# -*- coding: utf-8 -*-
"""
run_eval.py
===========
파싱 검증 파이프라인의 단일 진입점(러너).

두 가지 방식으로 쓴다.
  1) 문서셋 전체 평가:  run()             → DATA_DIR의 모든 문서를 돌며 WER/TEDS 계산,
                                             before→after 요약 + CSV 2개 저장.
  2) 문서 하나 진단:    diagnose(doc_id)  → 그 문서의 단어/문단 diff, 표 비교를
                                             콘솔에 자세히 출력(원인 파악용).

before/after는 있는 것만 비교한다. gt만 있으면 건너뛰고, gt+before만 있으면 before만,
gt+after만 있으면 after만 계산한다.

터미널 실행:  python run_eval.py
주피터 실행:  import run_eval; df, summary = run_eval.run()   (또는  %run run_eval.py)
"""

from pathlib import Path
import pandas as pd

import config
from gt_extract import extract_gt
from parsed_extract import load_parsed_txt
from metrics import compute_wer, average_teds
from xlsx_eval import evaluate_xlsx

_GT_EXTS = (".docx", ".pptx", ".xlsx")
# 긴 접미사부터 검사해야 "_before_gt"가 "_gt"로 잘못 매칭되지 않는다.
_ROLES = ("before_gt", "after_gt", "gt", "before", "after")
_ALL_ROLE_KEYS = ("gt", "before_gt", "after_gt", "before", "after")


# ══════════════════════════════════════════════════════════════════════
# 파일 수집: 폴더를 훑어 {문서id: {gt/before/after 경로}} 로 묶는다
# ══════════════════════════════════════════════════════════════════════
def _split_doc_id_and_role(stem: str):
    """
    파일명(확장자 제외)을 (문서id, 역할)로 분리.
      '1-1_before'          → ('1-1', 'before')
      '1-1_gt'              → ('1-1', 'gt')
      '1-1_before_gt'       → ('1-1', 'before_gt')   ← 가이드라인으로 원문 자체가
      '1-1_after_gt'        → ('1-1', 'after_gt')      바뀌는 "듀얼 GT" 케이스용

    듀얼 GT 케이스란: 가이드라인 적용은 "파서"가 아니라 "원문 문서 자체"를 고치는
    작업이므로, before 파싱 결과와 after 파싱 결과는 서로 다른 원본(전/후 문서)에서
    나온 것이다. 이 경우 before_gt/after_gt를 각각 따로 두면, before.txt는
    before_gt와, after.txt는 after_gt와 비교되어 "같은 문서인데 파서만 좋아졌는지"가
    아니라 "가이드라인 적용(문서 자체 개선 + 파서)의 최종 효과"를 정확히 잰다.
    before_gt/after_gt가 없으면 기존처럼 단일 _gt 하나를 before/after 둘 다에 쓴다
    (파서 알고리즘만 바뀌고 원문은 그대로인 경우).
    """
    for role in _ROLES:
        if stem.endswith("_" + role):
            return stem[: -(len(role) + 1)], role
    return None, None


def collect_docs() -> dict:
    """
    DATA_DIR 안의 파일을 문서id 기준으로 묶어
    {문서id: {"gt","before_gt","after_gt","before","after": path|None}} 반환.
    역할별 허용 확장자(gt류=docx/pptx/xlsx, before/after=txt)만 인정한다.
    """
    docs = {}
    if not config.DATA_DIR.exists():
        return docs
    for p in sorted(config.DATA_DIR.iterdir()):
        if not p.is_file():
            continue
        doc_id, role = _split_doc_id_and_role(p.stem)
        if role is None:
            continue
        ext = p.suffix.lower()
        if role in ("gt", "before_gt", "after_gt") and ext not in _GT_EXTS:
            continue
        if role in ("before", "after") and ext != ".txt":
            continue
        docs.setdefault(doc_id, {k: None for k in _ALL_ROLE_KEYS})[role] = p
    return dict(sorted(docs.items()))


def _gt_for_stage(files: dict, stage: str):
    """이 stage(before/after)에 쓸 GT 경로. 전용 {stage}_gt가 있으면 그것, 없으면 공용 gt."""
    return files.get(f"{stage}_gt") or files.get("gt")


# ══════════════════════════════════════════════════════════════════════
# WER 전처리(표 중복 제거 + 줄바꿈 유연화)를 한데 묶은 공유 헬퍼
#   run() / diagnose() / _make_html_report()가 전부 이걸 거쳐 같은 값을 내도록 한다.
#   (예전에 함수마다 로직이 달라 리포트와 표 값이 어긋나는 버그가 있었다.)
# ══════════════════════════════════════════════════════════════════════
def _wer_for_stage(pred_text, pred_tables, gt_text, gt_tables, is_pptx, prof):
    """
    docx/pptx 한 stage의 WER을 계산한다.
    ★ 표중복은 삭제/수정하지 않는다 - WER은 중복 포함 원문 그대로 계산한다.
      대신 리포트 표시용으로 '표 밖에 중복 추출된 라인'을 감지해 함께 반환한다.
    반환: (wer, 감지된_중복라인_리스트)
    """
    from wer_preprocess import detect_table_dup_text, relax_linebreaks
    from metrics import table_to_matrix, compute_wer
    from slide_eval import pptx_slide_wer

    # 표중복 "감지만"(제거 안 함) - 리포트 표시용
    dup_lines = []
    if pred_tables:
        matrices = [table_to_matrix(h) for h in pred_tables]
        dup_lines = detect_table_dup_text(pred_text, matrices)

    if is_pptx:
        # 슬라이드 단위 정렬 WER. pred_text는 중복 포함 원문 그대로.
        wer = pptx_slide_wer(gt_text, pred_text, prof["tokenizer"], prof["normalize"])["wer"] \
            if isinstance(gt_text, list) else None
        return wer, dup_lines
    else:
        # 줄바꿈 유연화만 적용(표중복은 그대로 둠)
        pred_text = relax_linebreaks(pred_text, gt_text)
        wer = compute_wer(pred_text, gt_text, tokenizer=prof["tokenizer"],
                          normalize_kwargs=prof["normalize"])["wer"]
        return wer, dup_lines


# ══════════════════════════════════════════════════════════════════════
# 문서셋 전체 평가
# ══════════════════════════════════════════════════════════════════════
def run():
    docs = collect_docs()
    if not docs:
        print(f"[경고] {config.DATA_DIR} 안에 '_gt/_before/_after' 파일이 없습니다.")
        return None, None

    rows = []
    for doc_id, files in docs.items():
        gt_before, gt_after = _gt_for_stage(files, "before"), _gt_for_stage(files, "after")
        if gt_before is None and gt_after is None:
            print(f"[건너뜀] {doc_id}: GT(_gt / _before_gt / _after_gt)가 없어 비교할 수 없습니다.")
            continue

        any_gt = gt_before or gt_after
        ext = any_gt.suffix.lower()
        is_xlsx = ext == ".xlsx"
        is_pptx = ext == ".pptx"
        dual_gt = files["before_gt"] is not None or files["after_gt"] is not None
        row = {"doc_id": doc_id, "type": ext[1:], "dual_gt": dual_gt}

        if is_xlsx:
            # 엑셀: TEDS(표) + WER(시트명/캡션 등 표 아닌 텍스트) 병행
            for stage in ("before", "after"):
                gt_path = _gt_for_stage(files, stage)
                if gt_path is None or files[stage] is None:
                    row[f"teds_{stage}"] = None
                    row[f"wer_{stage}"] = None
                    continue
                res = evaluate_xlsx(gt_path, files[stage])
                row[f"teds_{stage}"] = res["teds"]
                row[f"wer_{stage}"] = res["wer"]
        else:
            # docx/pptx: WER (활성 전처리 프로파일) + TEDS
            # 듀얼 GT면 stage마다 원문이 다르므로 GT 추출도 stage별로 각각 한다
            # (before_gt==after_gt인 단일 GT 모드에서는 캐시해서 같은 파일을 두 번
            #  추출하지 않도록 한다).
            prof = config.active_profile()
            _cache = {}

            def _extract_for(gt_path):
                key = str(gt_path)
                if key not in _cache:
                    if is_pptx:
                        from gt_extract import extract_pptx_slides
                        _cache[key] = extract_pptx_slides(str(gt_path))
                    else:
                        _cache[key] = extract_gt(str(gt_path))
                return _cache[key]

            for stage in ("before", "after"):
                gt_path = _gt_for_stage(files, stage)
                if gt_path is None or files[stage] is None:
                    row[f"wer_{stage}"] = row[f"teds_{stage}"] = None
                    continue
                gt_main, gt_tables = _extract_for(gt_path)  # gt_main: pptx면 slides 리스트, 아니면 text
                pred_text, pred_tables = load_parsed_txt(files[stage])
                # WER: 표 중복 본문 제거 + 줄바꿈 유연화 포함(공유 헬퍼)
                wer, _removed = _wer_for_stage(pred_text, pred_tables, gt_main, gt_tables,
                                               is_pptx, prof)
                row[f"wer_{stage}"] = wer
                # TEDS: 표는 그대로(전처리 영향 없음)
                row[f"teds_{stage}"] = average_teds(pred_tables, gt_tables)

        rows.append(row)

    if not rows:
        print("[경고] 비교 가능한 문서가 없습니다(모든 문서에 GT가 없음).")
        return None, None

    df = pd.DataFrame(rows)
    df.to_csv(config.REPORT_PER_DOC_CSV, index=False, encoding="utf-8-sig")

    summary = compute_summary(df)
    summary.to_csv(config.REPORT_SUMMARY_CSV, index=False, encoding="utf-8-sig")

    _print_report(df, summary)
    return df, summary


def compute_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    지표별 before/after 평균과 개선폭(after-before), 개선율(%)을 계산한다.
    데이터에 실제로 있는 지표(컬럼)만 요약에 넣는다.
      - 낮을수록 좋음: wer, cer          → 개선율 = (before - after) / before * 100
      - 높을수록 좋음: teds_content, cell_f1 → 개선율 = (after - before) / before * 100
    값이 없는 문서(before/after 미제공 또는 다른 문서 타입)는 평균에서 자동 제외(skipna).
    """
    # (표시명, before컬럼, after컬럼, 낮을수록 좋은지)
    metric_specs = [
        ("wer", "wer_before", "wer_after", True),
        # ("wer_norm", "wer_norm_before", "wer_norm_after", True),  # [WER정규화]
        ("teds_content", "teds_before", "teds_after", False),
    ]

    def _row(metric, before_col, after_col, lower_is_better):
        b = df[before_col].mean(skipna=True)
        a = df[after_col].mean(skipna=True)
        if pd.isna(b) or pd.isna(a):
            imp = None
        elif lower_is_better:
            imp = ((b - a) / b * 100) if b != 0 else None
        else:
            imp = ((a - b) / b * 100) if b != 0 else None
        return {
            "metric": metric,
            "before_mean": round(b, 4) if pd.notna(b) else None,
            "after_mean": round(a, 4) if pd.notna(a) else None,
            "absolute_change": round(a - b, 4) if pd.notna(a) and pd.notna(b) else None,
            "improvement_pct": round(imp, 2) if imp is not None else None,
            "direction": "감소(개선)" if lower_is_better else "증가(개선)",
        }

    records = []
    for metric, bcol, acol, lower in metric_specs:
        for col in (bcol, acol):
            if col not in df.columns:
                df[col] = None
        # before/after 둘 다 값이 하나도 없으면 그 지표는 이 데이터셋에 없는 것 → 생략
        if df[bcol].notna().any() or df[acol].notna().any():
            records.append(_row(metric, bcol, acol, lower))
    return pd.DataFrame(records)


def _print_report(df: pd.DataFrame, summary: pd.DataFrame):
    print("=" * 72)
    print(f"문서별 상세 결과 (전처리 프로파일: {config.EVAL_PROFILE}, "
          f"토큰: {config.active_profile()['tokenizer']})")
    print("=" * 72)
    print(df.to_string(index=False))

    print("\n" + "=" * 72)
    print("문서셋 전체 요약 (Before → After)")
    print("=" * 72)
    for _, r in summary.iterrows():
        print(f"- {r['metric']:<13}: {r['before_mean']} → {r['after_mean']}  "
              f"({r['direction']} {r['improvement_pct']}%)")
    print(f"\n[상세 CSV] {config.REPORT_PER_DOC_CSV}")
    print(f"[요약 CSV] {config.REPORT_SUMMARY_CSV}")


# ══════════════════════════════════════════════════════════════════════
# 문서 1건 진단 (원인 파악)
# ══════════════════════════════════════════════════════════════════════
def diagnose(doc_id: str, stage: str = "before", make_html: bool = False):
    """
    문서 하나를 골라 어디서 틀렸는지 자세히 본다.
      stage : "before" 또는 "after" (없는 stage를 고르면 안내 후 종료)
      make_html=True : report_html로 GT/Before/After 3-way HTML 리포트도 생성
    """
    from metrics import (print_paragraph_diff, print_wer_diff,
                         table_to_matrix, compare_tables_side_by_side)

    files = collect_docs().get(doc_id)
    gt_path = _gt_for_stage(files, stage) if files else None
    if files is None or gt_path is None:
        print(f"[오류] {doc_id}의 GT(_gt 또는 _{stage}_gt)를 찾을 수 없습니다.")
        return
    if files[stage] is None:
        print(f"[오류] {doc_id}에 {stage} 파일이 없습니다.")
        return

    if gt_path.suffix.lower() == ".xlsx":
        # 엑셀: 전용 전처리 후 TEDS + 정리된 표 미리보기
        from xlsx_eval import _merge_matrices, evaluate_xlsx
        from parsed_extract import load_parsed_matrices
        res = evaluate_xlsx(gt_path, files[stage])
        _, gt_html = extract_gt(str(gt_path))
        gt_m = _merge_matrices([table_to_matrix(h) for h in gt_html])
        pred_m = _merge_matrices(load_parsed_matrices(files[stage]))
        print("=" * 60)
        print(f"[{doc_id} / {stage}] 엑셀 요약  TEDS-content {res['teds']}  "
              f"(정리 후 GT {res['n_gt_rows']}행 / 파서 {res['n_pred_rows']}행)")
        print("\n[정리된 GT 표 (상위 8행)]")
        for r in gt_m[:8]:
            print("  ", r)
        print("\n[정리된 파서 표 (상위 8행)]")
        for r in pred_m[:8]:
            print("  ", r)
        if make_html:
            _make_html_report(doc_id, files)
        return

    prof = config.active_profile()
    is_pptx = gt_path.suffix.lower() == ".pptx"
    gt_text, gt_tables = extract_gt(str(gt_path))   # 표시/표비교용(flat 텍스트)
    pred_text, pred_tables = load_parsed_txt(files[stage])

    # WER은 run()과 동일하게 계산(표중복은 삭제 안 함, 줄바꿈 유연화만; pptx는 슬라이드 정렬)
    if is_pptx:
        from gt_extract import extract_pptx_slides
        gt_main, _ = extract_pptx_slides(str(gt_path))   # 슬라이드 리스트
    else:
        gt_main = gt_text
    wer_val, dup_lines = _wer_for_stage(pred_text, pred_tables, gt_main, gt_tables, is_pptx, prof)

    teds = average_teds(pred_tables, gt_tables)
    print("=" * 60)
    print(f"[{doc_id} / {stage}] 요약  (프로파일 {config.EVAL_PROFILE})  "
          f"WER {wer_val:.4f}  |  TEDS-content {teds}")
    if dup_lines:
        print(f"\n[표 내용 중복 추출] 표 밖에 중복 추출된 것으로 '표시'된 라인 {len(dup_lines)}개 "
              f"(WER 계산엔 그대로 포함됨):")
        for d in dup_lines:
            print("   -", d)

    print("\n" + "=" * 60 + "\n[문단 단위 diff] (큰 그림: 어디서 갈라지는지)")
    print_paragraph_diff(pred_text, gt_text)

    print("\n" + "=" * 60 + "\n[단어 단위 diff] (어떤 단어가 문제인지)")
    if is_pptx:
        # pptx는 슬라이드 단위 정렬 WER이라 통짜 diff는 위 요약 숫자와 다를 수 있음(참고용).
        print("(참고: PPT는 슬라이드 단위 정렬로 WER을 내므로, 아래 통짜 diff의 합계는")
        print(" 위 요약 WER과 다를 수 있습니다. 슬라이드별 상세는 by_unit()을 쓰세요.)")
        print_wer_diff(pred_text, gt_text, tokenizer=prof["tokenizer"], normalize_kwargs=prof["normalize"])
    else:
        # docx: WER 숫자와 일치하도록 줄바꿈 유연화만 적용한 pred로 diff(표중복은 그대로)
        from wer_preprocess import relax_linebreaks
        _pred = relax_linebreaks(pred_text, gt_text)
        print_wer_diff(_pred, gt_text, tokenizer=prof["tokenizer"], normalize_kwargs=prof["normalize"])

    if gt_tables:
        print("\n" + "=" * 60 + "\n[표 비교] (첫 번째 표)")
        print("GT   :", table_to_matrix(gt_tables[0]))
        if pred_tables:
            print("PRED :", table_to_matrix(pred_tables[0]))
            print("\n셀 단위 비교:")
            print(compare_tables_side_by_side(pred_tables[0], gt_tables[0]).to_string(index=False))
        else:
            print("PRED : (표를 찾지 못함 - txt에 <table> 태그가 있는지 확인)")

    if make_html:
        _make_html_report(doc_id, files)


def by_unit(doc_id: str, stage: str = "before"):
    """
    문서를 단위별로 쪼개 점수를 본다.
      - PPT : 슬라이드별 WER (파서 ## 섹션과 정렬한 결과)
      - 엑셀 : 시트별 TEDS (GT 시트 ↔ 파서 표를 순서대로 짝지음)
    "어느 슬라이드/시트에서 점수가 깨지는지"를 콕 집어 볼 때 쓴다.
    """
    files = collect_docs().get(doc_id)
    gt_path = _gt_for_stage(files, stage) if files else None
    if files is None or gt_path is None or files[stage] is None:
        print(f"[오류] {doc_id}의 GT 또는 {stage} 파일이 없습니다.")
        return
    ext = gt_path.suffix.lower()
    prof = config.active_profile()

    if ext == ".pptx":
        from gt_extract import extract_pptx_slides
        from slide_eval import pptx_slide_wer
        gt_slides, _ = extract_pptx_slides(str(gt_path))
        pred_text, _ = load_parsed_txt(files[stage])
        res = pptx_slide_wer(gt_slides, pred_text, prof["tokenizer"], prof["normalize"])
        print(f"[{doc_id} / {stage}] 슬라이드 정렬 WER {res['wer']:.4f} "
              f"({'슬라이드 정렬됨' if res['aligned'] else '헤더 없어 통짜 비교'})")
        if res.get("pairs"):
            print("-" * 60)
            for p in res["pairs"]:
                gt_no = "-" if p["gt_idx"] is None else f"S{p['gt_idx']+1}"
                pr_no = "-" if p["pred_idx"] is None else f"#{p['pred_idx']+1}"
                tag = {"match": "", "gt_only": "  ← 파서가 슬라이드 놓침",
                       "pred_only": "  ← 파서 잉여 슬라이드"}[p["kind"]]
                print(f"  GT {gt_no:>4} ↔ 파서 {pr_no:>4} : WER {p['wer']:.4f} "
                      f"(S{p['S']} D{p['D']} I{p['I']}/N{p['N']}){tag}")
        return

    if ext == ".xlsx":
        from xlsx_eval import _merge_matrices, _matrix_to_html
        from metrics import table_to_matrix, compute_teds_content
        from parsed_extract import load_parsed_matrices
        _, gt_html = extract_gt(str(gt_path))
        gt_sheets = [table_to_matrix(h) for h in gt_html]          # 시트별
        pred_tables = load_parsed_matrices(files[stage])           # 파서 표들
        print(f"[{doc_id} / {stage}] 시트별 TEDS  (GT 시트 {len(gt_sheets)}개 / 파서 표 {len(pred_tables)}개)")
        print("-" * 60)
        for k in range(max(len(gt_sheets), len(pred_tables))):
            gt_m = _merge_matrices([gt_sheets[k]]) if k < len(gt_sheets) else []
            pr_m = _merge_matrices([pred_tables[k]]) if k < len(pred_tables) else []
            if not gt_m and not pr_m:
                continue
            teds = compute_teds_content(_matrix_to_html(pr_m), _matrix_to_html(gt_m))["teds_content"]
            print(f"  시트/표 {k+1}: TEDS {teds:.4f}  (GT {len(gt_m)}행 / 파서 {len(pr_m)}행)")
        return

    print("[안내] by_unit은 PPT(슬라이드)/엑셀(시트) 전용입니다. docx는 diagnose를 쓰세요.")


def compare_profiles(doc_id: str, stage: str = "before"):
    """
    문서 하나(docx/pptx)를 여러 전처리 케이스로 WER을 내서 비교한다.
    "이 문서 점수가 나쁜 게 띄어쓰기 밀림 때문인지, 서식 때문인지, 진짜 내용 때문인지"를
    프로파일을 바꿔가며 확인할 때 쓴다. (엑셀은 셀 F1을 쓰므로 이 함수 대상이 아님.)
    """
    files = collect_docs().get(doc_id)
    gt_path = _gt_for_stage(files, stage) if files else None
    if files is None or gt_path is None or files[stage] is None:
        print(f"[오류] {doc_id}의 GT(_gt/_{stage}_gt) 또는 {stage} 파일이 없습니다.")
        return
    if gt_path.suffix.lower() == ".xlsx":
        print("[안내] 엑셀은 셀 F1을 쓰므로 프로파일 비교 대상이 아닙니다.")
        return

    gt_text, _ = extract_gt(str(gt_path))
    pred_text, _ = load_parsed_txt(files[stage])

    print(f"[{doc_id} / {stage}] 전처리 케이스별 WER")
    print("-" * 50)
    for name, prof in config.PROFILES.items():
        try:
            wer = compute_wer(pred_text, gt_text, tokenizer=prof["tokenizer"],
                              normalize_kwargs=prof["normalize"])["wer"]
            print(f"  {name:<15}(토큰 {prof['tokenizer']:<8}): WER {wer:.4f}")
        except ImportError as e:
            print(f"  {name:<15}: (건너뜀 - {e})")


def _make_html_report(doc_id: str, files: dict = None, output_path: str = None):
    """
    GT/Before/After 3-way HTML 리포트 생성(report_html.py 필요).
    듀얼 GT(before_gt/after_gt) 문서는 before쪽엔 before_gt, after쪽엔 after_gt를 쓴다
    (파서만 비교하는 게 아니라 "원문 개선+파싱" 최종 결과를 3-way로 보여줘야 하므로).
    """
    from report_html import generate_report

    files = files or collect_docs().get(doc_id)
    gt_before_path = _gt_for_stage(files, "before") if files else None
    gt_after_path = _gt_for_stage(files, "after") if files else None
    if files is None or (gt_before_path is None and gt_after_path is None):
        print(f"[오류] {doc_id}의 GT(_gt / _before_gt / _after_gt)를 찾을 수 없습니다.")
        return

    out = output_path or str(config.OUTPUT_DIR / f"report_{doc_id}.html")
    any_gt = gt_before_path or gt_after_path

    # ── 엑셀: 텍스트(WER) 대신, 정리된 표(패딩 제거)만 3-way로 비교 ──
    if any_gt.suffix.lower() == ".xlsx":
        from xlsx_eval import _merge_matrices, _matrix_to_html, evaluate_xlsx
        from metrics import table_to_matrix
        from parsed_extract import load_parsed_matrices

        def _gt_tbl(gt_path):
            if gt_path is None:
                return []
            _, gt_html = extract_gt(str(gt_path))
            return [_matrix_to_html(_merge_matrices([table_to_matrix(h) for h in gt_html]))]

        def _pred_tbl(stage):
            if not files[stage]:
                return []
            return [_matrix_to_html(_merge_matrices(load_parsed_matrices(files[stage])))]

        teds_b = evaluate_xlsx(gt_before_path, files["before"])["teds"] if (gt_before_path and files["before"]) else None
        teds_a = evaluate_xlsx(gt_after_path, files["after"])["teds"] if (gt_after_path and files["after"]) else None

        # GT 패널은 after_gt(최종본)를 우선 보여주고, before_gt만 있으면 그것을 보여준다.
        gt_display = _gt_tbl(gt_after_path or gt_before_path)

        path = generate_report("", gt_display, "", _pred_tbl("before"), "", _pred_tbl("after"),
                               None, None, teds_b, teds_a, output_path=out)
        print(f"[HTML 리포트] {path}")
        return

    # ── docx/pptx: 본문 텍스트 + 표 ──
    is_pptx = any_gt.suffix.lower() == ".pptx"

    def _extract(gt_path):
        """리포트 표시용 GT 본문/표. pptx는 슬라이드를 이어붙인 flat 텍스트로."""
        if gt_path is None:
            return "", []
        if is_pptx:
            from gt_extract import extract_pptx_slides
            slides, tables = extract_pptx_slides(str(gt_path))
            return "\n".join(s for s in slides if s.strip()), tables
        return extract_gt(str(gt_path))

    def _extract_wer_gt(gt_path):
        """WER 계산용 GT. pptx는 슬라이드 리스트 그대로(슬라이드 정렬에 필요)."""
        if gt_path is None:
            return None, []
        if is_pptx:
            from gt_extract import extract_pptx_slides
            return extract_pptx_slides(str(gt_path))   # (slides_list, tables)
        return extract_gt(str(gt_path))

    gt_before_text, gt_before_tables = _extract(gt_before_path)
    gt_after_text, gt_after_tables = _extract(gt_after_path or gt_before_path)
    before_text, before_tables = load_parsed_txt(files["before"]) if files["before"] else ("", [])
    after_text, after_tables = load_parsed_txt(files["after"]) if files["after"] else ("", [])

    prof = config.active_profile()

    def _stage_wer_and_dups(stage, gt_path, pred_text, pred_tables):
        if not files[stage] or gt_path is None:
            return None, []
        gt_main, gt_tables = _extract_wer_gt(gt_path)
        return _wer_for_stage(pred_text, pred_tables, gt_main, gt_tables, is_pptx, prof)

    wer_b, before_dups = _stage_wer_and_dups("before", gt_before_path, before_text, before_tables)
    wer_a, after_dups = _stage_wer_and_dups("after", gt_after_path or gt_before_path, after_text, after_tables)
    teds_b = average_teds(before_tables, gt_before_tables) if gt_before_path else None
    teds_a = average_teds(after_tables, gt_after_tables) if (gt_after_path or gt_before_path) else None

    # ★ 표중복은 본문에서 제거하지 않는다 - 원문 그대로 리포트에 넘긴다.
    #   before_dups/after_dups(감지된 중복 라인)는 리포트에서 색으로 '표시'만 하는 데 쓴다.
    before_text_for_report = before_text
    after_text_for_report = after_text

    # 줄바꿈 유연화만 리포트 본문에 반영(로그도 수집). (docx만; pptx는 슬라이드 단위)
    relax_log_b, relax_log_a = [], []
    if not is_pptx:
        from wer_preprocess import relax_linebreaks
        if isinstance(gt_before_text, str) and before_text_for_report:
            before_text_for_report, relax_log_b = relax_linebreaks(
                before_text_for_report, gt_before_text, return_log=True)
        gt_a_text = gt_after_text if isinstance(gt_after_text, str) else gt_before_text
        if isinstance(gt_a_text, str) and after_text_for_report:
            after_text_for_report, relax_log_a = relax_linebreaks(
                after_text_for_report, gt_a_text, return_log=True)

    # gt_text 인자엔 before쪽 GT를, gt_after_text엔 after쪽 GT를 넣어
    # (듀얼 GT면) before는 before_gt와, after는 after_gt와 각각 비교되도록 한다.
    path = generate_report(
        gt_before_text or gt_after_text, gt_before_tables or gt_after_tables,
        before_text_for_report, before_tables, after_text_for_report, after_tables,
        wer_b, wer_a, teds_b, teds_a, output_path=out,
        gt_after_text=gt_after_text, gt_after_tables=gt_after_tables,
        before_dups=before_dups, after_dups=after_dups,
        relax_log_before=relax_log_b, relax_log_after=relax_log_a)
    print(f"[HTML 리포트] {path}")
    return path


if __name__ == "__main__":
    run()
