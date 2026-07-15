# -*- coding: utf-8 -*-
"""
xlsx_eval.py
============
GT가 엑셀(.xlsx)일 때 쓰는 별도 평가 코드. 지표는 docx/pptx와 같은 TEDS-content를 쓴다.

엑셀은 왜 전처리가 필요한가
---------------------------
엑셀은 문서 전체가 표이고, 파서는 시트 전체를 훑어 빈 칸을 null로 채운 마크다운
파이프 표(| 셀 | 셀 |)로 내보낸다. 그래서 실제 표 밖까지 null 패딩 행이 잔뜩 붙고,
빈 헤더는 "Unnamed: 1"(pandas 흔적)로 나온다. 이걸 그대로 TEDS에 넣으면 패딩
때문에 노드 수가 부풀어 점수가 왜곡된다.

그래서 TEDS를 계산하기 "전에" 아래처럼 엑셀 전용 정리를 한다.
  1) 셀 표기 통일: null / Unnamed: N / nan 등을 빈 칸("")으로  (metrics._normalize_cell)
  2) 여러 표(파서의 md 표 블록, GT의 여러 시트)를 한 표로 세로로 이어붙임
  3) 완전히 빈 행/열을 (가장자리뿐 아니라 중간까지) 전부 제거 → 패딩 제거
  4) 정리된 GT/파서 표 하나끼리 TEDS-content 계산

이렇게 하면 "실제 값이 있는 표 영역"끼리만 구조+내용을 비교하게 되어, 패딩에 휘둘리지
않는 TEDS 점수가 나온다. (표를 하나로 합쳐 비교하므로, 파서가 표를 몇 조각으로 나눠
내보냈는지와도 무관하다.)
"""

import html as html_lib

from gt_extract import extract_gt
from parsed_extract import load_parsed_matrices
from metrics import table_to_matrix, _normalize_cell, compute_teds_content


def _clean_matrix(rows: list) -> list:
    """셀 표기 통일 후 완전히 빈 행/열을 모두 제거한다."""
    m = [[_normalize_cell(c) for c in row] for row in rows]
    m = [row for row in m if any(c != "" for c in row)]     # 빈 행 제거
    if not m:
        return []
    n_cols = max(len(r) for r in m)
    m = [r + [""] * (n_cols - len(r)) for r in m]           # 열 길이 맞춤
    keep = [c for c in range(n_cols) if any(row[c] != "" for row in m)]  # 빈 열 제거
    return [[row[c] for c in keep] for row in m]


def _merge_matrices(matrices: list) -> list:
    """여러 표(행렬)를 세로로 이어붙여 하나로 만든 뒤 정리한다."""
    rows = [row for mat in matrices for row in mat]
    return _clean_matrix(rows)


def _matrix_to_html(matrix: list) -> str:
    body = "".join(
        "<tr>" + "".join(f"<td>{html_lib.escape(c)}</td>" for c in row) + "</tr>"
        for row in matrix
    )
    return f"<table>{body}</table>"


def evaluate_xlsx(gt_path, parsed_path) -> dict:
    """
    엑셀 GT와 파서 결과(txt)를 TEDS-content로 비교한다.
    반환: {"teds": 0~1 또는 None, "n_gt_rows", "n_pred_rows"}
    파서 파일이 없으면 teds=None.

    엑셀도 시트명(## 시트명)이나 셀 밖 텍스트 같은 "표가 아닌 텍스트"가 있으면
    그 부분은 WER로 함께 잰다(표는 TEDS, 텍스트는 WER 병행). 시트 간 연결 캡션처럼
    값은 같아도 텍스트로만 드러나는 차이를 TEDS가 놓치는 걸 보완한다.
    반환에 wer 키가 추가된다(잴 텍스트가 없으면 wer=None).
    """
    if parsed_path is None:
        return {"teds": None, "wer": None, "n_gt_rows": None, "n_pred_rows": None}

    # GT: 시트명(text) + 시트별 표(HTML)
    gt_text, gt_html_tables = extract_gt(str(gt_path))
    gt_matrix = _merge_matrices([table_to_matrix(h) for h in gt_html_tables])

    # 파서: 파이프/HTML 표 → 행렬 → 하나로 병합/정리
    from parsed_extract import load_parsed_txt
    pred_text, _ = load_parsed_txt(parsed_path)          # 표를 뺀 본문(시트명 ## 등)
    pred_matrix = _merge_matrices(load_parsed_matrices(parsed_path))

    # 표(TEDS)
    if not gt_matrix and not pred_matrix:
        teds = 1.0
    else:
        teds = compute_teds_content(_matrix_to_html(pred_matrix), _matrix_to_html(gt_matrix))["teds_content"]

    # 표 아닌 텍스트(시트명 ## 등)가 GT나 파서에 있으면 WER 병행
    wer = None
    if gt_text.strip() or pred_text.strip():
        from metrics import compute_wer
        import config
        prof = config.active_profile()
        wer = compute_wer(pred_text, gt_text, tokenizer=prof["tokenizer"],
                          normalize_kwargs=prof["normalize"])["wer"]

    return {"teds": teds, "wer": wer, "n_gt_rows": len(gt_matrix), "n_pred_rows": len(pred_matrix)}
