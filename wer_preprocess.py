# -*- coding: utf-8 -*-
"""
wer_preprocess.py
=================
WER 계산에만 적용하는 "구조적 전처리" 두 가지. (TEDS/표에는 절대 적용하지 않는다.)

여기 있는 처리는 normalize.py(문자/기호 표기 통일)와 성격이 다르다. normalize는
"같은 글자를 같게" 만드는 표면 정규화이고, 여기는 "표와 본문의 관계" 같은 구조를
근거로 본문 라인을 덜어내거나(1) 토큰 비교를 줄바꿈 경계에서 유연하게(2) 만든다.

────────────────────────────────────────────────────────────────────────
(1) 표 안 내용이 본문에도 중복 추출된 경우, 본문 쪽 중복만 제거  (remove_table_dup_text)
────────────────────────────────────────────────────────────────────────
파서가 표를 <table>로도 내보내고, 같은 내용을 표 밖 본문 텍스트로도 한 번 더
쏟아내는 경우가 있다(표 셀을 줄줄이 풀어써서). 이때 본문에 남은 중복은 실제 문서에
없던 잉여라서 WER 삽입(I) 오류를 부풀린다. 그래서 WER 계산 "전에", 표 셀에 있던
내용과 일치하는 본문 라인을 본문에서 제거한다.

  - 대상: 엑셀 제외(docx/pptx). 그리고 "완전히 빈 표"(구조만 있고 셀이 전부 빈 표)는
          중복 판단 기준으로 쓰지 않는다(빈 셀과 빈 본문 라인이 엮이지 않도록).
  - "중복"은 라인 전체가 어떤 표 셀(또는 한 행의 셀들을 이어붙인 것)과 정규화 후
    같을 때만으로 본다. 부분 포함은 건드리지 않는다(과삭제 방지).
  - HTML 리포트에서는 제거하지 않는다. 대신 어떤 라인이 중복으로 빠졌는지 목록을
    함께 돌려주어("표 내용 중복 추출") 리포트가 따로 표시할 수 있게 한다.
  - "글 사이 중복"(표와 무관하게 본문 안에서 같은 문장이 반복되는 것)은 손대지 않는다.
    여기서는 오직 "표 셀 내용과 겹치는 본문 라인"만 제거 대상이다.

────────────────────────────────────────────────────────────────────────
(2) 줄바꿈 경계 유연화  (relax_linebreaks)
────────────────────────────────────────────────────────────────────────
파서가 원문에 없던 줄바꿈을 넣거나(한 문장을 두 줄로 쪼갬), 있어야 할 줄바꿈을
빠뜨리는 경우가 있다. WER은 결국 공백으로 이어붙여 토큰을 세므로 줄바꿈 자체는
토큰에 큰 영향이 없지만, "줄 단위로 무엇을 지우고 남기는지"를 판단하는 (1) 같은
처리에서는 줄 경계가 어긋나면 오정렬이 난다.

이 함수는 GT와 파서 텍스트를 줄 단위로 정렬하되, "바로 앞뒤 줄을 이었을 때 상대편의
한 줄과 같아지는가"만 국소적으로 확인해서, 같아지면 그 경계를 GT 쪽 표기에 맞춰
합치거나 나눈다. 전체를 자유롭게 재정렬하지 않고 인접(±1줄)만 보므로, 뒤 내용이
통째로 밀려 계산이 무너지는 일이 없다.

기본은 보수적으로 "인접 2줄 병합 → 상대 1줄과 일치" 케이스만 처리한다.
"""

import re

from normalize import normalize_for_wer


# ══════════════════════════════════════════════════════════════════════
# (1) 표 중복 본문 라인 제거
# ══════════════════════════════════════════════════════════════════════
def _norm_line(s: str) -> str:
    """중복 판정용 정규화. WER 정규화와 같은 기준 + 앞뒤 공백 제거."""
    return normalize_for_wer(s).strip()


def _table_cell_set(matrix: list) -> set:
    """표 한 개의 '비지 않은 셀값' 집합(정규화). 완전 빈 표면 빈 집합."""
    s = set()
    for row in matrix:
        for c in row:
            nc = _norm_line(c)
            if nc:
                s.add(nc)
    return s


def detect_table_dup_text(text: str, tables_matrices: list, dup_ratio: float = 0.9):
    """
    본문 text에서 '표 내용이 표 밖에도 통째로 중복 추출된 라인'을 찾아 표시한다.
    ★ 삭제/수정하지 않는다. WER은 중복을 그대로 포함해 계산한다. 이 함수는 오직
      report.html에서 그 라인들을 색으로 표시하기 위한 "감지"용이다.

    반환: 중복으로 판단된 라인 문자열 리스트(원문 그대로).

    판정 방식:
      - 표마다 '비지 않은 셀값 집합'을 만든다. 완전히 빈 표(구조만 있고 셀이 전부 빈
        표)는 판정에서 제외한다.
      - 본문 라인을 공백/파이프로 조각내어 표 셀값과 맞춘다. 한 라인이 표 셀들로만
        이루어져 있으면(예: "저축은행 429 322" = 한 행) 그 라인이 덮는 셀들을 집계.
      - 표 셀의 dup_ratio(기본 90%) 이상이 본문에 중복 추출됐으면 → 그 라인들을 중복으로
        표시. 소수(90% 미만)만 겹치면 표시 안 함(한두 개 우연 중복 보존).
      - 표와 무관한 '글 사이 중복'은 표 셀값이 아니므로 대상 아님.
    """
    cell_sets = [s for s in (_table_cell_set(m) for m in tables_matrices) if s]
    if not cell_sets:
        return []

    lines = text.split("\n")

    def _line_cells(line):
        s = line.strip()
        if not s:
            return []
        if "|" in s:
            parts = [p.strip() for p in s.strip("|").split("|")]
        else:
            parts = s.split()
        return [_norm_line(p) for p in parts if _norm_line(p)]

    dup_idx = set()
    for cells in cell_sets:
        matched_idx, covered = [], set()
        for i, ln in enumerate(lines):
            lc = _line_cells(ln)
            if not lc:
                continue
            hits = [c for c in lc if c in cells]
            if hits and len(hits) >= max(1, len(lc) // 2):
                matched_idx.append(i)
                covered.update(hits)
        if matched_idx and len(covered) / len(cells) >= dup_ratio:
            dup_idx.update(matched_idx)

    return [lines[i] for i in sorted(dup_idx)]


# ══════════════════════════════════════════════════════════════════════
# (2) 줄바꿈 경계 유연화
# ══════════════════════════════════════════════════════════════════════
def relax_linebreaks(pred_text: str, gt_text: str, return_log: bool = False):
    """
    파서 텍스트의 줄바꿈을 GT 줄바꿈 경계에 국소적으로 맞춘다.
    지금은 가장 흔한 두 경우만 보수적으로 처리한다.

      (a) 파서가 GT 한 줄을 두 줄로 쪼갠 경우:
          파서의 인접 두 줄 pred[i], pred[i+1]을 이어붙인 게 GT 어떤 한 줄과
          같으면 → 파서의 두 줄을 한 줄로 합친다.
      (b) 파서가 GT 두 줄을 한 줄로 붙인 경우:
          파서의 한 줄 pred[i]가 GT 인접 두 줄 gt[j], gt[j+1]을 이어붙인 것과
          같으면 → 파서 한 줄을 두 줄로 나눈다.

    "이어붙임"은 공백으로 잇고 정규화해서 비교한다. GT 줄 집합을 기준으로만
    보므로, 앞뒤가 밀려 전체가 어긋나는 일은 없다(±1줄 국소 확인).
    파서 라인 자체(원문)는 유지하고 경계만 바꾼다.

    return_log=True면 (결과텍스트, 변경로그) 튜플을 반환한다.
    변경로그 = [{"kind":"merge"/"split", "from":[...], "to":[...]}, ...]
    """
    gt_lines = [ln for ln in gt_text.split("\n")]
    gt_norm = [_norm_line(ln) for ln in gt_lines]
    gt_norm_set = set(n for n in gt_norm if n)
    # GT 인접 2줄 병합 문자열 집합(케이스 b 판정용)
    gt_pair_merged = {}
    for j in range(len(gt_lines) - 1):
        merged = _norm_line(gt_lines[j] + " " + gt_lines[j + 1])
        if merged:
            gt_pair_merged[merged] = (gt_lines[j], gt_lines[j + 1])

    pred_lines = pred_text.split("\n")
    out, log = [], []
    i = 0
    while i < len(pred_lines):
        cur = pred_lines[i]
        nxt = pred_lines[i + 1] if i + 1 < len(pred_lines) else None

        # (a) 파서 두 줄 병합 == GT 한 줄  → 합침
        if nxt is not None:
            merged = _norm_line(cur + " " + nxt)
            if merged and merged in gt_norm_set \
               and _norm_line(cur) not in gt_norm_set:
                out.append(cur + " " + nxt)
                log.append({"kind": "merge", "from": [cur, nxt], "to": [cur + " " + nxt]})
                i += 2
                continue

        # (b) 파서 한 줄 == GT 두 줄 병합  → 나눔
        cur_norm = _norm_line(cur)
        if cur_norm and cur_norm in gt_pair_merged and cur_norm not in gt_norm_set:
            a, b = gt_pair_merged[cur_norm]
            out.append(a)
            out.append(b)
            log.append({"kind": "split", "from": [cur], "to": [a, b]})
            i += 1
            continue

        out.append(cur)
        i += 1

    result = "\n".join(out)
    if return_log:
        return result, log
    return result
