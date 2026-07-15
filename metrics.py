# -*- coding: utf-8 -*-
"""
metrics.py
==========
파싱 품질 지표와 진단 도구를 모은 모듈.

  1) WER (Word Error Rate)   : 본문 텍스트 비교 (낮을수록 좋음). 유일한 텍스트 지표.
                               전처리 방식과 토큰 단위(어절/형태소)는 "프로파일"로 바꾼다
                               (config.PROFILES 참고). 띄어쓰기/줄바꿈 밀림은 형태소
                               토큰화 프로파일로 흡수한다.
  2) TEDS-content            : 표 구조+셀 내용 비교 (0~1, 높을수록 좋음).
                               docx/pptx 표와 엑셀(전처리 후) 모두 이 지표를 쓴다.
                               표 전용 2단계(행→셀) 편집거리로 직접 계산한다(외부 라이브러리 불필요).
  + WER 정규화 wer/(1+wer)(선택), 셀 F1(선택, 기본 미사용), 진단 도구
"""

import re
from collections import Counter
from difflib import SequenceMatcher

from normalize import normalize_for_wer


# ══════════════════════════════════════════════════════════════════════
# 1. WER (Word Error Rate)
# ══════════════════════════════════════════════════════════════════════
# 정의: WER = (S + D + I) / N
#   S=치환, D=삭제, I=삽입 단어 수, N=GT 단어 수.
# 계산 원리는 CER(글자 단위)와 같은 Levenshtein 거리이고, 비교 단위만 "단어(토큰)"다.
#
# 한국어에서 "단어"를 정하는 두 방식:
#   - eojeol(어절)   : 공백 기준으로 자름. 별도 설치 불필요. 국내 STT/OCR 평가의 관행 기본값.
#                      "아닙니다"/"아니에요"를 서로 다른 토큰으로 봐 활용형 차이도 오류로 셈.
#   - morpheme(형태소): 형태소 분석기(kiwipiepy)로 어간/어미/조사를 쪼갠 뒤 비교.
#                      활용형·조사 차이에 좀 더 관대. pip install kiwipiepy 필요.

_kiwi_instance = None  # Kiwi()는 초기화 비용이 있어 최초 1회만 만들어 재사용


def _get_kiwi():
    global _kiwi_instance
    if _kiwi_instance is None:
        try:
            from kiwipiepy import Kiwi
        except ImportError as e:
            raise ImportError(
                "형태소 단위 WER(tokenizer='morpheme')에는 kiwipiepy가 필요합니다.\n"
                "설치: pip install kiwipiepy"
            ) from e
        _kiwi_instance = Kiwi()
    return _kiwi_instance


def tokenize_eojeol(text: str) -> list:
    """어절(공백) 단위 토큰화."""
    return text.split()


def tokenize_morpheme(text: str) -> list:
    """형태소 단위 토큰화(kiwipiepy). 각 형태소의 표면형(form)만 반환."""
    return [tok.form for tok in _get_kiwi().tokenize(text)]


def _tokenize(text: str, tokenizer: str) -> list:
    if tokenizer == "eojeol":
        return tokenize_eojeol(text)
    if tokenizer == "morpheme":
        return tokenize_morpheme(text)
    raise ValueError(f"알 수 없는 tokenizer: {tokenizer} (eojeol 또는 morpheme)")


def _levenshtein_sdi(pred_tokens: list, gt_tokens: list):
    """토큰 리스트 Levenshtein 거리로 (S, D, I)를 구한다."""
    n, m = len(pred_tokens), len(gt_tokens)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        dp[i][0] = i
    for j in range(m + 1):
        dp[0][j] = j
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if pred_tokens[i - 1] == gt_tokens[j - 1]:
                dp[i][j] = dp[i - 1][j - 1]
            else:
                dp[i][j] = 1 + min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1])

    # back-trace. pred가 행(i), gt가 열(j)이므로:
    #   i만 줄면 → pred에만 있는 토큰 = 삽입(I)
    #   j만 줄면 → gt에만 있는 토큰   = 삭제(D, 파서가 놓침)
    i, j = n, m
    S = D = I = 0
    while i > 0 or j > 0:
        if i > 0 and j > 0 and pred_tokens[i - 1] == gt_tokens[j - 1] and dp[i][j] == dp[i - 1][j - 1]:
            i -= 1; j -= 1
        elif i > 0 and j > 0 and dp[i][j] == dp[i - 1][j - 1] + 1:
            S += 1; i -= 1; j -= 1
        elif i > 0 and dp[i][j] == dp[i - 1][j] + 1:
            I += 1; i -= 1
        else:
            D += 1; j -= 1
    return S, D, I


# ══════════════════════════════════════════════════════════════════════
# WER 정규화  (지금은 꺼짐 - 필요할 때 켜기)
# ----------------------------------------------------------------------
# WER은 상한이 없어서, 파서가 GT보다 훨씬 많은 잡음을 뱉으면(삽입 I 폭증) WER이
# 1.5·3.0·10.0처럼 1을 넘어 문서마다 스케일이 들쑥날쑥해진다. 그럴 때 아래 식으로
# [0,1) 범위에 눌러 담아 before→after 비교를 직관적으로 만들 수 있다.
#     wer / (1 + wer)
# 클리핑(min(wer,1))과 달리 1.5·3.0·10.0을 각각 0.6·0.75·0.91로 서로 다르게
# 눌러(정보 손실 없이) 극단 케이스 간 차이가 살아있고, 파라미터가 없어 재현 가능하다.
#
# 켜는 법: 이 파일과 run_eval.py 두 곳에서, 줄 끝에 대괄호 마커가 붙어 있는
#          주석 줄들(아래 함수 3줄 + compute_wer 반환의 wer_norm 줄 +
#          run_eval.py의 저장 줄 + 요약 줄)의 맨 앞 # 를 지우면 된다.
#          에디터에서 대괄호 마커로 검색하면 그 줄들이 전부 잡힌다.
#
# def normalized_wer(wer: float) -> float:          # [WER정규화]
#     """WER 정규화: wer / (1 + wer) → [0, 1) 범위로 압축."""   # [WER정규화]
#     return wer / (1.0 + wer)                      # [WER정규화]
# ══════════════════════════════════════════════════════════════════════


def compute_wer(pred_text: str, gt_text: str, tokenizer: str = "eojeol",
                apply_normalization: bool = True, normalize_kwargs: dict = None) -> dict:
    """
    본문 텍스트 두 개를 비교해 WER과 세부 내역(S/D/I/N)을 반환한다.
    apply_normalization=True면 normalize_for_wer()를 먼저 적용해 표기 차이를 제거한다.
    normalize_kwargs로 전처리 옵션(프로파일)을 넘길 수 있다(예: {"strip_bold": True}).
    """
    if apply_normalization:
        opts = normalize_kwargs or {}
        pred_text = normalize_for_wer(pred_text, **opts)
        gt_text = normalize_for_wer(gt_text, **opts)

    pred_tokens = _tokenize(pred_text, tokenizer)
    gt_tokens = _tokenize(gt_text, tokenizer)

    S, D, I = _levenshtein_sdi(pred_tokens, gt_tokens)
    N = len(gt_tokens)
    wer = (S + D + I) / N if N > 0 else 0.0

    return {
        "wer": wer,
        # "wer_norm": normalized_wer(wer),   # [WER정규화]
        "S": S, "D": D, "I": I, "N": N, "tokenizer": tokenizer,
    }


# ══════════════════════════════════════════════════════════════════════
# 2. TEDS-content (표 비교)
# ══════════════════════════════════════════════════════════════════════
# 표는 table>tr>td 로 얕고 규칙적인 구조라, 범용 트리편집거리(zss) 대신 "표 전용
# 2단계 편집거리"로 직접 계산한다. (외부 라이브러리 불필요 + 동작이 결정적)
#   1단계(행 정렬)   : GT 행들과 파서 행들을 편집거리로 정렬(행 추가/삭제/치환 허용)
#                      → 행이 밀리거나 빠져도 대응됨
#   2단계(셀 정렬)   : 짝지어진 두 행 안에서 셀을 편집거리로 정렬
#                      → 셀 치환 비용 = 1 - 텍스트 유사도, 셀 추가/삭제 = 1
#   행 추가/삭제 비용 = 1(tr) + 그 행의 셀 수(td들)
#   노드 수 = 1(table) + 행 수 + 전체 셀 수
#   TEDS-content = 1 - 편집거리 / max(GT 노드 수, 파서 노드 수)   (0~1, 1이면 완전 동일)
#
# 빈 셀 표기(null / Unnamed: 1 등)는 _normalize_cell에서 빈 칸으로 통일된다.

_EMPTY_CELL_TOKENS = {"", "null", "none", "nan", "n/a"}
_UNNAMED_CELL = re.compile(r"^unnamed:?\s*\d*$", re.IGNORECASE)   # pandas의 이름 없는 열 헤더


def _normalize_cell(text: str) -> str:
    """셀 텍스트 정규화: 앞뒤 공백 제거 + 빈 셀 표기(null/Unnamed 등)를 ''로 통일."""
    t = (text or "").strip()
    if t.lower() in _EMPTY_CELL_TOKENS or _UNNAMED_CELL.match(t):
        return ""
    return t


def _text_similarity(a: str, b: str) -> float:
    """두 셀 텍스트의 문자 단위 유사도(0~1). difflib 기반."""
    if a == "" and b == "":
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def _row_edit_cost(row_a: list, row_b: list) -> float:
    """두 행(셀 리스트) 사이 셀 단위 편집거리. 셀 치환=1-유사도, 셀 추가/삭제=1."""
    na, nb = len(row_a), len(row_b)
    dp = [[0.0] * (nb + 1) for _ in range(na + 1)]
    for i in range(na + 1):
        dp[i][0] = i
    for j in range(nb + 1):
        dp[0][j] = j
    for i in range(1, na + 1):
        for j in range(1, nb + 1):
            sub = dp[i - 1][j - 1] + (1.0 - _text_similarity(row_a[i - 1], row_b[j - 1]))
            dp[i][j] = min(sub, dp[i - 1][j] + 1.0, dp[i][j - 1] + 1.0)
    return dp[na][nb]


def _row_weight(row: list) -> float:
    """행 하나를 통째로 추가/삭제하는 비용 = tr 노드 1 + 셀 수."""
    return 1.0 + len(row)


def _table_edit_distance(rows_a: list, rows_b: list) -> float:
    """행 단위 정렬(치환=행 내부 셀 편집거리, 추가/삭제=행 통째 비용)로 표 편집거리 계산."""
    na, nb = len(rows_a), len(rows_b)
    dp = [[0.0] * (nb + 1) for _ in range(na + 1)]
    for i in range(1, na + 1):
        dp[i][0] = dp[i - 1][0] + _row_weight(rows_a[i - 1])
    for j in range(1, nb + 1):
        dp[0][j] = dp[0][j - 1] + _row_weight(rows_b[j - 1])
    for i in range(1, na + 1):
        for j in range(1, nb + 1):
            sub = dp[i - 1][j - 1] + _row_edit_cost(rows_a[i - 1], rows_b[j - 1])
            dp[i][j] = min(sub, dp[i - 1][j] + _row_weight(rows_a[i - 1]),
                           dp[i][j - 1] + _row_weight(rows_b[j - 1]))
    return dp[na][nb]


def _matrix_nodes(rows: list) -> int:
    """표의 노드 수 = table(1) + 행 수 + 전체 셀 수."""
    return 1 + len(rows) + sum(len(r) for r in rows)


def compute_teds_content(pred_html: str, gt_html: str) -> dict:
    """파서 표(HTML)와 GT 표(HTML)를 비교해 TEDS-content 점수를 계산."""
    gt_rows = table_to_matrix(gt_html)      # 셀은 _normalize_cell 적용된 상태로 반환됨
    pred_rows = table_to_matrix(pred_html)

    dist = _table_edit_distance(gt_rows, pred_rows)
    max_nodes = max(_matrix_nodes(gt_rows), _matrix_nodes(pred_rows))
    teds = 1.0 if max_nodes == 0 else max(0.0, 1.0 - dist / max_nodes)
    return {"teds_content": teds, "edit_distance": dist, "max_nodes": max_nodes}


def average_teds(pred_tables: list, gt_tables: list):
    """
    GT 표와 파서 표를 등장 순서대로 짝지어 TEDS-content 평균을 낸다.
    개수가 다르면 짝지어지는 개수까지만 계산한다. 짝이 없으면 None.
    """
    n = min(len(pred_tables), len(gt_tables))
    if n == 0:
        return None
    scores = [compute_teds_content(pred_tables[i], gt_tables[i])["teds_content"] for i in range(n)]
    return sum(scores) / len(scores)


# ── 셀 F1 (위치/패딩에 영향받지 않는 표 비교) ─────────────────────────────
# TEDS는 표의 "구조"까지 보므로, 엑셀처럼 파서가 전체 시트를 훑어 빈 칸을 null로
# 채워 내보내면 실제 표 밖 패딩 셀 때문에 구조가 어긋나 점수가 왜곡된다. 셀 F1은
# 표를 "비어있지 않은 셀 값들의 모음(다중집합)"으로만 보고, 위치나 빈칸 패딩과 무관하게
# "파서가 옳은 값들을 얼마나 담았는가"를 잰다.
#   precision = 겹친 값 수 / 파서 값 수,  recall = 겹친 값 수 / GT 값 수,  F1 = 2PR/(P+R)
# (구조는 보지 않으므로, 값은 맞는데 자리가 뒤섞인 경우도 잘 나온다. 구조까지 보고
#  싶으면 TEDS를 함께 쓰면 된다.)

def _nonempty_cells(matrices: list) -> Counter:
    """행렬 리스트에서 비어있지 않은 셀 값들을 모아 다중집합(Counter)으로."""
    counter = Counter()
    for matrix in matrices:
        for row in matrix:
            for cell in row:
                v = _normalize_cell(cell)
                if v != "":
                    counter[v] += 1
    return counter


def compute_cell_f1(pred_matrices: list, gt_matrices: list) -> dict:
    """
    파서 표들과 GT 표들을 셀 값 다중집합으로 비교해 precision/recall/F1을 낸다.
    입력은 '행렬(list of list) 리스트'다(표 여러 개를 통째로 넘겨도 됨).
    """
    pred = _nonempty_cells(pred_matrices)
    gt = _nonempty_cells(gt_matrices)
    overlap = sum((pred & gt).values())      # 다중집합 교집합 크기
    n_pred, n_gt = sum(pred.values()), sum(gt.values())

    if n_pred == 0 and n_gt == 0:
        return {"cell_f1": 1.0, "cell_precision": 1.0, "cell_recall": 1.0,
                "n_pred_cells": 0, "n_gt_cells": 0}
    precision = overlap / n_pred if n_pred else 0.0
    recall = overlap / n_gt if n_gt else 0.0
    f1 = (2 * precision * recall / (precision + recall)) if (precision + recall) else 0.0
    return {"cell_f1": f1, "cell_precision": precision, "cell_recall": recall,
            "n_pred_cells": n_pred, "n_gt_cells": n_gt}


# ══════════════════════════════════════════════════════════════════════
# 3. 진단 도구 (숫자가 왜 그렇게 나왔는지 눈으로 확인)
# ══════════════════════════════════════════════════════════════════════
def _align_tokens(pred_tokens: list, gt_tokens: list) -> list:
    """토큰 정렬 결과를 [(op, gt_token, pred_token), ...]로 반환.
    op: match/sub/del(GT에만 있음=누락)/ins(PRED에만 있음=잡음)."""
    n, m = len(pred_tokens), len(gt_tokens)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        dp[i][0] = i
    for j in range(m + 1):
        dp[0][j] = j
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if pred_tokens[i - 1] == gt_tokens[j - 1]:
                dp[i][j] = dp[i - 1][j - 1]
            else:
                dp[i][j] = 1 + min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1])

    i, j, ops = n, m, []
    while i > 0 or j > 0:
        if i > 0 and j > 0 and pred_tokens[i - 1] == gt_tokens[j - 1] and dp[i][j] == dp[i - 1][j - 1]:
            ops.append(("match", gt_tokens[j - 1], pred_tokens[i - 1])); i -= 1; j -= 1
        elif i > 0 and j > 0 and dp[i][j] == dp[i - 1][j - 1] + 1:
            ops.append(("sub", gt_tokens[j - 1], pred_tokens[i - 1])); i -= 1; j -= 1
        elif i > 0 and dp[i][j] == dp[i - 1][j] + 1:
            ops.append(("ins", None, pred_tokens[i - 1])); i -= 1
        else:
            ops.append(("del", gt_tokens[j - 1], None)); j -= 1
    ops.reverse()
    return ops


def print_wer_diff(pred_text: str, gt_text: str, tokenizer: str = "eojeol",
                   apply_normalization: bool = True, normalize_kwargs: dict = None,
                   max_lines: int = 300) -> None:
    """단어 단위로 치환/누락/삽입을 한 줄씩 출력한다(일치는 생략)."""
    if apply_normalization:
        opts = normalize_kwargs or {}
        pred_text = normalize_for_wer(pred_text, **opts)
        gt_text = normalize_for_wer(gt_text, **opts)
    pred_tokens = _tokenize(pred_text, tokenizer)
    gt_tokens = _tokenize(gt_text, tokenizer)
    alignment = _align_tokens(pred_tokens, gt_tokens)

    S = sum(1 for op, *_ in alignment if op == "sub")
    D = sum(1 for op, *_ in alignment if op == "del")
    I = sum(1 for op, *_ in alignment if op == "ins")
    N = len(gt_tokens)
    wer = (S + D + I) / N if N > 0 else 0.0
    print(f"WER={wer:.4f}  S={S} D={D} I={I} / N={N}")
    print("-" * 60)

    shown = 0
    for op, gt_tok, pred_tok in alignment:
        if op == "match":
            continue
        if op == "sub":
            print(f"[치환 S]  GT: {gt_tok}   PRED: {pred_tok}")
        elif op == "del":
            print(f"[누락 D]  GT: {gt_tok}   PRED: (없음)")
        elif op == "ins":
            print(f"[삽입 I]  GT: (없음)   PRED: {pred_tok}")
        shown += 1
        if shown >= max_lines:
            print(f"... (이하 생략, max_lines={max_lines})")
            break
    if shown == 0:
        print("(치환/누락/삽입 없음 - GT와 완전히 일치)")


def print_paragraph_diff(pred_text: str, gt_text: str, max_lines: int = 200) -> None:
    """
    문단(줄) 단위로 GT/PRED가 어디서 갈라지는지 보여준다(정규화 전 원문 기준).
    [=] 동일 / [GT≠][PRED≠] 같은 자리에서 다름 / [GT-] GT에만 / [PRED+] PRED에만.
    """
    import difflib
    gt_paras = [p for p in gt_text.split("\n") if p.strip()]
    pred_paras = [p for p in pred_text.split("\n") if p.strip()]
    matcher = difflib.SequenceMatcher(None, gt_paras, pred_paras, autojunk=False)

    shown = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if shown >= max_lines:
            print(f"... (이하 생략, max_lines={max_lines})")
            break
        if tag == "equal":
            for k in range(i1, i2):
                print(f"[=] {gt_paras[k]}"); shown += 1
        elif tag == "replace":
            gt_chunk, pred_chunk = gt_paras[i1:i2], pred_paras[j1:j2]
            for k in range(max(len(gt_chunk), len(pred_chunk))):
                if k < len(gt_chunk):
                    print(f"[GT≠]   {gt_chunk[k]}"); shown += 1
                if k < len(pred_chunk):
                    print(f"[PRED≠] {pred_chunk[k]}"); shown += 1
        elif tag == "delete":
            for k in range(i1, i2):
                print(f"[GT-]   {gt_paras[k]}"); shown += 1
        elif tag == "insert":
            for k in range(j1, j2):
                print(f"[PRED+] {pred_paras[k]}"); shown += 1


def table_to_matrix(table_html: str) -> list:
    """<table> HTML을 파서가 인식한 그대로 행렬(list of list)로 반환."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(table_html, "html.parser")
    table_tag = soup.find("table")
    if table_tag is None:
        return []
    matrix = []
    for tr in table_tag.find_all("tr", recursive=True):
        if tr.find_parent("table") is not table_tag:
            continue
        matrix.append([_normalize_cell(c.get_text()) for c in tr.find_all(["td", "th"], recursive=False)])
    return matrix


def compare_tables_side_by_side(pred_html: str, gt_html: str):
    """
    GT/PRED 표를 같은 행·열 위치끼리 나란히 놓고 셀 텍스트 유사도를 보여주는
    pandas DataFrame(row, col, gt_cell, pred_cell, similarity)을 반환한다.
    개수가 다른 칸은 None으로 채워 "통째로 사라짐/추가됨"도 드러난다.
    """
    import pandas as pd
    gt_matrix = table_to_matrix(gt_html)
    pred_matrix = table_to_matrix(pred_html)

    rows = []
    for r in range(max(len(gt_matrix), len(pred_matrix))):
        gt_row = gt_matrix[r] if r < len(gt_matrix) else []
        pred_row = pred_matrix[r] if r < len(pred_matrix) else []
        for c in range(max(len(gt_row), len(pred_row))):
            gt_cell = gt_row[c] if c < len(gt_row) else None
            pred_cell = pred_row[c] if c < len(pred_row) else None
            sim = (_text_similarity(gt_cell or "", pred_cell or "")
                   if (gt_cell is not None or pred_cell is not None) else None)
            rows.append({
                "row": r, "col": c, "gt_cell": gt_cell, "pred_cell": pred_cell,
                "similarity": round(sim, 3) if sim is not None else None,
            })
    return pd.DataFrame(rows)