# -*- coding: utf-8 -*-
"""
slide_eval.py
=============
PPT를 "슬라이드 단위"로 맞춰서 WER을 내는 코드.

왜 필요한가
-----------
파서는 새 슬라이드의 시작을 "## 제목" 같은 마크다운 헤더로 표시한다(원본 PPT는
슬라이드 사이가 그냥 나뉘어 있음). 문서를 통째로 한 줄로 이어 WER을 내면, 슬라이드
하나가 빠지거나 순서가 바뀌었을 때 그 뒤가 전부 밀려 WER이 무너진다.

그래서 여기서는
  1) GT는 슬라이드별 텍스트로(gt_extract.extract_pptx_slides),
     파서는 "## 헤더" 기준으로 섹션을 나눠 각각 '슬라이드 조각 리스트'를 만들고,
  2) 두 리스트를 슬라이드 단위로 정렬(Needleman-Wunsch)한다.
     - GT 슬라이드 i ↔ 파서 섹션 j 를 맞추는 비용 = 그 둘의 단어 편집거리
     - 한쪽 슬라이드를 못 맞추면(빠짐/추가) 그 슬라이드 단어 수만큼 비용
     이렇게 하면 슬라이드가 하나 빠져도 그 슬라이드만 손해 보고 뒤는 안 밀린다.
  3) 맞춰진 슬라이드 쌍마다 WER을 내고, 전체를 합쳐 문서 WER도 낸다.

split_sections()의 헤더 기준(기본 "#로 시작하는 줄")과 정렬은 파서 규칙에 맞춰 조정할
수 있다. 파서에 헤더가 없으면(섹션이 1개 이하) 슬라이드 정렬 없이 통짜 WER로 넘어간다.
"""

import re

from normalize import normalize_for_wer
from metrics import _tokenize, _levenshtein_sdi

# 슬라이드 시작 표시. 이 파서는 "## " (h2)만 슬라이드 경계로 쓴다.
# ("#"이나 "### "는 슬라이드 안 소제목일 수 있어 경계로 보지 않는다.)
# 다른 파서라 규칙이 다르면 이 정규식만 바꾸면 된다.
_HEADER_LINE = re.compile(r"^\s*##\s")


def split_sections(text: str) -> list:
    """
    파서 텍스트를 헤더(#로 시작하는 줄) 기준으로 섹션(슬라이드 조각)으로 나눈다.
    헤더 줄이 나오면 새 섹션이 시작된다. 첫 헤더 앞의 내용은 섹션 0이 된다.
    """
    sections, cur = [], []
    for line in text.split("\n"):
        if _HEADER_LINE.match(line) and cur:
            sections.append("\n".join(cur))
            cur = [line]
        else:
            cur.append(line)
    if cur:
        sections.append("\n".join(cur))
    return [s for s in sections if s.strip()]


def _edit_total(pred_tokens, gt_tokens) -> int:
    """두 토큰열의 편집거리(S+D+I)."""
    S, D, I = _levenshtein_sdi(pred_tokens, gt_tokens)
    return S + D + I


def align_slides_wer(gt_segments: list, pred_segments: list,
                     tokenizer: str = "eojeol", normalize_kwargs: dict = None) -> dict:
    """
    GT 슬라이드 조각들과 파서 섹션들을 슬라이드 단위로 정렬하고 WER을 낸다.

    각 슬라이드 쌍을 토큰화하기 전에, 파서 조각의 줄바꿈을 그 GT 슬라이드에
    국소적으로 맞춘다(relax_linebreaks). 슬라이드 경계로 이미 정렬돼 있으므로
    이 줄바꿈 보정이 슬라이드 밖으로 번져 뒤를 밀지 않는다. 슬라이드 내용이
    (줄바꿈 차이만 빼면) 동일하면 그 슬라이드는 WER 0으로 통과된다.

    반환:
      {"wer": 문서 전체 WER,
       "pairs": [{"gt_idx","pred_idx","wer","S","D","I","N","kind"}, ...]}
        kind: "match"(둘 다 있음) / "gt_only"(GT 슬라이드가 파서에서 빠짐)
              / "pred_only"(파서에만 있는 슬라이드)
    """
    from wer_preprocess import relax_linebreaks

    nk = normalize_kwargs or {}
    gt_tok = [_tokenize(normalize_for_wer(s, **nk), tokenizer) for s in gt_segments]
    # 파서 섹션은 대응 GT 슬라이드를 아직 모르므로, 정렬 비용 계산 단계에서는
    # 원본 줄바꿈 그대로 토큰화한다(줄바꿈은 어절 토큰에 큰 영향 없음).
    pred_tok = [_tokenize(normalize_for_wer(s, **nk), tokenizer) for s in pred_segments]
    n, m = len(gt_tok), len(pred_tok)

    # Needleman-Wunsch: 슬라이드를 원소로 보고 정렬
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        dp[i][0] = dp[i - 1][0] + len(gt_tok[i - 1])       # GT 슬라이드 전부 삭제
    for j in range(1, m + 1):
        dp[0][j] = dp[0][j - 1] + len(pred_tok[j - 1])     # 파서 슬라이드 전부 삽입
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            match = dp[i - 1][j - 1] + _edit_total(pred_tok[j - 1], gt_tok[i - 1])
            skip_gt = dp[i - 1][j] + len(gt_tok[i - 1])
            skip_pred = dp[i][j - 1] + len(pred_tok[j - 1])
            dp[i][j] = min(match, skip_gt, skip_pred)

    # 역추적으로 슬라이드 쌍 복원
    pairs = []
    i, j = n, m
    while i > 0 or j > 0:
        if i > 0 and j > 0 and dp[i][j] == dp[i - 1][j - 1] + _edit_total(pred_tok[j - 1], gt_tok[i - 1]):
            pairs.append((i - 1, j - 1)); i -= 1; j -= 1
        elif i > 0 and dp[i][j] == dp[i - 1][j] + len(gt_tok[i - 1]):
            pairs.append((i - 1, None)); i -= 1
        else:
            pairs.append((None, j - 1)); j -= 1
    pairs.reverse()

    # 쌍별 WER + 전체 집계 (매칭된 쌍은 줄바꿈 유연화 후 다시 토큰화)
    results, tot_edit, tot_n = [], 0, 0
    for gi, pj in pairs:
        if gi is not None and pj is not None:
            # 이 슬라이드 쌍에 한해 파서 줄바꿈을 GT 슬라이드에 맞춘 뒤 재토큰화
            relaxed = relax_linebreaks(pred_segments[pj], gt_segments[gi])
            p_tok = _tokenize(normalize_for_wer(relaxed, **nk), tokenizer)
            S, D, I = _levenshtein_sdi(p_tok, gt_tok[gi])
            N = len(gt_tok[gi]); kind = "match"
        elif gi is not None:
            N = len(gt_tok[gi]); S = 0; D = N; I = 0; kind = "gt_only"     # 파서가 슬라이드 통째로 놓침
        else:
            N = 0; S = 0; D = 0; I = len(pred_tok[pj]); kind = "pred_only" # 파서가 만든 잉여 슬라이드
        wer = (S + D + I) / N if N > 0 else (1.0 if I > 0 else 0.0)
        results.append({"gt_idx": gi, "pred_idx": pj, "wer": wer,
                        "S": S, "D": D, "I": I, "N": N, "kind": kind})
        tot_edit += S + D + I; tot_n += N

    doc_wer = tot_edit / tot_n if tot_n > 0 else 0.0
    return {"wer": doc_wer, "pairs": results}


def pptx_slide_wer(gt_slides: list, pred_text: str,
                   tokenizer: str = "eojeol", normalize_kwargs: dict = None) -> dict:
    """
    GT 슬라이드 리스트 + 파서 텍스트(표 제거된 본문)를 받아 슬라이드 정렬 WER을 낸다.
    파서에 헤더(##)가 거의 없어 섹션이 1개 이하이면, 슬라이드 정렬을 포기하고
    통짜 WER(슬라이드 경계 무시)로 계산한다.
    """
    pred_segments = split_sections(pred_text)
    if len(pred_segments) <= 1:
        # 헤더가 없어 섹션을 못 나눔 → GT도 통짜로 이어붙여 일반 WER
        from metrics import compute_wer
        gt_joined = "\n".join(gt_slides)
        r = compute_wer(pred_text, gt_joined, tokenizer=tokenizer, normalize_kwargs=normalize_kwargs)
        return {"wer": r["wer"], "pairs": None, "aligned": False}

    res = align_slides_wer(gt_slides, pred_segments, tokenizer, normalize_kwargs)
    res["aligned"] = True
    return res