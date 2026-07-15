# -*- coding: utf-8 -*-
"""
normalize.py
============
WER을 계산하기 직전에 텍스트에 적용하는 전처리(정규화) 모듈.

목적
----
WER은 "토큰(단어)이 정확히 같은가"만 보는 지표라, 내용은 같은데 표기만 다른 경우
(전각/반각, 곡선/직선 따옴표, 공백 개수 등)까지 전부 오류로 세어버린다. 이런
"내용과 무관한 표기 차이"를 먼저 없애서, WER이 실제 내용 차이만 반영하도록 만든다.

공정 비교 원칙 (이 모듈을 이해하는 핵심)
---------------------------------------
서식(제목/볼드/취소선/리스트)은 GT와 파서 출력이 "같은 표기 체계"일 때만 공정하게
비교된다. 그래서 각 서식은 아래 둘 중 하나로만 처리한다.

  (A) 채점한다  : GT 추출기(gt_extract.py)가 그 서식을 Markdown 기호로 복원하고,
                 여기서는 그 기호를 "지우지 않고 통일만" 한다. → 파서가 서식을
                 살렸는지/놓쳤는지가 WER에 그대로 반영된다.
  (B) 무시한다  : GT/파서 양쪽에서 그 기호를 똑같이 "제거"한다. → 점수에 영향 없음.

절대 하면 안 되는 조합:
  - GT는 복원했는데 여기서 기호를 지움  → 채점이 사라짐(복원한 의미가 없음)
  - GT는 복원 안 했는데 여기서 기호를 남김 → 파서 기호가 매번 삽입 오류로 잡힘(불공정)

기본 처리 항목
--------------
  채점(기호 유지·통일):  제목(#), 볼드(**), 취소선(~~), 리스트 불릿(→ "- "로 통일)
  무시(기호 제거):        인용(>), 구분선(---), 링크/이미지, 인라인 코드(`)
  표기 통일:              유니코드(NFKC), 공백/줄바꿈, 따옴표/대시, 영문 소문자
  기본 꺼둠:              문장부호 제거 (숫자 안 콤마까지 지워 실제 오류를 가릴 수 있어서)

각 항목은 normalize_for_wer()의 인자로 개별 on/off 할 수 있다.
"""

import re
import unicodedata

# ── Markdown 기호 탐지용 정규식 ──────────────────────────────────────────
_MD_HEADING = re.compile(r"^\s{0,3}#{1,6}\s*", re.MULTILINE)   # "# ", "## " ...
_MD_BOLD_MARK = re.compile(r"\*\*")                            # 볼드 표시 **
_MD_STRIKE_MARK = re.compile(r"~~")                            # 취소선 표시 ~~
_MD_BOLD_INNER = re.compile(r"\*\*\s*(.+?)\s*\*\*", re.DOTALL) # **  텍스트  ** 안쪽 공백 정리용
_MD_STRIKE_INNER = re.compile(r"~~\s*(.+?)\s*~~", re.DOTALL)   # ~~  텍스트  ~~ 안쪽 공백 정리용
_MD_STRIKE_WHOLE = re.compile(r"~~.+?~~", re.DOTALL)           # ~~텍스트~~ 통째로(내용까지 삭제할 때)
_MD_BLOCKQUOTE = re.compile(r"^\s*>\s?", re.MULTILINE)         # 인용 >
_MD_HR = re.compile(r"^\s*([-*_]\s?){3,}\s*$", re.MULTILINE)   # 구분선 ---, ***
# 대시(-)만 있는 줄. PPT 파서가 슬라이드 경계를 "-" 한 줄로 넣는 경우가 있는데,
# GT엔 그런 줄이 없어 그 "-"가 오류로 잡힌다. 내용이 없는 대시 줄은 양쪽에서 제거한다.
# ("- 내용"처럼 뒤에 글자가 있는 목록 항목은 대시 뒤에 내용이 있으므로 걸리지 않는다.)
_SEP_LINE = re.compile(r"^[ \t]*[-–—]+[ \t]*$", re.MULTILINE)
_MD_INLINE_CODE = re.compile(r"`([^`]*)`")                     # `코드` → 코드
_MD_LINK_IMAGE = re.compile(r"!?\[([^\]]*)\]\([^)]*\)")        # [텍스트](url), ![alt](url) → 텍스트

# 리스트 불릿(순서 없는 목록)으로 쓰이는 여러 기호를 한 종류로 통일하기 위한 패턴.
# 줄 맨 앞의 기호 + 뒤따르는 공백을 "- "로 바꾼다. 뒤에 공백이 반드시 있어야 하므로
# 구분선("---")이나 곱셈/강조 기호와 헷갈리지 않는다.
# 순서 있는 목록(1. 2. / ① / 가.)의 번호는 "내용"이므로 건드리지 않는다.
_LIST_BULLET = re.compile(r"^[ \t]*[-*+•·▪◦‣・]\s+", re.MULTILINE)

# 따옴표/대시 표기 통일 (곡선→직선, 엔/엠 대시→하이픈)
_QUOTE_DASH_MAP = str.maketrans({
    "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
    "\u2013": "-", "\u2014": "-",
})

# 눈에 안 보이는 제로폭/제어 문자. 화면상 아무것도 없지만 코드값이 있어, 같은
# 글자인데 토큰이 안 맞게 만드는 원인이 된다. (ZWSP, ZWNJ, ZWJ, BOM/ZWNBSP,
# soft hyphen 등) 전각/반각을 맞추는 NFKC로는 지워지지 않아 따로 제거한다.
_ZERO_WIDTH = re.compile(r"[\u200b\u200c\u200d\u2060\ufeff\u00ad]")


def normalize_for_wer(
    text: str,
    unify_unicode: bool = True,
    strip_zero_width: bool = True,      # 눈에 안 보이는 제로폭/제어 문자 제거
    normalize_whitespace: bool = True,
    strip_headers: bool = False,        # 채점 대상: 기본은 "#" 유지
    strip_bold: bool = False,           # 채점 대상: 기본은 "**" 유지
    strip_strike: bool = False,         # 채점 대상: 기본은 "~~" 유지
    drop_strike_content: bool = False,  # True면 취소선 "내용까지" 삭제(취소선=삭제된 글로 취급)
    normalize_list_bullet: bool = True, # 채점 대상: 여러 불릿 기호를 "- "로 통일
    strip_blockquote: bool = True,      # 무시: 인용 기호 제거
    strip_hr: bool = True,              # 무시: 구분선(--- 등) 제거
    strip_separator_line: bool = True,  # 무시: 대시(-)만 있는 줄(PPT 슬라이드 구분 등) 제거
    strip_inline_code: bool = True,     # 무시: 백틱만 제거(안쪽 텍스트는 유지)
    strip_link_image: bool = True,      # 무시: 링크/이미지 문법 제거(표시 텍스트는 유지)
    unify_quotes_dashes: bool = True,
    lower_latin: bool = True,
    strip_punctuation: bool = False,    # 위험: 숫자 콤마까지 지울 수 있어 기본 꺼둠
    custom_regex_subs: list = None,     # 특정 상황 교정용 (pattern, replacement) 목록
) -> str:
    """
    WER 계산 전 전처리 파이프라인. 위 공정 비교 원칙에 맞춘 기본값을 사용한다.
    GT 텍스트와 파서 텍스트에 "동일하게" 적용해야 공정한 비교가 된다.
    """
    if unify_unicode:
        # NFKC: 전각/반각 통일 + 한글 자모 조합 방식 통일까지 한 번에 처리
        text = unicodedata.normalize("NFKC", text)

    if strip_zero_width:
        text = _ZERO_WIDTH.sub("", text)

    # 무시할 서식(링크/코드/인용/구분선)은 먼저 정리
    if strip_link_image:
        text = _MD_LINK_IMAGE.sub(r"\1", text)
    if strip_inline_code:
        text = _MD_INLINE_CODE.sub(r"\1", text)
    if strip_blockquote:
        text = _MD_BLOCKQUOTE.sub("", text)
    if strip_hr:
        text = _MD_HR.sub("", text)
    if strip_separator_line:
        text = _SEP_LINE.sub("", text)

    # 제목(#)
    if strip_headers:
        text = _MD_HEADING.sub("", text)

    # 볼드(**): 지우거나(무시), 안쪽 공백만 정리(채점 - 기호는 유지)
    if strip_bold:
        text = _MD_BOLD_MARK.sub("", text)
    else:
        text = _MD_BOLD_INNER.sub(r"**\1**", text)

    # 취소선(~~): 내용까지 삭제 / 기호만 삭제 / 안쪽 공백만 정리(채점)
    if drop_strike_content:
        text = _MD_STRIKE_WHOLE.sub("", text)
    elif strip_strike:
        text = _MD_STRIKE_MARK.sub("", text)
    else:
        text = _MD_STRIKE_INNER.sub(r"~~\1~~", text)

    # 리스트 불릿: 여러 기호를 "- "로 통일(채점). GT/파서 양쪽에서 같은 기호가 되도록.
    if normalize_list_bullet:
        text = _LIST_BULLET.sub("- ", text)

    if unify_quotes_dashes:
        text = text.translate(_QUOTE_DASH_MAP)

    if lower_latin:
        text = text.lower()

    # 특정 상황을 겨냥한 사용자 정의 regex 치환(프로파일에서 주입).
    # 예: [(r"(?<=\d)\s+(?=\d)", "")]  # 숫자 사이 공백 제거
    if custom_regex_subs:
        for pattern, repl in custom_regex_subs:
            text = re.sub(pattern, repl, text)

    if strip_punctuation:
        # 마침표/쉼표/물음표 등 문장부호 제거. #, *, ~, - 같은 서식 기호는 건드리지 않는다.
        text = re.sub(r"[.,!?;:\"'()\[\]{}]", "", text)

    if normalize_whitespace:
        # 탭/줄바꿈/연속 공백 → 공백 1칸, 앞뒤 공백 제거.
        # 어절 토큰화(text.split())는 공백류를 이미 동일하게 취급하므로 토큰 결과를
        # 바꾸지 않고, 빈 토큰 같은 잡음만 없앤다.
        text = re.sub(r"\s+", " ", text).strip()

    return text