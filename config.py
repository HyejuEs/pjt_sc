# -*- coding: utf-8 -*-
"""
config.py
=========
경로와 옵션 설정 파일. 실제 사용 시 DATA_DIR 경로만 본인 환경에 맞게 바꾸면 된다.

폴더 구조
---------
gt/before/after 파일을 폴더 하나(DATA_DIR)에 전부 넣는다. 파일명으로 역할을 구분한다.

    data/
    ├── 1-1_gt.docx        ← 정답(원본 문서). 확장자는 .docx / .pptx / .xlsx
    ├── 1-1_before.txt     ← 가이드라인 적용 "전" 파서 결과
    ├── 1-1_after.txt      ← 가이드라인 적용 "후" 파서 결과
    ├── 1-2_gt.pptx
    ├── 1-2_after.txt      ← before 없이 after만 있어도 됨(after만 비교)
    ├── 1-3_gt.xlsx
    └── 1-3_before.txt     ← after 없이 before만 있어도 됨(before만 비교)

파일명 규칙
-----------
    {문서id}_gt.(docx|pptx|xlsx)
    {문서id}_before.txt
    {문서id}_after.txt

  * {문서id}는 자유롭게 정하면 된다(예: 1-1, doc_0001, card_fee 등).
    같은 문서의 gt/before/after는 이 {문서id}가 같아야 매칭된다.
  * gt는 반드시 있어야 한다(정답이 없으면 비교 불가). before/after는 둘 중
    하나만 있어도 되고, 있는 것만 비교한다.
  * before/after txt 안에는 원문 순서대로 텍스트(#, **, ~~, - 등 Markdown 포함)와
    표(<table>...</table> HTML)가 섞여 있다고 가정한다.
"""

from pathlib import Path

# ── 경로 (여기만 환경에 맞게 수정) ─────────────────────────────────────
DATA_DIR = Path("./testset")       # gt/before/after 파일을 모두 넣는 폴더

# ── 전처리 프로파일(케이스) ────────────────────────────────────────────
# 텍스트 지표는 WER 하나만 쓴다. 대신 "상황별 전처리 케이스"를 프로파일로 만들어
# 골라 쓴다. 각 프로파일은 토큰 단위(어절/형태소)와 normalize_for_wer 옵션을 묶는다.
# 아래 EVAL_PROFILE 값만 바꾸면 run_eval이 그 케이스로 WER을 계산한다.
#
#   "default"        : 기본. 제목/볼드/취소선/리스트를 채점, 표기 통일. 어절 단위.
#   "spacing_robust" : 띄어쓰기/줄바꿈 밀림 대응. 형태소(kiwipiepy) 단위로 재분절해
#                      띄어쓰기 차이를 흡수한다("전월 실적"↔"전월실적"에 휘둘리지 않음).
#                      ※ 어절 단위로는 공백을 지우면 토큰이 1개가 돼 WER이 깨지므로,
#                        띄어쓰기 밀림을 WER 하나로 흡수하려면 형태소 단위가 사실상 필수다.
#                        pip install kiwipiepy 필요.
#   "content_only"   : 서식은 채점하지 않고 순수 내용만. 헤더/볼드/취소선 기호 제거.
#
# 특정 패턴만 교정하고 싶으면 "normalize"에 custom_regex_subs=[(정규식, 치환)]를 넣는다.
# 예) 숫자 사이 공백 제거: {"custom_regex_subs": [(r"(?<=\d)\s+(?=\d)", "")]}
EVAL_PROFILE = "default"

PROFILES = {
    "default":        {"tokenizer": "eojeol",   "normalize": {}},
    "spacing_robust": {"tokenizer": "morpheme", "normalize": {}},
    "content_only":   {"tokenizer": "eojeol",
                       "normalize": {"strip_headers": True, "strip_bold": True, "strip_strike": True}},
}


def active_profile() -> dict:
    return PROFILES[EVAL_PROFILE]


# (구버전 호환용) 단독으로 토크나이저만 참조하던 코드를 위해 남겨둠.
WER_TOKENIZER = PROFILES[EVAL_PROFILE]["tokenizer"]



# ── 결과 출력 경로 ────────────────────────────────────────────────────
OUTPUT_DIR = Path("./output")
OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

REPORT_PER_DOC_CSV = OUTPUT_DIR / "report_per_document.csv"
REPORT_SUMMARY_CSV = OUTPUT_DIR / "report_summary.csv"
