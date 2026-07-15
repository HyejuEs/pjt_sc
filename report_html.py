# -*- coding: utf-8 -*-
"""
report_html.py
===============
GT / Before / After 세 문서를 한 페이지에서 색으로 비교하는 HTML 리포트 생성 모듈.
콘솔 diff(print_wer_diff 등)와 달리 문서 전체를 시각적으로 보여준다.

  1) 본문 텍스트를 GT / Before / After 3열 표로, 문단 단위 정렬해 전체 표시
  2) 문단 안에서 다른 단어를 색으로 하이라이트
     (치환: 노란색 / 누락(GT에만): 빨간 취소선 / 삽입(PRED에만): 초록 밑줄)
  3) 하이라이트 구간의 공백류 문자를 눈에 보이는 기호(·, ⬚, →, ↵)로 드러내
     "겉보기엔 같은데 다르다"의 원인(보이지 않는 문자)을 바로 확인
  4) 표는 GT/Before/After를 같은 행·열 위치에 겹쳐 실제 <table>로 렌더링하고
     셀 텍스트가 다르면 배경색으로 표시

보통은 run_eval.diagnose(doc_id, make_html=True)로 호출한다. 직접 쓰려면:
    from report_html import generate_report
    generate_report(gt_text, gt_tables, before_text, before_tables,
                    after_text, after_tables, wer_before, wer_after,
                    teds_before, teds_after, output_path="report.html")
"""

import re
import html as html_lib
import difflib
from pathlib import Path

from metrics import table_to_matrix, _text_similarity

# ── 공백류 문자를 "눈에 보이는 기호"로 바꾸는 매핑 ──────────────────────────
# 하이라이트(치환/누락/삽입) 구간 "안에" 있는 공백류만 이렇게 바꿉니다.
# (하이라이트 밖의, 즉 정상적으로 일치하는 부분의 공백은 원래 그대로 보여줍니다 -
#  그래야 레이아웃이 깨지지 않고, 정말 "문제가 된 구간"의 공백만 도드라져 보입니다)
_VISIBLE_WS_MAP = {
    " ": "·",       # 일반 반각 스페이스
    "\u3000": "⬚",  # 전각(全角) 스페이스 - 한국어 문서에서 자간 맞추기용으로 자주 씀
    "\u00a0": "␣",  # 줄바꿈 방지 공백(non-breaking space)
    "\t": "→",
    "\n": "↵\n",
}


def _reveal_whitespace(token: str) -> str:
    """토큰이 공백류 문자로만 이루어져 있으면, 각 문자를 눈에 보이는 기호로 변환."""
    if token != "" and token.strip() == "":
        return "".join(_VISIBLE_WS_MAP.get(ch, ch) for ch in token)
    return token


def _tokenize_ws(text: str) -> list:
    """단어와 공백(연속 공백류 포함)을 각각 토큰으로 분리 (재조립 가능하도록)."""
    return re.findall(r"\S+|\s+", text)


def _diff_pair_html(gt_text: str, pred_text: str):
    """
    문단 하나의 GT/PRED 텍스트를 토큰 단위(re.findall(\\S+|\\s+))로 정렬 비교해서,
    (gt_html, pred_html) 튜플로 반환. 하이라이트:
      - 치환(replace) : GT쪽은 노란 배경 + 취소선, PRED쪽은 노란 배경
      - GT에만 있음(delete, 누락) : 빨간 배경 + 취소선
      - PRED에만 있음(insert, 삽입) : 초록 배경 + 밑줄
    하이라이트 구간 안의 공백류 문자는 _reveal_whitespace로 눈에 보이게 바꿉니다.
    """
    gt_tokens = _tokenize_ws(gt_text)
    pred_tokens = _tokenize_ws(pred_text)
    matcher = difflib.SequenceMatcher(None, gt_tokens, pred_tokens, autojunk=False)

    gt_parts, pred_parts = [], []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            chunk = html_lib.escape("".join(gt_tokens[i1:i2]))
            gt_parts.append(chunk)
            pred_parts.append(chunk)
        elif tag == "replace":
            gt_chunk = "".join(_reveal_whitespace(t) for t in gt_tokens[i1:i2])
            pred_chunk = "".join(_reveal_whitespace(t) for t in pred_tokens[j1:j2])
            gt_parts.append(f'<span class="d-sub">{html_lib.escape(gt_chunk)}</span>')
            pred_parts.append(f'<span class="d-sub">{html_lib.escape(pred_chunk)}</span>')
        elif tag == "delete":
            gt_chunk = "".join(_reveal_whitespace(t) for t in gt_tokens[i1:i2])
            gt_parts.append(f'<span class="d-del">{html_lib.escape(gt_chunk)}</span>')
        elif tag == "insert":
            pred_chunk = "".join(_reveal_whitespace(t) for t in pred_tokens[j1:j2])
            pred_parts.append(f'<span class="d-ins">{html_lib.escape(pred_chunk)}</span>')

    return "".join(gt_parts), "".join(pred_parts)


def _align_pred_to_gt(gt_paragraphs: list, pred_paragraphs: list):
    """
    gt_paragraphs 순서를 기준으로, pred_paragraphs를 정렬해서
    [(gt_idx_or_None, pred_idx_or_None, anchor), ...] 리스트로 반환.
    anchor: gt_idx가 None(=GT에는 없고 PRED에만 있는 문단)일 때, "GT의 몇 번째
    문단 앞자리에 끼워 넣어야 하는지"를 나타내는 위치 값 (3-way 병합에 사용).
    """
    matcher = difflib.SequenceMatcher(None, gt_paragraphs, pred_paragraphs, autojunk=False)
    pairs = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                pairs.append((i1 + k, j1 + k, None))
        elif tag == "replace":
            n = max(i2 - i1, j2 - j1)
            for k in range(n):
                gi = i1 + k if k < (i2 - i1) else None
                pj = j1 + k if k < (j2 - j1) else None
                pairs.append((gi, pj, i2 if gi is None else None))
        elif tag == "delete":
            for k in range(i2 - i1):
                pairs.append((i1 + k, None, None))
        elif tag == "insert":
            for k in range(j2 - j1):
                pairs.append((None, j1 + k, i1))
    return pairs


def _build_three_way_rows(gt_paragraphs: list, before_paragraphs: list, after_paragraphs: list) -> list:
    """
    GT를 기준 축으로 삼아 before/after 문단을 정렬한 뒤,
    [{"gt": str_or_None, "before": str_or_None, "after": str_or_None}, ...] 형태의
    행(row) 리스트로 병합합니다 (문서 순서 그대로, 생략 없이 전체).
    """
    before_pairs = _align_pred_to_gt(gt_paragraphs, before_paragraphs)
    after_pairs = _align_pred_to_gt(gt_paragraphs, after_paragraphs)

    before_at = {gi: before_paragraphs[pj] for gi, pj, _ in before_pairs if gi is not None and pj is not None}
    after_at = {gi: after_paragraphs[pj] for gi, pj, _ in after_pairs if gi is not None and pj is not None}

    before_extra = [(anchor, before_paragraphs[pj]) for gi, pj, anchor in before_pairs if gi is None and pj is not None]
    after_extra = [(anchor, after_paragraphs[pj]) for gi, pj, anchor in after_pairs if gi is None and pj is not None]
    before_extra.sort(key=lambda x: x[0])
    after_extra.sort(key=lambda x: x[0])

    rows = []
    bi, ai = 0, 0
    n = len(gt_paragraphs)
    for gt_idx in range(n + 1):  # n까지 돌아서 맨 끝(trailing insert)도 처리
        while bi < len(before_extra) and before_extra[bi][0] == gt_idx:
            rows.append({"gt": None, "before": before_extra[bi][1], "after": None})
            bi += 1
        while ai < len(after_extra) and after_extra[ai][0] == gt_idx:
            rows.append({"gt": None, "before": None, "after": after_extra[ai][1]})
            ai += 1
        if gt_idx < n:
            rows.append({
                "gt": gt_paragraphs[gt_idx],
                "before": before_at.get(gt_idx),
                "after": after_at.get(gt_idx),
            })
    return rows


def _slide_divider_row(cols: int, label: str) -> str:
    """슬라이드/섹션 경계를 표 안에서 시각적으로 구분하는 행."""
    return (f'<tr class="slide-divider"><td colspan="{cols}">'
            f'━━ {html_lib.escape(label)} ━━</td></tr>')


def _wrap_dup(html_str, raw_text, dup_norm_set):
    """raw_text가 표중복으로 감지된 라인이면 파란 배경+뱃지로 표시(삭제 아님).
    삽입(초록 d-ins)으로 잡혔더라도 중복이면 파란색만 보이게 d-ins를 벗겨낸다."""
    if not dup_norm_set or raw_text is None:
        return html_str
    from wer_preprocess import _norm_line
    if _norm_line(raw_text) not in dup_norm_set:
        return html_str
    # 초록 삽입 표시가 씌워져 있으면 제거하고 파란색만 적용
    inner = html_str
    if 'class="d-ins"' in inner:
        # <span class="d-ins">...</span> → 안쪽 내용만
        inner = re.sub(r'<span class="d-ins">(.*?)</span>', r'\1', inner, flags=re.DOTALL)
    return f'<span class="dup-line">{inner}<span class="dup-badge">표 내용 중복</span></span>'


def _render_paragraph_table(rows: list, before_dup_norm=None, after_dup_norm=None) -> str:
    """3열(GT/Before/After) 문단 비교 표를 HTML 문자열로 렌더링 (전체, 생략 없음).
    GT 문단이 '## '로 시작하면(슬라이드/섹션 제목) 그 앞에 구분 행을 넣어 경계를 보여준다.
    before/after 칸이 표중복으로 감지되면 파란색으로 표시(삭제 안 함, WER엔 포함)."""
    body_rows = []
    for row in rows:
        gt, before, after = row["gt"], row["before"], row["after"]

        # 슬라이드/섹션 경계 표시
        if gt is not None and gt.lstrip().startswith("## "):
            body_rows.append(_slide_divider_row(3, gt.lstrip()[3:].strip()))

        if gt is None:
            # GT에는 없는 문단(before 또는 after 쪽에만 새로 생긴 문단)
            gt_html = '<span class="d-empty">(해당 없음)</span>'
            before_html = f'<span class="d-ins">{html_lib.escape(before)}</span>' if before else '<span class="d-empty">(해당 없음)</span>'
            after_html = f'<span class="d-ins">{html_lib.escape(after)}</span>' if after else '<span class="d-empty">(해당 없음)</span>'
        else:
            gt_html = html_lib.escape(gt)
            if before is None:
                before_html = '<span class="d-del">(누락됨 - GT에는 있는데 Before에서 사라짐)</span>'
            else:
                gt_diff_b, before_html = _diff_pair_html(gt, before)
            if after is None:
                after_html = '<span class="d-del">(누락됨 - GT에는 있는데 After에서 사라짐)</span>'
            else:
                gt_diff_a, after_html = _diff_pair_html(gt, after)
            # GT 칸은 before 기준 diff를 기본으로(둘 다 없으면 after 기준)
            if before is not None:
                gt_html = gt_diff_b
            elif after is not None:
                gt_html = gt_diff_a

        # 표중복 감지 라인이면 파란색 표시
        before_html = _wrap_dup(before_html, before, before_dup_norm)
        after_html = _wrap_dup(after_html, after, after_dup_norm)

        body_rows.append(f"""
        <tr>
          <td class="col-gt">{gt_html}</td>
          <td class="col-before">{before_html}</td>
          <td class="col-after">{after_html}</td>
        </tr>""")

    return f"""
    <table class="para-table">
      <thead><tr><th>GT (원본)</th><th>Before (가이드라인 적용 전)</th><th>After (가이드라인 적용 후)</th></tr></thead>
      <tbody>{"".join(body_rows)}</tbody>
    </table>"""


def _render_single_table_html(matrix: list, label: str, css_class: str) -> str:
    """table_to_matrix() 결과(list of list)를 실제 <table> HTML로 렌더링."""
    if not matrix:
        return f'<div class="table-block {css_class}"><div class="table-label">{label}</div><div class="d-empty">(표를 찾지 못했거나 비어있음)</div></div>'
    rows_html = []
    for row in matrix:
        cells = "".join(f"<td>{html_lib.escape(c)}</td>" for c in row)
        rows_html.append(f"<tr>{cells}</tr>")
    return f'<div class="table-block {css_class}"><div class="table-label">{label}</div><table class="cell-table">{"".join(rows_html)}</table></div>'


def _render_table_cell_diff(gt_matrix: list, pred_matrix: list, pred_label: str) -> str:
    """
    GT 행렬과 PRED 행렬을 같은 행/열 위치끼리 겹쳐서, 셀마다 일치/유사/불일치를
    색으로 표시한 <table>을 렌더링. (구조 자체가 다르면 - 예: 행/열 개수가 다르면 -
    빈 칸으로 표시되어 "표 구조가 깨졌다"는 것도 바로 드러남)
    """
    n_rows = max(len(gt_matrix), len(pred_matrix))
    if n_rows == 0:
        return '<div class="d-empty">(비교할 표 없음)</div>'
    rows_html = []
    for r in range(n_rows):
        gt_row = gt_matrix[r] if r < len(gt_matrix) else []
        pred_row = pred_matrix[r] if r < len(pred_matrix) else []
        n_cols = max(len(gt_row), len(pred_row))
        cells = []
        for c in range(n_cols):
            gt_cell = gt_row[c] if c < len(gt_row) else None
            pred_cell = pred_row[c] if c < len(pred_row) else None
            if gt_cell is None:
                cells.append(f'<td class="cell-extra">{html_lib.escape(pred_cell or "")}<div class="cell-tag">{pred_label}에만 존재</div></td>')
                continue
            if pred_cell is None:
                cells.append(f'<td class="cell-missing">{html_lib.escape(gt_cell)}<div class="cell-tag">{pred_label}에서 누락</div></td>')
                continue
            sim = _text_similarity(gt_cell, pred_cell)
            css = "cell-match" if sim >= 0.999 else ("cell-partial" if sim >= 0.6 else "cell-mismatch")
            cells.append(
                f'<td class="{css}">GT: {html_lib.escape(gt_cell)}<br>{pred_label}: {html_lib.escape(pred_cell)}'
                f'<div class="cell-tag">유사도 {sim:.2f}</div></td>'
            )
        rows_html.append(f"<tr>{''.join(cells)}</tr>")
    return f'<table class="cell-diff-table">{"".join(rows_html)}</table>'


def _render_tables_section(gt_tables: list, before_tables: list, after_tables: list) -> str:
    """문서 안의 모든 표를 등장 순서대로, GT 원형 + Before/After 셀 비교로 렌더링 (전체, 생략 없음)."""
    n = max(len(gt_tables), len(before_tables), len(after_tables))
    if n == 0:
        return "<p>문서에 표가 없습니다.</p>"

    blocks = []
    for i in range(n):
        gt_html = gt_tables[i] if i < len(gt_tables) else ""
        before_html = before_tables[i] if i < len(before_tables) else ""
        after_html = after_tables[i] if i < len(after_tables) else ""

        gt_matrix = table_to_matrix(gt_html) if gt_html else []
        before_matrix = table_to_matrix(before_html) if before_html else []
        after_matrix = table_to_matrix(after_html) if after_html else []

        blocks.append(f"""
        <div class="table-section">
          <h3>표 {i + 1}</h3>
          <div class="table-columns">
            {_render_single_table_html(gt_matrix, "GT (원본)", "gt")}
          </div>
          <h4>Before vs GT (셀 단위 비교)</h4>
          {_render_table_cell_diff(gt_matrix, before_matrix, "Before")}
          <h4>After vs GT (셀 단위 비교)</h4>
          {_render_table_cell_diff(gt_matrix, after_matrix, "After")}
        </div>""")
    return "".join(blocks)


_CSS = """
body { font-family: -apple-system, "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; margin: 24px; color: #1a1a1a; background:#fafafa; }
h1 { font-size: 22px; }
h2 { font-size: 18px; margin-top: 36px; border-bottom: 2px solid #333; padding-bottom: 6px; }
h3 { font-size: 15px; margin-top: 24px; }
h4 { font-size: 13px; color: #555; margin-bottom: 4px; }
.summary { display:flex; gap:16px; flex-wrap:wrap; margin-bottom: 24px; }
.summary .card { background:white; border:1px solid #ddd; border-radius:8px; padding:12px 18px; min-width:160px; }
.summary .card .label { font-size:12px; color:#777; }
.summary .card .value { font-size:20px; font-weight:bold; }
.legend { font-size:12px; margin-bottom: 12px; }
.legend span { padding: 1px 6px; border-radius: 3px; margin-right: 10px; }
.d-sub { background:#fff3b0; }
.d-del { background:#ffd0d0; text-decoration: line-through; }
.d-ins { background:#c9f2c9; text-decoration: underline; }
.d-empty { color:#aaa; font-style: italic; }
table.para-table { border-collapse: collapse; width: 100%; background:white; table-layout: fixed; }
table.para-table th, table.para-table td { border: 1px solid #ddd; padding: 8px 10px; vertical-align: top; font-size: 13px; line-height:1.6; word-break: break-word; }
table.para-table th { background:#333; color:white; text-align:left; }
table.para-table col { width: 33.3%; }
.col-gt { background: #fbfbfb; }
.table-section { background:white; border:1px solid #ddd; border-radius:8px; padding:16px; margin-top:16px; }
.table-columns { display:flex; gap:16px; flex-wrap:wrap; }
.table-block .table-label { font-weight:bold; font-size:12px; margin-bottom:4px; color:#555; }
table.cell-table { border-collapse: collapse; }
table.cell-table td { border:1px solid #ccc; padding:6px 10px; font-size:13px; }
table.cell-diff-table { border-collapse: collapse; margin-top:6px; margin-bottom:16px; }
table.cell-diff-table td { border:1px solid #ccc; padding:6px 10px; font-size:12px; vertical-align:top; position:relative; }
.cell-match { background:#c9f2c9; }
.cell-partial { background:#fff3b0; }
.cell-mismatch { background:#ffd0d0; }
.cell-missing { background:#ffd0d0; }
.cell-extra { background:#d6e4ff; }
.cell-tag { font-size:10px; color:#666; margin-top:4px; }
.note { font-size:12px; color:#777; }
.dup-wrap { display:flex; gap:16px; flex-wrap:wrap; }
.dup-col { flex:1; min-width:260px; background:white; border:1px solid #ddd; border-radius:8px; padding:12px 16px; }
.dup-list { margin:6px 0; padding-left:18px; }
.dup-list li { font-size:12px; color:#555; margin:2px 0; }
.slide-divider td { background:#eef2ff; color:#3730a3; font-weight:600; text-align:center;
  font-size:12px; letter-spacing:1px; padding:6px; }
.relax-note { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px;
  padding:10px 14px; font-size:12px; color:#166534; margin:8px 0; }
.relax-note ul { margin:6px 0 0; padding-left:18px; }
.dup-line { background:#e0f2fe; border-left:3px solid #0284c7; padding:2px 6px; }
.dup-badge { display:inline-block; font-size:10px; background:#0284c7; color:white;
  border-radius:4px; padding:1px 5px; margin-left:6px; vertical-align:middle; }
"""


def _render_two_way_table(gt_paragraphs: list, pred_paragraphs: list,
                          gt_label: str, pred_label: str, dup_norm=None) -> str:
    """GT 한 개 ↔ 파서 한 개를 2열로 비교하는 표(듀얼 GT일 때 stage별로 각각 렌더).
    GT 문단이 '## '로 시작하면 슬라이드/섹션 구분 행을 넣는다.
    파서 문단이 표중복으로 감지되면 파란색 표시(삭제 안 함)."""
    pairs = _align_pred_to_gt(gt_paragraphs, pred_paragraphs)
    body_rows = []
    for gi, pj, _ in pairs:
        # 슬라이드/섹션 경계
        if gi is not None and gt_paragraphs[gi].lstrip().startswith("## "):
            body_rows.append(_slide_divider_row(2, gt_paragraphs[gi].lstrip()[3:].strip()))

        pred_raw = pred_paragraphs[pj] if pj is not None else None
        if gi is not None and pj is not None:
            gt_diff, pred_html = _diff_pair_html(gt_paragraphs[gi], pred_paragraphs[pj])
            gt_html = gt_diff
        elif gi is not None:
            gt_html = html_lib.escape(gt_paragraphs[gi])
            pred_html = '<span class="d-del">(누락됨 - GT에는 있는데 파서 결과에서 사라짐)</span>'
        else:
            gt_html = '<span class="d-empty">(해당 없음)</span>'
            pred_html = f'<span class="d-ins">{html_lib.escape(pred_paragraphs[pj])}</span>'
        pred_html = _wrap_dup(pred_html, pred_raw, dup_norm)
        body_rows.append(
            f'<tr><td class="col-gt">{gt_html}</td>'
            f'<td class="col-before">{pred_html}</td></tr>'
        )
    return f"""
    <table class="para-table">
      <thead><tr><th>{html_lib.escape(gt_label)}</th><th>{html_lib.escape(pred_label)}</th></tr></thead>
      <tbody>{"".join(body_rows)}</tbody>
    </table>"""


def _render_dup_section(before_dups: list, after_dups: list) -> str:
    """표 셀과 겹쳐 WER에서 제외한 본문 라인들을 '표 내용 중복 추출'로 보여준다.
    (실제로 삭제하는 건 WER 계산에서만이고, 여기서는 참고용으로 남긴다.)"""
    if not before_dups and not after_dups:
        return ""

    def _list(dups):
        if not dups:
            return "<p>(없음)</p>"
        items = "".join(f"<li>{html_lib.escape(d)}</li>" for d in dups)
        return f"<ul class='dup-list'>{items}</ul>"

    return f"""
<h2>3. 표 내용 중복 추출 (표 밖에 중복 추출된 원문 - 표시만, WER 계산엔 포함)</h2>
<p class="note">아래 라인들은 표 안 셀 내용이 표 밖 본문에도 통째로(약 90%+) 중복 추출된
것으로 감지된 것입니다. WER 계산에서 삭제하거나 수정하지 않고 그대로 포함하며,
위 본문 비교에서 파란색으로 표시됩니다. 여기서는 목록으로만 정리합니다.</p>
<div class="dup-wrap">
  <div class="dup-col"><h3>Before</h3>{_list(before_dups)}</div>
  <div class="dup-col"><h3>After</h3>{_list(after_dups)}</div>
</div>"""


def _render_relax_section(relax_before: list, relax_after: list) -> str:
    """줄바꿈 유연화로 합쳐지거나(merge) 나뉜(split) 라인을 리포트에 표시."""
    relax_before = relax_before or []
    relax_after = relax_after or []
    if not relax_before and not relax_after:
        return ""

    def _fmt(log):
        if not log:
            return "<p>(없음)</p>"
        items = []
        for e in log:
            if e["kind"] == "merge":
                items.append(f"<li>두 줄 합침: {html_lib.escape(' / '.join(e['from']))} "
                             f"→ {html_lib.escape(e['to'][0])}</li>")
            else:
                items.append(f"<li>한 줄 나눔: {html_lib.escape(e['from'][0])} "
                             f"→ {html_lib.escape(' / '.join(e['to']))}</li>")
        return f"<ul>{''.join(items)}</ul>"

    return f"""
<h2>4. 줄바꿈 유연화 (WER 계산 시 GT 경계에 맞춘 내역)</h2>
<p class="note">파서 줄바꿈이 GT와 다를 때, 바로 인접한 줄만 확인해 GT 한 줄과 같아지면
합치거나 나눴습니다(±1줄만, 밀림 방지). 아래는 그렇게 조정된 줄들입니다.</p>
<div class="dup-wrap">
  <div class="dup-col"><h3>Before ({len(relax_before)}건)</h3><div class="relax-note">{_fmt(relax_before)}</div></div>
  <div class="dup-col"><h3>After ({len(relax_after)}건)</h3><div class="relax-note">{_fmt(relax_after)}</div></div>
</div>"""


def generate_report(
    gt_text: str, gt_tables: list,
    before_text: str, before_tables: list,
    after_text: str, after_tables: list,
    wer_before: float, wer_after: float,
    teds_before, teds_after,
    output_path: str = "report.html",
    title: str = "파싱 품질 비교 리포트",
    gt_after_text: str = None,      # 듀얼 GT: after 쪽 GT 본문(없으면 gt_text와 동일 취급)
    gt_after_tables: list = None,   # 듀얼 GT: after 쪽 GT 표
    before_dups: list = None,       # WER에서 제외된 표중복 본문 라인(before)
    after_dups: list = None,        # WER에서 제외된 표중복 본문 라인(after)
    relax_log_before: list = None,  # 줄바꿈 유연화 내역(before)
    relax_log_after: list = None,   # 줄바꿈 유연화 내역(after)
) -> str:
    """
    GT/Before/After 전체를 하나의 HTML 리포트로 저장합니다. (문서 전체, 잘림 없음)

    듀얼 GT(gt_after_text가 주어지고 gt_text와 다름)면, before는 gt_text(=before_gt)와,
    after는 gt_after_text(=after_gt)와 각각 2열로 비교한다(서로 다른 원본을 한 GT
    열에 억지로 합치지 않는다). 단일 GT면 기존 3열(GT/Before/After) 비교를 쓴다.
    """
    gt_before_paras = [p for p in gt_text.split("\n") if p.strip()]
    before_paragraphs = [p for p in before_text.split("\n") if p.strip()]
    after_paragraphs = [p for p in after_text.split("\n") if p.strip()]

    dual_gt = gt_after_text is not None and gt_after_text != gt_text
    gt_after_paras = [p for p in (gt_after_text or gt_text).split("\n") if p.strip()]

    # 표중복 감지 라인 → 정규화 집합(본문에서 색 표시용)
    from wer_preprocess import _norm_line
    before_dup_norm = set(_norm_line(d) for d in (before_dups or []) if _norm_line(d))
    after_dup_norm = set(_norm_line(d) for d in (after_dups or []) if _norm_line(d))

    if dual_gt:
        # stage별 2-way 비교 두 개
        para_html = (
            "<h3>Before: 가이드라인 적용 전 원본(before_gt) ↔ 파서 결과(before.txt)</h3>"
            + _render_two_way_table(gt_before_paras, before_paragraphs,
                                     "GT (before_gt)", "Before (파서)", dup_norm=before_dup_norm)
            + "<h3>After: 가이드라인 적용 후 원본(after_gt) ↔ 파서 결과(after.txt)</h3>"
            + _render_two_way_table(gt_after_paras, after_paragraphs,
                                     "GT (after_gt)", "After (파서)", dup_norm=after_dup_norm)
        )
    else:
        rows = _build_three_way_rows(gt_before_paras, before_paragraphs, after_paragraphs)
        para_html = _render_paragraph_table(rows, before_dup_norm, after_dup_norm)

    # 표 섹션: 듀얼 GT면 after쪽 GT 표를 대표로(없으면 gt_tables)
    tables_section_html = _render_tables_section(
        gt_after_tables if (dual_gt and gt_after_tables is not None) else gt_tables,
        before_tables, after_tables)

    dup_section_html = _render_dup_section(before_dups or [], after_dups or [])
    relax_section_html = _render_relax_section(relax_log_before, relax_log_after)

    def _fmt(v):
        return f"{v:.4f}" if isinstance(v, (int, float)) else str(v)

    html_doc = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>{html_lib.escape(title)}</title>
<style>{_CSS}</style>
</head>
<body>
<h1>{html_lib.escape(title)}</h1>

<div class="summary">
  <div class="card"><div class="label">WER (낮을수록 좋음) - Before</div><div class="value">{_fmt(wer_before)}</div></div>
  <div class="card"><div class="label">WER (낮을수록 좋음) - After</div><div class="value">{_fmt(wer_after)}</div></div>
  <div class="card"><div class="label">TEDS-content (높을수록 좋음) - Before</div><div class="value">{_fmt(teds_before)}</div></div>
  <div class="card"><div class="label">TEDS-content (높을수록 좋음) - After</div><div class="value">{_fmt(teds_after)}</div></div>
</div>

<h2>1. 본문 텍스트 비교 (문단 전체, 생략 없음)</h2>
<div class="legend">
  <span class="d-sub">치환(문구가 바뀜)</span>
  <span class="d-del">누락(GT엔 있는데 사라짐)</span>
  <span class="d-ins">삽입(파서가 새로 만든 잡음)</span>
  &nbsp;·&nbsp; 하이라이트 안의 <code>·</code>=공백 <code>⬚</code>=전각공백 <code>→</code>=탭 <code>↵</code>=줄바꿈 (눈에 안 보이던 문자를 드러낸 것)
</div>
{para_html}

<h2>2. 표 비교 (등장 순서대로, 전체)</h2>
{tables_section_html}

{dup_section_html}

{relax_section_html}
</body>
</html>"""

    Path(output_path).write_text(html_doc, encoding="utf-8")
    return output_path


if __name__ == "__main__":
    # 단독 실행 시: 첫 번째 문서로 3-way 리포트를 만든다.
    # 특정 문서를 지정하려면 run_eval.diagnose(doc_id, make_html=True)를 쓰면 된다.
    import run_eval
    docs = run_eval.collect_docs()
    if not docs:
        print("data 폴더에 문서가 없습니다. config.py의 DATA_DIR을 확인하세요.")
    else:
        run_eval._make_html_report(next(iter(docs)))