# -*- coding: utf-8 -*-
"""
parsed_extract.py
=================
파서 결과 txt에서 "본문 텍스트"와 "표"를 분리한다.

파서는 표를 두 가지 형태로 낼 수 있어 둘 다 인식한다.
  (1) HTML 표 : <table>...</table>              (docx/pptx 파서 출력에서 주로 관찰)
  (2) 마크다운 파이프 표 : | 셀 | 셀 |            (xlsx 파서 출력에서 주로 관찰)
                          (엑셀은 빈 칸을 null로 채우고 | 로 열을 구분해 내보냄)

반환
----
  - text   : 표 블록을 뺀 나머지 텍스트(#, **, ~~, - 등 Markdown 기호는 유지)  → WER/CER용
  - tables : 표들을 등장 순서대로 <table> HTML 문자열로 변환한 리스트          → TEDS용
             (마크다운 파이프 표도 내부적으로 <table> HTML로 바꿔 통일한다.)

엑셀 전용 비교(셀 F1)를 위해, 파이프 표를 "행렬(list of list)"로 바로 주는
load_parsed_matrices()도 제공한다.
"""

import re
import html as html_lib
from pathlib import Path
from typing import List, Tuple

# <table ...> ... </table> (여는/닫는 태그의 공백 변형 허용)
_TABLE_PATTERN = re.compile(r"<\s*table[^>]*>.*?<\s*/\s*table\s*>", re.DOTALL | re.IGNORECASE)

# 마크다운 표의 헤더 구분행: | --- | :---: | 처럼 대시(-)/콜론(:)/파이프만 있는 줄
_MD_SEP_ROW = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


def _looks_like_pipe_row(line: str) -> bool:
    """'| a | b |' 형태의 표 행처럼 보이는지."""
    s = line.strip()
    return s.startswith("|") and s.count("|") >= 2


def _split_pipe_row(line: str) -> List[str]:
    """'| a | b |' → ['a', 'b'] (양끝 파이프 제거 후 | 로 분리)."""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _parse_pipe_tables(text: str) -> Tuple[str, List[List[List[str]]]]:
    """
    파이프 표 블록들을 찾아 (표를 뺀 텍스트, 표 행렬 리스트)로 나눈다.
    연속된 파이프 행들이 표 한 개다. 헤더 구분행(| --- |)은 건너뛴다.
    파이프 행이 아닌 줄(빈 줄, '## 제목' 등)을 만나면 표 블록이 끝난다.
    """
    matrices, current, kept_lines = [], [], []
    for line in text.split("\n"):
        if _looks_like_pipe_row(line):
            if not _MD_SEP_ROW.match(line):        # 구분행은 표에 넣지 않음
                current.append(_split_pipe_row(line))
        else:
            if current:
                matrices.append(current)
                current = []
            kept_lines.append(line)                # 표가 아닌 줄만 본문으로 남김
    if current:
        matrices.append(current)
    return "\n".join(kept_lines), matrices


def _matrix_to_html(matrix: List[List[str]]) -> str:
    rows = []
    for row in matrix:
        cells = "".join(f"<td>{html_lib.escape(c)}</td>" for c in row)
        rows.append(f"<tr>{cells}</tr>")
    return f"<table>{''.join(rows)}</table>"


def split_text_and_tables(raw_text: str) -> Tuple[str, List[str]]:
    """
    txt 전체를 (표 뺀 텍스트, 표 HTML 리스트)로 나눈다.
    HTML 표를 먼저 뽑아내고, 남은 텍스트에서 파이프 표를 뽑아 HTML로 변환해 합친다.
    """
    html_tables = [m.group(0) for m in _TABLE_PATTERN.finditer(raw_text)]
    text_wo_html = _TABLE_PATTERN.sub("\n", raw_text)

    text_only, pipe_matrices = _parse_pipe_tables(text_wo_html)
    pipe_tables = [_matrix_to_html(m) for m in pipe_matrices]

    text_only = re.sub(r"\n{2,}", "\n", text_only).strip()
    return text_only, html_tables + pipe_tables


def load_parsed_txt(path) -> Tuple[str, List[str]]:
    """txt 경로를 받아 split_text_and_tables 결과를 반환. 없으면 ('', [])."""
    path = Path(path)
    if not path.exists():
        return "", []
    return split_text_and_tables(path.read_text(encoding="utf-8"))


def load_parsed_matrices(path) -> List[List[List[str]]]:
    """
    파서 txt에서 파이프 표들을 "행렬(list of list) 리스트"로 반환(엑셀 셀 F1용).
    HTML 표가 있으면 그것도 행렬로 변환해 함께 반환한다.
    """
    path = Path(path)
    if not path.exists():
        return []
    raw = path.read_text(encoding="utf-8")

    matrices = []
    # HTML 표 → 행렬
    from metrics import table_to_matrix
    for html in _TABLE_PATTERN.findall(raw):
        m = table_to_matrix(html)
        if m:
            matrices.append(m)
    # 파이프 표 → 행렬
    _, pipe_matrices = _parse_pipe_tables(_TABLE_PATTERN.sub("\n", raw))
    matrices.extend(pipe_matrices)
    return matrices