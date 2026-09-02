#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Google Drive `read_file_content` 덤프(markdown 표) 읽기 유틸.

`mcp__Google_Drive__read_file_content` 는 스프레드시트의 모든 탭을 시트 이름 없이
markdown 표로 이어붙여 돌려준다. 하네스는 큰 결과를 {"fileContent": "..."} JSON 파일로
저장하므로, 여기서는 JSON 과 markdown 원문을 모두 받는다.

task-master 저장소의 parse_board.py 에서 이식했다.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

_MERGED = re.compile(r"\\?\[merged\\?\]\s*")
_ESCAPED_PUNCT = re.compile(r"\\([-~+*_#.\[\]()|>`!])")


def unescape(text: str) -> str:
    """markdown 렌더링 산물(\\[merged\\], \\-, \\~ 등)을 원래 문자로 되돌린다."""
    if text is None:
        return ""
    text = _MERGED.sub("", text)
    text = _ESCAPED_PUNCT.sub(r"\1", text)
    text = text.replace("\\\\", "\\")
    # 표 렌더러가 줄바꿈을 지운 자리에 남은 연속 공백을 정리한다.
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def split_row(line: str) -> list[str]:
    """markdown 표 한 행을 셀 리스트로 자른다.

    빈 셀을 버리면 컬럼이 밀리므로 절대 필터링하지 않는다.
    """
    return [c.strip() for c in line.strip().strip("|").split("|")]


def is_separator(cells: list[str]) -> bool:
    return bool(cells) and all(re.fullmatch(r":?-+:?", c or "-") for c in cells if c)


def load_dump(path: str | Path) -> str:
    """JSON({"fileContent": ...}) 또는 markdown 원문 파일을 텍스트로 돌려준다."""
    raw = Path(path).read_text(encoding="utf-8")
    stripped = raw.lstrip()
    if stripped.startswith("{"):
        try:
            return json.loads(raw)["fileContent"]
        except (json.JSONDecodeError, KeyError, TypeError):
            pass
    return raw


def find_header_blocks(lines: list[str], anchor_labels) -> list[tuple[int, list[str]]]:
    """앵커 라벨이 모두 들어있는 헤더 행을 전부 찾는다. (index, 셀 리스트) 목록.

    같은 표가 여러 탭에 복제돼 있으면 여러 번 잡힌다. 호출자가 첫 블록을 고른다.
    """
    found = []
    for idx, line in enumerate(lines):
        if "|" not in line:
            continue
        cells = [unescape(c) for c in split_row(line)]
        if all(any(label == c for c in cells) for label in anchor_labels):
            found.append((idx, cells))
    return found
