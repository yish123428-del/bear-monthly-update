#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""현황판/메일 파일명 규칙과 최신 베이스 탐색.

  현황판: BEAR_아이디어_제안_현황판_26년9월_v1.xlsx  (같은 월이 있으면 v2, v3 …)
  메일  : BEAR_아이디어_제안_현황_공유_26년9월_1주차.eml
"""
from __future__ import annotations

import re
from pathlib import Path

STATUS_PREFIX = "BEAR_아이디어_제안_현황판_"
EMAIL_PREFIX = "BEAR_아이디어_제안_현황_공유_"
_STATUS_RE = re.compile(r"BEAR_아이디어_제안_현황판_(\d{2})년\s*(\d{1,2})월_v(\d+)\.xlsx$")
# '년' 오타('냔' 등)도 시트에 실제로 존재하므로 관대하게 받는다.
_MONTH_RE = re.compile(r"(\d{2})\s*[년냔]\s*(\d{1,2})\s*월")


def parse_month(label: str) -> tuple[int, int]:
    """'26년 9월' → (26, 9)."""
    m = _MONTH_RE.search(label or "")
    if not m:
        raise ValueError(f"월 라벨 형식이 아님: {label!r}")
    return int(m.group(1)), int(m.group(2))


def compact_month(label: str) -> str:
    """'26년 9월' → '26년9월' (파일명용)."""
    yy, mm = parse_month(label)
    return f"{yy}년{mm}월"


def normalize_month(label: str) -> str:
    """'24년  11월' / '24년11월' → '24년 11월'."""
    yy, mm = parse_month(label)
    return f"{yy}년 {mm}월"


def status_name(report_month: str, version: int) -> str:
    return f"{STATUS_PREFIX}{compact_month(report_month)}_v{version}.xlsx"


def email_name(report_month: str) -> str:
    return f"{EMAIL_PREFIX}{compact_month(report_month)}_1주차.eml"


def parse_status_name(name: str):
    m = _STATUS_RE.search(Path(name).name)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def latest_base(data_dir: str | Path, before: str | None = None) -> Path | None:
    """data/ 에서 (연, 월, 버전)이 가장 큰 현황판.

    before='26년 9월' 이면 그 보고월보다 앞선 파일 중 최신 (같은 달 재실행 시 v1 을 베이스로 쓰지 않도록).
    """
    limit = parse_month(before) if before else None
    best = None
    for p in Path(data_dir).glob("*.xlsx"):
        key = parse_status_name(p.name)
        if not key:
            continue
        if limit and key[:2] >= limit:
            continue
        if best is None or key > best[0]:
            best = (key, p)
    return best[1] if best else None


def next_version(output_dir: str | Path, report_month: str) -> int:
    """output/<월>/ 과 data/ 어디에도 없으면 1, 있으면 최대 버전 + 1."""
    yy, mm = parse_month(report_month)
    top = 0
    for p in Path(output_dir).rglob("*.xlsx"):
        key = parse_status_name(p.name)
        if key and key[:2] == (yy, mm):
            top = max(top, key[2])
    return top + 1
