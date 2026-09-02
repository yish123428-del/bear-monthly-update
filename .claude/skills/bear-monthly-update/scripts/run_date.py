#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BEAR 월간 업데이트 실행일 게이트.

규칙 (사용자 확인, 2026-09):
  * 보고월의 1일이 월요일이면 1일.
  * 1일이 화요일이면 전월 말일(그 주 월요일).
  * 1일이 수~일요일이면 그 달의 첫 월요일.
  * 그 날이 공휴일/주말이면 다음 영업일로 미룬다.

Routine 은 실행일 "후보일"마다 깨어나므로, 오늘이 실행일이거나 실행일이 지났는데
아직 그 보고월을 처리하지 않았으면(runs.json 에 없음, grace_days 이내) 실행으로 판정한다.
"""
from __future__ import annotations

import argparse
import calendar
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

SKILL_DIR = Path(__file__).resolve().parent.parent
EXIT_NOT_RUN_DAY = 3


# ---------- 휴일 ----------

def load_holidays(years, config=None, warn=print) -> set[date]:
    """holidays 패키지(KR, 대체휴일 포함) + config.schedule.extra_holidays/skip_dates."""
    days: set[date] = set()
    try:
        import holidays as _hol
        days.update(_hol.KR(years=list(years)).keys())
    except ImportError:
        warn("[warn] `holidays` 패키지가 없어 config 의 extra_holidays 만 사용합니다.")
    sched = (config or {}).get("schedule", {})
    for key in ("extra_holidays", "skip_dates"):
        for s in sched.get(key, []) or []:
            days.add(date.fromisoformat(s))
    return days


# ---------- 규칙 ----------

def first_weekday(year: int, month: int, weekday: int = 0) -> date:
    d = date(year, month, 1)
    return d + timedelta(days=(weekday - d.weekday()) % 7)


def compute_run_date(year: int, month: int, holidays: set[date]) -> tuple[date, str]:
    """보고월(year, month)의 실행일과 사유."""
    d1 = date(year, month, 1)
    wd = d1.weekday()  # Mon=0
    if wd == 0:
        d, reason = d1, "1일이 월요일"
    elif wd == 1:
        d, reason = d1 - timedelta(days=1), "1일이 화요일 → 전월 말일(월요일)"
    else:
        d, reason = first_weekday(year, month, 0), "그 달의 첫 월요일"
    shifted = False
    while d.weekday() >= 5 or d in holidays:
        d += timedelta(days=1)
        shifted = True
    if shifted:
        reason += " → 공휴일/주말이라 다음 영업일"
    return d, reason


def add_months(year: int, month: int, n: int) -> tuple[int, int]:
    idx = year * 12 + (month - 1) + n
    return idx // 12, idx % 12 + 1


def month_label(year: int, month: int) -> str:
    """'26년 9월' 형식."""
    return f"{year % 100}년 {month}월"


def decide(today: date, holidays: set[date], runs: dict | None = None,
           grace_days: int = 10) -> dict:
    """오늘 실행해야 하는지 판정한다.

    오늘이 속한 달과 다음 달을 보고월 후보로 본다 (1일=화요일이면 전월 말일에 실행하므로
    다음 달 보고월이 오늘 실행일일 수 있다).
    """
    runs = runs or {}
    candidates = []
    for n in (0, 1):
        y, m = add_months(today.year, today.month, n)
        run_date, reason = compute_run_date(y, m, holidays)
        label = month_label(y, m)
        py, pm = add_months(y, m, -1)
        info = {
            "report_month": label,
            "report_label": f"{m}월",
            "prev_month": f"{pm}월",
            "prev_month_label": month_label(py, pm),
            "run_date": run_date.isoformat(),
            "reason": reason,
            "already_done": label in runs,
        }
        candidates.append((run_date, info))

    for run_date, info in candidates:
        if info["already_done"]:
            continue
        if run_date <= today <= run_date + timedelta(days=grace_days):
            late = (today - run_date).days
            return {
                "run": True,
                "today": today.isoformat(),
                **info,
                "late_days": late,
                "note": "정규 실행일" if late == 0 else f"실행일에서 {late}일 지연된 캐치업 실행",
            }

    # 실행 아님 → 다음 실행일 안내
    upcoming = [(rd, i) for rd, i in candidates if rd > today and not i["already_done"]]
    if not upcoming:
        y, m = add_months(today.year, today.month, 2)
        rd, reason = compute_run_date(y, m, holidays)
        upcoming = [(rd, {"report_month": month_label(y, m), "run_date": rd.isoformat(), "reason": reason})]
    rd, i = min(upcoming, key=lambda x: x[0])
    return {
        "run": False,
        "today": today.isoformat(),
        "next_run_date": i["run_date"],
        "next_report_month": i["report_month"],
        "reason": i["reason"],
    }


def next_runs(start: date, n: int, holidays: set[date]) -> list[dict]:
    out = []
    y, m = start.year, start.month
    while len(out) < n:
        rd, reason = compute_run_date(y, m, holidays)
        if rd >= start:
            out.append({"report_month": month_label(y, m), "run_date": rd.isoformat(),
                        "weekday": "월화수목금토일"[rd.weekday()], "reason": reason})
        y, m = add_months(y, m, 1)
    return out


# ---------- CLI ----------

def load_json(path) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    return json.loads(p.read_text(encoding="utf-8"))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(SKILL_DIR / "config.json"))
    ap.add_argument("--runs", default=None, help="runs.json 경로 (기본: config.paths.runs_state)")
    ap.add_argument("--today", default=None, help="YYYY-MM-DD (테스트·재현용)")
    ap.add_argument("--check", action="store_true", help="오늘 실행 여부 판정 (실행이면 exit 0, 아니면 3)")
    ap.add_argument("--next", type=int, default=0, metavar="N", help="향후 N개 실행일 표시")
    args = ap.parse_args(argv)

    config = load_json(args.config)
    tz = ZoneInfo(config.get("schedule", {}).get("timezone", "Asia/Seoul"))
    today = date.fromisoformat(args.today) if args.today else datetime.now(tz).date()
    holidays = load_holidays(range(today.year - 1, today.year + 3), config, warn=lambda s: print(s, file=sys.stderr))

    if args.next:
        rows = next_runs(today, args.next, holidays)
        print("보고월      실행일        요일  사유")
        for r in rows:
            print(f"{r['report_month']:<10} {r['run_date']}  {r['weekday']}    {r['reason']}")
        if not args.check:
            return 0

    repo_root = SKILL_DIR.parent.parent.parent
    runs_path = args.runs or str(repo_root / config.get("paths", {}).get("runs_state", "data/runs.json"))
    runs = load_json(runs_path)
    grace = int(config.get("schedule", {}).get("grace_days", 10))
    result = decide(today, holidays, runs, grace)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["run"] else EXIT_NOT_RUN_DAY


if __name__ == "__main__":
    sys.exit(main())
