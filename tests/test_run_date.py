"""실행일 규칙 테스트. 휴일 집합을 직접 주입하므로 holidays 패키지 없이도 돈다."""
import sys
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / ".claude" / "skills" / "bear-monthly-update" / "scripts"))

import run_date as rd  # noqa: E402

# 2026~2027 공휴일 중 규칙에 영향을 주는 날 (holidays.KR 값과 대조함)
HOL = {
    date(2026, 10, 5),   # 개천절(10/3 토) 대체휴일
    date(2027, 3, 1),    # 삼일절
    date(2027, 5, 3),    # 노동절(5/1 토) 대체휴일
    date(2027, 10, 4),   # 개천절(10/3 일) 대체휴일
    date(2027, 2, 8), date(2027, 2, 9),  # 설 연휴
}


class TestComputeRunDate(unittest.TestCase):
    TABLE = [
        # (보고월, 기대 실행일, 사유 키워드)
        ((2026, 9), date(2026, 8, 31), "화요일"),      # 1일=화 → 전월 말 월요일
        ((2026, 10), date(2026, 10, 6), "다음 영업일"),  # 첫 월요일 10/5 가 대체휴일
        ((2026, 11), date(2026, 11, 2), "첫 월요일"),
        ((2026, 12), date(2026, 11, 30), "화요일"),
        ((2027, 1), date(2027, 1, 4), "첫 월요일"),
        ((2027, 2), date(2027, 2, 1), "1일이 월요일"),
        ((2027, 3), date(2027, 3, 2), "다음 영업일"),   # 삼일절 월요일
        ((2027, 5), date(2027, 5, 4), "다음 영업일"),   # 노동절 대체휴일
        ((2027, 6), date(2027, 5, 31), "화요일"),
        ((2027, 10), date(2027, 10, 5), "다음 영업일"),
        ((2027, 11), date(2027, 11, 1), "1일이 월요일"),
    ]

    def test_table(self):
        for (y, m), expected, kw in self.TABLE:
            got, reason = rd.compute_run_date(y, m, HOL)
            self.assertEqual(got, expected, f"{y}-{m}: {reason}")
            self.assertIn(kw, reason)

    def test_prev_month_monday_holiday_shifts_to_first(self):
        # 1일=화요일인데 전월 말 월요일이 휴일이면 그 다음 영업일(=1일 화요일)로
        got, _ = rd.compute_run_date(2026, 9, HOL | {date(2026, 8, 31)})
        self.assertEqual(got, date(2026, 9, 1))

    def test_weekend_after_holiday(self):
        # 금요일까지 밀리면 토·일을 건너뛴다
        got, _ = rd.compute_run_date(2027, 11, {date(2027, 11, 1), date(2027, 11, 2),
                                                date(2027, 11, 3), date(2027, 11, 4), date(2027, 11, 5)})
        self.assertEqual(got, date(2027, 11, 8))


class TestDecide(unittest.TestCase):
    def test_run_day(self):
        r = rd.decide(date(2026, 8, 31), HOL)
        self.assertTrue(r["run"])
        self.assertEqual(r["report_month"], "26년 9월")
        self.assertEqual(r["prev_month"], "8월")
        self.assertEqual(r["late_days"], 0)

    def test_catchup_within_grace(self):
        r = rd.decide(date(2026, 9, 3), HOL, runs={})
        self.assertTrue(r["run"])
        self.assertEqual(r["report_month"], "26년 9월")
        self.assertEqual(r["late_days"], 3)

    def test_already_done(self):
        r = rd.decide(date(2026, 9, 3), HOL, runs={"26년 9월": {"run_date": "2026-08-31"}})
        self.assertFalse(r["run"])
        self.assertEqual(r["next_report_month"], "26년 10월")
        self.assertEqual(r["next_run_date"], "2026-10-06")

    def test_beyond_grace(self):
        r = rd.decide(date(2026, 9, 15), HOL, runs={}, grace_days=10)
        self.assertFalse(r["run"])

    def test_not_run_day(self):
        r = rd.decide(date(2026, 10, 5), HOL, runs={})
        self.assertFalse(r["run"])
        self.assertEqual(r["next_run_date"], "2026-10-06")

    def test_next_runs_listing(self):
        rows = rd.next_runs(date(2026, 9, 1), 3, HOL)
        self.assertEqual([r["run_date"] for r in rows], ["2026-10-06", "2026-11-02", "2026-11-30"])


class TestHolidaysPackage(unittest.TestCase):
    def test_matches_package_if_installed(self):
        try:
            import holidays  # noqa: F401
        except ImportError:
            self.skipTest("holidays 미설치")
        real = rd.load_holidays([2026, 2027], {}, warn=lambda s: None)
        for d in HOL:
            self.assertIn(d, real, f"{d} 가 holidays.KR 에 없음")


if __name__ == "__main__":
    unittest.main()
