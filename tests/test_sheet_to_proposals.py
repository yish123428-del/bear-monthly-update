"""시트 덤프 → 신규 제안 변환 테스트."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / ".claude" / "skills" / "bear-monthly-update"
sys.path.insert(0, str(SKILL / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))

try:
    import openpyxl  # noqa: F401
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

FIXTURE = ROOT / "tests" / "fixtures" / "sheet_dump_sample.md"
CONFIG = json.loads((SKILL / "config.json").read_text(encoding="utf-8"))


@unittest.skipUnless(HAVE_OPENPYXL, "openpyxl 미설치")
class TestParseDump(unittest.TestCase):
    def setUp(self):
        import dumpio
        import sheet_to_proposals as s2p
        self.s2p = s2p
        self.rows, self.warnings = s2p.parse_dump(dumpio.load_dump(FIXTURE), CONFIG["sheet"]["anchor_labels"])

    def test_first_block_only(self):
        months = [r["제안월"] for r in self.rows]
        self.assertNotIn("25년 1월", months, "두 번째(구버전) 탭이 섞이면 안 됨")
        self.assertEqual(len(self.rows), 8)

    def test_month_normalized(self):
        self.assertEqual(self.rows[0]["제안월"], "24년 11월")   # '24냔 11월' 오타
        self.assertEqual(self.rows[3]["제안월"], "26년 7월")    # '26년  7월' 이중 공백

    def test_columns_aligned(self):
        r = self.rows[1]
        self.assertEqual(r["접수자"], "김민수")
        self.assertEqual(r["KOL"], "김민수PM")
        self.assertEqual(r["대분류"], "복합제")
        self.assertEqual(r["현황"], "진행중")
        self.assertEqual(r["세부현황"], "1차 타당성 검토")
        self.assertEqual(self.rows[0]["소분류"], "-")  # \- 언이스케이프


@unittest.skipUnless(HAVE_OPENPYXL, "openpyxl 미설치")
class TestNormalize(unittest.TestCase):
    def setUp(self):
        import sheet_to_proposals as s2p
        self.s2p = s2p
        self.m = CONFIG["mapping"]

    def test_bonbu(self):
        for raw, want in [("ETC마케팅본부", "MKT"), ("ETC 마케팅", "MKT"), ("ETC병원", "영업본부"),
                          ("ETC 병원본부", "영업본부"), ("ETC로컬본부", "영업본부")]:
            self.assertEqual(self.s2p.norm_bonbu(raw, self.m), want)

    def test_team_alias_and_suffix(self):
        self.assertEqual(self.s2p.norm_team("호피안", "MKT", self.m), "호흡피부안과사업팀")
        self.assertEqual(self.s2p.norm_team("호흡피부안과", "MKT", self.m), "호흡피부안과사업팀")
        self.assertEqual(self.s2p.norm_team("소화기2", "MKT", self.m), "소화기2사업팀")
        self.assertEqual(self.s2p.norm_team("병원경기2", "영업본부", self.m), "병원경기2")
        self.assertEqual(self.s2p.norm_division("마케팅1", "MKT", self.m), "마케팅1사업부")
        self.assertEqual(self.s2p.norm_division("서울2", "영업본부", self.m), "서울2사업부")
        self.assertEqual(self.s2p.norm_team("디지털1", "MKT", self.m), "디지털헬스1팀")

    def test_display_fields_for_sales(self):
        row = {"제안월": "26년 7월", "본부": "ETC 병원본부", "사업부": "서울2", "팀": "병원경기2", "접수자": "권봉기",
               "품목분류": "의약품", "개발형태": "자체 개발", "대분류": "복합제", "현황": "진행중", "세부현황": "1차 타당성 검토"}
        p = self.s2p.to_proposal(row, self.m)
        self.assertEqual((p["본부"], p["본부표시"]), ("영업본부", ""))
        self.assertEqual((p["팀"], p["팀표시"]), ("병원경기2", "병원경기2사무소"))
        self.assertEqual(p["사업부"], "서울2사업부")
        mkt = self.s2p.to_proposal({**row, "본부": "ETC마케팅본부", "사업부": "마케팅1", "팀": "소화기2"}, self.m)
        self.assertEqual((mkt["본부표시"], mkt["팀표시"]), ("MKT", "소화기2사업팀"))

    def test_review_status(self):
        self.assertEqual(self.s2p.norm_review("1차 타당성 검토", "진행중", self.m), "1차 타당성 검토중")
        self.assertEqual(self.s2p.norm_review("1차 피드백", "검토 중단", self.m), "검토 중단")
        self.assertEqual(self.s2p.norm_review("NBTS 제안", "진행중", self.m), "NBTS 제안")

    def test_proposal_fields(self):
        row = {"제안월": "26년 4월", "본부": "ETC마케팅", "사업부": "마케팅1", "팀": "소화기2", "접수자": "김민수",
               "품목분류": "의약품", "개발형태": "자체 개발", "대분류": "복합제", "소분류": "복용편의성 개선",
               "제품명": "우루사 100mg 실리스칸 175mg", "성분": "UDCA+실리마린", "현황": "진행중",
               "세부현황": "1차 타당성 검토"}
        p = self.s2p.to_proposal(row, self.m)
        self.assertEqual(p["제안유형"], "의약품 자체개발")
        self.assertEqual(p["아이디어유형"], "복합제")
        self.assertEqual(p["제안자"], "김민수")
        self.assertEqual(p["건수"], 1)
        self.assertEqual(p["제안내용"], "우루사 100mg 실리스칸 175mg 복합제(UDCA+실리마린)")
        self.assertTrue(p["_draft"])


@unittest.skipUnless(HAVE_OPENPYXL, "openpyxl 미설치")
class TestSelectionEndToEnd(unittest.TestCase):
    def setUp(self):
        import helpers
        import sheet_to_proposals as s2p
        self.s2p = s2p
        self.tmp = tempfile.TemporaryDirectory()
        self.base = helpers.make_base_xlsx(Path(self.tmp.name) / "BEAR_아이디어_제안_현황판_26년8월_v1.xlsx")
        self.meta = s2p.run(str(FIXTURE), "26년 9월", Path(self.tmp.name) / "work", CONFIG, str(self.base))
        self.data = json.loads((Path(self.tmp.name) / "work" / "proposals.json").read_text(encoding="utf-8"))
        self.props = self.data["proposals"]

    def tearDown(self):
        self.tmp.cleanup()

    def keys(self):
        return [(p["월"], p["제안자"], p["_reason"]) for p in self.props]

    def test_selected_rows(self):
        keys = self.keys()
        self.assertIn(("26년 8월", "김혁(C)", "전월 신규"), keys)
        self.assertIn(("26년 8월", "최태우", "전월 신규"), keys)
        self.assertIn(("26년 8월", "홍길동", "전월 신규"), keys)
        # 26년 4월 이영민은 베이스 시트2에 없어 캐치업, 김민수는 이미 있어 제외
        self.assertTrue(any(k[1] == "이영민" and "캐치업" in k[2] for k in keys))
        self.assertFalse(any(k[1] == "김민수" for k in keys))
        self.assertFalse(any(k[1] == "권봉기" for k in keys))
        self.assertEqual(len(self.props), 4)

    def test_dedupe(self):
        self.assertEqual(sum(1 for p in self.props if p["제안자"] == "김혁(C)"), 1)
        self.assertTrue(any("중복 행 제거" in w for w in self.meta["warnings"]))

    def test_unmatched_flag(self):
        by = {p["제안자"]: p for p in self.props}
        self.assertTrue(by["홍길동"]["_unmatched"])
        self.assertEqual(by["홍길동"]["팀"], "미지의팀사업팀")
        self.assertFalse(by["김혁(C)"]["_unmatched"])
        self.assertFalse(by["최태우"]["_unmatched"])  # 북부2 는 사무소 섹션에 있음
        self.assertTrue(any("팀 미매칭" in w for w in self.meta["warnings"]))

    def test_counts_and_warnings(self):
        self.assertEqual(self.meta["new_count"], 4)
        self.assertEqual(self.meta["total_count"], 84 + 4)
        self.assertEqual(self.meta["sheet_row_count"], 8)
        self.assertTrue(any("누계 불일치" in w for w in self.meta["warnings"]))

    def test_long_ingredient_truncated(self):
        by = {p["제안자"]: p for p in self.props}
        self.assertLess(len(by["최태우"]["제안내용"]), 120)
        self.assertIn("Casein", by["최태우"]["_source"]["성분"])  # 원문은 보존

    def test_outputs_exist(self):
        work = Path(self.tmp.name) / "work"
        self.assertTrue((work / "new_proposals.xlsx").exists())
        self.assertTrue((work / "review.md").exists())
        review = (work / "review.md").read_text(encoding="utf-8")
        self.assertIn("팀 미매칭", review)
        wb = openpyxl.load_workbook(work / "new_proposals.xlsx")
        ws = wb.active
        self.assertEqual(ws.cell(row=3, column=1).value, "제안 연월")
        self.assertEqual(ws.max_row, 3 + 4)

    def test_no_base(self):
        meta = self.s2p.run(str(FIXTURE), "26년 9월", Path(self.tmp.name) / "work2", CONFIG, None)
        self.assertEqual(meta["new_count"], 3)          # 캐치업 없음
        self.assertEqual(meta["total_count"], 8)        # 시트 행 수 사용
        self.assertEqual(meta["total_count_source"], "sheet_rows")


if __name__ == "__main__":
    unittest.main()
