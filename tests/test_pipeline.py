"""naming 규칙 + prepare→build 엔드투엔드 (렌더링 생략, dry-run)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / ".claude" / "skills" / "bear-monthly-update"
sys.path.insert(0, str(SKILL / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))

import naming  # noqa: E402

try:
    import openpyxl  # noqa: F401
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

FIXTURE = ROOT / "tests" / "fixtures" / "sheet_dump_sample.md"


class TestNaming(unittest.TestCase):
    def test_names(self):
        self.assertEqual(naming.status_name("26년 9월", 1), "BEAR_아이디어_제안_현황판_26년9월_v1.xlsx")
        self.assertEqual(naming.email_name("26년 12월"), "BEAR_아이디어_제안_현황_공유_26년12월_1주차.eml")
        self.assertEqual(naming.normalize_month("24년  11월"), "24년 11월")
        self.assertEqual(naming.normalize_month("24냔11월"), "24년 11월")
        self.assertEqual(naming.parse_status_name("BEAR_아이디어_제안_현황판_26년9월_v2.xlsx"), (26, 9, 2))

    def test_latest_base_and_version(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            for n in ["BEAR_아이디어_제안_현황판_26년7월_v1.xlsx", "BEAR_아이디어_제안_현황판_26년8월_v1.xlsx",
                      "BEAR_아이디어_제안_현황판_26년8월_v2.xlsx", "BEAR_아이디어_제안_현황판_26년9월_v1.xlsx", "기타.xlsx"]:
                (d / n).write_bytes(b"")
            self.assertEqual(naming.latest_base(d).name, "BEAR_아이디어_제안_현황판_26년9월_v1.xlsx")
            self.assertEqual(naming.latest_base(d, before="26년 9월").name, "BEAR_아이디어_제안_현황판_26년8월_v2.xlsx")
            self.assertEqual(naming.next_version(d, "26년 9월"), 2)
            self.assertEqual(naming.next_version(d, "26년 10월"), 1)


@unittest.skipUnless(HAVE_OPENPYXL, "openpyxl 미설치")
class TestPipelineEndToEnd(unittest.TestCase):
    def setUp(self):
        import helpers
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        (root / "data").mkdir()
        helpers.make_base_xlsx(root / "data" / "BEAR_아이디어_제안_현황판_26년8월_v1.xlsx")
        cfg = json.loads((SKILL / "config.json").read_text(encoding="utf-8"))
        cfg["paths"] = {"data_dir": str(root / "data"), "output_dir": str(root / "output"),
                        "work_dir": str(root / "work"), "runs_state": str(root / "data" / "runs.json")}
        self.cfg = root / "config.json"
        self.cfg.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        self.root = root

    def tearDown(self):
        self.tmp.cleanup()

    def run_cli(self, *args):
        res = subprocess.run([sys.executable, str(SKILL / "scripts" / "pipeline.py"), "--config", str(self.cfg), *args],
                             capture_output=True, text=True, timeout=300)
        return res

    def test_prepare_then_build(self):
        res = self.run_cli("prepare", "--dump", str(FIXTURE), "--report-month", "26년 9월")
        self.assertEqual(res.returncode, 0, res.stderr)
        meta = json.loads(res.stdout)
        self.assertEqual(meta["new_count"], 4)
        proposals = Path(meta["work_dir"]) / "proposals.json"
        self.assertTrue(proposals.exists())

        # dry-run: data/·runs.json 미변경
        res = self.run_cli("build", "--proposals", str(proposals), "--dry-run", "--skip-render")
        self.assertEqual(res.returncode, 0, res.stderr)
        man = json.loads(res.stdout)
        self.assertTrue(Path(man["status_xlsx"]).exists())
        self.assertTrue(Path(man["eml"]).exists())
        self.assertEqual(man["version"], 1)
        self.assertEqual(man["unmatched_teams"], ["미지의팀사업팀"])
        self.assertTrue(man["attention"])
        self.assertIn("[확인 필요]", man["notify_subject"])
        self.assertFalse((self.root / "data" / "runs.json").exists())
        self.assertEqual(len(list((self.root / "data").glob("*.xlsx"))), 1)

        eml = Path(man["eml"]).read_bytes().decode("utf-8", errors="replace")
        self.assertIn("Subject:", eml)
        self.assertIn("multipart", eml)

        # 실제 build: data/ 복사 + runs.json, 두 번째는 already_done
        res = self.run_cli("build", "--proposals", str(proposals), "--skip-render")
        self.assertEqual(res.returncode, 0, res.stderr)
        man = json.loads(res.stdout)
        self.assertEqual(man["version"], 2)
        runs = json.loads((self.root / "data" / "runs.json").read_text(encoding="utf-8"))
        self.assertIn("26년 9월", runs)
        self.assertTrue((self.root / "data" / "BEAR_아이디어_제안_현황판_26년9월_v2.xlsx").exists())

        res = self.run_cli("build", "--proposals", str(proposals), "--skip-render")
        self.assertEqual(res.returncode, 4)

        # 다음 달은 새 파일을 베이스로 (26년 10월 → 9월 v2)
        self.assertEqual(naming.latest_base(self.root / "data", before="26년 10월").name,
                         "BEAR_아이디어_제안_현황판_26년9월_v2.xlsx")


if __name__ == "__main__":
    unittest.main()
