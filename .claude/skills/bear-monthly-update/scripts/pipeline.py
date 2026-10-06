#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BEAR 월간 업데이트 오케스트레이터.

  prepare : 시트 덤프 → work/<월>/{new_proposals.xlsx, proposals.json, review.md}
            (Claude 가 proposals.json 의 제안내용/이메일표기 문구를 다듬을 수 있게 build 와 분리)
  build   : proposals.json → 현황판 .xlsx 갱신 → 시트 이미지 → 공유 메일 .eml → output/<월>/
            성공 시 새 현황판을 data/ 에 복사하고 data/runs.json 에 기록한다.

두 단계 모두 MCP 를 호출하지 않는다 (시트 수집·Drive 업로드·Gmail 알림은 SKILL.md 의 Claude 몫).
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import naming  # noqa: E402
import sheet_to_proposals as s2p  # noqa: E402
import update_status_excel as use  # noqa: E402

SKILL_DIR = Path(__file__).resolve().parent.parent
REPO_ROOT = SKILL_DIR.parent.parent.parent
EXIT_ALREADY_DONE = 4
EXIT_NO_BASE = 5


def load_config(path=None) -> dict:
    p = Path(path) if path else SKILL_DIR / "config.json"
    return json.loads(p.read_text(encoding="utf-8"))


def repo_path(config: dict, key: str) -> Path:
    return REPO_ROOT / config.get("paths", {}).get(key, key)


def load_runs(config: dict) -> dict:
    p = repo_path(config, "runs_state")
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def save_runs(config: dict, runs: dict):
    p = repo_path(config, "runs_state")
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(runs, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def resolve_base(config: dict, report_month: str, explicit: str | None) -> Path | None:
    if explicit:
        return Path(explicit)
    return naming.latest_base(repo_path(config, "data_dir"), before=report_month)


def prev_month_label(report_month: str) -> str:
    yy, mm = naming.parse_month(report_month)
    return f"{mm - 1 if mm > 1 else 12}월"


def now_kst() -> str:
    return datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S KST")


# ---------- prepare ----------

def cmd_prepare(args) -> int:
    config = load_config(args.config)
    report_month = naming.normalize_month(args.report_month)
    base = resolve_base(config, report_month, args.base)
    work = repo_path(config, "work_dir") / naming.compact_month(report_month)
    meta = s2p.run(args.dump, report_month, work, config,
                   str(base) if base else None, args.months, args.sheet_modified_time)
    meta["work_dir"] = str(work)
    print(json.dumps(meta, ensure_ascii=False, indent=2))
    return 0


# ---------- build ----------

def placeholder_png(path: Path):
    from PIL import Image
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", (800, 200), "white").save(path)


def render_images(status_xlsx: Path, images_dir: Path, warnings: list[str], skip: bool) -> tuple[Path, Path]:
    status_png = images_dir / "status_table.png"
    summary_png = images_dir / "summary_table.png"
    images_dir.mkdir(parents=True, exist_ok=True)
    if skip or shutil.which("soffice") is None:
        warnings.append("시트 이미지 렌더링 건너뜀 (soffice 없음 또는 --skip-render) — 메일 본문 이미지는 빈 자리표시자")
        placeholder_png(status_png)
        placeholder_png(summary_png)
        return status_png, summary_png
    cmd = [sys.executable, str(SKILL_DIR / "scripts" / "render_sheet_images.py"),
           "--xlsx", str(status_xlsx), "--output-dir", str(images_dir)]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if res.returncode != 0 or not status_png.exists():
        warnings.append("시트 이미지 렌더링 실패: " + (res.stderr or res.stdout)[-800:])
        placeholder_png(status_png)
    if not summary_png.exists():
        placeholder_png(summary_png)
    return status_png, summary_png


def build_eml(config: dict, proposals: list[dict], meta: dict, status_png: Path, summary_png: Path,
              attachment: Path, out_path: Path):
    email_json = out_path.parent / "proposals_email.json"
    clean = [{k: v for k, v in p.items() if not k.startswith("_")} for p in proposals]
    email_json.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")
    recipients = SKILL_DIR / config.get("recipients", "assets/recipients.json")
    cmd = [sys.executable, str(SKILL_DIR / "scripts" / "build_email_eml.py"),
           "--recipients", str(recipients),
           "--report-month", meta["report_month"],
           "--prev-month", prev_month_label(meta["report_month"]),
           "--new-count", str(meta["new_count"]),
           "--total-count", str(meta["total_count"]),
           "--proposals-json", str(email_json),
           "--status-image", str(status_png),
           "--summary-image", str(summary_png),
           "--attachment", str(attachment),
           "--output", str(out_path)]
    subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=120)


def cmd_build(args) -> int:
    config = load_config(args.config)
    proposals, meta = s2p.proposals_from_json(args.proposals)
    report_month = naming.normalize_month(args.report_month or meta.get("report_month"))
    meta["report_month"] = report_month
    if args.new_count is not None:
        meta["new_count"] = args.new_count
    if args.total_count is not None:
        meta["total_count"] = args.total_count
    meta.setdefault("new_count", sum(int(p.get("건수", 1)) for p in proposals))
    warnings: list[str] = list(meta.get("warnings", []))

    runs = load_runs(config)
    if report_month in runs and not args.force:
        print(json.dumps({"error": "already_done", "report_month": report_month, "run": runs[report_month]},
                         ensure_ascii=False))
        return EXIT_ALREADY_DONE

    base = resolve_base(config, report_month, args.base or meta.get("base"))
    if not base or not Path(base).exists():
        print(json.dumps({"error": "no_base", "message":
                          f"{report_month} 이전 현황판이 {repo_path(config, 'data_dir')} 에 없습니다. "
                          "BEAR_아이디어_제안_현황판_YY년M월_vN.xlsx 를 data/ 에 넣어 주세요."}, ensure_ascii=False))
        return EXIT_NO_BASE
    if "total_count" not in meta:
        meta["total_count"] = (s2p.load_base_info(base)["total"] or 0) + meta["new_count"]

    out_dir = repo_path(config, "output_dir") / naming.compact_month(report_month)
    version = args.version or naming.next_version(repo_path(config, "output_dir"), report_month)
    out_dir.mkdir(parents=True, exist_ok=True)
    status_path = out_dir / naming.status_name(report_month, version)
    eml_path = out_dir / naming.email_name(report_month)

    # 1) 입력 xlsx 재생성 (Claude 가 proposals.json 을 고쳤을 수 있음)
    s2p.write_new_proposals_xlsx(proposals, out_dir / "new_proposals.xlsx", report_month)
    (out_dir / "proposals.json").write_text(
        json.dumps({"_meta": meta, "proposals": proposals}, ensure_ascii=False, indent=2), encoding="utf-8")

    # 2) 현황판 갱신
    plain = [{k: v for k, v in p.items() if not k.startswith("_")} for p in proposals]
    warnings += use.run(str(base), plain, report_month, prev_month_label(report_month),
                        meta["new_count"], meta["total_count"], str(status_path),
                        status_updates=meta.get("status_updates"))

    # 3) 이미지 → 4) .eml
    status_png, summary_png = render_images(status_path, out_dir / "images", warnings, args.skip_render)
    build_eml(config, proposals, meta, status_png, summary_png, status_path, eml_path)

    # 5) review.md 복사(있으면) + manifest
    src_review = Path(args.proposals).parent / "review.md"
    if src_review.exists():
        shutil.copyfile(src_review, out_dir / "review.md")
    unmatched = [p["팀"] for p in proposals if p.get("_unmatched")]
    attention = bool(unmatched) or any(("불일치" in w or "캐치업" in w or "실패" in w) for w in warnings)
    subject = (f"{config['notify'].get('attention_prefix', '[확인 필요]') + ' ' if attention else ''}"
               f"{config['notify'].get('subject_prefix', '[BEAR 자동화]')} {report_month} 현황판·메일 생성 완료 "
               f"(신규 {meta['new_count']}건, 총 {meta['total_count']}건)")
    manifest = {
        "report_month": report_month,
        "version": version,
        "new_count": meta["new_count"],
        "total_count": meta["total_count"],
        "base": str(base),
        "status_xlsx": str(status_path),
        "eml": str(eml_path),
        "images": [str(status_png), str(summary_png)],
        "review": str(out_dir / "review.md"),
        "unmatched_teams": unmatched,
        "status_updates": meta.get("status_updates") or [],
        "warnings": warnings,
        "attention": attention,
        "notify_subject": subject,
        "dry_run": bool(args.dry_run),
        "built_at": now_kst(),
    }

    if not args.dry_run:
        data_dir = repo_path(config, "data_dir")
        data_dir.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(status_path, data_dir / status_path.name)
        runs[report_month] = {"run_date": manifest["built_at"][:10], "version": version,
                              "new_count": meta["new_count"], "total_count": meta["total_count"],
                              "status_xlsx": status_path.name}
        save_runs(config, runs)
        manifest["data_copy"] = str(data_dir / status_path.name)

    (out_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


# ---------- CLI ----------

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=None)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("prepare", help="시트 덤프 → work/<월>/ 신규 제안 초안")
    p.add_argument("--dump", required=True)
    p.add_argument("--report-month", required=True)
    p.add_argument("--base", default=None)
    p.add_argument("--months", nargs="*", default=None)
    p.add_argument("--sheet-modified-time", default=None)
    p.set_defaults(func=cmd_prepare)

    b = sub.add_parser("build", help="proposals.json → 현황판·이미지·.eml → output/<월>/")
    b.add_argument("--proposals", required=True, help="work/<월>/proposals.json")
    b.add_argument("--report-month", default=None)
    b.add_argument("--base", default=None)
    b.add_argument("--version", type=int, default=None)
    b.add_argument("--new-count", type=int, default=None)
    b.add_argument("--total-count", type=int, default=None)
    b.add_argument("--force", action="store_true", help="runs.json 에 이미 있어도 진행 (vN 증가)")
    b.add_argument("--dry-run", action="store_true", help="data/·runs.json 을 건드리지 않음")
    b.add_argument("--skip-render", action="store_true", help="LibreOffice 렌더링 생략 (테스트)")
    b.set_defaults(func=cmd_build)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
