#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
현황판 .xlsx의 시트 1·2를 PNG 이미지로 렌더링.

LibreOffice headless로 .xlsx → PDF → PNG 변환. 이메일 본문에 인라인 삽입할
두 이미지를 만든다:
  - status_table.png  (시트 '1. 제안 현황')
  - summary_table.png (시트 '2. 누적 제안 요약(26.01~)')

전략: 다른 시트는 hidden 처리하고 단일 시트만 PDF로 변환한 뒤 1페이지를 PNG로.
페이지 가로 1장에 맞도록 fit_to_page를 강제로 켠다.
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))


def ensure_korean_fonts():
    """assets/fonts/에 번들된 NanumGothic을 사용자 ~/.fonts에 설치 (없을 때만).
    LibreOffice 렌더링에서 한글 깨짐 방지."""
    user_fonts = Path.home() / ".fonts"
    targets = ["NanumGothic.ttf", "NanumGothicBold.ttf"]
    if all((user_fonts / f).exists() for f in targets):
        return
    src_dir = Path(__file__).resolve().parent.parent / "assets" / "fonts"
    user_fonts.mkdir(parents=True, exist_ok=True)
    copied = False
    for fname in targets:
        src = src_dir / fname
        dst = user_fonts / fname
        if src.exists() and not dst.exists():
            shutil.copyfile(src, dst)
            copied = True
    if copied:
        try:
            subprocess.run(["fc-cache", "-f", str(user_fonts)],
                           check=False, capture_output=True, timeout=30)
        except FileNotFoundError:
            pass


def libreoffice_to_pdf(xlsx_path, out_dir, timeout=180):
    cmd = [
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(out_dir), str(xlsx_path),
    ]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    if res.returncode != 0:
        print("soffice stderr:", res.stderr, file=sys.stderr)
        raise RuntimeError("LibreOffice PDF 변환 실패")
    pdf_path = Path(out_dir) / (Path(xlsx_path).stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("PDF 파일 못 찾음: " + str(pdf_path))
    return pdf_path


def pdf_pages_to_png(pdf_path, prefix, dpi=180):
    """PDF의 모든 페이지를 PNG로. prefix-1.png, prefix-2.png ... 반환.

    PyMuPDF(pymupdf)를 우선 사용하고, 없으면 pdftoppm(poppler)으로 폴백한다.
    (Claude Code 샌드박스에는 pdftoppm 이 없다.)"""
    prefix = Path(prefix)
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf  # 구버전 이름
        except ImportError:
            pymupdf = None
    if pymupdf is not None:
        doc = pymupdf.open(str(pdf_path))
        out = []
        for i, page in enumerate(doc, 1):
            p = prefix.parent / (prefix.name + "-" + str(i) + ".png")
            page.get_pixmap(dpi=dpi).save(str(p))
            out.append(p)
        doc.close()
        return out
    cmd = ["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(prefix)]
    subprocess.run(cmd, check=True, timeout=120)
    return sorted(prefix.parent.glob(prefix.name + "-*.png"))


def join_pngs_vertically(png_paths, out_path):
    """여러 PNG를 세로로 이어붙임. PIL 사용."""
    from PIL import Image
    imgs = [Image.open(p) for p in png_paths]
    width = max(im.width for im in imgs)
    height = sum(im.height for im in imgs)
    canvas = Image.new("RGB", (width, height), "white")
    y = 0
    for im in imgs:
        canvas.paste(im, (0, y))
        y += im.height
    canvas.save(out_path)


def crop_whitespace(png_path):
    """페이지 하단 여백 자르기 (PIL 사용). 흰색 행이 연속되는 부분 제거."""
    from PIL import Image, ImageChops
    im = Image.open(png_path).convert("RGB")
    bg = Image.new("RGB", im.size, "white")
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox:
        # 약간의 여백 추가 (10px)
        l, t, r, b = bbox
        l = max(0, l - 10)
        t = max(0, t - 10)
        r = min(im.width, r + 10)
        b = min(im.height, b + 10)
        im.crop((l, t, r, b)).save(png_path)


def render_sheet(src_xlsx, sheet_name, out_png, tmpdir, max_pages=2):
    """단일 시트를 PNG로. 다른 시트는 hidden, fit-to-page 활성화."""
    tmp_xlsx = Path(tmpdir) / "render_target.xlsx"
    shutil.copyfile(src_xlsx, tmp_xlsx)
    wb = openpyxl.load_workbook(tmp_xlsx)
    if sheet_name not in wb.sheetnames:
        raise RuntimeError("시트 '" + sheet_name + "' 없음")

    # 다른 시트를 삭제 (LibreOffice는 sheet_state="hidden"을 무시할 때가 있음)
    for n in list(wb.sheetnames):
        if n != sheet_name:
            del wb[n]
    wb.active = wb.sheetnames.index(sheet_name)

    ws = wb[sheet_name]
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 가로만 1장에 맞춤
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = False

    wb.save(tmp_xlsx)

    pdf = libreoffice_to_pdf(tmp_xlsx, tmpdir)
    prefix = Path(tmpdir) / "page"
    pages = pdf_pages_to_png(pdf, prefix, dpi=180)
    if not pages:
        raise RuntimeError("PNG 페이지 추출 실패")
    use_pages = pages[:max_pages]
    if len(use_pages) == 1:
        shutil.copyfile(use_pages[0], out_png)
    else:
        join_pngs_vertically(use_pages, out_png)
    crop_whitespace(out_png)
    return out_png


# ---------- Chromium(HTML) 백엔드 ----------

CHROME_CANDIDATES = ["/opt/pw-browsers/chromium*/chrome-linux*/chrome",
                     "/opt/pw-browsers/chromium*/chrome-linux*/headless_shell"]


def find_chrome():
    import glob
    for pat in CHROME_CANDIDATES:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    for name in ("chromium", "chromium-browser", "google-chrome", "chrome"):
        p = shutil.which(name)
        if p:
            return p
    return None


def soffice_has_calc():
    """libreoffice-calc 가 설치돼 있는지 (core 만 있으면 시트를 못 연다)."""
    if shutil.which("soffice") is None:
        return False
    return any(Path(d).glob("libscalclo.so") for d in ("/usr/lib/libreoffice/program",))


def render_sheet_html(src_xlsx, sheet_name, out_png, tmpdir, stop_regex=None, scale=2):
    """openpyxl → HTML → Chromium 스크린샷 → 여백 자르기."""
    import xlsx_html
    wb = openpyxl.load_workbook(src_xlsx)
    if sheet_name not in wb.sheetnames:
        raise RuntimeError("시트 '" + sheet_name + "' 없음")
    ws = wb[sheet_name]
    font_dir = Path(__file__).resolve().parent.parent / "assets" / "fonts"
    doc, warnings = xlsx_html.sheet_to_html(ws, font_dir, stop_regex=stop_regex)
    for w in warnings:
        print("[warn] " + w, file=sys.stderr)
    html_path = Path(tmpdir) / "sheet.html"
    html_path.write_text(doc, encoding="utf-8")
    width, height = xlsx_html.estimate_size(ws, html_doc=doc)
    chrome = find_chrome()
    if not chrome:
        raise RuntimeError("Chromium 을 찾지 못함")
    cmd = [chrome, "--headless=new", "--no-sandbox", "--disable-gpu", "--hide-scrollbars",
           "--force-device-scale-factor=" + str(scale),
           "--screenshot=" + str(out_png), "--window-size=" + str(width) + "," + str(height),
           "file://" + str(html_path)]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if not Path(out_png).exists():
        print(res.stderr[-1500:], file=sys.stderr)
        raise RuntimeError("Chromium 스크린샷 실패")
    crop_whitespace(out_png)
    return out_png


def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--output-dir", required=True)
    ap.add_argument("--backend", choices=["auto", "html", "soffice"], default="auto",
                    help="auto: Chromium(HTML) 우선, 없으면 LibreOffice Calc")
    ap.add_argument("--status-stop-regex", default=r"사무소별|^사무소$",
                    help="[html] 시트1에서 이 패턴이 나오는 행 앞에서 끊음 (기본: '3. 영업부 사무소별' 섹션 제외 — soffice 2페이지와 동일 범위)")
    ap.add_argument("--max-pages-status", type=int, default=2,
                    help="[soffice] 시트1을 몇 페이지까지 합칠지 (기본 2 — MKT + 영업부 사업부별 까지)")
    ap.add_argument("--max-pages-summary", type=int, default=2,
                    help="[soffice] 시트2를 몇 페이지까지 합칠지 (기본 2 — 검토결과 컬럼 spillover 대응)")
    return ap.parse_args()


def main():
    args = parse_args()
    ensure_korean_fonts()
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    backend = args.backend
    if backend == "auto":
        backend = "html" if find_chrome() else "soffice"
    if backend == "soffice" and not soffice_has_calc():
        raise SystemExit("LibreOffice Calc 가 없어 시트를 렌더링할 수 없습니다 (apt-get install libreoffice-calc) — "
                         "또는 Chromium 이 있는 환경에서 --backend html 을 쓰세요.")
    print("renderer: " + backend, file=sys.stderr)

    # 시트 이름 자동 탐지
    wb = openpyxl.load_workbook(args.xlsx, read_only=True)
    sheet1 = next((n for n in wb.sheetnames if n.startswith("1.") and "제안" in n), "1. 제안 현황")
    sheet2 = next((n for n in wb.sheetnames if n.startswith("2.") and "누적" in n), None)
    wb.close()

    with tempfile.TemporaryDirectory() as tmp:
        if backend == "html":
            render_sheet_html(args.xlsx, sheet1, out_dir / "status_table.png", tmp,
                              stop_regex=args.status_stop_regex)
        else:
            render_sheet(args.xlsx, sheet1, out_dir / "status_table.png",
                         tmp, max_pages=args.max_pages_status)
    print("v " + str(out_dir / "status_table.png"))

    if sheet2:
        with tempfile.TemporaryDirectory() as tmp:
            if backend == "html":
                render_sheet_html(args.xlsx, sheet2, out_dir / "summary_table.png", tmp)
            else:
                render_sheet(args.xlsx, sheet2, out_dir / "summary_table.png",
                             tmp, max_pages=args.max_pages_summary)
        print("v " + str(out_dir / "summary_table.png"))


if __name__ == "__main__":
    main()
