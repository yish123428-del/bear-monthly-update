#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BEAR monthly email .eml builder."""
import argparse, json, mimetypes, re, sys, uuid
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

PROPOSAL_BLOCK_TEMPLATE = '''&nbsp; &nbsp; &nbsp;{idx}. {bonbu}{sep_a}{sabu}{sep_b}<span style="font-weight: bold;">{team} {person}님 {count}건</span>
<ul class="list-disc flex flex-col gap-1 pl-8 mb-3">
<li class="whitespace-normal break-words pl-2"><span style="font-weight: bold;">{idea_type} {ic}건</span> ({detail})</li></ul>
'''

DEFAULT_GOOGLE_SHEET = "https://docs.google.com/spreadsheets/d/1hX4OX8PVpNA0SNdX-tqWcF3BdNRBJxz14IgUQzi5r8g/edit?usp=sharing"


def build_proposals_html(proposals_text_file=None, proposals_json=None):
    if proposals_text_file:
        return Path(proposals_text_file).read_text(encoding="utf-8")
    if not proposals_json:
        return ""
    parts = []
    for idx, p in enumerate(proposals_json, 1):
        # 본부표시/팀표시 가 있으면 메일 본문에는 그것을 쓴다 (예: 영업부는 본부 생략, '병원경기2사무소')
        bonbu = p["본부표시"] if "본부표시" in p else (p.get("본부") or "")
        sabu = p.get("사업부") or ""
        team = p.get("팀표시") or p["팀"]
        sep_a = " " if bonbu else ""
        sep_b = " " if sabu else ""
        if p.get("이메일표기"):
            li_text = p["이메일표기"]
            html = ('&nbsp; &nbsp; &nbsp;' + str(idx) + '. ' + bonbu + sep_a + sabu + sep_b
                + '<span style="font-weight: bold;">' + team + ' ' + p["제안자"]
                + '님 ' + str(p["건수"]) + '건</span>\n'
                + '<ul class="list-disc flex flex-col gap-1 pl-8 mb-3">\n'
                + '<li class="whitespace-normal break-words pl-2">' + li_text + '</li></ul>\n')
        else:
            html = PROPOSAL_BLOCK_TEMPLATE.format(
                idx=idx, bonbu=bonbu, sep_a=sep_a, sabu=sabu, sep_b=sep_b,
                team=team, person=p["제안자"], count=p["건수"],
                idea_type=p["아이디어유형"], ic=p["건수"], detail=p["제안내용"],
            )
        parts.append(html)
    return "".join(parts)


def render_template(template_html, ctx):
    out = template_html
    for k, v in ctx.items():
        out = out.replace("{{" + k + "}}", str(v))
    return out


def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=None)
    ap.add_argument("--recipients", required=True)
    ap.add_argument("--report-month", required=True)
    ap.add_argument("--prev-month", required=True)
    ap.add_argument("--new-count", type=int, required=True)
    ap.add_argument("--total-count", type=int, required=True)
    ap.add_argument("--proposals-text-file")
    ap.add_argument("--proposals-json")
    ap.add_argument("--status-image", required=True)
    ap.add_argument("--summary-image", required=True)
    ap.add_argument("--attachment", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--report-month-label", default=None)
    return ap.parse_args()


def main():
    args = parse_args()
    template_path = Path(args.template) if args.template else (
        Path(__file__).resolve().parent.parent / "assets" / "email_template.html"
    )
    template = template_path.read_text(encoding="utf-8")
    with open(args.recipients, "r", encoding="utf-8") as f:
        rcpt = json.load(f)

    proposals_json = None
    if args.proposals_json:
        with open(args.proposals_json, "r", encoding="utf-8") as f:
            proposals_json = json.load(f)
    proposals_html = build_proposals_html(args.proposals_text_file, proposals_json)

    report_label = args.report_month_label
    if not report_label:
        m = re.match(r"(\d+)월", args.prev_month)
        if m:
            n = int(m.group(1)) + 1
            if n > 12: n = 1
            report_label = str(n) + "월"
        else:
            report_label = args.prev_month

    status_cid_id = "status-" + uuid.uuid4().hex[:16] + "@bear.local"
    summary_cid_id = "summary-" + uuid.uuid4().hex[:16] + "@bear.local"

    body_html = render_template(template, {
        "REPORT_MONTH_LABEL": report_label, "PREV_MONTH": args.prev_month,
        "NEW_COUNT": args.new_count, "TOTAL_COUNT": args.total_count,
        "NEW_PROPOSALS_BLOCK": proposals_html,
        "GOOGLE_SHEET_URL": rcpt.get("google_sheet_url", DEFAULT_GOOGLE_SHEET),
        "STATUS_IMAGE_CID": status_cid_id, "SUMMARY_IMAGE_CID": summary_cid_id,
        "LOGO_URL": rcpt["signature"]["logo_url"],
        "SIG_NAME_KR": rcpt["signature"]["name_kr"],
        "SIG_NAME_EN": rcpt["signature"]["name_en"],
        "SIG_TEAM_KR": rcpt["signature"]["team_kr"],
        "SIG_TEAM_EN": rcpt["signature"]["team_en"],
        "SIG_ADDRESS": rcpt["signature"]["address"],
        "SIG_MOBILE": rcpt["signature"]["mobile"],
        "SIG_EMAIL": rcpt["signature"]["email"],
    })

    plain = ("안녕하세요, 신제품기획팀 이성희입니다.\n\n"
        + "ㅇ " + report_label + " 1주차 접수현황 : "
        + args.prev_month + " " + str(args.new_count) + "건 총 " + str(args.total_count) + "건 접수\n\n"
        "ㅇ 사업팀 별 아이디어 제안 현황 / 누적 제안 요약은 첨부 파일 또는 본문 이미지 참고.\n\n"
        "이성희 / 신제품기획팀\nshlee203@daewoong.co.kr\n")

    msg = EmailMessage()
    msg["Subject"] = "[공유] BEAR 아이디어 제안 현황 공유(" + report_label + " 1주차)"
    msg["From"] = formataddr((rcpt["from"]["name"], rcpt["from"]["email"]))
    msg["To"] = ", ".join(formataddr((r["name"], r["email"])) for r in rcpt.get("to", []))
    if rcpt.get("cc"):
        msg["Cc"] = ", ".join(formataddr((r["name"], r["email"])) for r in rcpt["cc"])
    now = datetime.now(timezone(timedelta(hours=9)))
    msg["Date"] = formatdate(now.timestamp(), localtime=True)
    msg["Message-ID"] = make_msgid(domain="daewoong.co.kr")

    msg.set_content(plain)
    msg.add_alternative(body_html, subtype="html")

    html_part = msg.get_payload()[1]
    for cid_id, img_path in [(status_cid_id, args.status_image), (summary_cid_id, args.summary_image)]:
        with open(img_path, "rb") as f:
            data = f.read()
        html_part.add_related(data, maintype="image", subtype="png", cid="<" + cid_id + ">")
    for part in html_part.walk():
        if part.get_content_type().startswith("image/"):
            if "Content-Disposition" in part:
                del part["Content-Disposition"]
            part.add_header("Content-Disposition", "inline")

    with open(args.attachment, "rb") as f:
        att_data = f.read()
    ctype, _ = mimetypes.guess_type(args.attachment)
    if ctype is None:
        ctype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    maintype, subtype = ctype.split("/", 1)
    msg.add_attachment(att_data, maintype=maintype, subtype=subtype, filename=Path(args.attachment).name)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "wb") as f:
        f.write(bytes(msg))
    print("v saved: " + str(out))


if __name__ == "__main__":
    main()
