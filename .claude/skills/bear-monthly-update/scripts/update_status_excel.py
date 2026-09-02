#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BEAR 아이디어 제안 현황판 업데이트.

이전 달 현황판(.xlsx)을 베이스로 삼아, 신규 제안 데이터(.xlsx)를 반영해서
새 현황판 파일을 생성한다.

업데이트 대상:
- 시트 '1. 제안 현황': 제목 월 갱신, 누계 진행 결과 문구 갱신,
  팀별 행의 해당 월 카운트 +1, 합계/누계 갱신.
- 시트 '2. 누적 제안 요약(26.01~)': 제목 월 갱신, 신규 제안 행 추가.
"""
import argparse
import json
import re
import shutil
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter


# ---------- 신규 제안 파싱 ----------

NEW_PROP_HEADERS = {
    "제안 연월": ["제안 연월", "제안연월", "연월"],
    "본부":     ["본부"],
    "사업부":   ["사업부"],
    "팀":       ["팀/사무소", "팀", "사무소", "팀/사업부"],
    "제안자":   ["제안자", "이름"],
    "건수":     ["건수", "수량"],
    "제안유형": ["제안 유형", "제안유형"],
    "아이디어유형": ["아이디어 유형", "아이디어유형"],
    "제안내용": ["제안 내용", "제안내용", "내용"],
    "검토결과": ["검토 결과", "검토결과", "결과"],
    "이메일표기": ["이메일 표기", "이메일표기", "표기"],
}


def find_header_row(ws):
    for r in range(1, 6):
        row_vals = [c.value for c in ws[r]]
        if any(v and "제안" in str(v) and "연월" in str(v) for v in row_vals):
            return r
    return 1


def map_columns(ws, header_row):
    headers = {c.value: c.column for c in ws[header_row] if c.value}
    mapping = {}
    for key, candidates in NEW_PROP_HEADERS.items():
        for cand in candidates:
            for h, col in headers.items():
                if str(h).strip() == cand:
                    mapping[key] = col
                    break
            if key in mapping:
                break
    return mapping


def load_new_proposals(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header_row = find_header_row(ws)
    cols = map_columns(ws, header_row)
    required = ["제안 연월", "팀", "제안자", "건수", "제안유형", "아이디어유형", "제안내용", "검토결과"]
    missing = [r for r in required if r not in cols]
    if missing:
        raise SystemExit(
            "신규 제안 파일에 필수 컬럼이 없음: " + str(missing)
            + ". 발견된 컬럼: " + str(list(cols.keys()))
        )

    proposals = []
    for r in range(header_row + 1, ws.max_row + 1):
        cell = lambda key: ws.cell(row=r, column=cols[key]).value if key in cols else None
        team = cell("팀")
        if not team:
            continue
        p = {
            "월": cell("제안 연월"),
            "본부": (cell("본부") or "").strip() if isinstance(cell("본부"), str) else (cell("본부") or ""),
            "사업부": cell("사업부") or "",
            "팀": str(team).strip(),
            "제안자": cell("제안자"),
            "건수": int(cell("건수") or 1),
            "제안유형": cell("제안유형"),
            "아이디어유형": cell("아이디어유형"),
            "제안내용": cell("제안내용"),
            "검토결과": cell("검토결과"),
            "이메일표기": cell("이메일표기"),
        }
        proposals.append(p)
    return proposals


# ---------- 시트 매칭 유틸 ----------

def normalize(name):
    if not name:
        return ""
    return re.sub(r"\s+", "", str(name))


def find_team_row(ws, team_name, search_col=3, start_row=10, end_row=None):
    end_row = end_row or ws.max_row
    target = normalize(team_name)
    if not target:
        return None
    exact = None
    norm_match = []
    partial_match = []
    for r in range(start_row, end_row + 1):
        v = ws.cell(row=r, column=search_col).value
        if not v:
            continue
        sv = str(v)
        if sv == team_name:
            exact = r
            break
        nv = normalize(sv)
        if nv == target:
            norm_match.append(r)
        elif target in nv or nv in target:
            partial_match.append(r)
    if exact is not None:
        return exact
    if norm_match:
        return norm_match[0]
    if partial_match:
        return partial_match[0]
    return None


# ---------- 시트 1 업데이트 ----------

def get_month_columns(ws, header_row, month_text):
    """헤더 행 또는 그 다음 행에서 month_text가 있는 컬럼 인덱스 리스트.
    시트1 MKT/영업부 섹션은 2행 헤더 (예: row 31 '제안 합계/26년', row 32 '1월~12월').
    그래서 header_row와 header_row+1 둘 다 스캔."""
    cols = []
    for r_off in (0, 1):
        for cell in ws[header_row + r_off]:
            if cell.value and str(cell.value).strip() == month_text:
                cols.append(cell.column)
        if cols:
            break
    return cols


def find_total_column(ws, header_row):
    for r_off in (0, 1):
        for cell in ws[header_row + r_off]:
            if cell.value and "누계" in str(cell.value):
                return cell.column
    return None


def find_subtotal_column(ws, header_row):
    """'제안 합계' / '리스트업' 등 섹션 합계 컬럼."""
    for r_off in (0, 1):
        for cell in ws[header_row + r_off]:
            v = cell.value
            if not v:
                continue
            sv = str(v).strip()
            if "제안 합계" in sv or sv == "리스트업":
                return cell.column
    return None


def update_sheet1_count(ws, team_name, month_text, n, debug_section="", header_labels=None):
    """team_name 행의 month_text 컬럼에 +n.

    header_labels 를 주면 그 라벨('팀'/'사업부'/'사무소')로 시작하는 섹션만 본다.
    숨김 처리된 헤더 행(예: 상단 구버전 표)은 건너뛴다.
    """
    # 모든 가능한 헤더 행 위치 수집 ('팀'/'사업부'/'사무소' 단어가 들어있는 셀)
    matched = []
    for r in range(1, ws.max_row + 1):
        for cell in ws[r]:
            if cell.value:
                sv = str(cell.value).strip()
                if sv in ("팀", "사업부", "사무소"):
                    matched.append(r)
                    break

    updated_anywhere = False
    warnings = []
    for header_row in matched:
        next_h = next((h for h in matched if h > header_row), ws.max_row + 1)
        rd = ws.row_dimensions.get(header_row)
        if rd is not None and rd.hidden:
            continue
        team_col = None
        label = None
        for cell in ws[header_row]:
            v = cell.value
            if not v:
                continue
            sv = str(v).strip()
            if sv in ("팀", "사업부", "사무소"):
                team_col = cell.column
                label = sv
                break
        if team_col is None:
            continue
        if header_labels and label not in header_labels:
            continue
        # 데이터 시작은 header_row + 2 또는 +3 (서브헤더와 빈줄 고려)
        data_start = header_row + 1
        # 만약 header_row+1이 서브헤더(월 이름 포함)면 data_start = header_row + 3
        sub_row = ws[header_row + 1] if header_row + 1 <= ws.max_row else []
        if any(v and re.match(r"\d+월", str(v).strip()) for v in (c.value for c in sub_row)):
            data_start = header_row + 3
        row_idx = find_team_row(ws, team_name, search_col=team_col,
                                 start_row=data_start, end_row=next_h - 1)
        if row_idx is None:
            continue
        month_cols = get_month_columns(ws, header_row, month_text)
        if not month_cols:
            warnings.append(
                "[" + debug_section + "] " + team_name +
                " 행은 찾았으나 헤더에 '" + month_text + "' 컬럼 없음 (header_row="
                + str(header_row) + ")"
            )
            continue
        mcol = month_cols[0]
        cur = ws.cell(row=row_idx, column=mcol).value
        try:
            cur = int(cur) if cur not in (None, "", "-") else 0
        except (TypeError, ValueError):
            cur = 0
        ws.cell(row=row_idx, column=mcol).value = cur + n

        # 누계/제안 합계 컬럼은 보통 =SUM(...) 수식이므로 건드리지 않음.
        # 수식이 아니라 정수면 +n으로 갱신.
        for col_finder, label in [(find_total_column, "누계"),
                                  (find_subtotal_column, "제안 합계")]:
            col = col_finder(ws, header_row)
            if not col or col == mcol:
                continue
            cur = ws.cell(row=row_idx, column=col).value
            if isinstance(cur, str) and cur.startswith("="):
                continue  # 수식 — Excel이 재계산하게 둠
            try:
                cur_int = int(cur) if cur not in (None, "", "-") else 0
            except (TypeError, ValueError):
                continue
            ws.cell(row=row_idx, column=col).value = cur_int + n

        updated_anywhere = True
        print(
            "  v [" + (debug_section or "sheet1") + "] '" + team_name
            + "' row=" + str(row_idx) + " col="
            + get_column_letter(mcol) + "(" + month_text + ") +" + str(n),
            file=sys.stderr,
        )

    if not updated_anywhere:
        warnings.append("매칭 실패: 팀 '" + team_name + "' (월=" + month_text + ")")
    return updated_anywhere, warnings


def update_sheet1_title_and_total(ws, report_month_label, prev_month, new_count, total_count):
    for r in range(1, 35):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if not v or not isinstance(v, str):
                continue
            if "사업팀 별 아이디어 현황판" in v or "아이디어 현황판" in v:
                ws.cell(row=r, column=c).value = re.sub(
                    r"\(\d+년\s*\d+월\)", "(" + report_month_label + ")", v
                )
            elif "누계 진행 결과" in v:
                ws.cell(row=r, column=c).value = (
                    "[ 누계 진행 결과 _" + prev_month + " "
                    + str(new_count) + "건 총 " + str(total_count) + "건 ] "
                )


# ---------- 시트 2 ----------

def update_sheet2(ws, proposals, report_month_label):
    for r in range(1, 5):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "누적 제안 요약" in v and re.search(r"\(\d+년\s*\d+월\)", v):
                ws.cell(row=r, column=c).value = re.sub(
                    r"\(\d+년\s*\d+월\)", "(" + report_month_label + ")", v
                )

    header_row = None
    for r in range(1, 10):
        vals = [c.value for c in ws[r]]
        if "제안 연월" in vals:
            header_row = r
            break
    if header_row is None:
        raise RuntimeError("시트 2 헤더 '제안 연월'을 찾을 수 없음")

    headers_map = {c.value: c.column for c in ws[header_row] if c.value}
    first_col = headers_map.get("제안 연월")
    insert_row = header_row + 1
    while ws.cell(row=insert_row, column=first_col).value not in (None, ""):
        insert_row += 1

    last_data_row = insert_row - 1
    field_to_header = {
        "월": "제안 연월",
        "팀": "사업팀/사무소",
        "제안자": "제안자",
        "제안유형": "제안 유형",
        "아이디어유형": "아이디어 유형",
        "제안내용": "제안 내용",
        "검토결과": "검토 결과",
    }

    for p in proposals:
        for fkey, hkey in field_to_header.items():
            col = headers_map.get(hkey)
            if not col:
                continue
            ws.cell(row=insert_row, column=col).value = p.get(fkey)
            if last_data_row > header_row:
                src = ws.cell(row=last_data_row, column=col)
                dst = ws.cell(row=insert_row, column=col)
                if src.has_style:
                    dst.font = copy(src.font)
                    dst.fill = copy(src.fill)
                    dst.border = copy(src.border)
                    dst.alignment = copy(src.alignment)
                    dst.number_format = src.number_format
        insert_row += 1


# ---------- main ----------

def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", required=True)
    ap.add_argument("--new-proposals", required=True)
    ap.add_argument("--report-month", required=True)
    ap.add_argument("--prev-month", required=True)
    ap.add_argument("--new-count", type=int, required=True)
    ap.add_argument("--total-count", type=int, required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--team-mapping")
    return ap.parse_args()


def run(base, new_proposals, report_month, prev_month, new_count, total_count,
        output, alias=None):
    """베이스 현황판에 신규 제안을 반영해 output 에 저장. 경고 문자열 목록을 돌려준다.

    new_proposals 는 .xlsx 경로 또는 이미 로드된 제안 dict 목록.
    """
    if isinstance(new_proposals, (str, Path)):
        proposals = load_new_proposals(new_proposals)
    else:
        proposals = list(new_proposals)
    print("신규 제안 " + str(len(proposals)) + "건 로드", file=sys.stderr)

    alias = alias or {}
    for p in proposals:
        if p["팀"] in alias:
            p["팀"] = alias[p["팀"]]

    shutil.copyfile(base, output)
    wb = openpyxl.load_workbook(output)
    all_warnings = []

    if "1. 제안 현황" in wb.sheetnames:
        ws1 = wb["1. 제안 현황"]
        update_sheet1_title_and_total(
            ws1, report_month, prev_month, new_count, total_count
        )
        for p in proposals:
            m = re.search(r"(\d+월)", str(p["월"]))
            if not m:
                all_warnings.append("월 추출 실패: " + str(p["월"]))
                continue
            month_text = m.group(1)
            ok, warns = update_sheet1_count(
                ws1, p["팀"], month_text, p["건수"], debug_section="제안현황"
            )
            all_warnings.extend(warns)
            # 영업부 제안은 '2. 영업부(사업부별)' 표의 사업부 행도 +n (현황판 운영 관행)
            division = str(p.get("사업부") or "").strip()
            if division and str(p.get("본부") or "").strip() != "MKT":
                ok2, warns2 = update_sheet1_count(
                    ws1, division, month_text, p["건수"], debug_section="사업부별",
                    header_labels=("사업부",)
                )
                all_warnings.extend(warns2)
    else:
        all_warnings.append("시트 '1. 제안 현황' 없음")

    sheet2_name = next(
        (n for n in wb.sheetnames if n.startswith("2.") and "누적" in n), None
    )
    if sheet2_name:
        update_sheet2(wb[sheet2_name], proposals, report_month)
    else:
        all_warnings.append("시트 '2. 누적 제안 요약(...)' 없음")

    wb.save(output)
    for w in all_warnings:
        print("[warn] " + w, file=sys.stderr)
    print("v 저장: " + str(output), file=sys.stderr)
    return all_warnings


def main():
    args = parse_args()
    alias = {}
    if args.team_mapping:
        with open(args.team_mapping, "r", encoding="utf-8") as f:
            alias = json.load(f)
    run(args.base, args.new_proposals, args.report_month, args.prev_month,
        args.new_count, args.total_count, args.output, alias)


if __name__ == "__main__":
    main()
