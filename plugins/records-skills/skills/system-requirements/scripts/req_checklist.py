#!/usr/bin/env python3
"""기록관리 관련 시스템 기능요건 점검표: 국가기록원 표준의 기능요건(필수 M·해당시 필수 CM·선택 O)을 점검표로 만들고,
기관이 채운 이행 결과를 받아 필수요건 충족률과 미충족 목록을 장·절별로 낸다. 시스템 구축·고도화 과업지시서, 감리, 기관평가 증빙 준비에 쓴다.
근거(국가기록원, 2026-11-20 내려받아 원문에서 요건 문장을 옮김 — assets/req_*.csv):
  NAK 6:2022(v1.5) 기록관리시스템 기능 요건 · NAK 7:2022(v1.5) 영구기록관리시스템 기능 요건
  NAK 20:2022(v1.3) 정부산하공공기관 등의 기록관리를 위한 시스템 기능요건 · NAK 19-1:2012 전자기록생산시스템 기록관리 기능요건
  NAK 19-2:2013 제2부 전자문서시스템 · NAK 19-3:2015 제3부 업무관리시스템
  (19-1·19-2는 원문에 M/O 표시가 없어 '~하여야 한다'를 M, '~할 수 있다'를 O로 나눔)

사용:  python req_checklist.py --system 기록관리시스템 -o 산출/기능요건_점검표.xlsx                 (빈 점검표)
       python req_checklist.py --system 기록관리시스템 --results 채운점검표.xlsx -o 산출/기능요건_결과.xlsx  (결과 집계)
시스템: 기록관리시스템(NAK 6) · 영구기록관리시스템(NAK 7) · 정부산하(NAK 20) · 생산시스템(NAK 19-1) · 전자문서시스템(NAK 19-2) · 업무관리시스템(NAK 19-3)
결과 열: 번호, 이행(충족·부분·미충족·해당없음), 근거(화면·문서·시험 결과)
원칙:  요건 문장은 원문 그대로다. 이행 판정은 기관·사업자가 시험해서 채운다. 원문 추출이 깨진 문장이 있으면 원문 PDF로 확인한다.
"""
import argparse
import csv
import os
import re
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
SYS = {"기록관리시스템": "6", "영구기록관리시스템": "7", "정부산하": "20", "생산시스템": "19-1", "전자문서시스템": "19-2", "업무관리시스템": "19-3"}


def load(code):
    with open(os.path.join(HERE, "..", "assets", f"req_{code}.csv"), encoding="utf-8") as f:
        return list(csv.DictReader(f))


def read_results(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10]) if "번호" in r and any(c.startswith("이행") for c in r)), None)
    if hi is None:
        return {}
    h = rows[hi]
    ni, ii = h.index("번호"), next(i for i, c in enumerate(h) if c.startswith("이행"))
    gi = next((i for i, c in enumerate(h) if c.startswith("근거")), None)
    return {r[ni]: (r[ii], r[gi] if gi is not None and gi < len(r) else "") for r in rows[hi + 1:] if len(r) > ni and r[ni]}


def norm(v):
    v = re.sub(r"\s", "", v or "")
    return "충족" if v in ("충족", "Y", "O", "예", "이행") else "부분" if "부분" in v else "미충족" if v in ("미충족", "N", "X", "아니오", "미이행") \
        else "해당없음" if "해당" in v else ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--system", required=True, choices=list(SYS))
    ap.add_argument("--results")
    ap.add_argument("--only-m", action="store_true", help="필수(M·CM)만")
    ap.add_argument("-o", "--out", default="산출/기능요건_점검표.xlsx")
    a = ap.parse_args()
    reqs = load(SYS[a.system])
    if a.only_m:
        reqs = [r for r in reqs if r["구분"] in ("M", "CM")]
    res = read_results(a.results) if a.results else {}
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검표"
    ws.append(["표준", "번호", "장", "절", "요건(원문)", "구분", "이행", "근거"])
    dv = DataValidation(type="list", formula1='"충족,부분,미충족,해당없음"', allow_blank=True)
    ws.add_data_validation(dv)
    stat = defaultdict(lambda: defaultdict(int))
    miss = []
    for r in reqs:
        st, ev = res.get(r["번호"], ("", ""))
        st = norm(st)
        ws.append([r["표준"], r["번호"], r["장"], r["절"], r["요건"], r["구분"], st, ev])
        dv.add(f"G{ws.max_row}")
        ws.cell(ws.max_row, 5).alignment = Alignment(wrap_text=True, vertical="top")
        if a.results:
            stat[r["장"]][(r["구분"], st or "미기재")] += 1
            if r["구분"] in ("M", "CM") and st in ("미충족", "부분", ""):
                miss.append(r | {"이행": st or "미기재", "근거": ev})
                ws.cell(ws.max_row, 7).fill = red
    sm = wb.create_sheet("요약", 0)
    total_m = sum(1 for r in reqs if r["구분"] in ("M", "CM"))
    sm.append([f"{a.system} 기능요건 — {reqs[0]['표준'] if reqs else ''}", ""])
    sm.append(["요건 수", f"{len(reqs)}(필수 {total_m} · 선택 {len(reqs) - total_m})"])
    if a.results:
        ok_m = sum(1 for r in reqs if r["구분"] in ("M", "CM") and norm(res.get(r["번호"], ("", ""))[0]) in ("충족", "해당없음"))
        sm.append(["필수요건 충족(해당없음 포함)", f"{ok_m}/{total_m} ({ok_m / max(total_m, 1):.1%})"])
        sm.append(["필수요건 미충족·부분·미기재", len(miss)])
        sm.append([])
        sm.append(["장", "필수 충족", "필수 부분", "필수 미충족", "필수 미기재", "선택 충족"])
        for ch, d in stat.items():
            sm.append([ch, d[("M", "충족")] + d[("CM", "충족")], d[("M", "부분")] + d[("CM", "부분")], d[("M", "미충족")] + d[("CM", "미충족")],
                       d[("M", "미기재")] + d[("CM", "미기재")], d[("O", "충족")]])
        ms = wb.create_sheet("필수 미충족")
        ms.append(["번호", "장", "절", "요건", "이행", "근거"])
        for r in miss:
            ms.append([r["번호"], r["장"], r["절"], r["요건"], r["이행"], r["근거"]])
        ms.column_dimensions["D"].width = 90
    else:
        sm.append(["쓰는 법", "'점검표'의 이행 칸(충족·부분·미충족·해당없음)과 근거를 채운 뒤 --results로 다시 돌리면 충족률을 냄"])
    sm.append([])
    sm.append(["구분", "M 필수 · CM 해당시 필수(시스템 유형·기관 환경에 따라) · O 선택"])
    for sh, widths in ((ws, (14, 8, 18, 18, 90, 6, 10, 30)), (sm, (30, 60, 12, 12, 12, 12))):
        sh["A1"].font = Font(bold=True)
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.system} 요건 {len(reqs)}(필수 {total_m})" + (f" · 필수 미충족·부분·미기재 {len(miss)}" if a.results else "") + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
