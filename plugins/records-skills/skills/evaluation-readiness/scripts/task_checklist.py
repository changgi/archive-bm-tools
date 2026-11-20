#!/usr/bin/env python3
"""기록관리 수행업무 점검표: 「2026년 기록물관리 지침」의 <참고> 처리과 기록관리 수행업무 리스트(28개)·기록관 기록관리 수행업무 리스트(24개)를
수행 여부(O·X) 칸이 있는 엑셀로 만들고, 채운 결과로 수행률과 미수행 항목·그 일을 돕는 킷 스킬을 정리한다.

사용:  python task_checklist.py --role 처리과|기록관|전체 [--org 기관·부서명] -o 산출/수행업무_점검표.xlsx
       python task_checklist.py --results 채운_점검표.xlsx -o 산출/수행업무_결과.xlsx
근거:  국가기록원 「2026년 기록물관리 지침」 제2장·제3장 <참고> 수행업무 리스트(2026-11-20 내려받아 확인, assets/guideline2026_tasklist.csv).
       --kind 공공기관: 「2026년 공공기관 기록물관리 지침」 처리과 수행업무 체크리스트(예시) 28·기록관 21(처리과 1-1 기록물관리책임자 지정 포함)
       --kind 대학: 「2026년 대학 기록물관리 지침」 처리과 27·기록관 18 (국가기록원 표준·지침·매뉴얼 > 지침 탭, 2026-11-21 받음)
       법령이 바뀐 항목은 '비고'에 적었다(예: 처리과 3-5 이관시기 연장 — 2027. 1. 1. 시행 시행령 제32조①).
"""
import argparse
import csv
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = {"일반": "guideline2026_tasklist.csv", "공공기관": "public2026_tasklist.csv", "대학": "university2026_tasklist.csv"}
TITLE = {"일반": "「2026년 기록물관리 지침」", "공공기관": "「2026년 공공기관 기록물관리 지침」", "대학": "「2026년 대학 기록물관리 지침」"}


def load(kind="일반"):
    with open(os.path.join(HERE, "..", "assets", ASSETS[kind]), encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--role", choices=["처리과", "기록관", "전체"], default="전체")
    ap.add_argument("--org", default="")
    ap.add_argument("--kind", choices=list(ASSETS), default="일반",
                    help="지침 종류: 일반(중앙·지자체 등, 처리과 28·기록관 24) / 공공기관(처리과 28·기록관 21) / 대학(처리과 27·기록관 18)")
    ap.add_argument("--results", help="O·X를 채운 점검표(이 스크립트가 만든 것)")
    ap.add_argument("-o", "--out", default="산출/수행업무_점검표.xlsx")
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    if a.results:
        src = openpyxl.load_workbook(a.results, data_only=True)["점검표"]
        rows = [r for r in src.iter_rows(min_row=3, values_only=True) if r and r[2]]
        s = wb.active
        s.title = "요약"
        miss = wb.create_sheet("미수행·미기재")
        miss.append(["대상", "구분", "번호", "수행업무", "수행 여부", "도움 되는 킷 스킬", "비고"])
        stat = {}
        for role, sec, no, txt, ox, skill, note in [list(r) + [None] * (7 - len(r)) for r in rows]:
            v = str(ox or "").strip().upper().replace("○", "O").replace("×", "X")
            st = stat.setdefault(role, [0, 0, 0])
            st[0] += 1
            if v == "O":
                st[1] += 1
            else:
                if v not in ("X", ""):
                    v = f"{v}(확인)"
                st[2] += v == ""
                miss.append([role, sec, no, txt, v or "미기재", skill, note])
        s.append(["대상", "항목 수", "수행(O)", "수행률", "미기재"])
        for role, (n, o, blank) in stat.items():
            s.append([role, n, o, f"{o / n * 100:.0f}%" if n else "", blank])
        s.append([])
        s.append(["근거", "「2026년 기록물관리 지침」 <참고> 수행업무 리스트"])
        os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
        wb.save(a.out)
        todo = miss.max_row - 1
        print(" · ".join(f"{k} {v[1]}/{v[0]}" for k, v in stat.items()) + f" · 미수행·미기재 {todo} → {a.out}")
        return 1 if todo else 0
    items = [r for r in load(a.kind) if a.role == "전체" or r["대상"] == a.role]
    ws = wb.active
    ws.title = "점검표"
    ws.append([f"{a.org} 기록관리 수행업무 점검표 ({TITLE[a.kind]} <참고> 수행업무 리스트)".strip()])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append(["대상", "구분", "번호", "수행업무", "수행 여부(O·X)", "도움 되는 킷 스킬", "비고(법령 변경 등)"])
    for c in ws[2]:
        c.font = Font(bold=True)
    dv = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(dv)
    yellow = PatternFill("solid", fgColor="FFF2CC")
    for r in items:
        ws.append([r["대상"], r["구분"], r["번호"], r["수행업무"], "", r["킷 스킬"], r["비고"]])
        dv.add(ws.cell(ws.max_row, 5))
        if r["비고"]:
            ws.cell(ws.max_row, 7).fill = yellow
    for col, w in zip("ABCDEFG", (8, 26, 6, 80, 12, 30, 50)):
        ws.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.kind} {a.role} {len(items)}개 항목 · 법령 변경 비고 {sum(1 for r in items if r['비고'])} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
