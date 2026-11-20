#!/usr/bin/env python3
"""중앙행정기관 고유업무 보존기간 대조: 기관 기록관리기준표의 고유 단위과제를 국가기록원 「중앙행정기관 기록물 보존기간 준칙」(2021)의
같은 기관 고유업무 단위과제(30년 이상 책정 대상)와 이름으로 맞춰 보고, 기준표의 보존기간이 준칙보다 짧으면 표시한다.

근거: 국가기록원 기록관리 자료실 > 기록물분류기준 > 보존기간 책정기준 「중앙행정기관 기록물 보존기간 준칙(2021년)」(고유, 시행 2021. 1. 1.,
      2026-11-20 내려받아 표 741개에서 추출) → assets/common/central_unique_2021.csv(44개 기관, 단위과제 4,128개: 영구 2,427·30년 939·준영구 752)
      준칙 적용범위: 기관 고유기능 단위과제. 공통기능은 공통업무 보존기간 준칙을 따르고, 준칙에 없는 과제는 유사 과제의 보존기간을 따르며,
      다른 법령에 특별히 정한 경우는 그 법령을 따른다(준칙 1. 일반사항).
      원자력안전위원회·개인정보보호위원회·검찰청·새만금개발청·질병관리청 등은 '추후 제정 예정'으로 빠져 있다.

사용:  python unique_task_lookup.py 기준표.xlsx --org 환경부 -o 산출/고유업무_준칙대조.xlsx
       python unique_task_lookup.py --list-orgs
기준표 열(머리글 이름으로 찾음): 단위과제명(단위과제), 보존기간, (선택) 단위과제코드
원칙:  이름 유사도로 후보를 고를 뿐이다. 기관 업무가 바뀌었을 수 있으니(2021 기준) 업무설명을 보고 담당자가 판단한다.
"""
import argparse
import csv
import difflib
import os
import re
import sys

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "common", "central_unique_2021.csv")
RANK = {"영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "1년": 2}


def key(s):
    return re.sub(r"[\s·ㆍ,()（）\-/]|및|등|관련|업무|관리$|운영$", "", s or "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table", nargs="?")
    ap.add_argument("--org", help="준칙의 기관명(예: 환경부). --list-orgs로 확인")
    ap.add_argument("--list-orgs", action="store_true")
    ap.add_argument("--threshold", type=float, default=0.8, help="이름 유사도 기준(기본 0.8)")
    ap.add_argument("-o", "--out", default="산출/고유업무_준칙대조.xlsx")
    a = ap.parse_args()
    with open(ASSET, encoding="utf-8-sig") as f:
        crit = list(csv.DictReader(f))
    orgs = sorted({c["기관"] for c in crit})
    if a.list_orgs or not a.table:
        print("준칙(2021) 기관 44곳: " + ", ".join(orgs))
        return 0
    org = next((o for o in orgs if a.org and re.sub(r"\s", "", a.org) in re.sub(r"\s", "", o)), None)
    if not org:
        print(f"입력 오류: --org '{a.org}'가 준칙 기관 목록에 없습니다. --list-orgs로 확인하세요.", file=sys.stderr)
        return 2
    mine = [c for c in crit if c["기관"] == org]
    if a.table.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.table, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.table, encoding="utf-8-sig")))
    hi = max(range(min(10, len(rows))), key=lambda i: sum(bool(c) for c in rows[i]))
    head = [re.sub(r"\s", "", h) for h in rows[hi]]
    find = lambda *ns: next((head.index(n) for n in ns if n in head), None)
    ni, pi, ci = find("단위과제명", "단위과제"), find("보존기간"), find("단위과제코드", "코드", "단위과제번호")
    if ni is None or pi is None:
        print("입력 오류: 단위과제명·보존기간 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대조"
    ws.append(["원본 행", "단위과제코드", "단위과제명", "보존기간", "준칙 코드", "준칙 단위과제", "준칙 보존기간", "근거(별표 1)", "유사도", "판정"])
    matched, short = set(), 0
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        if not any(r):
            continue
        name, per = r[ni] if ni < len(r) else "", re.sub(r"\s", "", r[pi] if pi < len(r) else "")
        if not name:
            continue
        k = key(name)
        best, score = None, 0.0
        for c in mine:
            sc = 1.0 if key(c["단위과제"]) == k else difflib.SequenceMatcher(None, k, key(c["단위과제"])).ratio()
            if sc > score:
                best, score = c, sc
        if not best or score < a.threshold:
            continue
        matched.add(best["코드"] + best["단위과제"])
        judge = "같음"
        if per in RANK and best["보존기간"] in RANK and RANK[per] < RANK[best["보존기간"]]:
            judge = f"기준표({per})가 준칙({best['보존기간']})보다 짧음 — 책정 사유 확인"
            short += 1
        elif per != best["보존기간"]:
            judge = "준칙과 다름(더 길거나 표기 다름)"
        ws.append([n, r[ci] if ci is not None and ci < len(r) else "", name, per, best["코드"], best["단위과제"], best["보존기간"], best["근거"], f"{score:.2f}", judge])
    miss = wb.create_sheet("준칙에만 있는 과제")
    miss.append(["준칙 코드", "중기능", "단위과제", "보존기간", "근거", "설명"])
    for c in mine:
        if c["코드"] + c["단위과제"] not in matched:
            miss.append([c["코드"], c["중기능"], c["단위과제"], c["보존기간"], c["근거"], c["설명"][:200]])
    s = wb.create_sheet("요약")
    for row in (["기관", org], ["준칙 단위과제", len(mine)], ["기준표와 맞춘 과제", len(matched)], ["준칙보다 짧은 과제", short],
                ["근거", "국가기록원 「중앙행정기관 기록물 보존기간 준칙」(2021. 1. 1. 시행) — 30년 이상 책정 고유업무"],
                ["주의", "2021년 기준이므로 그 뒤 직제·업무 변경을 반영하지 않았다. '준칙에만 있는 과제'는 기준표에 빠졌는지, 이름이 달라 못 맞췄는지 확인한다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{org}: 준칙 {len(mine)}과제 · 맞춤 {len(matched)} · 준칙보다 짧음 {short} → {a.out}")
    return 1 if short else 0


if __name__ == "__main__":
    sys.exit(main())
