#!/usr/bin/env python3
"""대학·의료기관 보존기간 가이드 대조: 기관 기록관리기준표의 단위과제를 국가기록원 기관유형별 가이드의 표준 단위과제와
이름(유사 단위과제 포함)으로 맞춰 보고, 기준표 보존기간이 가이드보다 짧으면 표시한다. 키워드로 가이드를 찾아볼 수도 있다.

근거: 국가기록원 표준·지침·매뉴얼 > 매뉴얼(2026-11-21 내려받아 추출)
      --sector 대학: 「2021 대학 기록물 보존기간 책정기준 가이드」(2021. 11.) 4.2 표준 단위과제별 보존기간 책정기준
                     → assets/sector/university_2021.csv(201과제: 중기능·소기능·단위과제 설명·책정사유·유사 단위과제)
                     국·공립·사립 10개 대학 전문요원 협업반 작성. 비밀기록물 과제는 '개별적용'(원본 비밀의 보존기간).
      --sector 의료: 「의료기관 표준 기록관리기준표 운영 가이드」(2016. 12.) Ⅳ 단위과제별 관리기준
                     → assets/sector/medical_2016.csv(135과제: 대·중·소기능·업무설명·책정사유·공개여부·접근권한)
                     10개 국립대병원 등 의료기관 분과 작성, 의료기관 고유기능(의무). 공통기능은 공통업무 준칙(common_task_criteria)을 본다.
사용:  python sector_guide_lookup.py 기준표.xlsx --sector 대학 -o 산출/대학가이드_대조.xlsx
       python sector_guide_lookup.py --sector 의료 --find 진료비
기준표 열(머리글 이름으로 찾음): 단위과제명(단위과제), 보존기간, (선택) 단위과제코드
원칙:  이름 유사도로 후보를 고를 뿐이다. 가이드는 권고이고 발간 뒤 법령·직제가 바뀌었을 수 있으니(대학 2021·의료 2016)
       설명·책정사유를 읽고 담당자가 판단한다. 개별 법령(의료법 시행규칙 진료기록 보존기간 등)이 있으면 그 법령이 우선이다.
"""
import argparse
import csv
import difflib
import os
import re
import sys

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "sector")
SECTOR = {"대학": ("university_2021.csv", "「2021 대학 기록물 보존기간 책정기준 가이드」(국가기록원, 2021. 11.)"),
          "의료": ("medical_2016.csv", "「의료기관 표준 기록관리기준표 운영 가이드」(국가기록원, 2016. 12.)")}
RANK = {"영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "1년": 2}


def key(s):
    return re.sub(r"[\s·ㆍ,()（）\-/]|및|등|관련|업무|관리$|운영$", "", s or "")


def load(sector):
    with open(os.path.join(ASSET, SECTOR[sector][0]), encoding="utf-8-sig") as f:
        crit = list(csv.DictReader(f))
    for c in crit:
        c["_names"] = [c["단위과제명"]] + [x.strip() for x in c.get("유사 단위과제", "").split(",") if x.strip()]
        c["_desc"] = c.get("단위과제 설명") or c.get("단위업무 설명", "")
        c["_func"] = " > ".join(x for x in (c.get("대기능", ""), c.get("중기능", ""), c.get("소기능", "")) if x)
    return crit


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table", nargs="?")
    ap.add_argument("--sector", choices=sorted(SECTOR), required=True)
    ap.add_argument("--find", help="키워드로 가이드 단위과제 찾기")
    ap.add_argument("--threshold", type=float, default=0.8)
    ap.add_argument("-o", "--out", default="산출/가이드_대조.xlsx")
    a = ap.parse_args()
    crit = load(a.sector)
    if a.find or not a.table:
        kw = re.sub(r"\s", "", a.find or "")
        hits = [c for c in crit if kw and (kw in re.sub(r"\s", "", " ".join(c["_names"]) + c["_desc"]))] if kw else crit[:20]
        for c in hits:
            print(f"{c['번호']:>3} {c['_func']} | {c['단위과제명']} | {c['보존기간']} | {c['_desc'][:60]}")
        print(f"{SECTOR[a.sector][1]} — {len(hits)}/{len(crit)}과제")
        return 0
    if a.table.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.table, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.table, encoding="utf-8-sig")))
    rows = rows or [[]]
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
    cols = ["원본 행", "단위과제코드", "단위과제명", "보존기간", "가이드 번호", "가이드 기능", "가이드 단위과제", "맞춘 이름", "가이드 보존기간", "유사도", "판정", "책정사유"]
    if a.sector == "의료":
        cols += ["가이드 공개여부", "가이드 접근권한"]
    ws.append(cols)
    matched, short = set(), 0
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        name = r[ni] if ni < len(r) else ""
        per = re.sub(r"\s", "", r[pi] if pi < len(r) else "")
        if not name:
            continue
        k = key(name)
        best, score, via = None, 0.0, ""
        for c in crit:
            for nm in c["_names"]:
                sc = 1.0 if key(nm) == k else difflib.SequenceMatcher(None, k, key(nm)).ratio()
                if sc > score:
                    best, score, via = c, sc, nm
        if not best or score < a.threshold:
            continue
        matched.add(best["번호"])
        g = best["보존기간"]
        if g not in RANK:
            judge = f"가이드 '{g}' — 책정사유대로 개별 판단"
        elif per in RANK and RANK[per] < RANK[g]:
            judge = f"기준표({per})가 가이드({g})보다 짧음 — 책정 사유 확인"
            short += 1
        elif per != g:
            judge = "가이드와 다름(더 길거나 표기 다름)"
        else:
            judge = "같음"
        row = [n, r[ci] if ci is not None and ci < len(r) else "", name, per, best["번호"], best["_func"], best["단위과제명"],
               via if via != best["단위과제명"] else "", g, f"{score:.2f}", judge, best["보존기간 책정사유"][:300]]
        if a.sector == "의료":
            row += [best.get("공개여부", ""), best.get("접근권한", "")]
        ws.append(row)
    miss = wb.create_sheet("가이드에만 있는 과제")
    miss.append(["번호", "기능", "단위과제", "보존기간", "설명"])
    for c in crit:
        if c["번호"] not in matched:
            miss.append([c["번호"], c["_func"], c["단위과제명"], c["보존기간"], c["_desc"][:200]])
    s = wb.create_sheet("요약")
    for row in (["가이드", SECTOR[a.sector][1]], ["가이드 단위과제", len(crit)], ["기준표와 맞춘 과제", len(matched)], ["가이드보다 짧은 과제", short],
                ["주의", "가이드는 권고다. 발간 뒤 법령·직제 변경, 개별 법령의 보존기간(우선)을 확인한다. '가이드에만 있는 과제'는 빠졌는지 이름이 다른지 본다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.sector}: 가이드 {len(crit)}과제 · 맞춤 {len(matched)} · 가이드보다 짧음 {short} → {a.out}")
    return 1 if short else 0


if __name__ == "__main__":
    sys.exit(main())
