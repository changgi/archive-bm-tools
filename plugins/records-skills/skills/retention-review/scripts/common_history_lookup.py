#!/usr/bin/env python3
"""옛 공통업무 단위과제 → 지금 공통업무 준칙 이어 찾기(개정 이력 대조).

  common_history_lookup.py 목록.xlsx [--edition 2008|2022] -o 산출/공통업무_이력대조.xlsx

목록(기준표나 철 목록)의 단위과제명을 옛 공통업무 보존기간표에서 찾고, 그 과제가 지금 어느 공통업무 단위과제로
이어졌는지(통합·명칭 변경·폐지·보존기간 조정)와 지금 보존기간을 보여 준다. 목록 보존기간이 옛 기준이나 지금 기준보다
짧으면 표시한다.

근거(국가기록원 기록관리 자료실 > 기록물분류기준 > 보존기간 책정기준, 2026-11-20 내려받음)
  - 「처리과공통업무 보존기간표(2008년)」 20과제·「기관공통업무 보존기간표(2008년)」 386과제(엑셀, 개정구분·개정결과 열 포함)
    → 개정결과의 분류번호(ATG·ANG)를 지금 「중앙행정기관 공통업무 보존기간 준칙(2022)」 과제로 연결
  - 「지방자치단체 처리과공통업무 보존기간표(2022년)」 14과제·「기관공통업무 보존기간표(2022년)」 190과제(HWPX)
    → 이름으로 지금 「지방자치단체 공통업무 보존기간 준칙」(2025) 과제와 연결
  → assets/common/common_task_history.csv(610행), 지금 기준은 assets/common/common_task_criteria.csv
"""
import argparse
import csv
import difflib
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
HIST = os.path.join(HERE, "..", "assets", "common", "common_task_history.csv")
CUR = os.path.join(HERE, "..", "assets", "common", "common_task_criteria.csv")
RANK = {"영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "2년": 2.5, "1년": 2}


def key(s):
    s = re.sub(r"^\((?:기관|처리과)공통\)\s*", "", s or "")
    return re.sub(r"[\s·ㆍ,()（）\-/]|및|관리$|업무$", "", s)


def load(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main() -> int:
    ap = argparse.ArgumentParser(description="옛 공통업무 단위과제의 개정 이력과 지금 보존기간")
    ap.add_argument("list")
    ap.add_argument("--edition", choices=["2008", "2022"], help="옛 판 한정(기본: 둘 다)")
    ap.add_argument("--threshold", type=float, default=0.85)
    ap.add_argument("-o", "--out", default="산출/공통업무_이력대조.xlsx")
    a = ap.parse_args()
    hist = [h for h in load(HIST) if not a.edition or h["판"].startswith(a.edition)]
    cur = load(CUR)
    by_code = {c["분류번호"]: c for c in cur}
    local = [c for c in cur if c["기관유형"] == "지방자치단체"]
    if a.list.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.list, encoding="utf-8-sig")))
    hi = max(range(min(10, len(rows))), key=lambda i: sum(bool(c) for c in rows[i]))
    head = [re.sub(r"\s", "", h) for h in rows[hi]]
    find = lambda *ns: next((head.index(n) for n in ns if n in head), None)
    ni, pi = find("단위과제명", "단위과제", "단위업무명"), find("보존기간", "현보존기간")
    if ni is None:
        print("입력 오류: 단위과제명 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대조"
    ws.append(["원본 행", "단위과제명", "목록 보존기간", "옛 판", "옛 단위과제", "옛 보존기간", "개정구분", "개정결과",
               "지금 단위과제", "지금 보존기간", "유사도", "판정"])
    flagged = 0
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        name = r[ni].strip() if ni < len(r) else ""
        if not name:
            continue
        per = re.sub(r"\s", "", r[pi]) if pi is not None and pi < len(r) else ""
        k = key(name)
        best, score = None, 0.0
        for h in hist:
            sc = 1.0 if key(h["단위과제명"]) == k else difflib.SequenceMatcher(None, k, key(h["단위과제명"])).ratio()
            if sc > score:
                best, score = h, sc
        if not best or score < a.threshold:
            continue
        nows = [by_code[c] for c in best["후속 분류번호"].split(",") if c and c in by_code]
        if not nows and best["판"].startswith("2022"):
            kk = key(best["단위과제명"])
            nows = [c for c in local if key(c["단위과제명"]) == kk][:1]
        now_name = " / ".join(f"{c['단위과제명']}({c['분류번호']})" for c in nows)
        now_per = max((c["보존기간"] for c in nows if c["보존기간"] in RANK), key=RANK.get, default="")
        notes = []
        if "폐지" == best["개정구분"].strip():
            notes.append("옛 판에서 폐지된 과제 — " + (best["개정내용"][:60] or "지금 과제 없음"))
        if per in RANK:
            if best["보존기간"] in RANK and RANK[per] < RANK[best["보존기간"]]:
                notes.append(f"목록({per})이 옛 기준({best['보존기간']})보다 짧음")
            if now_per and RANK[per] < RANK[now_per]:
                notes.append(f"목록({per})이 지금 기준({now_per})보다 짧음")
        if notes and any("짧음" in x for x in notes):
            flagged += 1
        ws.append([n, name, per, best["판"], best["단위과제명"], best["보존기간"], best["개정구분"], best["개정결과"],
                   now_name, now_per, f"{score:.2f}", " · ".join(notes) or "같은 흐름"])
    s = wb.create_sheet("근거")
    for row in (["2008 판", "처리과공통업무·기관공통업무 보존기간표(2008년, 시행 2008. 1. 1.) — 개정결과 분류번호로 지금 중앙 공통업무 준칙(2022)에 연결"],
                ["2022 판", "지방자치단체 처리과·기관공통업무 보존기간표(2022년, 시행 2022. 12. 8.) — 이름으로 지금 지방자치단체 공통업무 준칙(2025)에 연결"],
                ["주의", "이름 유사도로 후보를 고를 뿐이다. 통합된 과제는 지금 과제의 업무설명을 보고 담당자가 판단한다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"맞춘 과제 {ws.max_row - 1} · 확인 필요 {flagged} → {a.out}")
    return 1 if flagged else 0


if __name__ == "__main__":
    sys.exit(main())
