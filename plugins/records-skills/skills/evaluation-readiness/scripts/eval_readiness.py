#!/usr/bin/env python3
"""기록관리 평가 준비 점검: 증빙 문서 목록을 2026년('25년 실적) 평가지표에 맞춰 묶고, 빠진 지표와 증빙 형식 결함을 찾는다.

사용:  python eval_readiness.py 증빙목록.xlsx --type 공공기관|교육(지원)청 --from 2025-01-01 --to 2025-12-31 [--select 3-1,3-6,3-8] -o 산출/평가준비.xlsx
증빙목록 열(머리글 이름으로 찾음): 제목, 등록번호, 등록일자(생산일자), 결재(결재일·결재 여부), 지표(직접 지정하면 우선), 파일명
지표:  assets/indicators_2026.json (국가기록원 누리집 '2026년(25년 실적) 기록관리 평가지표 설명자료' PDF 2종에서 옮김)
점검:  ① 필수지표(1-x·2-x)마다 증빙 후보가 있는지 ② 선택지표(3-x)는 4개 이하·배점 합 20점 이하인지
       ③ 증빙마다 등록번호·등록일자가 있는지, 평가기준일(--from~--to) 안인지, 결재 확인이 되는지(설명자료 제출 유의사항 Ⅲ·Ⅳ·Ⅶ)
원칙:  지표 연결은 제목 낱말로 찾은 '후보'다. 실적 판단과 점수는 평가단이 한다. 제출은 사람이 한다.
"""
import argparse
import collections
import datetime as dt
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ALIAS = {"제목": ["제목", "문서제목", "건제목"], "번호": ["등록번호", "생산등록번호", "문서번호"], "일자": ["등록일자", "생산일자", "시행일자", "등록일"],
         "결재": ["결재", "결재일", "결재여부", "결재일자"], "지표": ["지표", "평가지표"], "파일": ["파일명", "파일"]}


def pdate(v):
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", str(v or ""))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--type", required=True, choices=["공공기관", "교육(지원)청"])
    ap.add_argument("--from", dest="d0", required=True)
    ap.add_argument("--to", dest="d1", required=True)
    ap.add_argument("--select", default="", help="선택지표 id, 쉼표로(예: 3-1,3-6,3-8)")
    ap.add_argument("-o", "--out", default="산출/평가준비.xlsx")
    a = ap.parse_args()
    spec = json.load(open(os.path.join(HERE, "..", "assets", "indicators_2026.json"), encoding="utf-8"))
    inds = {i["id"]: i for i in spec[a.type]}
    d0, d1 = pdate(a.d0), pdate(a.d1)
    import openpyxl
    ws = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    docs = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]

    sel = [s.strip() for s in a.select.split(",") if s.strip()]
    problems = []
    bad_sel = [s for s in sel if s not in inds or not s.startswith("3-")]
    if bad_sel:
        problems.append(["선택지표", ", ".join(bad_sel), "3-x 지표가 아님 또는 이 기관유형에 없음"])
    pts = sum(inds[s]["pts"] for s in sel if s in inds)
    if len(sel) > 4:
        problems.append(["선택지표", ",".join(sel), f"{len(sel)}개 — 최대 4개"])
    if pts > 20:
        problems.append(["선택지표", ",".join(sel), f"배점 합 {pts}점 — 20점 이하로"])
    targets = [i for i in inds.values() if i["id"][0] in "12"] + [inds[s] for s in sel if s in inds]

    links = collections.defaultdict(list)
    fmt = []
    for n, d in docs:
        title = d.get("제목", "")
        hit = [d["지표"]] if d.get("지표") in inds else [i["id"] for i in inds.values() if any(k.replace(" ", "") in title.replace(" ", "") for k in i["kw"])]
        for h in hit:
            links[h].append((n, title))
        flaws = []
        if not d.get("번호"):
            flaws.append("등록번호 없음")
        day = pdate(d.get("일자"))
        if not day:
            flaws.append("등록일자 없음")
        elif d0 and d1 and not (d0 <= day <= d1):
            flaws.append(f"등록일자 {day} — 평가기준일({d0}~{d1}) 밖")
        if "결재" in idx and not d.get("결재"):
            flaws.append("결재 확인 안 됨(문서등록카드 등 추가자료)")
        if not hit:
            flaws.append("연결된 지표 없음")
        fmt.append([n, title, ", ".join(hit), " · ".join(flaws) or "형식 문제 없음"])

    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "지표별 준비"
    s.append(["지표", "지표명", "배점", "구분", "증빙 후보 수", "증빙 후보(앞 3건)", "요구 증빙(설명자료)", "판정"])
    missing = 0
    for i in targets:
        c = links.get(i["id"], [])
        if not c:
            missing += 1
        s.append([i["id"], i["name"], i["pts"], "필수" if i["id"][0] in "12" else "선택", len(c), " / ".join(t for _, t in c[:3]),
                  i["evidence"], "후보 있음 — 내용 확인" if c else "증빙 없음"])
    f = wb.create_sheet("증빙 형식 점검")
    f.append(["원본 행", "제목", "연결 지표", "점검"])
    for r in fmt:
        f.append(r)
    p = wb.create_sheet("선택지표·원칙")
    p.append(["선택지표", ",".join(sel) or "(선택 안 함)", f"{len(sel)}개 · {pts}점"])
    for pr in problems:
        p.append(pr)
    p.append([])
    for rule in spec["_제출원칙"]:
        p.append(["제출 원칙", rule])
    p.append(["출처", spec["_출처"]])
    for sh, wd in ((s, (7, 40, 6, 6, 8, 60, 50, 18)), (f, (8, 50, 16, 60)), (p, (12, 70, 20))):
        for k, w in enumerate(wd):
            sh.column_dimensions[chr(65 + k)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    nflaw = sum(1 for r in fmt if r[3] != "형식 문제 없음")
    print(f"{a.type} · 대상 지표 {len(targets)} · 증빙 없음 {missing} · 형식 결함 문서 {nflaw} · 선택지표 문제 {len(problems)} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
