#!/usr/bin/env python3
"""개별 법령 법정 보존기간 점검: 기관 기록관리기준표의 단위과제를 개별 법령이 최소 보존기간을 정한 기록 종류와
낱말로 맞춰, 기준표 보존기간이 법정 기간보다 짧으면 '법정 최소 미달'로 표시하고 맞출 보존기간을 제안한다.

근거(law.go.kr 2026-11-21 원문 확인) → assets/law/statutory_minimum.csv (26종)
  - 「근로기준법」 제42조, 시행령 제22조: 근로자 명부·근로계약서·임금대장·고용·해고·퇴직·승급·휴가·서면 합의 3년(기산일 따로)
  - 「산업안전보건법」 제164조, 시행규칙 제241조: 선임 서류·산업재해 기록·유해성 조사 3년, 위원회 회의록 2년,
    작업환경측정·건강진단 5년(고시 물질 30년), 기관석면조사 3년
  - 「국세기본법」 제85조의3②: 장부·증거서류 법정신고기한부터 5년(역외거래 7년)
  - 「소방시설 설치 및 관리에 관한 법률 시행규칙」 제23조④: 자체점검 실시결과 보고서 점검이 끝난 날부터 2년
  - 「의료법 시행규칙」 제15조①: 진료기록부 등 9종
사용:  python statutory_check.py 기준표.xlsx [--law 근로기준법,산업안전보건법] -o 산출/법정보존기간_점검.xlsx
       python statutory_check.py --list   (법정 기간 표 보기)
기준표 열(머리글 이름으로 찾음): 단위과제명(단위과제), 보존기간, (선택) 단위과제코드, 단위과제 설명(업무설명·설명)
원칙:  법정 기간은 '최소'다. 기준표가 더 길면 그대로 둔다. 공공기록물 보존기간은 1·3·5·10·30년·준영구·영구라
       법정 2년이면 3년, 7년이면 10년처럼 그 이상인 가장 짧은 구분을 제안한다(시행령 제26조①).
       단위과제명으로 맞추면 '법정 최소 미달', 설명 낱말로만 맞추면 '확인 필요'로 낮춰 표시한다.
       낱말로 고른 후보일 뿐이니 '적용 대상'(예: 근로기준법은 공무원이 아닌 근로자)과 기산일을 담당자가 확인한다.
       공공기록물은 법정 기간이 지나도 공공기록물법 평가·폐기 절차로만 폐기한다.
"""
import argparse
import csv
import os
import re
import sys

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "law", "statutory_minimum.csv")
CLASSES = [1, 3, 5, 10, 30]


def load(laws=None):
    with open(ASSET, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if laws:
        rows = [r for r in rows if any(l.strip() and l.strip() in r["법령"] for l in laws)]
    return rows


def years(p):
    """보존기간 글자를 연수로: 영구·준영구는 아주 큰 값, 모르면 None."""
    p = re.sub(r"\s", "", p or "")
    if p in ("영구", "준영구"):
        return 1000 if p == "영구" else 999
    m = re.fullmatch(r"(\d+)년", p)
    return int(m.group(1)) if m else None


def suggest(n):
    return next((f"{c}년" for c in CLASSES if c >= n), "준영구")


def match(rows, name, desc=""):
    """단위과제명에서 먼저, 없으면 설명에서 법정 기록 종류를 찾는다. 같은 과제에 여러 종류면 모두 준다."""
    hit = [r for r in rows if re.search(r["낱말"], name or "")]
    if hit:
        return hit, "단위과제명"
    return [r for r in rows if desc and re.search(r["낱말"], desc)], "설명"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table", nargs="?")
    ap.add_argument("--law", help="쉼표로 나눈 법령 이름 일부(기본: 전부)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("-o", "--out", default="산출/법정보존기간_점검.xlsx")
    a = ap.parse_args()
    rows = load(a.law.split(",") if a.law else None)
    if a.list or not a.table:
        for r in rows:
            print(f"{r['법령']} {r['조항']} | {r['기록 종류']} | {r['최소 보존기간']} | 기산 {r['기산'] or '-'} | {r['적용 대상']}")
        print(f"법정 보존기간 {len(rows)}종 (law.go.kr 2026-11-21 확인)")
        return 0
    if a.table.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        data = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.table, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        data = list(csv.reader(open(a.table, encoding="utf-8-sig")))
    data = data or [[]]
    hi = max(range(min(10, len(data))), key=lambda i: sum(bool(c) for c in data[i]))
    head = [re.sub(r"\s", "", h) for h in data[hi]]
    find = lambda *ns: next((head.index(n) for n in ns if n in head), None)
    ni, pi = find("단위과제명", "단위과제"), find("보존기간")
    ci, di = find("단위과제코드", "코드", "단위과제번호"), find("단위과제설명", "업무설명", "설명")
    if ni is None or pi is None:
        print("입력 오류: 단위과제명·보존기간 열이 필요합니다.", file=sys.stderr)
        return 2
    cell = lambda r, i: r[i] if i is not None and i < len(r) else ""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "단위과제코드", "단위과제명", "기준표 보존기간", "법령", "조항", "기록 종류", "법정 최소", "기산", "적용 대상",
               "판정", "제안 보존기간", "비고"])
    short, maybe, checked = 0, 0, 0
    for n, r in enumerate(data[hi + 1:], hi + 2):
        name, per = cell(r, ni), re.sub(r"\s", "", cell(r, pi))
        if not name:
            continue
        hits, via = match(rows, name, cell(r, di))
        if not hits:
            continue
        checked += 1
        y = years(per)
        for h in hits:
            need = years(h["최소 보존기간"])
            if y is None:
                judge, sug = "보존기간 표기 확인", ""
            elif y < need and via == "설명":
                judge, sug = "확인 필요 — 설명 낱말로만 맞춤(해당 기록이 이 과제에 철해지는지 확인)", suggest(need)
                maybe += 1
            elif y < need:
                judge, sug = "법정 최소 미달 — 기준표 보존기간을 늘림", suggest(need)
                short += 1
            else:
                judge, sug = "충족", ""
            ws.append([n, cell(r, ci), name, per, h["법령"], h["조항"], h["기록 종류"], h["최소 보존기간"], h["기산"], h["적용 대상"],
                       judge, sug, h["비고"]])
    tb = wb.create_sheet("법정 보존기간")
    tb.append(["법령", "조항", "시행", "기록 종류", "최소 보존기간", "기산", "적용 대상", "찾는 낱말", "비고"])
    for r in rows:
        tb.append([r["법령"], r["조항"], r["시행"], r["기록 종류"], r["최소 보존기간"], r["기산"], r["적용 대상"], r["낱말"].replace("\\s*", ""), r["비고"]])
    s = wb.create_sheet("요약")
    for row in (["법정 보존기간 종류", len(rows)], ["법정 기록에 해당한 단위과제", checked], ["법정 최소 미달", short], ["확인 필요(설명 낱말로만 맞춤)", maybe],
                ["확인일", "law.go.kr 2026-11-21 원문(조문은 시행일 열 참조)"],
                ["주의", "법정 기간은 최소다. 적용 대상·기산일을 확인한다. 공공기록물 보존기간 구분(1·3·5·10·30년·준영구·영구)으로 올려 맞추고, "
                         "법정 기간이 지나도 공공기록물법 평가·폐기 절차로만 폐기한다. 법령이 바뀌면 records-qa-gate 근거 변화 감시가 알린다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"법정 보존기간 {len(rows)}종 · 해당 단위과제 {checked} · 법정 최소 미달 {short} · 확인 필요 {maybe} → {a.out}")
    return 1 if short else 0


if __name__ == "__main__":
    sys.exit(main())
