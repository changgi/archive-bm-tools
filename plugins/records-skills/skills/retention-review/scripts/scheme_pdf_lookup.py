#!/usr/bin/env python3
"""기관별 기록물분류기준표(2005~2012 공개본 PDF) 찾기·읽기·대조.

국가기록원 기록관리 자료실 > 기록물분류기준 > 기록물분류기준표 탭의 2,416건(지방자치단체 1,170·교육청 768·국공립대학 259·
특별지방행정기관 219, 시행일 2005~2012)을 2026-11-20 목록에서 옮긴 색인(assets/legacy/scheme_pdf_index.csv)으로,
2000~2012년대 생산 기록물의 당시 단위업무·보존기간을 확인한다.

  1) 판 찾기:   scheme_pdf_lookup.py --org 강릉시 [--year 2009]
                → 그 기관의 판(시행일)과 원문 PDF 주소. --year가 있으면 그해 말 기준으로 시행 중이던 판을 표시.
                (스크립트는 내려받지 않는다. 주소를 브라우저로 열어 PDF를 받는다.)
  2) PDF 읽기:  scheme_pdf_lookup.py --pdf 20080131_강원도강릉시.pdf -o 산출/분류기준표.xlsx
                → 처리과기관코드·처리과/기능(원문 줄)·단위업무코드·단위업무명·보존기간 표. 못 읽은 줄은 따로.
  3) 목록 대조: scheme_pdf_lookup.py --pdf ….pdf --list 목록.xlsx -o 산출/분류기준표_대조.xlsx
                → 목록의 단위업무코드(또는 단위업무명)로 당시 보존기간을 찾고, 지금 보존기간이 더 짧으면 표시.

주의: 2012판 PDF 일부는 글꼴 대응표가 어긋나 받침 ㄴ·ㄹ이 ㄲ·ㄳ로 나온다(젂→전, 읷→일). 그런 PDF는 자동으로 고치고
      '글자 보정' 표시를 남긴다. 줄바꿈으로 보존기간이 다음 줄로 넘어간 행은 '못 읽은 줄'에 남기니 원문을 본다.
"""
import argparse
import csv
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, "..", "assets", "legacy", "scheme_pdf_index.csv")
ROW = re.compile(r"^(\d{7})\s*(.*?)\s*([A-Z]{1,2}[A-Z0-9]\d{5,6}|\d{8})\s*(.+?)\s+(영구|준영구?|\d+년)(?=\s|$)")
RANK = {"영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "2년": 2.5, "1년": 2}
ROW2 = re.compile(r"^(\d{7})\s*(.*?)\s*([A-Z]{1,2}[A-Z0-9]\d{5,6}|\d{8})\s*(.+?)(?:\s+(\d{1,2}))?\s+(원본|병행|전자|복제|사본)(?=\s|$)")
GARBLE = ("보졲", "기갂", "기죾", "젂문")


def norm(s):
    return re.sub(r"[\s·.()（）]", "", s or "")


def fix_garble(t):
    """받침 ㄲ(2)→ㄴ, ㄳ(3)→ㄹ로 되돌린다(어긋난 글꼴 대응표)."""
    out = []
    for ch in t:
        if "가" <= ch <= "힣":
            j = (ord(ch) - 0xAC00) % 28
            if j == 2:
                ch = chr(ord(ch) + 2)
            elif j == 3:
                ch = chr(ord(ch) + 5)
        out.append(ch)
    return "".join(out)


def period(p):
    return "준영구" if p.startswith("준영") else p


def read_pdf(path):
    import pdfplumber
    rows, miss, fixed = [], [], False
    with pdfplumber.open(path) as pdf:
        texts = [pg.extract_text() or "" for pg in pdf.pages]
    if any(g in t for t in texts[:5] for g in GARBLE):
        texts, fixed = [fix_garble(t) for t in texts], True
    for pno, t in enumerate(texts, 1):
        lines = t.split("\n")
        for i, ln in enumerate(lines):
            m = ROW.match(ln)
            if m:
                rows.append([pno, m[1], m[2], m[3], m[4].strip(), period(m[5]), ""])
                continue
            m = ROW2.match(ln)
            if m:  # 보존기간 칸이 줄바꿈으로 위아래 줄에 나뉜 행: '준영'+'구', '20'+'년'
                prev, nxt = (lines[i - 1] if i else ""), (lines[i + 1] if i + 1 < len(lines) else "")
                per = ""
                if m[5] and re.match(r"^\s*(?:\S+\s+)*?년(?=\s|$)", nxt) or (m[5] and re.search(r"(^|\s)년(\s|$)", nxt)):
                    per = m[5] + "년"
                elif "준영" in prev and re.search(r"(^|\s)구(\s|$)", nxt):
                    per = "준영구"
                elif m[5] in ("1", "3", "5", "10", "20", "30"):
                    per = m[5] + "년"  # '년' 글자가 원문 텍스트에서 빠진 경우
                if per:
                    rows.append([pno, m[1], m[2], m[3], m[4].strip(), per, "줄 넘김 복원"])
                    continue
            if re.match(r"^\d{7}", ln):
                miss.append([pno, ln])
    return rows, miss, fixed



def find(org, year):
    with open(INDEX, encoding="utf-8-sig") as f:
        idx = list(csv.DictReader(f))
    key = norm(org)
    hits = [r for r in idx if key in norm(r["기관명"])]
    names = sorted({r["기관명"] for r in hits})
    if not hits:
        print(f"'{org}'에 맞는 기관이 색인에 없습니다(지방자치단체·교육청·국공립대학·특별지방행정기관만 있음).", file=sys.stderr)
        return 2
    for name in names:
        vers = sorted((r for r in hits if r["기관명"] == name), key=lambda r: r["시행일자"])
        cur = None
        if year:
            past = [r for r in vers if r["시행일자"] <= f"{year}-12-31"]
            cur = past[-1] if past else None
        print(f"■ {name} ({vers[0]['기관구분']}) — 판 {len(vers)}개")
        for r in vers:
            mark = "  ← %d년 말 시행 중" % year if cur is r else ""
            print(f"  {r['시행일자']}  {r['원문 PDF']}{mark}")
        if year and not cur:
            print(f"  ※ {year}년보다 앞선 판이 없음 — 가장 이른 판({vers[0]['시행일자']})을 참고하고 당시 기관 기준표를 확인")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="기관별 기록물분류기준표(2005~2012 PDF) 찾기·읽기·대조")
    ap.add_argument("--org", help="기관명 일부(예: 강릉시, 완도교육청)")
    ap.add_argument("--year", type=int, help="생산연도: 그해 말 시행 중이던 판 표시")
    ap.add_argument("--pdf", help="내려받은 기록물분류기준표 PDF")
    ap.add_argument("--list", help="대조할 목록(xlsx/csv): 단위업무코드 또는 단위업무명, 보존기간 열")
    ap.add_argument("-o", "--out", default="산출/기록물분류기준표.xlsx")
    a = ap.parse_args()
    if a.org and not a.pdf:
        return find(a.org, a.year)
    if not a.pdf:
        ap.print_help()
        return 2
    try:
        rows, miss, fixed = read_pdf(a.pdf)
    except ModuleNotFoundError:
        print("pdfplumber 꾸러미가 필요합니다. 담당자 허락 후 `python -m pip install pdfplumber`", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "분류기준표"
    ws.append(["쪽", "처리과기관코드", "처리과·대/중/소기능(원문 줄)", "단위업무코드", "단위업무명", "보존기간", "비고"])
    for r in rows:
        ws.append(r)
    ms = wb.create_sheet("못 읽은 줄")
    ms.append(["쪽", "원문 줄(줄바꿈으로 칸이 나뉨 — 원문 확인)"])
    for r in miss:
        ms.append(r)
    flagged = 0
    if a.list:
        if a.list.lower().endswith((".xlsx", ".xlsm")):
            lr = [["" if v is None else str(v).strip() for v in r] for r in openpyxl.load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
        else:
            lr = list(csv.reader(open(a.list, encoding="utf-8-sig")))
        head = [norm(h) for h in lr[0]]
        col = lambda *ns: next((head.index(n) for n in ns if n in head), None)
        ci, ni, pi = col("단위업무코드", "단위과제코드", "단위업무번호"), col("단위업무명", "단위과제명", "단위업무"), col("보존기간", "현보존기간")
        if ci is None and ni is None:
            print("입력 오류: 목록에 단위업무코드 또는 단위업무명 열이 필요합니다.", file=sys.stderr)
            return 2
        by_code, by_name = {}, {}
        for r in rows:
            by_code.setdefault(r[3], r)
            by_name.setdefault(norm(r[4]), r)
        cs = wb.create_sheet("대조", 0)
        cs.append(["원본 행", "단위업무코드", "단위업무명", "지금 보존기간", "당시 단위업무", "당시 보존기간", "판정"])
        for n, r in enumerate(lr[1:], 2):
            g = lambda i, r=r: r[i].strip() if i is not None and i < len(r) else ""
            if not (g(ci) or g(ni)):
                continue
            hit = by_code.get(g(ci)) or by_name.get(norm(g(ni)))
            cur = re.sub(r"\s", "", g(pi))
            if not hit:
                cs.append([n, g(ci), g(ni), cur, "", "", "이 PDF에 없음 — 다른 판·처리과 확인"])
                flagged += 1
                continue
            judge = "같음" if cur == hit[5] else "다름"
            if cur in RANK and hit[5] in RANK and RANK[cur] < RANK[hit[5]]:
                judge = f"지금({cur})이 당시({hit[5]})보다 짧음 — 재평가 때 확인"
                flagged += 1
            cs.append([n, g(ci), g(ni), cur, f"{hit[3]} {hit[4]}", hit[5], judge])
    s = wb.create_sheet("요약")
    for row in (["원문 PDF", os.path.basename(a.pdf)], ["읽은 행", len(rows)], ["못 읽은 줄", len(miss)],
                ["글자 보정", "받침 ㄲ→ㄴ·ㄳ→ㄹ 보정함(글꼴 대응표 어긋남)" if fixed else "없음"],
                ["근거", "국가기록원 기록관리 자료실 > 기록물분류기준 > 기록물분류기준표(기관별 공개본)"],
                ["주의", "처리과명·기능명은 원문 줄을 그대로 둔 것이다(줄바꿈된 이름은 잘릴 수 있음). 판정 전 PDF 원문을 본다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(rows)}행 읽음 · 못 읽은 줄 {len(miss)}" + (" · 글자 보정" if fixed else "") + (f" · 확인 필요 {flagged}" if a.list else "") + f" → {a.out}")
    return 1 if flagged else 0


if __name__ == "__main__":
    sys.exit(main())
