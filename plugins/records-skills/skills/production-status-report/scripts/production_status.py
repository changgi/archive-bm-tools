#!/usr/bin/env python3
"""기록물 생산현황 통보 준비: 등록정보 목록을 집계하고, 통보 전에 고칠 것을 찾는다.

사용:  python production_status.py 등록정보.xlsx --year 2025 [--deadline 2026-05-31] -o 산출/생산현황_2025.xlsx
목록 열(머리글 이름으로 찾음): 생산부서, 생산연도(또는 생산일자), 단위과제, 보존기간, 기록물형태(전자·비전자),
       유형(일반문서·시청각·행정박물 등), 공개구분, 등록번호(또는 철번호), 제목, 쪽수, 실물쪽수, 비밀구분
근거:  공공기록물법 제19조 — 공공기관은 매년 기록물 생산현황을 관할 기록물관리기관에 통보.
       시행령 제33조·제42조 — 기록물·기록물철 등록정보를 시스템으로 제출(제20조①·제23조②의 등록정보).
       시행규칙 제21조 — 생산현황은 기록물 정리가 끝나 등록정보에 반영된 뒤 작성.
       시행규칙(정리) — 등록번호 누락 확인, 비전자기록물은 등록정보상 쪽수와 실물 쪽수 일치 확인.
       2026. 10. 2. 시행 개정 시행령으로 통보 주체·대상이 바뀌었다는 해설이 있다. 날짜(--deadline)와 제출 경로는 기관에서 확인해 넣는다.
원칙:  집계와 점검만 한다. 제출은 사람이 시스템으로 한다. 비밀 표시 행은 수만 세고 제목을 옮기지 않는다.
"""
import argparse
import collections
import os
import re
import sys

ALIAS = {
    "생산부서": ["생산부서", "처리과", "부서", "생산부서(처리과)명", "처리과명", "생산기관"], "생산연도": ["생산연도", "생산년도", "연도", "생산시작년도"], "생산일자": ["생산일자", "등록일자"],
    "단위과제": ["단위과제", "단위과제명", "단위업무명/단위과제명", "단위업무명", "과제카드명"], "보존기간": ["보존기간"], "형태": ["기록물형태", "전자구분", "형태"],
    "유형": ["유형", "기록물유형", "기록물구분"], "공개구분": ["공개구분", "공개여부", "열람조건"], "등록번호": ["등록번호", "생산등록번호", "철번호", "분류번호"],
    "제목": ["제목", "철제목", "건제목", "기록물철제목", "기록물건제목", "기록물철명"], "쪽수": ["쪽수", "면수"], "실물쪽수": ["실물쪽수", "실제쪽수"], "비밀": ["비밀구분", "비밀여부"],
}
PERIODS = ["영구", "준영구", "30년", "10년", "5년", "3년", "1년"]


def norm_period(v):
    s = str(v or "").replace(" ", "")
    if s in PERIODS:
        return s
    m = re.fullmatch(r"(\d+)(년)?", s)
    return f"{m[1]}년" if m else s


def read(path):
    import openpyxl
    if path.lower().endswith(".csv"):
        import csv
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    else:
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(str(c or "").strip() in names for c in r) >= 3), 0)
    head = [str(c or "").strip() for c in rows[hi]]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head:
                idx[k] = head.index(n)
                break
    data = []
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        if any(c not in (None, "") for c in r):
            data.append((n, {k: ("" if r[i] is None or i >= len(r) else r[i]) for k, i in idx.items()}))
    return data, idx


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, required=True, help="통보 대상 생산연도(전년도)")
    ap.add_argument("--deadline", help="기관이 확인한 통보 기한 YYYY-MM-DD(표시용)")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    out = a.out or f"산출/생산현황_{a.year}.xlsx"
    data, idx = read(a.listing)
    need = [k for k in ("생산부서", "보존기간", "등록번호") if k not in idx]
    if "생산연도" not in idx and "생산일자" not in idx:
        need.append("생산연도/생산일자")
    if need:
        print("입력 오류: 필요한 열이 없습니다 — " + ", ".join(need))
        return 2

    issues, secret, target = [], 0, []
    seen = collections.Counter(str(r.get("등록번호", "")).strip() for _, r in data)
    for ln, r in data:
        y = r.get("생산연도") or ""
        if not y and r.get("생산일자"):
            m = re.search(r"(\d{4})", str(r["생산일자"]))
            y = m[1] if m else ""
        try:
            y = int(str(y)[:4])
        except ValueError:
            issues.append((ln, "", "생산연도 없음·형식 오류", "제20조 등록정보"))
            continue
        if y != a.year:
            continue
        no = str(r.get("등록번호", "")).strip()
        if str(r.get("비밀", "")).strip() not in ("", "일반", "N", "아니오", "해당없음"):
            secret += 1
            continue
        target.append(r)
        p = norm_period(r.get("보존기간"))
        if not no:
            issues.append((ln, "", "등록번호 없음", "시행규칙 정리 — 누락 등록번호 표기"))
        elif seen[no] > 1:
            issues.append((ln, no, f"등록번호 중복({seen[no]}건)", "등록정보 확인"))
        if p not in PERIODS:
            issues.append((ln, no, f"보존기간 '{r.get('보존기간')}' 표준 구분 아님", "시행령 제26조"))
        if not str(r.get("단위과제", "")).strip() and "단위과제" in idx:
            issues.append((ln, no, "단위과제 없음", "시행령 제25조 기록관리기준표"))
        if "공개구분" in idx and not str(r.get("공개구분", "")).strip():
            issues.append((ln, no, "공개구분 없음", "법 제35조·정보공개법 제9조"))
        form = str(r.get("형태", ""))
        if "비전자" in form and "실물쪽수" in idx:
            a1, a2 = str(r.get("쪽수", "")).strip(), str(r.get("실물쪽수", "")).strip()
            if a1 and a2 and a1 != a2:
                issues.append((ln, no, f"쪽수 불일치(등록 {a1} · 실물 {a2})", "시행규칙 정리 — 쪽수 일치 확인"))
            elif not a2:
                issues.append((ln, no, "비전자기록물인데 실물 쪽수 확인 기록 없음", "시행규칙 정리"))

    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    by_dept = collections.Counter(str(r.get("생산부서", "")).strip() or "(부서 없음)" for r in target)
    by_period = collections.Counter(norm_period(r.get("보존기간")) for r in target)
    by_form = collections.Counter(str(r.get("형태", "")).strip() or "(형태 없음)" for r in target)
    by_type = collections.Counter(str(r.get("유형", "")).strip() or "(유형 없음)" for r in target)
    ws.append([f"{a.year}년 생산 기록물 생산현황(통보 준비용)"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["대상 건수", len(target)])
    ws.append(["비밀 표시(별도 통보, 제목 미기재)", secret])
    ws.append(["고칠 것", len(issues)])
    ws.append(["통보 기한", a.deadline or "기관 확인 필요(2026. 10. 2. 시행 개정 시행령 확인)"])
    ws.append([])
    for title, c in (("보존기간별", by_period), ("형태별", by_form), ("유형별", by_type)):
        ws.append([title])
        ws[f"A{ws.max_row}"].font = Font(bold=True)
        order = PERIODS if title == "보존기간별" else sorted(c)
        for k in order + [k for k in c if k not in order]:
            if c.get(k):
                ws.append([k, c[k]])
        ws.append([])
    w2 = wb.create_sheet("부서×보존기간")
    cols = [p for p in PERIODS if by_period.get(p)] + [p for p in by_period if p not in PERIODS]
    w2.append(["생산부서"] + cols + ["합계"])
    for d in sorted(by_dept):
        cc = collections.Counter(norm_period(r.get("보존기간")) for r in target if (str(r.get("생산부서", "")).strip() or "(부서 없음)") == d)
        w2.append([d] + [cc.get(p, 0) for p in cols] + [sum(cc.values())])
    w2.append(["합계"] + [by_period[p] for p in cols] + [len(target)])
    w3 = wb.create_sheet("고칠 것")
    w3.append(["원본 행", "등록번호", "내용", "근거"])
    for row in issues:
        w3.append(list(row))
    for sh, widths in ((ws, (34, 40)), (w2, [18] + [8] * (len(cols) + 1)), (w3, (8, 20, 44, 34))):
        for i, wdt in enumerate(widths):
            sh.column_dimensions[chr(65 + i)].width = wdt
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)
    print(f"{a.year}년 생산 {len(target)}건(비밀 {secret}건 별도) · 부서 {len(by_dept)} · 고칠 것 {len(issues)} → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
