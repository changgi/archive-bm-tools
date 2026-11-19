#!/usr/bin/env python3
"""문서번호 추출: 기록물 목록의 한 열에서 문서번호를 뽑고, 물결표(~) 범위를 낱개로 펼친다.

  기획행정담당관-351-92~94   →  …-351-92, …-351-93, …-351-94
  총무과-1203~1205          →  총무과-1203, 총무과-1204, 총무과-1205
  총무과-1203~총무과-1205   →  위와 같음
  0089~0091                 →  자릿수(앞자리 0) 보존

사용:  python extract_docno.py 목록.xlsx --col 구기록물문서번호 [--id-col 철관리번호] -o 산출/문서번호.xlsx
산출:  철관리번호 / 문서번호(원문 조각) / 문서번호 전체(펼친 낱개) 3열, 범위가 너무 크면(--max-range) 펼치지 않고 표시
비고:  기존 extract_docno.py를 스킬로 승격하며 재작성한 판. 원본 스크립트가 있으면 결과를 대조할 것.
"""
import argparse
import os
import re
import sys

# 기관·부서명(한글·영문·숫자·괄호) + -숫자 묶음 1~3개
DOCNO = re.compile(r"([가-힣A-Za-z0-9()·]+(?:-\d+){1,3})(?:\s*~\s*((?:[가-힣A-Za-z0-9()·]+-)?(?:\d+-)*\d+))?")


def expand(start: str, end: str | None, max_range: int):
    if not end:
        return [start], ""
    head, last = start.rsplit("-", 1)
    end_last = end.rsplit("-", 1)[-1] if "-" in end else end
    if not (last.isdigit() and end_last.isdigit()):
        return [f"{start}~{end}"], "범위 해석 불가"
    a, b = int(last), int(end_last)
    if b < a:
        return [f"{start}~{end}"], "끝 번호가 시작보다 작음"
    if b - a + 1 > max_range:
        return [f"{start}~{end}"], f"범위 {b - a + 1}건 — 펼치지 않음(--max-range {max_range})"
    width = len(last) if last.startswith("0") else 0
    return [f"{head}-{str(n).zfill(width)}" for n in range(a, b + 1)], ""


def extract(cell: str, max_range: int):
    out = []
    for m in DOCNO.finditer(cell or ""):
        nums, note = expand(m.group(1), m.group(2), max_range)
        out.append((m.group(0), nums, note))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--col", default="구기록물문서번호")
    ap.add_argument("--id-col", default="철관리번호")
    ap.add_argument("--header-row", type=int, help="머리글 행 번호(기본: --col이 든 첫 행)")
    ap.add_argument("--max-range", type=int, default=500)
    ap.add_argument("-o", "--out", default="산출/문서번호.xlsx")
    a = ap.parse_args()

    import openpyxl
    ws = openpyxl.load_workbook(a.table, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    h = (a.header_row - 1) if a.header_row else next((i for i, r in enumerate(rows[:20]) if a.col in r), None)
    if h is None or a.col not in rows[h]:
        print(f"입력 오류: '{a.col}' 열을 찾지 못했습니다.", file=sys.stderr)
        return 2
    head = rows[h]
    ci, ii = head.index(a.col), (head.index(a.id_col) if a.id_col in head else None)

    wb = openpyxl.Workbook()
    out = wb.active
    out.title = "문서번호"
    out.append([a.id_col, "문서번호", "문서번호 전체", "비고"])
    n_in = n_out = 0
    for r in rows[h + 1:]:
        if ci >= len(r) or not r[ci]:
            continue
        rid = r[ii] if ii is not None and ii < len(r) else ""
        for raw, nums, note in extract(r[ci], a.max_range):
            n_in += 1
            for num in nums:
                out.append([rid, raw, num, note])
                n_out += 1
    for c, w in zip("ABCD", (18, 34, 34, 36)):
        out.column_dimensions[c].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"문서번호 조각 {n_in}개 → 낱개 {n_out}개 → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
