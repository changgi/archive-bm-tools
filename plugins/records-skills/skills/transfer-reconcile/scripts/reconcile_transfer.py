#!/usr/bin/env python3
"""이관 대조: 이관 목록(엑셀·CSV)과 실제 파일 폴더를 맞대어 빠진 것·남는 것·중복을 보고한다.

사용:  python reconcile_transfer.py 이관목록.xlsx 파일폴더/ [--col 파일명] [--key-col 관리번호] -o 산출/이관대조.xlsx
       --col  : 목록에서 파일 이름(또는 상대 경로)이 든 열. 확장자는 무시하고 비교할 수 있다(--ignore-ext).
       --header-row : 머리글이 몇 번째 행인지(기본 자동: '관리번호'나 '파일명'이 든 첫 행)
원칙:  폴더는 읽기만 한다. 결과는 새 엑셀 한 개(요약·불일치 두 장).
종료:  0 일치 · 1 불일치 있음 · 2 입력 오류
"""
import argparse
import collections
import csv
import os
import sys
import unicodedata

HEAD_HINTS = ("관리번호", "파일명", "철관리번호", "기록물건등록번호", "전자파일명")


def norm(s: str, ignore_ext: bool) -> str:
    s = unicodedata.normalize("NFC", s.strip().replace("\\", "/")).split("/")[-1]
    return os.path.splitext(s)[0].lower() if ignore_ext else s.lower()


def read_rows(path, header_row):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    if header_row:
        h = header_row - 1
    else:
        h = next((i for i, r in enumerate(rows[:20]) if any(c in HEAD_HINTS for c in r)), 0)
    head = rows[h]
    return h + 1, head, [(h + 2 + i, dict(zip(head, r))) for i, r in enumerate(rows[h + 1:]) if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listfile")
    ap.add_argument("folder")
    ap.add_argument("--col", default="파일명")
    ap.add_argument("--key-col", default="관리번호")
    ap.add_argument("--header-row", type=int)
    ap.add_argument("--ignore-ext", action="store_true")
    ap.add_argument("-o", "--out", default="산출/이관대조.xlsx")
    a = ap.parse_args()

    try:
        hrow, head, rows = read_rows(a.listfile, a.header_row)
    except Exception as e:
        print(f"입력 오류: {e}", file=sys.stderr)
        return 2
    if a.col not in head:
        print(f"입력 오류: '{a.col}' 열이 없습니다. 머리글({hrow}행): {head}", file=sys.stderr)
        return 2

    files = {}
    for root, _, fs in os.walk(a.folder):
        for fn in fs:
            if fn.startswith("."):
                continue
            files.setdefault(norm(fn, a.ignore_ext), []).append(os.path.relpath(os.path.join(root, fn), a.folder))

    listed = collections.defaultdict(list)
    issues = []
    for rownum, r in rows:
        name = r.get(a.col, "")
        key = r.get(a.key_col, "")
        if not name:
            issues.append(("목록 빈칸", rownum, key, "", "파일명 칸이 비어 있음"))
            continue
        listed[norm(name, a.ignore_ext)].append((rownum, key, name))

    for n, items in listed.items():
        if len(items) > 1:
            issues.append(("목록 중복", ", ".join(str(i[0]) for i in items), items[0][1], items[0][2], f"{len(items)}번 등재"))
        if n not in files:
            for rownum, key, name in items:
                issues.append(("파일 없음", rownum, key, name, "목록에는 있으나 폴더에 없음"))
        elif len(files[n]) > 1:
            issues.append(("폴더 중복", items[0][0], items[0][1], items[0][2], " / ".join(files[n])))
    for n, paths in files.items():
        if n not in listed:
            for p in paths:
                issues.append(("목록 누락", "", "", p, "폴더에는 있으나 목록에 없음"))

    import openpyxl
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "요약"
    counts = collections.Counter(i[0] for i in issues)
    s.append(["항목", "건수"])
    s.append(["목록 행", len(rows)])
    s.append(["폴더 파일", sum(len(v) for v in files.values())])
    for k in ("파일 없음", "목록 누락", "목록 중복", "폴더 중복", "목록 빈칸"):
        s.append([k, counts.get(k, 0)])
    d = wb.create_sheet("불일치")
    d.append(["유형", "목록 행", a.key_col, "파일", "설명"])
    for i in issues:
        d.append(list(i))
    for ws in (s, d):
        ws.column_dimensions["A"].width = 14
    d.column_dimensions["D"].width = 48
    d.column_dimensions["E"].width = 40
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"목록 {len(rows)}행 · 파일 {sum(len(v) for v in files.values())}개 · 불일치 {len(issues)}건 → {a.out}")
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
