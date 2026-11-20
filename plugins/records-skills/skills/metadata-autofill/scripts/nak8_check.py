#!/usr/bin/env python3
"""기록관리 메타데이터 필수 요소 점검: 등록부·목록 CSV/XLSX가 NAK 8:2022(v2.3) 「기록관리 메타데이터 표준」의
필수 요소(부속서 C 요소표)를 갖췄는지, 세트로 입력해야 할 요소가 짝을 이루는지, 등록 후 바꿀 수 없는 요소가 바뀌지 않았는지 본다.
근거: NAK 8:2022 5.1(상위 필수 12개: 생산자·기록계층·기록식별자·기록물명·전자기록물 여부·유형·크기·분류·일시·보존기간·보존장소·권한),
      부속서 C 요소표, 부속서 D 생산자 유형별 필수 (국가기록원, 2026-11-20 내려받아 확인).

사용:  python nak8_check.py 등록부.csv -o 산출/메타데이터점검.xlsx [--before 이전_등록부.csv]
열 이름: 요소 이름("제목", "생산일시") 또는 번호를 붙인 이름("4.1 제목", "13.1생산일시") 어느 쪽이든 된다.
원칙:  읽기만 한다. 비밀(17.1) 요소 열은 값을 읽지 않는다(있다는 사실만 적음). 값이 맞는지(분류명이 기준표와 같은지 등)는 보지 않는다.
"""
import argparse
import csv
import os
import re
import sys

# (번호, 요소명, 열 별칭들) — 어떤 환경에서든 입력해야 하는 '필수' 요소 중 등록부 한 행에서 볼 수 있는 것
REQUIRED = [
    ("1.1", "생산자유형", ()), ("1.2", "기관명", ("생산기관명",)), ("2", "기록계층", ()), ("3.1", "기본식별자", ("기록식별자",)),
    ("4.1", "제목", ("기록물명",)), ("7", "전자기록물여부", ("전자기록물 여부",)), ("8.1", "기록유형", ()),
    ("11.1", "용량", ("크기",)), ("11.2", "단위", ()), ("12.1", "분류체계유형", ()), ("12.2.2", "분류명", ()),
    ("13.1", "생산일시", ()), ("13.2", "종료일시", ()), ("15.1", "보존기간", ()), ("16", "보존장소", ()), ("17.3.1", "공개구분", ()),
]
# 컨테이너가 쓰였으면(앞 요소에 값이 있으면) 함께 있어야 하는 요소
SETS = [
    (("3.2.2", "시스템식별자"), ("3.2.1", "시스템식별자유형")), (("3.3.2", "보조식별자"), ("3.3.1", "보조식별자유형")),
    (("4.2.2", "기타제목명"), ("4.2.1", "기타제목유형")), (("5.2", "기술내용"), ("5.1", "기술유형")),
    (("6.2", "주제명"), ("6.1", "주제유형")), (("9.2", "포맷버전"), ("9.1", "포맷명")),
    (("14.4.1", "비치사유"), ("14.4.2", "비치기간")), (("14.4.2", "비치기간"), ("14.4.1", "비치사유")),
    (("18.2", "소장위치"), ("18.1", "소장처")), (("23.2", "무결성체크값"), ("23.1", "무결성체크법")),
    (("22.2", "관계대상식별자"), ("22.1", "관계유형")),
]
# 등록 후 바꿀 수 없는 요소(5.1): 1 생산자, 3.1 기본식별자, 4.1 제목, 13.1 생산일시, 14.6 생산경로
IMMUTABLE = [("1.2", "기관명"), ("1.6", "개인명"), ("4.1", "제목"), ("13.1", "생산일시")]
# 부속서 D: 이 생산자 유형은 개인명이 필수
NEED_PERSON = ("기안자", "검토자", "협조자", "결재자", "배부자", "접수자", "접수문서업무담당자", "보고수신자", "기록물철담당자")
SECRET = re.compile(r"^\s*17\.1|비밀")
CLOSED = ("비공개", "부분공개")
DATE = re.compile(r"(\d{4})[-./]?\s*(\d{1,2})?[-./]?\s*(\d{1,2})?")


def norm(h):
    return re.sub(r"[\s()·]", "", re.sub(r"^\s*[\d.]+\s*", "", h or ""))


def read(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        ws = load_workbook(path, read_only=True, data_only=True).active
        rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    return (rows[0], rows[1:]) if rows else ([], [])


def column_map(header):
    """요소 번호 → 열 위치. 번호가 붙은 머리글이 우선, 없으면 이름·별칭으로 찾는다."""
    idx, secret_cols = {}, []
    names = {n: c for c, n, _ in REQUIRED}
    for c, n in [x for s in SETS for x in s] + IMMUTABLE + [("1.6", "개인명"), ("17.3.2", "비공개사유"), ("9.1", "포맷명")]:
        names.setdefault(n, c)
    alias = {norm(a): c for c, _, al in REQUIRED for a in al}
    for i, h in enumerate(header):
        if SECRET.search(h or ""):
            secret_cols.append(i)
            continue
        m = re.match(r"^\s*([\d]+(?:\.\d+)*)", h or "")
        if m:
            idx.setdefault(m.group(1), i)
    for i, h in enumerate(header):
        if i in secret_cols:
            continue
        k = norm(h)
        code = names.get(k) or alias.get(k)
        if code and code not in idx:
            idx[code] = i
    return idx, secret_cols


def val(row, idx, code):
    i = idx.get(code)
    return (row[i] if i is not None and i < len(row) else "").strip()


def ymd(s):
    m = DATE.search(s or "")
    return tuple(int(x) if x else 0 for x in m.groups()) if m else None


def check(header, rows, before=None):
    idx, secret_cols = column_map(header)
    missing_cols = [f"{c} {n}" for c, n, _ in REQUIRED if c not in idx]
    out = []
    old = {}
    if before:
        bh, br = before
        bidx, _ = column_map(bh)
        old = {val(r, bidx, "3.1"): (r, bidx) for r in br if val(r, bidx, "3.1")}
    for k, r in enumerate(rows, 2):
        if not any(c.strip() for c in r):
            continue
        rid = val(r, idx, "3.1") or f"{k}행"
        probs = [f"필수 {c} {n} 비어 있음" for c, n, _ in REQUIRED if c in idx and not val(r, idx, c)]
        for (c1, n1), (c2, n2) in SETS:
            if c1 in idx and val(r, idx, c1) and not (c2 in idx and val(r, idx, c2)):
                probs.append(f"{c1} {n1}이 있으면 {c2} {n2}도 함께(세트)")
        if val(r, idx, "17.3.1").replace(" ", "") in CLOSED and "17.3.2" in idx and not val(r, idx, "17.3.2"):
            probs.append("공개구분이 비공개·부분공개면 17.3.2 비공개사유 필요")
        if "1.1" in idx and val(r, idx, "1.1") in NEED_PERSON and "1.6" in idx and not val(r, idx, "1.6"):
            probs.append(f"생산자유형 '{val(r, idx, '1.1')}'은 1.6 개인명 필수(부속서 D)")
        a, b = ymd(val(r, idx, "13.1")), ymd(val(r, idx, "13.2"))
        if a and b and b < a:
            probs.append("13.2 종료일시가 13.1 생산일시보다 앞섬")
        e = val(r, idx, "7").upper().replace(" ", "")
        if e in ("Y", "전자", "예", "TRUE") and "9.1" in idx and not val(r, idx, "9.1"):
            probs.append("전자기록물이면 9 포맷(9.1 포맷명) 해당시 필수")
        if rid in old:
            orow, oidx = old[rid]
            for c, n in IMMUTABLE:
                if c in idx and c in oidx and val(orow, oidx, c) and val(orow, oidx, c) != val(r, idx, c):
                    probs.append(f"등록 후 변경 불가 요소 {c} {n} 바뀜('{val(orow, oidx, c)}' → '{val(r, idx, c)}')")
        out.append([k, rid, len(probs), " · ".join(probs), "통과" if not probs else "고칠 것 있음"])
    return out, missing_cols, len(secret_cols)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("register")
    ap.add_argument("-o", "--out", default="산출/메타데이터점검.xlsx")
    ap.add_argument("--before", help="이전 등록부(같은 형식) — 등록 후 변경 불가 요소 대조")
    a = ap.parse_args()
    header, rows = read(a.register)
    rows_out, missing_cols, nsecret = check(header, rows, read(a.before) if a.before else None)
    notes = []
    if missing_cols:
        notes.append("필수 요소 열이 없음: " + ", ".join(missing_cols))
    if nsecret:
        notes.append(f"비밀(17.1) 관련 열 {nsecret}개는 값을 읽지 않음 — 비밀기록물은 담당자가 별도 점검")
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    head = ["행", "기본식별자", "문제 수", "문제", "판정"]
    if a.out.lower().endswith(".xlsx"):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "점검"
        ws.append(head)
        for r in rows_out:
            ws.append(r)
        s = wb.create_sheet("요약")
        s.append(["항목", "값"])
        s.append(["점검 행", len(rows_out)])
        s.append(["고칠 것 있음", sum(1 for r in rows_out if r[2])])
        for n in notes:
            s.append(["알림", n])
        s.append(["근거", "NAK 8:2022(v2.3) 5.1·부속서 C·D"])
        wb.save(a.out)
    else:
        with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(head)
            w.writerows(rows_out)
            for n in notes:
                w.writerow(["", "(알림)", "", n, ""])
    bad = sum(1 for r in rows_out if r[2])
    print(f"{len(rows_out)}행 · 고칠 것 있음 {bad}" + (f" · {notes[0]}" if missing_cols else "") + f" → {a.out}")
    return 1 if bad or missing_cols else 0


if __name__ == "__main__":
    sys.exit(main())
