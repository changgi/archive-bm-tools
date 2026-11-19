#!/usr/bin/env python3
"""검수 게이트: 산출 목록을 기록의 4대 속성 기준으로 점검하고, 하나라도 실패하면 내보내지 않는다.

사용:
  python qa_gate.py 산출/목록.xlsx \
      [--source 원본/이관목록.xlsx --key 관리번호] \
      [--manifest 산출/bag/manifest-sha256.txt --data-dir 산출/bag] \
      [-o 산출/검수결과.md]

점검 항목
  비밀 차단 비밀 구분 열에 값이 있거나 셀에 'Ⅱ급 비밀'·'대외비' 같은 표시가 있으면 즉시 실패(AI 처리 대상 아님)
  진본성   --source를 주면 관리번호별 생산기관·생산일자가 원본 목록과 같은지
  무결성   --manifest를 주면 파일 해시가 manifest와 같은지
  신뢰성   '근거' 열이 있으면 비어 있는 행이 없는지
  이용가능성 필수 열 빈칸, 공개구분 값(공개·부분공개·비공개), 비공개 사유의 호수(1~8호)와 구체 사유,
           '공개' 행에 개인정보 후보가 남아 있는지, 생산일자 형식
종료 코드: 0 통과 · 1 실패 · 2 입력 오류
"""
import argparse
import csv
import datetime as dt
import hashlib
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "pii-redact", "scripts"))
try:
    from pii_scan import scan_text  # 같은 저장소의 pii-redact 스킬을 재사용
except ImportError:
    scan_text = None

REQUIRED_DEFAULT = ["제목", "생산기관", "생산일자", "보존기간", "공개구분"]
ALIASES = {"제목": ["제목", "기록물건제목", "기록물철제목"], "관리번호": ["관리번호", "철관리번호", "기록물건등록번호"]}
DISCLOSURE = {"공개", "부분공개", "비공개"}
REASON_RX = re.compile(r"제?\s?([1-8])\s?호")
SECRET_COL_RX = re.compile(r"(비밀|보안등급|secret|classif)", re.I)
SECRET_EMPTY = {"", "-", "없음", "해당없음", "해당 없음", "일반", "n", "no", "x"}
SECRET_MARK_RX = re.compile(r"(Ⅰ|Ⅱ|Ⅲ|I{1,3}|[123])\s?급\s?비밀|대외비|(?<![가-힣])비밀기록물")
DATE_RX = re.compile(r"^\d{4}([-.]\d{1,2}([-.]\d{1,2})?)?$")
KST = dt.timezone(dt.timedelta(hours=9))


def read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    rows = [r for r in rows if any(r)]
    head, body = rows[0], rows[1:]
    return head, [dict(zip(head, r + [""] * (len(head) - len(r)))) for r in body]


def col(head, name):
    for cand in ALIASES.get(name, [name]):
        if cand in head:
            return cand
    return None


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--required", default=",".join(REQUIRED_DEFAULT))
    ap.add_argument("--source")
    ap.add_argument("--key", default="관리번호")
    ap.add_argument("--manifest")
    ap.add_argument("--data-dir")
    ap.add_argument("-o", "--out", default="산출/검수결과.md")
    a = ap.parse_args()

    try:
        head, rows = read_table(a.table)
    except Exception as e:
        print(f"입력 오류: {e}", file=sys.stderr)
        return 2
    fails = {"비밀 차단": [], "진본성": [], "무결성": [], "신뢰성": [], "이용가능성": []}
    note = []

    # 비밀 차단: 비밀 구분 열 또는 비밀 표시
    secret_cols = [h for h in head if SECRET_COL_RX.search(h) and "비공개" not in h]
    for i, r in enumerate(rows, 2):
        for c in secret_cols:
            if r[c].strip().lower() not in SECRET_EMPTY:
                fails["비밀 차단"].append(f"{i}행 {c} = '{r[c]}' — 비밀기록물은 AI 처리 대상이 아님")
        marks = {m.group(0) for v in r.values() for m in SECRET_MARK_RX.finditer(v)}
        if marks:
            fails["비밀 차단"].append(f"{i}행 비밀 표시 발견: {', '.join(sorted(marks))}")

    # 이용가능성: 필수 열
    for name in [s.strip() for s in a.required.split(",") if s.strip()]:
        c = col(head, name)
        if not c:
            fails["이용가능성"].append(f"필수 열 없음: {name}")
            continue
        for i, r in enumerate(rows, 2):
            if not r[c]:
                fails["이용가능성"].append(f"{i}행 {c} 빈칸")

    disc = col(head, "공개구분")
    reason = col(head, "비공개사유") or col(head, "비공개 사유")
    date = col(head, "생산일자")
    for i, r in enumerate(rows, 2):
        if disc and r[disc]:
            v = r[disc].replace(" ", "")
            if v not in DISCLOSURE:
                fails["이용가능성"].append(f"{i}행 공개구분 값이 표준과 다름: '{r[disc]}'")
            elif v in ("부분공개", "비공개"):
                txt = r.get(reason, "") if reason else ""
                if not REASON_RX.search(txt):
                    fails["이용가능성"].append(f"{i}행 {v}인데 정보공개법 제9조 호수(1~8호)가 없음")
                elif len(REASON_RX.sub("", txt).strip(" ,·-()")) < 4:
                    fails["이용가능성"].append(f"{i}행 비공개 사유가 호수뿐이고 구체 사유가 없음")
            elif v == "공개" and scan_text:
                hits = scan_text(" | ".join(r.values()))
                if hits:
                    kinds = ", ".join(sorted({h[0] for h in hits}))
                    fails["이용가능성"].append(f"{i}행 '공개'인데 개인정보 후보 남음: {kinds}")
        if date and r[date] and not DATE_RX.match(r[date]):
            fails["이용가능성"].append(f"{i}행 생산일자 형식 확인 필요: '{r[date]}'")
    if not scan_text:
        note.append("pii-redact 스킬을 찾지 못해 개인정보 점검을 건너뜀")

    # 신뢰성: 근거 열
    ev = col(head, "근거")
    if ev:
        for i, r in enumerate(rows, 2):
            if not r[ev]:
                fails["신뢰성"].append(f"{i}행 판단 근거 빈칸")
    else:
        note.append("'근거' 열이 없어 신뢰성(판단 근거) 점검을 건너뜀")

    # 진본성: 원본 목록 대조
    if a.source:
        sh, srows = read_table(a.source)
        k_out, k_src = col(head, a.key), col(sh, a.key)
        if not (k_out and k_src):
            fails["진본성"].append(f"대조 키 열 '{a.key}'가 한쪽에 없음")
        else:
            src = {r[k_src]: r for r in srows}
            for i, r in enumerate(rows, 2):
                s = src.get(r[k_out])
                if s is None:
                    fails["진본성"].append(f"{i}행 {r[k_out]}: 원본 목록에 없음")
                    continue
                for name in ("생산기관", "생산일자"):
                    co, cs = col(head, name), col(sh, name)
                    if co and cs and r[co] != s[cs]:
                        fails["진본성"].append(f"{i}행 {r[k_out]} {name} 불일치: '{r[co]}' ≠ 원본 '{s[cs]}'")
    else:
        note.append("--source가 없어 진본성(원본 대조) 점검을 건너뜀")

    # 무결성: manifest 대조
    if a.manifest:
        base = a.data_dir or os.path.dirname(a.manifest)
        with open(a.manifest, encoding="utf-8") as f:
            for line in f:
                if not line.strip():
                    continue
                digest, rel = line.rstrip("\n").split(maxsplit=1)
                p = os.path.join(base, rel)
                if not os.path.exists(p):
                    fails["무결성"].append(f"파일 없음: {rel}")
                elif sha256(p) != digest:
                    fails["무결성"].append(f"해시 불일치: {rel}")
    else:
        note.append("--manifest가 없어 무결성(해시) 점검을 건너뜀")

    total = sum(len(v) for v in fails.values())
    now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    lines = [f"# 검수 결과 — {'통과' if total == 0 else '실패'}", "",
             f"- 대상: `{a.table}` ({len(rows)}행)", f"- 시각: {now}", f"- 실패 {total}건", "",
             "| 속성 | 실패 | 첫 항목 |", "| --- | --- | --- |"]
    for k, v in fails.items():
        lines.append(f"| {k} | {len(v)} | {v[0] if v else '—'} |")
    for k, v in fails.items():
        if v:
            lines += ["", f"## {k}", *[f"- {x}" for x in v]]
    if note:
        lines += ["", "## 건너뛴 점검", *[f"- {x}" for x in note]]
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"{'통과' if total == 0 else '실패'} · 실패 {total}건 · 건너뜀 {len(note)}건 → {a.out}")
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
