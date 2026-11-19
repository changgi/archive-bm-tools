#!/usr/bin/env python3
"""인수 기록 일괄 점검: 폴더 하나를 받아 포맷 분포·위험 파일·중복·개인정보 후보를 한 장 보고서로 묶는다.

사용:  python triage.py 인수폴더/ -o 산출/점검_2026-11
산출:  보고서.md (요약) · 포맷식별.csv · 개인정보_탐지.csv · 중복.csv
방식:  format-identify와 pii-redact 스킬의 스크립트를 그대로 불러 쓰고, 중복(SHA-256 같음)과 크기 통계만 여기서 계산한다.
       Brunnhilde(Siegfried 기반 기록 보고 도구)와 같은 '먼저 훑어보기' 단계로, 판단은 하지 않는다.
"""
import argparse
import collections
import csv
import datetime as dt
import hashlib
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILLS = os.path.dirname(os.path.dirname(HERE))
KST = dt.timezone(dt.timedelta(hours=9))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run(script, *args):
    p = subprocess.run([sys.executable, os.path.join(SKILLS, script), *args], capture_output=True, text=True)
    return (p.stdout + p.stderr).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--outdir", default="산출/점검")
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    fmt_csv, pii_csv, dup_csv = (os.path.join(a.outdir, n) for n in ("포맷식별.csv", "개인정보_탐지.csv", "중복.csv"))
    fmt_msg = run("format-identify/scripts/identify.py", a.folder, "-o", fmt_csv)
    pii_msg = run("pii-redact/scripts/pii_scan.py", a.folder, "-o", pii_csv)

    files, by_hash, by_ext, total = [], collections.defaultdict(list), collections.Counter(), 0
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            p = os.path.join(root, fn)
            rel = os.path.relpath(p, a.folder)
            size = os.path.getsize(p)
            total += size
            by_ext[os.path.splitext(fn)[1].lower() or "(없음)"] += 1
            by_hash[sha256(p)].append(rel)
            files.append(rel)
    dups = {h: ps for h, ps in by_hash.items() if len(ps) > 1}
    with open(dup_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["sha256", "파일 수", "파일"])
        for h, ps in dups.items():
            w.writerow([h, len(ps), " | ".join(ps)])

    flagged = []
    with open(fmt_csv, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r.get("표시"):
                flagged.append((r["파일"], r["표시"]))
    with open(pii_csv, encoding="utf-8-sig") as f:
        pii_rows = list(csv.DictReader(f))
    pii_files = collections.Counter(r["파일"] for r in pii_rows if r["유형"] != "읽기 실패")
    pii_kinds = collections.Counter(r["유형"] for r in pii_rows if r["유형"] != "읽기 실패")

    now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    L = ["# 인수 기록 점검 보고서", "", f"- 대상: `{a.folder}`", f"- 시각: {now}",
         f"- 파일 {len(files)}개 · {total:,}바이트", "", "## 한눈에", "",
         "| 항목 | 결과 |", "| --- | --- |",
         f"| 포맷 위험 표시 | {len(flagged)}개 |",
         f"| 중복 묶음 | {len(dups)}묶음 ({sum(len(v) for v in dups.values())}개 파일) |",
         f"| 개인정보 후보 | {sum(pii_kinds.values())}건 · {len(pii_files)}개 파일 |", "",
         "## 확장자 분포", "", "| 확장자 | 파일 수 |", "| --- | --- |",
         *[f"| {k} | {v} |" for k, v in by_ext.most_common()], ""]
    if flagged:
        L += ["## 포맷 위험 표시", "", *[f"- `{p}` — {m}" for p, m in flagged[:50]], ""]
    if dups:
        L += ["## 중복", "", *[f"- {' = '.join(f'`{x}`' for x in ps)}" for ps in list(dups.values())[:30]], ""]
    if pii_kinds:
        L += ["## 개인정보 후보 (값은 가림, 세부는 개인정보_탐지.csv)", "",
              *[f"- {k}: {v}건" for k, v in pii_kinds.most_common()],
              *[f"- `{p}` {n}건" for p, n in pii_files.most_common(20)], ""]
    L += ["## 다음 단계", "", "- 포맷 위험 파일은 담당자가 열어 확인 → 필요하면 보존포맷 변환 검토",
          "- 중복은 어느 쪽이 원본인지 확인 후 목록에 반영(삭제는 사람이)",
          "- 개인정보 후보가 있는 파일은 공개구분 판단 때 부분공개 검토",
          "- 이관 목록이 있으면 transfer-reconcile로 대조", "",
          "## 도구 출력", "", f"- 포맷 식별: {fmt_msg}", f"- 개인정보: {pii_msg}"]
    with open(os.path.join(a.outdir, "보고서.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print(f"파일 {len(files)} · 위험 {len(flagged)} · 중복 {len(dups)}묶음 · 개인정보 후보 {sum(pii_kinds.values())}건 → {a.outdir}/보고서.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
