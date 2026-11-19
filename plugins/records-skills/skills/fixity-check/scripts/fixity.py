#!/usr/bin/env python3
"""무결성 기준선 만들기와 정기 대조 (SHA-256).

  python fixity.py baseline 보존/ -o 산출/기준선_2026-11.csv
  python fixity.py verify   보존/ 산출/기준선_2026-11.csv -o 산출/대조결과.csv

verify 결과 유형: 변경(해시 다름) · 사라짐 · 새 파일
종료 코드: 0 이상 없음 · 1 이상 있음
AIC(Archive Integrity Checker)가 설치돼 있으면 그쪽 aic-baseline / aic-check를 우선 쓰고,
이 스크립트는 AIC가 없는 PC에서 쓰는 최소판이다(같은 CSV 열 구성: path, size, sha256, mtime).
"""
import argparse
import csv
import datetime as dt
import hashlib
import os
import sys

KST = dt.timezone(dt.timedelta(hours=9))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def walk(base):
    for root, _, fs in os.walk(base):
        for fn in sorted(fs):
            p = os.path.join(root, fn)
            yield os.path.relpath(p, base).replace(os.sep, "/"), p


def baseline(base, out):
    n = 0
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["path", "size", "sha256", "mtime"])
        for rel, p in walk(base):
            st = os.stat(p)
            w.writerow([rel, st.st_size, sha256(p), dt.datetime.fromtimestamp(st.st_mtime, KST).isoformat(timespec="seconds")])
            n += 1
    print(f"기준선 {n}개 파일 → {out}")
    return 0


def verify(base, ref, out):
    with open(ref, encoding="utf-8-sig", newline="") as f:
        known = {r["path"]: r for r in csv.DictReader(f)}
    seen, issues = set(), []
    for rel, p in walk(base):
        seen.add(rel)
        if rel not in known:
            issues.append(["새 파일", rel, "", sha256(p)])
        else:
            d = sha256(p)
            if d != known[rel]["sha256"]:
                issues.append(["변경", rel, known[rel]["sha256"], d])
    for rel in known:
        if rel not in seen:
            issues.append(["사라짐", rel, known[rel]["sha256"], ""])
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["유형", "path", "기준선 sha256", "현재 sha256"])
        w.writerows(issues)
    print(f"대조 {len(seen)}개 · 이상 {len(issues)}건 → {out}")
    return 1 if issues else 0


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("baseline")
    b.add_argument("base")
    b.add_argument("-o", "--out", default="산출/기준선.csv")
    v = sub.add_parser("verify")
    v.add_argument("base")
    v.add_argument("ref")
    v.add_argument("-o", "--out", default="산출/대조결과.csv")
    a = ap.parse_args()
    return baseline(a.base, a.out) if a.cmd == "baseline" else verify(a.base, a.ref, a.out)


if __name__ == "__main__":
    sys.exit(main())
