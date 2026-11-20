#!/usr/bin/env python3
"""전자매체 간이 검사(NAK 12:2022 부속서 D): 매체의 데이터 전체를 다른 저장공간으로 복사한 뒤 복사 전·후의
폴더 수·파일 수·크기를 비교하고, 파일마다 SHA-256을 대조한다. 복사 속도가 지나치게 느리면 표시한다.

사용:  python media_verify.py 매체경로(예: E:/) 산출/매체검사/DVD-0012 [--min-mbps 2] [--report 산출/매체검사_DVD-0012.csv]
원칙:  매체는 읽기만 한다(쓰기 없음). 복사본은 대상 폴더에만 만든다. 대상 폴더가 비어 있지 않으면 멈춘다.
       결과가 같으면 '검사 시점에' 이상이 없다는 뜻이다. 느리거나 다르면 복구 전문가 자문을 권한다.
"""
import argparse
import csv
import hashlib
import os
import shutil
import sys
import time


def tree(root):
    dirs = files = size = 0
    for d, ds, fs in os.walk(root):
        dirs += len(ds)
        for f in fs:
            files += 1
            size += os.path.getsize(os.path.join(d, f))
    return dirs, files, size


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--min-mbps", type=float, default=2.0, help="이보다 느리면 물리 상태 의심(이 킷의 기본 2MB/s)")
    ap.add_argument("--report")
    a = ap.parse_args()
    if os.path.exists(a.dst) and os.listdir(a.dst):
        print(f"대상 폴더가 비어 있지 않습니다: {a.dst}", file=sys.stderr)
        return 2
    A = tree(a.src)
    t0 = time.time()
    shutil.copytree(a.src, a.dst, dirs_exist_ok=True)
    sec = max(time.time() - t0, 1e-6)
    B = tree(a.dst)
    rows, diff = [], 0
    for d, _, fs in os.walk(a.src):
        for f in sorted(fs):
            s = os.path.join(d, f)
            rel = os.path.relpath(s, a.src)
            t = os.path.join(a.dst, rel)
            hs, ht = sha(s), (sha(t) if os.path.exists(t) else "")
            ok = hs == ht
            diff += not ok
            rows.append([rel, os.path.getsize(s), hs, ht, "같음" if ok else "다름"])
    mbps = A[2] / 1048576 / sec
    slow = A[2] > 50 * 1048576 and mbps < a.min_mbps
    same = A == B and diff == 0
    verdict = "이상 없음(검사 시점)" if same and not slow else "전문가 자문 권고"
    rep = a.report or os.path.join(os.path.dirname(a.dst.rstrip("/\\")) or ".", os.path.basename(a.dst.rstrip("/\\")) + "_검사.csv")
    with open(rep, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["항목", "복사 전(A)", "복사 후(B)", "같음"])
        for n, x, y in (("폴더 수", A[0], B[0]), ("파일 수", A[1], B[1]), ("전체 크기(바이트)", A[2], B[2])):
            w.writerow([n, x, y, "예" if x == y else "아니오"])
        w.writerow(["복사 속도", f"{mbps:.1f}MB/s", "", "느림 — 배드섹터 등 물리 상태 의심" if slow else ""])
        w.writerow(["판정", verdict])
        w.writerow([])
        w.writerow(["파일", "크기", "SHA-256(매체)", "SHA-256(복사본)", "대조"])
        w.writerows(rows)
    print(f"폴더 {A[0]}/{B[0]} · 파일 {A[1]}/{B[1]} · 크기 {A[2]}/{B[2]} · 해시 다름 {diff} · {mbps:.1f}MB/s → {verdict} ({rep})")
    return 0 if verdict.startswith("이상 없음") else 1


if __name__ == "__main__":
    sys.exit(main())
