#!/usr/bin/env python3
"""BagIt 1.0 이관 패키지 만들기 (RFC 8493). 원본은 복사만 하고 옮기거나 지우지 않는다.

사용:  python make_bag.py 원본폴더/ 산출/bag_이관2026/ [--info 기관=○○시청 --info 담당=홍길동]
산출:  bagit.txt · bag-info.txt · manifest-sha256.txt · tagmanifest-sha256.txt · data/
검증:  python make_bag.py --verify 산출/bag_이관2026/
조건:  이관 대조(reconcile_transfer.py) 불일치가 0건일 때만 실행한다(SKILL.md 절차).
"""
import argparse
import datetime as dt
import hashlib
import os
import shutil
import sys

KST = dt.timezone(dt.timedelta(hours=9))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def rel_files(base):
    for root, _, fs in os.walk(base):
        for fn in sorted(fs):
            yield os.path.relpath(os.path.join(root, fn), base).replace(os.sep, "/")


def write_tagmanifest(bag):
    tags = [t for t in ("bagit.txt", "bag-info.txt", "manifest-sha256.txt") if os.path.exists(os.path.join(bag, t))]
    with open(os.path.join(bag, "tagmanifest-sha256.txt"), "w", encoding="utf-8") as f:
        for t in tags:
            f.write(f"{sha256(os.path.join(bag, t))}  {t}\n")


def make(src, bag, info):
    if os.path.exists(bag) and os.listdir(bag):
        sys.exit(f"이미 내용이 있는 폴더입니다: {bag} (덮어쓰지 않음)")
    data = os.path.join(bag, "data")
    shutil.copytree(src, data)  # copy2로 수정 시각 보존
    lines, total, count = [], 0, 0
    for rel in rel_files(data):
        p = os.path.join(data, rel)
        lines.append(f"{sha256(p)}  data/{rel}")
        total += os.path.getsize(p)
        count += 1
    with open(os.path.join(bag, "manifest-sha256.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    with open(os.path.join(bag, "bagit.txt"), "w", encoding="utf-8") as f:
        f.write("BagIt-Version: 1.0\nTag-File-Character-Encoding: UTF-8\n")
    meta = {"Bagging-Date": dt.datetime.now(KST).strftime("%Y-%m-%d"),
            "Payload-Oxum": f"{total}.{count}", "Source-Organization": "", **info}
    with open(os.path.join(bag, "bag-info.txt"), "w", encoding="utf-8") as f:
        for k, v in meta.items():
            if v:
                f.write(f"{k}: {v}\n")
    write_tagmanifest(bag)
    print(f"패키지 완료: 파일 {count}개 · {total:,}바이트 → {bag}")


def verify(bag) -> int:
    bad = 0
    for mf in ("manifest-sha256.txt", "tagmanifest-sha256.txt"):
        with open(os.path.join(bag, mf), encoding="utf-8") as f:
            for line in f:
                if not line.strip():
                    continue
                digest, rel = line.rstrip("\n").split(maxsplit=1)
                p = os.path.join(bag, rel)
                if not os.path.exists(p) or sha256(p) != digest:
                    print(f"불일치: {rel}")
                    bad += 1
    listed = set()
    with open(os.path.join(bag, "manifest-sha256.txt"), encoding="utf-8") as f:
        listed = {l.split(maxsplit=1)[1].strip() for l in f if l.strip()}
    for rel in rel_files(os.path.join(bag, "data")):
        if f"data/{rel}" not in listed:
            print(f"manifest에 없는 파일: data/{rel}")
            bad += 1
    print("검증 통과" if bad == 0 else f"검증 실패 {bad}건")
    return 0 if bad == 0 else 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?")
    ap.add_argument("bag")
    ap.add_argument("--info", action="append", default=[], help="키=값 (bag-info.txt에 기록)")
    ap.add_argument("--verify", action="store_true")
    a = ap.parse_args()
    if a.verify:
        return verify(a.bag)
    info = dict(x.split("=", 1) for x in a.info)
    make(a.src, a.bag, info)
    return 0


if __name__ == "__main__":
    sys.exit(main())
