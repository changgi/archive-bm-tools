#!/usr/bin/env python3
"""두 추출기 맞대기: 법정 보존기간 목록(statutory_catalog.csv, 조문 문장 단위 추출)과
위임 추적 목록(statutory_items.csv, 항·호·단서·위임 분해)을 같은 법령·같은 조문끼리 비교해 기간이 어긋나는 곳을 찾는다.

어긋남은 어느 한쪽의 추출 오류일 가능성이 크다(예: '원수는 5년간, 정수는 3년간'에서 한 기간만 잡음, 날짜 '1월 1일'을 '1일'로 읽음).
결과는 사람이 원문을 보고 판단할 점검 표다 — 어느 쪽이 맞는지는 정하지 않는다.

사용:
  python statutory_crosscheck.py                     # 요약
  python statutory_crosscheck.py -o 산출/맞대기.csv  # 어긋난 조문 목록(목록 기간·추적 기간·문장)
  python statutory_crosscheck.py --law 산업안전보건법
"""
from __future__ import annotations

import argparse
import collections
import csv
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LAW = os.path.join(HERE, "..", "assets", "law")


def pn(p):
    """기간 → 비교용('10년이상'·'10년(발급일부터)' → '10년', '준영구'·'영구')."""
    p = re.sub(r"\s", "", p or "")
    if p.startswith("준영구"):
        return "준영구"
    if p.startswith("영구"):
        return "영구"
    m = re.match(r"(\d+)(년|개월|일)", p)
    return m.group(1) + m.group(2) if m else None


def load(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def compare(cat_rows, item_rows, law=""):
    jo = lambda s: s.split("(")[0]
    cat, item = collections.defaultdict(set), collections.defaultdict(set)
    sent = collections.defaultdict(list)
    for r in cat_rows:
        if r.get("판정") == "의무 아님 가능성" or (law and law not in r["법령"]):
            continue
        p = pn(r["보존기간"])
        if p:
            cat[(r["법령"], jo(r["조문"]))].add(p)
            sent[(r["법령"], jo(r["조문"]))].append(r["문장"][:200])
    for r in item_rows:
        if r.get("판정") == "의무 아님 가능성" or r["상태"] in ("위임 못 찾음", "방법만 위임(기간 없음)") or (law and law not in r["법령"]):
            continue
        p = pn(r["보존기간"])
        if p:
            item[(r["법령"], jo(r["조문"]))].add(p)
    both = set(cat) & set(item)
    out = []
    for k in sorted(both):
        a, b = cat[k], item[k]
        if a == b:
            continue
        kind = "추적에 빠진 기간" if a > b else ("목록에 빠진 기간" if b > a else "서로 다름")
        out.append({"법령": k[0], "조문": k[1], "어긋남": kind, "목록 기간": "·".join(sorted(a)), "추적 기간": "·".join(sorted(b)),
                    "목록에만": "·".join(sorted(a - b)), "추적에만": "·".join(sorted(b - a)), "문장": " ‖ ".join(dict.fromkeys(sent[k]))[:400]})
    stats = {"목록 조문": len(cat), "추적 조문": len(item), "겹침": len(both), "기간 다름": len(out),
             "목록에만 있는 조문": len(set(cat) - set(item)), "추적에만 있는 조문": len(set(item) - set(cat))}
    return out, stats


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--catalog", default=os.path.join(LAW, "statutory_catalog.csv"))
    ap.add_argument("--items", default=os.path.join(LAW, "statutory_items.csv"))
    ap.add_argument("--law", default="")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    out, st = compare(load(a.catalog), load(a.items), a.law)
    print(" · ".join(f"{k} {v:,}" for k, v in st.items()))
    if st["겹침"]:
        print(f"겹치는 조문 가운데 기간이 같은 비율 {100 * (st['겹침'] - st['기간 다름']) / st['겹침']:.1f}%")
    print(" · ".join(f"{k} {v}" for k, v in collections.Counter(r["어긋남"] for r in out).most_common()))
    if a.out:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["법령", "조문", "어긋남", "목록 기간", "추적 기간", "목록에만", "추적에만", "문장"])
            w.writeheader()
            w.writerows(out)
        print(f"→ {a.out}")
    else:
        for r in out[:15]:
            print(f"- {r['법령']} {r['조문']}: 목록 {r['목록 기간']} / 추적 {r['추적 기간']} ({r['어긋남']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
