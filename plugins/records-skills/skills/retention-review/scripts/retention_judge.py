#!/usr/bin/env python3
"""보존 의무 판정기 — 법령 문장이 '기록을 N년 보존(보관·유지·파기)하라'는 의무를 정한 문장인지 확률로 답한다.
외부 API를 쓰지 않는다(비용 없음·인터넷 없음·즉시). 글자 2~4자 조각에 가중치를 매긴 로지스틱 회귀 모형
(assets/law/retention_judge_model.json, 255KB)을 순수 파이썬으로 계산한다.

판정은 확률을 주고, 반영 여부는 코드(아래 decide)가 정한다:
  p ≥ 0.7 '의무' · 0.3 ≤ p < 0.7 '확인 필요' · p < 0.3 '의무 아님 가능성'
  다만 성격이 '지나면 파기(상한)'인 행과 '명칭: N년' 목록형 조문은 모형이 약한 꼴이라 판정을 낮추지 않는다('확인 필요' 이상).
시험(assets/law/retention_judge_gold.json, 사람이 판정한 문장 143개 중 처음 보는 문장 40개 기준):
  p ≥ 0.3이면 실제 의무 문장을 모두 잡음(재현 1.0), p ≥ 0.7 정밀 0.74. → '의무 아님 가능성'은 지우지 말고 사람 확인으로만 쓴다.

사용:  python retention_judge.py "회의록은 5년간 보존하여야 한다."
       python retention_judge.py --catalog [statutory_catalog.csv] [--write]   # 목록 전체에 '판정' 칸을 매김
       python retention_judge.py --gold                                         # 시험 문장으로 성능 확인
"""
import argparse
import csv
import json
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
MODEL = os.path.join(HERE, "..", "assets", "law", "retention_judge_model.json")
GOLD = os.path.join(HERE, "..", "assets", "law", "retention_judge_gold.json")
CATALOG = os.path.join(HERE, "..", "assets", "law", "statutory_catalog.csv")
HI, LO = 0.7, 0.3
LISTED = re.compile(r"[:：]\s*(?:\d+\s*(?:년|개월)|영구|준영구)")
_M = None


def _model():
    global _M
    if _M is None:
        with open(MODEL, encoding="utf-8") as f:
            _M = json.load(f)
    return _M


def feats(s):
    s = re.sub(r"\d+", "0", s or "")
    out = set()
    for w in s.split():
        w = f" {w} "
        for n in (2, 3, 4):
            for i in range(len(w) - n + 1):
                out.add(w[i:i + n])
    return out


def prob(sentence):
    """이 문장이 기록 보존 의무를 정할 확률(0~1)."""
    m = _model()
    f = feats(sentence)
    if not f:
        return 0.0
    z = m["b"] + sum(m["w"].get(x, 0.0) for x in f) / math.sqrt(len(f))
    return 1 / (1 + math.exp(-z))


def decide(p, 성격="", sentence=""):
    """판정 → 행동(코드가 정함)."""
    if p >= HI:
        return "의무"
    if p >= LO or "상한" in (성격 or "") or LISTED.search(sentence or ""):
        return "확인 필요"
    return "의무 아님 가능성"


def judge(sentence, 성격=""):
    p = prob(sentence)
    return round(p, 3), decide(p, 성격, sentence)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("text", nargs="?")
    ap.add_argument("--catalog", nargs="?", const=CATALOG)
    ap.add_argument("--write", action="store_true", help="--catalog 파일에 '판정' 칸을 써 넣음")
    ap.add_argument("--gold", action="store_true")
    a = ap.parse_args()
    if a.gold:
        with open(GOLD, encoding="utf-8") as f:
            G = json.load(f)
        for grp in sorted({g["묶음"] for g in G}):
            S = [g for g in G if g["묶음"] == grp]
            for th in (LO, HI):
                tp = sum(prob(g["문장"]) >= th and g["정답"] for g in S)
                fp = sum(prob(g["문장"]) >= th and not g["정답"] for g in S)
                fn = sum(prob(g["문장"]) < th and g["정답"] for g in S)
                print(f"{grp} ({len(S)}문장) p≥{th}: 정밀 {tp / max(1, tp + fp):.2f} · 재현 {tp / max(1, tp + fn):.2f}")
        return 0
    if a.catalog:
        with open(a.catalog, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        cols = list(rows[0].keys()) if rows else []
        if "판정" not in cols:
            cols.append("판정")
        cnt = {}
        for r in rows:
            r["판정"] = judge(r.get("문장", ""), r.get("성격", ""))[1]
            cnt[r["판정"]] = cnt.get(r["판정"], 0) + 1
        print(f"{len(rows)}건: " + " · ".join(f"{k} {v:,}" for k, v in sorted(cnt.items())))
        if a.write:
            with open(a.catalog, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)
        return 0
    if not a.text:
        ap.error("문장이나 --catalog, --gold 중 하나를 주세요")
    p, d = judge(a.text)
    print(f"{p:.3f}\t{d}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
