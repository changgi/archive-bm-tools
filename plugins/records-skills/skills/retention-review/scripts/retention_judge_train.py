#!/usr/bin/env python3
"""보존 의무 판정기 모형 다시 만들기(개발용 — 담당자는 쓸 일 없음). scikit-learn이 필요하다.

입력: 국가법령정보센터에서 받아 둔 법령·행정규칙·자치법규 본문 텍스트 폴더(statutory_catalog_build.py가 받는 것과 같음).
방법: ① 기간 낱말(N년·N개월·N일·영구)이 든 문장을 모은다 ② 엄격한 규칙으로 '확실히 의무'·'확실히 아님'만 약하게 이름 붙인다
      (애매한 문장은 학습에서 뺌) ③ 글자 2~4자 조각으로 로지스틱 회귀를 학습 ④ |가중치| ≥ 0.1만 남겨 JSON으로 저장.
검증: retention_judge.py --gold (사람이 판정한 143문장 — '2차 검증'은 학습·규칙 설계 때 보지 않은 문장).
사용:  python retention_judge_train.py 본문폴더1 [본문폴더2 …] [-o ../assets/law/retention_judge_model.json]
"""
import argparse
import glob
import json
import math
import os
import random
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PER = re.compile(r"(\d+\s*(?:년|개월|일|시간)(?:\s*(?:간|동안|이상|이내|까지))?|영구|준영구)")
KV = re.compile(r"보존|보관|보전|유지|비치|파기|폐기|관리")
OBJ = re.compile(r"기록|서류|장부|자료|대장|문서|파일|로그|영상|녹음|녹화|정보|증빙|명부|일지|원부|사본|내역|회의록|결과|부철|원본|데이터|서식|일계표|보고서|명단|증명")
NEG = re.compile(r"(?<!경력증명)(?:고용|근로관계|신분|(?<!증명)경력(?!증명)|보유하는|보유할|보유한|보유자|휴가|징역|벌금|임기|정지|자격|이내에\s*(?:신고|신청|제출|통보|발급|등기|처분)"
                 r"|주기|마다|판매|시료|우편물|폐기물|동물|진열|지정기간|보호기간|단위로|처분하여야)")
META = re.compile(r"보존기간\s*(?:이|은|을)?\s*\d+\s*년\s*(?:이하|이상|미만|초과)인|보존기간\s*\d+\s*년\s*이상\s*기록물|보존가치|이관하여야")
DEAD = re.compile(r"이내에[^.]{0,40}(?:제출|통보|보고|지급|교육|알려|통지|발급|등록|신고|신청)")
LAPSE = re.compile(r"(?:\d+\s*(?:년|개월|일))이\s*(?:경과|지난|지나)[^.]{0,20}(?:파기|폐기)")


def weak(s):
    near = [s[max(0, m.start() - 60):m.end() + 30] for m in PER.finditer(s)]
    if not near or not KV.search(s):
        return 0
    if META.search(s) or DEAD.search(s) and not re.search(r"보존|보관|보전|비치", s):
        return 0
    hit = any(KV.search(w) and OBJ.search(w) for w in near) or bool(LAPSE.search(s) and OBJ.search(s))
    if hit and not NEG.search(s):
        return 1
    if NEG.search(s) and not any(OBJ.search(w) for w in near):
        return 0
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("dirs", nargs="+")
    ap.add_argument("-o", "--out", default=os.path.join(HERE, "..", "assets", "law", "retention_judge_model.json"))
    a = ap.parse_args()
    try:
        from sklearn.feature_extraction import DictVectorizer
        from sklearn.linear_model import LogisticRegression
    except ImportError:
        print("scikit-learn이 필요합니다(개발용). 담당자 허락을 받고 설치하세요.", file=sys.stderr)
        return 2
    sys.path.insert(0, HERE)
    from retention_judge import feats
    random.seed(7)
    pos, neg = [], []
    for d in a.dirs:
        for fn in glob.glob(os.path.join(d, "*")):
            t = re.sub(r"\s+", " ", open(fn, encoding="utf-8", errors="ignore").read())
            for s in re.split(r"(?<=다\.)\s+|(?=[①-⑳])", t):
                s = s.strip()
                if 15 < len(s) < 400 and PER.search(s) and "href" not in s:
                    lab = weak(s)
                    (pos if lab == 1 else neg if lab == 0 else []).append(s)
    neg = random.sample(neg, min(len(neg), len(pos) * 4))
    X = [{f: 1 / math.sqrt(max(1, len(fs))) for f in fs} for fs in (feats(s) for s in pos + neg)]
    dv = DictVectorizer()
    m = LogisticRegression(C=2, max_iter=3000).fit(dv.fit_transform(X), [1] * len(pos) + [0] * len(neg))
    W = {f: round(float(w), 3) for f, w in zip(dv.feature_names_, m.coef_[0]) if abs(w) >= 0.1}
    json.dump({"b": round(float(m.intercept_[0]), 4), "w": W}, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(f"의무 {len(pos):,} · 아님 {len(neg):,}문장으로 학습 → 특징 {len(W):,}개 → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
