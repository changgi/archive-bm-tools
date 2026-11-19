#!/usr/bin/env python3
"""단위과제·보존기간 후보 추천: 기록물 제목을 국가기록원 공통업무 보존기간 준칙(또는 기관 기준표)과 비교해 비슷한 단위과제 3개를 근거와 함께 제안한다.

사용:  python suggest_task.py 목록.xlsx --common 중앙행정기관 [--title-col 제목] [--top 3] -o 산출/단위과제_후보.xlsx
       python suggest_task.py 목록.xlsx --criteria 기관기준표.xlsx -o ...
방식:  글자 2~3그램 TF-IDF 코사인 유사도(외부 패키지·인터넷 없음). 비교 대상 = 단위과제명 + 업무설명 + 기록물유형.
       목록에 현재 보존기간이 있으면 1순위 후보와 다를 때 '확인'으로 표시한다.
근거:  국가기록원 누리집 업무안내·자료 > 기록관리 자료실 > 기록물 분류기준의 공통업무 보존기간 준칙 7종(assets/common_task_criteria.csv, 2026-11-20 받음).
       시행령 제26조②③(책정표준에 따라 책정, 관할 영구기록물관리기관과 협의해 확정).
원칙:  후보일 뿐이다. 단위과제 지정·보존기간 책정은 처리과와 기록관이 한다. 점수가 낮으면(기본 0.15 미만) '후보 약함'으로 둔다.
"""
import argparse
import collections
import csv
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
COMMON = os.path.join(HERE, "..", "..", "retention-review", "assets", "common", "common_task_criteria.csv")


def grams(s):
    s = re.sub(r"[\s\W_]+", "", str(s or ""))
    return [s[i:i + n] for n in (2, 3) for i in range(len(s) - n + 1)]


def read_any(path):
    if path.lower().endswith(".csv"):
        return list(csv.DictReader(open(path, encoding="utf-8-sig")))
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10]) if sum(1 for c in r if c) >= 2), 0)
    head = rows[hi]
    return [dict(zip(head, r)) for r in rows[hi + 1:] if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--common", help="중앙행정기관|지방자치단체|공공기관|교육(행정)기관|학교|대학")
    ap.add_argument("--criteria", help="기관 기준표(단위과제명·보존기간·업무설명 열)")
    ap.add_argument("--title-col", default="제목")
    ap.add_argument("--top", type=int, default=3)
    ap.add_argument("--min-score", type=float, default=0.15)
    ap.add_argument("-o", "--out", default="산출/단위과제_후보.xlsx")
    a = ap.parse_args()
    if a.common:
        crit = [r for r in csv.DictReader(open(COMMON, encoding="utf-8-sig")) if r["기관유형"] == a.common]
    elif a.criteria:
        crit = read_any(a.criteria)
    else:
        print("입력 오류: --common 또는 --criteria가 필요합니다.")
        return 2
    if not crit:
        print("입력 오류: 비교할 단위과제가 없습니다(기관유형 이름 확인).")
        return 2
    g = lambda r, *ks: next((r[k] for k in ks if r.get(k)), "")
    docs = [grams(g(r, "단위과제명") * 2 + " " + g(r, "업무설명", "단위과제 업무설명") + " " + g(r, "기록물유형")) for r in crit]
    df = collections.Counter(t for d in docs for t in set(d))
    N = len(docs)
    idf = {t: math.log((N + 1) / (c + 1)) + 1 for t, c in df.items()}

    def vec(tokens):
        tf = collections.Counter(tokens)
        v = {t: c * idf.get(t, math.log(N + 1) + 1) for t, c in tf.items()}
        n = math.sqrt(sum(x * x for x in v.values())) or 1.0
        return {t: x / n for t, x in v.items()}
    cv = [vec(d) for d in docs]

    items = read_any(a.listing)
    if not items or a.title_col not in items[0]:
        print(f"입력 오류: 목록에 '{a.title_col}' 열이 없습니다.")
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "후보"
    ws.append(["원본 행", "제목", "현재 보존기간", "순위", "단위과제명", "분류번호", "보존기간", "점수", "업무설명(앞부분)", "책정사유(앞부분)", "출처", "판정"])
    weak = diff = 0
    for n, it in enumerate(items, 2):
        title = it.get(a.title_col, "")
        q = vec(grams(title))
        scored = sorted(((sum(q.get(t, 0) * w for t, w in c.items()), i) for i, c in enumerate(cv)), reverse=True)[:a.top]
        cur = re.sub(r"\s", "", it.get("보존기간", "") or "")
        for rank, (sc, i) in enumerate(scored, 1):
            r = crit[i]
            verdict = ""
            if rank == 1:
                if sc < a.min_score:
                    verdict, weak = "후보 약함 — 사람이 지정", weak + 1
                elif cur and cur != g(r, "보존기간"):
                    verdict, diff = f"현재 {cur} ≠ 후보 {g(r, '보존기간')} — 확인", diff + 1
            ws.append([n, title if rank == 1 else "", cur if rank == 1 else "", rank, g(r, "단위과제명"), g(r, "분류번호", "단위과제코드"),
                       g(r, "보존기간"), round(sc, 3), g(r, "업무설명", "단위과제 업무설명")[:80], g(r, "보존기간책정사유")[:60], g(r, "출처"), verdict])
    for i, w in enumerate((8, 34, 10, 6, 26, 12, 8, 7, 50, 40, 30, 26)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"제목 {len(items)} · 비교 단위과제 {N} · 후보 약함 {weak} · 보존기간 다름 {diff} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
