#!/usr/bin/env python3
"""정수점검·상태검사 일정: 보존 기록물 목록에서 다음 점검 연도와 지난 점검을 찾는다.

사용:  python inspection_schedule.py 보존목록.xlsx --year 2026 -o 산출/점검일정.xlsx
목록 열(머리글 이름으로 찾음): 관리번호, 제목, 유형(종이·행정박물·영화필름·오디오·비디오·사진·필름·전자매체 등),
       상태등급(1·2·3, 종이·행정박물만), 최근정수점검(연도 또는 날짜), 최근상태검사, 생산연도, 서고, 서가
근거:  시행규칙 제31조·별표 14 <개정 2026. 9. 11.>
         정수점검 모두 2년 / 상태검사: 종이·행정박물 1등급 30년·2등급 15년·3등급 10년, 영화필름 2년, 오디오·비디오 3년,
         사진·필름 10년, 전자기록물 등 보존매체 5년
       생산 후 70년 지난 기록물 중 영구기록물관리기관의 장이 정하는 것은 주기의 2배 범위에서 따로 정할 수 있다(--long-factor).
       점검 결과는 별지 제7호서식 '보존기록물의 정수점검·상태검사서'로 작성한다(서식 칸은 기관 서식에 맞춘다).
원칙:  일정 계산만 한다. 점검과 사후조치(복원·탈산 등)는 사람이 한다.
"""
import argparse
import collections
import os
import re
import sys

COUNT_CYCLE = 2
STATE = {("종이", 1): 30, ("종이", 2): 15, ("종이", 3): 10, "영화필름": 2, "오디오비디오": 3, "사진필름": 10, "전자매체": 5}
ALIAS = {"관리번호": ["관리번호", "등록번호", "철번호"], "제목": ["제목", "철제목"], "유형": ["유형", "기록물유형", "매체"],
         "등급": ["상태등급", "상태검사등급", "등급"], "정수": ["최근정수점검", "정수점검", "최근 정수점검"],
         "상태": ["최근상태검사", "상태검사", "최근 상태검사"], "생산연도": ["생산연도", "생산년도"], "서고": ["서고"], "서가": ["서가"],
         "보존기간": ["보존기간"], "재질": ["재질", "용지", "매체형태", "기록물형태"]}
FRAGILE = re.compile(r"감열|감광|청사진|트레이싱|습식복사|갱지|신문용지|저급|재활용지|자기|테이프|비디오|오디오|카세트|VHS|필름|질산|초산|아세테이트|슬라이드")


def kind(t):
    s = re.sub(r"\s", "", str(t))
    if "영화" in s:
        return "영화필름"
    if any(k in s for k in ("오디오", "비디오", "녹음", "녹화", "테이프")):
        return "오디오비디오"
    if any(k in s for k in ("사진", "필름", "슬라이드", "마이크로")):
        return "사진필름"
    if any(k in s for k in ("전자", "매체", "디스크", "LTO", "광디스크", "HDD")):
        return "전자매체"
    if any(k in s for k in ("종이", "문서", "도면", "카드", "행정박물", "박물", "비전자")):
        return "종이"
    return None


def year(v):
    m = re.search(r"(1[89]\d{2}|20\d{2})", str(v or ""))
    return int(m[1]) if m else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--long-factor", type=float, default=1.0, help="생산 70년 경과 지정 기록물의 주기 배수(최대 2)")
    ap.add_argument("--soon", type=int, default=1, help="며칠 아닌 몇 해 안에 돌아오면 '예정'으로 볼지")
    ap.add_argument("-o", "--out", default="산출/점검일정.xlsx")
    a = ap.parse_args()
    factor = min(max(a.long_factor, 1.0), 2.0)
    import openpyxl
    ws = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head:
                idx[k] = head.index(n)
                break
    out = openpyxl.Workbook()
    w1 = out.active
    w1.title = "점검 일정"
    w1.append(["관리번호", "제목", "유형", "상태등급", "서고·서가", "최근 정수점검", "다음 정수점검", "최근 상태검사", "상태검사 주기(년)", "다음 상태검사", "판정"])
    prio = []
    w2 = out.create_sheet("확인 필요")
    w2.append(["원본 행", "관리번호", "내용"])
    cnt = collections.Counter()
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        if not any(r):
            continue
        g = lambda k: r[idx[k]] if k in idx and idx[k] < len(r) else ""
        no, k = g("관리번호"), kind(g("유형"))
        if not k:
            w2.append([n, no, f"별표 14 구분에 맞출 수 없는 유형: '{g('유형')}'"])
            cnt["확인"] += 1
            continue
        grade = None
        if k == "종이":
            m = re.search(r"[123]", g("등급"))
            grade = int(m[0]) if m else None
            if not grade:
                w2.append([n, no, "종이·행정박물인데 상태검사 등급(1~3)이 없음 — 상태검사 주기를 정할 수 없음"])
                cnt["확인"] += 1
        cyc = STATE[(k, grade)] if k == "종이" and grade else (STATE[k] if k != "종이" else None)
        py = year(g("생산연도"))
        f = factor if (py and a.year - py >= 70) else 1.0
        cc = COUNT_CYCLE * f
        lc, ls = year(g("정수")), year(g("상태"))
        nc = int(lc + cc) if lc else None
        ns = int(ls + cyc * f) if (ls and cyc) else None
        verdict = []
        for label, nxt, last in (("정수점검", nc, lc), ("상태검사", ns, ls)):
            if last is None:
                verdict.append(f"{label} 기록 없음")
                cnt[f"{label} 기록 없음"] += 1
            elif nxt is not None and nxt < a.year:
                verdict.append(f"{label} {a.year - nxt}년 지남")
                cnt[f"{label} 지남"] += 1
            elif nxt is not None and nxt <= a.year + a.soon:
                verdict.append(f"{label} {nxt}년 예정")
                cnt[f"{label} 예정"] += 1
        if ls is None:
            per = re.sub(r"\s", "", g("보존기간"))
            frag = FRAGILE.search(" ".join((g("유형"), g("재질"), g("제목"))))
            sc = (3 if per in ("영구", "준영구") else 0) + (min((a.year - py) // 20, 5) if py else 0) + (3 if frag else 0)
            why = [x for x in (f"보존기간 {per}" if per in ("영구", "준영구") else "", f"생산 {a.year - py}년 경과" if py else "",
                               f"취약 재질({frag.group(0)})" if frag else "") if x]
            prio.append([sc, no, g("제목"), g("유형"), per, py or "", " · ".join(why)])
        w1.append([no, g("제목"), {"종이": "종이·행정박물", "오디오비디오": "오디오·비디오", "사진필름": "사진·필름", "전자매체": "전자기록물 등 보존매체"}.get(k, k),
                   grade or "", " ".join(x for x in (g("서고"), g("서가")) if x), lc or "", nc or "", ls or "",
                   (cyc * f if cyc else ""), ns or "", " · ".join(verdict) or "정상"])
    wp = out.create_sheet("첫 상태검사 우선순위")
    wp.append(["순위", "점수", "관리번호", "제목", "유형", "보존기간", "생산연도", "우선 이유"])
    for k_, x in enumerate(sorted(prio, key=lambda x: -x[0]), 1):
        wp.append([k_] + x)
    wp.append([])
    wp.append(["", "", "", "근거: 국가기록원 「기록물 상태검사 지침」(2025. 6.) 2.3.2 — 전체 대상이 원칙이나 대량이면 영구·준영구, 오래된 기록물, 보존성이 취약한 재질(감열지·감광지·트레이싱지·저급용지, 자성 오디오·비디오 테이프, 필름 매체)부터"])
    w3 = out.create_sheet("기준")
    for line in [["근거", "시행규칙 제31조·별표 14 <개정 2026. 9. 11.>"], ["정수점검", "모든 유형 2년"],
                 ["상태검사", "종이·행정박물 1등급 30년, 2등급 15년, 3등급 10년 / 영화필름 2년 / 오디오·비디오 3년 / 사진·필름 10년 / 전자기록물 등 보존매체 5년"],
                 ["70년 경과 배수", f"{factor} (영구기록물관리기관의 장이 정하는 기록물만, 최대 2배)"],
                 ["결과 서식", "별지 제7호서식 보존기록물의 정수점검·상태검사서"]]:
        w3.append(line)
    for sh, wd in ((w1, (14, 30, 18, 8, 14, 12, 12, 12, 14, 12, 34)), (w2, (8, 14, 60)), (w3, (16, 90))):
        for i, x in enumerate(wd):
            sh.column_dimensions[chr(65 + i)].width = x
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    out.save(a.out)
    print(f"{a.year}년 기준 · " + " · ".join(f"{k} {v}" for k, v in sorted(cnt.items())) + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
