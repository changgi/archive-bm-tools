#!/usr/bin/env python3
"""정보공개 청구 처리기한 계산과 결정 점검.

사용:  python foi_deadlines.py 청구목록.xlsx --today 2026-11-19 [--calc 민법|근무일] [--holidays 추가휴일.txt] -o 산출/정보공개_기한.xlsx
목록 열(머리글 이름으로 찾음, 없는 열은 빈칸으로 본다):
  접수번호, 청구일, 연장통지일, 결정일, 결정(공개·부분공개·비공개·정보부존재), 비공개사유,
  제3자통지일, 공개일, 이의신청일, 이의연장통지일, 이의결정일
기한 근거(정보공개법):
  제11조① 청구를 받은 날부터 10일 이내 결정, ② 부득이하면 그 기간이 끝나는 날의 다음 날부터 10일 범위 연장(연장 사실·사유 문서 통지)
  제9조① 비공개 대상 1~8호, 제13조④ 비공개 결정 시 이유·불복방법 통지
  제18조③ 이의신청은 받은 날부터 7일 이내 결정(부득이하면 7일 범위 연장)
  시행령 제12조① 공개 결정한 날부터 10일 이내 공개일시 통지(제3자 비공개 요청에도 공개 결정한 경우는 제외)
  제21조③ 제3자 비공개 요청에도 공개하면 공개 결정일과 공개 실시일 사이 최소 30일
  제18조① 청구 후 20일이 지나도록 결정이 없으면 청구인은 이의신청 가능
  (2026-11-20 국가법령정보센터 원문 확인: 법 [시행 2023. 11. 17.], 시행령 [시행 2026. 1. 2.])
기간 계산: 기본은 민법 방식(초일 불산입, 말일이 토·일·공휴일이면 다음 근무일).
  --calc 근무일 은 토·일·공휴일을 빼고 센다(기관 운영 기준이 그렇다면만).
  공휴일은 holidays 패키지(KR)를 쓰고 --holidays 파일(한 줄에 YYYY-MM-DD)로 임시공휴일 등을 더한다. 패키지가 없으면 주말만 뺀다(경고).
원칙: 계산과 점검만 한다. 결정·통지는 사람이 한다. 청구인 이름·연락처 열은 결과에 옮기지 않는다.
"""
import argparse
import datetime as dt
import os
import re
import sys

COLS = ["접수번호", "청구일", "연장통지일", "결정일", "결정", "비공개사유", "제3자통지일", "제3자비공개요청", "공개일", "이의신청일", "이의연장통지일", "이의결정일"]
REASON_RX = re.compile(r"([1-8])\s*호")


def parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", str(v))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


class Cal:
    def __init__(self, extra, mode):
        self.mode, self.note, self.h = mode, "", set(extra)
        try:
            import holidays
            self.kr = holidays.KR()
            self.note = f"holidays {holidays.__version__} (KR)"
        except ImportError:
            self.kr = {}
            self.note = "holidays 패키지 없음: 주말만 제외(공휴일 미반영 — 확인 필요)"
        if extra:
            self.note += f" + 추가 휴일 {len(extra)}일"

    def off(self, d):
        return d.weekday() >= 5 or d in self.kr or d in self.h

    def add(self, start, n):
        """start(초일)를 빼고 n일째 되는 날. 민법: 말일만 근무일로 밀고, 근무일: 근무일만 센다."""
        d = start
        if self.mode == "근무일":
            k = 0
            while k < n:
                d += dt.timedelta(days=1)
                if not self.off(d):
                    k += 1
            return d
        d = start + dt.timedelta(days=n)
        while self.off(d):
            d += dt.timedelta(days=1)
        return d


def read_rows(path):
    if path.lower().endswith(".csv"):
        import csv
        with open(path, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    else:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10]) if sum(1 for c in r if str(c or "").strip() in COLS) >= 2), 0)
    head = [str(c or "").strip() for c in rows[hi]]
    idx = {c: head.index(c) for c in COLS if c in head}
    out = []
    for r in rows[hi + 1:]:
        if not any(c not in (None, "") for c in r):
            continue
        out.append({c: (r[i] if i < len(r) else None) for c, i in idx.items()})
    return out, [c for c in COLS if c not in idx]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--today")
    ap.add_argument("--calc", choices=["민법", "근무일"], default="민법")
    ap.add_argument("--holidays")
    ap.add_argument("--warn-days", type=int, default=3, help="기한 며칠 전부터 '임박'으로 표시")
    ap.add_argument("-o", "--out", default="산출/정보공개_기한.xlsx")
    a = ap.parse_args()
    today = parse_date(a.today) if a.today else dt.date.today()
    extra = []
    if a.holidays:
        extra = [d for d in (parse_date(l) for l in open(a.holidays, encoding="utf-8")) if d]
    cal = Cal(extra, a.calc)
    rows, missing = read_rows(a.listing)
    if "청구일" in missing:
        print("입력 오류: '청구일' 열이 없습니다.")
        return 2

    import openpyxl
    from openpyxl.styles import PatternFill
    wb = openpyxl.Workbook()
    w1 = wb.active
    w1.title = "결정 기한"
    w1.append(["접수번호", "청구일", "결정기한(10일)", "연장 시 최종기한", "결정일", "상태", "공개실시기한", "공개일", "공개 상태"])
    w2 = wb.create_sheet("결정 점검")
    w2.append(["접수번호", "결정", "점검", "근거"])
    w3 = wb.create_sheet("이의신청")
    w3.append(["접수번호", "이의신청일", "결정기한(7일)", "연장 시 최종기한(+7일)", "이의결정일", "상태"])
    red, yel = PatternFill("solid", fgColor="FBEBEB"), PatternFill("solid", fgColor="F9EEDA")
    cnt = {"초과": 0, "임박": 0, "점검": 0}

    def state(done, due, due2):
        if done:
            if done <= due:
                return "기한 내 처리"
            return "연장 기한 내 처리" if due2 and done <= due2 else "기한 넘겨 처리"
        lim = due2 or due
        left = (lim - today).days
        if left < 0:
            return f"기한 초과 {-left}일"
        return f"진행 중 D-{left}" + (" (임박)" if left <= a.warn_days else "")

    for r in rows:
        no = str(r.get("접수번호") or "")
        req = parse_date(r.get("청구일"))
        if not req:
            w2.append([no, "", "청구일 없음 또는 형식 오류", "제11조①"])
            cnt["점검"] += 1
            continue
        ext = parse_date(r.get("연장통지일"))
        dec = parse_date(r.get("결정일"))
        due = cal.add(req, 10)
        due2 = cal.add(due, 10) if ext else None
        st = state(dec, due, due2)
        kind = str(r.get("결정") or "").strip()
        pub = parse_date(r.get("공개일"))
        objected = str(r.get("제3자비공개요청") or "").strip() not in ("", "N", "아니오", "없음")
        pdue = cal.add(dec, 10) if dec and kind in ("공개", "부분공개") and not objected else None
        pst = ""
        if dec and kind in ("공개", "부분공개") and objected:
            earliest = dec + dt.timedelta(days=30)
            pst = (f"제3자 요청: {earliest} 이후 공개 가능" if not pub else
                   ("30일 간격 지킴" if (pub - dec).days >= 30 else f"30일 간격 미달({(pub - dec).days}일) — 제21조③"))
        if pdue:
            pst = ("기한 내 공개" if pub <= pdue else "공개일 기한 넘김") if pub else \
                (f"공개 예정 D-{(pdue - today).days}" if pdue >= today else f"공개 기한 초과 {(today - pdue).days}일")
        w1.append([no, req.isoformat(), due.isoformat(), due2.isoformat() if due2 else "", dec.isoformat() if dec else "",
                   st, pdue.isoformat() if pdue else "", pub.isoformat() if pub else "", pst])
        if "초과" in st or "넘겨" in st or "넘김" in pst or "초과" in pst:
            cnt["초과"] += 1
            for c in w1[w1.max_row]:
                c.fill = red
        elif "임박" in st:
            cnt["임박"] += 1
            for c in w1[w1.max_row]:
                c.fill = yel

        def chk(msg, basis):
            w2.append([no, kind, msg, basis])
            cnt["점검"] += 1
        if not dec and (today - req).days >= 20:
            chk(f"청구 후 {(today - req).days}일 경과, 결정 없음 — 청구인 이의신청 가능", "제18조①(20일 경과한 날부터 30일 이내)")
        if "미달" in pst:
            chk(pst, "제21조③")
        if ext and ext > due:
            chk(f"연장 통지일({ext})이 원래 기한({due}) 뒤", "제11조② 기간 이내 연장·통지")
        if dec and dec > due and not ext:
            chk("기한을 넘겨 결정했는데 연장 통지 기록 없음", "제11조②")
        if kind in ("비공개", "부분공개"):
            reason = str(r.get("비공개사유") or "")
            ho = sorted(set(REASON_RX.findall(reason)))
            if not ho:
                chk("비공개 사유에 제9조제1항 호(1~8호)가 없음", "제9조①, 제13조④")
            elif len(re.sub(r"(제?\s*9\s*조)?\s*(제?\s*1\s*항)?|[1-8]\s*호|[,·\s]|에 해당|해당", "", reason)) < 6:
                chk(f"사유가 호({','.join(ho)}호)만 있고 구체적 이유가 없음", "제13조④ 비공개 이유 명시")
        if kind and kind not in ("공개", "부분공개", "비공개", "정보부존재", "민원 이첩", "취하", "종결"):
            chk(f"결정 구분 '{kind}'이 표준 구분과 다름", "확인")
        tp = parse_date(r.get("제3자통지일"))
        if tp and dec and kind in ("공개", "부분공개") and (dec - tp).days < 3:
            chk(f"제3자 통지({tp}) 후 3일 안에 공개 결정({dec}) — 비공개 요청 기간 확인", "제21조①")
        ob = parse_date(r.get("이의신청일"))
        if ob:
            od = parse_date(r.get("이의결정일"))
            odue = cal.add(ob, 7)
            w3.append([no, ob.isoformat(), odue.isoformat(), cal.add(odue, 7).isoformat(),
                       od.isoformat() if od else "", state(od, odue, cal.add(odue, 7) if parse_date(r.get("이의연장통지일")) else None)])

    w4 = wb.create_sheet("계산 기준")
    for line in [["오늘", today.isoformat()], ["기간 계산", a.calc + (" (초일 불산입, 말일이 휴일이면 다음 근무일)" if a.calc == "민법" else " (토·일·공휴일 제외)")],
                 ["휴일 목록", cal.note], ["없는 열", ", ".join(missing) or "없음"],
                 ["주의", "기간 계산 방식은 기관의 정보공개 운영 기준으로 확인하고, 결과는 검토 자료로만 쓴다."]]:
        w4.append(line)
    for ws, widths in ((w1, (14, 12, 14, 16, 12, 18, 14, 12, 16)), (w2, (14, 10, 52, 28)), (w3, (14, 12, 14, 20, 12, 18)), (w4, (12, 70))):
        for i, wdt in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = wdt
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"청구 {len(rows)}건 · 기한 초과·넘김 {cnt['초과']} · 임박 {cnt['임박']} · 점검 {cnt['점검']} → {a.out}  [{cal.note}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
