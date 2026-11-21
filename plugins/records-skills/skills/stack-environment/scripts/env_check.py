#!/usr/bin/env python3
"""서고 환경 점검: 온습도 기록계 자료를 시행령 별표 6 「시설ㆍ장비 및 보존환경 기준」에 견주어
기준 밖 비율·가장 길게 이어진 이탈 시간·하루 변동·기록 공백을 서고별로 보고, 필요한 서고 면적과 장비 점검표를 만든다.

사용:  python env_check.py 온습도.csv [다른파일.xlsx ...] --org 기록관|특수기록관|영구 [--media "1서고=종이,2서고=자기"]
         [--books 25000 --shelving 고정식|이동식 --area 180] [--gap-hours 3] -o 산출/서고환경.xlsx
자료 열(머리글 이름으로 찾음): 시각(일시·측정일시·datetime·time), 온도(℃·temp), 습도(%·RH·humidity), [서고(위치·기록계·장소)], [매체]
       서고 열이 없으면 파일 이름을 서고 이름으로 쓴다.
기준(시행령 별표 6, 2026. 8. 25. 개정, law.go.kr 2026-11-20 확인):
  종이 20±2℃·50±5% / 전자 20±2℃·40±5% / 행정박물 20±2℃·50±5% (영구·특수기록관)
  시청각 필름매체 0±2℃·30±5%, 자기매체 15±2℃·40±5% (흑백 사진필름·마이크로필름은 자기매체 기준) (영구·특수기록관)
  기록관은 종이·전자만 정함
  조명 서고 100~300럭스(자외선 차단등), 전시 50~200럭스 / 공기질(영구) PM-10 50㎍/㎥, SO2·NOx 0.05ppm 이하, HCHO 120㎍/㎥, VOC 400㎍/㎥
  서고 면적 종이 고정식 1만권당 99㎡, 이동식은 고정식 면적의 40~60% / 하중 고정식 750, 이동식 1,000~1,200㎏f/㎡
원칙:  점검과 계산만 한다. 설비 조치는 사람이 한다. 기록계 시각은 기록계 설정 그대로 읽는다(KST 가정).
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys
from collections import defaultdict

STD = {  # (온도 중심, 허용, 습도 중심, 허용)
    "종이": (20, 2, 50, 5), "전자": (20, 2, 40, 5), "행정박물": (20, 2, 50, 5),
    "필름": (0, 2, 30, 5), "자기": (15, 2, 40, 5),
}
ORG_MEDIA = {"영구": set(STD), "특수기록관": set(STD), "기록관": {"종이", "전자"}}
MEDIA_ALIAS = {"종이": "종이", "문서": "종이", "전자": "전자", "전자기록물": "전자", "행정박물": "행정박물", "박물": "행정박물",
               "필름": "필름", "필름매체": "필름", "영화필름": "필름", "컬러필름": "필름",
               "자기": "자기", "자기매체": "자기", "테이프": "자기", "흑백필름": "자기", "흑백사진필름": "자기", "마이크로필름": "자기"}
COLS = {"시각": ["시각", "일시", "측정일시", "측정시각", "날짜시간", "datetime", "date time", "time", "timestamp", "date"],
        "온도": ["온도", "온도(℃)", "온도(°c)", "temp", "temperature", "temp(°c)", "temperature(°c)"],
        "습도": ["습도", "습도(%)", "상대습도", "rh", "rh(%)", "humidity", "humidity(%rh)", "humidity(%)"],
        "서고": ["서고", "위치", "기록계", "장소", "logger", "location"], "매체": ["매체", "보존매체", "기록물유형"]}
EQUIP = {
    "영구": ["항온·항습 설비, 환경적응장비(시청각)", "온·습도계 서고당 1대", "자동소화시설(서고는 가스식)", "폐쇄회로 감시장치",
            "탈산·소독 장비(용역이면 생략 가능)", "복원·시청각 장비(용역이면 생략 가능)", "매체수록 장비(전자매체·마이크로필름, 용역이면 생략 가능)",
            "공기질: PM-10 50㎍/㎥, SO2 0.05ppm 이하, NOx 0.05ppm 이하, HCHO 120㎍/㎥, VOC 400㎍/㎥",
            "조명: 서고 100~300럭스(자외선 차단등), 전시관 50~200럭스", "신축·개축: 불연재료·유해물질 미방출 자재"],
    "특수기록관": ["항온·항습 설비", "온·습도계 서고당 1대", "자동소화시설(서고는 가스식)", "폐쇄회로 감시장치", "소독 장비(용역이면 생략 가능)",
              "매체수록 장비(용역이면 생략 가능)", "조명: 서고 100~300럭스(자외선 차단등), 전시관 50~200럭스",
              "시청각이 소량이면 저온함 활용 가능", "하중 미충족 시 기록물·서가 중량 조정", "신축·개축: 불연재료·유해물질 미방출 자재"],
    "기록관": ["항온·항습 설비", "온·습도계 서고당 1대", "가스식 휴대형 소화기", "이중 잠금장치", "매체수록 장비(용역이면 생략 가능)",
            "조명: 서고 100~300럭스(자외선 차단등), 전시관 50~200럭스", "하중 미충족 시 기록물·서가 중량 조정", "신축·개축: 불연재료·유해물질 미방출 자재"],
}

# NAK 11:2025(v2.0) 기록물관리기관 보존환경 기준 — 국가기록원, 2026-11-20 내려받아 확인
NAK11_CHECK = [
    ["위치·구조", "서고에 창문이 없는가(있으면 밀폐·자외선 차단 처리)"],
    ["위치·구조", "서고 안이나 바로 위·옆에 물을 쓰는 공간(화장실·배관·급수)이 없는가"],
    ["위치·구조", "내화 구조·불연 자재로 되어 있는가"],
    ["위치·구조", "출입문 잠금장치가 있고 출입을 관리하는가"],
    ["위치·구조", "자기매체 서고는 전자파 차폐가 되어 있는가(해당 시)"],
    ["서가", "최하단이 바닥에서 8.5~15cm 떨어져 있는가"],
    ["서가", "외벽에서 떨어져 설치했는가"],
    ["서가", "통로 폭 0.7~1m, 중앙통로 1.5~2m를 확보했는가"],
    ["서가", "서고 면적 대비 서가 면적 비율이 40~70%인가"],
    ["온습도 측정", "온습도계 분해능이 1℃·1% 이하인가"],
    ["온습도 측정", "기류 0.3m/s 이하, 벽·서가에서 30cm 떨어진 곳에 두었는가"],
    ["온습도 측정", "서고 대표지점과 외벽 가까운 극단지점을 함께 재는가"],
    ["온습도 측정", "디지털 측정기는 3회 측정해 평균을 쓰는가"],
    ["온습도 측정", "측정 때 날씨·실외 온습도·측정자를 기록하는가"],
    ["온습도 측정", "온습도가 급격히 바뀌지 않게 관리하는가(하루 변동)"],
]


def num(v):
    m = re.search(r"-?\d+(?:\.\d+)?", str(v or "").replace(",", "."))
    return float(m[0]) if m else None


def when(v):
    if isinstance(v, dt.datetime):
        return v
    s = str(v or "").strip().replace("T", " ").replace("/", "-").replace(".", "-")
    s = re.sub(r"\s+", " ", s).rstrip("-")
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d %p %I:%M:%S", "%Y-%m-%d %H", "%Y%m%d%H%M%S", "%Y%m%d %H%M", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    s2 = s.replace("오전", "AM").replace("오후", "PM")
    for f in ("%Y-%m-%d %p %I:%M:%S", "%Y-%m-%d %p %I:%M"):
        try:
            return dt.datetime.strptime(s2, f)
        except ValueError:
            pass
    return None


def read_table(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return [r for r in csv.reader(f)]
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"{path}: 문자 인코딩을 읽지 못함")


def find_cols(rows):
    for hi, r in enumerate(rows[:15]):
        low = [str(c or "").strip().lower() for c in r]
        idx = {}
        for k, al in COLS.items():
            for a_ in al:
                if a_.lower() in low and k not in idx:
                    idx[k] = low.index(a_.lower())
        if {"시각", "온도", "습도"} <= idx.keys():
            return hi, idx
    return None, None


def norm_media(v):
    s = re.sub(r"\s", "", str(v or ""))
    return MEDIA_ALIAS.get(s, "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--org", choices=["기록관", "특수기록관", "영구"], default="기록관")
    ap.add_argument("--media", default="", help='서고별 매체 "1서고=종이,2서고=자기" (없으면 매체 열, 그것도 없으면 종이)')
    ap.add_argument("--gap-hours", type=float, default=3, help="이 시간보다 긴 기록 공백을 표시")
    ap.add_argument("--books", type=int, help="보존 권 수(종이) — 필요 서고 면적 계산")
    ap.add_argument("--shelving", choices=["고정식", "이동식"], default="고정식")
    ap.add_argument("--area", type=float, help="실제 서고 면적(㎡)")
    ap.add_argument("--pages", type=int, help="보존 매수(종이) — 100매 = 1권으로 권 수 환산(NAK 11:2025)")
    ap.add_argument("--shelf-m", type=float, help="서가에 꽂힌 길이 합계(m) — 1권 15~20mm로 권 수 범위 환산(NAK 11:2025)")
    ap.add_argument("--calibrated", help="온습도계 마지막 검·교정일 YYYY-MM-DD — 1년이 지나면 표시(NAK 11:2025 매년 검·교정)")
    ap.add_argument("--today", help="판정일(기본: 오늘)")
    ap.add_argument("--audio", type=int, help="오디오 매체 수(1만개당 25㎡)")
    ap.add_argument("--video", type=int, help="비디오 매체 수(1만개당 45㎡)")
    ap.add_argument("--albums", type=int, help="사진·필름 앨범 권 수(1만권당 145㎡)")
    ap.add_argument("--film-cans", type=int, help="영화필름 캔 수(1천캔당 30㎡)")
    ap.add_argument("--workers", type=int, help="작업실 근무 인원(1명당 7㎡, 장비공간 별도)")
    ap.add_argument("--seats", type=int, help="열람실 근무 인원+열람 좌석(1명당 7㎡)")
    ap.add_argument("--newbuild", action="store_true", help="보존시설 신축·리모델링 점검표(NAK 24:2023) 시트를 더함")
    ap.add_argument("--completed", help="신축·리모델링 완공일 YYYY-MM-DD — 6개월 이상 온습도 안정화 후 입고")
    ap.add_argument("-o", "--out", default="산출/서고환경.xlsx")
    a = ap.parse_args()
    mmap = {}
    for part in filter(None, (p.strip() for p in a.media.split(","))):
        if "=" in part:
            k, v = part.split("=", 1)
            mmap[k.strip()] = norm_media(v) or v.strip()
    data = defaultdict(list)     # 서고 -> [(시각, 온도, 습도, 매체)]
    bad_rows = 0
    for path in a.files:
        rows = read_table(path)
        hi, idx = find_cols(rows)
        if hi is None:
            print(f"{path}: 시각·온도·습도 열을 찾지 못함", file=sys.stderr)
            return 2
        default = os.path.splitext(os.path.basename(path))[0]
        for r in rows[hi + 1:]:
            if not any(c not in (None, "") for c in r):
                continue
            g = lambda k: r[idx[k]] if k in idx and idx[k] < len(r) else None
            t, te, rh = when(g("시각")), num(g("온도")), num(g("습도"))
            if not t or te is None or rh is None:
                bad_rows += 1
                continue
            stack = str(g("서고") or default).strip()
            data[stack].append((t, te, rh, norm_media(g("매체"))))
    if not data:
        print("읽은 측정값이 없습니다.", file=sys.stderr)
        return 2

    import openpyxl
    from openpyxl.styles import PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    sm = wb.active
    sm.title = "요약"
    sm.append(["서고", "매체", "기준 온도", "기준 습도", "측정 수", "기간", "온도 범위 밖 %", "습도 범위 밖 %", "둘 중 하나라도 밖 %",
               "가장 긴 이탈(시간)", "하루 온도 변동 최대(℃)", "하루 습도 변동 최대(%)", "기록 공백", "판정"])
    ev = wb.create_sheet("이탈 구간")
    ev.append(["서고", "시작", "끝", "지속(시간)", "온도 최저~최고", "습도 최저~최고", "무엇이 벗어났나"])
    dy = wb.create_sheet("일별")
    dy.append(["서고", "날짜", "온도 평균", "온도 최저", "온도 최고", "습도 평균", "습도 최저", "습도 최고", "범위 밖 측정 수"])
    mo = wb.create_sheet("월별")
    mo.append(["서고", "월", "측정 수", "온도 평균", "습도 평균", "범위 밖 %"])
    worst = []
    niso = []
    for stack in sorted(data):
        pts = sorted(data[stack])
        media = mmap.get(stack) or next((m for *_, m in pts if m), "") or "종이"
        if media not in STD:
            media = "종이"
        notes = []
        if media not in ORG_MEDIA[a.org]:
            notes.append(f"{a.org} 기준표에 '{media}' 칸이 없음 — 영구·특수기록관 기준으로 봄")
        tc, tt, hc, ht = STD[media]
        out_t = [abs(p[1] - tc) > tt for p in pts]
        out_h = [abs(p[2] - hc) > ht for p in pts]
        out_any = [x or y for x, y in zip(out_t, out_h)]
        # 이탈 구간
        runs, start = [], None
        for i, o in enumerate(out_any + [False]):
            if o and start is None:
                start = i
            elif not o and start is not None:
                seg = pts[start:i]
                hrs = (seg[-1][0] - seg[0][0]).total_seconds() / 3600
                what = "·".join(w for w, f in (("온도", any(out_t[start:i])), ("습도", any(out_h[start:i]))) if f)
                runs.append((hrs, seg, what))
                start = None
        longest = max((r[0] for r in runs), default=0)
        for hrs, seg, what in runs:
            if hrs >= 1 or len(seg) >= 3:
                ev.append([stack, seg[0][0].strftime("%Y-%m-%d %H:%M"), seg[-1][0].strftime("%Y-%m-%d %H:%M"), round(hrs, 1),
                           f"{min(p[1] for p in seg):.1f}~{max(p[1] for p in seg):.1f}", f"{min(p[2] for p in seg):.0f}~{max(p[2] for p in seg):.0f}", what])
        # 일별·월별
        byday, bymon = defaultdict(list), defaultdict(list)
        for p, o in zip(pts, out_any):
            byday[p[0].date()].append((p, o))
            bymon[p[0].strftime("%Y-%m")].append((p, o))
        dt_max = dh_max = 0
        for d, lst in sorted(byday.items()):
            ts, hs = [p[1] for p, _ in lst], [p[2] for p, _ in lst]
            dt_max, dh_max = max(dt_max, max(ts) - min(ts)), max(dh_max, max(hs) - min(hs))
            dy.append([stack, d.isoformat(), round(sum(ts) / len(ts), 1), min(ts), max(ts), round(sum(hs) / len(hs), 1), min(hs), max(hs),
                       sum(o for _, o in lst)])
        for m, lst in sorted(bymon.items()):
            mo.append([stack, m, len(lst), round(sum(p[1] for p, _ in lst) / len(lst), 1), round(sum(p[2] for p, _ in lst) / len(lst), 1),
                       round(100 * sum(o for _, o in lst) / len(lst), 1)])
        gaps = [(b[0] - a_[0]).total_seconds() / 3600 for a_, b in zip(pts, pts[1:])]
        long_gaps = [g for g in gaps if g > a.gap_hours]
        pct = lambda xs: round(100 * sum(xs) / len(xs), 1)
        p_any = pct(out_any)
        verdict = "기준 안" if p_any == 0 else ("주의" if p_any < 5 and longest < 24 else "조치 필요")
        if long_gaps:
            notes.append(f"{a.gap_hours:g}시간 넘는 기록 공백 {len(long_gaps)}번(최장 {max(long_gaps):.0f}시간)")
        sm.append([stack, media, f"{tc}±{tt}℃", f"{hc}±{ht}%", len(pts), f"{pts[0][0]:%Y-%m-%d} ~ {pts[-1][0]:%Y-%m-%d}",
                   pct(out_t), pct(out_h), p_any, round(longest, 1), round(dt_max, 1), round(dh_max, 1),
                   len(long_gaps), verdict + (" · " + " · ".join(notes) if notes else "")])
        if verdict == "조치 필요":
            sm.cell(sm.max_row, 14).fill = red
        worst.append((p_any, stack))
        mavg = [(sum(p[1] for p, _ in lst) / len(lst), sum(p[2] for p, _ in lst) / len(lst)) for _, lst in sorted(bymon.items())]
        mdt = max((abs(b[0] - a_[0]) for a_, b in zip(mavg, mavg[1:])), default=0)
        mdh = max((abs(b[1] - a_[1]) for a_, b in zip(mavg, mavg[1:])), default=0)
        if media == "종이":
            niso.append((stack, max(p[1] for p in pts), min(p[2] for p in pts), max(p[2] for p in pts), dt_max, dh_max, mdt, mdh))
    if niso:  # 참고: NISO TR01-1995 「종이기록물 저장 환경 가이드라인」(국가기록포털 해외표준) — 법정 기준(별표 6)이 우선, 비교용
        nz = wb.create_sheet("참고 NISO TR01(종이)")
        nz.append(["서고", "최고 온도(℃)", "서고 전용 18.3℃(65°F) 이하", "열람 겸용 21.1℃(70°F) 이하", "습도 범위(%)", "30~50% 안",
                   "하루 온도 폭(℃)", "≤2.2℃(±2°F)", "하루 습도 폭(%)", "≤6%(±3%)", "월 평균 온도 변화(℃)", "≤1.7℃(3°F)", "월 평균 습도 변화(%)", "≤3%"])
        ok = lambda f: "예" if f else "아니오"
        for st, tmax, hmin, hmax, dtm, dhm, mdt, mdh in niso:
            nz.append([st, round(tmax, 1), ok(tmax <= 18.3), ok(tmax <= 21.1), f"{hmin:.0f}~{hmax:.0f}", ok(hmin >= 30 and hmax <= 50),
                       round(dtm, 1), ok(dtm <= 2.2), round(dhm, 1), ok(dhm <= 6), round(mdt, 1), ok(mdt <= 1.7), round(mdh, 1), ok(mdh <= 3)])
        nz.append([])
        nz.append(["안내", "NISO TR01-1995 제안값(비규범, 비교용): 온도 서고 전용 65°F·열람 겸용 70°F 이하, 습도 30~50%, 하루 ±2°F·±3%, 월 변화 3°F·3%, 창·형광등 빛은 415nm 이하 차단, SO2·NO2·O3 5~10ppb, 공기 여과 서고 90~95%. 국내 법정 기준은 시행령 별표 6(종이 20±2℃·50±5%)이다."])
    ar = wb.create_sheet("서고 면적")
    ar.append(["항목", "값", "근거"])
    if not a.books and a.pages:
        a.books = -(-a.pages // 100)
        ar.append(["권 수(매수 환산)", f"{a.books:,}권", f"{a.pages:,}매 ÷ 100매(A4 100매 = 1권, NAK 11:2025)"])
    if a.shelf_m:
        lo, hi = int(a.shelf_m * 1000 / 20), int(a.shelf_m * 1000 / 15)
        ar.append(["권 수(서가 길이 환산)", f"{lo:,}~{hi:,}권", f"{a.shelf_m:g}m ÷ 1권 15~20mm(NAK 11:2025)"])
        if not a.books:
            a.books = hi
            ar.append(["면적 계산에 쓴 권 수", f"{hi:,}권", "범위의 큰 값(보수적)"])
    if a.books:
        fixed = a.books / 10000 * 99
        if a.shelving == "고정식":
            ar.append(["필요 면적(고정식)", f"{fixed:.1f}㎡", "종이 1만권당 99㎡(시행령 별표 6)"])
            need_lo, need_hi = fixed, fixed
        else:
            need_lo, need_hi = fixed * 0.4, fixed * 0.6
            ar.append(["필요 면적(이동식)", f"{need_lo:.1f}~{need_hi:.1f}㎡", "고정식 면적의 40~60%"])
        if a.area:
            ok = a.area >= need_lo
            ar.append(["실제 면적", f"{a.area:.1f}㎡", "충족" if ok else f"부족 — 최소 {need_lo - a.area:.1f}㎡ 더 필요"])
        ar.append(["하중", "고정식 750㎏f/㎡, 이동식 1,000~1,200㎏f/㎡", "미충족 시 기록물·서가 중량 조정(기록관·특수기록관 비고)"])
    else:
        ar.append(["필요 면적", "--books(보존 권 수)를 주면 계산", "종이 1만권당 99㎡"])
    for lab, val, per, unit in (("오디오", a.audio, 10000, 25), ("비디오", a.video, 10000, 45), ("사진·필름 앨범", a.albums, 10000, 145),
                                ("영화필름", a.film_cans, 1000, 30), ("작업실", a.workers, 1, 7), ("열람실", a.seats, 1, 7)):
        if val:
            ar.append([f"필요 면적({lab})", f"{val / per * unit:.1f}㎡", f"{'1명' if per == 1 else f'{per:,}개(권·캔)'}당 {unit}㎡(시행령 별표 6, NAK 24 부속서)"])
    if a.area:
        size = "대형" if a.area >= 300 else "중형" if a.area >= 150 else "소형"
        ar.append(["서고 규모", size, "대형 300㎡ 이상 · 중형 150~300㎡ · 소형 150㎡ 이하(NAK 11:2025) — 온습도계 대표지점 수 판단에 참고"])
    eq = wb.create_sheet("시설·장비 점검표")
    eq.append([f"{a.org} 시설·장비 기준(시행령 별표 6)", "있음/없음", "비고"])
    for item in EQUIP[a.org]:
        eq.append([item, "", ""])
    ck = wb.create_sheet("서고 점검표(NAK 11)")
    ck.append(["구분", "점검 항목(NAK 11:2025 「기록물관리기관 보존환경 기준」)", "적합/부적합", "비고"])
    for line in NAK11_CHECK:
        ck.append([*line, "", ""])
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    if a.calibrated:
        try:
            cal = dt.date.fromisoformat(a.calibrated)
            due = cal.replace(year=cal.year + 1) if not (cal.month == 2 and cal.day == 29) else dt.date(cal.year + 1, 3, 1)
            ck.append(["검·교정", f"마지막 검·교정 {cal.isoformat()} → 다음 {due.isoformat()}", "부적합" if today > due else "적합",
                       "1년이 지남 — 공인 교정기관 검·교정 필요" if today > due else f"{(due - today).days}일 남음"])
        except ValueError:
            ck.append(["검·교정", f"날짜를 읽지 못함: {a.calibrated}", "", "YYYY-MM-DD"])
    else:
        ck.append(["검·교정", "온습도계 검·교정일을 --calibrated로 주면 1년 경과를 판정", "", ""])
    if a.newbuild:
        nb = wb.create_sheet("신축 점검표(NAK 24)")
        nb.append(["단계", "점검 항목(NAK 24:2023 「기록물 보존시설 신축 가이드라인」)", "적합/부적합", "비고"])
        for line in [("부지", "적법성 — 건축 관련 법령 조건에 맞는 곳"), ("부지", "안전성 — 지형·지반·수해·화재로부터 안전(수해 지역은 가능한 피함)"),
                     ("부지", "접근성 — 이관기관·열람자 접근 쉬움"), ("부지", "확장성 — 향후 서고 확장 가능"), ("부지", "공해와 격리 — 공장·교통 등 유해환경과 떨어짐"),
                     ("부지", "유사기관 연계 — 박물관·도서관·학술기관과 가까움"),
                     ("설계", "보존서고는 독립 건축 원칙, 기존 시설 활용 시 부속서(시행령 별표 6) 요구사항 충족"),
                     ("설계", "서고는 건물 내 독립 위치, 침수 대비 가급적 지상 배치"), ("설계", "서고 크기: 대형 300㎡ 초과 · 중형 150~300㎡ · 소형 150㎡ 미만"),
                     ("설계", "서고 하중: 고정식 750kgf/㎡ 이상, 이동식 1,000~1,200kgf/㎡(6단 복식서가, 단 길이 85~90㎝) — 설계 초기부터 전문가 협의"),
                     ("설계", "서가 열 사이 0.7~1m, 중앙통로 확보, 서고 전체에서 등거리 이격"), ("설계", "항온항습·공기정화, 화재감시·가스식 진압, 전기·조명 설비"),
                     ("공사", "서고 내부 건축자재 P.A.T.(Photographic Activity Test) 사전 실시 — 금속편법·요오드-아지드법"),
                     ("공사", "사용금지 건축자재를 시공 단계별로 확인"),
                     ("안정성", "완공 후 6개월 이상 온습도 점검·안정화 확인 뒤 기록물 입고"), ("안정성", "서고 상·하층부에 급·배수라인 통과 금지, 전기배선·전원장치 안전 점검"),
                     ("안정성", "완공 후 하중 정밀 점검(전문가 의뢰), 설계 하중과 비교")]:
            nb.append([*line, "", ""])
        if a.completed:
            try:
                c = dt.date.fromisoformat(a.completed)
                m = c.month + 6
                ok_day = dt.date(c.year + (m - 1) // 12, (m - 1) % 12 + 1, min(c.day, 28))
                today_ = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
                nb.append(["입고 가능", f"완공 {c.isoformat()} + 6개월 = {ok_day.isoformat()} 이후(온습도 안정화 확인 전제)",
                           "아직 이름" if today_ < ok_day else "기간 지남", ""])
            except ValueError:
                nb.append(["입고 가능", f"날짜를 읽지 못함: {a.completed}", "", ""])
        nb.column_dimensions["A"].width = 10
        nb.column_dimensions["B"].width = 100
    st = wb.create_sheet("기준")
    st.append(["매체", "온도", "습도", "적용 기관"])
    for m, (tc, tt, hc, ht) in STD.items():
        st.append([m, f"{tc}±{tt}℃", f"{hc}±{ht}%", "영구·특수기록관" + (", 기록관" if m in ("종이", "전자") else "")])
    st.append(["비고", "흑백 사진필름·마이크로필름은 자기매체 기준", "", "시행령 별표 6 <2026. 8. 25.>"])
    for ws, widths in ((sm, (14, 8, 10, 10, 8, 24, 10, 10, 10, 10, 12, 12, 8, 60)), (ev, (14, 17, 17, 10, 14, 14, 12)),
                       (eq, (70, 12, 30)), (ar, (20, 34, 50)), (ck, (12, 90, 12, 40)), (st, (12, 12, 12, 40))):
        for i, w in enumerate(widths):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    worst.sort(reverse=True)
    print(f"서고 {len(data)}곳 · 측정 {sum(len(v) for v in data.values())}개(읽지 못한 행 {bad_rows}) · 기준 밖 최대 "
          f"{worst[0][1]} {worst[0][0]}% → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
