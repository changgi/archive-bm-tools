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
    ar = wb.create_sheet("서고 면적")
    ar.append(["항목", "값", "근거"])
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
    eq = wb.create_sheet("시설·장비 점검표")
    eq.append([f"{a.org} 시설·장비 기준(시행령 별표 6)", "있음/없음", "비고"])
    for item in EQUIP[a.org]:
        eq.append([item, "", ""])
    st = wb.create_sheet("기준")
    st.append(["매체", "온도", "습도", "적용 기관"])
    for m, (tc, tt, hc, ht) in STD.items():
        st.append([m, f"{tc}±{tt}℃", f"{hc}±{ht}%", "영구·특수기록관" + (", 기록관" if m in ("종이", "전자") else "")])
    st.append(["비고", "흑백 사진필름·마이크로필름은 자기매체 기준", "", "시행령 별표 6 <2026. 8. 25.>"])
    for ws, widths in ((sm, (14, 8, 10, 10, 8, 24, 10, 10, 10, 10, 12, 12, 8, 60)), (ev, (14, 17, 17, 10, 14, 14, 12)),
                       (eq, (70, 12, 30)), (ar, (20, 34, 50)), (st, (12, 12, 12, 40))):
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
