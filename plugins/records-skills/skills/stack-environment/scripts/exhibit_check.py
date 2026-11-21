#!/usr/bin/env python3
"""기록물 전시 환경 점검 — 전시(대여 전시 포함) 목록으로 조도·자외선·전시 기간·재전시 간격을 보고, 전시 이력 기록부를 만든다.

근거:
  공공기록물법 시행령 별표 6: 전시관 조명 50~200럭스(국내 법정 범위)
  ANSI/NISO Z39.79-2001 「도서관·보존기록관 자료 전시 환경조건」(국가기록포털 해외표준 게시):
    3.1 전시 기간은 물건마다 정한다 — 빛에 민감한 자료를 반복 전시(2년에 1번 이하)할 때 통상 12주, 드물게(10년에 1번) 전시하면 더 길게, 일반적으로 52주 이하
    3.2 조도: 매우 민감 50lux 이하, 중간 민감 100lux 이하(민감도 모르면 매우 민감으로). 3.3 조도는 물건 표면에서, 설치·조명 교체 때 측정
    3.4 직사광선 금지 · 3.5.1 자외선 75 µW/lm 이하 · 3.6 전시 이력(전시명·기간·조도, 여러 쪽이면 펼친 쪽) 기록
    4.2 상대습도 설정값 35~50%, ±5%(30~55%) · 4.4 월 변화 5% 이하 · 5.1~5.2 온도 21℃ 이하 설정, ±3℃(25℃ 넘지 않음)
  수치는 기준을 그대로 옮겼고 문장은 새로 썼다. 기관 기준이 더 엄하면 기관 기준을 따른다.

입력: 전시 목록(xlsx·csv) — 열: 기록물(명), 민감도(매우·중간·낮음, 비우면 재질로 추정), 재질, 조도(lux), 자외선(µW/lm), 하루 조명시간, 시작일, 종료일,
      지난전시종료일, 쪽(펼친 쪽), 직사광선(Y), 습도설정, 온도설정
결과: xlsx — 점검 · 전시 이력 기록부 · 기준
예) python exhibit_check.py 전시목록.xlsx -o 산출/전시환경_점검.xlsx
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

ALIAS = {"명": ["기록물", "기록물명", "자료명", "제목"], "민감": ["민감도", "빛민감도"], "재질": ["재질", "매체"], "lux": ["조도", "조도(lux)", "lux"],
         "uv": ["자외선", "자외선(µW/lm)", "UV"], "시간": ["하루 조명시간", "조명시간", "일 조명시간"], "시작": ["시작일", "전시시작"], "종료": ["종료일", "전시종료"],
         "지난": ["지난전시종료일", "직전 전시 종료일", "지난 전시"], "쪽": ["쪽", "펼친 쪽", "전시면"], "직사": ["직사광선"], "rh": ["습도설정", "상대습도 설정"],
         "온도": ["온도설정", "온도 설정"]}
VERY = r"컬러|칼라|색|수채|잉크|만년필|볼펜|펠트|사인펜|신문|산성지|직물|섬유|가죽|양피|컬러사진|청사진|감광|팩스|감열|도면 복사|복사본|채색|한지 채색|필사"
MOD = r"인쇄|먹|흑백 ?사진|젤라틴|활판|탄소|연필"


def d(v):
    t = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y.%m.%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def num(v):
    m = re.search(r"\d+(\.\d+)?", str(v or ""))
    return float(m.group()) if m else None


def load(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        rows = [["" if v is None else str(v).strip() for v in r] for r in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)]
    else:
        rows = [[c.strip() for c in r] for r in csv.reader(open(path, encoding="utf-8-sig"))]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser(description="기록물 전시 환경 점검(시행령 별표 6, ANSI/NISO Z39.79)")
    ap.add_argument("listing")
    ap.add_argument("-o", "--out", default="산출/전시환경_점검.xlsx")
    a = ap.parse_args()
    items = load(a.listing)
    if not items:
        print("입력 오류: 목록이 비어 있습니다.", file=sys.stderr)
        return 2
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["기록물", "민감도(판단)", "조도", "전시 주수", "누적 노출(lux·h)", "판정"])
    hist = wb.create_sheet("전시 이력 기록부")
    hist.append(["기록물", "전시명(채움)", "시작일", "종료일", "전시 주수", "조도(lux)", "펼친 쪽", "점검 결과(빛 변화 여부)(채움)"])
    bad = 0
    for x in items:
        name = x.get("명", "")
        sens = x.get("민감", "")
        if not re.search(r"매우|중간|낮", sens):
            mat = x.get("재질", "")
            sens = "매우(추정)" if (re.search(VERY, mat) or not re.search(MOD, mat)) else "중간(추정)"
        very = sens.startswith("매우")
        lux, uv, hrs = num(x.get("lux")), num(x.get("uv")), num(x.get("시간")) or 8
        s0, s1, prev = d(x.get("시작")), d(x.get("종료")), d(x.get("지난"))
        weeks = round((s1 - s0).days / 7, 1) if s0 and s1 else None
        iss = []
        if lux is None:
            iss.append("조도 측정값 없음 — 물건 표면에서, 설치·조명 교체 때 측정(Z39.79 3.3)")
        else:
            lim = 50 if very else 100
            if lux > lim:
                iss.append(f"조도 {lux:g}lux > {lim}lux({'매우' if very else '중간'} 민감, Z39.79 3.2)")
            if lux > 200:
                iss.append("시행령 별표 6 전시관 50~200럭스를 넘음")
        if uv is not None and uv > 75:
            iss.append(f"자외선 {uv:g} µW/lm > 75(Z39.79 3.5.1) — 자외선 차단 필터")
        if re.match(r"(Y|예|있)", x.get("직사", "").strip()):
            iss.append("직사광선이 닿음 — 금지(Z39.79 3.4)")
        if weeks is not None:
            if weeks > 52:
                iss.append(f"전시 {weeks}주 > 52주(일반 상한)")
            elif very and weeks > 12:
                iss.append(f"빛에 민감한 자료 {weeks}주 > 12주(반복 전시 통상 기준) — 교체 전시·복제본 검토")
        elif not (s0 and s1):
            iss.append("전시 시작·종료일 없음 — 전시 기간은 물건마다 정함(Z39.79 3.1)")
        if prev and s0 and very and (s0 - prev).days < 730:
            iss.append(f"지난 전시 종료 뒤 {(s0 - prev).days}일 만에 재전시 — 민감 자료는 2년에 1번 이하가 통상")
        rh, tp = num(x.get("rh")), num(x.get("온도"))
        if rh is not None and not 35 <= rh <= 50:
            iss.append(f"습도 설정 {rh:g}% — 35~50%(±5%, Z39.79 4.2)")
        if tp is not None and tp > 21:
            iss.append(f"온도 설정 {tp:g}℃ — 21℃ 이하(±3℃, Z39.79 5.1)")
        if re.search(r"책|권|철|대장|일기|앨범", name + x.get("재질", "")) and not x.get("쪽"):
            iss.append("책·철 전시는 펼친 쪽을 기록(Z39.79 3.6) — 오래 전시하면 쪽을 바꿈")
        expo = round(lux * hrs * (s1 - s0).days) if (lux is not None and s0 and s1) else ""
        bad += bool(iss)
        ws.append([name, sens, lux if lux is not None else "", weeks if weeks is not None else "", expo, " · ".join(iss) or "통과"])
        hist.append([name, "", s0.isoformat() if s0 else "", s1.isoformat() if s1 else "", weeks if weeks is not None else "", lux if lux is not None else "", x.get("쪽", ""), ""])
    gd = wb.create_sheet("기준")
    for row in [("국내 법정", "시행령 별표 6 전시관 조명 50~200럭스"),
                ("조도", "매우 민감(컬러 잉크·수채·신문·감광·직물 등, 모르면 여기) 50lux 이하 · 중간 민감(먹·인쇄 잉크 고급지, 흑백 젤라틴 사진) 100lux 이하 — 비공개 시간엔 끄거나 낮춤"),
                ("자외선·직사광선", "자외선 75 µW/lm 이하, 직사광선·천창빛 금지, 조명 열 차단"),
                ("기간", "민감 자료 반복 전시(2년에 1번 이하) 통상 12주, 드물게(10년에 1번)면 더 길게, 일반 상한 52주"),
                ("습도·온도", "습도 35~50% 설정 ±5%(30~55%), 월 변화 5% 이하 · 온도 21℃ 이하 설정 ±3℃"),
                ("기록", "전시마다 이력(전시명·기간·조도·펼친 쪽)을 남기고 빛·습도·온도에 의한 변화를 정기 점검"),
                ("근거", "공공기록물법 시행령 별표 6, ANSI/NISO Z39.79-2001(국가기록포털 해외표준, 수치 인용·문장은 새로 씀)")]:
        gd.append(row)
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["F"].width = 100
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"전시 {len(items)}점 · 확인할 것 {bad} → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
