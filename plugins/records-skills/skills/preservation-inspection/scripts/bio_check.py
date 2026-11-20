#!/usr/bin/env python3
"""서고·기록물 생물학적 오염 점검: ATP(RLU) 표본 측정과 서고 부유미생물(CFU/m³) 측정값으로 소독·정밀검사 대상 후보를 고른다.
근거: 국가기록원 「미소독기록물 유해충·균 모니터링 및 기준 검증 연구(3차)」 제5절 기록물 소독 선별처리 기준 수정제안
        - 표본: 이관·미소독 기록물은 서가 단별 1권, 소장기록물은 서가별 1권. ATP 채취 면적 10cm×10cm, 인쇄 번짐을 막으려 공란에서 채취
        - 기준: 1,000 RLU 이상이면 소독 대상 또는 정밀검사 대상. 측정 주기 연 1회 이상
        - 종이 재질에 따라 RLU 검출 감도가 다름(트레싱지가 가장 민감, 박엽지가 가장 낮음) — 낮게 나온 박엽지는 정밀검사로 보완 검토
      실내공기질 관리 기준(같은 연구 인용): 부유세균 800 CFU/m³, 부유곰팡이 500 CFU/m³ — 사람 건강 기준이므로 서고에서는 '초과=즉시 조치',
        평소 값보다 크게 늘면 '추세 확인'으로 본다.

사용:  python bio_check.py 측정목록.xlsx -o 산출/생물오염_점검.xlsx
목록 열(머리글 이름으로 찾음): 구분(기록물·서고), 대상(관리번호·서고명), 측정일, RLU, 부유세균, 부유곰팡이(부유진균), 재질, 곰팡이육안(Y/N), 비고
원칙:  점검·후보 선정만 한다. 소독 방법(저산소·훈증·방사선 등)과 실시는 보존처리 담당자가 정한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"구분": ["구분", "측정구분"], "대상": ["대상", "관리번호", "서고", "서고명"], "일": ["측정일", "측정일자", "일자"], "RLU": ["RLU", "ATP", "ATP(RLU)"],
         "세균": ["부유세균", "부유세균(CFU/m3)", "세균"], "진균": ["부유곰팡이", "부유진균", "부유곰팡이(CFU/m3)", "진균"], "재질": ["재질", "종이재질", "용지"],
         "육안": ["곰팡이육안", "육안곰팡이", "곰팡이"], "비고": ["비고"]}
RLU_LIMIT, BACT, FUNGI = 1000, 800, 500


def num(s):
    m = re.search(r"[\d,]+(\.\d+)?", str(s or ""))
    return float(m[0].replace(",", "")) if m else None


def read(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(path, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        import csv
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    head = [re.sub(r"\s", "", h) for h in rows[0]]
    idx = {k: next((head.index(re.sub(r"\s", "", a)) for a in al if re.sub(r"\s", "", a) in head), None) for k, al in ALIAS.items()}
    return idx, [r for r in rows[1:] if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("list")
    ap.add_argument("--today")
    ap.add_argument("-o", "--out", default="산출/생물오염_점검.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    idx, rows = read(a.list)
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검 결과"
    ws.append(["원본 행", "구분", "대상", "측정일", "측정값", "판정", "조치"])
    last = {}
    hits = 0
    for n, r in enumerate(rows, 2):
        g = lambda k, r=r: (r[idx[k]] if idx[k] is not None and idx[k] < len(r) else "").strip()
        tgt, day = g("대상"), g("일")
        rlu, b, f = num(g("RLU")), num(g("세균")), num(g("진균"))
        out = []
        if re.match(r"(Y|예|유|있)", g("육안")):
            out.append(("곰팡이 육안 확인", "격리 후 소독 대상(ATP 값과 무관) — 지퍼백·별도 상자, 담당자 보호구"))
        if rlu is not None:
            if rlu >= RLU_LIMIT:
                out.append((f"ATP {rlu:g} RLU ≥ {RLU_LIMIT}", "소독 대상 또는 정밀검사(멸균 면봉 10×10cm 채취, CFU/100cm²) 대상"))
            elif re.search(r"박엽지|습자지", g("재질")) and rlu >= RLU_LIMIT * 0.3:
                out.append((f"ATP {rlu:g} RLU(박엽지는 RLU가 낮게 나옴)", "기준 미만이지만 정밀검사로 보완 검토"))
        if b is not None and b > BACT:
            out.append((f"부유세균 {b:g} CFU/m³ > {BACT}", "서고 환기·공조 필터 점검, 담당자 보호구, 원인 확인 후 재측정"))
        if f is not None and f > FUNGI:
            out.append((f"부유곰팡이 {f:g} CFU/m³ > {FUNGI}", "곰팡이 발생원(누수·결로·고습) 확인, 해당 서고 기록물 ATP 표본 측정 확대"))
        key = (g("구분"), tgt)
        if f is not None:
            prev = last.get(key)
            if prev and prev > 0 and f >= prev * 3 and f <= FUNGI:
                out.append((f"부유곰팡이 {prev:g} → {f:g}(3배 이상 증가)", "기준 이하이나 추세 확인 — 습도·누수 점검"))
            last[key] = f
        d = None
        try:
            d = dt.date.fromisoformat(day[:10].replace(".", "-")) if day else None
        except ValueError:
            pass
        if d and (today - d).days > 365:
            out.append(("측정 1년 지남", "연 1회 이상 측정(같은 연구 제안) — 다시 측정"))
        hits += bool(out)
        for j, k in out:
            ws.append([n, g("구분"), tgt, day, g("RLU") or g("세균") or g("진균"), j, k])
    s = wb.create_sheet("기준")
    for row in (["항목", "기준", "근거"],
                ["ATP 표본", "1,000 RLU 이상 → 소독·정밀검사 대상, 서가 단별(이관·미소독)·서가별(소장) 1권, 10×10cm 공란 채취, 연 1회 이상",
                 "국가기록원 「미소독기록물 유해충·균 모니터링 및 기준 검증 연구(3차)」 제5절"],
                ["종이 재질 감도", "트레싱지 가장 민감, 박엽지 가장 낮음 — 박엽지는 정밀검사 보완", "같은 연구"],
                ["부유세균", "800 CFU/m³", "실내공기질 관리 기준(같은 연구 인용)"],
                ["부유곰팡이", "500 CFU/m³", "실내공기질 관리 기준(2018 확대, 같은 연구 인용)"]):
        s.append(row)
    for i, w in enumerate((8, 10, 18, 12, 12, 34, 60)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"측정 {len(rows)}행 · 조치 필요 {hits}행 → {a.out}")
    return 1 if hits else 0


if __name__ == "__main__":
    sys.exit(main())
