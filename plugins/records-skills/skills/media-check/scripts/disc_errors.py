#!/usr/bin/env python3
"""광디스크 오류율 판정: CD·DVD·BD 검사 장비가 낸 오류율 값으로 보존 등급(계속 보존·빠른 이전·즉시 이전)과 다음 검사일을 낸다.
근거: 국가기록원 「전자기록물 이관·보존 매체의 기준 설정 및 규격화 연구」 표 15~17이 정리한 기준
  DVD(PI sum 8) ISO/IEC 29121 — 초기: 100 미만 양호, 100 이상 장기보존 부적합(빠른 시일 내 이전)
                               정기(3년 주기, 더 짧게 권장): 140 미만 양호, 140~280 신속한 대책, 280 이상 즉시 대책
                (참고) JIS Z 6017 같은 값(초기 100 단일), 중국 DA/T 38-2008 초기 80·정기 140/180/240 — 2등급이면 주기 1년, 3등급 6개월
  BD(R-SER / Burst) JIIMA — 초기: 3.5E-4 미만이고 800바이트 미만이면 양호, 아니면 즉시 대책
                            정기: 5E-4 미만·800바이트 이하 양호, 5E-4~1E-3 또는 800~1,900바이트 신속한 대책, 1E-3 이상 또는 1,900바이트 이상 즉시 대책
  CD(BLER) — 220이면 읽기 실패 가능성이 높은 한계(같은 연구 인용). 50 미만 양호·50~220 주의·220 이상 즉시 이전(이 스크립트의 구간)
  공통: 수정 불가능 오류(Uncorrectable)가 하나라도 있으면 즉시 이전하고 소실 데이터를 확인(DA/T 38).

사용:  python disc_errors.py 측정목록.xlsx [--standard iso|china] [--today 2026-11-20] -o 산출/광디스크_오류율.xlsx
목록 열(머리글 이름으로 찾음): 매체번호, 종류(CD·DVD·BD), 검사구분(초기·정기), 측정일, PI sum 8(PISum8·PIE), R-SER(SER), Burst(버스트), BLER, 수정불가(Uncorrectable)
원칙:  측정은 오류율 검사 장비로 사람이 한다. 장비마다 값이 달라 같은 장비로 추세를 보라고 연구도 권한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"번호": ["매체번호", "관리번호", "디스크번호"], "종류": ["종류", "매체종류", "매체"], "구분": ["검사구분", "구분"], "일": ["측정일", "검사일", "측정일자"],
         "pi": ["PIsum8", "PISum8", "PI sum 8", "PIE", "PIESUM8"], "ser": ["R-SER", "RSER", "SER"], "burst": ["Burst", "버스트", "BurstError"],
         "bler": ["BLER"], "uncor": ["수정불가", "Uncorrectable", "POF", "수정불가오류"]}


def num(s):
    s = str(s or "").strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        m = re.search(r"\d+(\.\d+)?([eE][-+]?\d+)?", s)
        return float(m[0]) if m else None


def grade(kind, initial, v, std):
    """(등급 1~3, 판정, 다음 검사까지 개월) — 값이 없으면 None"""
    k = kind.upper()
    if "BD" in k or "블루" in kind:
        ser, burst = v["ser"], v["burst"]
        if ser is None and burst is None:
            return None
        ser, burst = ser or 0, burst or 0
        if initial:
            return (1, "양호", 36) if ser < 3.5e-4 and burst < 800 else (3, "장기보존 부적합 — 즉시 대책(새 매체에 다시 기록)", 0)
        if ser >= 1e-3 or burst >= 1900:
            return 3, "즉시 대책(즉시 이전)", 0
        if ser >= 5e-4 or burst > 800:
            return 2, "신속한 대책(빠른 시일 내 이전)", 12
        return 1, "양호(계속 보존)", 36
    if "DVD" in k:
        pi = v["pi"]
        if pi is None:
            return None
        if initial:
            lim = 80 if std == "china" else 100
            return (1, "양호", 36 if std == "iso" else 24) if pi < lim else (2, f"초기 PI sum 8 {lim} 이상 — 장기보존 부적합, 새 매체에 다시 기록", 0)
        if std == "china":
            if pi > 240:
                return 3, "즉시 이전", 0
            if pi > 180:
                return 3, "경고 3등급 — 교체, 검사 주기 6개월", 6
            if pi > 140:
                return 2, "경고 2등급 — 빠른 시일 내 교체, 검사 주기 1년", 12
            return 1, "경고 1등급 — 계속 보존(주기 2년)", 24
        if pi >= 280:
            return 3, "즉시 대책(즉시 이전)", 0
        if pi >= 140:
            return 2, "신속한 대책(빠른 시일 내 이전)", 12
        return 1, "양호(계속 보존)", 36
    if "CD" in k:
        b = v["bler"]
        if b is None:
            return None
        if b >= 220:
            return 3, "BLER 220 이상 — 읽기 실패 위험, 즉시 이전", 0
        if b >= 50:
            return 2, "BLER 50~220 — 주의, 빠른 시일 내 이전 검토", 12
        return 1, "양호", 36
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("list")
    ap.add_argument("--standard", choices=["iso", "china"], default="iso", help="DVD 판정 기준: ISO/IEC 29121(기본) 또는 중국 DA/T 38-2008(가장 엄격)")
    ap.add_argument("--today")
    ap.add_argument("-o", "--out", default="산출/광디스크_오류율.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    if a.list.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if c is None else str(c).strip() for c in r] for r in load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        import csv
        rows = list(csv.reader(open(a.list, encoding="utf-8-sig")))
    head = [re.sub(r"\s", "", h).upper() for h in rows[0]]
    idx = {k: next((head.index(re.sub(r"\s", "", x).upper()) for x in al if re.sub(r"\s", "", x).upper() in head), None) for k, al in ALIAS.items()}
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "판정"
    ws.append(["원본 행", "매체번호", "종류", "검사구분", "측정일", "측정값", "등급", "판정", "다음 검사일", "비고"])
    urgent = 0
    for n, r in enumerate(rows[1:], 2):
        if not any(r):
            continue
        g = lambda k, r=r: (r[idx[k]] if idx[k] is not None and idx[k] < len(r) else "").strip()
        kind, initial = g("종류"), "초기" in g("구분")
        v = {k: num(g(k)) for k in ("pi", "ser", "burst", "bler", "uncor")}
        res = grade(kind, initial, v, a.standard)
        note = ""
        if v["uncor"]:
            res, note = (3, "수정 불가능 오류 있음 — 즉시 이전, 소실 데이터 확인", 0), "DA/T 38-2008"
        try:
            d = dt.date.fromisoformat(g("일")[:10].replace(".", "-"))
        except ValueError:
            d = None
        shown = " · ".join(f"{k} {g(k)}" for k in ("pi", "ser", "burst", "bler", "uncor") if g(k))
        if res is None:
            ws.append([n, g("번호"), kind, g("구분"), g("일"), shown, "", "판정 불가 — 종류(CD·DVD·BD)와 해당 오류율 값 확인", "", ""])
            continue
        lv, msg, months = res
        nxt = ""
        if d:
            y, mo = divmod(d.month - 1 + months, 12)
            nxt = dt.date(d.year + y, mo + 1, min(d.day, 28)).isoformat() if months else "즉시"
            if months and dt.date.fromisoformat(nxt) < today:
                note = (note + " · " if note else "") + "다음 검사일 지남"
        urgent += lv == 3
        ws.append([n, g("번호"), kind, g("구분"), g("일"), shown, lv, msg, nxt, note])
    s = wb.create_sheet("기준")
    for row in (["매체", "검사", "기준", "근거"],
                ["DVD", "초기", "PI sum 8 < 100 양호 (중국 80)", "ISO/IEC 29121 · JIS Z 6017 · DA/T 38-2008"],
                ["DVD", "정기(3년)", "< 140 양호 · 140~280 신속한 대책 · ≥ 280 즉시 대책 (중국 140/180/240, 주기 2년→1년→6개월)", "같은 표준"],
                ["BD", "초기", "R-SER < 3.5E-4 그리고 Burst < 800바이트", "JIIMA"],
                ["BD", "정기", "< 5E-4·≤ 800 양호 · 5E-4~1E-3·800~1,900 신속한 대책 · ≥ 1E-3·≥ 1,900 즉시 대책", "JIIMA"],
                ["CD", "정기", "BLER 220 읽기 한계(50·220 구간은 이 스크립트가 정함)", "연구 인용"],
                ["공통", "", "수정 불가능 오류 → 즉시 이전, 같은 장비로 추세를 볼 것", "DA/T 38-2008, 국가기록원 매체 기준 연구"]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(rows) - 1}행 · 즉시 대책 {urgent} → {a.out}")
    return 1 if urgent else 0


if __name__ == "__main__":
    sys.exit(main())
