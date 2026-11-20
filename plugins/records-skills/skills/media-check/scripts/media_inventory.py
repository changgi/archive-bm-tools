#!/usr/bin/env python3
"""기록매체 관리대장 점검: 보존·임시저장 매체 목록에서 매체 종류별 점검·되감기 주기, 쓰기 방식(-R/+R), 사용량, 보호수단,
마이그레이션 대상을 찾는다.
근거: NAK 12:2022(v3.1) 「기록매체 요건 및 관리기준」 5.4~5.7, 부속서 D·E (국가기록원, 2026-11-20 내려받아 확인)

사용:  python media_inventory.py 매체대장.xlsx [--today 2026-11-20] [--check-years 1] -o 산출/기록매체점검.xlsx
입력 열(머리글 이름으로 찾음): 매체번호, 매체종류(광디스크·DVD-R·BD-R·CD-RW·하드디스크·SSD·USB·자기테이프·LTO·비디오테이프·
       오디오테이프·플로피·LP), 용도(보존·임시·이관·백업), 수록일, 마지막점검일, 마지막되감기일, 용량, 사용량, 보호수단(WORM·쓰기방지·봉인·접근권한), 비고
규칙(NAK 12):
  광디스크: 보존용은 한 번 쓰는 -R/+R만(필수, 5.6.2.2) · 표기 용량의 4/5까지만 기록 · 주기 검사(부속서 D 또는 해시 비교)
  반도체(SSD·USB·메모리카드): 최소 1년 이하 주기로 전원 공급·동작 확인(5.7.3)
  하드디스크(임시저장): 주기적 동작 확인(5.5.3)
  자기테이프: 3년마다 되감기, 세워서 보관, 오래된 테이프는 빨리 디지털 변환·마이그레이션(5.4.3)
  플로피디스켓: 수명 1~3년 — 빨리 다른 매체로 마이그레이션(5.5.2)
  보호수단(부속서 E): WORM(매체·솔루션) 권장, 차선 접근권한·쓰기방지·봉인
원칙:  대장을 읽고 표시만 한다. 매체 검사 자체는 media_verify.py(부속서 D 간이 검사)로 한다. '점검 주기'의 기본 1년은 이 킷의 값이다(광디스크·하드디스크는 표준에 주기 숫자가 없음).
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"번호": ["매체번호", "관리번호", "매체 번호"], "종류": ["매체종류", "매체 종류", "종류", "매체"], "용도": ["용도"],
         "수록": ["수록일", "기록일", "제작일"], "점검": ["마지막점검일", "점검일", "최종점검일"], "되감기": ["마지막되감기일", "되감기일"],
         "용량": ["용량", "표기용량"], "사용": ["사용량", "기록량"], "보호": ["보호수단", "보호"], "비고": ["비고"]}


def kind_of(s):
    s = re.sub(r"\s", "", s or "").upper()
    if re.search(r"플로피|FLOPPY|디스켓", s):
        return "플로피"
    if re.search(r"SSD|USB|메모리|SD카드|반도체|플래시", s):
        return "반도체"
    if re.search(r"LTO|자기테이프|비디오테이프|오디오테이프|VHS|베타|DAT|카세트|릴|테이프", s):
        return "자기테이프"
    if re.search(r"하드|HDD|외장", s):
        return "하드디스크"
    if re.search(r"CD|DVD|BD|블루레이|광", s):
        return "광디스크"
    if re.search(r"LP|음반", s):
        return "음반"
    return ""


def day(v):
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(v or ""))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def size_gb(v):
    m = re.search(r"([\d.]+)\s*(TB|GB|MB)?", str(v or ""), re.I)
    if not m:
        return None
    n = float(m[1])
    u = (m[2] or "GB").upper()
    return n * 1024 if u == "TB" else n / 1024 if u == "MB" else n


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--today")
    ap.add_argument("--check-years", type=float, default=1.0, help="광디스크·하드디스크 점검 주기(년, 이 킷 기본 1)")
    ap.add_argument("-o", "--out", default="산출/기록매체점검.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매체 점검"
    ws.append(["매체번호", "매체종류", "분류", "용도", "수록일", "마지막 점검", "다음 점검", "사용률", "보호수단", "판정", "할 일"])
    stats = {"조치 필요": 0, "주의": 0, "통과": 0}
    for x in rows:
        k = kind_of(x.get("종류"))
        raw = (x.get("종류", "") + " " + x.get("비고", "")).upper().replace(" ", "")
        hard, soft = [], []
        last = day(x.get("점검")) or day(x.get("수록"))
        nxt = None
        period = {"반도체": 1.0, "광디스크": a.check_years, "하드디스크": a.check_years}.get(k)
        if period and last:
            nxt = last + dt.timedelta(days=round(365.25 * period))
            if today > nxt:
                (hard if k == "반도체" else soft).append(f"점검 주기 지남({nxt.isoformat()}) — 부속서 D 간이 검사 또는 해시 비교"
                                                       + (" · 반도체는 1년 이하 주기로 전원 공급(5.7.3)" if k == "반도체" else ""))
        elif period and not last:
            soft.append("점검일·수록일 없음")
        if k == "광디스크":
            if re.search(r"RW|RAM|RE\b|-RE", raw):
                hard.append("다시 쓸 수 있는 광디스크(RW·RAM·RE) — 보존용은 -R/+R만(5.6.2.2 필수)")
            if re.search(r"DL|이중층|더블", raw):
                soft.append("이중층 — 단면 단층(SS-SL) 권장")
            cap, used = size_gb(x.get("용량")), size_gb(x.get("사용"))
            if cap and used and used / cap > 0.8:
                soft.append(f"사용률 {used / cap:.0%} — 표기 용량의 4/5까지만 기록 권고")
        if k == "자기테이프":
            rw = day(x.get("되감기")) or day(x.get("수록"))
            if rw and (today - rw).days > 3 * 365:
                soft.append("3년 넘게 되감기하지 않음 — 3년마다 되감기, 세워서 보관(5.4.3)")
            rec = day(x.get("수록"))
            if rec and (today - rec).days > 10 * 365:
                soft.append("수록 10년 넘음 — 재생장치 구형화 전에 디지털 변환·마이그레이션 권고")
        if k == "플로피":
            hard.append("플로피디스켓(수명 1~3년) — 빨리 다른 매체로 마이그레이션(5.5.2)")
        if not k:
            soft.append("매체 종류를 알 수 없음")
        prot = x.get("보호", "")
        if k in ("광디스크", "하드디스크", "반도체", "자기테이프") and not prot and "보존" in x.get("용도", "보존"):
            soft.append("보호수단 없음 — WORM(매체·솔루션) 권장, 차선 접근권한·쓰기방지·봉인(부속서 E)")
        if k in ("하드디스크", "반도체"):
            soft.append("표식은 겉면에만(PCB에 붙이지 않음)") if "PCB" in raw else None
        verdict = "조치 필요" if hard else "주의" if soft else "통과"
        stats[verdict] += 1
        cap, used = size_gb(x.get("용량")), size_gb(x.get("사용"))
        ws.append([x.get("번호", ""), x.get("종류", ""), k, x.get("용도", ""), x.get("수록", ""), x.get("점검", ""), nxt.isoformat() if nxt else "",
                   f"{used / cap:.0%}" if cap and used else "", prot, verdict, " · ".join(hard + soft)])
        if hard:
            ws.cell(ws.max_row, 10).fill = red
    g = wb.create_sheet("기준(NAK 12)")
    g.append(["매체", "관리 기준"])
    for line in [("광디스크", "보존용은 -R/+R(WORM)만, 단면 단층 권장, 표기 용량의 4/5·최대 배속보다 낮게 기록, 주기 검사(부속서 D·해시)"),
                 ("하드디스크", "보존매체·임시저장·이관·백업, 임시저장이면 주기 동작 확인, 강한 자기장 주변 금지, 표식은 PCB에 붙이지 않음"),
                 ("반도체 저장매체", "정전기 제거 후 취급, 무전원 고온 장기 보관 주의, 최소 1년 이하 주기로 전원 공급·동작 확인"),
                 ("자기테이프", "사용 전 앞으로 감고 되감기, 기록 후 되감기, 세워서 보관, 3년마다 되감기, 오래된 것은 빨리 디지털 변환"),
                 ("플로피디스켓", "수명 1~3년, 빨리 마이그레이션"),
                 ("보호수단(부속서 E)", "1) 매체 WORM 2) WORM 솔루션 3) 접근권한 4) 쓰기방지 5) 봉인 — 3)~5)는 임의 변경 못 하게 추가 조치"),
                 ("간이 검사(부속서 D)", "매체 전체를 다른 저장공간에 복사 → 복사 전·후 폴더 수·파일 수·크기 비교, 복사가 지나치게 느리면 전문가 자문 — media_verify.py"),
                 ("이 킷의 값", f"광디스크·하드디스크 점검 주기 {a.check_years:g}년(표준은 '주기적'만 규정), 테이프 수록 10년 넘으면 변환 권고")]:
        g.append(list(line))
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
    for sh, widths in ((ws, (14, 16, 10, 8, 12, 12, 12, 8, 14, 10, 90)), (g, (20, 110))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"매체 {len(rows)}개 · 조치 필요 {stats['조치 필요']} · 주의 {stats['주의']} · 통과 {stats['통과']} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
