#!/usr/bin/env python3
"""기록물 기술(記述) 점검: 기술 목록(기록물군~건)이 NAK 13 기술규칙의 필수요소(참조코드·제목·일자·기술계층·생산자명·범위와 내용)를
갖췄는지, 일자 표기(연/월/일, 범위 '-', [대략]·?·[주로], 연호는 서기로), 계층별 일자 단위, 상위 계층 일자 범위, 계층 값,
전자기록물 계량단위를 지켰는지 본다.
근거: NAK 13:2022(v2.1) 「영구기록물 기술규칙」 4.1~4.4 (국가기록원, 2026-11-20 내려받아 확인)

사용:  python description_check.py 기술목록.xlsx -o 산출/기술점검.xlsx
입력 열(머리글 이름으로 찾음): 참조코드, 상위참조코드, 제목, 일자, 기술계층, 생산자명, 범위와내용, 규모(수량), 종류(전자·비전자)
원칙:  표기·누락만 본다. 내용(범위와 내용 서술)의 정확성은 기술 담당자가 판단한다. 연호를 서기로 바꾼 값은 '고쳐 쓸 값'으로만 제안한다.
"""
import argparse
import os
import re
import sys

ALIAS = {"코드": ["참조코드", "관리번호", "식별번호"], "상위": ["상위참조코드", "상위코드", "상위 참조코드"], "제목": ["제목", "표제"], "일자": ["일자", "생산일자", "생산기간"],
         "계층": ["기술계층", "계층"], "생산자": ["생산자명", "생산자", "생산기관"], "범위": ["범위와내용", "범위와 내용", "내용"], "규모": ["규모", "수량", "규모와 유형"],
         "종류": ["종류", "기록물종류"]}
REQ = [("코드", "참조코드"), ("제목", "제목"), ("일자", "일자"), ("계층", "기술계층"), ("생산자", "생산자명"), ("범위", "범위와 내용")]
LEVELS = ["기록물군", "기록물하위군", "기록물계열", "기록물하위계열", "기록물철", "기록물건"]
ERA = {"단기": -2333, "광무": 1896, "융희": 1906, "대정": 1911, "소화": 1925, "명치": 1867, "개국": 1391}
DATE_RX = re.compile(r"^(\[(대략|주로)\])?\d{3}[\d?](/\d{2}(/\d{2})?)?\??(\s*-\s*\d{3}[\d?](/\d{2}(/\d{2})?)?\??)?$")


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return idx, [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]


def years(d):
    ys = [int(y.replace("?", "0")) for y in re.findall(r"\d{3}[\d?]", d)]
    return (min(ys), max(ys)) if ys else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("-o", "--out", default="산출/기술점검.xlsx")
    a = ap.parse_args()
    idx, rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검 결과"
    ws.append(["행", "참조코드", "기술계층", "구분", "내용", "고쳐 쓸 값(제안)"])
    issues = 0
    byc = {x.get("코드", ""): x for _, x in rows if x.get("코드")}
    seen = set()
    for n, x in rows:
        out = []
        for k, label in REQ:
            if not x.get(k):
                out.append(("필수요소", f"{label} 빈 칸(4.1 필수요소)", ""))
        code = x.get("코드", "")
        if code and code in seen:
            out.append(("참조코드", "참조코드 중복", ""))
        seen.add(code)
        lv = re.sub(r"\s", "", x.get("계층", ""))
        if lv and lv not in LEVELS:
            out.append(("기술계층", f"'{lv}' — 기록물군·하위군·계열·하위계열·철·건 중 하나", ""))
        d = x.get("일자", "").strip()
        if d:
            fix = d
            for era, base in ERA.items():
                m = re.search(era + r"\s*(\d+)\s*년?", fix)
                if m:
                    fix = fix.replace(m[0], str(base + int(m[1])))
            fix = re.sub(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", lambda m_: f"{m_[1]}/{int(m_[2]):02d}/{int(m_[3]):02d}", fix)
            fix = re.sub(r"(\d{4})\s*년", r"\1", fix)
            fix = re.sub(r"(\d{4})[.-](\d{1,2})[.-](\d{1,2})", lambda m_: f"{m_[1]}/{int(m_[2]):02d}/{int(m_[3]):02d}", fix)
            fix = re.sub(r"\s*[~∼〜]\s*", "-", fix).replace("서기", "").strip()
            if not DATE_RX.match(d):
                note = "연호는 서기로 환산하고 연호는 추가설명(주기)에 적음 · " if any(e in d for e in ERA) else ""
                out.append(("일자 표기", note + "연/월/일은 '/', 범위는 '-', 추정은 [대략]·?, 집중 시기는 [주로](4.3.3)", fix if fix != d and DATE_RX.match(fix) else ""))
            if lv and lv != "기록물건" and re.search(r"\d{4}/\d{2}", d):
                out.append(("계층별 일자", f"{lv}은 생산 시작·종결 연도만(연/월/일은 건 단위)", re.sub(r"(\d{4})/\d{2}(/\d{2})?", r"\1", d)))
            par = byc.get(x.get("상위", ""))
            if par and years(d) and years(par.get("일자", "")):
                (c0, c1), (p0, p1) = years(d), years(par.get("일자", ""))
                if c0 < p0 or c1 > p1:
                    out.append(("상위 계층 일자", f"상위({par.get('코드')}) {p0}-{p1} 범위 밖({c0}-{c1})", ""))
        if x.get("상위") and x.get("상위") not in byc:
            out.append(("계층 연결", f"상위참조코드 {x.get('상위')}가 목록에 없음", ""))
        if "전자" in x.get("종류", "") and "비전자" not in x.get("종류", "") and x.get("규모") and not re.search(r"\d\s*(B|KB|MB|GB|TB|PB)\b", x.get("규모", "")):
            out.append(("규모", "전자기록물은 용량(B·KB·MB·GB·TB·PB)으로(4.3.5)", ""))
        for kind, msg, fx in out:
            ws.append([n, code, lv, kind, msg, fx])
            issues += 1
    rl = wb.create_sheet("기술요소(NAK 13)")
    rl.append(["영역", "기술요소"])
    for line in [("1 식별", "참조코드*, 제목*, 일자*, 기술계층*, 기술단위의 규모와 유형"), ("2 배경", "생산자명*, 행정연혁/개인이력, 기록물 이력, 수집·이관의 직접적 출처"),
                 ("3 내용과 구조", "범위와 내용*, 평가·폐기·처리일정, 추가수집 예상, 정리체계, 색인어"), ("4 접근과 이용", "접근환경, 이용환경, 언어, 물리적 특성·기술적 요구, 검색도구"),
                 ("5 관련자료", "원본·사본의 존재와 위치, 관련 기술단위, 출판물"), ("6 추가설명", "주기(연호 등)"), ("7 기술통제", "기술담당자, 규칙과 협약, 기술일자"),
                 ("* 필수요소", "다른 기관과 기술정보 교환을 위해 지정(4.1)"), ("일자 예", "기록물군 1960-2006 / 계열 1960-1975 / 철 1960-1960 / 건 1960/01/30, [대략]195?, [주로]1971-1979")]:
        rl.append(list(line))
    for sh, widths in ((ws, (6, 18, 14, 14, 80, 30)), (rl, (16, 100))):
        sh["A1"].font = Font(bold=True)
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"기술 {len(rows)}건 · 고칠 것 {issues} → {a.out}")
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
