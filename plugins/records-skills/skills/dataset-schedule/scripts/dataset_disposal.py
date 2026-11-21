#!/usr/bin/env python3
"""행정정보 데이터세트 평가·폐기 준비 — 관리기준표의 단위기능별 보존기간·기산일·처분제약으로 보존기간이 끝난 데이터 영역을 찾아
평가심의 상정 목록(영역별)과 폐기 결과 기록부를 만들고, 운영부서가 검증할 건수 확인 SQL을 적는다.

근거: 국가기록원 「행정정보데이터세트 기록관리 실행매뉴얼(지침종합)」(2020. 9.)
  3.3 평가 — 처분제약에 따른 보류기간과 보존기간이 모두 지난 뒤 평가심의에 상정, 데이터의 영역별(볼륨·범위·건수)로 상정
           영역별 삭제는 운영부서가 기술 검토로 전체 시스템 영향이 없음을 사전 검증
  3.4 삭제 — 폐기가 확정되면 삭제하고 결과를 기록화: 단위기능·범위·폐기량·폐기방식·폐기일시. 시스템 안은 시스템에서, 별도 스토리지·매체는 각각 파기
           삭제 SQL은 기관이 시스템에 맞게 자체 작성하고 유지보수와 충분히 검토(삭제지원 도구는 국가기록원 제공)
  1.2 관리대상 선정·관리기준표·평가심의 없이 운영부서가 임의로 삭제(백업 후 삭제 포함)하면 '기록물 무단 폐기'에 해당할 수 있음
  공공기록물법 시행령 제25조⑥·제26조·제43조(평가), NAK 35:2020

입력: 단위기능 목록(xlsx·csv) — 열: 단위기능(명), 테이블(쉼표로 여럿), 기산일컬럼, 보존기간, 처분제약(년, 소송·감사 등 보류), 연도별건수(예 '2014:1200;2015:980', 있으면)
결과: xlsx — 평가 상정 목록 · 평가심의서(영역별) · 폐기 결과 기록부(빈 칸) · 확인용 SQL(SELECT COUNT, DELETE는 주석) · 기준
이 스크립트는 DB에 접속하지 않고, 아무것도 지우지 않는다.

예) python dataset_disposal.py 단위기능.xlsx --system "가상민원시스템" --today 2026-11-21 -o 산출/데이터세트_평가.xlsx
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

ALIAS = {"기능": ["단위기능", "단위기능명", "업무기능"], "테이블": ["테이블", "테이블명", "대상테이블"], "컬럼": ["기산일컬럼", "기산일 컬럼", "기산일"],
         "보존": ["보존기간"], "제약": ["처분제약", "처분제약(년)", "보류기간"], "건수": ["연도별건수", "연도별 건수", "건수"]}
YEARS = {"1년": 1, "3년": 3, "5년": 5, "10년": 10, "30년": 30, "준영구": None, "영구": None}


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
    ap = argparse.ArgumentParser(description="행정정보 데이터세트 평가·폐기 준비(실행매뉴얼 2020 3.3·3.4)")
    ap.add_argument("listing")
    ap.add_argument("--system", default="")
    ap.add_argument("--today")
    ap.add_argument("-o", "--out", default="산출/데이터세트_평가.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    items = load(a.listing)
    if not items or "기능" not in items[0]:
        print("입력 오류: '단위기능' 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    fill = PatternFill("solid", fgColor="FFF4CC")
    wb = openpyxl.Workbook()
    ls = wb.active
    ls.title = "평가 상정 목록"
    ls.append(["단위기능", "테이블", "보존기간", "처분제약(년)", "평가 상정 가능한 기산연도(이하)", "대상 건수(연도별 합)", "확인할 것"])
    ev = wb.create_sheet("평가심의서(영역별)")
    ev.append(["연번", "시스템", "단위기능", "데이터 영역(범위)", "건수·볼륨", "보존기간", "기산일 기준", "처분제약", "처리과(운영부서) 의견(채움)",
               "시스템 영향도 검토(채움)", "기록물관리전문요원 의견(채움)", "평가심의회 결과(채움)"])
    sq = wb.create_sheet("확인용 SQL")
    sq.append(["단위기능", "테이블", "SQL(운영부서가 영향도 검증 후 사용 — 이 도구는 실행하지 않음)"])
    n_ev = 0
    for x in items:
        fn, tbls = x.get("기능", ""), [t.strip() for t in re.split(r"[,;/]", x.get("테이블", "")) if t.strip()]
        per = next((p for p in YEARS if p in x.get("보존", "")), "")
        hold = int(re.sub(r"\D", "", x.get("제약", "")) or 0)
        col = x.get("컬럼", "")
        notes = []
        if not per:
            notes.append(f"보존기간 '{x.get('보존', '')}' — 7종 중 하나로(관리기준표)")
        if not col:
            notes.append("기산일 컬럼 없음 — 업무가 끝나는 시점의 날짜 컬럼을 관리기준표에 지정")
        last = None
        if per and YEARS[per]:
            # 기산일(다음 해 1. 1.) + 보존기간 + 처분제약 이 지난 연도: 기산연도 Y의 만료 = 1. 1.(Y+1+기간+제약)
            last = today.year - 1 - YEARS[per] - hold
        elif per:
            notes.append(f"{per} — 평가·폐기 대상 아님(이관 협의 대상이면 SIARD 패키지: siard_check.py)")
        counts = {}
        for part in re.split(r"[;\n]", x.get("건수", "")):
            m = re.match(r"\s*(\d{4})\s*[:=]\s*([\d,]+)", part)
            if m:
                counts[int(m[1])] = int(m[2].replace(",", ""))
        due = sum(v for y, v in counts.items() if last is not None and y <= last)
        ls.append([fn, ", ".join(tbls), per, hold, last if last is not None else "", due if counts else "", " · ".join(notes) or "통과"])
        if last is not None and (not counts or due):
            n_ev += 1
            rng = f"{col} < {last + 1}-01-01" if col else f"기산연도 {last}년 이하"
            ev.append([n_ev, a.system, fn, rng, f"{due:,}건" if counts else "(건수 확인)", per, col, f"{hold}년" if hold else "없음", "", "", "", ""])
            for t in tbls or ["(테이블)"]:
                if col:
                    sq.append([fn, t, f"SELECT COUNT(*) FROM {t} WHERE {col} < DATE '{last + 1}-01-01';"])
                    sq.append([fn, t, f"-- (평가심의 폐기 확정·영향도 검증 후) DELETE FROM {t} WHERE {col} < DATE '{last + 1}-01-01';  -- 참조 무결성·첨부 파일 경로 먼저 확인"])
    rec = wb.create_sheet("폐기 결과 기록부")
    rec.append(["연번", "시스템", "단위기능", "폐기 범위", "폐기량(건·볼륨)", "폐기 방식(시스템 기능·삭제지원 도구·매체 파기)", "폐기 일시", "근거(평가심의 결과·문서번호)",
                "실행자(운영부서)", "확인자(기록관)", "별도 스토리지·백업 매체 처리"])
    gd = wb.create_sheet("기준")
    for row in [("평가 시점", "처분제약(소송·감사 등) 보류기간과 보존기간이 모두 지난 뒤 — 기산연도 Y의 만료는 Y+1년 1월 1일 + 보존기간 + 처분제약"),
                ("상정 단위", "데이터의 영역별(볼륨·범위·건수). 영역 삭제가 시스템 전체에 영향이 없는지 운영부서가 먼저 기술 검토"),
                ("폐기 기록화", "단위기능·범위·폐기량·폐기방식·폐기일시를 남김. 시스템 안은 시스템에서, 별도 스토리지·백업 매체도 각각 파기"),
                ("무단 폐기", "관리대상 선정·관리기준표·평가심의 없이 운영부서가 임의로 삭제(백업 후 삭제 포함)하면 기록물 무단 폐기에 해당할 수 있음"),
                ("SQL", "확인용 SELECT만 제안. DELETE는 주석으로만 두며 기관이 시스템에 맞게 작성·검증(국가기록원 삭제지원 도구 사용 가능)"),
                ("근거", "행정정보데이터세트 기록관리 실행매뉴얼(지침종합, 2020. 9.) 3.3·3.4·4.1, 시행령 제25조⑥·제43조, NAK 35:2020")]:
        gd.append(row)
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True)
    for r in ev.iter_rows(min_row=2, min_col=9, max_col=12):
        for c in r:
            c.fill = fill
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"단위기능 {len(items)} · 평가 상정 영역 {n_ev} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
