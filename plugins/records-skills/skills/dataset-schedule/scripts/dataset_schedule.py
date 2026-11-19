#!/usr/bin/env python3
"""행정정보 데이터세트 관리기준표 초안: 시스템의 테이블·컬럼 정의(스키마)에서 국가기록원 관리기준표 서식의 6개 영역을 미리 채운다.

사용:  python dataset_schedule.py 스키마.csv --org "가상시" --system "가상 민원시스템" --dbms "PostgreSQL 15" [--answers 답변.json] -o 산출/데이터세트_관리기준표.xlsx
       python dataset_schedule.py 시스템.sqlite ... (SQLite 파일이면 스키마를 직접 읽는다)
스키마 csv 열: 테이블명, 테이블설명, 컬럼명, 컬럼설명, 자료형  (DB의 information_schema·테이블정의서를 내보낸 것)
서식:  국가기록원 누리집 업무안내·자료 > 기록관리 자료실 > 행정정보 데이터세트 > '관리기준표 서식(한글) 및 작성방법'(2026-11-20 받음)
       1 관리기관정보(기관명·부서명·담당자명) 2 법령정보 3 시스템정보(시스템명·개요·구축년도·산출물·EA·DBMS·업무DB명)
       4 데이터정보(대표 데이터·비정형 데이터 종류/포맷/구동SW/저장방법·암호화·연계) 5 업무정보(업무명·내용·대상, 반복)
       6 기록관리정보(단위기능명·개요·주제어·보유권한·관리권한·접근권한·정보공개 구분·보존기간·책정사유·적용 범위·기산일·처분 제약·처분 방법, 업무정보와 1:1)
자동으로 채우는 것: DBMS, 업무DB(테이블 접두어 묶음), 대표 데이터, 비정형(BLOB·파일경로 컬럼), 암호화 후보(주민번호·enc 등 컬럼),
       단위기능 후보(테이블 묶음), 적용 범위(테이블 목록), 기산일 후보(종료·완료·퇴직·등록일 등 날짜 컬럼), 주제어(테이블 설명 낱말).
원칙:  '필수' 항목 중 스키마로 알 수 없는 것(법령, 업무내용, 보존기간·책정사유, 보유권한 등)은 비워 두고 '입력 필요'로 표시한다.
       보존기간은 task-suggest 스킬이나 기관 기록관리기준표를 근거로 담당자가 정하고, 관할 영구기록물관리기관과 협의한다(시행령 제26조①2·③, 제34조의3).
"""
import argparse
import collections
import csv
import json
import os
import re
import sys

DATE_HINT = [("종료", 5), ("완료", 5), ("퇴직", 5), ("폐지", 5), ("end", 4), ("close", 4), ("cmpl", 4), ("rt", 2), ("처리", 3), ("등록", 2),
             ("reg", 2), ("create", 1), ("일자", 1), ("date", 1), ("dt", 1)]
ENC = re.compile(r"jumin|rrn|ssn|resident|주민|enc|암호|passwd|password|card|계좌|acct", re.I)
FILE = re.compile(r"blob|clob|bytea|binary|file|path|첨부|attach|image|img", re.I)


def load_schema(path):
    rows = []
    if path.lower().endswith((".sqlite", ".db", ".sqlite3")):
        import sqlite3
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
        for (t,) in con.execute("select name from sqlite_master where type='table' and name not like 'sqlite_%'"):
            for cid, name, typ, *_ in con.execute(f'pragma table_info("{t}")'):
                rows.append({"테이블명": t, "테이블설명": "", "컬럼명": name, "컬럼설명": "", "자료형": typ})
        con.close()
        return rows
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            rows.append({k: (r.get(k) or "").strip() for k in ("테이블명", "테이블설명", "컬럼명", "컬럼설명", "자료형")})
    return rows


def prefix(t):
    parts = re.split(r"[_\-]", t.lower())
    parts = [p for p in parts if p not in ("tb", "tbl", "t", "m", "d", "h", "info", "mst", "dtl", "hist")]
    return parts[0] if parts else t.lower()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("schema")
    ap.add_argument("--org", required=True)
    ap.add_argument("--system", required=True)
    ap.add_argument("--dbms", default="")
    ap.add_argument("--answers", help="담당자가 아는 값(JSON): {'부서명':..., '법령/규정':..., '구축년도':..., '단위기능': {'접두어': {'보존기간':..., '책정사유':...}}}")
    ap.add_argument("-o", "--out", default="산출/데이터세트_관리기준표.xlsx")
    a = ap.parse_args()
    cols = load_schema(a.schema)
    if not cols:
        print("입력 오류: 테이블·컬럼을 읽지 못했습니다.")
        return 2
    ans = json.load(open(a.answers, encoding="utf-8")) if a.answers else {}
    tables = collections.OrderedDict()
    for c in cols:
        t = tables.setdefault(c["테이블명"], {"desc": c["테이블설명"], "cols": []})
        t["cols"].append(c)
        if c["테이블설명"] and not t["desc"]:
            t["desc"] = c["테이블설명"]
    groups = collections.OrderedDict()
    for t in tables:
        groups.setdefault(prefix(t), []).append(t)
    unstructured = [f"{c['테이블명']}.{c['컬럼명']}({c['자료형']})" for c in cols if FILE.search(c["컬럼명"] + " " + c["자료형"] + " " + c["컬럼설명"])]
    enc = [f"{c['테이블명']}.{c['컬럼명']}" for c in cols if ENC.search(c["컬럼명"] + " " + c["컬럼설명"])]

    def start_date(tlist):
        best = []
        for t in tlist:
            for c in tables[t]["cols"]:
                if not re.search(r"date|time|dt$|일자|일시|일$", c["컬럼명"] + " " + c["자료형"] + " " + c["컬럼설명"], re.I):
                    continue
                key = (c["컬럼명"] + c["컬럼설명"]).lower()
                score = max((w for h, w in DATE_HINT if h in key), default=0)
                best.append((score, f"{t}.{c['컬럼명']}" + (f"({c['컬럼설명']})" if c["컬럼설명"] else "")))
        best.sort(reverse=True)
        return [b for _, b in best[:3]]

    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "관리기준표 초안"
    need = PatternFill("solid", fgColor="F9EEDA")
    ws.append(["기관명", a.org, "시스템명", a.system])
    ws.append(["영역", "항목", "내용(초안)", "구분", "출처"])
    for c in ws[2]:
        c.font = Font(bold=True)

    def row(area, item, val, must, src):
        ws.append([area, item, val, must, src if val else "입력 필요"])
        if not val and must.startswith("필수"):
            for c in ws[ws.max_row]:
                c.fill = need

    row("1. 관리기관정보", "기관명", a.org, "필수", "인자")
    row("1. 관리기관정보", "부서명", ans.get("부서명", ""), "필수", "답변")
    row("1. 관리기관정보", "담당자명", ans.get("담당자명", ""), "선택", "답변")
    row("2. 법령정보", "법령/규정", ans.get("법령/규정", ""), "필수", "답변")
    row("3. 시스템정보", "시스템명", a.system, "필수", "인자")
    row("3. 시스템정보", "시스템 개요", ans.get("시스템 개요", ""), "선택", "답변")
    row("3. 시스템정보", "구축년도", ans.get("구축년도", ""), "필수", "답변")
    row("3. 시스템정보", "개발 산출물", ans.get("개발 산출물", "테이블 정의(이 초안의 원천 스키마)"), "선택", "스키마")
    row("3. 시스템정보", "EA 정보", ans.get("EA 정보", ""), "선택", "답변")
    row("3. 시스템정보", "DBMS 정보", a.dbms, "필수", "인자")
    row("3. 시스템정보", "업무 DB명", ", ".join(groups), "필수", "테이블 접두어 묶음")
    row("4. 데이터정보", "대표 데이터", ", ".join(f"{t}({tables[t]['desc']})" if tables[t]["desc"] else t for t in list(tables)[:8]), "필수", "스키마")
    row("4. 데이터정보", "비정형 데이터 종류·저장방법", "; ".join(unstructured[:10]), "필수(해당 시)", "BLOB·파일경로 컬럼")
    row("4. 데이터정보", "비정형 데이터 포맷·구동 SW", ans.get("비정형 포맷", ""), "필수(해당 시)", "답변")
    row("4. 데이터정보", "암호화 데이터 사용유무", ("암호화 후보 컬럼: " + ", ".join(enc[:10])) if enc else "", "선택", "컬럼명 패턴(확인 필요)")
    row("4. 데이터정보", "연계 시스템 유무", ans.get("연계 시스템", ""), "선택", "답변")
    unit = ans.get("단위기능", {})
    for g, tl in groups.items():
        u = unit.get(g, {})
        words = sorted({w for t in tl for w in re.findall(r"[가-힣]{2,}", tables[t]["desc"])})[:8]
        area5, area6 = f"5. 업무정보 [{g}]", f"6. 기록관리정보 [{g}]"
        row(area5, "업무명", u.get("업무명", g), "필수", "접두어(이름 확인)")
        row(area5, "업무내용", u.get("업무내용", ""), "필수", "답변")
        row(area5, "대상", u.get("대상", ""), "선택", "답변")
        row(area6, "단위기능명", u.get("업무명", g), "필수", "업무정보와 1:1")
        row(area6, "단위기능개요", "; ".join(f"{t}: {tables[t]['desc']}" for t in tl if tables[t]["desc"]), "필수", "테이블 설명")
        row(area6, "주제어", ", ".join(words), "필수", "테이블 설명 낱말")
        row(area6, "데이터의 보유권한", u.get("보유권한", ""), "필수", "답변")
        row(area6, "시스템의 관리권한", u.get("관리권한", ""), "필수", "답변")
        row(area6, "접근권한", u.get("접근권한", ""), "선택", "답변")
        row(area6, "정보공개 구분", u.get("정보공개 구분", ""), "필수", "답변")
        row(area6, "보존기간", u.get("보존기간", ""), "필수", "답변(task-suggest·기준표 근거)")
        row(area6, "보존기간 책정사유", u.get("책정사유", ""), "필수", "답변")
        row(area6, "적용 범위", ", ".join(tl), "필수", "테이블 목록")
        cand = start_date(tl)
        row(area6, "기산일", u.get("기산일", ("후보: " + " / ".join(c + " 값이 속하는 다음 연도 1월 1일" for c in cand)) if cand else ""), "필수", "날짜 컬럼 후보(확인 필요)")
        row(area6, "처분의 제약 발생 사항", u.get("처분 제약", ""), "선택", "답변")
        row(area6, "처분 방법", u.get("처분 방법", ""), "선택", "답변")
    for i, w in enumerate((24, 22, 80, 12, 24)):
        ws.column_dimensions[chr(65 + i)].width = w
    s2 = wb.create_sheet("스키마")
    s2.append(["테이블명", "테이블설명", "컬럼명", "컬럼설명", "자료형"])
    for c in cols:
        s2.append([c[k] for k in ("테이블명", "테이블설명", "컬럼명", "컬럼설명", "자료형")])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    blanks = sum(1 for r in ws.iter_rows(min_row=3, values_only=True) if not r[2] and str(r[3]).startswith("필수"))
    print(f"테이블 {len(tables)} · 단위기능 후보 {len(groups)} · 비정형 {len(unstructured)} · 암호화 후보 {len(enc)} · 필수 입력 필요 {blanks} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
