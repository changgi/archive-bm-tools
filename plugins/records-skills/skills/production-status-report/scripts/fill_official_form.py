#!/usr/bin/env python3
"""국가기록원 「기록물 생산현황 통보서식」(xlsx)에 등록정보를 채운다. 원본 서식은 고치지 않고 복사본에 쓴다.

사용:  python fill_official_form.py 등록정보.xlsx --form 2024년도기록물생산현황통보서식.xlsx --year 2025 -o 산출/생산현황_통보서식_2025.xlsx
       [--list-scope long|all|none]  서식 2-1·2-2 목록에 넣을 범위(기본 long: 보존기간 30년 이상만)
서식:  국가기록원 누리집 업무안내·자료 > 생산부터 보존까지 > 처리과 > '기록물 생산현황 통보서식(2024년)'
       서식 1  일반기록물(전자·비전자) 생산현황 — 일반문서 ①전자 ②비전자 ③혼합, 도면류 ④, 카드류 ⑤ × 보존기간 7종 × 철·건
       서식 2-1 생산목록(기록물철), 서식 2-2 생산목록(기록물건)
       주의사항(서식에 적힌 것): 모든 서식 변경·시트 삭제 금지. 중앙행정기관·특별행정기관은 보존기간 30년 이상 단위과제카드(철)·건 목록 제출,
       지자체·교육청·기타공공기관은 [서식 2] 제출 제외. 시스템 연계가 되는 기관은 RMS→CAMS 통보.
등록정보 열(머리글 이름으로 찾음): 철분류번호, 철제목, 등록번호, 건제목(제목), 생산부서, 처리과코드, 단위과제ID, 단위과제명, 생산연도,
       보존기간, 기록물형태(전자·비전자), 유형(일반문서·도면·카드), 권호수, 생산시스템구분코드, 등록일시, 공개구분, 비공개사유, 비밀구분
추가 입력(선택):
  --av 시청각목록.xlsx   → 서식 3(사진/필름류·녹음/동영상류 × 보존기간 건수)·서식 4(생산목록)
        열: 생산부서, 등록일시, 건제목, 유형구분(사진·필름 / 녹음·동영상), 매체형태, 건수, 수량, 등록번호, 내용요약, 공개구분, 비공개사유, 단위과제명, 보존기간, 비고
  --meetings 회의목록.xlsx → 서식 8(회의 종류 1~9 × 개최횟수 대면·영상/서면 × 회의록·속기록 생산건수)
        열: 회의명, 회의종류(1~9 번호, 시행령 제18조①의 호), 개최방식(대면·영상·서면), 회의록건수, 속기록건수
        회의종류 번호가 없으면 회의체 목록의 회의호수, 그것도 없으면 9(그 밖에)로 두고 '확인'으로 알린다.
        회의 목록에 '회의체' 열이 있으면 그것으로, 없으면 회의명에 들어 있는 가장 긴 회의체명으로 묶는다.
  --bodies 회의체목록.xlsx → 서식 9(회의록 생산목록, 회의체마다 한 줄 — 전년도 개최가 없어도 0으로 넣음). --meetings와 함께 쓴다.
        열: 회의체명, 속기록지정(지정회의/공란), 부서명, 회의호수(1~9), 근거구분(법률·대통령령·총리령·부령·조례·규칙·훈령·예규·고시),
            근거명칭(정식명칭), 근거조항(제○조 제○항), 최초구성일(YYYY.MM.DD), 회의기능, 위원장(직위·민간여부), 위원(구성 현황)
        점검: 근거구분·근거조항·구성일 형식, 회의호수, 서면 개최인데 회의록 0건, 속기록 지정회의인데 속기록·녹음 건수 < 개최횟수,
              회의체 목록에 없는 회의, 위원장·위원 칸에 사람 이름처럼 보이는 것(성명은 쓰지 않음).
  --artifacts 행정박물_점검.xlsx → 서식 5(유형별 점 수)·서식 6(보유목록). admin-artifacts-check 결과 파일을 넣는다.
  --pubs 간행물_점검.xlsx → 서식 7(생산량·송부량, 종 수). publications-check 결과 파일을 넣는다.
  --secret 비밀_점검.xlsx → 서식 10(생산·재분류 통계)·서식 11(생산 후 3년 지난 목록). classified-records-check 결과를 넣는다.
        비밀 전용 장비에서만 실행한다.
원칙:  서식의 구조·수식·병합은 그대로 두고 값 칸만 채운다. 등록정보의 비밀 표시 행은 서식 1·2에 넣지 않는다(서식 10·11은 --secret). 제출은 사람이 공문·시스템으로 한다.
"""
import argparse
import collections
import os
import re
import shutil
import sys

PCOL = {"영구": "E", "준영구": "G", "30년": "I", "10년": "K", "5년": "M", "3년": "O", "1년": "Q"}
LONG = ("영구", "준영구", "30년")
ALIAS = {"철": ["철분류번호", "기록물철분류번호", "분류번호", "철번호"], "철제목": ["철제목", "기록물철제목", "기록물철명"],
         "등록번호": ["등록번호", "생산등록번호", "생산(접수)등록번호"], "건제목": ["건제목", "제목", "기록물건제목"],
         "부서": ["생산부서", "처리과", "생산부서(처리과)명", "처리과명", "생산기관"], "처리과코드": ["처리과코드"], "과제ID": ["단위과제ID", "단위업무코드", "단위과제코드", "단위업무코드/단위과제ID"],
         "과제명": ["단위과제명", "단위업무명", "단위과제", "단위업무명/단위과제명"], "연도": ["생산연도", "생산년도", "생산시작년도"], "보존기간": ["보존기간"],
         "형태": ["기록물형태", "전자구분", "전자기록물여부"], "유형": ["유형", "기록물유형", "기록물구분"], "권호": ["권호수"],
         "시스템": ["생산시스템구분코드", "생산시스템구분 코드", "시스템구분코드"], "일시": ["등록일시", "생산(접수)등록일시", "생산일자"],
         "공개": ["공개구분"], "사유": ["비공개사유"], "비밀": ["비밀구분", "비밀여부"], "건식별자": ["건식별자"]}


def norm_p(v):
    s = re.sub(r"\s", "", str(v or ""))
    m = re.fullmatch(r"(\d+)년?", s)
    return f"{m[1]}년" if m else s


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head:
                idx[k] = head.index(n)
                break
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)], idx


SYS_CODES = {"01": "업무관리시스템('08년 이전)", "02": "전자문서시스템", "03": "기록관리시스템", "04": "업무관리시스템('08년 이후)",
             "05": "기타(행정 DB구축, 자료관시스템 마이그레이션)", "06": "구전자문서시스템"}  # 서식 2-1 기재요령(NAK 29-2:2020 시스템구분 코드)
ABBR = {"공공기록물법": "공공기록물 관리에 관한 법률", "기록물관리법": "공공기록물 관리에 관한 법률", "정보공개법": "공공기관의 정보공개에 관한 법률",
        "공공데이터법": "공공데이터의 제공 및 이용 활성화에 관한 법률", "교육자치법": "지방교육자치에 관한 법률", "지방교육자치법": "지방교육자치에 관한 법률",
        "민원처리법": "민원 처리에 관한 법률", "개인정보법": "개인정보 보호법", "청탁금지법": "부정청탁 및 금품등 수수의 금지에 관한 법률",
        "김영란법": "부정청탁 및 금품등 수수의 금지에 관한 법률", "부패방지권익위법": "부패방지 및 국민권익위원회의 설치와 운영에 관한 법률",
        "이해충돌방지법": "공직자의 이해충돌 방지법", "정부업무평가법": "정부업무평가 기본법", "행정기본법시행령": "행정기본법 시행령",
        "대통령기록물법": "대통령기록물 관리에 관한 법률", "국가재정법시행령": "국가재정법 시행령"}
PCOL3 = {"영구": "D", "준영구": "E", "30년": "F", "10년": "G", "5년": "H", "3년": "I", "1년": "J"}


def read_generic(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10]) if sum(1 for c in r if c) >= 3), 0)
    head = rows[hi]
    return [dict(zip(head, r)) for r in rows[hi + 1:] if any(r)], head


def insert_rows_keep_merges(ws, at, n):
    """at 행부터 n줄을 끼워 넣고, 그 아래 병합 범위도 같이 내린다(openpyxl은 병합을 옮기지 않는다)."""
    moved = [m for m in list(ws.merged_cells.ranges) if m.min_row >= at]
    for m in moved:
        ws.merged_cells.ranges.remove(m) if hasattr(ws.merged_cells.ranges, "remove") else ws.unmerge_cells(str(m))
    ws.insert_rows(at, n)
    for m in moved:
        ws.merge_cells(start_row=m.min_row + n, start_column=m.min_col, end_row=m.max_row + n, end_column=m.max_col)


def write_rows(ws, first, last_free, rows):
    free = last_free - first + 1
    if len(rows) > free:
        insert_rows_keep_merges(ws, last_free + 1, len(rows) - free)
    for i, vals in enumerate(rows):
        for j, v in enumerate(vals):
            ws.cell(row=first + i, column=1 + j, value=v)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--form", required=True, help="국가기록원 생산현황 통보서식 원본 xlsx")
    ap.add_argument("--year", type=int, required=True, help="생산연도")
    ap.add_argument("--list-scope", choices=["long", "all", "none"], default="long")
    ap.add_argument("--av", help="시청각기록물 목록(서식 3·4)")
    ap.add_argument("--meetings", help="회의 목록(서식 8)")
    ap.add_argument("--bodies", help="회의체 목록(서식 9), --meetings와 함께")
    ap.add_argument("--artifacts", help="admin-artifacts-check 결과(서식 5·6)")
    ap.add_argument("--pubs", help="publications-check 결과(서식 7)")
    ap.add_argument("--secret", help="classified-records-check 결과(서식 10·11), 비밀 전용 장비에서만")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    if os.path.abspath(a.out) == os.path.abspath(a.form) or os.path.exists(a.out):
        print("덮어쓰지 않음: 결과 파일 이름을 원본·기존 파일과 다르게 주세요.")
        return 2
    rows, idx = read(a.listing)
    if "보존기간" not in idx or "철" not in idx:
        print("입력 오류: 철분류번호와 보존기간 열이 필요합니다.")
        return 2
    secret = [r for r in rows if r.get("비밀", "") not in ("", "일반", "N", "아니오", "해당없음")]
    rows = [r for r in rows if r not in secret and (not r.get("연도") or str(r.get("연도"))[:4] == str(a.year))]

    import openpyxl
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    shutil.copyfile(a.form, a.out)
    wb = openpyxl.load_workbook(a.out)
    if "서식 1" not in wb.sheetnames:
        print("서식 오류: '서식 1' 시트가 없습니다. 국가기록원 통보서식 원본인지 확인하세요.")
        return 2

    files = collections.defaultdict(list)
    for r in rows:
        files[r["철"]].append(r)
    bad_sys = sorted({r.get("시스템", "") for r in rows if r.get("시스템", "") and r.get("시스템", "").zfill(2) not in SYS_CODES})
    if bad_sys:
        print("확인: 생산시스템구분 코드가 서식 기재요령(01~06)에 없음 — " + ", ".join(bad_sys[:10]))
    counts = collections.Counter()
    skipped = 0
    for no, items in files.items():
        p = norm_p(items[0].get("보존기간"))
        if p not in PCOL:
            skipped += 1
            continue
        kind = " ".join(x.get("유형", "") for x in items)
        forms = {("비전자" if "비전자" in x.get("형태", "") or x.get("형태", "") in ("N", "종이") else "전자") for x in items}
        if "도면" in kind:
            row = 11
        elif "카드" in kind:
            row = 12
        else:
            row = 8 if forms == {"전자"} else (9 if forms == {"비전자"} else 10)
        counts[(row, p, "철")] += 1
        counts[(row, p, "건")] += len(items)
    ws = wb["서식 1"]
    ws["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
    for row in (8, 9, 10, 11, 12):
        for p, col in PCOL.items():
            ws[f"{col}{row}"] = counts.get((row, p, "철"), 0)
            ws[f"{chr(ord(col) + 1)}{row}"] = counts.get((row, p, "건"), 0)

    n1 = n2 = 0
    if a.list_scope != "none":
        pick = [no for no, items in files.items() if a.list_scope == "all" or norm_p(items[0].get("보존기간")) in LONG]
        if "서식 2-1" in wb.sheetnames:
            w = wb["서식 2-1"]
            w["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
            body = []
            for i, no in enumerate(sorted(pick), 1):
                f = files[no][0]
                body.append([i, f.get("시스템", ""), f.get("처리과코드", ""), f.get("부서", ""), f.get("연도", "") or a.year, f.get("과제ID", ""),
                             f.get("과제명", ""), no, f.get("권호", ""), f.get("철제목", ""), f.get("유형", ""), norm_p(f.get("보존기간")), len(files[no])])
            write_rows(w, 6, 13, body)
            n1 = len(body)
        if "서식2-2" in wb.sheetnames:
            w = wb["서식2-2"]
            w["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
            body = []
            k = 0
            for no in sorted(pick):
                for x in files[no]:
                    k += 1
                    body.append([k, x.get("시스템", ""), x.get("처리과코드", ""), x.get("과제명", "") or x.get("과제ID", ""), x.get("철제목", ""),
                                 x.get("유형", ""), norm_p(x.get("보존기간")), x.get("건식별자", ""), x.get("일시", ""), x.get("등록번호", ""),
                                 x.get("건제목", ""), "N" if "비전자" in x.get("형태", "") else "Y", no, x.get("공개", ""), x.get("사유", "")])
            write_rows(w, 6, 11, body)
            n2 = len(body)
    av_n = mt_n = 0
    notes = []
    if a.av and "서식 3" in wb.sheetnames:
        avs, _ = read_generic(a.av)
        w3 = wb["서식 3"]
        w3["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
        c3 = collections.Counter()
        for x in avs:
            kind = 6 if re.search(r"녹음|동영상|영상|비디오|오디오", x.get("유형구분", "")) else 5
            p = norm_p(x.get("보존기간"))
            if p in PCOL3:
                c3[(kind, p)] += int(x.get("건수") or 1) if str(x.get("건수") or "1").isdigit() else 1
            else:
                notes.append(f"시청각 '{x.get('건제목', '')[:20]}' 보존기간 불명")
        for kind in (5, 6):
            for p, col in PCOL3.items():
                w3[f"{col}{kind}"] = c3.get((kind, p), 0)
        if "서식 4" in wb.sheetnames:
            w4 = wb["서식 4"]
            w4["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
            body = [[i, x.get("생산부서", ""), x.get("등록일시", ""), x.get("건제목", ""), x.get("유형구분", ""), x.get("매체형태", ""),
                     x.get("건수", ""), x.get("수량", ""), x.get("등록번호", ""), x.get("내용요약", ""), x.get("공개구분", ""),
                     x.get("비공개사유", ""), x.get("단위과제명", ""), norm_p(x.get("보존기간")), x.get("비고", "")] for i, x in enumerate(avs, 1)]
            if body:
                insert_rows_keep_merges(w4, 5, len(body))
                write_rows(w4, 5, 4 + len(body), body)
        av_n = len(avs)
    bodies = []
    if a.bodies:
        bodies, _ = read_generic(a.bodies)
    bname = lambda b: re.sub(r"\s", "", b.get("회의체명", ""))
    def body_of(x):
        if x.get("회의체"):
            return next((b for b in bodies if bname(b) == re.sub(r"\s", "", x["회의체"])), None)
        nm = re.sub(r"\s", "", x.get("회의명", ""))
        hits = [b for b in bodies if bname(b) and bname(b) in nm]
        return max(hits, key=lambda b: len(bname(b))) if hits else None
    b9_n = None
    if a.meetings and "서식 8" in wb.sheetnames:
        ms, _ = read_generic(a.meetings)
        for x in ms:
            b = body_of(x) if bodies else None
            x["_body"] = b
            if b is not None and not re.match(r"\s*[1-9]", str(x.get("회의종류", ""))) and re.match(r"\s*[1-9]", str(b.get("회의호수", ""))):
                x["회의종류"] = str(b["회의호수"]).strip()
        w8 = wb["서식 8"]
        w8["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
        c8 = collections.Counter()
        for x in ms:
            m = re.match(r"\s*([1-9])", str(x.get("회의종류", "")))
            k = int(m[1]) if m else 9
            if not m:
                notes.append(f"회의 '{x.get('회의명', '')[:20]}' 종류 번호 없음 → 9(그 밖에)로 둠, 확인")
            row = 5 + k
            mode = "서면" if "서면" in x.get("개최방식", "") else "대면"
            c8[(row, mode)] += 1
            num = lambda v: int(v) if str(v or "").strip().isdigit() else (1 if str(v or "").strip().upper() in ("Y", "O", "예") else 0)
            c8[(row, "회의록")] += num(x.get("회의록건수"))
            c8[(row, "속기록")] += num(x.get("속기록건수"))
        for row in range(6, 15):
            w8[f"C{row}"] = c8.get((row, "대면"), 0)
            w8[f"D{row}"] = c8.get((row, "서면"), 0)
            w8[f"F{row}"] = c8.get((row, "회의록"), 0)
            w8[f"G{row}"] = c8.get((row, "속기록"), 0)
        mt_n = len(ms)
        if bodies and "서식 9" in wb.sheetnames:
            num = lambda v: int(v) if str(v or "").strip().isdigit() else (1 if str(v or "").strip().upper() in ("Y", "O", "예") else 0)
            agg = {id(b): collections.Counter() for b in bodies}
            for x in ms:
                b = x.get("_body")
                if b is None:
                    notes.append(f"회의 '{x.get('회의명', '')[:24]}' — 회의체 목록에 없음(서식 9에 빠짐)")
                    continue
                c = agg[id(b)]
                c["서면" if "서면" in x.get("개최방식", "") else "대면"] += 1
                c["회의록"] += num(x.get("회의록건수"))
                c["속기록"] += num(x.get("속기록건수"))
                if "서면" in x.get("개최방식", "") and not num(x.get("회의록건수")):
                    c["서면무록"] += 1
            KINDS9 = ("법률", "대통령령", "총리령", "부령", "조례", "규칙", "훈령", "예규", "고시")
            body9 = []
            jj = {}
            jp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "jijeong_2026.csv")
            if os.path.exists(jp):
                import csv as _csv
                with open(jp, encoding="utf-8-sig") as fh:
                    for r in _csv.DictReader(fh):
                        jj[re.sub(r"\s|\(.*?\)", "", r["대상회의"])] = r
            for b in bodies:
                kn = re.sub(r"\s|\(.*?\)", "", b.get("회의체명", ""))
                hitj = jj.get(kn) or next((v for k, v in jj.items() if len(kn) >= 6 and (k.endswith(kn) or kn.endswith(k))), None)
                if hitj and "지정" not in b.get("속기록지정", ""):
                    b["속기록지정"] = "지정회의"
                    notes.append(f"서식 9 '{b.get('회의체명', '')[:20]}': 국가기록원 '속기록 생산의무 지정회의 현황'('26. 4. 1., {hitj['소관기관']})에 있음 — 속기록지정을 '지정회의'로 채움, 확인")
                c, nm = agg[id(b)], b.get("회의체명", "")
                chk = []
                if b.get("근거구분") and not any(k in b["근거구분"] for k in KINDS9):
                    chk.append(f"근거구분 '{b['근거구분']}'")
                if not b.get("근거명칭"):
                    chk.append("근거명칭 없음")
                elif re.sub(r"[「」\s]", "", b["근거명칭"]) in ABBR:
                    short = b["근거명칭"]
                    full = ABBR[re.sub(r"[「」\s]", "", short)]
                    chk.append(f"근거명칭 '{short}'은 약칭 — 「{full}」")
                if b.get("근거조항") and not re.search(r"제\s*\d+\s*조", b["근거조항"]):
                    chk.append(f"근거조항 '{b['근거조항']}' — 제○조 제○항")
                if b.get("최초구성일") and not re.fullmatch(r"\d{4}\.\d{2}\.\d{2}\.?", b["최초구성일"].strip()):
                    chk.append(f"최초구성일 '{b['최초구성일']}' — YYYY.MM.DD")
                if not re.match(r"\s*[1-9]\s*$", str(b.get("회의호수", ""))):
                    chk.append("회의호수(시행령 제18조① 1~9호) 없음")
                if c["서면무록"]:
                    chk.append(f"서면 개최 {c['서면무록']}회에 회의록 0건 — 서면이어도 회의록 생산")
                total = c["대면"] + c["서면"]
                if "지정" in b.get("속기록지정", "") and c["속기록"] < total:
                    chk.append(f"지정회의인데 속기록·녹음 {c['속기록']}건 < 개최 {total}회(시행령 제18조②, 녹음은 녹취록도)")
                if "지정" in b.get("속기록지정", "") and str(b.get("회의호수", "")).strip() in ("4", "7", "8", "9"):
                    chk.append(f"지정회의는 제18조① 1·2·3·5·6호 회의 중에서만 — 회의호수 {b['회의호수']} 확인")
                for k in ("위원장", "위원"):
                    if re.search(r"(^|[\s,(])[가-힣]{2,3}\s*(위원장|교수|위원|씨)?\s*($|[,)])", b.get(k, "")) and not re.search(r"장관|차관|국장|과장|실장|교수|민간|공무원|명|직위|시장|청장|원장", b.get(k, "")):
                        chk.append(f"{k} 칸에 성명이 있을 수 있음(직위·민간여부만)")
                if chk:
                    notes.append(f"서식 9 '{nm[:20]}': " + " · ".join(chk))
                body9.append([nm, "지정회의" if "지정" in b.get("속기록지정", "") else "", b.get("부서명", ""), b.get("회의호수", ""), b.get("근거구분", ""),
                              b.get("근거명칭", ""), b.get("근거조항", ""), b.get("최초구성일", ""), b.get("회의기능", ""), b.get("위원장", ""), b.get("위원", ""),
                              None, c["대면"], c["서면"], None, c["회의록"], c["속기록"]])
            w9 = wb["서식 9"]
            w9["N2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
            if len(body9) > 6:
                insert_rows_keep_merges(w9, 11, len(body9) - 6)
            for i, v in enumerate(body9):
                r = 5 + i
                for j, val in enumerate(v):
                    if val is not None:
                        w9.cell(row=r, column=1 + j, value=val)
                w9[f"L{r}"] = f"=M{r}+N{r}"
                w9[f"O{r}"] = f"=P{r}+Q{r}"
            if not body9:
                w9["A5"] = "해당없음"
            b9_n = len(body9)
    art_n = None
    if a.artifacts and "서식 5" in wb.sheetnames:
        src = openpyxl.load_workbook(a.artifacts, data_only=True)
        if "서식 5 수치" not in src.sheetnames or "서식 6 행" not in src.sheetnames:
            print("입력 오류: --artifacts에는 admin-artifacts-check 결과 파일을 넣으세요.")
            return 2
        w5 = wb["서식 5"]
        hdr = str(src["서식 5 수치"]["B1"].value or "")
        mm = re.search(r"(\d{4}\.\d{1,2}\.\d{1,2}\.)", hdr)
        if mm:
            w5["A2"] = f"({mm[1]} 기준)"
        vals = {str(r[0]).replace("∙", "·"): r[1] for r in src["서식 5 수치"].iter_rows(min_row=2, values_only=True)}
        for row in range(5, 12):
            name = str(w5[f"A{row}"].value or "").replace("∙", "·")
            w5[f"B{row}"] = int(vals.get(name, 0) or 0)
        body = [list(r) for r in src["서식 6 행"].iter_rows(min_row=2, values_only=True)]
        if "서식 6" in wb.sheetnames:
            w6 = wb["서식 6"]
            if mm:
                w6["A2"] = f"({mm[1]} 기준)"
            if body and body[0][0] != "해당없음":
                if len(body) > 4:
                    insert_rows_keep_merges(w6, 10, len(body) - 4)
                write_rows(w6, 6, 9 + max(0, len(body) - 4), body)
                art_n = len(body)
            else:
                w6["A6"] = "해당없음"
                art_n = 0
    pub_v = None
    if a.pubs and "서식 7" in wb.sheetnames:
        src = openpyxl.load_workbook(a.pubs, data_only=True)
        if "서식 7 수치" not in src.sheetnames:
            print("입력 오류: --pubs에는 publications-check 결과 파일을 넣으세요.")
            return 2
        prod, sent = src["서식 7 수치"]["A2"].value or 0, src["서식 7 수치"]["B2"].value or 0
        w7 = wb["서식 7"]
        w7["A2"] = f"(단위: 철(종) / 생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
        w7["A4"], w7["B4"] = int(prod), int(sent)
        pub_v = (prod, sent)
    sec_n = None
    if a.secret and "서식 10" in wb.sheetnames:
        src = openpyxl.load_workbook(a.secret, data_only=True)
        if "서식 10 수치" not in src.sheetnames or "서식 11 행" not in src.sheetnames:
            print("입력 오류: --secret에는 classified-records-check 결과 파일을 넣으세요.")
            return 2
        w10 = wb["서식 10"]
        w10["A2"] = f"(생산기간 : {a.year}.1.1.～ {a.year}.12.31.)"
        cols = [c.value for c in src["서식 10 수치"][1]][1:10]
        for r in src["서식 10 수치"].iter_rows(min_row=2, values_only=True):
            for c, v in zip(cols, r[1:10]):
                w10[f"{c}{r[0]}"] = int(v or 0)
        rows11 = collections.defaultdict(list)
        for r in src["서식 11 행"].iter_rows(min_row=2, values_only=True):
            rows11[r[0]].append(list(r[1:]))
        if "서식 11" in wb.sheetnames:
            w11 = wb["서식 11"]
            w11["A2"] = f"(생산기간 : {a.year - 3}.1.1.～ {a.year - 3}.12.31.)"
            slots = [("문서류", 6, True), ("도면·카드류", 10, False), ("시청각류", 14, False), ("책자류", 18, True), ("그밖의종류", 22, False)]
            for k, row, pages in reversed(slots):  # 아래 구획부터 끼워 넣어야 위 구획 위치가 안 바뀐다
                body = rows11.get(k, [])
                if len(body) > 1:
                    insert_rows_keep_merges(w11, row + 1, len(body) - 1)
                for i, v in enumerate(body):
                    vals = v if pages else v[:4] + v[5:]
                    for j, val in enumerate(vals):
                        w11.cell(row=row + i, column=1 + j, value=val)
            sec_n = sum(len(v) for v in rows11.values())
            if not sec_n:
                w11["A6"] = "해당없음"
    wb.save(a.out)
    if sec_n is not None:
        print(f"서식 10 비밀 통계 · 서식 11 목록 {sec_n}건 ({a.year - 3}년 생산분)")
    if art_n is not None or pub_v:
        print("서식 5·6 행정박물 " + ("-" if art_n is None else f"{art_n}행") + " · 서식 7 " + (f"생산 {pub_v[0]}종·송부 {pub_v[1]}종" if pub_v else "-"))
    if b9_n is not None:
        print(f"서식 9 회의체 {b9_n}개")
    if av_n or mt_n:
        print(f"서식 3·4 시청각 {av_n}건 · 서식 8 회의 {mt_n}회" + ("".join("\n  확인: " + n for n in notes[:30]) if notes else ""))
    total_f = sum(v for (r, p, t), v in counts.items() if t == "철")
    total_i = sum(v for (r, p, t), v in counts.items() if t == "건")
    print(f"서식 1: 철 {total_f} · 건 {total_i} · 서식 2-1 철 {n1} · 서식 2-2 건 {n2} · 비밀 제외 {len(secret)} · 보존기간 불명 철 {skipped} → {a.out}")
    print("다음: 엑셀에서 열어 합계 수식(총계·소계) 값을 확인하고, 기관 유형에 따라 서식 2 제출 여부를 정한다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
