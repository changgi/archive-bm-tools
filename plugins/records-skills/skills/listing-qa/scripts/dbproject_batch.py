#!/usr/bin/env python3
"""행정정보DB구축 기록물 일괄등록 파일 만들기·점검: 행정정보DB구축(종이기록물 목록·이미지 구축) 결과를
기록관리시스템(RMS)에 일괄등록하기 전에 철·건·특수목록 목록과 이미지 폴더를 규격대로 점검하고 Tab 구분 목록파일을 만든다.

사용:  python dbproject_batch.py 구축목록.xlsx [--images 이미지폴더] [--encoding cp949|utf-8] -o 산출/행정DB일괄등록/
       python dbproject_batch.py --template 빈서식.xlsx          # 머리글만 있는 입력 서식 만들기
입력:  시트 '기록물철등록부'(필수)·'기록물등록대장'·'특수목록'. 첫 행이 머리글이고 항목 이름은 규격 이름(띄어쓰기·괄호 무시).
       값은 사람이 쓰는 꼴도 된다: 보존기간 '영구·준영구·30년·10년…' → 40·30·25·10…, 일련번호·권호·건수·쪽수는 자리 맞춤.
       --images: 철 폴더(28자리)/건 폴더(19자리)/00000001.TIF … 구조를 점검한다.
근거:  「행정정보DB구축 기록물의 기록관리시스템 규격」(2010. 12. 제정, 2023. 4. 개정, 국가기록원)
         Ⅱ.3 입력항목(철 31·건 42·특수목록 7항목), 항목 체크사항 및 오류내용, 이미지 관련 체크사항 및 오류내용
         구분자 Tab, 한 행이 한 레코드, 값이 없어도 Tab으로 칸 유지. 길이는 Byte 기준.
         2023 개정: 구기록물 철·건 생산기관명 길이 100 → 256
원칙:  점검과 파일 만들기까지만 한다. 일괄등록 실행은 담당자가 RMS 작업디렉토리(rmsConfig.xml의 batch-record-path)에서 한다.
       비밀기록물(특수기록물 2번째 자리 Y)은 목록파일에서 빼고 건수만 알린다. 원본 이미지는 읽기만 한다.
미확인: 목록파일 이름과 문자 인코딩은 규격에 없다. 기관 RMS에서 확인해 --encoding과 파일 이름을 맞춘다(기본 cp949, 이름 DOCU/DETAIL/SPECIAL.TXT).
"""
import argparse
import datetime as dt
import os
import re
import sys

# (항목, 길이, 필수, 허용값 또는 None, 기본값)
DOC = [("처리과기관코드", 7, 1, None, ""), ("구기록물철생산기관명", 256, 1, None, ""), ("단위업무코드", 8, 1, None, "99999999"),
       ("생산년도", 4, 1, None, ""), ("소기능코드", 8, 0, None, ""), ("기록물철등록일련번호", 6, 1, None, ""), ("권호수", 3, 1, None, "001"),
       ("기록물철제목", 500, 1, None, ""), ("기록물철기타제목", 200, 0, None, ""), ("기록물형태", 1, 1, set("12345"), ""),
       ("시작년도", 4, 1, None, ""), ("종료연도", 4, 1, None, ""),
       ("보존기간", 2, 1, {"01", "03", "05", "10", "20", "25", "30", "40"}, ""), ("보존방법", 1, 1, set("123"), ""),
       ("보존장소", 1, 1, set("12"), ""), ("비치종결일자", 8, 0, None, ""), ("비치사유", 100, 0, None, ""), ("업무담당자", 40, 0, None, ""),
       ("구기록물철분류번호", 5, 0, None, ""), ("기록물구분(신/구)", 1, 1, {"2"}, "2"), ("수정여부", 1, 1, {"0"}, "0"),
       ("기록물등록건수", 6, 1, None, ""), ("기록물쪽수", 6, 1, None, ""), ("전자파일갯수", 6, 1, {"000000"}, "000000"),
       ("인수인계구분", 1, 1, {"0"}, "0"), ("인수인계처리과기관코드", 7, 0, None, ""), ("인수인계단위업무코드", 8, 0, None, ""),
       ("인수인계생산년도", 4, 0, None, ""), ("인수인계기록물철등록일련번호", 6, 0, None, ""), ("인수인계권호수", 3, 0, None, ""),
       ("내용요약", 1000, 0, None, "")]
ITEM = [("처리과기관코드", 7, 1, None, ""), ("구기록물생산기관명", 256, 1, None, ""), ("등록구분", 1, 1, set("12345678"), ""),
        ("생산(접수)등록일자", 12, 1, None, ""), ("생산(접수)등록번호", 13, 1, None, ""), ("구기록물문서번호", 30, 0, None, ""),
        ("분리등록번호(첨부번호)", 2, 1, {"00"}, "00"), ("제목", 500, 1, None, ""), ("기타제목", 200, 0, None, ""), ("쪽수", 6, 1, None, ""),
        ("결재권자(직위/직급)", 40, 0, None, ""), ("검토자", 100, 0, None, ""), ("협조자", 100, 0, None, ""), ("기안자(업무담당자)", 40, 0, None, ""),
        ("결재일자", 8, 0, None, ""), ("시행일자", 8, 0, None, ""), ("수신자(발신자)", 100, 0, None, ""), ("문서과배부번호", 6, 0, None, ""),
        ("생산기관등록번호", 30, 0, None, ""), ("첨부수", 4, 0, None, ""),
        ("첨부물의종류", 2, 0, {"01", "02", "03", "04", "05", "99"}, ""), ("첨부물기타", 20, 0, None, ""), ("언어", 3, 0, None, ""),
        ("전자기록물여부", 1, 1, {"2"}, "2"), ("분류번호", 28, 1, None, ""), ("특수기록물", 5, 0, None, ""), ("공개여부", 9, 0, None, ""),
        ("공개제한부분표시", 100, 0, None, ""), ("공개시기", 100, 0, None, ""),
        ("구기록물보존기간", 2, 1, {"01", "03", "05", "10", "20", "30", "40"}, ""),
        ("시청각내용요약", 500, 0, None, ""), ("시청각기록물형태", 50, 0, None, ""), ("기록물구분(신/구)", 1, 1, {"2"}, "2"),
        ("수정여부", 1, 1, {"0"}, "0"), ("반려여부", 1, 1, {"0"}, "0"), ("건내용요약", 500, 0, None, ""),
        ("포맷", 1, 0, set("123459"), ""), ("포맷기타", 20, 0, None, ""),
        ("저장매체", 2, 0, {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "99"}, ""), ("저장매체기타", 20, 0, None, ""),
        ("건일련번호", 6, 1, None, ""), ("건세부유형", 2, 0, {"11", "21", "22"}, "")]
SPEC = [("특수목록위치", 1, 1, set("12"), ""), ("기록물철분류번호", 28, 0, None, ""), ("기록물건등록번호", 19, 0, None, ""),
        ("일련번호", 3, 1, None, ""), ("특수목록1", 100, 0, None, ""), ("특수목록2", 100, 0, None, ""), ("특수목록3", 100, 0, None, "")]
SHEETS = (("기록물철등록부", DOC, "DOCU.TXT"), ("기록물등록대장", ITEM, "DETAIL.TXT"), ("특수목록", SPEC, "SPECIAL.TXT"))
PERIOD = {"영구": "40", "준영구": "30", "30년": "25", "20년": "20", "10년": "10", "5년": "05", "3년": "03", "1년": "01",
          "1": "01", "3": "03", "5": "05"}
PAD = {"기록물철등록일련번호": 6, "권호수": 3, "기록물등록건수": 6, "기록물쪽수": 6, "전자파일갯수": 6, "쪽수": 6, "건일련번호": 6,
       "분리등록번호(첨부번호)": 2, "일련번호": 3}
SECTION_FORM = {"1": "1", "2": "1", "3": "2", "4": "2", "5": "3", "6": "4", "7": "5", "8": "5"}  # 등록구분 → 기록물형태
SUBTYPE_FORM = {"11": "1", "21": "5", "22": "5"}


def key(h):
    return re.sub(r"[\s*①-⑨<>]|필수|옵션", "", str(h or ""))


def blen(v, enc):
    try:
        return len(v.encode(enc))
    except UnicodeEncodeError:
        return -1


def valid_date(v, fmt):
    try:
        dt.datetime.strptime(v, fmt)
        return True
    except ValueError:
        return False


def read_sheet(wb, name, spec):
    if name not in wb.sheetnames:
        return None
    rows = [["" if v is None else str(v).strip() for v in r] for r in wb[name].iter_rows(values_only=True)]
    if not rows:
        return []
    head = [key(h) for h in rows[0]]
    want = {key(f): i for i, (f, *_) in enumerate(spec)}
    pos = {want[h]: j for j, h in enumerate(head) if h in want}
    out = []
    for n, r in enumerate(rows[1:], 2):
        if not any(r):
            continue
        if all(r[pos[i]] == str(spec[i][1]) for i in pos if pos[i] < len(r)):
            continue  # --template의 길이 안내 행
        rec = {}
        for i, (f, L, req, codes, dflt) in enumerate(spec):
            v = r[pos[i]] if i in pos and pos[i] < len(r) else ""
            v = re.sub(r"\.0$", "", v) if re.fullmatch(r"\d+\.0", v) else v
            if f in ("보존기간", "구기록물보존기간"):
                v = PERIOD.get(v.replace(" ", ""), v)
            if f in PAD and v.isdigit():
                v = v.zfill(PAD[f])
            rec[f] = v or dflt
        out.append((n, rec))
    return out


def check(sheet, spec, recs, enc, issues):
    for n, r in recs:
        for f, L, req, codes, _ in spec:
            v = r[f]
            if req and not v:
                issues.append([sheet, n, f, f"{f} 항목이 누락되었습니다."])
                continue
            b = blen(v, enc)
            if b < 0:
                issues.append([sheet, n, f, f"{enc}로 쓸 수 없는 문자가 있습니다."])
            elif v and b > L:
                issues.append([sheet, n, f, f"{f} 항목이 {L}Byte를 초과하였습니다.({b}Byte)"])
            if v and codes and v not in codes:
                issues.append([sheet, n, f, f"{f} 항목이 잘못되었습니다.('{v}', 허용 {'·'.join(sorted(codes))})"])


def main() -> int:
    ap = argparse.ArgumentParser(description="행정정보DB구축 기록물 일괄등록 목록·이미지 점검")
    ap.add_argument("form", nargs="?")
    ap.add_argument("--images", help="이미지 최상위 폴더(철 폴더들이 있는 곳)")
    ap.add_argument("--encoding", default="cp949")
    ap.add_argument("--template", help="빈 입력 서식 xlsx를 만들고 끝냄")
    ap.add_argument("-o", "--out", default="산출/행정DB일괄등록")
    a = ap.parse_args()
    import openpyxl
    if a.template:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for name, spec, _ in SHEETS:
            ws = wb.create_sheet(name)
            ws.append([("*" if req else "") + f for f, L, req, *_ in spec])
            ws.append([f"{L}" for f, L, *_ in spec])
        wb.save(a.template)
        print(f"서식 → {a.template} (2행은 길이(Byte) 안내 — 지우고 입력)")
        return 0
    if not a.form:
        ap.error("입력 목록 xlsx가 필요합니다(또는 --template)")
    try:
        wb = openpyxl.load_workbook(a.form, read_only=True, data_only=True)
    except Exception as e:
        print(f"입력 오류: 엑셀을 열 수 없습니다({e})", file=sys.stderr)
        return 2
    data = {name: read_sheet(wb, name, spec) for name, spec, _ in SHEETS}
    if data["기록물철등록부"] is None:
        print("입력 오류: '기록물철등록부' 시트가 없습니다(--template으로 서식을 만드세요).", file=sys.stderr)
        return 2
    issues, secret = [], 0
    for name, spec, _ in SHEETS:
        if data[name]:
            check(name, spec, data[name], a.encoding, issues)
    docs = {}
    for n, r in data["기록물철등록부"]:
        cls = r["처리과기관코드"] + r["단위업무코드"] + r["생산년도"] + r["기록물철등록일련번호"] + r["권호수"]
        S = "기록물철등록부"
        if r["단위업무코드"] and r["단위업무코드"] != "99999999":
            issues.append([S, n, "단위업무코드", "단위업무코드가 '99999999'가 아닙니다."])
        if r["소기능코드"] and len(r["소기능코드"]) != 8:
            issues.append([S, n, "소기능코드", "소기능코드 항목이 8Byte 가 아닙니다."])
        for f, L in (("처리과기관코드", 7), ("생산년도", 4), ("기록물철등록일련번호", 6), ("권호수", 3), ("시작년도", 4), ("종료연도", 4)):
            if r[f] and not (len(r[f]) == L and (r[f].isdigit() or f == "처리과기관코드")):
                issues.append([S, n, f, f"{f} 항목이 {L}자리 숫자가 아닙니다."])
        if r["시작년도"] and r["종료연도"] and r["시작년도"] > r["종료연도"]:
            issues.append([S, n, "시작·종료연도", "시작년도가 종료연도보다 늦습니다."])
        if r["비치종결일자"] and not (re.fullmatch(r"\d{8}", r["비치종결일자"]) and valid_date(r["비치종결일자"], "%Y%m%d")):
            issues.append([S, n, "비치종결일자", "비치종결일자가 정확하지 않습니다."])
        if bool(r["비치종결일자"]) != bool(r["비치사유"]):
            issues.append([S, n, "비치", "비치종결일자와 비치사유 항목이 함께 존재하지 않습니다."])
        p, loc = r["보존기간"], r["보존장소"]
        if p in ("25", "30", "40") and loc == "1" or p in ("01", "03", "05", "10") and loc == "2":
            issues.append([S, n, "보존장소", f"보존기간 {p}·보존장소 {loc} — 규격 예시: 1 기록관(10년 이하)·2 전문관리기관(30년 이상)"])
        for f in ("기록물등록건수", "기록물쪽수"):
            if r[f] and not r[f].isdigit():
                issues.append([S, n, f, f"{f} 항목에 문자가 있습니다."])
        if r["기록물등록건수"].isdigit() and int(r["기록물등록건수"]) == 0:
            issues.append([S, n, "기록물등록건수", "기록물등록건수 항목이 '0'입니다."])
        if cls in docs:
            issues.append([S, n, "기록물철등록일련번호", f"기록물철 일련번호({r['기록물철등록일련번호']})가 중복되었습니다.({docs[cls][0]}행)"])
        docs[cls] = (n, r)
    agg, items = {}, {}
    keep_items = []
    for n, r in data["기록물등록대장"] or []:
        S = "기록물등록대장"
        if r["특수기록물"][1:2] == "Y":
            secret += 1
            continue
        keep_items.append((n, r))
        d, no = r["생산(접수)등록일자"], r["생산(접수)등록번호"]
        if d and not (re.fullmatch(r"\d{12}", d) and valid_date(d, "%Y%m%d%H%M")):
            issues.append([S, n, "생산(접수)등록일자", "생산(접수)등록일자 항목이 정확하지 않습니다.(날짜형식이 아님, 12자리)"])
        if no and not (len(no) == 13 and no[7:].isdigit() and no.startswith(r["처리과기관코드"])):
            issues.append([S, n, "생산(접수)등록번호", "생산(접수)등록번호 항목이 잘못되었습니다.(처리과기관코드7+등록일련번호6)"])
        for f in ("결재일자", "시행일자"):
            if r[f] and not valid_date(r[f], "%Y%m%d"):
                issues.append([S, n, f, f"{f} 항목이 잘못되었습니다.(날짜형식이 아님)"])
        for f in ("쪽수", "건일련번호", "문서과배부번호"):
            if r[f] and not r[f].isdigit():
                issues.append([S, n, f, f"{f} 항목이 잘못되었습니다.(숫자가 아님)"])
        if r["첨부수"] and r["첨부수"] not in ("0", "0000"):
            issues.append([S, n, "첨부수", "비전자기록물은 본문/첨부 구분이 없어 '0'으로 설정"])
        cls = r["분류번호"]
        if cls and (len(cls) != 28 or not cls[7:].isdigit()):
            issues.append([S, n, "분류번호", "철분류번호 항목의 형식이 잘못되었습니다.(28자리, 기관코드 외 숫자)"])
        elif cls and cls not in docs:
            issues.append([S, n, "분류번호", "철분류번호에 해당하는 기록물철이 존재하지 않습니다."])
        if r["특수기록물"] and not re.fullmatch(r"[YN]{5}", r["특수기록물"]):
            issues.append([S, n, "특수기록물", "특수기록물 항목의 형식이 잘못되었습니다.(Y/N 5자리)"])
        o = r["공개여부"]
        if o:
            if len(o) != 9:
                issues.append([S, n, "공개여부", "공개여부 항목의 자릿수가 일치하지 않습니다.(9Byte)"])
            elif o[0] not in "123":
                issues.append([S, n, "공개여부", "공개여부 항목 중 공개코드값이 잘못되었습니다.(1:공개,2:부분공개,3:비공개)"])
            elif not re.fullmatch(r"[YN]{8}", o[1:]):
                issues.append([S, n, "공개여부", "공개여부 항목 중 공개등급이 잘못되었습니다.(Y/N)"])
            elif o[0] == "1" and "Y" in o[1:]:
                issues.append([S, n, "공개여부", "공개(1)인데 비공개 사유 Y가 있음 — 확인"])
        sec, form = r["등록구분"], None
        if r["건세부유형"] and sec in SECTION_FORM and SUBTYPE_FORM.get(r["건세부유형"]) != SECTION_FORM[sec]:
            issues.append([S, n, "건세부유형", "등록구분에 해당하는 세부유형이 아닙니다."])
        if cls in docs and sec in SECTION_FORM and docs[cls][1]["기록물형태"] != SECTION_FORM[sec]:
            issues.append([S, n, "등록구분", f"등록구분 {sec}(기록물형태 {SECTION_FORM[sec]})이 철 기록물형태 {docs[cls][1]['기록물형태']}와 다름 — 확인"])
        k = (cls, r["건일련번호"])
        if k in items:
            issues.append([S, n, "건일련번호", f"기록물건 일련번호가 중복되었습니다.({items[k]}행)"])
        items[k] = n
        g = agg.setdefault(cls, [0, 0, set()])
        g[0] += 1
        g[1] += int(r["쪽수"]) if r["쪽수"].isdigit() else 0
        if d[:4] and no:
            g[2].add(d[:4] + no + r["분리등록번호(첨부번호)"])
    if data["기록물등록대장"] is not None:
        for cls, (n, r) in docs.items():
            c, pg, _ = agg.get(cls, [0, 0, set()])
            if c == 0:
                issues.append(["기록물철등록부", n, "기록물건", "기록물철에 기록물건이 존재하지 않습니다."])
                continue
            if r["기록물등록건수"].isdigit() and int(r["기록물등록건수"]) != c:
                issues.append(["기록물철등록부", n, "기록물등록건수", f"기록물철의 기록물등록건수({int(r['기록물등록건수'])})와 등록된 기록물건수({c})가 일치하지 않습니다."])
            if r["기록물쪽수"].isdigit() and int(r["기록물쪽수"]) != pg:
                issues.append(["기록물철등록부", n, "기록물쪽수", f"기록물철의 기록물쪽수({int(r['기록물쪽수'])})와 등록된 기록물건의 쪽수 합({pg})이 일치하지 않습니다."])
    item_ids = {x for g in agg.values() for x in g[2]}
    for n, r in data["특수목록"] or []:
        S = "특수목록"
        if r["특수목록위치"] == "1":
            if r["기록물건등록번호"]:
                issues.append([S, n, "기록물건등록번호", "특수목록위치가 1(철)이면 기록물건등록번호는 비움"])
            if r["기록물철분류번호"] not in docs:
                issues.append([S, n, "기록물철분류번호", "기록물철분류번호의 철이 존재하지 않습니다."])
        elif r["특수목록위치"] == "2":
            if r["기록물철분류번호"]:
                issues.append([S, n, "기록물철분류번호", "특수목록위치가 2(건)이면 기록물철분류번호는 비움"])
            if r["기록물건등록번호"] not in item_ids:
                issues.append([S, n, "기록물건등록번호", "기록물건등록번호 건이 존재하지 않습니다.(생산년도4+등록번호13+분리2)"])
        if r["일련번호"] and not (len(r["일련번호"]) == 3 and r["일련번호"].isdigit()):
            issues.append([S, n, "일련번호", "일련번호 항목이 잘못되었습니다."])
        if not (r["특수목록1"] or r["특수목록2"] or r["특수목록3"]):
            issues.append([S, n, "특수목록", "특수목록 항목이 누락되었습니다."])
        if any(re.search(r"\d{6}-?[1-4]\d{6}", r[f]) for f in ("특수목록1", "특수목록2", "특수목록3")):
            issues.append([S, n, "특수목록", "주민등록번호 꼴의 값 — 개인정보 처리 근거 확인(파일에는 그대로 씀, 점검표에는 값 안 남김)"])
    img = []
    if a.images:
        img = check_images(a.images, docs, agg)
        issues += img
    os.makedirs(a.out, exist_ok=True)
    written = {}
    for name, spec, fname in SHEETS:
        recs = keep_items if name == "기록물등록대장" else data[name]
        if not recs:
            continue
        lines = ["\t".join(r[f].replace("\t", " ").replace("\n", " ") for f, *_ in spec) for _, r in recs]
        with open(os.path.join(a.out, fname), "w", encoding=a.encoding, errors="replace", newline="\r\n") as fh:
            fh.write("\n".join(lines) + "\n")
        written[fname] = len(lines)
    rw = openpyxl.Workbook()
    ws = rw.active
    ws.title = "점검"
    ws.append(["시트", "원본 행", "항목", "부적합 시 오류내용(규격 문구)"])
    for x in issues:
        ws.append(x)
    s = rw.create_sheet("요약")
    s.append(["구분", "값"])
    for f, c in written.items():
        s.append([f, f"{c}행"])
    s.append(["철", len(docs)])
    s.append(["건", sum(g[0] for g in agg.values())])
    s.append(["비밀기록물(뺌)", secret])
    s.append(["점검 항목", len(issues)])
    s.append(["이미지 점검", f"{len(img)}건" if a.images else "안 함(--images)"])
    s.append(["근거", "「행정정보DB구축 기록물의 기록관리시스템 규격」(2023. 4. 개정) 입력항목·체크사항·이미지 체크사항"])
    s.append(["미확인", f"목록파일 이름·인코딩은 규격에 없음 — {a.encoding}로 씀, 기관 RMS에서 확인"])
    rw.save(os.path.join(a.out, "행정DB일괄등록_점검.xlsx"))
    print(" · ".join(f"{f} {c}행" for f, c in written.items()) + f" · 비밀 제외 {secret} · 점검 {len(issues)} → {a.out}")
    return 1 if issues else 0


def check_images(root, docs, agg):
    out, S = [], "이미지"
    if not os.path.isdir(root):
        return [[S, "", root, "이미지 폴더가 없습니다."]]
    seen_docs = set()
    for dname in sorted(os.listdir(root)):
        dp = os.path.join(root, dname)
        if not os.path.isdir(dp):
            continue
        if len(dname) != 28:
            out.append([S, "", dname, f"기록물철 디렉토리명({dname})이 28자리가 아닙니다."])
            continue
        if not os.access(dp, os.R_OK | os.X_OK):
            out.append([S, "", dname, f"철 디렉토리({dname})에 읽을 수 있는 권한이 없습니다."])
            continue
        if dname not in docs:
            out.append([S, "", dname, f"철 디렉토리({dname})에 해당하는 기록물철 분류번호가 존재하지 않습니다"])
        seen_docs.add(dname)
        subs = [x for x in sorted(os.listdir(dp)) if os.path.isdir(os.path.join(dp, x))]
        if not subs:
            out.append([S, "", dname, f"철 디렉토리({dname})에 건 디렉토리가 존재하지 않습니다"])
        known = agg.get(dname, [0, 0, set()])[2]
        for iname in subs:
            ip = os.path.join(dp, iname)
            where = f"{dname}/{iname}"
            if len(iname) != 19:
                out.append([S, "", where, f"기록물건 디렉토리명({iname})이 19자리가 아닙니다."])
                continue
            if agg and iname not in known:
                other = next((c for c, g in agg.items() if iname in g[2]), None)
                msg = (f"기록물건의 철분류번호({other})와 해당 철 디렉토리({dname})가 일치하지 않습니다." if other
                       else f"기록물건 디렉토리({iname})에 해당하는 기록물건이 존재하지 않습니다.")
                out.append([S, "", where, msg])
            files = sorted(x for x in os.listdir(ip) if os.path.isfile(os.path.join(ip, x)))
            if not files:
                out.append([S, "", where, f"기록물건 디렉토리({iname})에 이미지파일이 존재하지 않습니다."])
                continue
            nums = []
            for f in files:
                stem, ext = os.path.splitext(f)
                if len(stem) != 8:
                    out.append([S, "", f"{where}/{f}", f"기록물건 디렉토리({iname})에 이미지 파일의 파일명이({f})이 8자리가 아닙니다."])
                elif not stem.isdigit():
                    out.append([S, "", f"{where}/{f}", f"기록물건 디렉토리({iname})에 이미지 파일의 파일명이({f})이 숫자가 아닙니다."])
                else:
                    nums.append(int(stem))
                if ext.upper() not in (".TIF", ".TIFF"):
                    out.append([S, "", f"{where}/{f}", "이미지파일 확장자가 TIF가 아님(규격 예: 00000001.TIF) — 확인"])
                if os.path.getsize(os.path.join(ip, f)) == 0:
                    out.append([S, "", f"{where}/{f}", f"기록물건 디렉토리({iname})에 파일명이({f})인 파일의 크기가 '0'입니다."])
            if nums and sorted(nums) != list(range(1, len(nums) + 1)):
                out.append([S, "", where, f"기록물건 디렉토리({iname})에 파일명의 일련번호들을 확인하십시요.(1~{len(nums)} 연속이어야)"])
    for cls, (n, r) in docs.items():
        if cls not in seen_docs:
            out.append([S, n, cls, "목록의 철에 해당하는 철 디렉토리가 없음 — 이미지 미구축 여부 확인"])
    return out


if __name__ == "__main__":
    sys.exit(main())
