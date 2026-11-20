#!/usr/bin/env python3
"""기록물통합서비스플랫폼(RAMP) 일괄등록 파일 만들기·점검: 국가기록원 '기록물 일괄등록 서식'(xlsx)을 채운 목록으로
DOCU.TXT(기록물철등록부)·DETAIL.TXT(기록물등록대장)를 규격대로 만들고, 필수·길이·코드·철/건 정합을 점검한다.

사용:  python ramp_batch.py 일괄등록서식_채움.xlsx [--registrar 부서코드,직위,이름,ID] [--date 20261121090000]
                           [--encoding utf-8|cp949] -o 산출/RAMP일괄등록/
입력:  국가기록원 표준·지침·매뉴얼 > 규격·서식 「기록물통합서비스플랫폼 기록물 일괄등록 서식」의 '철 등록부'·'기록물등록대장' 시트
       (머리글 행은 '*시스템구분'으로 시작, 이 스크립트는 머리글 순서를 그대로 쓴다). '특수목록'·'전자파일' 시트에 값이 있으면 SPECIAL.TXT·FILE.TXT도 만든다.
       값은 사람이 쓰는 꼴로 넣어도 된다: 보존기간 '영구·준영구·30년·10년…' → 40·30·25·10…, 일련번호·권호·건수·쪽수는 자리 맞춤.
근거:  「기록물통합서비스플랫폼(RAMP) 기록물 일괄등록 규격」(2024. 10., 국가기록원 디지털혁신과)
         구분자 Tab, 한 행이 한 레코드, 값이 없어도 Tab으로 칸 유지, 입력항목 순서대로
         철분류번호(28) = 처리과기관코드(7)+단위업무코드(8, 구기록물 99999999)+생산년도(4)+철 등록일련번호(6)+권호수(3)
         생산(접수)등록번호(13) = 처리과기관코드(7)+등록일련번호(6), 등록일자 12자리(YYYYMMDDHHMM), 등록일자(등록자) 14자리
         보존기간 코드 01·03·05·10·20·25(30년)·30(준영구)·40(영구), 보존장소 1 기록관(10년 이하)·2 전문관리기관(30년 이상)
         비치종결일자와 비치사유는 함께, 공개여부 9자리(첫 자리 1 공개·2 부분공개·3 비공개 + 1~8호 Y/N), 특수기록물 5자리 Y/N
원칙:  파일을 만들고 점검만 한다. 업로드·등록은 담당자가 RAMP에서 한다. 비밀분류가 3급 이상인 행은 파일에서 빼고 담당자에게 넘긴다.
       업로드 문자 인코딩은 규격에 적혀 있지 않으니 기관 RAMP에서 확인한 값으로 --encoding을 준다(기본 utf-8).
"""
import argparse
import datetime as dt
import os
import re
import sys

DOCU_LEN = [2, 7, 100, 4, 1, 8, 8, 6, 3, 500, 200, 1, 4, 4, 2, 1, 1, 500, 2, 5, 8, 100, 40, 5, 1, 1, 6, 6, 6, 1, 7, 8, 4, 6, 3, 1000, 2, 1, 7, 100, 40, 35, 14]
DETAIL_LEN = [2, 7, 100, 1, 12, 13, 30, 2, 500, 200, 6, 40, 100, 100, 40, 8, 8, 100, 6, 30, 4, 2, 20, 3, 6, 1, 28, 5, 9, 100, 100, 2, 500, 50, 1, 1, 1, 500,
              1, 20, 2, 20, 6, 2, 2, 1000, 2, 20, 7, 100, 40, 35, 14]
PERIOD = {"영구": "40", "준영구": "30", "30년": "25", "20년": "20", "10년": "10", "5년": "05", "3년": "03", "1년": "01",
          "1": "01", "3": "03", "5": "05"}
CODES = {"시스템구분": {"03", "05", "06"}, "단위과제여부": {"1", "2"}, "기록물형태": {"1", "2", "3", "4", "5"},
         "보존기간": {"01", "03", "05", "10", "20", "25", "30", "40"}, "보존방법": {"1", "2", "3"}, "보존장소": {"1", "2"},
         "기록물구분(신/구)": {"1", "2"}, "수정여부": {"0", "1"}, "인수인계구분": {"0", "1", "2"}, "열람조건": {"01", "02", "03"},
         "등록구분": {str(i) for i in range(1, 9)}, "전자기록물여부": {"1", "2"}, "반려여부": {"0", "1"}, "구기록물보존기간": {"01", "03", "05", "10", "20", "25", "30", "40"}}
PAD = {"기록물철등록일련번호": 6, "권호수": 3, "기록물등록건수": 6, "기록물쪽수": 6, "전자파일개수": 6, "쪽수": 6, "분리등록번호(첨부번호)": 2, "건일련번호": 6,
       "열람조건": 2, "보존기간": 2, "구보존기간": 2, "구기록물보존기간": 2}


def base(h):
    h = re.sub(r"\(YYYY[^)]*\)|\[.*?\]|\s|\*", "", str(h or ""))
    return h.replace("인수인계처리과기관코드", "인수인계처리과기관코드")


def load(wb, name):
    if name not in wb.sheetnames:
        return None, []
    rows = [["" if v is None else str(v).strip() for v in r] for r in wb[name].iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows) if r and r[0].replace(" ", "") == "*시스템구분"), None)
    if hi is None:
        return None, []
    head = rows[hi]
    while head and not head[-1]:
        head = head[:-1]
    data = []
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        if not any(r[:len(head)]):
            break  # 예시 아래 코드표는 빈 행 뒤에 있다
        data.append((n, (r + [""] * len(head))[:len(head)]))
    return head, data


def norm(field, v):
    v = v.strip()
    if field in ("보존기간", "구보존기간", "구기록물보존기간"):
        v = PERIOD.get(v.replace(" ", ""), v)
    if field in PAD and v.isdigit():
        v = v.zfill(PAD[field])
    if field == "시스템구분" and v.isdigit():
        v = v.zfill(2)
    return v


def main() -> int:
    ap = argparse.ArgumentParser(description="RAMP 일괄등록 DOCU/DETAIL.TXT 만들기·점검")
    ap.add_argument("form")
    ap.add_argument("--registrar", help="등록자 부서코드,직위,이름,ID — 빈 칸을 채움")
    ap.add_argument("--date", help="등록일자 YYYYMMDDHHMMSS — 빈 칸을 채움(기본: 지금)")
    ap.add_argument("--encoding", default="utf-8")
    ap.add_argument("-o", "--out", default="산출/RAMP일괄등록")
    a = ap.parse_args()
    import openpyxl
    try:
        wb = openpyxl.load_workbook(a.form, read_only=True, data_only=True)
    except Exception as e:
        print(f"입력 오류: 엑셀을 열 수 없습니다({e})", file=sys.stderr)
        return 2
    reg = (a.registrar.split(",") + ["", "", "", ""])[:4] if a.registrar else None
    now = a.date or dt.datetime.now().strftime("%Y%m%d%H%M%S")
    issues, out = [], {}
    files = {}
    for sheet, fname, lens in (("철 등록부", "DOCU.TXT", DOCU_LEN), ("기록물등록대장", "DETAIL.TXT", DETAIL_LEN)):
        head, data = load(wb, sheet)
        if head is None:
            if sheet == "철 등록부":
                print("입력 오류: '철 등록부' 시트(머리글 '*시스템구분')가 없습니다.", file=sys.stderr)
                return 2
            continue
        if len(head) != len(lens):
            issues.append([sheet, "", "머리글", f"열 {len(head)}개 — 규격 {len(lens)}개와 다름(서식 최신판 확인)"])
        names = [base(h) for h in head]
        lines, recs = [], []
        for n, r in data:
            rec = {}
            vals = []
            for k, (h, nm) in enumerate(zip(head, names)):
                v = norm(nm, r[k] if k < len(r) else "")
                if reg and not v and nm in ("등록자부서코드", "등록자직위(직급)", "등록자명", "등록자ID"):
                    v = reg[("등록자부서코드", "등록자직위(직급)", "등록자명", "등록자ID").index(nm)]
                if not v and nm == "등록일자":
                    v = now
                if not v and nm in ("수정여부", "인수인계구분", "반려여부"):
                    v = "0"
                rec[nm] = v
                vals.append(v.replace("\t", " ").replace("\n", " "))
                L = lens[k] if k < len(lens) else 0
                if h.startswith("*") and not v:
                    issues.append([sheet, n, h.lstrip("*"), "필수 항목 빈칸"])
                if L and len(v) > L:
                    issues.append([sheet, n, h.lstrip("*"), f"길이 {len(v)} > 규격 {L}"])
                if v and nm in CODES and v not in CODES[nm]:
                    issues.append([sheet, n, h.lstrip("*"), f"'{v}' — 허용 코드 {sorted(CODES[nm])}"])
            if sheet == "기록물등록대장" and rec.get("비밀분류") in ("03", "04", "05"):
                issues.append([sheet, n, "비밀분류", "비밀기록물 — 일괄등록 파일에서 뺌, 담당자 처리"])
                continue
            recs.append((n, rec))
            lines.append("\t".join(vals))
        out[sheet] = recs
        files[fname] = lines
    # 철 점검
    docs = {}
    for n, r in out.get("철 등록부", []):
        key = r.get("처리과기관코드", "") + r.get("단위업무코드", "") + r.get("생산년도", "") + r.get("기록물철등록일련번호", "") + r.get("권호수", "")
        if len(key) != 28:
            issues.append(["철 등록부", n, "철분류번호", f"처리과기관코드7+단위업무코드8+생산년도4+일련번호6+권호3 = {len(key)}자리(28이어야)"])
        if key in docs:
            issues.append(["철 등록부", n, "철분류번호", f"중복({docs[key][0]}행과 같음)"])
        docs[key] = (n, r)
        p, loc = r.get("보존기간"), r.get("보존장소")
        if p in ("25", "30", "40") and loc == "1" or p in ("01", "03", "05", "10", "20") and loc == "2":
            issues.append(["철 등록부", n, "보존장소", f"보존기간 {p}인데 보존장소 {loc} — 30년 이상은 2(전문관리기관), 10년 이하는 1(기록관)"])
        if bool(r.get("비치종결일자")) != bool(r.get("비치사유")):
            issues.append(["철 등록부", n, "비치", "비치종결일자와 비치사유는 함께 입력"])
        if r.get("단위업무코드") and r.get("기록물구분(신/구)") == "2" and r["단위업무코드"] != "99999999":
            issues.append(["철 등록부", n, "단위업무코드", "구기록물은 '99999999' — 확인"])
        if r.get("시작연도") and r.get("종료연도") and r["시작연도"] > r["종료연도"]:
            issues.append(["철 등록부", n, "시작·종료연도", "시작연도가 종료연도보다 늦음"])
    # 건 점검
    agg = {}
    for n, r in out.get("기록물등록대장", []):
        cls = r.get("분류번호", "")
        if cls and cls not in docs:
            issues.append(["기록물등록대장", n, "분류번호", f"'{cls}' — 철 등록부에 없는 철"])
        no = r.get("생산(접수)등록번호", "")
        if no and (len(no) != 13 or not no.startswith(r.get("처리과기관코드", "-"))):
            issues.append(["기록물등록대장", n, "생산(접수)등록번호", "처리과기관코드7+등록일련번호6 = 13자리"])
        if r.get("생산(접수)등록일자") and not re.fullmatch(r"\d{12}", r["생산(접수)등록일자"]):
            issues.append(["기록물등록대장", n, "생산(접수)등록일자", "YYYYMMDDHHMM 12자리(없으면 협의해 월말일 0900)"])
        if r.get("공개여부") and not re.fullmatch(r"[1-3][YN]{8}", r["공개여부"]):
            issues.append(["기록물등록대장", n, "공개여부", "9자리: 첫 자리 1·2·3 + 비공개 사유 1~8호 Y/N"])
        if r.get("공개여부", "").startswith("1") and "Y" in r.get("공개여부", "")[1:]:
            issues.append(["기록물등록대장", n, "공개여부", "공개(1)인데 비공개 사유 Y가 있음"])
        if r.get("공개여부", "")[:1] in ("2", "3") and "Y" not in r.get("공개여부", "")[1:]:
            issues.append(["기록물등록대장", n, "공개여부", "부분공개·비공개인데 1~8호 사유가 모두 N"])
        if r.get("특수기록물") and not re.fullmatch(r"[YN]{5}", r["특수기록물"]):
            issues.append(["기록물등록대장", n, "특수기록물", "Y/N 5자리"])
        a_ = agg.setdefault(cls, [0, 0])
        a_[0] += 1
        a_[1] += int(r.get("쪽수") or 0) if (r.get("쪽수") or "0").isdigit() else 0
    if out.get("기록물등록대장"):
        for key, (n, r) in docs.items():
            c, pg = agg.get(key, [0, 0])
            if r.get("기록물등록건수", "").isdigit() and int(r["기록물등록건수"]) != c:
                issues.append(["철 등록부", n, "기록물등록건수", f"{int(r['기록물등록건수'])} ≠ 기록물등록대장 {c}건"])
            if r.get("기록물쪽수", "").isdigit() and int(r["기록물쪽수"]) != pg:
                issues.append(["철 등록부", n, "기록물쪽수", f"{int(r['기록물쪽수'])} ≠ 건 쪽수 합 {pg}"])
    for sheet, fname in (("특수목록", "SPECIAL.TXT"), ("전자파일", "FILE.TXT")):
        head, data = load(wb, sheet)
        if head and data:
            files[fname] = ["\t".join(v.replace("\t", " ") for v in r[:len(head)]) for _, r in data]
    os.makedirs(a.out, exist_ok=True)
    for fname, lines in files.items():
        with open(os.path.join(a.out, fname), "w", encoding=a.encoding, errors="strict", newline="\r\n") as fh:
            fh.write("\n".join(lines) + ("\n" if lines else ""))
    rw = openpyxl.Workbook()
    ws = rw.active
    ws.title = "점검"
    ws.append(["시트", "원본 행", "항목", "내용"])
    for x in issues:
        ws.append(x)
    s = rw.create_sheet("요약")
    for row in ([f, len(l)] for f, l in files.items()):
        s.append(row)
    s.append(["점검 항목", len(issues)])
    s.append(["근거", "「기록물통합서비스플랫폼(RAMP) 기록물 일괄등록 규격」(2024. 10.) — Tab 구분, 입력항목 순서, 철분류번호 28자리, 보존기간·보존장소·열람조건 코드"])
    s.append(["주의", f"인코딩 {a.encoding}로 썼다. 업로드 전 기관 RAMP 작업디렉토리(rmsConfig.xml의 batch-record-path)와 인코딩을 확인한다."])
    rw.save(os.path.join(a.out, "일괄등록_점검.xlsx"))
    print(" · ".join(f"{f} {len(l)}행" for f, l in files.items()) + f" · 점검 {len(issues)} → {a.out}")
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
