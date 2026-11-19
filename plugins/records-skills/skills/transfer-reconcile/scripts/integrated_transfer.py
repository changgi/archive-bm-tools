#!/usr/bin/env python3
"""기록물 통합(전자+비전자) 이관 준비: 기록관리시스템 건 목록에서 통합이관 대상(비전자문서건, 전자문서의 비전자 붙임(별송))을
찾고, 「2026년 기록물관리 지침」 참고 1의 '일반 문서류 건 색인목록'·'비전자 붙임(별송) 색인목록' 서식 초안을 만든다.

사용:  python integrated_transfer.py 건목록.xlsx [--until 2015] -o 산출/통합이관_준비.xlsx
목록 열(머리글 이름으로 찾음): 철분류번호, 철제목, 철구분(전자·혼합·비전자), 생산연도, 등록번호, 등록일자, 건명, 형태(전자·비전자),
       보낸기관, 받은기관, 쪽수, 붙임(붙임 표기·본문 등 붙임 정보가 든 열 — 여러 개면 모두 봄), 분리등록(예/아니오), 비밀구분
근거:  「2026년 기록물관리 지침」 제3장 참고 1 기록물 통합(전자+비전자) 이관 전 준비사항:
         대상 — 기록관리시스템으로 인수·등록한 비전자문서(~'15년 생산분): 전자철의 비전자 붙임(별송)·분리등록 첨부물,
                혼합철의 비전자문서건·비전자 붙임·분리등록첨부물, 비전자철의 비전자문서건·분리등록첨부물
         ① 대상 목록 확보: 비전자철·혼합철 내 비전자문서 건목록, 전자·혼합철 내 전자문서의 비전자 붙임 목록
            (국가기록원 육안검수결과 '붙임 별송 누락목록' 참조, 분리등록 정보, 기안문의 '붙임(별첨, 별송)' 표기)
         붙임(별첨) 표기 예시: 공시송달·따로보냄·따로붙임·모사전송·반송용봉투·반송용우표·반신용봉투·별도배포·별도보고·별도보관·
            별도보냄·별도송부(별송)·별도우송·별도자료·별도제출·별도첨부(별첨)·별지·별책·사송·수동송부·우편·인편송부·종이문서·
            직접송부·직접전달·책자인쇄·출력송부·택배·홍보책자·FAX·Nafis별송·그림·디스켓·비디오·슬라이드·인화사진·테이프·필름·
            CD·DAT·DVD·LD·LP·SP
         ② 실물확인: 철 분류번호·건 문서번호 표시 확인, 실물과 시스템 대조해 색인목록 작성(비고: 실물확인 여부)
원칙:  후보만 찾는다. 실물확인·색인목록 확정·이관은 사람이 한다. 비밀 표시 행은 넣지 않는다(시행령 제68조 별도 절차).
"""
import argparse
import collections
import os
import re
import sys

STRONG = ["공시송달", "따로보냄", "따로붙임", "모사전송", "반송용봉투", "반송용우표", "반신용봉투", "별도배포", "별도보고", "별도보관",
          "별도보냄", "별도송부", "별송", "별도우송", "별도자료", "별도제출", "별도첨부", "별첨", "별책", "사송", "수동송부", "인편송부",
          "종이문서", "직접송부", "직접전달", "책자인쇄", "출력송부", "택배", "홍보책자", "Nafis별송", "디스켓", "인화사진", "슬라이드"]
WEAK = ["별지", "우편", "그림", "비디오", "테이프", "필름", "FAX", "CD", "DAT", "DVD", "LD", "LP", "SP"]
LATIN = {"FAX", "CD", "DAT", "DVD", "LD", "LP", "SP"}
ALIAS = {"철번호": ["철분류번호", "기록물철분류번호", "분류번호"], "철제목": ["철제목", "기록물철제목"], "철구분": ["철구분", "철형태", "기록물철형태"],
         "연도": ["생산연도", "생산년도"], "등록번호": ["등록번호", "생산등록번호", "문서번호"], "일자": ["등록일자", "생산일자", "시행일자"],
         "건명": ["건명", "건제목", "제목"], "형태": ["형태", "기록물형태", "전자구분"], "보낸": ["보낸기관", "발신기관"], "받은": ["받은기관", "수신기관"],
         "쪽수": ["쪽수", "면수"], "분리": ["분리등록", "분리등록여부"], "비밀": ["비밀구분", "비밀여부"]}
ATT = ["붙임", "붙임표기", "첨부", "본문", "비고", "붙임정보"]


def hits(text):
    t = text or ""
    strong = [k for k in STRONG if k in t.replace(" ", "")]
    weak = []
    for k in WEAK:
        if k in LATIN:
            if re.search(rf"(?<![A-Za-z]){k}(?![A-Za-z])", t):
                weak.append(k)
        elif k in t:
            weak.append(k)
    return strong, weak


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--until", type=int, default=2015, help="대상 생산연도 상한(지침: ~'15년 생산분)")
    ap.add_argument("-o", "--out", default="산출/통합이관_준비.xlsx")
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Font
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v} | set(ATT)
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    att_cols = [i for i, h in enumerate(head) if h in ATT]
    items = []
    for r in rows[hi + 1:]:
        if not any(r):
            continue
        x = {k: (r[i] if i < len(r) else "") for k, i in idx.items()}
        x["_att"] = " / ".join(r[i] for i in att_cols if i < len(r) and r[i])
        items.append(x)
    secret = sum(1 for x in items if x.get("비밀"))
    items = [x for x in items if not x.get("비밀")]
    def yr(x):
        m = re.search(r"((?:19|20)\d{2})", x.get("연도", "") or x.get("일자", ""))
        return int(m[1]) if m else None
    nonelec, attach, sep = [], [], []
    for x in items:
        y = yr(x)
        if y and y > a.until:
            continue
        is_non = "비전자" in x.get("형태", "") or "종이" in x.get("형태", "")
        if is_non:
            nonelec.append(x)
            continue
        if str(x.get("분리", "")).strip() in ("예", "Y", "O", "1"):
            sep.append(x)
        s, w = hits(x["_att"] + " " + x.get("건명", ""))
        if s or w:
            attach.append((x, s, w))
    wb = openpyxl.Workbook()
    sm = wb.active
    sm.title = "요약"
    sm.append(["항목", "건수", "설명"])
    sm.append(["비전자문서 건", len(nonelec), "비전자철·혼합철 내 비전자문서 — 실물과 시스템 대조 후 건 색인목록"])
    sm.append(["비전자 붙임(별송) 후보 — 강한 단서", sum(1 for _, s, _ in attach if s), "별송·따로보냄·우편 송부 등 표기"])
    sm.append(["비전자 붙임(별송) 후보 — 약한 단서만", sum(1 for _, s, w in attach if not s), "그림·필름·CD 등 — 전자 파일일 수 있어 확인 필요"])
    sm.append(["분리등록 첨부물", len(sep), "분리등록 표시가 있는 건"])
    sm.append(["대상 연도", f"~{a.until}", "지침: 기록관리시스템 인수·등록 비전자문서(~'15년 생산분)"])
    if secret:
        sm.append(["비밀 표시 행(제외)", secret, "비밀기록물은 시행령 제68조 별도 절차"])
    sm.append(["다음", "", "국가기록원 '전자기록물 오류조치 공문'의 비전자 붙임(별송) 목록·육안검수결과와 대조하고 실물확인 열을 채운다"])
    c1 = wb.create_sheet("비전자 붙임 후보")
    c1.append(["철분류번호", "철제목", "등록번호", "건명", "강한 단서", "약한 단서", "붙임 표기·본문(앞 120자)"])
    for x, s, w in sorted(attach, key=lambda t: (not t[1], t[0].get("철번호", ""))):
        c1.append([x.get("철번호", ""), x.get("철제목", ""), x.get("등록번호", ""), x.get("건명", ""), ", ".join(s), ", ".join(w), x["_att"][:120]])

    def index_sheet(title, rows_by_file, cols, row_of):
        wsx = wb.create_sheet(title)
        for (fno, ftitle, fyear), rs in rows_by_file.items():
            wsx.append([f"기록물철 제목: {ftitle}"])
            wsx.cell(wsx.max_row, 1).font = Font(bold=True)
            wsx.append([f"분류번호 : {fno}", "", "", "", "", f"(생산연도 : {fyear} )"])
            wsx.append(cols)
            for i, x in enumerate(rs, 1):
                wsx.append(row_of(i, x))
            wsx.append([])
        return wsx

    by_file = collections.OrderedDict()
    for x in sorted(nonelec, key=lambda x: (x.get("철번호", ""), x.get("일자", ""))):
        by_file.setdefault((x.get("철번호", ""), x.get("철제목", ""), x.get("연도", "")), []).append(x)
    index_sheet("건 색인목록(서식)", by_file, ["일련번호", "등록일자", "등록번호", "건명", "보낸 기관", "받은 기관", "쪽표시", "비고(실물확인 여부)"],
                lambda i, x: [i, x.get("일자", ""), x.get("등록번호", ""), x.get("건명", ""), x.get("보낸", ""), x.get("받은", ""), x.get("쪽수", ""), ""])
    by_file2 = collections.OrderedDict()
    for x, s, w in sorted(attach, key=lambda t: t[0].get("철번호", "")):
        by_file2.setdefault((x.get("철번호", ""), x.get("철제목", ""), x.get("연도", "")), []).append((x, s, w))
    index_sheet("붙임(별송) 색인목록(서식)", by_file2, ["일련번호", "등록번호", "건명", "붙임(별송) 표기", "전자문서 여부", "비고(실물확인 여부)"],
                lambda i, t: [i, t[0].get("등록번호", ""), t[0].get("건명", ""), ", ".join(t[1] + t[2]), "전자", ""])
    if sep:
        c3 = wb.create_sheet("분리등록 첨부물")
        c3.append(["철분류번호", "철제목", "등록번호", "건명"])
        for x in sep:
            c3.append([x.get("철번호", ""), x.get("철제목", ""), x.get("등록번호", ""), x.get("건명", "")])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"비전자문서 건 {len(nonelec)} · 비전자 붙임 후보 {len(attach)}(강한 단서 {sum(1 for _, s, _ in attach if s)}) · 분리등록 {len(sep)}"
          + (f" · 비밀 제외 {secret}" if secret else "") + f" (~{a.until}년) → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
