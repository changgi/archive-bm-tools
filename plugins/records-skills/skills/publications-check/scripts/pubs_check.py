#!/usr/bin/env python3
"""간행물 점검: 기관 간행물 목록의 발간등록번호·발간주기·송부(책자 3부+PDF)·생략대상을 국가기록원 지침대로 보고, 서식 7 수치를 낸다.

사용:  python pubs_check.py 간행물목록.xlsx --year 2025 -o 산출/간행물_점검.xlsx
목록 열(머리글 이름으로 찾음): 제목, 발간등록번호, 발행일, 발간형태(책자·전자), 간행물유형(팸플릿·리플릿·소식지 등), 송부일, 송부부수, PDF송부(Y/N), 공개구분, 비밀구분
근거:  공공기록물법 제22조, 시행규칙 제36조(발간등록)·제37조(분류)
       국가기록원 「2026년 기록물관리 지침」 간행물 관리:
         원고 확정 후 인쇄 전에 발간등록번호 신청, 앞표지 좌측 상단에 표기
         발간 주기(번호 뒤 2자리): 단행본 01, 연속간행물 02~15. 연속간행물은 제목·발행기관·발간주기가 바뀌면 새 번호
         책자형태 간행물은 국가기록원에 책자 3부와 같은 내용의 전자파일(PDF) 1점을 함께 송부(관할 기록관·영구기록물관리기관에도 따로 송부)
         송부 기한: 발간 후 15일이 지나기 전까지(시행령 제55조④, 국가기록포털 발간등록 안내 2026-11-20 확인)
         발간등록 생략대상(수첩·팸플릿·리플릿·속보·만화·기관신문·추록, 홍보자료·소식지·안내지·지도, 장학·학습·교육원 교재·교과서, 관보·보도자료 등)은 신청·송부 생략
         비밀기록물을 책자로 만든 것은 간행물이 아니라 비밀기록물 절차
       생산현황 통보서식 서식 7: 생산량·송부량은 부수가 아닌 종(철) 수, 생략대상 제외, 없으면 0
원칙:  점검과 수치 계산만 한다. 발간등록 신청·송부는 담당자가 국가기록포털에서 한다.
"""
import argparse
import collections
import os
import re
import sys

NUM = re.compile(r"^(\d{2})-(\d{7})-(\d{6})-(\d{2})$")
SKIP = ["수첩", "팸플릿", "팜플렛", "리플릿", "리플렛", "속보", "만화", "웹툰", "기관신문", "신문", "추록", "홍보", "소식지", "안내지", "브로셔",
        "브로슈어", "지도", "장학", "학습교재", "교재", "교과서", "관보", "구보", "군보", "시보", "보도자료"]
TITLE_SKIP = ["수첩", "팸플릿", "팜플렛", "리플릿", "리플렛", "웹툰", "기관신문", "추록", "소식지", "안내지", "브로셔", "브로슈어", "교과서", "관보", "보도자료", "홍보물"]
ALIAS = {"제목": ["제목", "간행물명"], "번호": ["발간등록번호"], "발행일": ["발행일", "발간일"], "형태": ["발간형태", "형태"],
         "유형": ["간행물유형", "유형"], "송부일": ["송부일", "송부일자"], "부수": ["송부부수"], "pdf": ["PDF송부", "전자파일송부", "전자파일"],
         "공개": ["공개구분"], "비밀": ["비밀구분"]}


def _d(v):
    import datetime as _dt
    t = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y%m%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--today", help="송부 기한 판정일(기본: 오늘)")
    ap.add_argument("-o", "--out", default="산출/간행물_점검.xlsx")
    a = ap.parse_args()
    import datetime as _dt
    _today = _dt.date.fromisoformat(a.today) if a.today else _dt.date.today()
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    items = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]
    by_no = collections.defaultdict(list)
    for n, x in items:
        if x.get("번호"):
            by_no[x["번호"]].append(x.get("제목", ""))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "제목", "발간등록번호", "발간주기", "생략대상 후보", "송부", "판정"])
    produced, sent = set(), set()
    for n, x in items:
        issues = []
        title, no = x.get("제목", ""), x.get("번호", "")
        if x.get("비밀", "") not in ("", "일반", "N", "해당없음"):
            ws.append([n, "(비밀 표시)", "", "", "", "", "간행물이 아니라 비밀기록물 절차로 관리"])
            continue
        kind = x.get("유형", "")
        skip = next((k for k in SKIP if k in kind), "") or next((k for k in TITLE_SKIP if k in title), "")
        cycle = ""
        m = NUM.match(no)
        if no and not m:
            issues.append(f"발간등록번호 형식 '{no}' — NN-NNNNNNN-NNNNNN-NN(예: 11-1311153-000040-10)")
        if m:
            c = int(m[4])
            cycle = "단행본(01)" if c == 1 else (f"연속간행물({m[4]})" if 2 <= c <= 15 else f"주기 코드 {m[4]} 확인")
            if c == 1 and len(by_no[no]) > 1:
                issues.append(f"단행본 번호가 {len(by_no[no])}건에 중복 — 단행본은 새 번호를 받는다")
            if c != 1 and len({re.sub(r'\s|\d{4}년?|제?\d+[호권집]', '', t) for t in by_no[no]}) > 1:
                issues.append("같은 연속간행물 번호에 제목이 다른 간행물 — 제목이 바뀌었으면 변경신청으로 새 번호")
        if not no and not skip:
            issues.append("발간등록번호 없음 — 인쇄 전에 신청했어야 함(생략대상이면 무시)")
        if no and skip:
            issues.append(f"생략대상 후보('{skip}')인데 번호를 받음 — 생략대상은 송부해도 영구보존 대상이 아닐 수 있음")
        send = ""
        if no:
            is_book = "전자" not in x.get("형태", "") or "책자" in x.get("형태", "")
            copies = int(re.sub(r"\D", "", x.get("부수", "")) or 0)
            pdf = x.get("pdf", "").upper() in ("Y", "O", "예", "완료")
            if is_book:
                send = f"책자 {copies}부·PDF {'있음' if pdf else '없음'}"
                if copies < 3:
                    issues.append(f"책자 {copies}부 송부 — 국가기록원에 3부")
                if not pdf:
                    issues.append("책자와 같은 PDF 1점을 함께 송부·업로드")
            else:
                send = f"전자파일 {'있음' if pdf else '없음'}"
                if not pdf:
                    issues.append("전자간행물은 전자파일 1점을 업로드")
        if no and not skip:
            pd_, sd_ = _d(x.get("발행일")), _d(x.get("송부일"))
            if pd_ and sd_ and (sd_ - pd_).days > 15:
                issues.append(f"발간 후 {(sd_ - pd_).days}일 만에 송부 — 발간 후 15일이 지나기 전까지(시행령 제55조④)")
            elif pd_ and not sd_ and (_today - pd_).days > 15:
                issues.append(f"발간 후 {(_today - pd_).days}일 지났는데 송부일 없음 — 15일 안 3부(관할 기록관·관할 영구기록물관리기관·중앙기록물관리기관)")
        year_ok = str(a.year) in x.get("발행일", "")
        if year_ok and m and not skip:
            produced.add(no)
            if (x.get("송부일") or "").strip():
                sent.add(no)
        ws.append([n, title, no, cycle, skip, send, " · ".join(issues) or "통과"])
    s7 = wb.create_sheet("서식 7 수치")
    s7.append([f"{a.year}년 간행물 생산량(종) ①", "간행물 송부량(종) ②", "기준"])
    produced, sent = len(produced), len(sent)
    s7.append([produced, sent, "발행일이 해당 연도·형식이 맞는 발간등록번호(같은 번호는 1종)·생략대상 제외, 송부량은 송부일이 있는 종 수(부수 아님)"])
    for i, w in enumerate((8, 40, 24, 16, 14, 20, 70)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    bad = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[6] != "통과")
    print(f"간행물 {len(items)}종 · 고칠 것 있는 종 {bad} · 서식 7 생산량 {produced} · 송부량 {sent} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
