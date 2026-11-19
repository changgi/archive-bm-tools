#!/usr/bin/env python3
"""이관목록 만들기: 등록정보 목록에서 시행규칙 별지 제2호서식 '이관목록'(철 단위·건 단위)을 만든다.

사용:  python make_transfer_list.py 등록정보.xlsx --org "가상시 행정과(1234567)" --year 2026 [--giver "주무관 홍가상"] [--taker "기록연구사 김가상"] -o 산출/이관목록.xlsx
목록 열(머리글 이름으로 찾음): 철분류번호, 철제목, 등록번호, 제목(건 제목), 쪽수, 기록물형태, 공개구분, 공개제한쪽, 비고, 비밀구분, 단위과제
근거:  시행규칙 제20조의2(비전자기록물 이관목록은 별지 제2호서식) <신설 2026. 9. 11.>, 시행령 제32조⑤(단위과제별로 보존상자에 넣어 이관목록과 제출)
서식:  1. 기록물철 단위 — 분류번호 · 제목 · 수량(건수, 쪽수*) · 기록물형태 · 비고
       2. 건 단위(개별관리) — 등록번호 · 제목 · 쪽수* · 기록물형태 · 기록물철 분류번호 · 공개구분 · 공개제한 쪽표시 · 비고
       * 쪽수는 종이기록물만 적는다.
원칙:  목록만 만든다. 인계·인수 서명은 사람이 한다. 비밀 표시 행은 넣지 않고 수만 알린다(비밀기록물 이관은 별도 절차, 시행령 제68조).
"""
import argparse
import collections
import os
import sys

ALIAS = {"철분류번호": ["철분류번호", "기록물철분류번호", "분류번호", "철번호"], "철제목": ["철제목", "기록물철제목"],
         "등록번호": ["등록번호", "생산등록번호", "접수등록번호"], "제목": ["제목", "건제목", "기록물건제목"],
         "쪽수": ["쪽수", "면수"], "형태": ["기록물형태", "형태", "전자구분"], "공개구분": ["공개구분", "공개여부"],
         "공개제한쪽": ["공개제한쪽", "공개제한 쪽표시", "비공개쪽"], "비고": ["비고"], "비밀": ["비밀구분", "비밀여부"],
         "단위과제": ["단위과제", "단위과제명"], "관리단위": ["관리단위", "개별관리"]}
PAPER = ("종이", "비전자", "일반문서류")


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


def is_paper(form):
    return any(p in form for p in PAPER) and "전자" != form


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--org", required=True, help="기관명(처리과기관코드)")
    ap.add_argument("--year", required=True, help="이관연도")
    ap.add_argument("--giver", default="", help="인계자 직급 및 성명")
    ap.add_argument("--taker", default="", help="인수자 직급 및 성명")
    ap.add_argument("-o", "--out", default="산출/이관목록.xlsx")
    a = ap.parse_args()
    rows, idx = read(a.listing)
    if "철분류번호" not in idx:
        print("입력 오류: 기록물철 분류번호 열이 없습니다.")
        return 2
    secret = [r for r in rows if r.get("비밀", "") not in ("", "일반", "N", "아니오", "해당없음")]
    rows = [r for r in rows if r not in secret]

    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = openpyxl.Workbook()

    def header(ws, title, cols):
        ws.append(["이관목록"])
        ws["A1"].font = Font(bold=True, size=14)
        ws.append([title])
        ws.append([f"기관명(처리과기관코드): {a.org}", "", "", f"인계자 직급 및 성명: {a.giver}"])
        ws.append([f"이관연도: {a.year}", "", "", f"인수자 직급 및 성명: {a.taker}"])
        ws.append(cols)
        for c in ws[ws.max_row]:
            c.font, c.border, c.alignment = Font(bold=True), box, Alignment(horizontal="center", wrap_text=True)

    w1 = wb.active
    w1.title = "1. 기록물철 단위"
    header(w1, "1. 기록물철 단위로 이관하는 기록물", ["기록물철 분류번호", "기록물철 제목", "수량(건수)", "수량(쪽수*)", "기록물형태", "비고"])
    files = collections.OrderedDict()
    for r in rows:
        if r.get("관리단위", "") in ("건", "개별", "Y"):
            continue
        f = files.setdefault(r["철분류번호"], {"제목": r.get("철제목", ""), "건": 0, "쪽": 0, "형태": set(), "비고": set(), "과제": r.get("단위과제", "")})
        f["건"] += 1
        f["형태"].add(r.get("형태", ""))
        if is_paper(r.get("형태", "")) and str(r.get("쪽수", "")).isdigit():
            f["쪽"] += int(r["쪽수"])
        if r.get("비고"):
            f["비고"].add(r["비고"])
    for no, f in sorted(files.items(), key=lambda kv: (kv[1]["과제"], kv[0])):
        forms = "·".join(sorted(x for x in f["형태"] if x)) or ""
        paper = any(is_paper(x) for x in f["형태"])
        w1.append([no, f["제목"], f["건"], f["쪽"] if paper else "", forms, "; ".join(sorted(f["비고"]))])
        for c in w1[w1.max_row]:
            c.border = box
    w1.append(["* 종이기록물의 경우에만 기재"])

    w2 = wb.create_sheet("2. 건 단위")
    header(w2, "2. 개별관리기록물(건) 단위로 이관하는 기록물",
           ["등록번호", "제목", "쪽수*", "기록물형태", "기록물철 분류번호", "공개구분", "공개제한 쪽표시", "비고"])
    n_item = 0
    for r in rows:
        if r.get("관리단위", "") not in ("건", "개별", "Y"):
            continue
        n_item += 1
        w2.append([r.get("등록번호", ""), r.get("제목", ""), r.get("쪽수", "") if is_paper(r.get("형태", "")) else "",
                   r.get("형태", ""), r["철분류번호"], r.get("공개구분", ""), r.get("공개제한쪽", ""), r.get("비고", "")])
        for c in w2[w2.max_row]:
            c.border = box
    w2.append(["* 종이기록물의 경우에만 기재"])
    for ws, widths in ((w1, (20, 40, 10, 10, 14, 24)), (w2, (18, 40, 8, 12, 20, 10, 14, 20))):
        for i, wdt in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = wdt
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"이관목록: 철 {len(files)}개 · 건 단위 {n_item}건 · 비밀 표시 {len(secret)}건 제외 → {a.out}  (별지 제2호서식)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
