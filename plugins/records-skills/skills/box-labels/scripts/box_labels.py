#!/usr/bin/env python3
"""보존상자 편성·표지: 기록물철 목록을 생산연도·보존기간·단위과제별로 보존상자에 나눠 담는 편성안을 만들고,
시행규칙 별표 6 「보존상자 표지」(가로 9㎝ = 항목 3㎝ + 내용 6㎝, 세로 7.5㎝: 상자번호 1.5·생산연도 1.5·생산기관 1.5·업무명 3.0㎝)를
A4 한 장에 6개씩 인쇄할 수 있는 PDF·HTML로 만든다. 철당 100매가 넘는 철은 분권 후보로 표시한다.

사용:  python box_labels.py 기록물철목록.xlsx --org "○○시 기획과" [--box-pages 1000] [--start 1] [--prefix 2025-]
         [--font C:/Windows/Fonts/malgun.ttf] -o 산출/보존상자
목록 열(머리글 이름으로 찾음): 철분류번호, 철제목, 생산연도, 보존기간, 단위과제명(업무명), 쪽수(매수), 처리과(생산기관), 권호수
결과: 보존상자_편성안.xlsx, 보존상자_표지.html(브라우저 인쇄 — 배율 100%, 여백 '없음'), 보존상자_표지.pdf(reportlab이 있을 때)
근거(law.go.kr 2026-11-20 확인):
  시행규칙 제9조③ 철당 100매 이내, 넘으면 2권 이상으로 나누고 같은 제목·분류번호에 권호수만 다르게
  제9조④ 완결 문서는 별표 4 보존용 표지 + 별표 5 보존상자(중성판지 800g/㎡ 이상·두께 1㎜ 이상)에 단위과제별로
  제9조⑤ 상자 측면에 별표 6 보존상자 표지
  제14조② <신설 2026. 9. 11.> 기록물철을 생산연도별·보존기간별로 구분해 보존상자에 넣음
  시행령 제32조⑤ 비전자기록물 이관은 단위과제별로 보존상자에 넣어 이관목록과 함께
원칙:  편성안과 표지만 만든다. 상자 한 개에 들어가는 분량(--box-pages)은 기관 상자 규격에 맞춰 정한다.
"""
import argparse
import glob
import html
import os
import re
import sys

ALIAS = {"철번호": ["철분류번호", "기록물철분류번호", "분류번호", "관리번호"], "제목": ["철제목", "기록물철제목", "제목"],
         "생산연도": ["생산연도", "생산년도"], "보존기간": ["보존기간"], "업무": ["단위과제명", "단위과제", "업무명"],
         "쪽수": ["쪽수", "매수", "면수"], "기관": ["처리과", "생산부서", "생산기관", "생산부서(처리과)명"], "권호": ["권호수", "권호"]}
ORDER = {"영구": 0, "준영구": 1, "30년": 2, "10년": 3, "5년": 4, "3년": 5, "1년": 6}
FONT_CANDIDATES = ["C:/Windows/Fonts/malgun.ttf", "/Library/Fonts/AppleGothic.ttf", "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
                   "/usr/share/fonts/truetype/nanum/NanumGothic.ttf", os.path.expanduser("~/Library/Fonts/NanumGothic.ttf")]


def pages(v):
    m = re.search(r"\d+", str(v or ""))
    return int(m[0]) if m else 0


def find_font(arg):
    if arg and os.path.exists(arg):
        return arg
    for p in FONT_CANDIDATES + glob.glob("/usr/share/fonts/**/NanumGothic*.ttf", recursive=True):
        if os.path.exists(p):
            return p
    return None


def make_pdf(path, labels, font_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    if font_path:
        from reportlab.pdfbase.ttfonts import TTFont
        pdfmetrics.registerFont(TTFont("KR", font_path))
        fname, embedded = "KR", True
    else:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
        fname, embedded = "HYGothic-Medium", False
    c = canvas.Canvas(path, pagesize=A4)
    W, H = A4
    lw, lh = 90 * mm, 75 * mm
    x0, y0 = (W - 2 * lw - 10 * mm) / 2, H - 15 * mm
    rows_h = [15 * mm, 15 * mm, 15 * mm, 30 * mm]
    keys = ["상자번호", "생산연도", "생산기관", "업무명"]
    for i, lab in enumerate(labels):
        k = i % 6
        if i and k == 0:
            c.showPage()
        col, row = k % 2, k // 2
        x = x0 + col * (lw + 10 * mm)
        y = y0 - row * (lh + 12 * mm)
        c.setLineWidth(1.2)
        c.rect(x, y - lh, lw, lh)
        c.line(x + 30 * mm, y, x + 30 * mm, y - lh)
        yy = y
        for h, key in zip(rows_h, keys):
            c.setFont(fname, 12)
            c.drawCentredString(x + 15 * mm, yy - h / 2 - 4, key)
            val = str(lab[key])
            size = 13 if key != "업무명" else 12
            c.setFont(fname, size)
            lines = wrap(val, 14 if key == "업무명" else 18)
            for j, ln in enumerate(lines[:3]):
                c.drawString(x + 33 * mm, yy - h / 2 - 4 + (len(lines[:3]) - 1) * 7 - j * 14, ln)
            yy -= h
            if key != "업무명":
                c.line(x, yy, x + lw, yy)
        c.setFont(fname, 7)
        c.drawString(x, y - lh - 9, f"시행규칙 별표 6 보존상자 표지 · {lab['비고']}")
    c.save()
    return embedded


def wrap(s, n):
    out, cur = [], ""
    for ch in s:
        cur += ch
        if len(cur) >= n:
            out.append(cur)
            cur = ""
    if cur:
        out.append(cur)
    return out or [""]


def make_html(path, labels):
    css = """@page{size:A4;margin:0}body{margin:0;font-family:'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR',sans-serif}
.page{width:210mm;height:297mm;display:grid;grid-template-columns:90mm 90mm;column-gap:10mm;row-gap:12mm;justify-content:center;
align-content:start;padding-top:15mm;box-sizing:border-box;page-break-after:always}
table{border-collapse:collapse;width:90mm;height:75mm;table-layout:fixed}td{border:1.2pt solid #000;padding:0 2mm;font-size:12pt}
td.k{width:30mm;text-align:center}tr.s{height:15mm}tr.b{height:30mm}.n{font-size:7pt;margin-top:1mm}"""
    parts = [f"<!DOCTYPE html><html lang='ko'><head><meta charset='utf-8'><title>보존상자 표지</title><style>{css}</style></head><body>"]
    for i in range(0, len(labels), 6):
        parts.append("<div class='page'>")
        for lab in labels[i:i + 6]:
            e = lambda k: html.escape(str(lab[k]))
            parts.append(f"<div><table><tr class='s'><td class='k'>상자번호</td><td>{e('상자번호')}</td></tr>"
                         f"<tr class='s'><td class='k'>생산연도</td><td>{e('생산연도')}</td></tr>"
                         f"<tr class='s'><td class='k'>생산기관</td><td>{e('생산기관')}</td></tr>"
                         f"<tr class='b'><td class='k'>업무명</td><td>{e('업무명')}</td></tr></table>"
                         f"<div class='n'>시행규칙 별표 6 보존상자 표지 · {e('비고')}</div></div>")
        parts.append("</div>")
    parts.append("</body></html>")
    open(path, "w", encoding="utf-8").write("".join(parts))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--org", default="", help="생산기관(처리과) 이름 — 목록에 처리과 열이 있으면 그것을 씀")
    ap.add_argument("--box-pages", type=int, default=1000, help="상자 하나에 넣을 최대 매수(기관 상자 규격에 맞게)")
    ap.add_argument("--start", type=int, default=1, help="상자 일련번호 시작")
    ap.add_argument("--prefix", default="", help="상자번호 앞에 붙일 글자(예: 2025-)")
    ap.add_argument("--font", help="한글 TrueType 글꼴(.ttf) 경로 — PDF에 넣음")
    ap.add_argument("-o", "--out", default="산출/보존상자")
    a = ap.parse_args()
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head and k not in idx:
                idx[k] = head.index(n)
    items = [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} | {"_row": n} for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]
    warn = []
    for x in items:
        p = pages(x.get("쪽수"))
        x["_p"] = p
        if p > 100:
            warn.append([x["_row"], x.get("철번호", ""), x.get("제목", ""), f"{p}매 — 철당 100매 이내(시행규칙 제9조③): 같은 제목·분류번호로 분권하고 권호수만 다르게"])
        if not x.get("생산연도") or not x.get("보존기간"):
            warn.append([x["_row"], x.get("철번호", ""), x.get("제목", ""), "생산연도·보존기간이 있어야 상자를 나눌 수 있음(제14조②)"])
    key = lambda x: (x.get("생산연도", ""), ORDER.get(re.sub(r"\s", "", x.get("보존기간", "")), 9), x.get("업무", ""), x.get("기관", "") or a.org)
    groups = {}
    for x in sorted(items, key=lambda x: (*key(x), x.get("철번호", ""))):
        groups.setdefault(key(x), []).append(x)
    boxes, seq = [], a.start
    for (yr, _, task, org), lst in groups.items():
        cur, used = [], 0
        for x in lst:
            if cur and used + x["_p"] > a.box_pages:
                boxes.append((yr, lst[0].get("보존기간", ""), task, org, cur, used))
                cur, used = [], 0
            cur.append(x)
            used += x["_p"]
        if cur:
            boxes.append((yr, lst[0].get("보존기간", ""), task, org, cur, used))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "편성안"
    ws.append(["상자번호", "생산연도", "보존기간", "업무명(단위과제)", "생산기관", "철 수", "매수 합계", "철분류번호(권호)", "철제목"])
    labels = []
    for yr, ret, task, org, lst, used in boxes:
        no = f"{a.prefix}{seq:03d}"
        seq += 1
        for j, x in enumerate(lst):
            ws.append([no if j == 0 else "", yr if j == 0 else "", ret if j == 0 else "", task if j == 0 else "", org if j == 0 else "",
                       len(lst) if j == 0 else "", used if j == 0 else "",
                       x.get("철번호", "") + (f" ({x['권호']})" if x.get("권호") else ""), x.get("제목", "")])
        labels.append({"상자번호": no, "생산연도": yr, "생산기관": org, "업무명": task or "(단위과제명 없음)", "비고": f"보존기간 {ret} · {len(lst)}철"})
        if any(x["_p"] == 0 for x in lst):
            warn.append(["", no, "", "쪽수 없는 철이 있어 매수 합계가 작게 잡힘"])
    chk = wb.create_sheet("확인할 것")
    chk.append(["원본 행", "철분류번호·상자번호", "철제목", "내용"])
    for w in warn:
        chk.append(w)
    if not warn:
        chk.append(["", "", "", "없음"])
    g = wb.create_sheet("기준")
    for line in [["상자 구분", "생산연도별·보존기간별(시행규칙 제14조②), 단위과제별(제9조④, 시행령 제32조⑤)"],
                 ["상자 규격", "별표 5 — 중성판지 800g/㎡ 이상, 두께 1㎜ 이상"], ["표지", "별표 6 — 상자 측면, 9㎝×7.5㎝"],
                 ["철당 분량", "100매 이내(제9조③)"], ["상자당 매수", f"--box-pages {a.box_pages}(기관 상자 규격에 맞게)"]]:
        g.append(line)
    for i, w in enumerate((12, 10, 10, 30, 20, 8, 10, 26, 40)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(a.out, exist_ok=True)
    xl = os.path.join(a.out, "보존상자_편성안.xlsx")
    wb.save(xl)
    make_html(os.path.join(a.out, "보존상자_표지.html"), labels)
    pdf_note = ""
    try:
        font = find_font(a.font)
        embedded = make_pdf(os.path.join(a.out, "보존상자_표지.pdf"), labels, font)
        pdf_note = " · PDF" + ("" if embedded else "(한글 글꼴 미포함 — --font로 .ttf를 주면 넣음)")
    except ImportError:
        pdf_note = " · PDF 없음(reportlab 미설치 — HTML을 인쇄)"
    print(f"{len(items)}철 → 상자 {len(boxes)}개 · 확인할 것 {len(warn)}{pdf_note} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
