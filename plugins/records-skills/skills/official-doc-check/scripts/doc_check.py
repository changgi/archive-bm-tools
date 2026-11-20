#!/usr/bin/env python3
"""공문서 표기 점검: 기안문·보고서(txt·docx·hwpx)의 날짜·시간·금액 표기, 항목 기호 순서, 붙임 표기, '끝' 표시, 등록번호 형식을
「행정업무의 운영 및 혁신에 관한 규정」 제7조와 같은 규정 시행규칙 제2조·제4조에 맞춰 보고 고쳐 쓸 값을 제안한다.
근거: 규정 [시행 2026. 10. 2.] 제7조④⑤, 시행규칙 [시행 2023. 6. 28.] 제2조①②·제4조④⑤ (law.go.kr 2026-11-20 확인),
      공공기록물법 시행규칙 별표 1(등록번호 '처리과명-연도별 일련번호')

사용:  python doc_check.py 문서.hwpx [문서2.docx ...] -o 산출/공문서점검.xlsx
규칙:
  날짜  숫자로, 연·월·일 글자 대신 온점 — 2026. 11. 20. (2026년 11월 20일, 2026-11-20, 2026.11.20 → 제안)
  시간  24시각제, 시·분 글자 대신 쌍점 — 14:30 (오후 2시 30분, 14시 30분 → 제안)
  금액  아라비아 숫자 다음 괄호에 한글 — 금113,560원(금일십일만삼천오백육십원)
  항목  1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮, 같은 단계는 오름차순·가나다순(□ ○ - ㆍ 등 특수기호는 허용)
  붙임  본문 내용 다음 줄 '붙임', 첨부물 명칭과 수량(1부), 둘 이상이면 1. 2.로 구분
  끝    본문(붙임) 마지막 글자에서 한 글자 띄우고 '끝.' — 표로 끝나면 표 아래 '끝.' 또는 '이하 빈칸'
원칙:  원본을 고치지 않는다. '특별한 사유가 있으면 다른 방법'도 허용되므로 결과는 제안이다. 비밀 표시가 있는 문서는 열지 않는다.
"""
import argparse
import os
import re
import sys
import zipfile

NUM = "영일이삼사오육칠팔구"
HANGUL_SEQ = "가나다라마바사아자차카타파하"
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"
CIRCLED_H = "㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻"
LEVELS = [  # (이름, 정규식, 값 → 순번)
    ("1.", re.compile(r"^(\d{1,2})\.\s"), lambda m: int(m[1])),
    ("가.", re.compile(r"^([가-하])\.\s"), lambda m: HANGUL_SEQ.find(m[1]) + 1),
    ("1)", re.compile(r"^(\d{1,2})\)\s"), lambda m: int(m[1])),
    ("가)", re.compile(r"^([가-하])\)\s"), lambda m: HANGUL_SEQ.find(m[1]) + 1),
    ("(1)", re.compile(r"^\((\d{1,2})\)\s"), lambda m: int(m[1])),
    ("(가)", re.compile(r"^\(([가-하])\)\s"), lambda m: HANGUL_SEQ.find(m[1]) + 1),
    ("①", re.compile(r"^([①-⑮])\s?"), lambda m: CIRCLED.find(m[1]) + 1),
    ("㉮", re.compile(r"^([㉮-㉻])\s?"), lambda m: CIRCLED_H.find(m[1]) + 1),
]
SECRET = re.compile(r"(Ⅰ|Ⅱ|Ⅲ|I{1,3})\s?급\s?비밀|대외비")


def hangul_amount(n: int) -> str:
    if n == 0:
        return "영"
    units, out, g = ["", "만", "억", "조"], [], 0
    while n:
        part, n = n % 10000, n // 10000
        if part:
            s = ""
            for d, u in zip(f"{part:04d}", ["천", "백", "십", ""]):
                if d != "0":
                    s += NUM[int(d)] + u
            out.append(s + units[g])
        g += 1
    return "".join(reversed(out))


def read_lines(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        import docx
        d = docx.Document(path)
        lines = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                lines.append(" | ".join(c.text for c in r.cells))
        return lines
    if ext == ".hwpx":
        lines = []
        with zipfile.ZipFile(path) as z:
            for n in sorted(x for x in z.namelist() if re.match(r"Contents/section\d+\.xml", x)):
                xml = z.read(n).decode("utf-8", "replace")
                for p in re.findall(r"<hp:p\b.*?</hp:p>", xml, re.S):
                    lines.append("".join(re.findall(r"<hp:t>([^<]*)</hp:t>", p)))
        return lines
    for enc in ("utf-8-sig", "cp949"):
        try:
            return open(path, encoding=enc).read().splitlines()
        except UnicodeDecodeError:
            continue
    return []


def check_lines(lines):
    out = []  # (줄, 구분, 내용, 제안)
    for i, raw in enumerate(lines, 1):
        s = raw.strip()
        if not s:
            continue
        for m in re.finditer(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", s):
            out.append((i, "날짜", m[0], f"{m[1]}. {int(m[2])}. {int(m[3])}."))
        for m in re.finditer(r"(?<![\d.])(\d{4})-(\d{1,2})-(\d{1,2})(?![\d-])", s):
            out.append((i, "날짜", m[0], f"{m[1]}. {int(m[2])}. {int(m[3])}."))
        for m in re.finditer(r"(?<![\d.])(\d{4})\.(\d{1,2})\.(\d{1,2})\.?(?!\d)", s):
            out.append((i, "날짜", m[0], f"{m[1]}. {int(m[2])}. {int(m[3])}.  (온점 다음 한 칸 띄움)"))
        for m in re.finditer(r"(오전|오후)?\s*(\d{1,2})\s*시\s*(?:(\d{1,2})\s*분)?", s):
            if re.match(r"\d", s[m.end():m.end() + 1] or "x") or "시간" in s[m.start():m.end() + 2]:
                continue
            h = int(m[2]) + (12 if m[1] == "오후" and int(m[2]) < 12 else 0)
            out.append((i, "시간", m[0].strip(), f"{h:02d}:{int(m[3] or 0):02d}"))
        for m in re.finditer(r"(금\s*)?(\d{1,3}(?:,\d{3})+|\d{4,})\s*원(?!\s*\()", s):
            n = int(m[2].replace(",", ""))
            out.append((i, "금액", m[0], f"금{n:,}원(금{hangul_amount(n)}원)"))
        for m in re.finditer(r"금\s*([\d,]+)\s*원\s*\(\s*금?([가-힣]+)원\s*\)", s):
            n = int(m[1].replace(",", ""))
            if m[2].replace("금", "") != hangul_amount(n):
                out.append((i, "금액", m[0], f"한글 금액이 숫자와 다름 → 금{n:,}원(금{hangul_amount(n)}원)"))
    # 항목 기호 순서
    stack = []  # [(단계, 마지막 순번)]
    for i, raw in enumerate(lines, 1):
        s = raw.strip()
        if re.match(r"^붙\s*임", s):
            break  # 붙임 목록의 번호는 따로 본다
        lv = next(((k, f(m)) for k, (name, rx, f) in enumerate(LEVELS) for m in [rx.match(s)] if m), None)
        if not lv:
            continue
        k, v = lv
        while stack and stack[-1][0] > k:
            stack.pop()
        if stack and stack[-1][0] == k:
            if v != stack[-1][1] + 1:
                out.append((i, "항목 순서", s[:20], f"{LEVELS[k][0]} 단계 {stack[-1][1]} 다음은 {stack[-1][1] + 1}번째여야 함"))
            stack[-1] = (k, v)
        else:
            if stack and k > stack[-1][0] + 1:
                out.append((i, "항목 단계", s[:20], f"{LEVELS[stack[-1][0]][0]} 다음 하위 항목은 {LEVELS[stack[-1][0] + 1][0]} 형태"))
            elif not stack and k > 0 and v and any(LEVELS[0][1].match(l.strip()) for l in lines):
                pass
            if v != 1:
                out.append((i, "항목 순서", s[:20], f"{LEVELS[k][0]} 단계는 처음 항목부터(1번째) 시작"))
            stack.append((k, v))
    # 붙임
    idx = next((i for i, l in enumerate(lines) if re.match(r"^\s*붙\s*임", l)), None)
    if idx is not None:
        att = [lines[idx].strip()] + [l.strip() for l in lines[idx + 1:idx + 8] if l.strip()]
        first = re.sub(r"^붙\s*임\s*", "", att[0])
        items = [first] if first else []
        for l in att[1:]:
            if re.match(r"^\d+\.\s", l):
                items.append(l)
            else:
                break
        if len(items) > 1 and not all(re.match(r"^\d+\.\s", x) for x in items):
            out.append((idx + 1, "붙임", att[0][:30], "첨부물이 둘 이상이면 1. 2.로 구분"))
        for x in items:
            if not re.search(r"\d+\s*(부|매|권|건|개|점|장)", x):
                out.append((idx + 1, "붙임", x[:30], "첨부물 명칭 다음에 수량(예: 1부)"))
    # 끝 표시
    body = [l for l in lines if l.strip()]
    tail = next((l for l in reversed(body[-12:]) if re.search(r"끝\.?\s*$|이하\s*빈칸", l)), None)
    if not tail:
        out.append((len(lines), "끝 표시", "", "본문(붙임) 마지막 글자에서 한 글자 띄우고 '끝.' — 표로 끝나면 표 아래 '끝.' 또는 '이하 빈칸'"))
    else:
        m = re.search(r"(\S)(\s*)끝\.?\s*$", tail)
        if m and len(m[2].replace("\t", " ")) != 1 and not re.match(r"^\s*끝", tail):
            out.append((lines.index(tail) + 1, "끝 표시", tail.strip()[-15:], "마지막 글자와 '끝' 사이는 한 칸(두 칸·붙여 쓰기 아님)"))
        if re.search(r"끝\s*$", tail):
            out.append((lines.index(tail) + 1, "끝 표시", tail.strip()[-10:], "'끝' 다음 온점: 끝."))
    # 등록번호
    for i, l in enumerate(lines, 1):
        for m in re.finditer(r"(?:시행|접수)\s+([가-힣A-Za-z0-9]+)\s*-\s*(\d+)", l):
            if len(m[2]) > 6:
                out.append((i, "등록번호", m[0], "처리과명-연도별 일련번호 형식인지 확인(공공기록물법 시행규칙 별표 1)"))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", default="산출/공문서점검.xlsx")
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검 결과"
    ws.append(["파일", "줄", "구분", "찾은 표기", "제안"])
    total, skipped = 0, []
    for f in a.files:
        if re.search(r"(^|[\\/])비밀([\\/]|$)", f):
            skipped.append(f)
            continue
        lines = read_lines(f)
        if any(SECRET.search(l) for l in lines[:40]):
            skipped.append(f)
            continue
        for i, k, found, sug in check_lines(lines):
            ws.append([os.path.basename(f), i, k, found, sug])
            total += 1
    g = wb.create_sheet("기준")
    for line in [("날짜", "2026. 11. 20. — 연·월·일 글자 대신 온점(규정 제7조⑤)"), ("시간", "14:30 — 24시각제, 시·분 글자 대신 쌍점(제7조⑤)"),
                 ("금액", "금113,560원(금일십일만삼천오백육십원)(시행규칙 제2조②)"), ("항목", "1. 가. 1) 가) (1) (가) ① ㉮, 필요하면 □ ○ - ㆍ(제2조①)"),
                 ("붙임", "본문 다음 줄 '붙임' + 명칭·수량, 둘 이상이면 항목 구분(제4조④)"), ("끝", "마지막 글자에서 한 글자 띄우고 '끝.', 표는 표 아래 '끝.' 또는 '이하 빈칸'(제4조⑤)"),
                 ("숫자·용지", "아라비아 숫자(제7조④), A4 210×297㎜(제7조⑥)"), ("예외", "특별한 사유가 있으면 다른 방법으로 표시할 수 있음(제7조⑤ 단서) — 결과는 제안")]:
        g.append(list(line))
    if skipped:
        g.append(["열지 않음(비밀)", ", ".join(os.path.basename(x) for x in skipped)])
    for sh, widths in ((ws, (24, 6, 10, 30, 60)), (g, (14, 90))):
        sh["A1"].font = Font(bold=True)
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(a.files) - len(skipped)}개 문서 · 제안 {total}" + (f" · 비밀로 건너뜀 {len(skipped)}" if skipped else "") + f" → {a.out}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
