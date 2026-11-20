#!/usr/bin/env python3
"""개인정보 탐지: 주민등록번호·외국인등록번호·전화번호·이메일·카드번호·계좌번호(문맥)·여권번호·운전면허번호·사업자등록번호(검증식)·
법인등록번호(문맥)·군번(문맥)·사건번호 후보를 찾는다. 추가 유형에서 같은 숫자 5자리 이상(11111·00000)은 더미로 보고 뺀다.
근거(추가 유형): 국가기록원 「전자기록물 공개재분류를 위한 비공개정보 필터링 및 마스킹 기술 적용방안 연구」 — 제6호 정규식 10종
(이메일·여권·주민·외국인·사업자·군번·운전면허·카드·주소·소송번호). 주소·인명은 정규식으로 잡지 않는다(연구도 주소DB·기계학습을 씀).

사용:  python pii_scan.py 입력(파일 또는 폴더) -o 산출/개인정보_탐지.csv
지원:  .txt .csv .tsv .md .xlsx .pdf(텍스트층) .hwpx(본문 XML)
원칙:  결과에는 값을 가린 형태만 적는다(원문을 다시 퍼뜨리지 않기 위해).
한계:  2020년 10월 이후 발급 주민등록번호는 뒷자리 검증식이 맞지 않을 수 있어,
       검증식이 안 맞아도 '후보'로 남긴다. 탐지 0건이 개인정보 0건을 뜻하지는 않는다.
"""
import argparse
import csv
import os
import re
import sys
import zipfile

PATTERNS = {
    "주민·외국인등록번호": re.compile(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\s?[-–]?\s?([1-8])(\d{6})(?!\d)"),
    "휴대전화": re.compile(r"(?<!\d)01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}(?!\d)"),
    "일반전화": re.compile(r"(?<!\d)0(2|[3-6][1-5])[-.)\s]\d{3,4}[-.\s]\d{4}(?!\d)"),
    "이메일": re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"),
    "카드번호": re.compile(r"(?<!\d)(?:\d{4}[-\s]?){3}\d{4}(?!\d)"),
    "계좌번호(문맥)": re.compile(r"(계좌|입금|은행|예금주)[^\n\d]{0,12}(\d{2,6}[-\s]\d{2,6}[-\s]\d{2,8}(?:[-\s]\d{1,4})?)"),
    # 아래는 국가기록원 「전자기록물 공개재분류를 위한 비공개정보 필터링 및 마스킹 기술 적용방안 연구」의 제6호 정규식 10종에 맞춰 더한 것
    "여권번호": re.compile(r"(?<![A-Za-z0-9])(?:[MSRODG]\d{8}|[MSRODG]\d{3}[A-Z]\d{4})(?![A-Za-z0-9])"),
    "운전면허번호": re.compile(r"(?<!\d)(?:1[1-9]|2[0-8])[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{2}(?!\d)"),
    "사업자등록번호": re.compile(r"(?<!\d)\d{3}-\d{2}-\d{5}(?!\d)"),
    "법인등록번호(문맥)": re.compile(r"(법인등록번호|법인번호)[^\n\d]{0,6}(\d{6}\s?-\s?\d{7})"),
    "군번(문맥)": re.compile(r"(군번)[^\n\d]{0,6}(\d{2}-\d{5,8})"),
    "사건번호": re.compile(r"(?<!\d)(?:19|20)\d{2}\s?(?:가합|가단|가소|나|다|라|머|드|느|노|도|고합|고단|고정|구합|구단|누|두|호|즈|카|타|하)\s?\d{2,6}(?!\d)"),
}
NEW_KINDS = {"여권번호", "운전면허번호", "사업자등록번호", "법인등록번호(문맥)", "군번(문맥)", "사건번호"}
CONTEXT_KINDS = {"계좌번호(문맥)", "법인등록번호(문맥)", "군번(문맥)"}
DUMMY = re.compile(r"(\d)\1{4,}")  # 같은 숫자 5자리 이상(11111·00000)은 예시·더미로 보고 뺀다(같은 연구)


def bizno_ok(d: str) -> bool:
    w = [1, 3, 7, 1, 3, 7, 1, 3, 5]
    s = sum(int(x) * k for x, k in zip(d[:9], w)) + int(d[8]) * 5 // 10
    return (10 - s % 10) % 10 == int(d[9])


def rrn_checksum_ok(digits: str) -> bool:
    w = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5]
    s = sum(int(d) * k for d, k in zip(digits[:12], w))
    return (11 - s % 11) % 10 == int(digits[12])


def luhn_ok(num: str) -> bool:
    d = [int(c) for c in num][::-1]
    total = sum(d[0::2]) + sum(sum(divmod(x * 2, 10)) for x in d[1::2])
    return total % 10 == 0


def mask(value: str) -> str:
    keep = 2 if len(value) > 6 else 1
    return value[:keep] + "*" * max(len(value) - keep * 2, 1) + value[-keep:]


def mask_all(ctx: str) -> str:
    """문맥에 함께 걸린 다른 개인정보도 가린다(옆 값이 문맥으로 새지 않게)."""
    for kind, rx in PATTERNS.items():
        def rep(m, kind=kind):
            raw = m.group(2) if kind in CONTEXT_KINDS else m.group(0)
            return m.group(0).replace(raw, mask(raw))
        ctx = rx.sub(rep, ctx)
    return re.sub(r"\d{5,}", lambda m: m.group(0)[:1] + "*" * (len(m.group(0)) - 2) + m.group(0)[-1:], ctx)


def scan_text(text: str):
    """(유형, 가린 값, 검증, 앞뒤 문맥) 목록을 돌려준다."""
    found = []
    for kind, rx in PATTERNS.items():
        for m in rx.finditer(text):
            raw = m.group(2) if kind in CONTEXT_KINDS else m.group(0)
            digits = re.sub(r"\D", "", raw)
            if kind in NEW_KINDS and DUMMY.search(digits):
                continue
            check = ""
            if kind == "사업자등록번호":
                check = "검증식 일치" if bizno_ok(digits) else "검증식 불일치 — 다른 번호일 수 있음"
            if kind == "주민·외국인등록번호":
                check = "검증식 일치" if rrn_checksum_ok(digits) else "검증식 불일치(2020.10 이후 번호일 수 있음)"
            elif kind == "카드번호":
                if not luhn_ok(digits):
                    continue  # 카드번호는 Luhn이 표준이라 불일치는 버린다
                check = "Luhn 일치"
            ctx = text[max(0, m.start() - 12): m.end() + 12].replace("\n", " ")
            ctx = mask_all(ctx.replace(raw, mask(raw)))
            found.append((kind, mask(raw), check, ctx))
    corp = {v for k, v, _, _ in found if k == "법인등록번호(문맥)"}
    return [f for f in found if not (f[0] == "주민·외국인등록번호" and f[1] in corp)]  # 같은 값이 법인등록번호로 잡혔으면 하나만


def read_units(path: str):
    """(위치, 텍스트) 단위로 읽는다. 위치는 행·셀·쪽 번호."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md"):
        with open(path, encoding="utf-8", errors="replace") as f:
            for i, line in enumerate(f, 1):
                yield f"{i}행", line
    elif ext in (".csv", ".tsv"):
        with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
            for i, row in enumerate(csv.reader(f, delimiter="\t" if ext == ".tsv" else ","), 1):
                yield f"{i}행", " | ".join(row)
    elif ext == ".xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for r, row in enumerate(ws.iter_rows(values_only=True), 1):
                text = " | ".join("" if v is None else str(v) for v in row)
                if text.strip(" |"):
                    yield f"{ws.title}!{r}행", text
    elif ext == ".pdf":
        from pypdf import PdfReader
        for i, page in enumerate(PdfReader(path).pages, 1):
            yield f"{i}쪽", page.extract_text() or ""
    elif ext == ".hwpx":
        with zipfile.ZipFile(path) as z:
            for name in sorted(n for n in z.namelist() if n.startswith("Contents/section")):
                xml = z.read(name).decode("utf-8", "replace")
                yield name, re.sub(r"<[^>]+>", " ", xml)


def iter_files(target: str):
    if os.path.isdir(target):
        for root, _, files in os.walk(target):
            for fn in sorted(files):
                yield os.path.join(root, fn)
    else:
        yield target


def main() -> int:
    ap = argparse.ArgumentParser(description="개인정보 후보 탐지 (값은 가려서 기록)")
    ap.add_argument("target")
    ap.add_argument("-o", "--out", default="산출/개인정보_탐지.csv")
    a = ap.parse_args()
    rows = []
    for path in iter_files(a.target):
        try:
            for where, text in read_units(path):
                for kind, masked, check, ctx in scan_text(text):
                    rows.append([path, where, kind, masked, check, ctx])
        except Exception as e:  # 읽지 못한 파일도 결과에 남긴다
            rows.append([path, "", "읽기 실패", "", str(e)[:100], ""])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["파일", "위치", "유형", "가린 값", "검증", "문맥(가림)"])
        w.writerows(rows)
    hits = sum(1 for r in rows if r[2] != "읽기 실패")
    print(f"탐지 {hits}건 · 읽기 실패 {len(rows) - hits}건 → {a.out}")
    return 1 if hits else 0


if __name__ == "__main__":
    sys.exit(main())
