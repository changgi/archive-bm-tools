#!/usr/bin/env python3
"""메타데이터 후보 뽑기: 문서 본문에서 제목·생산기관·생산일자·문서번호·수신자 후보를 찾아, 찾은 위치(근거)와 함께 표로 만든다.

사용:  python extract_candidates.py 문서폴더/ [--existing 목록.xlsx --key 파일명] -o 산출/메타데이터_후보.xlsx
대상:  .txt .md .docx .hwpx .pdf(텍스트층)
방식:  공문서 서식(행정업무의 운영 및 혁신에 관한 규정의 기안문 구성)의 줄을 먼저 읽는다.
         '제목'·'수신'·'시행 ○○과-123 (2026. 3. 2.)'·'접수'·발신명의('○○시장' 등)
       서식이 없으면 첫 의미 있는 줄을 제목 후보로, 본문의 날짜·문서번호를 후보로 둔다.
       --existing을 주면 목록의 빈칸만 '채울 후보'로 표시하고, 이미 있는 값과 다르면 '불일치'로 표시한다.
원칙:  값은 '후보'다. 모든 후보에 근거(파일·쪽·줄)를 붙인다. 기록관리시스템에 직접 넣지 않는다.
       Claude가 후보를 검토해 값을 고를 때도 근거 열을 채워 records-qa-gate의 신뢰성 점검을 통과하게 한다.
"""
import argparse
import os
import re
import sys
import zipfile

DATE_RX = re.compile(r"(?<!\d)((?:19|20)\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})\s*[.일]?")
DOCNO_RX = re.compile(r"([가-힣A-Za-z0-9()·]{2,}-\d{1,6}(?:-\d{1,6})?)")
ORG_SIGN_RX = re.compile(r"^\s*([가-힣]{2,}(?:특별시|광역시|특별자치시|특별자치도|도|시|군|구)?(?:시장|군수|구청장|도지사|교육감|장관|청장|처장|위원장|원장|이사장|총장|공사사장|대표))\s*$")
FIELD_RX = {
    "제목": re.compile(r"^\s*제\s*목\s*[:：]?\s*(.+)$"),
    "수신": re.compile(r"^\s*수\s*신\s*(?:자)?\s*[:：]?\s*(.+)$"),
    "시행": re.compile(r"^\s*시\s*행\s*[:：]?\s*(.+)$"),
    "접수": re.compile(r"^\s*접\s*수\s*[:：]?\s*(.+)$"),
}


def lines_of(path):
    """(위치, 줄) 목록."""
    ext = os.path.splitext(path)[1].lower()
    out = []
    if ext in (".txt", ".md"):
        with open(path, encoding="utf-8", errors="replace") as f:
            out = [(f"{i}줄", l.rstrip("\n")) for i, l in enumerate(f, 1)]
    elif ext == ".docx":
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "replace")
        paras = [re.sub(r"<[^>]+>", "", p) for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)]
        out = [(f"{i}문단", p) for i, p in enumerate(paras, 1)]
    elif ext == ".hwpx":
        with zipfile.ZipFile(path) as z:
            names = sorted(n for n in z.namelist() if n.startswith("Contents/section"))
            i = 0
            for n in names:
                xml = z.read(n).decode("utf-8", "replace")
                for p in re.findall(r"<hp:p[ >].*?</hp:p>", xml, re.S):
                    i += 1
                    out.append((f"{i}문단", re.sub(r"<[^>]+>", "", p)))
    elif ext == ".pdf":
        from pypdf import PdfReader
        for pn, page in enumerate(PdfReader(path).pages, 1):
            for ln, l in enumerate((page.extract_text() or "").splitlines(), 1):
                out.append((f"{pn}쪽 {ln}줄", l))
    return [(w, l.strip()) for w, l in out if l.strip()]


def norm_date(m):
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{y:04d}-{mo:02d}-{d:02d}" if 1 <= mo <= 12 and 1 <= d <= 31 else None


def candidates(path):
    L = lines_of(path)
    c = {k: [] for k in ("제목", "생산기관", "생산일자", "문서번호", "수신자")}
    for where, line in L:
        for field, rx in FIELD_RX.items():
            m = rx.match(line)
            if not m:
                continue
            val = m.group(1).strip()
            if field == "제목":
                c["제목"].append((val, where, "서식 '제목' 줄"))
            elif field == "수신":
                c["수신자"].append((val, where, "서식 '수신' 줄"))
            elif field in ("시행", "접수"):
                dn = DOCNO_RX.search(val)
                if dn:
                    c["문서번호"].append((dn.group(1), where, f"서식 '{field}' 줄"))
                dm = DATE_RX.search(val)
                if dm and norm_date(dm):
                    c["생산일자"].append((norm_date(dm), where, f"서식 '{field}' 줄"))
        sm = ORG_SIGN_RX.match(line)
        if sm:
            c["생산기관"].append((sm.group(1), where, "발신명의 줄"))
    if not c["제목"] and L:
        first = next(((w, l) for w, l in L if len(l) >= 4 and not DATE_RX.search(l)), L[0])
        c["제목"].append((first[1][:120], first[0], "첫 의미 있는 줄(서식 없음)"))
    if not c["생산일자"]:
        for where, line in L:
            for dm in DATE_RX.finditer(line):
                if norm_date(dm):
                    c["생산일자"].append((norm_date(dm), where, "본문 날짜"))
        c["생산일자"] = c["생산일자"][:3]
    if not c["문서번호"]:
        for where, line in L:
            for dn in DOCNO_RX.finditer(line):
                if re.search(r"\d{4}-\d{2}-\d{2}", dn.group(1)) is None:
                    c["문서번호"].append((dn.group(1), where, "본문 번호 형태"))
        c["문서번호"] = c["문서번호"][:3]
    return c, len(L)


def read_existing(path, key):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    head = rows[0]
    return {r[head.index(key)]: dict(zip(head, r)) for r in rows[1:] if key in head and r[head.index(key)]}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--existing")
    ap.add_argument("--key", default="파일명")
    ap.add_argument("-o", "--out", default="산출/메타데이터_후보.xlsx")
    a = ap.parse_args()
    existing = read_existing(a.existing, a.key) if a.existing else {}
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "후보"
    ws.append(["파일", "요소", "후보 값", "근거 위치", "근거 종류", "기존 값", "상태"])
    n_files = n_fill = n_diff = 0
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            p = os.path.join(root, fn)
            if os.path.splitext(fn)[1].lower() not in (".txt", ".md", ".docx", ".hwpx", ".pdf"):
                continue
            n_files += 1
            try:
                cand, n_lines = candidates(p)
            except Exception as e:
                ws.append([fn, "", "", "", f"읽기 실패: {str(e)[:60]}", "", "확인 필요"])
                continue
            ex = existing.get(fn, {})
            for field, items in cand.items():
                old = ex.get(field, "")
                if not items:
                    ws.append([fn, field, "", "", "후보 없음", old, "사람이 입력" if not old else "기존 값 유지"])
                    continue
                for val, where, kind in items:
                    if not a.existing:
                        state = "후보"
                    elif not old:
                        state = "채울 후보"
                        n_fill += 1
                    elif old.replace(" ", "") == val.replace(" ", ""):
                        state = "일치"
                    else:
                        state = "불일치 — 확인"
                        n_diff += 1
                    ws.append([fn, field, val, where, kind, old, state])
    for col, w in zip("ABCDEFG", (24, 10, 44, 12, 22, 24, 14)):
        ws.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"파일 {n_files}개 · 채울 후보 {n_fill} · 불일치 {n_diff} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
