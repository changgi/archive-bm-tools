#!/usr/bin/env python3
"""회의록 점검: 회의록 파일마다 필수 구성 요소 8가지가 들어 있는지 보고, 생산의무 회의 목록과 대조해 회의록이 없는 회의를 찾는다.

사용:  python minutes_check.py 회의록폴더/ [--meetings 생산의무회의목록.xlsx] -o 산출/회의록_점검.xlsx
대상:  .txt .md .docx .hwpx .pdf(텍스트층)
근거:  공공기록물법 제17조②, 시행령 제18조(회의록의 작성·관리, 대상 회의 1~9호)
       국가기록원 「2026년 기록물관리 지침」 회의록 구성 요소: 회의명, 회의 개최 부서(기관), 일시 및 장소, 참석자·배석자 명단,
       회의 진행순서, 상정 안건, 발언 요지, 결정사항 및 표결내용. (기록관리 평가지표 '회의록 관리의 적절성'도 같은 항목을 본다)
       (특수)기록관은 생산의무 회의 목록을 확보하고 회의별 회의록 작성 현황을 점검한다(같은 지침).
원칙:  글자로 항목 이름이나 그에 해당하는 표시를 찾을 뿐이다. 내용이 충실한지는 사람이 본다. 회의록 본문은 결과에 옮기지 않는다.
"""
import argparse
import os
import re
import sys
import zipfile

ELEMENTS = {
    "회의명": [r"회\s*의\s*명", r"회의\s*명칭", r"제\s*\d+\s*차\s*.*회의"],
    "개최 부서(기관)": [r"주관\s*(부서|기관)", r"개최\s*(부서|기관)", r"담당\s*부서"],
    "일시 및 장소": [r"일\s*시", r"장\s*소"],
    "참석자·배석자 명단": [r"참\s*석\s*자", r"출석\s*(위원|자)", r"배\s*석\s*자"],
    "회의 진행순서": [r"진행\s*순서", r"순\s*서", r"개\s*회", r"폐\s*회"],
    "상정 안건": [r"안\s*건", r"의\s*안", r"상정"],
    "발언 요지": [r"발언", r"위원\s*[:：]", r"위원장\s*[:：]", r"의견"],
    "결정사항 및 표결내용": [r"결정\s*사항", r"의\s*결", r"표\s*결", r"원안\s*가결", r"가결", r"부결"],
}


def text_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md"):
        return open(path, encoding="utf-8", errors="replace").read()
    if ext == ".docx":
        with zipfile.ZipFile(path) as z:
            return re.sub(r"<[^>]+>", " ", z.read("word/document.xml").decode("utf-8", "replace"))
    if ext == ".hwpx":
        with zipfile.ZipFile(path) as z:
            return " ".join(re.sub(r"<[^>]+>", " ", z.read(n).decode("utf-8", "replace")) for n in sorted(z.namelist()) if n.startswith("Contents/section"))
    if ext == ".pdf":
        from pypdf import PdfReader
        return "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    return None


def norm(s):
    return re.sub(r"[\s()（）·・\-_]", "", str(s or ""))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--meetings", help="생산의무 회의 목록 xlsx(열: 회의명[, 개최일, 주관부서])")
    ap.add_argument("-o", "--out", default="산출/회의록_점검.xlsx")
    a = ap.parse_args()
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "구성 요소"
    ws.append(["파일"] + list(ELEMENTS) + ["빠진 요소 수", "판정"])
    found_titles = []
    n_files = n_bad = 0
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            p = os.path.join(root, fn)
            try:
                t = text_of(p)
            except Exception as e:
                ws.append([fn] + [""] * len(ELEMENTS) + ["", f"읽기 실패: {str(e)[:40]}"])
                continue
            if t is None:
                continue
            n_files += 1
            marks = ["○" if any(re.search(rx, t) for rx in rxs) else "✕" for rxs in ELEMENTS.values()]
            miss = marks.count("✕")
            n_bad += bool(miss)
            ws.append([fn] + marks + [miss, "구성 요소 갖춤" if not miss else "빠진 요소 확인"])
            found_titles.append(norm(fn) + " " + norm(t[:300]))
    m_rows = []
    if a.meetings:
        mw = openpyxl.load_workbook(a.meetings, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in mw.iter_rows(values_only=True)]
        head = rows[0]
        ni = head.index("회의명") if "회의명" in head else 0
        di = head.index("개최일") if "개최일" in head else None
        for r in rows[1:]:
            if not any(r):
                continue
            name = r[ni]
            key = norm(name)
            date = r[di] if di is not None and di < len(r) else ""
            dkey = re.sub(r"\D", "", date)[:8]
            hit = any(key and key in ft and (not dkey or dkey in re.sub(r"\D", "", ft)) for ft in found_titles)
            m_rows.append([name, date, "회의록 있음" if hit else "회의록 못 찾음 — 생산 여부 확인"])
        w2 = wb.create_sheet("생산의무 회의 대조")
        w2.append(["회의명", "개최일", "판정"])
        for r in m_rows:
            w2.append(r)
    w3 = wb.create_sheet("기준")
    for line in [["근거", "공공기록물법 제17조②, 시행령 제18조, 국가기록원 「2026년 기록물관리 지침」 회의록 구성 요소"],
                 ["주의", "글자 표시만 찾는다. 발언 요지가 충실한지, 표결 결과가 맞는지는 사람이 확인한다."]]:
        w3.append(line)
    for i, w in enumerate([30] + [11] * len(ELEMENTS) + [10, 16]):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    miss_m = sum(1 for r in m_rows if "못 찾음" in r[2])
    print(f"회의록 {n_files}건 · 요소 빠짐 {n_bad}건" + (f" · 생산의무 회의 {len(m_rows)} 중 회의록 못 찾음 {miss_m}" if a.meetings else "") + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
