#!/usr/bin/env python3
"""폐기 심의자료 만들기: 보존기간이 끝난 기록물로 기록물평가심의서·생산부서 의견조회표·심의 안건(HWPX)을 만든다.

사용:  python build_dossier.py 산출/보존기간_검토.xlsx [--year 2026] [--include-next] -o 산출/평가심의
입력:  retention-review 결과(시트 '만료 예정') 또는 관리번호·제목·보존기간·만료일 열이 있는 목록
산출:  평가심의서.xlsx        — 심의 대상, 생산부서·전문요원·심의 의견 칸(선택 목록), 부서별 건수
       의견조회_<부서>.xlsx    — 생산부서별로 나눈 의견조회표 (생산부서 열이 있을 때)
       심의안건.hwpx          — python-hwpx가 설치돼 있으면 안건 개요와 목록 (한글에서 열어 확인할 것)
절차:  생산부서 의견조회 → 기록물관리 전문요원 심사 → 기록물평가심의회 심의 (NAK 5-1:2014 기록물 평가·폐기 절차)
원칙:  의견 칸은 비워 둔다. '검토 신호'는 보류·재책정을 검토해 볼 단서일 뿐 판단이 아니다. 폐기 집행은 하지 않는다.
"""
import argparse
import datetime as dt
import os
import re
import sys

KST = dt.timezone(dt.timedelta(hours=9))
SIGNALS = {
    "소송·분쟁": r"소송|소청|분쟁|행정심판|이의신청|재판",
    "감사·징계": r"감사|징계|조사",
    "재산·권리": r"계약|보상|토지|부동산|재산|채권|채무|등기|인허가|허가|면허",
    "시설·공사": r"준공|공사|설계|시설물",
    "재난·사고": r"재난|재해|사고|피해",
    "역사적 가치": r"최초|기념|역사|개청|창립|백서",
}
OPINIONS = '"폐기,보존기간 재책정,보류"'
COLS = ["번호", "관리번호", "제목", "생산부서", "단위과제", "보존기간", "종료연도", "만료일", "검토 신호",
        "생산부서 의견", "생산부서 사유", "전문요원 심사 의견", "전문요원 사유", "심의 결과", "근거"]
ALIAS = {"관리번호": ["관리번호", "철관리번호", "기록물건등록번호"], "제목": ["제목", "기록물철제목", "기록물건제목"],
         "생산부서": ["생산부서", "처리과", "생산기관", "부서"], "단위과제": ["단위과제", "단위과제명", "단위과제코드"],
         "보존기간": ["보존기간"], "종료연도": ["종료연도", "종료년도", "생산년도"], "만료일": ["만료일"], "분류": ["분류"]}


def read(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["만료 예정"] if "만료 예정" in wb.sheetnames else wb.worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    h = next((i for i, r in enumerate(rows[:20]) if any(c in r for c in ALIAS["관리번호"])), 0)
    head = rows[h]
    get = lambda r, k: next((r[head.index(c)] for c in ALIAS[k] if c in head and head.index(c) < len(r)), "")
    return [{k: get(r, k) for k in ALIAS} for r in rows[h + 1:] if any(r)], any(c in head for c in ALIAS["생산부서"])


def signals(title):
    return " · ".join(k for k, rx in SIGNALS.items() if re.search(rx, title or ""))


def write_xlsx(path, rows, title):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "심의 대상"
    ws.append([title])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(COLS)
    for c in ws[2]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="EEF1F4")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, r in enumerate(rows, 1):
        ws.append([i, r["관리번호"], r["제목"], r["생산부서"], r["단위과제"], r["보존기간"], r["종료연도"], r["만료일"],
                   r["검토 신호"], "", "", "", "", "", "보존기간 만료(NAK 5-1:2014 평가·폐기 절차)"])
    if rows:
        dv = DataValidation(type="list", formula1=OPINIONS, allow_blank=True)
        ws.add_data_validation(dv)
        for col in ("J", "L", "N"):
            dv.add(f"{col}3:{col}{len(rows) + 2}")
    for col, w in zip("ABCDEFGHIJKLMNO", (6, 14, 36, 14, 16, 9, 9, 12, 16, 14, 24, 14, 24, 12, 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C3"
    # 생산부서(처리과) 의견 작성 안내 — 「2026년 기록물관리 지침」 제2장 제4절 1. 기록물의 평가 관련 처리과 의견 제출
    gd = wb.create_sheet("의견 작성 안내")
    for line in (["처리과 의견 작성 안내 (「2026년 기록물관리 지침」 제2장 제4절 1)"],
                 ["누가", "처리과 기록물관리책임자가 업무담당자 의견을 받아 취합·조정하고, 처리과 장 결재 후 기록관으로 제출"],
                 ["무엇을", "건(철)마다 '폐기'·'보존기간 재책정'·'보류' 중 하나와 그 사유"],
                 ["검토사항 1 — 유관법령", "유관 법령에 해당 기록물의 보존기간과 관련된 내용이 규정되어 있는지"],
                 ["검토사항 2 — 업무활용", "관련 업무·사업이 끝났는지, 소송이 진행 중인지, 업무에 참고할 필요가 있는지"],
                 ["검토사항 3 — 보존기간 책정기준", "기록관리표(기준표), 시행령 별표 1 보존기간 책정기준, 공통업무 보존기간 준칙(NAK 39 제정 예정)"],
                 ["주의", "처리과는 어떤 기록물도 스스로 폐기할 수 없다. 폐기 결정은 기록관이 전문요원 심사·기록물평가심의회 심의를 거쳐 한다."],
                 ["'검토 신호' 열", "제목에 소송·감사·재산·공사·재난 등이 보여 붙인 단서일 뿐, 판단이 아니다."]):
        gd.append(line)
    gd.column_dimensions["A"].width = 26
    gd.column_dimensions["B"].width = 100
    # 시행규칙 별지 제10호서식 기록물평가심의서 — 1. 기록물철 단위로 평가하는 경우
    f10 = wb.create_sheet("별지 제10호(철 단위)")
    f10.append(["기록물평가심의서 (시행규칙 별지 제10호서식) — 1. 기록물철 단위로 평가하는 경우"])
    f10["A1"].font = Font(bold=True, size=12)
    f10.append(["기록물철 분류번호", "생산연도", "기록물철 제목", "보존기간 만료일", "처리과 의견", "", "기록물관리 전문요원 의견", "", "심의회 의견"])
    f10.append(["", "", "", "", "처리 의견", "사유", "평가 의견", "사유", ""])
    for rr in (2, 3):
        for c in f10[rr]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="EEF1F4")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for a1, a2 in (("A2", "A3"), ("B2", "B3"), ("C2", "C3"), ("D2", "D3"), ("E2", "F2"), ("G2", "H2"), ("I2", "I3")):
        f10.merge_cells(f"{a1}:{a2}")
    for r in rows:
        f10.append([r["관리번호"], r["종료연도"], r["제목"], r["만료일"], "", "", "", "", ""])
    if rows:
        dv2 = DataValidation(type="list", formula1=OPINIONS, allow_blank=True)
        f10.add_data_validation(dv2)
        for col in ("E", "G", "I"):
            dv2.add(f"{col}4:{col}{len(rows) + 3}")
    for col, w in zip("ABCDEFGHI", (18, 9, 36, 13, 12, 24, 12, 24, 12)):
        f10.column_dimensions[col].width = w
    cnt = {}
    for r in rows:
        cnt[r["생산부서"] or "(부서 미기재)"] = cnt.get(r["생산부서"] or "(부서 미기재)", 0) + 1
    s = wb.create_sheet("부서별 건수")
    s.append(["생산부서", "건수"])
    for k, v in sorted(cnt.items(), key=lambda x: -x[1]):
        s.append([k, v])
    wb.save(path)


def write_hwpx(path, rows, year):
    try:
        from hwpx import HwpxDocument
    except ImportError:
        return False
    d = HwpxDocument.new()
    d.add_heading(f"{year}년 기록물 평가 심의 안건", 1)
    d.add_paragraph(f"1. 심의 대상: 보존기간이 끝난 기록물 {len(rows)}건")
    d.add_paragraph("2. 절차: 생산부서 의견조회 → 기록물관리 전문요원 심사 → 기록물평가심의회 심의")
    d.add_paragraph("3. 검토 신호(소송·감사·재산·공사·재난·역사적 가치 관련 제목)는 보류·재책정을 검토할 단서이며 판단이 아님")
    d.add_paragraph("4. 폐기는 심의회 의결 뒤 담당자가 집행함")
    d.add_paragraph("")
    show = rows[:50]
    t = d.add_table(len(show) + 1, 5)
    for c, h in enumerate(["관리번호", "제목", "보존기간", "만료일", "검토 신호"]):
        t.set_cell_text(0, c, h)
    for i, r in enumerate(show, 1):
        for c, k in enumerate(["관리번호", "제목", "보존기간", "만료일", "검토 신호"]):
            t.set_cell_text(i, c, r[k] or "-")
    if len(rows) > 50:
        d.add_paragraph(f"※ 나머지 {len(rows) - 50}건은 붙임 평가심의서.xlsx 참조")
    try:
        d.save_to_path(path)  # python-hwpx 6.x
    except AttributeError:
        d.save(path)  # python-hwpx 7.0부터 save(path)로 합쳐짐
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--year", type=int, default=dt.datetime.now(KST).year)
    ap.add_argument("--include-next", action="store_true", help="'내년 만료'도 미리 포함")
    ap.add_argument("--hold", help="폐기 금지 통보 목록(한 줄에 관리번호, 또는 '단위과제:이름'·'제목:낱말'). 시행령 제54조의4")
    ap.add_argument("-o", "--outdir", default="산출/평가심의")
    a = ap.parse_args()
    rows, has_dept = read(a.table)
    keep = {"만료 지남", "올해 만료"} | ({"내년 만료"} if a.include_next else set())
    rows = [r for r in rows if not r["분류"] or r["분류"] in keep]
    held = []
    if a.hold:
        ids, rules = set(), []
        for line in open(a.hold, encoding="utf-8-sig"):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            k, _, v = line.partition(":")
            if v and k.strip() in ("단위과제", "제목", "생산부서"):
                rules.append((k.strip(), v.strip()))
            else:
                ids.add(line)
        def is_held(r):
            return str(r["관리번호"]).strip() in ids or any(v and v in str(r.get(k) or "") for k, v in rules)
        held = [r for r in rows if is_held(r)]
        rows = [r for r in rows if not is_held(r)]
    for r in rows:
        r["검토 신호"] = signals(r["제목"])
    os.makedirs(a.outdir, exist_ok=True)
    write_xlsx(os.path.join(a.outdir, "평가심의서.xlsx"), rows, f"{a.year}년 기록물 평가심의서 (안)")
    n_dept = 0
    if has_dept:
        groups = {}
        for r in rows:
            groups.setdefault(r["생산부서"] or "부서미기재", []).append(r)
        for dept, rs in groups.items():
            safe = re.sub(r"[\\/:*?\"<>|\s]+", "_", dept)
            write_xlsx(os.path.join(a.outdir, f"의견조회_{safe}.xlsx"), rs, f"{a.year}년 기록물 폐기 의견조회 — {dept}")
            n_dept += 1
    if held:
        for r in held:
            r["검토 신호"] = signals(r["제목"])
        write_xlsx(os.path.join(a.outdir, "폐기금지_제외.xlsx"), held,
                   f"{a.year}년 폐기 금지 기록물 — 심의 대상에서 제외(시행령 제54조의4①, 해제 통보 후 다시 심의)")
    hw = write_hwpx(os.path.join(a.outdir, "심의안건.hwpx"), rows, a.year)
    flagged = sum(1 for r in rows if r["검토 신호"])
    print(f"심의 대상 {len(rows)}건 · 폐기 금지 제외 {len(held)}건 · 검토 신호 {flagged}건 · 의견조회표 {n_dept}개 · HWPX {'생성' if hw else '건너뜀(python-hwpx 없음)'} → {a.outdir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
