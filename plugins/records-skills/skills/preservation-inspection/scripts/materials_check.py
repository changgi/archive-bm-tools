#!/usr/bin/env python3
"""30년 이상 보존기록물의 기록재료 점검: 목록의 용지·필기구(잉크)·보존용기가 공공기록물법 시행규칙 별표 15(제38조, 영 제61조)에 맞는지 본다.
근거(law.go.kr 2026-11-20 확인): 시행규칙 별표 15
  - 종이: 한지류 또는 보존용지 1종(보존복사용지·보존백상지·보존아트지). 보존용지 1종 = 화학펄프 100%, pH 7.5 이상, 탄산칼슘 2% 이상
  - 보존용 판지(표지·용기): pH 7.5 이상, 탄산칼슘 3% 이상
  - 잉크: 먹, 탄소형·안료형, 보존용잉크 / 필기구: 사무용 프린터 토너, 탄소형 사인펜, 흑색 안료형·보존용 필기구 — pH 7.0 이상, 내광성 4호 이상, 산·알칼리(표백제 포함)에 변하지 않을 것

사용:  python materials_check.py 목록.xlsx -o 산출/기록재료_점검.xlsx
목록 열(머리글 이름으로 찾음): 철분류번호(또는 관리번호), 제목, 보존기간, 용지(종이), 필기구(잉크), 용기(표지·상자), 용지pH, 탄산칼슘, 판지pH, 판지탄산칼슘, 잉크pH, 내광성
--important: 국가기록원 「중요문서용 기록용지 표준화 및 시제품 개발연구」 표 80 규격(안)도 본다(백색도 80%·pH 7.5~10·Kappa 5 이하·탄산칼슘 3%·
       인열강도 공식서한류 600/협약류 1,200 mN·인장강도 80 N/15mm(90 gsm 기준, 평량으로 환산)·가속노화 ΔE 3.5 이하·인열 유지 87.5%). 연구 제안이라 '참고'로 표시.
원칙:  목록 값으로 '확인 필요'를 고를 뿐이다. 실제 재질 판정은 pH 측정·제품 인증서로 한다. 30년 미만은 별표 15 대상이 아니라 건너뛴다.
"""
import argparse
import os
import re
import sys

LONG = ("영구", "준영구", "30년")
ALIAS = {"번호": ["철분류번호", "관리번호", "분류번호", "등록번호"], "제목": ["제목", "철제목", "건제목"], "보존": ["보존기간"],
         "용지": ["용지", "종이", "용지종류"], "필기": ["필기구", "잉크", "필기구(잉크)", "출력방식"], "용기": ["용기", "표지", "보존상자", "보존용기"],
         "용지pH": ["용지pH", "종이pH", "pH"], "탄산": ["탄산칼슘", "용지탄산칼슘"], "판지pH": ["판지pH"], "판지탄산": ["판지탄산칼슘"],
         "잉크pH": ["잉크pH"], "내광": ["내광성"],
         "백색도": ["백색도"], "카파": ["Kappa", "KappaNo", "카파값", "카파"], "인열": ["인열강도", "인열강도(CD)"], "인장": ["인장강도", "인장강도(MD)"],
         "평량": ["평량", "gsm"], "변색": ["가속노화변색도", "변색도", "ΔE"], "유지율": ["가속노화인열강도", "인열강도유지율"], "용도": ["용도", "용지용도"]}
PAPER_OK = re.compile(r"한지|보존복사용지|보존백상지|보존아트지|보존용지\s*1종")
PAPER_BAD = re.compile(r"신문|갱지|재생|감열|팩스|트레이싱|일반\s*복사|중성지\s*아님|산성")
INK_OK = re.compile(r"먹|탄소|안료|보존용|토너|레이저")
INK_BAD = re.compile(r"염료|수성|볼펜|형광|잉크젯|연필|감열|만년필")
BOX_OK = re.compile(r"보존용\s*(판지|상자|표지)|중성|무산")


def num(s):
    m = re.search(r"\d+(\.\d+)?", str(s or ""))
    return float(m[0]) if m else None


def read(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(path, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        import csv
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    head = [re.sub(r"\s", "", h) for h in rows[0]]
    idx = {k: next((head.index(re.sub(r"\s", "", a)) for a in al if re.sub(r"\s", "", a) in head), None) for k, al in ALIAS.items()}
    return idx, rows[1:]


def check(get):
    out = []
    paper, ink, box = get("용지"), get("필기"), get("용기")
    if not paper:
        out.append(("용지", "용지 정보 없음 — 확인 필요"))
    elif PAPER_BAD.search(paper):
        out.append(("용지", f"'{paper}' — 한지류·보존용지 1종이 아님(복사·이관 시 보존용지로 복제 검토)"))
    elif not PAPER_OK.search(paper):
        out.append(("용지", f"'{paper}' — 보존용지 1종 여부 확인 필요"))
    ph, ca = num(get("용지pH")), num(get("탄산"))
    if ph is not None and ph < 7.5:
        out.append(("용지", f"용지 pH {ph} — 7.5 미만(restoration_plan의 탈산 후보: pH 6.5 이하)"))
    if ca is not None and ca < 2:
        out.append(("용지", f"탄산칼슘 {ca}% — 2% 미만"))
    if not ink:
        out.append(("필기구", "필기구·잉크 정보 없음 — 확인 필요"))
    elif INK_BAD.search(ink):
        out.append(("필기구", f"'{ink}' — 먹·탄소형·안료형·보존용 잉크·토너가 아님(번짐·퇴색 위험)"))
    elif not INK_OK.search(ink):
        out.append(("필기구", f"'{ink}' — 보존용 여부 확인 필요"))
    iph, lf = num(get("잉크pH")), num(get("내광"))
    if iph is not None and iph < 7.0:
        out.append(("필기구", f"잉크 pH {iph} — 7.0 미만"))
    if lf is not None and lf < 4:
        out.append(("필기구", f"내광성 {lf:g}호 — 4호 미만"))
    if box and not BOX_OK.search(box):
        out.append(("용기", f"'{box}' — 보존용 판지 표지·용기 여부 확인 필요"))
    bph, bca = num(get("판지pH")), num(get("판지탄산"))
    if bph is not None and bph < 7.5:
        out.append(("용기", f"판지 pH {bph} — 7.5 미만"))
    if bca is not None and bca < 3:
        out.append(("용기", f"판지 탄산칼슘 {bca}% — 3% 미만"))
    return out


def important(get):
    """국가기록원 「중요문서용 기록용지 표준화 및 시제품 개발연구」 표 80 표준규격(안) — 공식서한류·협약류. 확정 표준이 아닌 연구 제안."""
    out = []
    pact = "협약" in get("용도")
    g = num(get("평량")) or 90.0
    for key, lim, label, low in (("백색도", 80, "백색도 80% 이상", True), ("탄산", 3, "예비 알칼리(탄산칼슘) 3% 이상", True),
                                 ("카파", 5, "Kappa No. 5 이하", False), ("변색", 3.5, "가속노화 후 변색도 ΔE 3.5 이하", False),
                                 ("유지율", 87.5, "가속노화 후 인열강도 초기의 87.5% 이상", True)):
        v = num(get(key))
        if v is not None and (v < lim if low else v > lim):
            out.append(("중요문서용(안)", f"{label} {'미달' if low else '초과'}({v:g})"))
    ph = num(get("용지pH"))
    if ph is not None and not 7.5 <= ph <= 10.0:
        out.append(("중요문서용(안)", f"pH 7.5~10.0 밖({ph:g})"))
    tear, ten = num(get("인열")), num(get("인장"))
    need = (1200 if pact else 600) * g / 90
    if tear is not None and tear < need:
        out.append(("중요문서용(안)", f"인열강도(CD) {tear:g} mN — {'협약류' if pact else '공식서한류'} {need:.0f} 이상(90 gsm 기준 {1200 if pact else 600}을 평량 {g:g}로 환산)"))
    if ten is not None and not pact and ten < 80 * g / 90:
        out.append(("중요문서용(안)", f"인장강도(MD) {ten:g} N/15mm — 공식서한류 {80 * g / 90:.0f} 이상(90 gsm 기준 80)"))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("list")
    ap.add_argument("--important", action="store_true", help="중요문서용 기록용지 표준규격(안)(국가기록원 연구 표 80, 공식서한류·협약류)도 본다 — 백색도·Kappa·인열·인장·가속노화")
    ap.add_argument("-o", "--out", default="산출/기록재료_점검.xlsx")
    a = ap.parse_args()
    idx, rows = read(a.list)
    if idx["보존"] is None:
        print("입력 오류: 보존기간 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검 결과"
    ws.append(["원본 행", "번호", "제목", "보존기간", "항목", "내용"])
    target = skipped = bad = 0
    for n, r in enumerate(rows, 2):
        get = lambda k, r=r: (r[idx[k]] if idx[k] is not None and idx[k] < len(r) else "").strip()
        if not any(r):
            continue
        if not re.sub(r"\s", "", get("보존")).startswith(LONG):
            skipped += 1
            continue
        target += 1
        res = check(get) + (important(get) if a.important else [])
        bad += bool(res)
        for item, msg in res:
            ws.append([n, get("번호"), get("제목"), get("보존"), item, msg])
    s = wb.create_sheet("요약")
    for row in (["30년 이상 대상", target], ["30년 미만 건너뜀", skipped], ["확인 필요 기록물", bad],
                ["근거", "공공기록물법 시행규칙 별표 15(제38조, 영 제61조) — 종이 pH 7.5·탄산칼슘 2%, 판지 pH 7.5·3%, 잉크 pH 7.0·내광성 4호"]):
        s.append(row)
    for i, w in enumerate((8, 16, 36, 10, 10, 70)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"30년 이상 {target}건 · 확인 필요 {bad}건 · 30년 미만 건너뜀 {skipped} → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
