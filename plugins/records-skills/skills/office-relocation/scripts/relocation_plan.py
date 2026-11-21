#!/usr/bin/env python3
"""청사 이전 대비 기록물 이송·재배치 계획: 보유현황 목록으로 통계·보존기간 확인·평가폐기 후보·관리번호·이송 박스·배가표·추진상황 점검표를 만든다.

사용:  python relocation_plan.py 보유현황목록.xlsx --year 2026 [--per-box 10] [--org 기관명] -o 산출/청사이전.xlsx
목록 열(머리글 이름으로 찾음): 부서(처리과·생산부서), 기록물철 제목(철제목·제목), 생산연도, 보존기간, 보관위치(위치), 권수, 쪽수, 비밀구분
근거:  국가기록원 「공공기관 청사 이전 대비 기록관리 업무 매뉴얼」(2014. 12.) 제2장(기본계획·보유현황·평가폐기·이송재배치)·부록 1 추진상황 점검표
       「지방이전 정부기관 기록물관리 지침」(2011) Ⅲ.5 박스 작업(박스번호·권수·생산부서), Ⅳ 보존서고 면적(보존량의 약 140%, 이동식 서가 면적 50% 안팎 감축,
       하중 고정식 750·이동식 1,000~1,200kgf/㎡), Ⅳ.4 유형별 소요예산(예시), 참고 1 이전비용 산출 내역서(항목·관리경비 5%·기업이윤 10%·부가세 10%)
원칙:  계획과 목록만 만든다. 폐기·이송은 절차에 따라 사람이 한다. 비밀구분 철은 목록에 넣지 않고 건수만 센다.
"""
import argparse
import collections
import csv
import os
import re
import sys

ALIAS = {"부서": ["부서", "처리과", "생산부서", "부서명"], "제목": ["기록물철제목", "철제목", "제목", "기록물철 제목"], "연도": ["생산연도", "생산년도"],
         "보존": ["보존기간"], "위치": ["보관위치", "위치", "보관장소"], "권": ["권수", "수량"], "쪽": ["쪽수"], "비밀": ["비밀구분", "비밀여부"]}
YEARS = {"1년": 1, "3년": 3, "5년": 5, "10년": 10, "20년": 20, "30년": 30}
ORDER = {"영구": 0, "준영구": 1, "30년": 2, "20년": 3, "10년": 4, "5년": 5, "3년": 6, "1년": 7}
CHECK = [("1. 기본계획", "기관 청사 이전 추진계획(일정) 숙지", "기록관"),
         ("1. 기본계획", "기록물관리 기본계획 수립·시행(현황·문제점, 이송 사전준비, 인적·물적 인프라, 기록관 설치·운영, 예산, 추진체계, 월·주 단위 일정)", "기록관"),
         ("2. 보유현황", "보유현황 조사 계획(대상·방법·일정) 수립, 위탁·자체 결정, 처리부서 통보·조사 매뉴얼 교육", "기록관"),
         ("2. 보유현황", "조사 준비물(작업공간·카트·일련번호 스티커·보존기간표 라벨·장갑·마스크·보존용 표지·이송박스 등)", "기록관"),
         ("2. 보유현황", "처리부서·문서고 보관 기록물철에 일련번호 스티커 부착, 목록 작성(비밀기록물은 따로)", "처리과"),
         ("2. 보유현황", "보존기간 누락 확인 → 처리부서 의견조회 → 보존기간별 목록·통계 취합", "기록관"),
         ("2. 보유현황", "관리번호 부여·보존기간표 레이블 배포·부착, 미편철 기록물 편철, 생산부서→생산연도→보존기간 순 서가 배치", "기록관·처리과"),
         ("3. 평가·폐기", "보존기간 경과 대상 목록 → 평가·폐기 계획 → 처리부서 의견조회 → 전문요원 심사 → 평가심의회 → 폐기 확정·집행(폐기증명) → 보류·연장 사후조치", "기록관"),
         ("4. 이송·재배치", "이송 대상량 파악, 이송계획 수립, 사전준비(이송용 박스·박스별 목록·박스 인식 스티커·신청사 서가 정리·인력·차량·선탑자)", "기록관"),
         ("4. 이송·재배치", "이송용 박스 작업(박스별 목록 동봉, 스티커 부착)", "처리과"),
         ("4. 이송·재배치", "반출·이송(이송차량 직원 선탑), 신청사 서고 배가(서가·박스 번호 일치), 해포·배치, 목록-실물 대조(인계서)", "기록관"),
         ("4. 이송·재배치", "이전 대상 전자문서 사전조사, 전자문서시스템·데이터 훼손 대비 복제", "기록관·정보화부서"),
         ("5. 설계·시공", "보존시설 규모 산정(보유량 기준), 보존공간·사무공간·열람공간 면적, 고정식 서가 750kgf/㎡·이동식 1,000~1,200kgf/㎡ 하중, 항온항습·온습도계·소화·CCTV, 전산·열람장비", "기록관"),
         ("6. 기록관 설치·운영", "현황조사·분석, 설치·운영계획, 직제 반영, 전문요원·적정인력 배치, 기록관 운영규정·업무분장, 공간·시스템·시설 구축", "기록관"),
         ("보안", "기록물·일반자료 폐기 때 파쇄·입회 등 보안 준수, 개인정보 노출 방지", "기록관·처리과")]


def main() -> int:
    ap = argparse.ArgumentParser(description="청사 이전 대비 기록물 이송·재배치 계획")
    ap.add_argument("list")
    ap.add_argument("--year", type=int, required=True, help="기준 연도(평가·폐기 후보 계산)")
    ap.add_argument("--per-box", type=int, default=10, help="이송 박스 하나에 담을 권 수")
    ap.add_argument("--org", default="")
    ap.add_argument("--growth-years", type=int, default=10, help="신청사 서고 산정에 넣을 앞으로의 연수")
    ap.add_argument("-o", "--out", default="산출/청사이전.xlsx")
    a = ap.parse_args()
    if a.list.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.list, encoding="utf-8-sig")))
    head = [re.sub(r"\s", "", h) for h in rows[0]]
    idx = {k: next((head.index(re.sub(r"\s", "", x)) for x in al if re.sub(r"\s", "", x) in head), None) for k, al in ALIAS.items()}
    if idx["제목"] is None:
        print("입력 오류: 기록물철 제목 열이 필요합니다.", file=sys.stderr)
        return 2
    items, secret = [], 0
    for n, r in enumerate(rows[1:], 2):
        g = lambda k, r=r: r[idx[k]].strip() if idx[k] is not None and idx[k] < len(r) else ""
        if not any(r):
            continue
        if g("비밀") and not re.fullmatch(r"(일반|없음|해당없음|-|N|평문)", g("비밀")):
            secret += 1
            continue
        per = re.sub(r"\s", "", g("보존"))
        y = re.search(r"(19|20)\d{2}", g("연도"))
        items.append({"행": n, "부서": g("부서") or "(부서 없음)", "제목": g("제목"), "연도": int(y[0]) if y else None, "보존": per,
                      "위치": g("위치"), "권": int(g("권")) if g("권").isdigit() else 1, "쪽": g("쪽")})
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    st = wb.active
    st.title = "보유현황 통계"
    pers = sorted({x["보존"] for x in items if x["보존"] in ORDER}, key=ORDER.get)
    st.append(["부서"] + pers + ["보존기간 확인 필요", "합계(권)"])
    by = collections.defaultdict(collections.Counter)
    for x in items:
        by[x["부서"]][x["보존"] if x["보존"] in ORDER else "확인"] += x["권"]
    for d in sorted(by):
        st.append([d] + [by[d][p] for p in pers] + [by[d]["확인"], sum(by[d].values())])
    st.append([])
    st.append(["보관위치", "권수"])
    lc = collections.Counter()
    for x in items:
        lc[x["위치"] or "(위치 없음)"] += x["권"]
    for loc, c in lc.most_common():
        st.append([loc, c])
    if secret:
        st.append([])
        st.append(["비밀구분 철", f"{secret}건 — 목록에서 뺐음, 비밀기록물 담당자가 따로 처리"])
    ck = wb.create_sheet("보존기간 확인")
    ck.append(["원본 행", "부서", "기록물철 제목", "생산연도", "보존기간", "확인할 것"])
    for x in items:
        why = []
        if not x["보존"]:
            why.append("보존기간 없음 — 처리부서 의견조회")
        elif x["보존"] not in ORDER:
            why.append(f"'{x['보존']}' — 영구·준영구·30년·10년·5년·3년·1년 중 하나")
        if not x["연도"]:
            why.append("생산연도 없음")
        if why:
            ck.append([x["행"], x["부서"], x["제목"], x["연도"] or "", x["보존"], " · ".join(why)])
    dz = wb.create_sheet("평가·폐기 후보")
    dz.append(["원본 행", "부서", "기록물철 제목", "생산연도", "보존기간", "보존기간 만료(연말)", "비고"])
    n_due = 0
    for x in items:
        if x["보존"] in YEARS and x["연도"] and x["연도"] + YEARS[x["보존"]] < a.year:
            n_due += 1
            dz.append([x["행"], x["부서"], x["제목"], x["연도"], x["보존"], x["연도"] + YEARS[x["보존"]],
                       "평가 절차(의견조회·전문요원 심사·평가심의회) 뒤에만 폐기 — 이송 전 임의 폐기 금지"])
    items.sort(key=lambda x: (x["부서"], x["연도"] or 0, ORDER.get(x["보존"], 9), x["제목"]))
    lb = wb.create_sheet("관리번호·레이블")
    lb.append(["관리번호", "생산부서", "생산연도", "보존기간", "기록물철 제목", "권수", "레이블 문구"])
    for k, x in enumerate(items, 1):
        x["관리번호"] = f"{k:05d}"
        lb.append([x["관리번호"], x["부서"], x["연도"] or "", x["보존"], x["제목"], x["권"], f"{x['관리번호']} | {x['부서']} | {x['연도'] or ''} | {x['보존']}"])
    bx = wb.create_sheet("이송 박스")
    bx.append(["박스번호", "부서", "보존기간", "관리번호", "기록물철 제목", "권수", "박스 스티커"])
    box, cur, used = 0, None, 0
    boxes = collections.OrderedDict()
    for x in items:
        key = (x["부서"], x["보존"])
        if key != cur or used + x["권"] > a.per_box:
            box += 1
            cur, used = key, 0
        used += x["권"]
        bno = f"B{box:04d}"
        boxes.setdefault(bno, []).append(x)
        bx.append([bno, x["부서"], x["보존"], x["관리번호"], x["제목"], x["권"], f"{a.org} {bno} {x['부서']} {x['보존']}".strip()])
    pl = wb.create_sheet("배가표")
    pl.append(["박스번호", "부서", "보존기간", "관리번호 범위", "권수", "신청사 서고", "서가 번호", "배가 확인", "목록-실물 대조"])
    for bno, xs in boxes.items():
        pl.append([bno, xs[0]["부서"], xs[0]["보존"], f"{xs[0]['관리번호']}~{xs[-1]['관리번호']}", sum(x["권"] for x in xs), "", "", "", ""])
    lab = wb.create_sheet("박스 라벨")
    lab.append(["박스번호", "권수", "생산부서", "보존기간", "관리번호 범위", "동봉 목록"])
    for k, (bno, xs) in enumerate(boxes.items(), 1):
        lab.append([f"{k:05d}", sum(x["권"] for x in xs), xs[0]["부서"], xs[0]["보존"], f"{xs[0]['관리번호']}~{xs[-1]['관리번호']}", "박스별 기록물 목록 1부"])
    lab.append([])
    lab.append(["지침 예시", "박스 앞면에 박스번호(00001)·권수(박스 안 철 수)·생산부서를 붙이고, 박스별 목록 1부를 넣어 봉함"])
    # 신청사 서고 산정
    total = sum(x["권"] for x in items)
    yrs = collections.Counter()
    for x in items:
        if re.fullmatch(r"\d{4}", str(x["연도"])):
            yrs[int(x["연도"])] += x["권"]
    recent = [yrs.get(y, 0) for y in range(a.year - 5, a.year)]
    per_year = round(sum(recent) / 5) if any(recent) else 0
    proj = total + per_year * a.growth_years
    need = round(proj * 1.4)
    kind = "A형(1만 권 기준)" if need < 40000 else "B형(4만 권 이상)" if need < 150000 else "C형(15만 권 이상)" if need < 300000 else "D형(30만 권 이상)"
    sz = wb.create_sheet("신청사 서고 산정")
    sz.append(["항목", "값", "근거·설명"])
    for row in (["현재 보유량(권)", total, "보유현황 목록 합계(비밀 제외)"],
                ["최근 5년 연평균 생산(권)", per_year, f"{a.year - 5}~{a.year - 1} 생산연도 권수 평균 — 추가 이관 예상량 추정에 씀"],
                [f"{a.growth_years}년 뒤 예상량(권)", proj, "현재 + 연평균 × 기간(--growth-years)"],
                ["필요 수용량(권)", need, "작업공간을 고려해 보존기록물량 대비 약 140% 확보(지침 Ⅳ.2)"],
                ["서가 방식", "이동식 검토", "이동식 서가(모빌렉)는 소요면적 50% 안팎 감축 — 하중 1,000~1,200kgf/㎡ 이상 필요, 고정식은 750kgf/㎡ 이상"],
                ["비밀기록물", f"{secret}건", "별도 전용서고 필요(공공기록물법 제32조, 시행령 제66조) — 목록에서 뺀 건수"],
                ["소요예산 유형(예시)", kind, "지침 Ⅳ.4 예시(2011, 백만원): A 33.2 · B 74.4 · C 225.3 · D 431.2 — 현재 단가로 다시 산출"],
                ["보존환경", "stack-environment로 점검", "현행 기준은 시행령 별표 6(지침의 표는 2011년 값)"]):
        sz.append(row)
    # 이전비용 견적 점검(수량·단가는 기관 견적으로 채움, 금액은 수식)
    ct = wb.create_sheet("이전비용 견적 점검")
    ct.append(["구분", "항목", "수량", "단가(원)", "일수·할증", "금액(원)", "지침 예시(2011, 참고)"])
    rows = [("운반임", "화물차(5톤 윙바디 등) 대수", None, None, 1, "65대 × 290,000원 × 1일"),
            ("작업임", "반출·이송·상차·하차·이송·반입(R/T)", None, None, 1.5, "1,147.9 R/T × 2,056~2,267원 × 할증 1.5, 6항목"),
            ("포장·정리", "발지 작업 인원", None, None, None, "12명 × 84,686원 × 10일"),
            ("포장·정리", "착지 작업 인원", None, None, None, "12명 × 84,686원 × 10일"),
            ("옵션비", "모빌렉 해체·설치(련)", None, None, 1, "180련 × 400,000원"),
            ("옵션비", "항온항습기 이전", None, None, 1, "1대 × 45,000,000원(100평)"),
            ("옵션비", "공기청정기 이전", None, None, 1, "3대 × 1,500,000원"),
            ("포장재료", "서류박스·테이프·물품송표·바구니", 1, None, 1, "17,900,000원")]
    for i, (g, item, q, u, d, ex) in enumerate(rows, 2):
        ct.append([g, item, q, u, d, f"=IF(AND(C{i}<>\"\",D{i}<>\"\"),C{i}*D{i}*IF(E{i}=\"\",1,E{i}),\"\")", ex])
    n = len(rows) + 1
    ct.append(["소계", "", "", "", "", f"=SUM(F2:F{n})", ""])
    ct.append(["관리경비", "소계의 5%", "", "", "", f"=ROUND(F{n + 1}*0.05,0)", "5%"])
    ct.append(["기업이윤", "(소계+관리경비)의 10%", "", "", "", f"=ROUND((F{n + 1}+F{n + 2})*0.1,0)", "10%"])
    ct.append(["부가세", "(소계+관리경비+기업이윤)의 10%", "", "", "", f"=ROUND((F{n + 1}+F{n + 2}+F{n + 3})*0.1,0)", "10%"])
    ct.append(["보험료", "이송 중 분실·파손 대비", "", "", "", "", "5,000,000원"])
    ct.append(["합계", "천 원 단위 절사", "", "", "", f"=ROUNDDOWN(SUM(F{n + 1}:F{n + 5}),-3)", ""])
    ct.append([])
    ct.append(["주의", "작업임 6항목은 항목마다 한 줄씩 늘려 적는다. 수량·단가는 기관이 받은 견적으로 채운다(예시 단가를 그대로 쓰지 않음)."])
    ct.column_dimensions["B"].width = 36
    ct.column_dimensions["G"].width = 44
    sz.column_dimensions["C"].width = 70
    cs = wb.create_sheet("추진상황 점검표")
    cs.append(["영역", "점검항목", "수행주체", "준비", "진행", "완료", "산출물·비고"])
    for row in CHECK:
        cs.append(list(row) + ["", "", "", ""])
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    cs.column_dimensions["B"].width = 100
    for col in "BE":
        lb.column_dimensions[col].width = 40
        bx.column_dimensions[col].width = 40
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"철 {len(items)}건({sum(x['권'] for x in items)}권) · 보존기간 확인 {ck.max_row - 1} · 평가·폐기 후보 {n_due} · 이송 박스 {len(boxes)}"
          + (f" · 비밀 {secret}건 제외" if secret else "") + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
