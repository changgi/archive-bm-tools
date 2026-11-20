#!/usr/bin/env python3
"""복원처리 계획: 상태검사 결과(훼손도 등급)·중요도·활용도·pH로 복원처리 방법(심화·예방·응급)과 우선순위, 탈산 처리 대상을 나누고,
처리 기록(처리 전·후 상태, 사진, 처리자, 재료) 칸이 있는 복원처리 대장을 만든다.
근거: NAK 25:2026(v2.1) 「종이기록물 보존 및 복원에 관한 표준」 4.4·4.5·6.1·7.2·부속서 C (국가기록원, 2026-11-20 내려받아 확인)

사용:  python restoration_plan.py 상태검사결과.xlsx -o 산출/복원처리계획.xlsx
입력 열(머리글 이름으로 찾음): 관리번호, 제목, 훼손도(1~3등급), 중요도(상·중·하), 활용도(상·중·하), pH, 대형(예), 재난피해(예), 긴급요청(예), 상태(바스러짐·황변 등)
규칙(부속서 C, 상태검사 3등급 기준):
  심화복원 — 훼손도 3등급 + 중요도 상 + 활용도 상: 전 과정(상태조사→해체→건식·습식 세척→수리흔적 제거→결실 보강→건조·제책→복원처리서)
  예방복원 — 3등급(또는 대형·구조 취약)인데 중요도·활용도가 그보다 낮음: 훼손요인 제거·구조 안정화·편철 개선
  응급복원 — 재난 피해 또는 즉시 활용(등록·열람·스캐닝) 긴급 요청: 곰팡이 제거·세척→낱장 해체→최소 보강→재편철
  탈산 처리 이력('탈산연도' 열)이 있으면: pH 6.8 미만은 재탈산·품질검사 검토, pH가 없으면 표본 측정(약 1%, 권당 3곳) 안내, 백화현상 표시
    (국가기록원 「탈산 처리 기록물의 상태평가 및 개선방안 연구」 — 1,035권 표본, 미 의회도서관 처리 하한 pH 6.8 인용)
  탈산(7.2) — pH 6.5 이하 또는 바스러짐·심한 황변 → 처리 후 pH 7.5~10.0 목표(필기재료·종이 안정성 사전 실험)
원칙:  분류 후보만 낸다. 처리 방법은 관리자와 복원 전문가가 협의해 정한다(4.2). 원형 유지·최소 개입·가역성.
"""
import argparse
import os
import re
import sys

ALIAS = {"번호": ["관리번호", "철관리번호", "등록번호"], "제목": ["제목", "철제목", "기록물철제목"], "훼손": ["훼손도", "종이 훼손도", "훼손등급", "훼손도 등급"],
         "중요": ["중요도"], "활용": ["활용도"], "ph": ["pH", "PH", "ph"], "대형": ["대형", "대형기록물"], "탈산이력": ["탈산연도", "탈산처리연도", "탈산이력", "탈산처리"], "재난": ["재난피해", "재난 피해"],
         "긴급": ["긴급요청", "긴급", "즉시활용"], "상태": ["상태", "비고", "훼손 상태"]}
LV = {"상": 3, "중": 2, "하": 1}


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def yes(v):
    return bool(re.search(r"^(예|Y|O|○|있음|해당|1|true)", (v or "").strip(), re.I))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("-o", "--out", default="산출/복원처리계획.xlsx")
    a = ap.parse_args()
    rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "복원처리 계획"
    ws.append(["우선순위", "관리번호", "제목", "훼손도", "중요도", "활용도", "pH", "처리 방법(후보)", "탈산", "처리 절차", "확인할 것"])
    PROC = {"심화복원": "상태조사 → 해체 → 건식 세척 → 습식 세척 → 테이프 등 수리흔적 제거 → 결실부 보강 → 건조·제책 → 복원처리서",
            "예방복원": "촬영·조사 → 세척 → 훼손요인 제거 → 결실 보강 → 편철 상태 개선 → 복원처리서",
            "응급복원": "곰팡이 제거·세척 → 낱장 해체 → 최소한의 결실부 보강 → 재편철",
            "보존 관리": "적정 보존환경·보존용품(폴더·상자)으로 관리, 다음 상태검사에서 재확인"}
    out, cnt = [], {}
    for x in rows:
        g = re.search(r"[123]", x.get("훼손", ""))
        g = int(g[0]) if g else None
        imp, use = LV.get(re.sub(r"\s", "", x.get("중요", ""))[:1]), LV.get(re.sub(r"\s", "", x.get("활용", ""))[:1])
        ph = re.search(r"\d+(\.\d+)?", x.get("ph", ""))
        ph = float(ph[0]) if ph else None
        notes = []
        if yes(x.get("재난")) or yes(x.get("긴급")):
            m = "응급복원"
        elif g == 3 and imp == 3 and use == 3:
            m = "심화복원"
        elif g == 3 or (yes(x.get("대형")) and (g or 0) >= 2):
            m = "예방복원"
        else:
            m = "보존 관리"
        if g is None:
            notes.append("훼손도 등급 없음 — 시행령 별표 3 기준으로 상태검사")
        if m != "응급복원" and g == 3 and (imp is None or use is None):
            notes.append("중요도·활용도가 없어 심화·예방을 가를 수 없음(예방복원으로 둠)")
        brittle = re.search(r"바스러|부스러|심한\s*황변|황변\s*심", x.get("상태", ""))
        done = x.get("탈산이력", "").strip()
        deacid = (ph is not None and ph <= 6.5) or bool(brittle)
        if done and done not in ("-", "없음", "N", "아니오"):
            if ph is not None and ph < 6.8:
                notes.append(f"탈산 처리 기록물인데 pH {ph:g} < 6.8(미 의회도서관 처리 하한) — 탈산 효과 저하, 재탈산·품질검사 검토(국가기록원 탈산 상태평가 연구)")
            elif ph is None:
                notes.append("탈산 처리 기록물 — 사후 품질검사: 약 1% 표본, 권당 첫 장·1/4·1/2 지점 3곳 pH·백색도 측정, 백화현상 육안 확인")
            if re.search(r"백화|하얗|흰\s*가루", x.get("상태", "")):
                notes.append("백화현상 — MgO 과량 흡착, 용제 린스로 줄어들 수 있음(보존처리 담당 판단)")
            deacid = deacid and not (ph is not None and ph >= 6.8)
        if deacid:
            notes.append("탈산 대상 — 필기재료·종이의 약품 안정성 사전 실험, 처리 후 pH 7.5~10.0 목표(7.2)")
        if ph is None and m in ("심화복원", "예방복원"):
            notes.append("pH 측정 필요(지시약·냉수추출법·pH-meter)")
        score = {"응급복원": 400, "심화복원": 300, "예방복원": 200, "보존 관리": 0}[m] + (imp or 0) * 10 + (use or 0) * 10 + (g or 0)
        cnt[m] = cnt.get(m, 0) + 1
        out.append((-score, x, g, ph, m, deacid, notes))
    for i, (_, x, g, ph, m, deacid, notes) in enumerate(sorted(out, key=lambda z: z[0]), 1):
        ws.append([i, x.get("번호", ""), x.get("제목", ""), f"{g}등급" if g else "", x.get("중요", ""), x.get("활용", ""), "" if ph is None else ph, m,
                   "대상" if deacid else "", PROC[m], " · ".join(notes)])
    lg = wb.create_sheet("복원처리 기록(4.5·6.1)")
    lg.append(["관리번호", "처리 전 상태(재질: 섬유·필기재료·기록방식 / 물리: 크기·두께·중량·백색도 / 훼손: pH·색차·육안)", "처리 전 사진", "처리 과정 사진",
               "처리 후 사진", "처리 방법", "사용 재료·약품(접착제·한지·펄프 pH 6~9)", "테스트 결과", "처리자", "처리 기간", "처리 후 상태", "보존 위치·용품"])
    for _, x, g, ph, m, deacid, notes in sorted(out, key=lambda z: z[0]):
        if m != "보존 관리":
            lg.append([x.get("번호", ""), f"pH {ph}" if ph is not None else "", "", "", "", m, "", "", "", "", "", ""])
    rl = wb.create_sheet("기준(NAK 25)")
    for line in [("원칙", "원형 유지, 원본에 최소한의 개입, 재료·기술은 가역성, 예방적 기술(환경통제·취급·보존용기·대체물)이 먼저 — 복원은 최종 선택(4.2·4.3)"),
                 ("선정", "상태검사 결과와 활용·보존 계획을 고려해 우선순위, 자주 전시·활용되는 기록물은 우선, 복원 전문가 의견 반영(4.4)"),
                 ("기록", "처리 전·후 상태, 사진(전·중·후), 처리자, 재료·약품·방법·테스트 결과를 기록물과 함께 영구 보존(4.5·6.1)"),
                 ("탈산", "pH 6.5 이하 또는 바스러짐·심한 황변 → pH 7.5~10.0, 전용 약제, 장비 또는 소량은 분무·습식(7.2)"),
                 ("실무체계", "심화(3등급·중요 상·활용 상) / 예방(3등급 중 중요·활용 중·하, 대형·구조 취약) / 응급(재난 피해·즉시 활용)(부속서 C)"),
                 ("보존용품", "공공기록물법 제29조, 시행령 제61조, 시행규칙 제38조 별표 15")]:
        rl.append(list(line))
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
    for sh, widths in ((ws, (8, 16, 30, 8, 8, 8, 6, 12, 8, 80, 60)), (lg, (16, 60, 12, 12, 12, 12, 30, 20, 10, 12, 20, 20)), (rl, (12, 120))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(rows)}건 · " + " · ".join(f"{k} {v}" for k, v in cnt.items()) + f" · 탈산 대상 {sum(z[5] for z in out)} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
