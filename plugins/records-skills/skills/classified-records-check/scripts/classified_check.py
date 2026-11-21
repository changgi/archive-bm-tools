#!/usr/bin/env python3
"""비밀기록물 점검: 비밀관리기록부에서 내보낸 '관리정보 목록'으로 예고문·보존기간·이관 시기를 점검하고,
생산현황 서식 10(생산·재분류 통계)과 서식 11(생산 후 3년 지난 목록) 값을 만든다.

★ 보안: 비밀 전용 전산장비(폐쇄망)에서 담당자가 직접 실행한다. 입력·결과 파일을 AI 서비스·클라우드·메일로 옮기지 않는다.
        이 스크립트는 네트워크를 쓰지 않고, 화면에는 건수만 찍는다(제목을 출력하지 않음).

사용:  python classified_check.py 비밀목록.xlsx --year 2025 [--today 2026-11-20] [--list-scope long|all] -o 비밀_점검.xlsx
       --year = 생산현황 대상 연도(전년도). 서식 11 목록은 --year - 3년 생산분(시행령 제71조②: 생산 후 3년이 지난 다음 연도 5.31.).
목록 열(머리글 이름으로 찾음): 식별번호(관리번호), 생산부서, 제목, 유형(문서류·도면카드류·시청각류·책자류·그밖), 구분(전자·비전자 / 시청각은 아날로그·디지털),
       수량, 면수, 비밀등급(1~3), 보존기간, 보호기간(일자 또는 경우), 예고문처리(이관·일반문서로 재분류·파기), 원본구분(원본·사본),
       생산일자, 재분류일자, 재분류내용(일반문서·등급·보존기간), 변경전등급, 변경전보존기간, 기록관이관일, 영구이관일
       (선택) 합철(Y: 일반문서로 재분류 후 일반문서철에 합철 → 그 철의 이관시기), 이관시기연장(Y: 특수기록관 소관 비공개, 30년까지 연장 → 생산 후 30년 지난 다음 연도)
근거(law.go.kr 2026-11-20 확인):
  공공기록물법 제33조·제34조, 시행령 제68조(① 처리과→기록관: 일반문서 재분류·보호기간 만료·생산 후 30년,
    ② 기록관→영구기록물관리기관: 30년 이상은 인수한 다음 연도 중, 다만 기산일부터 5년이 안 지났으면 남은 기간 보존 <2026.8.25. 개정>,
    ④ 비전자는 봉인봉투에 건·권별), 제69조(기록관리표 비밀 관련 단위과제 비밀 지정), 제70조(일반문서 재분류 시 편철·공개여부 구분),
    제71조(① 생산·해제·재분류 현황 5.31.→기록관, 8.31.→영구기록물관리기관, ② 목록은 생산 후 3년 지난 다음 연도 5.31., 제목 비밀정보 삭제 가능,
    ③ 기록관은 30년 이상 목록을 8.31.까지)
  「비밀기록물 관리 실무지침」(2020. 11.): 보존기간 말일 ≥ 보호기간 말일, 유형1~3 이관 전 정리(대각선 삭제·재분류 근거·공개여부·관리기록부 정리·봉인),
  「영구기록물관리기관 비밀기록물 관리지침」(2025. 12.): 서고(일반·별도·전용) 시설 기준, 정수점검 2년, 열람·반출. 이 지침의 기록관→영구 '기산일부터 10년'은 개정 시행령(2026.8.25.) 전 문구.
  「2026년 기록물관리 지침」 비밀기록물(제3장 제6절·제4장 제4절, 이관 시기 문구는 개정 전): 원본만 대상(사본 제외), 보존기간 ≥ 보호기간, 보호기간에 '처리 후·불필요 시·참고 후' 같은
    불확실한 표현 금지, 원본 예고문은 '이관' 또는 '일반문서로 재분류'(사본은 '파기' 가능), 처리과→기록관 이관은 사유가 생긴 다음 연도 중.
    지침의 '기록관→영구 생산 후 10년'은 개정 시행령 제68조②(기산일부터 5년)와 달라 시행령을 따른다.
  NAK 17:2021(v1.3) 6.2: 원본 예고문은 이관·일반문서로 재분류, 사본은 파기·일반문서로 재분류(원본이 '이관'이면 사본은 반드시 '파기')
  생산현황 통보서식 서식 10·11 기재요령(등급별 합 = 보존기간별 합, 재분류는 변경 전 기준, 등급은 숫자, 목록 보존기간은 30년·준영구·영구)
원칙:  점검·계산만. 재분류·이관·봉인·통보는 비밀취급인가를 받은 담당자가 한다.
"""
import argparse
import collections
import datetime as dt
import os
import re
import sys

KINDS = ["문서류", "도면·카드류", "시청각류", "책자류", "그밖의종류"]
# 서식 10 열: (전자/디지털, 비전자/아날로그)
COLS = {"문서류": ("E", "F"), "도면·카드류": ("H", "I"), "시청각류": ("L", "K"), "책자류": ("N", "O"), "그밖의종류": ("P", "P")}
PERIODS = ["영구", "준영구", "30년", "10년", "5년", "3년", "1년"]
LONG = ("영구", "준영구", "30년")
YEARS = {"영구": 999, "준영구": 998, "30년": 30, "10년": 10, "5년": 5, "3년": 3, "1년": 1}
VAGUE = ["처리 후", "처리후", "불필요 시", "불필요시", "참고 후", "참고후", "필요시", "필요 시"]
ALIAS = {"번호": ["식별번호", "관리번호", "비밀관리번호"], "부서": ["생산부서"], "제목": ["제목"], "유형": ["유형", "기록물유형"],
         "구분": ["구분", "전자구분", "형태"], "수량": ["수량", "건수", "매수", "권수"], "면수": ["면수"], "등급": ["비밀등급", "등급"],
         "보존": ["보존기간"], "보호": ["보호기간"], "예고": ["예고문처리", "예고문", "보호기간 만료 시 처리"], "원본": ["원본구분", "원본여부"],
         "생산": ["생산일자", "생산일"], "재분류일": ["재분류일자", "재분류일"], "재분류": ["재분류내용", "재분류구분"],
         "전등급": ["변경전등급"], "전보존": ["변경전보존기간"], "기록관": ["기록관이관일", "기록관 이관일"], "영구": ["영구이관일", "영구기록물관리기관 이관일"],
         "검토": ["최근재분류검토일", "재분류검토일"],
         "정수": ["최근정수점검일", "정수점검일", "정수점검"], "서고": ["보관장소", "서고"],
         "연장": ["이관시기연장", "특수기록관연장", "30년연장"], "합철": ["합철", "일반문서철합철", "합철여부"]}


def d(v):
    s = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y%m%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def kind(v):
    s = re.sub(r"[\s·∙ㆍ]", "", v or "")
    for k in KINDS:
        if re.sub(r"[·]", "", k).startswith(s[:2]) and s:
            return k
    return "그밖의종류" if s else ""


def per(v):
    s = re.sub(r"\s", "", str(v or ""))
    return s if s in PERIODS else (s + "년" if s + "년" in PERIODS else "")


def grade(v):
    m = re.search(r"([123]|Ⅰ|Ⅱ|Ⅲ|I{1,3})", str(v or ""))
    return {"Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "I": 1, "II": 2, "III": 3}.get(m[1], int(m[1]) if m and m[1].isdigit() else 0) if m else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, required=True, help="생산현황 대상 연도(전년도)")
    ap.add_argument("--today")
    ap.add_argument("--list-scope", choices=["long", "all"], default="long", help="서식 11: 기록관→영구는 30년 이상(long), 처리과→기록관은 all")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    items = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "식별번호", "유형", "등급", "보존기간", "판정"])
    tr = wb.create_sheet("이관 시기")
    rv = wb.create_sheet("재분류 검토서")
    rv.append(["식별번호", "생산부서", "생산연도", "비밀등급", "예고문(보호기간)", "형태", "제목(비밀정보는 지움)", "대상 사유", "주요 내용(채움)",
               "재분류 의견: 비밀유지·등급변경·보호기간변경·일반문서·파기(채움)", "공개여부·정보공개법 제9조①호(채움)", "생산기관 의견(채움)"])
    pr = wb.create_sheet("이관 준비")
    pr.append(["식별번호", "유형", "서고(영구기관 기준)", "이관 방법", "이관 전 정리(처리과·기록관이 마쳐야 할 것)"])
    tr.append(["식별번호", "단계", "사유", "이관할 연도", "상태"])
    c10 = collections.Counter()
    lst = collections.defaultdict(list)
    copies = bad = 0
    for n, x in items:
        if "사본" in x.get("원본", ""):
            copies += 1
            if "이관" in x.get("예고", ""):  # NAK 17:2021 6.2: 사본은 파기 또는 일반문서로 재분류
                bad += 1
                ws.append([n, x.get("번호", ""), kind(x.get("유형")), grade(x.get("등급")) or x.get("등급", ""), per(x.get("보존")) or x.get("보존", ""),
                           "사본 예고문에 '이관' — 사본은 '파기' 또는 '일반문서로 재분류'(NAK 17 6.2)"])
            continue
        issues = []
        k, g, p = kind(x.get("유형")), grade(x.get("등급")), per(x.get("보존"))
        made, prot_txt = d(x.get("생산")), x.get("보호", "")
        prot = d(prot_txt)
        if not k:
            issues.append("유형 없음 — 문서류·도면카드류·시청각류·책자류·그밖의종류")
        if not g:
            issues.append(f"비밀등급 '{x.get('등급', '')}' — 1·2·3급")
        if not p:
            issues.append(f"보존기간 '{x.get('보존', '')}' — 7종 중")
        if not made:
            issues.append("생산일자 없음")
        if any(v in prot_txt for v in VAGUE):
            issues.append(f"보호기간 '{prot_txt}' — 도래가 명확한 일자·경우로(불확실 표현 금지)")
        elif not prot_txt:
            issues.append("보호기간 없음")
        # 보존기간 말일(생산 다음 해 1.1. 기산 → 생산연도+기간 12.31.)이 보호기간 말일보다 같거나 나중(비밀기록물 관리 실무지침 2020 Ⅲ.1.나)
        if prot and made and p and YEARS[p] < 900 and prot > dt.date(made.year + YEARS[p], 12, 31):
            need = next((q for q in ("1년", "3년", "5년", "10년", "30년") if dt.date(made.year + YEARS[q], 12, 31) >= prot), "준영구 이상")
            issues.append(f"보존기간 {p}(말일 {made.year + YEARS[p]}.12.31.)이 보호기간 말일({prot.isoformat()})보다 앞섬 — 최소 {need}(시행령 제67조)")
        ye = x.get("예고", "")
        if ye and "파기" in ye:
            issues.append("원본 예고문에 '파기' — 원본은 '이관' 또는 '일반문서로 재분류'")
        elif not ye:
            issues.append("예고문 처리(이관·일반문서로 재분류) 없음")
        # 처리과 → 기록관(시행령 제68조①)
        rc, rk = d(x.get("재분류일")), x.get("재분류", "")
        to_rec, moved1 = None, d(x.get("기록관"))
        if rc and "일반" in rk and re.match(r"(Y|예|합철)", x.get("합철", "").strip()):
            to_rec, why = None, ""
            tr.append([x.get("번호", ""), "처리과→기록관", "일반문서로 재분류 후 일반문서철에 합철 — 그 철의 이관시기를 따름(NAK 17 8.2.3)", "", "철 기준"])
        elif rc and "일반" in rk:
            to_rec, why = rc.year + 1, "유형1 일반문서 재분류"
        elif prot and prot <= today:
            to_rec, why = prot.year + 1, "유형2 보호기간 만료"
        elif made and made.replace(year=made.year + 30) <= today:
            to_rec, why = made.year + 31, "유형3 생산 후 30년(봉인봉투, 전담 관리요원)"
        if to_rec:
            ty = why[:3]
            prep = {"유형1": ("일반서고", "일반기록물과 같은 절차",
                              "비밀등급 표시 대각선 삭제 · 첫 면 여백에 '예고문 등에 따라 일반문서로 재분류(일자)'+직위·성명 기명날인 · 첫 페이지 우측 상단에 공개여부(비공개면 괄호 안 호수) · 비밀관리기록부 '비밀등급~사본번호' 적색 2선 삭제, 처리방법 '일반재분류' 표시, 처리자·확인자 결재, 이관일·인계자 기재"),
                    "유형2": ("별도서고(일반서고보다 보안 강화: 이중 잠금·CCTV·분리 보관)", "일반기록물 절차, 유형1과 구분해 이관",
                              "비밀등급 표시 대각선 삭제 · 첫 면에 '예고문에 따른 비밀 보호기간 만료(일자)'+기명날인 · 첫 페이지 우측 상단 공개여부·비공개 호수 · 전자 등록 시 스캔 사본 첨부 금지(접근제한)"),
                    "유형3": ("전용서고(이중 잠금·CCTV·서가별 잠금, Ⅰ급은 금고)", "비밀 전담 관리요원, 건·철별 봉인봉투(간인) 또는 국정원장이 정한 보안 전송망",
                              "이관 전 비밀 재분류 실시 · 보호기간 30년 이상·만료 여부 확인 · 끝부분에 비밀열람기록전 첨부 · 「보안업무규정」에 따라 열람·복사(인가자, 전담요원 입회)")}.get(ty)
            if prep:
                pr.append([x.get("번호", ""), ty, *prep])
        # 재분류 대상(실무지침 2020·영구기관 지침 2025 Ⅱ.7, 2010 재분류 매뉴얼 흐름: 대상 선정→검토서→생산기관 의견조회→심의회→확정·후속)
        why_rv = []
        if prot and prot <= today:
            why_rv.append("보호기간 만료")
        if made and dt.date(made.year + 31, 1, 1) <= today:
            why_rv.append("기산일부터 30년 지남" + (" — 보호기간 남음: 생산기관 동의 필요" if prot and prot > today else ""))
        lr = d(x.get("검토"))
        if "검토" in idx and prot and prot > today and (not lr or (today - lr).days > 183):
            issues.append("보호기간 만료 전 주기적 사전검토가 6개월 넘게 없음(실무지침 2020 Ⅲ.1.가, 2010 매뉴얼: 원본 연 2회 6월·12월)")
        if why_rv and not (rc and "일반" in rk):
            rv.append([x.get("번호", ""), x.get("부서", ""), made.year if made else "", g or x.get("등급", ""), prot_txt, k, x.get("제목", ""), " · ".join(why_rv), "", "", "", ""])
        cy = d(x.get("정수"))
        if x.get("정수", "") != "" or "정수" in idx:
            if not cy or cy.replace(year=cy.year + 2) < today:
                issues.append("정수점검 2년 주기 지남(또는 기록 없음) — 철 단위 육안 점검 후 목록 현행화(영구기록물관리기관 비밀기록물 관리지침 2025)")
        if to_rec and not moved1:
            tr.append([x.get("번호", ""), "처리과→기록관", why, to_rec, "늦음" if to_rec < today.year else ("올해" if to_rec == today.year else "예정")])
        # 기록관 → 영구기록물관리기관(시행령 제68조②, 2026.8.25. 개정)
        if moved1 and p in LONG and not d(x.get("영구")) and made:
            legacy = made.year + 1 + 5 <= 2026 and moved1 <= dt.date(2026, 8, 25)   # 부칙 제19조: 시행 당시 기록관 보유·기산일+5년 경과
            y = max(moved1.year + 1, made.year + 1 + (10 if legacy else 5))
            ext = re.match(r"(Y|예|연장)", x.get("연장", "").strip())
            if ext:
                y = max(y, made.year + 31)
            why = "특수기록관 소관 비공개, 이관시기 30년까지 연장 — 생산 후 30년 지난 다음 연도(NAK 17 8.3.2 비고)" if ext else ("30년 이상, 개정 시행령 시행(2026. 8. 25.) 당시 기산일부터 5년이 지난 것 — 종전 규정(10년) 적용(부칙 제19조, 종전 제68조② 원문 확인)"
                   if legacy else "30년 이상, 인수한 다음 연도(기산일부터 5년 전이면 그때까지 보존)")
            tr.append([x.get("번호", ""), "기록관→영구", why, y, "늦음" if y < today.year else ("올해" if y == today.year else "예정")])
        # 서식 10 (가) 생산현황: 해당 연도 생산분
        q = int(re.sub(r"\D", "", x.get("수량", "")) or 1)
        if made and made.year == a.year and k and g and p:
            col = COLS[k][1 if re.search(r"비전자|아날로그|종이", x.get("구분", "")) else 0]
            c10[(5 + g, col)] += q
            c10[(9 if p not in LONG else 10, col)] += q
        # (나)·(다) 재분류현황: 해당 연도에 재분류, 변경 전 기준
        if rc and rc.year == a.year and k:
            col = COLS[k][1 if re.search(r"비전자|아날로그|종이", x.get("구분", "")) else 0]
            g0, p0 = grade(x.get("전등급")) or g, per(x.get("전보존")) or p
            if "일반" in rk:
                if g0:
                    c10[(10 + g0, col)] += q
                if p0:
                    c10[(14 if p0 not in LONG else 15, col)] += q
            else:
                if "등급" in rk and g0:
                    c10[(15 + g0, col)] += q
                if "보존" in rk and p0:
                    c10[(19 if p0 not in LONG else 20, col)] += q
        # 서식 11: 생산 후 3년 지난 목록
        if made and made.year == a.year - 3 and k and (a.list_scope == "all" or p in LONG):
            elec = not re.search(r"비전자|아날로그|종이", x.get("구분", ""))
            lst[k].append([x.get("번호", ""), x.get("부서", ""), x.get("제목", ""), "" if elec else x.get("면수", ""), g, p, prot_txt,
                           "전자" if elec else "비전자"])
        bad += bool(issues)
        ws.append([n, x.get("번호", ""), k, g or x.get("등급", ""), p or x.get("보존", ""), " · ".join(issues) or "통과"])
    s10 = wb.create_sheet("서식 10 수치")
    s10.append(["서식 행", "E", "F", "H", "I", "K", "L", "N", "O", "P", "구분"])
    labels = {6: "(가) I급", 7: "(가) II급", 8: "(가) III급", 9: "(가) 30년 미만", 10: "(가) 30년·준영구·영구", 11: "(나) I급", 12: "(나) II급",
              13: "(나) III급", 14: "(나) 30년 미만", 15: "(나) 30년 이상", 16: "(다) I급", 17: "(다) II급", 18: "(다) III급", 19: "(다) 30년 미만",
              20: "(다) 30년 이상"}
    for r in range(6, 21):
        s10.append([r] + [c10.get((r, c), 0) for c in "EFHIKLNOP"] + [labels[r]])
    mism = []
    for a1, b1, lab in ((range(6, 9), (9, 10), "(가)"), (range(11, 14), (14, 15), "(나)")):
        for c in "EFHIKLNOP":
            if sum(c10.get((r, c), 0) for r in a1) != sum(c10.get((r, c), 0) for r in b1):
                mism.append(f"{lab} {c}열")
    s11 = wb.create_sheet("서식 11 행")
    s11.append(["유형", "연번", "식별번호", "생산부서", "제목", "면수", "비밀등급", "보존기간", "보호기간", "구분"])
    for k in KINDS:
        for i, r in enumerate(lst[k], 1):
            s11.append([k, i, *r])
    for i, w in enumerate((8, 18, 12, 6, 10, 90)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    late = sum(1 for r in tr.iter_rows(min_row=2, values_only=True) if r[4] == "늦음")
    print(f"원본 {len(items) - copies}건(사본 {copies}건 제외) · 고칠 것 {bad} · 이관 늦음 {late} · 서식 11 목록 {sum(len(v) for v in lst.values())}건"
          + (f" · 등급별≠보존기간별 {', '.join(mism)}" if mism else "") + f" → {a.out}")
    print("결과 파일은 비밀 전용 장비 밖으로 옮기지 않는다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
