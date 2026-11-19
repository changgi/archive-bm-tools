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
근거(law.go.kr 2026-11-20 확인):
  공공기록물법 제33조·제34조, 시행령 제68조(① 처리과→기록관: 일반문서 재분류·보호기간 만료·생산 후 30년,
    ② 기록관→영구기록물관리기관: 30년 이상은 인수한 다음 연도 중, 다만 기산일부터 5년이 안 지났으면 남은 기간 보존 <2026.8.25. 개정>,
    ④ 비전자는 봉인봉투에 건·권별), 제69조(기록관리표 비밀 관련 단위과제 비밀 지정), 제70조(일반문서 재분류 시 편철·공개여부 구분),
    제71조(① 생산·해제·재분류 현황 5.31.→기록관, 8.31.→영구기록물관리기관, ② 목록은 생산 후 3년 지난 다음 연도 5.31., 제목 비밀정보 삭제 가능,
    ③ 기록관은 30년 이상 목록을 8.31.까지)
  「2026년 기록물관리 지침」 비밀기록물(제3장 제6절·제4장 제4절, 이관 시기 문구는 개정 전): 원본만 대상(사본 제외), 보존기간 ≥ 보호기간, 보호기간에 '처리 후·불필요 시·참고 후' 같은
    불확실한 표현 금지, 원본 예고문은 '이관' 또는 '일반문서로 재분류'(사본은 '파기' 가능), 처리과→기록관 이관은 사유가 생긴 다음 연도 중.
    지침의 '기록관→영구 생산 후 10년'은 개정 시행령 제68조②(기산일부터 5년)와 달라 시행령을 따른다.
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
         "전등급": ["변경전등급"], "전보존": ["변경전보존기간"], "기록관": ["기록관이관일", "기록관 이관일"], "영구": ["영구이관일", "영구기록물관리기관 이관일"]}


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
    tr.append(["식별번호", "단계", "사유", "이관할 연도", "상태"])
    c10 = collections.Counter()
    lst = collections.defaultdict(list)
    copies = bad = 0
    for n, x in items:
        if "사본" in x.get("원본", ""):
            copies += 1
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
        if prot and made and p and YEARS[p] < 900 and prot > made.replace(year=made.year + YEARS[p]):
            issues.append(f"보존기간 {p}이 보호기간({prot.isoformat()})보다 짧음 — 보존기간 ≥ 보호기간")
        ye = x.get("예고", "")
        if ye and "파기" in ye:
            issues.append("원본 예고문에 '파기' — 원본은 '이관' 또는 '일반문서로 재분류'")
        elif not ye:
            issues.append("예고문 처리(이관·일반문서로 재분류) 없음")
        # 처리과 → 기록관(시행령 제68조①)
        rc, rk = d(x.get("재분류일")), x.get("재분류", "")
        to_rec, moved1 = None, d(x.get("기록관"))
        if rc and "일반" in rk:
            to_rec, why = rc.year + 1, "유형1 일반문서 재분류"
        elif prot and prot <= today:
            to_rec, why = prot.year + 1, "유형2 보호기간 만료"
        elif made and made.replace(year=made.year + 30) <= today:
            to_rec, why = made.year + 31, "유형3 생산 후 30년(봉인봉투, 전담 관리요원)"
        if to_rec and not moved1:
            tr.append([x.get("번호", ""), "처리과→기록관", why, to_rec, "늦음" if to_rec < today.year else ("올해" if to_rec == today.year else "예정")])
        # 기록관 → 영구기록물관리기관(시행령 제68조②, 2026.8.25. 개정)
        if moved1 and p in LONG and not d(x.get("영구")) and made:
            y = max(moved1.year + 1, made.year + 1 + 5)
            tr.append([x.get("번호", ""), "기록관→영구", "30년 이상, 인수한 다음 연도(기산일부터 5년 전이면 그때까지 보존)", y,
                       "늦음" if y < today.year else ("올해" if y == today.year else "예정")])
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
