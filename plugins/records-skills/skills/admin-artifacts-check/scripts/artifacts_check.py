#!/usr/bin/env python3
"""행정박물 점검: 관리카드(기본정보 11항목)·유형·재질·등록번호·관인 폐기공고번호·공직자선물 4항목을 보고,
이관대상의 이관 기한(시행령 별표 5)을 계산하고, 생산현황 서식 5(유형별 수량)·서식 6(보유목록) 값을 만든다.

사용:  python artifacts_check.py 행정박물목록.xlsx --asof 2025-12-31 [--today 2026-11-19] -o 산출/행정박물_점검.xlsx
목록 열(머리글 이름으로 찾음): 등록번호, 생산부서, 행정박물명, 생산일자, 종료일자, 유형, 재질, 크기, 수량, 내용요약, 관련번호,
       폐기공고번호, 수령인, 수령일, 증정인, 수령경위, 이관대상(이관·제외·공란), 사유발생일, 이관일, 비고
근거(2026-11-19 law.go.kr·국가기록원 확인):
  공공기록물법 제24조, 시행령 제57조(①관리대상 별표 4, ②관리번호·등록정보 생산과 생산현황 보고, ③이관대상 지정,
    ④이관시기 별표 5, ⑤협의로 이관시기 연장, ⑥미지정 행정박물 폐기는 제43조①절차)
  별표 4 유형 7: 관인류, 견본류, 상징류, 기념류, 상장·상패류, 사무집기류, 그 밖의 유형
  별표 5 이관시기: 관인류 폐기하는 때 / 견본류 생산 후 60일 이내 / 상징류 신규 상징물 제작 시·기관 폐지 시 /
    기념류 행사·사업 종료 시 / 상장·상패류 수상 후 1년 이내 / 사무집기류 활용 종료 시 / 그 밖의 유형 영구기록물관리기관의 장이 정하는 시기
  「2026년 기록물관리 지침」 행정박물 관리: 관리카드 기본정보 11항목 필수, 등록번호 처리과코드[7]+등록연도[4]+일련번호[6],
    재질 13종 중 택1, 관인류는 폐기공고번호, 공직자선물은 '그 밖의 유형(공직자선물)'과 수령인·수령일·증정인·수령경위,
    공직자선물 생산·종료일자는 수령일. 처리과→기록관 5.31, 기록관→영구기록물관리기관 8.31 통보.
  생산현황 통보서식 서식 5(유형별 점 수, 없으면 0)·서식 6(보유목록, 없으면 '해당없음', 일자 yyyymmdd)
원칙:  점검과 계산만 한다. 이관대상 지정은 영구기록물관리기관이, 이관·연장 협의·폐기 심의는 사람이 한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

TYPES = ["관인류", "견본류", "상징류", "기념류", "상장·상패류", "사무집기류", "그 밖의 유형"]
MATERIALS = ["금속", "토재·도자기", "석재", "유리·보석", "초재", "목재", "골각·패각", "종이", "모피", "직물", "합성수지", "고무", "기타"]
REQUIRED = ["등록번호", "생산부서", "행정박물명", "생산일자", "종료일자", "유형", "재질", "크기", "수량", "내용요약", "관련번호"]
ALIAS = {"등록번호": ["등록번호", "관리번호", "생산(접수)등록번호"], "생산부서": ["생산부서", "생산부서(처리과)명", "처리과"],
         "행정박물명": ["행정박물명", "명칭"], "생산일자": ["생산일자"], "종료일자": ["종료일자"], "유형": ["유형"], "재질": ["재질"],
         "크기": ["크기"], "수량": ["수량", "수량(점)"], "내용요약": ["내용요약"], "관련번호": ["관련번호"], "폐기공고번호": ["폐기공고번호"],
         "수령인": ["수령인"], "수령일": ["수령일"], "증정인": ["증정인"], "수령경위": ["수령경위"],
         "이관대상": ["이관대상", "이관대상 지정여부", "이관대상지정여부"], "사유발생일": ["사유발생일", "이관사유발생일"],
         "이관일": ["이관일", "이관일자"], "비고": ["비고"]}
# 국가기록원 「행정박물 상태검사·점검 원내 실무 매뉴얼」(2019) 형태 분류·재질 세부 범위 → 비어 있는 유형·재질 후보(판정 아님)
FORM_TYPE = [(r"국새|관인|직인|인장|도장|압인", "관인류"), (r"견본|우표|기념주화|화폐|유가증권|채권", "견본류"),
             (r"기\b|기발|현수막|휘장|배지|엠블럼|상징|현판", "상징류"), (r"상패|상장|메달|트로피|훈장|감사패|공로패", "상장·상패류"),
             (r"책상|의자|탁자|캐비닛|책장|집기|명패", "사무집기류"), (r"기념품|기념패|기념비|기념|선물", "기념류")]
MAT_HINT = [(r"금속|금제|은제|동제|청동|황동|철제|주석|스테인리스|메달|주화|도금", "금속"), (r"나무|목재|현판|목각", "목재"), (r"도자기|도기|자기|토기|그릇|화병", "토재·도자기"),
            (r"직물|면|모직|견직|현수막|깃발|기\b|의복|넥타이|스카프|손수건", "직물"), (r"플라스틱|아크릴|합성수지|PVC", "합성수지"),
            (r"유리|크리스털|보석|옥|비취|수정", "유리·보석"), (r"뼈|뿔|상아|조개|자개", "골각·패각"), (r"돌|석재|대리석|화강암|옥돌", "석재"),
            (r"밀짚|짚|대나무|왕골|초재", "초재"), (r"가죽|모피", "모피"), (r"고무", "고무"), (r"종이|책자|포스터|팸플릿|인쇄물|상장\b", "종이")]
REGNO = re.compile(r"^(\d{7})-?(\d{4})-?(\d{6})$")
BY_EVENT = {"관인류": "관인을 폐기하는 때(신규관인 제작·명칭 변경·기관 폐지)", "상징류": "신규 상징물 제작 시·기관 폐지 시",
            "기념류": "행사·사업 종료 시", "사무집기류": "해당 사무집기류 활용 종료 시"}


def norm_type(v):
    s = re.sub(r"\s", "", v or "")
    if "공직자선물" in s or s.startswith("그밖"):
        return "그 밖의 유형", "공직자선물" in s
    s = s.replace("∙", "·").replace("ㆍ", "·").replace("상장상패", "상장·상패")
    return next((t for t in TYPES if t.replace(" ", "") in (s, s + "류")), ""), False


def norm_mat(v):
    s = re.sub(r"\s", "", v or "").replace("∙", "·").replace("ㆍ", "·").replace(",", "·")
    return next((m for m in MATERIALS if m == s or (m != "기타" and s == m.replace("·", ""))), "")


def parse_date(v):
    s = re.sub(r"\s", "", str(v or ""))[:10]
    for f in ("%Y-%m-%d", "%Y%m%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s.rstrip("."), f).date(), True
        except ValueError:
            pass
    if re.fullmatch(r"\d{4}", s):
        return dt.date(int(s), 12, 31), False
    return None, False


def ymd(v):
    d, full = parse_date(v)
    return "" if not d else (d.strftime("%Y%m%d") if full else str(d.year))


def add_year(d):
    try:
        return d.replace(year=d.year + 1)
    except ValueError:
        return d.replace(year=d.year + 1, day=28)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--asof", help="보유현황 기준일(기본: 전년도 12.31.)")
    ap.add_argument("--today", help="이관 기한 판정일(기본: 오늘)")
    ap.add_argument("-o", "--out", default="산출/행정박물_점검.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    asof = dt.date.fromisoformat(a.asof) if a.asof else dt.date(today.year - 1, 12, 31)
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    items = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]
    seen = {}
    for n, x in items:
        if x.get("등록번호"):
            seen.setdefault(x["등록번호"], []).append(n)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "등록번호", "행정박물명", "유형", "판정"])
    due_ws = wb.create_sheet("이관 기한")
    due_ws.append(["등록번호", "행정박물명", "유형", "이관시기(별표 5)", "기한", "상태"])
    held, counts, bad = [], {t: 0 for t in TYPES}, 0
    for n, x in items:
        issues = []
        for k in REQUIRED:
            if not x.get(k):
                issues.append(f"{k} 비어 있음" + (" — 없으면 '없음'" if k == "관련번호" else ""))
        no = x.get("등록번호", "")
        m = REGNO.match(no)
        if no and not m:
            issues.append(f"등록번호 '{no}' — 처리과코드7+등록연도4+일련번호6(예: 1741069-2022-000704)")
        if no and len(seen[no]) > 1:
            issues.append(f"등록번호가 {len(seen[no])}행에 중복 — 세트·동일 복수는 한 번호에 수량 합산")
        typ, gift = norm_type(x.get("유형"))
        if x.get("유형") and not typ:
            issues.append(f"유형 '{x['유형']}' — 별표 4의 7개 유형 중 택1")
        if x.get("재질") and not norm_mat(x["재질"]):
            issues.append(f"재질 '{x['재질']}' — 대표 재질 13종 중 택1")
        nm = (x.get("행정박물명", "") + " " + x.get("내용요약", ""))
        if not x.get("유형"):
            hint = next((t for pat, t in FORM_TYPE if re.search(pat, nm)), "")
            if hint:
                issues.append(f"유형 후보: {hint}(이름·내용요약 기준 — 확인 후 기입)")
        if not x.get("재질"):
            hints = [m for pat, m in MAT_HINT if re.search(pat, nm)]
            if hints:
                issues.append(f"재질 후보: {'·'.join(dict.fromkeys(hints))} 중 대표 재질 1개(복합재질은 주된 재질)")
        p, pf = parse_date(x.get("생산일자"))
        e, ef = parse_date(x.get("종료일자"))
        for k, d in (("생산일자", p), ("종료일자", e)):
            if x.get(k) and not d:
                issues.append(f"{k} '{x[k]}' — yyyymmdd(불분명하면 연도만)")
        if p and e and e < p and pf and ef:
            issues.append("종료일자가 생산일자보다 앞섬")
        if m and p and m[2] < str(p.year):
            issues.append(f"등록연도 {m[2]}가 생산일자 {p.year}보다 앞섬")
        q = re.sub(r"[,\s점]", "", x.get("수량", ""))
        if x.get("수량") and not (q.isdigit() and int(q) > 0):
            issues.append(f"수량 '{x['수량']}' — 점 단위 숫자")
        if x.get("크기") and not re.search(r"\d", x["크기"]):
            issues.append("크기 — 세로×가로×높이(cm), 부정형은 최대 길이")
        if typ == "관인류" and not x.get("폐기공고번호"):
            issues.append("관인류는 폐기공고번호 필수")
        if gift:
            miss = [k for k in ("수령인", "수령일", "증정인", "수령경위") if not x.get(k)]
            if miss:
                issues.append("공직자선물 " + "·".join(miss) + " 필수")
            r_, _ = parse_date(x.get("수령일"))
            if r_ and (p != r_ or e != r_):
                issues.append("공직자선물은 생산일자·종료일자를 수령일로")
        mark = x.get("이관대상", "")
        if mark not in ("", "이관", "제외"):
            issues.append(f"이관대상 '{mark}' — '이관'·'제외'·공란(미평가)")
        moved, _ = parse_date(x.get("이관일"))
        if mark == "이관" and typ and not moved:
            ev, _ = parse_date(x.get("사유발생일"))
            if typ == "견본류" and p:
                due, rule = p + dt.timedelta(days=60), "생산 후 60일 이내"
            elif typ == "상장·상패류" and p:
                due, rule = add_year(p), "수상 후 1년 이내(수상일=생산일자)"
            elif typ in BY_EVENT:
                due, rule = ev, BY_EVENT[typ] + ("" if ev else " — 사유발생일을 적으면 기한을 판정")
            else:
                due, rule = None, "영구기록물관리기관의 장이 정하는 시기" if typ == "그 밖의 유형" else "생산일자 필요"
            state = "확인 필요" if not due else ("기한 지남 — 이관 또는 연장 협의(시행령 제57조⑤)" if due < today else f"{(due - today).days}일 남음")
            due_ws.append([no, x.get("행정박물명", ""), typ, rule, due.isoformat() if due else "", state])
        if not (moved and moved <= asof) and (not p or p <= asof) and typ:
            q_n = int(q) if q.isdigit() else 1
            counts[typ] += q_n
            held.append([no, x.get("생산부서", ""), x.get("행정박물명", ""), ymd(x.get("생산일자")), ymd(x.get("종료일자")),
                         "그 밖의 유형(공직자선물)" if gift else typ, norm_mat(x.get("재질")) or x.get("재질", ""), x.get("크기", ""), q_n,
                         x.get("폐기공고번호", ""), x.get("수령인", ""), ymd(x.get("수령일")), x.get("증정인", ""), x.get("수령경위", ""),
                         mark, x.get("비고", "")])
        bad += bool(issues)
        ws.append([n, no, x.get("행정박물명", ""), typ or x.get("유형", ""), " · ".join(issues) or "통과"])
    s5 = wb.create_sheet("서식 5 수치")
    s5.append(["기록물구분", f"수량(점) {asof.year}.{asof.month}.{asof.day}. 기준"])
    for t in TYPES:
        s5.append([t, counts[t]])
    s5.append(["합계", sum(counts.values())])
    s6 = wb.create_sheet("서식 6 행")
    s6.append(["연번", "등록번호", "생산부서(처리과)명", "행정박물명", "생산일자", "종료일자", "유형", "재질", "크기", "수량(점)", "관인류 폐기공고번호",
               "수령인", "수령일", "증정인", "수령경위", "이관대상 지정여부", "비고"])
    for i, r in enumerate(held, 1):
        s6.append([i, *r])
    if not held:
        s6.append(["해당없음"])
    for i, w in enumerate((8, 22, 40, 14, 80)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    over = sum(1 for r in due_ws.iter_rows(min_row=2, values_only=True) if str(r[5]).startswith("기한 지남"))
    print(f"행정박물 {len(items)}건 · 고칠 것 {bad} · 이관 기한 지남 {over} · 서식 5 합계 {sum(counts.values())}점 · 서식 6 {len(held)}행 → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
