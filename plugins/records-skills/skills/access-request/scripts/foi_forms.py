#!/usr/bin/env python3
"""정보공개 서식과 통계: 청구 목록으로 정보공개법 시행규칙 별지 제4호 정보공개 처리대장, 제11호 이의신청 처리대장,
제12호 정보공개 운영실태(총괄·청구방법별·공개/수령방법별·결정기간별·비공개 사유별), 별표 수수료 계산, 제7호 결정 통지서 기재 점검을 만든다.
근거: 「공공기관의 정보공개에 관한 법률 시행규칙」 [시행 2025. 9. 19.] 별표·별지 제3·4·7·11·12호 (law.go.kr 2026-11-20 확인)

사용:  python foi_forms.py 청구목록.xlsx [--year 2026] -o 산출/정보공개_서식.xlsx
입력 열(머리글 이름으로 찾음): 접수번호, 청구일, 정보내용, 공개방법(열람·시청/사본·출력물/전자파일/복제·인화물), 청구방법(직접 출석·우편·팩스·정보통신망),
       담당부서, 결정(공개·부분공개·비공개), 공개내용, 비공개사유, 근거조항, 결정일(결정통지일), 공개일, 수령방법(직접 방문·우편·팩스·정보통신망·전자우편),
       소속(본청·소속기관), 사본규격(A3이상·B4이하), 사본장수, 열람시간(분), 변환(예), 지움(예), 이의신청일, 사건명, 주문, 신청취지, 이유, 이의결정일
수수료(별표): 종이 사본 A3 이상 첫 장 300원·추가 100원, B4 이하 첫 장 250원·추가 50원 / 열람 1일 1시간 무료, 넘으면 30분마다 1,000원 /
       전자파일 복제 무료 / 전자파일 변환이 필요하면 사본 수수료의 1/2, 부분공개 지움+변환이면 사본 수수료와 같게 / 100원 미만 버림 / 매체비용·우송료 별도
원칙:  청구인 성명·연락처는 옮기지 않는다(처리대장 '청구인' 칸은 정보공개시스템에서 채움). 결정·통지는 사람이 한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"번호": ["접수번호"], "청구일": ["청구일", "접수일"], "내용": ["정보내용", "청구내용", "정보 내용"], "공개방법": ["공개방법", "공개 방법"],
         "청구방법": ["청구방법", "청구 방법"], "부서": ["담당부서", "처리부서"], "결정": ["결정", "결정구분", "결정 구분"], "공개내용": ["공개내용", "공개 내용"],
         "사유": ["비공개사유", "비공개 사유"], "근거": ["근거조항", "비공개 근거 조항"], "결정일": ["결정일", "결정통지일"], "공개일": ["공개일"],
         "수령": ["수령방법", "수령 방법"], "소속": ["소속", "기관"], "규격": ["사본규격", "규격"], "장수": ["사본장수", "장수"], "열람": ["열람시간", "열람시간(분)"],
         "변환": ["변환"], "지움": ["지움", "부분공개 지움"], "이의일": ["이의신청일"], "사건": ["사건명"], "주문": ["주문", "주문 내용"], "취지": ["신청취지", "신청 취지"],
         "이유": ["이유", "처리 결과 요지"], "이의결정일": ["이의결정일"]}


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else (v.date().isoformat() if isinstance(v, dt.datetime) else str(v).strip()) for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def day(v):
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", v or "")
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def yes(v):
    return bool(re.match(r"(예|Y|O|○|있음|1)", (v or "").strip(), re.I))


def fee(x):
    """(수수료 원, 산정 명세)"""
    parts, total = [], 0
    n = re.search(r"\d+", x.get("장수", ""))
    n = int(n[0]) if n else 0
    if n:
        big = "A3" in x.get("규격", "").upper() or "이상" in x.get("규격", "")
        first, extra = (300, 100) if big else (250, 50)
        paper = first + extra * (n - 1)
        if yes(x.get("지움")):
            paper_fee, how = paper, "부분공개 지움+변환 = 사본 수수료와 같게"
        elif yes(x.get("변환")):
            paper_fee, how = paper / 2, "전자파일 변환 = 사본 수수료의 1/2"
        elif "전자" in x.get("공개방법", "") and "사본" not in x.get("공개방법", ""):
            paper_fee, how = 0, "전자파일 복제 무료"
        else:
            paper_fee, how = paper, "종이 사본"
        total += paper_fee
        parts.append(f"{'A3 이상' if big else 'B4 이하'} {n}장 {first}+{extra}×{n - 1}={paper}원 → {how} {int(paper_fee)}원")
    m = re.search(r"\d+", x.get("열람", ""))
    if m and int(m[0]) > 60:
        extra_min = int(m[0]) - 60
        v = -(-extra_min // 30) * 1000
        total += v
        parts.append(f"열람 {m[0]}분 — 1시간 넘은 {extra_min}분 → 30분마다 1,000원 {v}원")
    total = int(total // 100 * 100)
    return total, " · ".join(parts) + (" · 100원 미만 버림, 매체비용·우송료 별도" if parts else "")


def kind(v):
    v = re.sub(r"\s", "", v or "")
    return "부분 공개" if "부분" in v else "비공개" if "비공개" in v else "전부 공개" if "공개" in v else ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--year", type=int, help="운영실태 집계 연도(청구일 기준) — 없으면 전체")
    ap.add_argument("-o", "--out", default="산출/정보공개_서식.xlsx")
    a = ap.parse_args()
    rows = read(a.table)
    if a.year:
        rows = [x for x in rows if (day(x.get("청구일")) or dt.date(a.year, 1, 1)).year == a.year]
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    r4 = wb.active
    r4.title = "별지 제4호 처리대장"
    r4.append(["■ 공공기관의 정보공개에 관한 법률 시행규칙 [별지 제4호서식] 정보공개 처리대장"])
    r4.append(["접수번호", "접수일", "청구인", "정보 내용", "공개방법", "담당 부서", "결정 구분", "공개 내용", "비공개(부분 공개) 내용 및 사유", "결정 통지일", "공개일",
               "수령 방법", "비고", "수수료(별표 계산)", "수수료 산정 명세", "확인할 것"])
    chk = 0
    for x in rows:
        k = kind(x.get("결정"))
        notes = []
        if k in ("부분 공개", "비공개"):
            if not re.search(r"[1-8]\s*호", x.get("사유", "") + x.get("근거", "")):
                notes.append("비공개 근거 조항(법 제9조제1항 각 호)이 없음 — 제7호 통지서 '근거 조항' 칸 필수")
            if len(re.sub(r"제?\s*[1-8]\s*호|\s", "", x.get("사유", ""))) < 4:
                notes.append("비공개 내용 및 사유(정보별)가 짧음")
            if "5" in re.findall(r"([1-8])\s*호", x.get("사유", "") + x.get("근거", "")) and not re.search(r"단계|종료", x.get("사유", "")):
                notes.append("제5호(의사결정·내부검토)면 통지서에 단계와 종료 예정일")
        if k in ("전부 공개", "부분 공개") and not x.get("수령"):
            notes.append("수령 방법 없음")
        f, how = fee(x)
        chk += bool(notes)
        r4.append([x.get("번호", ""), x.get("청구일", ""), "(정보공개시스템)", x.get("내용", ""), x.get("공개방법", ""), x.get("부서", ""), k, x.get("공개내용", ""),
                   x.get("사유", ""), x.get("결정일", ""), x.get("공개일", ""), x.get("수령", ""), "이의신청" if x.get("이의일") else "",
                   f if how else "", how, " · ".join(notes)])
        if notes:
            r4.cell(r4.max_row, 16).fill = red
    r11 = wb.create_sheet("별지 제11호 이의신청 처리대장")
    r11.append(["■ 공공기관의 정보공개에 관한 법률 시행규칙 [별지 제11호서식] 이의신청 처리대장"])
    r11.append(["접수번호(처리대장 번호)", "이의신청일", "사건명", "청구인", "주문 내용", "신청 취지", "이유(처리 결과 요지)", "결정 통지일"])
    for x in rows:
        if x.get("이의일"):
            r11.append([x.get("번호", ""), x.get("이의일", ""), x.get("사건", ""), "(정보공개시스템)", x.get("주문", ""), x.get("취지", ""), x.get("이유", ""), x.get("이의결정일", "")])
    r12 = wb.create_sheet("별지 제12호 운영실태")
    r12.append(["■ [별지 제12호서식] 정보공개 운영실태" + (f" — {a.year}년" if a.year else "")])
    orgs = sorted({x.get("소속", "") or "본청" for x in rows}) or ["본청"]

    def table(title, cols, keyf):
        r12.append([])
        r12.append([title])
        r12.append(["구분", "계", *cols])
        tot = [0] * len(cols)
        for o in orgs:
            line = [0] * len(cols)
            for x in rows:
                if (x.get("소속", "") or "본청") != o:
                    continue
                for c in keyf(x):
                    if c in cols:
                        line[cols.index(c)] += 1
            tot = [p + q for p, q in zip(tot, line)]
            r12.append([o, sum(line), *line])
        r12.append(["계", sum(tot), *tot])

    table("2. 가. 총괄(처리 현황)", ["전부 공개", "부분 공개", "비공개"], lambda x: [kind(x.get("결정"))])
    table("2. 나. 청구방법별", ["직접 출석", "우편", "팩스", "정보통신망"],
          lambda x: ["직접 출석" if re.search(r"출석|방문|직접", x.get("청구방법", "")) else "우편" if "우편" in x.get("청구방법", "")
                     else "팩스" if "팩스" in x.get("청구방법", "") else "정보통신망" if x.get("청구방법") else ""])
    pm = lambda x: ["열람·시청" if re.search(r"열람|시청", x.get("공개방법", "")) else "사본·출력물" if re.search(r"사본|출력", x.get("공개방법", ""))
                    else "전자파일" if "전자" in x.get("공개방법", "") else "복제·인화물" if re.search(r"복제|인화", x.get("공개방법", "")) else "기타" if x.get("공개방법") else ""]
    table("2. 다. 공개방법별", ["열람·시청", "사본·출력물", "전자파일", "복제·인화물", "기타"], pm)
    table("2. 다. 수령방법별", ["직접방문", "우편", "팩스", "정보통신망", "기타"],
          lambda x: ["직접방문" if re.search(r"방문|직접", x.get("수령", "")) else "우편" if "우편" in x.get("수령", "") else "팩스" if "팩스" in x.get("수령", "")
                     else "정보통신망" if re.search(r"정보통신|전자우편|온라인|시스템", x.get("수령", "")) else "기타" if x.get("수령") else ""])

    def period(x):
        c, d = day(x.get("청구일")), day(x.get("결정일"))
        if not c or not d:
            return [""]
        n = (d - c).days
        return ["즉시" if n <= 0 else "10일 이내" if n <= 10 else "20일 이내" if n <= 20 else "20일 초과"]
    table("2. 라. 공개 여부 결정기간별(달력일 기준)", ["즉시", "10일 이내", "20일 이내", "20일 초과"], period)
    reasons = [f"제{i}호" for i in range(1, 9)]
    table("4. 비공개 사유별(부분 공개·비공개, 사유에 적힌 호마다)", reasons,
          lambda x: [f"제{h}호" for h in sorted(set(re.findall(r"([1-8])\s*호", x.get("사유", "") + x.get("근거", ""))))] if kind(x.get("결정")) in ("부분 공개", "비공개") else [])
    r12.append([])
    r12.append(["주의", "1. 정보공개창구 설치 현황은 기관이 채움. 결정기간은 달력일로 셌음 — 기관 집계 기준이 다르면 고쳐 씀. 3. 처리 현황 목록은 처리대장 사본으로 갈음 가능."])
    g = wb.create_sheet("별지 제7호 기재 항목")
    g.append(["칸", "적을 내용"])
    for line in [("청구 내용·공개 내용", "공개·부분 공개 때 공개할 정보"), ("공개 일시·장소", "수수료 추가 납부가 필요하면 시행령 제12조에 따라 공개일이 바뀔 수 있음"),
                 ("공개 방법", "열람·시청 / 사본·출력물 / 전자파일 / 복제·인화물 / 기타"), ("수령 방법", "직접 방문 / 우편 / 팩스 전송 / 정보통신망 / 전자우편 등"),
                 ("납부 금액", "① 수수료 ② 우송료 ③ 감면액, 계(①+②-③), 수수료 산정 명세, 납입계좌"),
                 ("비공개 근거 조항", "법 제9조제1항 각 호(부분 공개·비공개 필수)"), ("비공개 내용 및 사유", "제5호(의사결정·내부검토)면 단계와 종료 예정일"),
                 ("결정기간 연장(별지 제3호)", "접수일·접수번호, 당초 결정기간, 연장 사유, 연장 결정기간 — 법 제11조②"),
                 ("시행", "처리과명-연도별 일련번호(시행일), 기관장 직인")]:
        g.append(list(line))
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
        for col in "ABCDEFGHIJKLMNOP":
            sh.column_dimensions[col].width = 14
    r4.column_dimensions["O"].width = 60
    r4.column_dimensions["P"].width = 60
    g.column_dimensions["A"].width = 24
    g.column_dimensions["B"].width = 90
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"청구 {len(rows)}건 · 확인할 것 {chk} · 이의신청 {r11.max_row - 2} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
