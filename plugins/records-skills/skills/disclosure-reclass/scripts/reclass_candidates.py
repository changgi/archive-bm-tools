#!/usr/bin/env python3
"""공개 재분류 대상 추리기: 비공개·부분공개 기록물 목록에서 올해 심의에 올릴 후보를 나누고, 비공개 사유 표기를 점검한다.

사용:  python reclass_candidates.py 비공개목록.xlsx [--year 2026] -o 산출/공개재분류_후보.xlsx

나누는 기준 (근거는 references/reclass-rules.md)
  ① 30년 경과   생산연도 종료 후 30년이 지난 비공개 기록물 — 공개를 원칙으로 심의(연장하려면 심의회 심의)
  ①-2 연장 신청 기한  생산 후 29~30년 — 비공개를 유지하려면 30년 지난 해의 전년도 말까지 연장 신청(시행령 제72조③)
  ② 5년 주기    마지막 재분류 연도부터 5년이 된 비공개 기록물
  ③ 6호 유예    비공개 사유가 정보공개법 제9조제1항 제6호(개인정보)뿐 → 생산연도 종료 후 30년까지 재분류를 하지 않을 수 있음
  ④ 이력 없음   재분류 연도가 비어 있어 주기를 셀 수 없음 — 담당자 확인
  ⑤ 사유 결함   비공개인데 호수(1~8호)나 구체 사유가 없음, 또는 제1호(다른 법률 비밀·비공개)인데 근거 법령 조항이 없음
                (「2026년 기록물관리 지침」: 제1호로 비공개하면 법령 조항을 사유에 기재 — 예: 「공익신고자 보호법」 제12조)
  --transfer   이관 대상 목록이면 '이관 전 재분류' 시트를 더한다: 최근 5년 안에 공개여부를 재분류한 것은 생략 가능,
               그 밖의 비공개·부분공개는 이관 전에 건별 공개여부 재분류(공공기록물법 제35조, 지침 제3장 이관 준비)
  ⑥ 제외(비밀)  비밀 구분이 있거나 비밀 표시가 있는 기록 — 이 도구로 다루지 않음
검토서(NAK 16-1:2020 부속서 F 16칸)·기준서(부속서 A, 단위과제 열이 있으면 업무별)·공개관리 체크리스트(부속서 J) 시트도 만든다.
--score-col  다른 도구가 낸 비공개 확률(0~1)이 있으면 '비고'에 검토 우선순위 구간(0.10 이하·검토 필요·0.90 이상)을 붙인다
             (국가기록원 2025 연구 3.1.3.3 국외 사례 임계값). 구간은 검토 순서를 정하는 데만 쓰고 공개 여부를 정하지 않는다.
결과는 '후보'일 뿐이다. 공개 여부는 기록물공개심의회가 처리과 의견을 들어 정한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

KST = dt.timezone(dt.timedelta(hours=9))
REASON_RX = re.compile(r"제?\s?([1-8])\s?호")
SECRET_COL_RX = re.compile(r"(비밀|보안등급)")
SECRET_EMPTY = {"", "-", "없음", "해당없음", "해당 없음", "일반"}
SECRET_MARK_RX = re.compile(r"(Ⅰ|Ⅱ|Ⅲ|I{1,3}|[123])\s?급\s?비밀|대외비")
YEAR_COLS = ["생산년도", "생산연도", "종료년도", "생산일자"]
LAST_COLS = ["재분류일", "최근재분류연도", "재분류연도", "최종재분류일"]
BASIS = {
    "30년 경과": "공공기록물법 제35조 — 생산 30년 경과 비공개 기록물은 공개 원칙, 연장은 심의회 심의",
    "5년 주기": "공공기록물법 제35조제2항 — 비공개 기록물 5년 주기 공개여부 재분류",
    "연장 신청 기한": "시행령 제72조③ — 비공개 기간 연장이 필요하면 기산일부터 30년이 지난 해의 전년도 말까지 연장시기·사유를 영구기록물관리기관에 제출",
    "6호 유예": "공공기록물법 제35조제1항 단서 — 정보공개법 제9조제1항제6호 사유는 생산연도 종료 후 30년까지 재분류 미실시 가능",
    "이력 없음": "재분류 연도가 없어 5년 주기를 셀 수 없음 — 담당자 확인",
    "사유 결함": "비공개·부분공개는 정보공개법 제9조제1항 각 호와 구체적 사유를 함께 적어야 함",
    "제외(비밀)": "비밀기록물은 별도 절차(시행령 제68조·제70조) — AI 처리 대상 아님",
}
ORDER = ["30년 경과", "연장 신청 기한", "5년 주기", "6호 유예", "이력 없음", "사유 결함", "제외(비밀)"]


def year_of(v: str):
    m = re.search(r"(1[89]\d{2}|20\d{2})", v or "")
    return int(m.group(1)) if m else None


def read_table(path):
    import openpyxl
    if path.lower().endswith(".xlsx"):
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        import csv
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    h = next((i for i, r in enumerate(rows[:20]) if "공개구분" in r), 0)
    head = rows[h]
    return head, [dict(zip(head, r + [""] * (len(head) - len(r)))) for r in rows[h + 1:] if any(r)]


def classify(r, head, year):
    """한 행의 분류(여러 개 가능)와 비고를 돌려준다."""
    secret_cols = [c for c in head if SECRET_COL_RX.search(c) and "비공개" not in c]
    if any(r[c].strip() not in SECRET_EMPTY for c in secret_cols) or any(SECRET_MARK_RX.search(v) for v in r.values()):
        return ["제외(비밀)"], ""
    disc = r.get("공개구분", "").replace(" ", "")
    if disc not in ("비공개", "부분공개"):
        return [], ""
    out, notes = [], []
    reason = r.get("비공개사유", "") or r.get("비공개 사유", "")
    nums = {int(n) for n in REASON_RX.findall(reason)}
    if not nums or len(REASON_RX.sub("", reason).strip(" ,·-()")) < 4:
        out.append("사유 결함")
    elif 1 in nums and not re.search(r"「[^」]+」|제\s*\d+\s*조", reason):
        out.append("사유 결함")
        notes.append("제1호는 근거 법령 조항을 적어야 함(예: 「공익신고자 보호법」 제12조)")
    py = next((year_of(r[c]) for c in YEAR_COLS if c in r and year_of(r[c])), None)
    if py is None:
        notes.append("생산연도를 읽지 못함")
    elif year - py > 30:
        out.append("30년 경과")
        return out, f"생산 {py}년 · {year - py}년 경과"
    elif year - py in (29, 30):
        # 기산일 = 생산(종료)연도 다음 해 1월 1일 → 30년이 지난 해 = py+31, 그 전년도 말 = py+30년 12월 31일
        out.append("연장 신청 기한")
        notes.append(f"비공개를 30년 넘겨 유지하려면 {py + 30}년 12월 31일까지 연장 신청")
    if nums == {6} and py is not None:
        out.append("6호 유예")
        return out, f"{py + 31}년부터 30년 경과 심의 대상"
    last = next((year_of(r[c]) for c in LAST_COLS if c in r and year_of(r[c])), None)
    if last is None:
        out.append("이력 없음")
    elif year - last >= 5:
        out.append("5년 주기")
        notes.append(f"마지막 재분류 {last}년")
    else:
        notes.append(f"다음 주기 {last + 5}년")
    return out, " · ".join(notes)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--year", type=int, default=dt.datetime.now(KST).year)
    ap.add_argument("--transfer", action="store_true", help="이관 대상 목록: 이관 전 공개여부 재분류가 필요한 것을 따로 표시")
    ap.add_argument("--score-col", help="다른 도구(AI 분류기 등)가 낸 '비공개 확률' 열 이름(0~1). 주면 검토 우선순위 구간을 붙인다")
    ap.add_argument("--low", type=float, default=0.10, help="이 값 이하 = 공개 쪽 확실(기본 0.10)")
    ap.add_argument("--high", type=float, default=0.90, help="이 값 이상 = 비공개 쪽 확실(기본 0.90)")
    ap.add_argument("-o", "--out", default="산출/공개재분류_후보.xlsx")
    a = ap.parse_args()
    head, rows = read_table(a.table)
    if "공개구분" not in head:
        print("입력 오류: '공개구분' 열이 없습니다.", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    summ = wb.active
    summ.title = "요약"
    sheets = {k: wb.create_sheet(k) for k in ORDER}
    for ws in sheets.values():
        ws.append([*head, "분류", "근거", "비고"])
    counts = dict.fromkeys(ORDER, 0)

    def band(r):
        """국가기록원 「AI를 활용한 비공개기록물 공개재분류 및 비식별화 서비스 모델 연구」(2025) 3.1.3.3 국외 사례(미 국무부 외교전문)의
        3구간: 0.10 이하 확실한 공개 · 0.10 초과 0.90 미만 검토 필요 · 0.90 이상 확실한 비공개. 어느 구간이든 결정은 심의회가 한다."""
        if not a.score_col or a.score_col not in r:
            return ""
        try:
            v = float(str(r[a.score_col]).strip())
        except ValueError:
            return "점수 없음 — 검토 필요"
        if not 0 <= v <= 1:
            return f"점수 {v} — 0~1 밖, 확인"
        return f"공개 쪽 확실({v:.2f} ≤ {a.low})" if v <= a.low else (f"비공개 쪽 확실({v:.2f} ≥ {a.high})" if v >= a.high else f"검토 필요({v:.2f})")
    for r in rows:
        kinds, note = classify(r, head, a.year)
        if band(r):
            note = (note + " · " if note else "") + "검토 우선순위: " + band(r)
        for k in kinds:
            counts[k] += 1
            sheets[k].append([*(r[c] for c in head), k, BASIS[k], note])
    if a.transfer:
        wt = wb.create_sheet("이관 전 재분류")
        wt.append([*head, "판정", "근거"])
        need = skip = 0
        for r in rows:
            kinds, _ = classify(r, head, a.year)
            if "제외(비밀)" in kinds or r.get("공개구분", "").replace(" ", "") not in ("비공개", "부분공개"):
                continue
            last = next((year_of(r[c]) for c in LAST_COLS if c in r and year_of(r[c])), None)
            if last is not None and a.year - last < 5:
                skip += 1
                wt.append([*(r[c] for c in head), f"생략 가능(최근 재분류 {last}년)", "최근 5년 안에 공개여부를 재분류한 기록물은 이관 전 재분류 생략 가능"])
            else:
                need += 1
                wt.append([*(r[c] for c in head), "이관 전 재분류 필요", "이관 대상 비공개·부분공개 기록물은 건별 공개여부 재분류 후 이관"])
        summ.append(["이관 전 재분류", f"필요 {need} · 생략 가능 {skip}", "「2026년 기록물관리 지침」 이관대상 기록물 공개여부 재분류"])
    # 별지 제13호의2 비공개로 재분류된 기록물의 비공개 유형별 현황(법 제35조④ 공고, 시행규칙 제40조의2)
    res_col = next((c for c in ("재분류결과", "재분류 결과", "심의결과", "재분류 후 공개구분") if c in head), None)
    rs_col = next((c for c in ("재분류 비공개사유", "재분류비공개사유", "재분류 사유") if c in head), None)
    ry_col = next((c for c in ("재분류연도", "재분류일", "재분류일자", "최종재분류일") if c in head), None)
    labels = ["법령상 비밀·비공개(제1호)", "국방 등 국익침해(제2호)", "국민의 생명 등 공익침해(제3호)", "재판관련 정보 등(제4호)",
              "공정한 업무수행 지장 등(제5호)", "개인사생활 침해(제6호)", "법인 등 영업상 비밀침해(제7호)", "특정인의 이익·불이익(제8호)"]
    f13 = wb.create_sheet("별지 13의2")
    f13.append(["■ 공공기록물 관리에 관한 법률 시행규칙 [별지 제13호의2서식] 비공개로 재분류된 기록물의 비공개 유형별 현황"])
    by, total, partial, nores = [0] * 8, 0, 0, 0
    for r in rows:
        kinds, _ = classify(r, head, a.year)
        if "제외(비밀)" in kinds:
            continue
        if res_col:
            if ry_col and year_of(r.get(ry_col, "")) not in (None, a.year):
                continue
            res = r.get(res_col, "").replace(" ", "")
            if res == "부분공개":
                partial += 1
            if res != "비공개":
                continue
        elif r.get("공개구분", "").replace(" ", "") != "비공개":
            continue
        total += 1
        m = REASON_RX.findall(r.get(rs_col, "") if rs_col else "") or REASON_RX.findall(r.get("비공개사유", "") or r.get("비공개 사유", ""))
        n = int(m[0]) if m else 0
        if 1 <= n <= 8:
            by[n - 1] += 1
        else:
            nores += 1
    f13.append(["재분류된 비공개 총 건수", *labels])
    f13.append([total, *by])
    f13.append([])
    f13.append(["※ 비공개 유형은 「공공기관의 정보공개에 관한 법률」 제9조제1항의 비공개 사유로 하되, 동일 건에 비공개 사유가 2개 이상인 경우에는 주요 비공개 사유를 표기"])
    f13.append(["이 킷의 계산", "사유에 적힌 첫 번째 호를 주요 사유로 봄 — 다르면 고쳐 적음"])
    if res_col:
        f13.append(["기준", f"'{res_col}' 열이 '비공개'인 건" + (f", '{ry_col}'이 {a.year}년인 건" if ry_col else "") + f" · 부분공개로 재분류 {partial}건은 이 서식에 넣지 않음"])
    else:
        f13.append(["주의", "재분류 결과 열(재분류결과)이 없어 현재 공개구분이 '비공개'인 건으로 셈 — 심의 뒤 결과 열을 넣고 다시 돌림"])
    if nores:
        f13.append(["확인", f"호수를 읽지 못한 비공개 {nores}건(총 건수에는 포함)"])
    # NAK 16-1:2020 부속서 F 공개재분류 검토서(16칸) — 심의 대상(30년 경과·연장 신청·5년 주기·이력 없음)으로 채움
    pick = lambda *names: next((c for c in names if c in head), None)
    cols = [pick("관리번호", "철관리번호", "서고관리번호", "철분류번호"), pick("기록물철명", "철제목", "철명"), pick("생산기관", "생산부서", "처리과"),
            pick("생산년도", "생산연도"), pick("등록번호"), pick("건명", "제목"), pick("시작쪽", "시작 쪽"), pick("끝쪽", "끝 쪽"), "공개구분",
            pick("공개제한유형", "비공개사유", "비공개 사유")]
    rv = wb.create_sheet("검토서(NAK 16-1 부속서 F)")
    rv.append(["① 관리번호", "② 기록물철명", "③ 생산기관", "④ 생산년도", "⑤ 등록번호", "⑥ 건명", "⑦ 시작쪽", "⑧ 끝쪽", "⑨ 공개구분", "⑩ 공개제한유형",
               "⑪ 재분류 검토일", "⑫ 재분류 검토의견", "⑬ 재분류 공개구분", "⑭ 재분류 공개제한유형", "⑮ 비공개 대상정보", "⑯ 공개제한 쪽", "후보 분류"])
    nrev = 0
    for r in rows:
        kinds, _ = classify(r, head, a.year)
        k = [x for x in kinds if x in ("30년 경과", "연장 신청 기한", "5년 주기", "이력 없음")]
        if not k:
            continue
        nrev += 1
        rv.append([r.get(c, "") if c else "" for c in cols] + ["", "", "", "", "", "", "·".join(k)])
    rv.append([])
    rv.append(["⑪~⑯은 기록물관리 전문요원이 검토해 채움. 처리과 의견을 듣고 협의·조정한 뒤 공개재분류심의회에 올림(NAK 16-1 6.1.1.2~6.1.1.4)"])
    gs = wb.create_sheet("기준서(NAK 16-1 부속서 A)")
    task_c = pick("단위과제명", "단위과제", "단위업무")
    gs.append(["대상 업무", "업무내용", "업무레벨(중기능·소기능·단위과제)", "생산기관 및 소관부서", "관련 법규 및 판례", "기록물 유형", "기록물 세부내용",
               "기존 공개이력(건수)", "재분류 검토의견", "비공개 대상정보"])
    if task_c:
        agg = {}
        for r in rows:
            if r.get("공개구분", "").replace(" ", "") not in ("비공개", "부분공개"):
                continue
            t_ = r.get(task_c, "") or "(단위과제 없음)"
            g = agg.setdefault(t_, {"dept": set(), "n": 0, "why": set()})
            g["n"] += 1
            if cols[2]:
                g["dept"].add(r.get(cols[2], ""))
            g["why"].update(REASON_RX.findall(r.get(cols[9], "") if cols[9] else ""))
        for t_, g in sorted(agg.items()):
            gs.append([t_, "", "단위과제", ", ".join(sorted(x for x in g["dept"] if x)), "", "", "", f"비공개·부분공개 {g['n']}건",
                       "", ", ".join(f"제{w}호" for w in sorted(g["why"]))])
    else:
        gs.append(["(단위과제명 열이 있으면 업무별로 한 줄씩 채움)"])
    gs.append([])
    gs.append(["작성 절차(4.2.2.2)", "① 업무 분석(고유·공통, 수행부서) ② 기록물 유형화 ③ 공개 여부에 영향을 주는 규정·조례·지침의 비공개 근거와 기간 ④ 판례·보도기사로 비공지성·비공개 필요성 상실 확인 — 결재 또는 심의로 공식화"])
    ck = wb.create_sheet("공개관리 체크리스트")
    ck.append(["영역", "점검 항목(NAK 16-1 부속서 J)", "이행 여부", "근거자료"])
    for line in [("일반", "처리과별·업무유형별 비공개 대상정보 세부기준을 수립했는가?(정보공개법 제9조③)"),
                 ("일반", "세부기준에 기록관의 공개재분류 결과를 계속 반영하는가?"), ("일반", "공개재분류 기준서를 업무 단위로 작성했는가?"),
                 ("일반", "비공지성·비공개 필요성을 잃은 기록물을 정기적으로 점검해 기준에 반영하는가?"),
                 ("업무부서", "생산할 때 공개 여부를 정하고, 부분공개·비공개면 사유(제9조①각 호와 구체 사유)를 적는가?"),
                 ("업무부서", "매년 정리기간과 이관 때 공개값을 다시 보는가?"),
                 ("기록관", "비공개 기록물을 5년마다 재분류하는가?"), ("기록관", "전문요원이 기준서와 업무단위별 검토서를 작성하는가?"),
                 ("기록관", "처리과 의견을 듣고 협의·조정해 기준서·검토서에 반영하는가?"), ("기록관", "공개재분류심의회를 구성·운영하는가?"),
                 ("기록관", "영구기록물관리기관으로 이관하기 전에 비공개 기록물을 재분류하는가?")]:
        ck.append([*line, "", ""])
    for sh, widths in ((rv, (14, 30, 14, 8, 14, 30, 6, 6, 8, 14, 12, 30, 10, 14, 30, 10, 20)), (gs, (24, 30, 14, 24, 24, 16, 24, 18, 30, 20)), (ck, (10, 90, 10, 20))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    summ.append(["검토서(부속서 F)", f"{nrev}건", "NAK 16-1:2020 — 전문요원 검토 후 심의회"])
    summ.append(["별지 13의2", f"비공개로 재분류 {total}건", "법 제35조④ 공고, 시행규칙 제40조의2"])
    summ.append(["기준 연도", a.year])
    summ.append(["입력 행", len(rows)])
    for k in ORDER:
        summ.append([k, counts[k], BASIS[k]])
    summ.append([])
    summ.append(["주의", "후보 목록입니다. 공개 여부는 기록물공개심의회가 처리과 의견을 들어 정합니다."])
    summ.column_dimensions["A"].width = 14
    summ.column_dimensions["C"].width = 80
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.year}년 기준 · " + " · ".join(f"{k} {counts[k]}" for k in ORDER) + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
