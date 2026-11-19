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
    for r in rows:
        kinds, note = classify(r, head, a.year)
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
