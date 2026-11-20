#!/usr/bin/env python3
"""보존기간 검토: 기록물 목록의 보존기간을 기록관리기준표와 대조하고, 만료 예정·영구기록물관리기관 이관 대상을 추린다.

사용:  python check_retention.py 목록.xlsx --criteria 기록관리기준표.csv [--year 2026] [--scheme current|draft] -o 산출/보존기간_검토.xlsx

기준표 열(이름은 아래 별칭 중 하나):  단위과제코드 · 단위과제명 · 보존기간
목록 열:                          관리번호 · 단위과제코드(또는 단위과제명) · 보존기간 · 종료년도(또는 생산년도·생산일자)

계산 (references/retention-rules.md)
  기산일    종료(생산)연도 다음 해 1월 1일  → 만료연도 = 종료연도 + 보존연수
  이관 대상  보존기간 30년 이상은 기산일부터 5년이 지난 다음 연도 중 이관(시행령 제40조①, 2026. 8. 25. 개정)
             단, 2026. 8. 25. 시행 당시 기산일부터 5년이 이미 지난 것(종료연도 2020년 이전)은 종전 규정 10년(부칙 제17조①)
  70년 평가  --self-keep(자체 보존 공공기관): 준영구는 기산일부터 70년 경과 후 평가, 2028. 1. 1.부터 영구도(제43조②, 부칙 제5조)
             → 이관연도 = 종료연도 + 1 + N, N = --transfer-years(기본 5, 개정 전 규정은 10)
  허용 값    current: 영구·준영구·30년·10년·5년·3년·1년  /  draft: 입법예고안 5종(3·5·10·20년·영구)도 허용하고 표시 — 2026. 8. 25. 개정 시행령은 7종 유지(미채택)
결과는 검토 목록일 뿐이다. 보존기간 변경·폐기는 기록물평가심의회를 거쳐 사람이 정한다.
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

KST = dt.timezone(dt.timedelta(hours=9))
CURRENT = ["영구", "준영구", "30년", "10년", "5년", "3년", "1년"]
DRAFT_EXTRA = ["20년"]
YEARS = {"준영구": 70, "30년": 30, "20년": 20, "10년": 10, "5년": 5, "3년": 3, "1년": 1}  # 준영구는 정렬용(만료 계산 안 함)
CODE = ["단위과제코드", "단위과제ID", "단위과제 코드", "기록물철분류번호"]
NAME = ["단위과제명", "단위과제"]
RET = ["보존기간"]
END = ["종료년도", "종료연도", "생산년도", "생산연도", "생산일자"]
KEY = ["관리번호", "철관리번호", "기록물건등록번호"]
TITLE = ["제목", "기록물철제목", "기록물건제목"]
BASIS = {
    "허용 값 아님": "공공기록물법 시행령 제26조·별표1 보존기간 구분",
    "개정안 값": "보존기간 5종(3·5·10·20년·영구) 입법예고안 값 — 현행 시행령 제26조①은 7종",
    "기준표에 없음": "단위과제가 기록관리기준표에 없음 — 기준표 갱신 또는 단위과제 확인",
    "기준표보다 짧음": "기준표보다 짧은 보존기간 — 조기 폐기 위험, 근거 확인",
    "기준표보다 김": "기준표와 다름 — 직권 연장 근거 확인",
    "만료 지남": "보존기간 만료 — 기록물평가심의회 평가 대상",
    "올해 만료": "기산일(종료연도 다음 해 1월 1일)부터 보존연수 경과",
    "내년 만료": "다음 해 평가 준비 대상",
    "이관 대상": "보존기간 30년 이상: 기산일부터 N년 경과한 다음 연도 중 영구기록물관리기관 이관(시행령 제40조①, 현행 N=5)",
    "이관 대상(종전)": "개정 시행령(2026. 8. 25.) 시행 당시 기산일부터 5년이 이미 지난 30년 이상 기록물 — 종전 규정(기산일부터 10년 경과 다음 연도) 적용(부칙 제17조①). 시행 전 승인받은 연기는 그 일정(②)",
    "70년 평가": "자체 보존 공공기관: 준영구는 기산일부터 70년(동종·대량·저가치는 50년) 경과 후 평가(시행령 제43조②). 2028. 1. 1.부터 영구도 같음(부칙 제1조제3호·제5조)",
}


def pick(head, names):
    return next((n for n in names if n in head), None)


def year_of(v):
    m = re.search(r"(1[89]\d{2}|20\d{2})", v or "")
    return int(m.group(1)) if m else None


def norm_ret(v):
    v = (v or "").replace(" ", "")
    if re.fullmatch(r"\d+", v):
        v += "년"
    return v


def read_table(path):
    if path.lower().endswith(".xlsx"):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    h = next((i for i, r in enumerate(rows[:20]) if "보존기간" in r), 0)
    head = rows[h]
    return head, [dict(zip(head, r + [""] * (len(head) - len(r)))) for r in rows[h + 1:] if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--criteria", help="기관 기록관리기준표(기록관리표) csv·xlsx")
    ap.add_argument("--common", help="기관 기준표 대신 국가기록원 공통업무 보존기간 준칙을 쓴다: "
                    "중앙행정기관|지방자치단체|공공기관|교육(행정)기관|학교|대학 (assets/common/common_task_criteria.csv)")
    ap.add_argument("--common-basis", choices=["준칙", "nak39안"], default="준칙",
                    help="--common 자료: 현행 준칙(기본) 또는 NAK 39:2026 제정(안)(예고 2026-11-02, 확정 전 — assets/common/nak39_2026_draft.csv)")
    ap.add_argument("--year", type=int, default=dt.datetime.now(KST).year)
    ap.add_argument("--scheme", choices=["current", "draft"], default="current")
    ap.add_argument("--transfer-years", type=int, default=5, help="30년 이상 기록물 이관: 기산일부터 경과 연수(현행 5, 개정 전 10)")
    ap.add_argument("--self-keep", action="store_true", help="시행령 제3조 공공기관처럼 영구기록물관리기관으로 이관하지 않고 자체 보존(제40조④) — 이관 대신 30년 만료·70년 평가")
    ap.add_argument("-o", "--out", default="산출/보존기간_검토.xlsx")
    a = ap.parse_args()

    if a.common:
        import csv as _csv
        import tempfile
        src = os.path.join(os.path.dirname(__file__), "..", "assets", "common",
                           "nak39_2026_draft.csv" if a.common_basis == "nak39안" else "common_task_criteria.csv")
        if a.common_basis == "nak39안":
            print("알림: NAK 39:2026 제정(안)은 확정 고시 전이다. 판단 근거로 쓰기 전 확정본과 대조할 것(references/nak39-draft-diff.md).", file=sys.stderr)
        with open(src, encoding="utf-8-sig") as fh:
            rows_c = [r for r in _csv.DictReader(fh) if r["기관유형"] == a.common]
        if not rows_c:
            print(f"입력 오류: 공통업무 준칙에 '{a.common}' 기관유형이 없습니다.", file=sys.stderr)
            return 2
        tmp = tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, encoding="utf-8-sig", newline="")
        w = _csv.writer(tmp)
        w.writerow(["단위과제코드", "단위과제명", "보존기간", "비고"])
        for r in rows_c:
            w.writerow([r["분류번호"], r["단위과제명"], r["보존기간"], (r.get("검토메모") + " · " if r.get("검토메모") else "") + r["출처"]])
        tmp.close()
        a.criteria = tmp.name
    if not a.criteria:
        print("입력 오류: --criteria 또는 --common 중 하나가 필요합니다.", file=sys.stderr)
        return 2
    ch, crows = read_table(a.criteria)
    th, rows = read_table(a.table)
    c_code, c_name, c_ret = pick(ch, CODE), pick(ch, NAME), pick(ch, RET)
    t_code, t_name, t_ret, t_end = pick(th, CODE), pick(th, NAME), pick(th, RET), pick(th, END)
    t_key, t_title = pick(th, KEY), pick(th, TITLE)
    if not (c_ret and (c_code or c_name)) or not (t_ret and (t_code or t_name)):
        print("입력 오류: 기준표·목록 모두 보존기간과 단위과제(코드 또는 이름) 열이 있어야 합니다.", file=sys.stderr)
        return 2
    by_code = {r[c_code]: norm_ret(r[c_ret]) for r in crows if c_code and r[c_code]}
    by_name = {r[c_name]: norm_ret(r[c_ret]) for r in crows if c_name and r[c_name]}
    allowed = CURRENT + (DRAFT_EXTRA if a.scheme == "draft" else [])

    check, expiry, transfer = [], [], []
    for r in rows:
        ident = [r.get(t_key, "") if t_key else "", r.get(t_title, "") if t_title else ""]
        task = (r.get(t_code) if t_code else "") or (r.get(t_name) if t_name else "")
        ret = norm_ret(r[t_ret])
        ref = by_code.get(r.get(t_code, "")) if t_code and r.get(t_code) else None
        if ref is None and t_name and r.get(t_name):
            ref = by_name.get(r[t_name])
        issues = []
        if ret not in allowed:
            issues.append("허용 값 아님")
        elif ret in DRAFT_EXTRA:
            issues.append("개정안 값")
        if ref is None:
            issues.append("기준표에 없음")
        elif ret in YEARS and ref in YEARS and ret != ref:
            issues.append("기준표보다 짧음" if YEARS[ret] < YEARS[ref] else "기준표보다 김")
        elif ret != ref and ref and ret in allowed and "영구" in (ret, ref):
            issues.append("기준표보다 짧음" if ref == "영구" else "기준표보다 김")
        for k in issues:
            check.append([*ident, task, ret, ref or "", k, BASIS[k]])

        end = year_of(r.get(t_end, "")) if t_end else None
        if end is None:
            continue
        if a.self_keep and ret in ("영구", "준영구"):
            ey = end + 1 + 70              # 기산일부터 70년 경과 후
            if ret == "준영구" or a.year >= 2028:
                if a.year >= ey:
                    expiry.append([*ident, task, ret, end, f"{ey}-01-01 이후", "70년 평가", BASIS["70년 평가"]])
                elif a.year >= ey - 20:
                    expiry.append([*ident, task, ret, end, f"{ey}-01-01 이후", "70년 평가", "동종·대량·저가치면 50년 경과 후 평가 가능(시행령 제43조②) — 해당하는지 확인"])
            continue
        if not a.self_keep and (ret in ("영구", "준영구") or (ret in YEARS and YEARS[ret] >= 30)):
            old = a.transfer_years == 5 and end + 1 + 5 <= 2026   # 2026. 8. 25. 시행 당시 기산일+5년 경과 → 부칙 제17조① 종전 규정
            ty = end + 1 + (10 if old else a.transfer_years)
            if a.year >= ty:
                k = "이관 대상(종전)" if old else "이관 대상"
                transfer.append([*ident, task, ret, end, ty, k, BASIS[k]])
            continue
        if ret not in YEARS:
            continue
        n = YEARS[ret]
        exp = end + n
        kind = "만료 지남" if exp < a.year else "올해 만료" if exp == a.year else "내년 만료" if exp == a.year + 1 else None
        if kind:
            expiry.append([*ident, task, ret, end, f"{exp}-12-31", kind, BASIS[kind]])

    import openpyxl
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "요약"
    s.append(["기준 연도", a.year])
    s.append(["허용 값", " · ".join(allowed)])
    s.append(["목록 행", len(rows)])
    s.append(["보존기간 점검 항목", len(check)])
    s.append(["만료 예정·지남", len(expiry)])
    s.append(["영구기관 이관 대상", len(transfer)])
    s.append([])
    s.append(["주의", "검토 목록입니다. 보존기간 변경·폐기는 기록물평가심의회를 거쳐 사람이 정합니다."])
    s.column_dimensions["A"].width = 18
    s.column_dimensions["B"].width = 70
    idh = [t_key or "관리번호", t_title or "제목"]
    w = wb.create_sheet("보존기간 점검")
    w.append([*idh, "단위과제", "목록 보존기간", "기준표 보존기간", "분류", "근거"])
    for x in check:
        w.append(x)
    w = wb.create_sheet("만료 예정")
    w.append([*idh, "단위과제", "보존기간", "종료연도", "만료일", "분류", "근거"])
    for x in sorted(expiry, key=lambda x: x[6]):
        w.append(x)
    w = wb.create_sheet("이관 대상")
    w.append([*idh, "단위과제", "보존기간", "종료연도", "이관연도", "분류", "근거"])
    for x in transfer:
        w.append(x)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.year}년 기준 · 점검 {len(check)} · 만료 {len(expiry)} · 이관 {len(transfer)} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
