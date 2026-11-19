#!/usr/bin/env python3
"""별지 제7호서식 「보존기록물 정수점검ㆍ상태검사서」 작업지 만들기·결과 반영.

사용:
  1) 작업지:  python inspection_form.py form 보존목록.xlsx --year 2026 [--all] -o 산출/별지7_2026.xlsx
     올해 돌아온(지난 것 포함) 정수점검·상태검사 대상을 서식 칸 순서로 채운 작업지. --all이면 대상과 관계없이 전부.
  2) 결과 반영: python inspection_form.py apply 산출/별지7_2026_작성.xlsx 보존목록.xlsx -o 산출/보존목록_갱신.xlsx
     작성된 작업지에서 점검일자·검사일자와 검사후등급을 읽어 보존목록 '사본'의 최근정수점검·최근상태검사·상태등급을 고친다.
     종이·행정박물의 상태등급 = 그 기록물 검사 항목 검사후등급 중 가장 나쁜 값(3이 가장 나쁨). 3등급은 사후조치 후보로 따로 적는다.
서식(law.go.kr, 시행규칙 별지 제7호서식 <개정 2026. 9. 11.>, 2026-11-20 확인):
  1. 정수점검서: 제 호·점검일자·담당자 확인 / 시스템번호·관리번호·철제목·기록물유형·소장위치·보존상자번호·육안상태점검·점검목록 조치사항(삭제·수정·추가·기타 등)·비고
  2. 상태검사서: 제 호(기록물형태)·검사일자·담당자 확인 / 일련번호·서고번호·서가번호(·앨범번호)·관리번호 / 검사내용 및 검사전등급(검사후등급)
     가. 종이류: 파손·잉크탈색·재질변색·건조화·재질오염·산성화
     나. 사진·필름류: (재질) 파손·감광유제층·긁힘·구겨짐·재질오염·수침
     다. 오디오류: (재질) 칼라·수축변형·부스러짐·도포층파괴·파손·긁힘·구겨짐·재질오염, 오디오신호 잡음
     라. 비디오·영화필름류: (재질) 칼라·수축변형·부스러짐·도포층파괴·파손·긁힘·구겨짐·재질오염, 잡음: 오디오·영상·제어신호
     마. 행정박물류: 파손·오염·변색·부식·그 밖의 사항
  ※ 다·라의 항목 묶음은 서식 본문 글자 순서와 괄호 수로 맞춘 것이다. 제출용은 기관이 쓰는 원본 서식(HWP)과 칸을 대조한다.
  전자기록물 등 보존매체(상태검사 5년, 별표 14)는 상태검사서에 칸이 없어 '보존매체' 시트로 따로 둔다.
근거: 시행규칙 제31조(정수점검·상태검사, 결과에 따른 사후조치)·별표 14·별지 제7호서식.
원칙: 작업지를 만들고 결과를 사본에 옮길 뿐, 점검·등급 판정·복원 조치는 사람이 한다. 원본 보존목록은 고치지 않는다.
"""
import argparse
import datetime as dt
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import inspection_schedule as isch  # noqa: E402

SECTIONS = {
    "가. 종이류": (["일련번호", "서고번호", "서가번호", "관리번호"], ["파손", "잉크탈색", "재질변색", "건조화", "재질오염", "산성화"]),
    "나. 사진·필름류": (["일련번호", "서고번호", "서가번호", "앨범번호", "관리번호"], ["파손", "감광유제층", "긁힘", "구겨짐", "재질오염", "수침"]),
    "다. 오디오류": (["일련번호", "서고번호", "서가번호", "관리번호"],
                  ["칼라", "수축변형", "부스러짐", "도포층파괴", "파손", "긁힘", "구겨짐", "재질오염", "오디오신호 잡음"]),
    "라. 비디오·영화필름류": (["일련번호", "서고번호", "서가번호", "관리번호"],
                      ["칼라", "수축변형", "부스러짐", "도포층파괴", "파손", "긁힘", "구겨짐", "재질오염", "잡음(오디오신호)", "잡음(영상신호)", "잡음(제어신호)"]),
    "마. 행정박물류": (["일련번호", "서고번호", "서가번호", "관리번호"], ["파손", "오염", "변색", "부식", "그 밖의 사항"]),
}
COUNT_HEAD = ["시스템번호", "관리번호", "철제목", "기록물유형", "소장위치", "보존상자번호", "육안상태점검", "점검목록 조치사항(삭제ㆍ수정ㆍ추가ㆍ기타 등)", "비고"]
EXTRA = {"시스템번호": ["시스템번호", "시스템 번호"], "상자": ["보존상자번호", "상자번호"], "앨범": ["앨범번호"]}


def section_of(t):
    s = re.sub(r"\s", "", str(t))
    k = isch.kind(t)
    if k == "종이":
        return "마. 행정박물류" if "박물" in s else "가. 종이류"
    if k == "사진필름":
        return "나. 사진·필름류"
    if k == "오디오비디오":
        return "다. 오디오류" if re.search(r"오디오|녹음|음성|카세트|릴", s) and not re.search(r"비디오|영상|녹화", s) else "라. 비디오·영화필름류"
    if k == "영화필름":
        return "라. 비디오·영화필름류"
    if k == "전자매체":
        return "보존매체"
    return None


def load(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    alias = {**isch.ALIAS, **EXTRA}
    names = {n for v in alias.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in alias.items() for n in al if n in head}
    items = [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]
    return items, hi, head, idx


def form(a):
    import openpyxl
    from openpyxl.styles import Alignment, Font
    items, _, _, _ = load(a.listing)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "1. 정수점검서"
    ws.append(["[별지 제7호서식] 보존기록물 정수점검ㆍ상태검사서 — 1. 보존기록물 정수점검서"])
    ws.append(["제     호", "", "점검일자", "", "담당자 확인", ""])
    ws.append([])
    ws.append(COUNT_HEAD)
    sec = {k: [] for k in SECTIONS}
    media, unknown, nc = [], [], 0
    for x in items:
        k = isch.kind(x.get("유형", ""))
        last_c, last_s = isch.year(x.get("정수")), isch.year(x.get("상태"))
        loc = " ".join(v for v in (x.get("서고", ""), x.get("서가", "")) if v)
        if a.all or last_c is None or last_c + isch.COUNT_CYCLE <= a.year:
            ws.append([x.get("시스템번호", ""), x.get("관리번호", ""), x.get("제목", ""), x.get("유형", ""), loc, x.get("상자", ""), "", "", ""])
            nc += 1
        s = section_of(x.get("유형", ""))
        if s is None:
            unknown.append(x.get("관리번호", ""))
            continue
        g = int(x["등급"]) if str(x.get("등급", "")).strip() in ("1", "2", "3") else None
        cyc = isch.STATE.get((k, g)) if k == "종이" else isch.STATE.get(k)
        due = a.all or last_s is None or (cyc is not None and last_s + cyc <= a.year) or (k == "종이" and g is None)
        if not due:
            continue
        if s == "보존매체":
            media.append([x.get("관리번호", ""), x.get("제목", ""), x.get("유형", ""), loc, last_s or "", "", ""])
        else:
            sec[s].append((x, g))
    for name, (ids, checks) in SECTIONS.items():
        w = wb.create_sheet(name.replace("·", "")[:28])
        w.append([f"2. 보존기록물 상태검사서 — {name}"])
        w.append(["제     호", "(기록물형태:      )", "검사일자", "", "담당자 확인", ""])
        w.append([])
        w.append(ids + ["검사내용 및 등급 →"] + [x for c in checks for x in (f"{c} 검사전", f"{c} 검사후")])
        for i, (x, g) in enumerate(sec[name], 1):
            vals = {"일련번호": i, "서고번호": x.get("서고", ""), "서가번호": x.get("서가", ""), "앨범번호": x.get("앨범", ""), "관리번호": x.get("관리번호", "")}
            w.append([vals[c] for c in ids] + [""] + [g if (j % 2 == 0 and g) else "" for j in range(len(checks) * 2)])
        w.freeze_panes = w.cell(row=5, column=len(ids) + 2)
        for c in w[4]:
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.font = Font(bold=True)
    wm = wb.create_sheet("보존매체")
    wm.append(["전자기록물 등 보존매체 — 상태검사 5년(별표 14). 별지 제7호 상태검사서에 칸이 없어 기관 방식으로 기록"])
    wm.append(["관리번호", "제목", "유형", "서고·서가", "최근 상태검사", "검사일자", "검사 결과"])
    for r in media:
        wm.append(r)
    for c in ws[4]:
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.font = Font(bold=True)
    for col, wdt in zip("ABCDEFGHI", (12, 18, 40, 12, 14, 12, 14, 30, 12)):
        ws.column_dimensions[col].width = wdt
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"정수점검 {nc}건 · 상태검사 " + " · ".join(f"{k.split('. ')[1]} {len(v)}" for k, v in sec.items()) + f" · 보존매체 {len(media)}"
          + (f" · 유형 불명 {len(unknown)}(작업지에 없음)" if unknown else "") + f" → {a.out}")
    print("검사전 등급은 보존목록의 상태등급을 미리 적은 것이다(종이·행정박물). 다·라 칸 묶음은 원본 서식과 대조할 것.")
    return 0


def apply(a):
    import openpyxl
    src = openpyxl.load_workbook(a.filled, data_only=True)
    count_dates, state = {}, {}

    def ydate(v):
        if isinstance(v, (dt.date, dt.datetime)):
            return v.year
        return isch.year(v)

    if "1. 정수점검서" in src.sheetnames:
        w = src["1. 정수점검서"]
        y = ydate(w["D2"].value)
        for r in w.iter_rows(min_row=5, values_only=True):
            if r[1] and (r[6] not in (None, "")):
                count_dates[str(r[1]).strip()] = y
    for name, (ids, checks) in SECTIONS.items():
        sn = name.replace("·", "")[:28]
        if sn not in src.sheetnames:
            continue
        w = src[sn]
        y = ydate(w["D2"].value)
        ki = ids.index("관리번호")
        for r in w.iter_rows(min_row=5, values_only=True):
            if not r[ki]:
                continue
            after = [r[len(ids) + 1 + 2 * j + 1] for j in range(len(checks))]
            grades = [int(v) for v in after if str(v or "").strip() in ("1", "2", "3")]
            bad = [checks[j] for j, v in enumerate(after) if str(v or "").strip() == "3"]
            if grades:
                state[str(r[ki]).strip()] = (y, max(grades), bad, name)
    if os.path.exists(a.out) or os.path.abspath(a.out) == os.path.abspath(a.listing):
        print("멈춤: 원본이나 이미 있는 파일에는 쓰지 않습니다.")
        return 2
    _, hi, head, idx = load(a.listing)
    wb = openpyxl.load_workbook(a.listing)
    ws = wb.worksheets[0]
    need = {"정수": "최근정수점검", "상태": "최근상태검사", "등급": "상태등급"}
    for k, label in need.items():
        if k not in idx:
            ws.cell(row=hi + 1, column=ws.max_column + 1, value=label)
            idx[k] = ws.max_column - 1
    no_col = idx.get("관리번호")
    changed, follow, missing_date = 0, [], 0
    for row in ws.iter_rows(min_row=hi + 2):
        no = str(row[no_col].value or "").strip() if no_col is not None else ""
        if not no:
            continue
        if no in count_dates:
            if count_dates[no]:
                row[idx["정수"]].value = count_dates[no]
                changed += 1
            else:
                missing_date += 1
        if no in state:
            y, g, bad, name = state[no]
            if y:
                row[idx["상태"]].value = y
            else:
                missing_date += 1
            if name in ("가. 종이류", "마. 행정박물류"):
                row[idx["등급"]].value = g
            changed += 1
            if bad:
                follow.append([no, name, ", ".join(bad), "사후조치(복원·탈산·재질 교체 등) 검토 — 시행규칙 제31조"])
    if follow:
        wf = wb.create_sheet("사후조치 후보")
        wf.append(["관리번호", "구분", "검사후 3등급 항목", "할 일"])
        for r in follow:
            wf.append(r)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"반영 {changed}칸(정수점검 {len(count_dates)}건·상태검사 {len(state)}건) · 사후조치 후보 {len(follow)}"
          + (f" · 점검·검사일자 칸이 비어 연도를 못 넣은 것 {missing_date}" if missing_date else "") + f" → {a.out}")
    print("다음: inspection_schedule.py로 갱신된 사본의 다음 점검 연도를 다시 계산한다.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest="cmd", required=True)
    f = sp.add_parser("form")
    f.add_argument("listing")
    f.add_argument("--year", type=int, required=True)
    f.add_argument("--all", action="store_true")
    f.add_argument("-o", "--out", required=True)
    p = sp.add_parser("apply")
    p.add_argument("filled")
    p.add_argument("listing")
    p.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    return form(a) if a.cmd == "form" else apply(a)


if __name__ == "__main__":
    sys.exit(main())
