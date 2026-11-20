#!/usr/bin/env python3
"""기록관리기준표(기록관리표) 점검: 단위과제별 관리항목이 NAK 4:2025(v2.3) 「기록관리기준표 작성 및 관리 절차」에 맞게 채워졌는지 본다.
지난해 표를 함께 주면 신설·변경 단위과제(고시 대상) 목록도 만든다.

사용:  python schedule_check.py 기준표.xlsx [--previous 지난해기준표.xlsx] -o 산출/기준표_점검.xlsx
열(머리글 이름으로 찾음): 단위과제코드, 단위과제명, 처리과, 소기능, 업무설명, 보존기간, 보존기간책정사유, 비치여부, 비치사유, 비치기간,
       보존장소, 보존방법, 공개여부, 비공개사유, 공개재검토시점, 접근권한
근거:  국가기록원고시 제2025-7호(2025. 11. 3. 시행) = NAK 4:2025(v2.3)
         4.3.1 업무설명 — 생산맥락을 알 수 있게 업무 성격·처리 절차를 기술
         4.3.2 보존기간 7종 중 하나, 책정 사유는 근거와 기준을 구체적으로
         4.3.3 비치기록물 — 비치여부(비치·미비치)·비치사유·비치기간(5년 범위, 연장 시 추가 기재)
         4.3.4 보존장소 — 30년 미만은 기록관·특수기록관, 30년 이상은 영구기록물관리기관(지정된 동종대량 등은 협의)
         4.3.5 보존방법 — 원본과 보존매체 함께 / 원본 그대로 / 원본 폐기·보존매체만 (시행령 제29조①·별표 2)
         4.3.6 공개여부 — 공개·부분공개·비공개, 부분공개·비공개는 정보공개법 비공개 사유·구체적 비공개 정보·공개 재검토 시점
         4.3.7 접근권한 — 목록 및 원문 열람 / 목록만 열람 / 열람 불가 (+출력·저장 가능 여부)
       시행령 제25조① <2026. 8. 25.> 단위과제별 기록관리표: 보존기간·책정 사유, 처리과, 소기능, 단위과제 개요
       시행령 제25조⑤ 신설·보존기간 변경 단위과제는 정리기간 직후 관보·공보·홈페이지에 고시
       시행령 제26조③ <2026. 8. 25.> 책정한 보존기간은 관할 영구기록물관리기관과 협의해 확정
원칙:  점검과 초안만 만든다. 보존기간 확정·고시는 기관이 절차에 따라 한다.
"""
import argparse
import collections
import os
import re
import sys

PERIODS = ["영구", "준영구", "30년", "10년", "5년", "3년", "1년"]
LONG = {"영구", "준영구", "30년"}
METHODS = {"원본과 보존매체 함께": ["원본과보존매체", "병행"], "원본 그대로": ["원본그대로", "원본을그대로", "원본보존"],
           "원본 폐기·보존매체만": ["보존매체만", "원본은폐기", "원본폐기"]}
ACCESS = {"목록 및 원문 열람": ["목록및원문", "원문열람", "목록·원문", "목록과원문"], "목록만 열람": ["목록만"], "열람 불가": ["열람불가", "불가"]}
ALIAS = {"코드": ["단위과제코드", "단위과제ID", "코드"], "명": ["단위과제명", "단위과제"], "처리과": ["처리과", "부서"], "소기능": ["소기능", "소기능명"],
         "설명": ["업무설명", "단위과제 업무설명", "단위과제개요", "개요"], "보존기간": ["보존기간"], "사유": ["보존기간책정사유", "책정사유", "보존기간 책정 사유"],
         "비치": ["비치여부", "비치"], "비치사유": ["비치사유"], "비치기간": ["비치기간"], "장소": ["보존장소"], "방법": ["보존방법"],
         "공개": ["공개여부", "공개구분"], "비공개사유": ["비공개사유", "비공개 사유"], "재검토": ["공개재검토시점", "재검토시점", "공개 재검토 시점"],
         "접근": ["접근권한"]}
BIG = re.compile(r"예비타당성|예타|민간투자|민자사업|BTL|BTO|대규모\s*투자|총사업비\s*[3-9]\d{2}\s*억|총사업비\s*\d[,\d]*\s*천억|총사업비\s*\d+\s*조")
VAGUE = re.compile(r"^(업무\s*참고|참고용|필요|중요|기타|해당\s*없음|관련\s*기록|-)$")


def norm(s):
    return re.sub(r"\s", "", str(s or ""))


def read(path):
    import openpyxl
    if path.lower().endswith(".csv"):
        import csv
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    else:
        rows = [["" if v is None else str(v).strip() for v in r] for r in
                openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head:
                idx[k] = head.index(n)
                break
    return [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)], idx


def match(value, table):
    v = norm(value)
    return next((k for k, keys in table.items() if any(x in v for x in keys)), None)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--previous")
    ap.add_argument("--min-desc", type=int, default=20, help="업무설명 최소 글자 수")
    ap.add_argument("-o", "--out", default="산출/기준표_점검.xlsx")
    a = ap.parse_args()
    rows, idx = read(a.table)
    if "보존기간" not in idx or ("코드" not in idx and "명" not in idx):
        print("입력 오류: 단위과제(코드 또는 명)와 보존기간 열이 필요합니다.")
        return 2
    issues = []

    def add(n, r, item, msg, basis):
        issues.append([n, r.get("코드", ""), r.get("명", ""), item, msg, basis])

    for n, r in rows:
        p = norm(r.get("보존기간"))
        if p not in PERIODS:
            add(n, r, "보존기간", f"'{r.get('보존기간')}' — 7종 중 하나가 아님", "NAK 4 4.3.2 · 시행령 제26조①")
        if "사유" in idx:
            s = str(r.get("사유", "")).strip()
            if not s:
                add(n, r, "책정 사유", "비어 있음", "NAK 4 4.3.2 · 시행령 제25조①1")
            elif VAGUE.match(s) or len(norm(s)) < 8:
                add(n, r, "책정 사유", f"'{s}' — 근거·기준이 구체적이지 않음(별표 1 해당 호·관계 법령 등)", "NAK 4 4.3.2")
        if "설명" in idx and len(norm(r.get("설명"))) < a.min_desc:
            add(n, r, "업무설명", f"{len(norm(r.get('설명')))}자 — 업무 성격·처리 절차가 드러나지 않음", "NAK 4 4.3.1")
        if "처리과" in idx and not r.get("처리과"):
            add(n, r, "처리과", "비어 있음", "시행령 제25조①2")
        if "소기능" in idx and not r.get("소기능"):
            add(n, r, "소기능", "비어 있음", "시행령 제25조①3")
        if "비치" in idx:
            b = norm(r.get("비치"))
            if b not in ("비치", "미비치", "Y", "N", "예", "아니오"):
                add(n, r, "비치여부", f"'{r.get('비치')}' — 비치·미비치로 구분", "NAK 4 4.3.3")
            if b in ("비치", "Y", "예"):
                if "비치사유" in idx and not r.get("비치사유"):
                    add(n, r, "비치사유", "비치인데 사유 없음", "NAK 4 4.3.3")
                m = re.search(r"(\d+)", str(r.get("비치기간", "")))
                if not m:
                    add(n, r, "비치기간", "비치인데 기간 없음", "NAK 4 4.3.3 · 시행령 제31조②")
                elif int(m[1]) > 5 and "연장" not in str(r.get("비치기간", "")):
                    add(n, r, "비치기간", f"{m[1]}년 — 5년 범위 초과(연장이면 연장 기간을 따로 적음)", "시행령 제31조②")
        if "장소" in idx and p in PERIODS:
            place = norm(r.get("장소"))
            want_perm = p in LONG
            is_perm = "영구" in place
            if not place:
                add(n, r, "보존장소", "비어 있음", "NAK 4 4.3.4")
            elif want_perm and not is_perm and "협의" not in place and "지정" not in place:
                add(n, r, "보존장소", f"보존기간 {p}인데 '{r.get('장소')}' — 30년 이상은 영구기록물관리기관(지정·협의 예외)", "NAK 4 4.3.4")
            elif not want_perm and is_perm:
                add(n, r, "보존장소", f"보존기간 {p}인데 영구기록물관리기관 — 30년 미만은 기록관·특수기록관", "NAK 4 4.3.4")
        if "방법" in idx and not match(r.get("방법"), METHODS):
            add(n, r, "보존방법", f"'{r.get('방법')}' — 시행령 제29조① 3가지 중 하나가 아님", "NAK 4 4.3.5 · 시행령 별표 2")
        if "공개" in idx:
            d = norm(r.get("공개"))
            if d not in ("공개", "부분공개", "비공개"):
                add(n, r, "공개여부", f"'{r.get('공개')}' — 공개·부분공개·비공개로 구분", "NAK 4 4.3.6")
            elif d != "공개":
                if not re.search(r"[1-8]\s*호", str(r.get("비공개사유", ""))):
                    add(n, r, "공개여부", "부분공개·비공개인데 정보공개법 제9조① 호가 없음", "NAK 4 4.3.6")
                if "재검토" in idx and not r.get("재검토"):
                    add(n, r, "공개여부", "공개 재검토 시점 없음", "NAK 4 4.3.6")
        txt = f"{r.get('명', '')} {r.get('설명', '')}"
        if BIG.search(txt) and p in PERIODS and p not in ("영구", "준영구"):
            add(n, r, "보존기간", f"대규모 투자사업으로 보이는데 {p} — '사업' 단위 단위과제로 준영구 이상 책정(예타 통과·민투심 심의·지자체 총사업비 300억 이상)",
                "NAK 4 4.3.2 비고 · 행정안전부 재정정책과-3671(2012.8.9.)")
        if "접근" in idx and not match(r.get("접근"), ACCESS):
            add(n, r, "접근권한", f"'{r.get('접근')}' — 목록 및 원문 열람 / 목록만 열람 / 열람 불가 중 하나가 아님", "NAK 4 4.3.7")

    notice = []
    if a.previous:
        prev, _ = read(a.previous)
        key = "코드" if "코드" in idx else "명"
        pm = {r.get(key): r for _, r in prev}
        for n, r in rows:
            old = pm.get(r.get(key))
            if old is None:
                notice.append([r.get("코드", ""), r.get("명", ""), "신설", "", r.get("보존기간", ""), r.get("설명", "")])
            elif norm(old.get("보존기간")) != norm(r.get("보존기간")):
                notice.append([r.get("코드", ""), r.get("명", ""), "보존기간 변경", old.get("보존기간", ""), r.get("보존기간", ""), r.get("설명", "")])

    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "요약"
    s.append(["기록관리기준표 점검 (NAK 4:2025 v2.3 · 국가기록원고시 제2025-7호)"])
    s["A1"].font = Font(bold=True, size=12)
    s.append(["단위과제 수", len(rows)])
    c = collections.Counter(i[3] for i in issues)
    for k, v in c.most_common():
        s.append([k, v])
    s.append(["고시 대상(신설·보존기간 변경)", len(notice) if a.previous else "지난해 표 없음(--previous)"])
    d = wb.create_sheet("점검 결과")
    d.append(["원본 행", "단위과제코드", "단위과제명", "항목", "내용", "근거"])
    for i in issues:
        d.append(i)
    if a.previous:
        g = wb.create_sheet("고시 대상")
        g.append(["단위과제코드", "단위과제명", "구분", "이전 보존기간", "보존기간", "단위과제 업무설명"])
        for i in notice:
            g.append(i)
        g.append([])
        g.append(["신설·변경 단위과제 보존기간은 매년 10월 31일까지 관할 영구기록물관리기관과 협의·확정(11월 1일 이후 사유는 발생 즉시 협의), 영구기록물관리기관은 12월 31일까지 회신 — NAK 4 5.3.1·5.3.2"])
        g.append(["정리기간 종료 직후 관보(지자체는 공보) 또는 홈페이지에 고시 — 시행령 제25조⑤. 보존기간은 관할 영구기록물관리기관과 협의 후 확정 — 시행령 제26조③"])
    for sh in wb.worksheets:
        for i, wdt in enumerate((14, 16, 28, 14, 60, 30)):
            sh.column_dimensions[chr(65 + i)].width = wdt
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"단위과제 {len(rows)} · 점검 {len(issues)}" + (f" · 고시 대상 {len(notice)}" if a.previous else "") + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
