#!/usr/bin/env python3
"""이관 시기 관리: 기록물철 목록으로 처리과→기록관, 기록관→영구기록물관리기관 이관 연도를 계산하고,
기한이 지났거나 올해 이관할 것, 연장·연기가 필요한 것, 비치기록물 기간 만료·연장 신청 기한을 나눈다.
연장이 필요한 철은 시행규칙 별지 제1호의2 「기록물 이관 기간 연장ㆍ일정 연기 목록」 칸으로 옮긴다.

사용:  python transfer_schedule.py 기록물철목록.xlsx --year 2026 [--rule auto|current|2027] [--org 기록관|미설치]
         [--today 2026-11-20] [--org-name "○○시"] -o 산출/이관일정.xlsx
목록 열(머리글 이름으로 찾음): 철분류번호(또는 등록번호), 철제목, 처리과(생산부서), 처리과코드, 생산연도, 종료연도,
       보존기간, 기록물형태(전자·비전자), 비치여부, 비치종료일(또는 비치기간), 기록관이관일, 영구이관일,
       연장여부, 이관희망시기, 연장사유, 문서등록번호(건 단위일 때)
근거(law.go.kr 2026-11-20 확인):
  시행령 제26조③ 기산일 = 종료(생산)된 다음 연도 1월 1일
  시행령 제32조 처리과→기록관
    현행(2026. 12. 31.까지): ① 매년 전년도 생산·접수분 이관, 비전자는 기산일부터 2년 범위 보관 후 철 단위 이관 가능
                             ③ 기록관 미설치·수시 참고 필요 시 기산일부터 10년 범위 연장
    [시행 2027. 1. 1.]:     ① 기산일부터 1년 안 이관, 부득이하면 처리과의 장이 기록관과 협의해 1년 범위 연장
                             ③ 기록관·특수기록관 미설치 공공기관은 기산일부터 1년 안, 부득이하면 5년 범위 연장
  시행규칙 제20조 [시행 2027. 1. 1.] 연장 협의 때 별지 제1호의2 목록 첨부(현행은 별지 제3호 이관시기 연장신청서)
  시행령 제31조 비치기록물: 5년 범위 지정, 끝나기 60일 전까지 연장 신청, 종료되면 다음 연도 중 기록관 이관,
    기록관은 기산일+5년 지난 30년 이상을 인수한 다음 연도 중 영구로
  시행령 제40조 기록관→영구: 30년 이상(30년·준영구·영구), 기산일부터 5년 지난 다음 연도 중(이관연도 = 종료연도+6),
    제시 일정 연기는 제시받은 날부터 30일 안 협의, 최대 1년(③). 제3조 공공기관은 자체 보존(④)
원칙:  계산과 목록 작성만 한다. 연장 협의·이관 실행은 사람이 한다. 부칙의 경과규정은 원문으로 확인한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"철번호": ["철분류번호", "기록물철분류번호", "분류번호", "등록번호", "관리번호"], "제목": ["철제목", "기록물철제목", "제목"],
         "처리과": ["처리과", "생산부서", "생산부서(처리과)명", "처리과명"], "처리과코드": ["처리과코드", "처리과기관코드", "기관코드"],
         "생산연도": ["생산연도", "생산년도"], "종료연도": ["종료연도", "종료년도", "철종료연도"], "보존기간": ["보존기간"],
         "형태": ["기록물형태", "형태", "전자여부", "전자/비전자"], "비치": ["비치여부", "비치"], "비치종료": ["비치종료일", "비치기간종료일", "비치만료일"],
         "비치기간": ["비치기간"], "기록관이관": ["기록관이관일", "기록관이관일자", "처리과이관일", "이관일"],
         "영구이관": ["영구이관일", "영구기록물관리기관이관일", "영구기관이관일"], "연장": ["연장여부", "연장"],
         "희망": ["이관희망시기", "이관희망연도"], "사유": ["연장사유", "사유"], "문서번호": ["문서등록번호", "건등록번호"]}
LONG = ("영구", "준영구", "30년")
TODAY_RULE_SWITCH = dt.date(2027, 1, 1)


def year(v):
    m = re.search(r"(19|20)\d{2}", str(v or ""))
    return int(m[0]) if m else None


def date(v):
    s = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y%m%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def is_yes(v):
    return str(v or "").strip() in ("비치", "Y", "y", "예", "O", "o", "연장", "1")


def nonelec(v):
    s = str(v or "")
    return "비전자" in s or s.strip().upper() in ("N", "비")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, help="기준 연도(기본: 오늘의 연도)")
    ap.add_argument("--today", help="판정일(기본: 오늘)")
    ap.add_argument("--rule", choices=["auto", "current", "2027"], default="auto",
                    help="처리과→기록관 규정: current(2026. 12. 31.까지), 2027(2027. 1. 1. 시행), auto(판정일로 고름)")
    ap.add_argument("--org", choices=["기록관", "미설치"], default="기록관", help="기록관(특수기록관)이 있는 기관인지")
    ap.add_argument("--org-name", default="", help="별지 제1호의2의 기관명")
    ap.add_argument("-o", "--out", default="산출/이관일정.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    yr = a.year or today.year
    rule = a.rule if a.rule != "auto" else ("2027" if today >= TODAY_RULE_SWITCH else "current")

    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head and k not in idx:
                idx[k] = head.index(n)
    if "종료연도" not in idx and "생산연도" not in idx:
        print("종료연도 또는 생산연도 열이 필요합니다.", file=sys.stderr)
        return 2
    items = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]

    wb = openpyxl.Workbook()
    sch = wb.active
    sch.title = "이관 일정"
    sch.append(["원본 행", "철분류번호", "철제목", "처리과", "보존기간", "형태", "종료연도", "기산일",
                "기록관 이관 기한(현행)", "기록관 이관 기한(2027~)", "적용 규정", "기록관 이관 상태", "영구 이관 연도", "영구 이관 상태", "확인할 것"])
    ext = wb.create_sheet("연장 목록(별지 1의2)")
    ext.append(["■ 공공기록물 관리에 관한 법률 시행규칙 [별지 제1호의2서식] — 기록물 이관 기간 연장·일정 연기 목록 (시행 2027. 1. 1.)"])
    ext.append([f"기관명 : {a.org_name}"])
    ext.append(["일련번호", "처리과 기관코드", "기록물철 분류번호", "문서등록번호*", "기록물철(건) 제목", "생산연도", "이관희망시기", "사유"])
    bichi = wb.create_sheet("비치기록물")
    bichi.append(["원본 행", "철분류번호", "철제목", "처리과", "비치 종료일", "연장 신청 기한(60일 전)", "상태", "확인할 것"])
    counts = {"기한 지남": 0, f"{yr}년 이관": 0, "연장 필요": 0, "영구 지남": 0, f"영구 {yr}년": 0, "비치 신청 임박": 0}
    ext_n = 0
    for n, x in items:
        notes = []
        end = year(x.get("종료연도")) or year(x.get("생산연도"))
        if not x.get("종료연도"):
            notes.append("종료연도 없음 — 생산연도로 계산")
        if not end:
            sch.append([n, x.get("철번호", ""), x.get("제목", ""), x.get("처리과", ""), x.get("보존기간", ""), x.get("형태", ""),
                        "", "", "", "", rule, "확인 필요", "", "", "종료·생산연도를 읽지 못함"])
            continue
        base = dt.date(end + 1, 1, 1)
        ne = nonelec(x.get("형태"))
        # 처리과 → 기록관
        cur_due = dt.date(end + (2 if ne else 1), 12, 31)            # 현행: 전년도분 이관, 비전자는 기산일부터 2년 범위 보관
        new_due = dt.date(base.year, 12, 31)                          # 2027~: 기산일부터 1년 안
        cur_max = dt.date(base.year + 9, 12, 31)                     # 현행 ③ 미설치·수시 참고 시 기산일부터 10년 범위 연장
        new_max = dt.date(base.year + (5 if a.org == "미설치" else 1), 12, 31)
        due, limit = (new_due, new_max) if rule == "2027" else (cur_due, cur_max)
        moved = date(x.get("기록관이관"))
        bichi_on = is_yes(x.get("비치"))
        bend = date(x.get("비치종료"))
        if bichi_on and not bend:
            notes.append("비치인데 비치 종료일 없음 — 지정 통보의 비치기간 끝나는 날을 적으세요")
        state = ""
        if moved:
            state = "이관됨"
        elif bichi_on:
            if bend:
                due = dt.date(bend.year + 1, 12, 31)                   # 시행령 제31조⑥ 종료되면 다음 연도 중
                state = "비치 중" if bend >= today else ("기한 지남" if due < today else f"{due.year}년 이관")
            else:
                state = "비치 중(종료일 미상)"
        else:
            if due < today:
                want = x.get("희망", "")
                if is_yes(x.get("연장")) or want:
                    wy = year(want)
                    if wy and dt.date(wy, 12, 31) > limit:
                        state = "연장 한도 넘음"
                        notes.append(f"이관희망시기 {wy} — {'2027 규정은 1년' if rule == '2027' else '현행은 기산일부터 10년'} 범위({limit.year}년)까지"
                                     + (" · 미설치 기관은 5년" if rule == "2027" and a.org == "미설치" else ""))
                    else:
                        state = "연장 중"
                else:
                    state = "기한 지남"
            elif due.year == yr:
                state = f"{yr}년 이관"
            else:
                state = f"{due.year}년 이관 예정"
        if rule == "2027" and not moved and not bichi_on and cur_due > new_due and new_due < TODAY_RULE_SWITCH <= cur_due + dt.timedelta(days=1):
            notes.append(f"현행 규정으로는 {cur_due.isoformat()}까지(비전자 2년 보관) — 개정 부칙 경과규정 확인")
        if state in ("기한 지남", "연장 한도 넘음") or (state == "연장 중" and rule == "2027"):
            counts["연장 필요"] += state != "연장 중"
            ext_n += 1
            ext.append([ext_n, x.get("처리과코드", ""), x.get("철번호", ""), x.get("문서번호", ""), x.get("제목", ""),
                        year(x.get("생산연도")) or "", x.get("희망", ""), x.get("사유", "") or "(사유를 적으세요 — 부득이한 사유)"])
        if state == "기한 지남":
            counts["기한 지남"] += 1
        if state == f"{yr}년 이관":
            counts[f"{yr}년 이관"] += 1
        # 비치기록물 시트
        if bichi_on:
            ask = bend - dt.timedelta(days=60) if bend else None
            bstate = ""
            if bend:
                if bend < today:
                    bstate = "비치 종료 — 다음 연도 중 기록관 이관(제31조⑥)"
                elif ask <= today:
                    bstate = "연장 신청 기한 지남(종료 60일 전)"
                elif (ask - today).days <= 30:
                    bstate = f"연장 신청 {(ask - today).days}일 남음"
                    counts["비치 신청 임박"] += 1
                else:
                    bstate = "비치 중"
            bn = []
            m = re.search(r"(\d+)\s*년", x.get("비치기간", ""))
            if m and int(m[1]) > 5:
                bn.append(f"비치기간 {m[1]}년 — 지정은 5년 범위(제31조②), 연장은 직전 기간 범위(④)")
            bichi.append([n, x.get("철번호", ""), x.get("제목", ""), x.get("처리과", ""), bend.isoformat() if bend else "",
                          ask.isoformat() if ask else "", bstate or "종료일 필요", " · ".join(bn)])
        # 기록관 → 영구
        ret = re.sub(r"\s", "", x.get("보존기간", ""))
        ty, tstate = "", ""
        if any(ret.startswith(k) for k in LONG):
            ty = end + 6
            if bichi_on and moved:
                ty = max(moved.year + 1, ty)                           # 제31조⑦ 인수한 다음 연도 중
            p_moved = date(x.get("영구이관"))
            if bichi_on and not moved and "기록관이관" in idx:
                ty, tstate = "", "비치 중 — 기록관이 인수한 다음 연도 중(제31조⑦)"
            elif p_moved:
                tstate = "이관됨"
            elif ty < yr:
                tstate = "지남 — 제시 일정·연기 협의(제40조③, 최대 1년) 확인"
                counts["영구 지남"] += 1
            elif ty == yr:
                tstate = f"{yr}년 대상 — 영구기록물관리기관이 제시한 일정"
                counts[f"영구 {yr}년"] += 1
            else:
                tstate = "예정"
        elif ret and not re.fullmatch(r"(10|5|3|1)년", ret):
            notes.append(f"보존기간 '{x.get('보존기간')}' — 7종 중 하나로")
        sch.append([n, x.get("철번호", ""), x.get("제목", ""), x.get("처리과", ""), x.get("보존기간", ""), "비전자" if ne else "전자",
                    end, base.isoformat(), cur_due.isoformat(), new_due.isoformat(), "2027~" if rule == "2027" else "현행", state,
                    ty, tstate, " · ".join(notes)])
    if ext_n == 0:
        ext.append(["해당 없음"])
    ext.append(["* \"건\"단위로 기록물을 이관하는 경우에만 기재"])
    base_ws = wb.create_sheet("기준")
    for line in [
        ["판정일", today.isoformat(), "적용 규정", "2027. 1. 1. 시행 규정" if rule == "2027" else "현행(2026. 12. 31.까지)"],
        ["기산일", "종료(생산)된 다음 연도 1월 1일", "시행령 제26조③", ""],
        ["처리과→기록관(현행)", "매년 전년도분. 비전자는 기산일부터 2년 범위 보관 후 철 단위. 기록관 미설치·수시 참고 시 기산일부터 10년 범위 연장", "시행령 제32조①③", "연장신청: 시행규칙 별지 제3호(현행 제20조)"],
        ["처리과→기록관(2027~)", "기산일부터 1년 안. 부득이하면 기록관과 협의해 1년 범위 연장(미설치 기관은 5년)", "시행령 제32조①③ [시행 2027. 1. 1.]", "협의 때 별지 제1호의2 목록 첨부(시행규칙 제20조)"],
        ["비치기록물", "5년 범위 지정, 끝나기 60일 전까지 연장 신청, 종료되면 다음 연도 중 기록관 이관", "시행령 제31조", ""],
        ["기록관→영구", "30년 이상, 기산일부터 5년 지난 다음 연도 중(이관연도 = 종료연도+6), 제시 일정 연기는 30일 안 협의·최대 1년", "시행령 제40조", "공공기관(제3조)은 자체 보존"],
        ["암호화 기록물", "생산·접수 때 암호화한 기록물은 목록과 해제 정보를 첨부해 이관 — encrypted_scan.py", "시행령 제31조의2 [시행 2027. 1. 1.]", ""],
        ["경과규정", "개정 부칙의 경과규정은 법령 원문으로 확인", "", ""],
    ]:
        base_ws.append(line)
    for ws, widths in ((sch, (8, 22, 36, 14, 8, 8, 8, 12, 14, 14, 8, 18, 10, 30, 50)), (ext, (8, 14, 22, 16, 40, 10, 14, 40)),
                       (bichi, (8, 22, 36, 14, 12, 18, 34, 50)), (base_ws, (20, 70, 30, 40))):
        for i, w in enumerate(widths):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(items)}철 · 규정 {'2027~' if rule == '2027' else '현행'} · 기록관 이관 기한 지남 {counts['기한 지남']} · {yr}년 이관 {counts[f'{yr}년 이관']}"
          f" · 연장 목록 {ext_n} · 영구 이관 지남 {counts['영구 지남']} · 영구 {yr}년 대상 {counts[f'영구 {yr}년']}"
          f" · 비치 연장 신청 임박 {counts['비치 신청 임박']} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
