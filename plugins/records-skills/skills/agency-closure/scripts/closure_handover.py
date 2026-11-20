#!/usr/bin/env python3
"""한시·폐지기관 기록물 인계인수 준비: 기관 폐지 유형에 따라 누가 기록물을 관리하는지 정하고, 유형별 이관목록(부속서 C),
기록물 인계인수서(부속서 B) 수량, 단계별 체크리스트(부속서 A), 일정(인수 종료 뒤 3개월 안 통보)을 만든다.
근거: NAK 15:2021(v2.0) 「한시 및 폐지기관 기록관리」 (국가기록원, 2026-11-20 내려받아 확인), 공공기록물법 제25조

사용:  python closure_handover.py 기록물목록.xlsx --type 폐지-승계없음 --org "○○위원회" --to "○○기록원" [--done 2026-12-15] -o 산출/인계인수.xlsx
유형(--type): 통합 · 한시 · 폐지-승계 · 폐지-승계없음 · 일부승계 · 공공기관전환 · 민간전환
입력 열(머리글 이름으로 찾음): 생산기관, 처리과, 기록물철제목(행정박물명), 수량, 생산연도, 보존기간, 기록물형태(문서·카드·도면·사진·시청각·행정박물·전자),
       공개여부, 비공개사유, 세부매체, 저작권동의, 유형·형태·재질·크기(행정박물), 기산일(공공기관 전환)
원칙:  목록과 서식만 만든다. 인계인수 실행·폐기는 하지 않는다. 승계기관이 없으면 보존기간 10년 이하도 이관 대상이므로 폐기하지 말라고 표시한다.
"""
import argparse
import datetime as dt
import os
import re
import sys

WHO = {"통합": "통합기관이 관리(전자문서·업무관리시스템 권한과 기록물은 통합기관 기록관리시스템으로, 웹기록물·데이터세트·비전자는 인계목록과 함께)",
       "한시": "업무 승계기관으로 인계, 없으면 관할 영구기록물관리기관으로 이관(활동 종료 전 정리·분류·이관 사전 협의)",
       "폐지-승계": "업무 승계기관이 관리 — 인계절차 착수 전까지 인계계획을 중앙기록물관리기관에 통보, 승계기관은 인수 종료 시 결과 통보",
       "폐지-승계없음": "보존기간에 관계없이 기록물 일체를 지체 없이 관할 영구기록물관리기관으로 이관(활동 종료 전 사전 협의)",
       "일부승계": "업무추진에 필요한 최소한의 기록물만 승계기관으로 — 목록은 생산기관명으로 작성해 관할 영구기록물관리기관에 통보, 재정리 불가, 인계 뒤 사본 처리",
       "공공기관전환": "보존기간 30년 이상(국가·지자체 기관일 때 생산분)은 기산일부터 10년이 지난 다음 연도 중 관할 영구기록물관리기관으로 이관",
       "민간전환": "모든 기록물을 지체 없이 관할 영구기록물관리기관으로 이관(공공기관이 민간으로 전환되면 30년 이상 중 영구기록물관리기관이 지정한 것)"}
ALIAS = {"기관": ["생산기관"], "처리과": ["처리과", "생산부서"], "제목": ["기록물철제목", "철제목", "제목", "행정박물명"], "수량": ["수량", "쪽수", "권수"],
         "연도": ["생산연도", "생산년도"], "보존": ["보존기간"], "형태": ["기록물형태", "형태구분", "형태"], "공개": ["공개여부", "공개구분"],
         "사유": ["비공개사유", "비공개 사유"], "매체": ["세부매체"], "저작권": ["저작권동의", "저작권 동의 여부", "저작권동의여부"],
         "유형": ["유형"], "재질": ["재질"], "크기": ["크기"], "기산": ["기산일"]}
CHECK = {
    "A.1 사전 조치": ["기록물 생산업무 종결 처리 및 각 처리과 고지", "인계인수 대상 기록물 확인(종이·시청각·행정박물)", "인계인수 대상 목록 작성(부속서 C)",
                   "각 처리과 사무실 보유 기록물 임의 폐기 금지 고지", "처리과 미정리 기록물 현황 파악 및 기록관 이관 여부", "DB 수록 완료 기록물 목록·실물 확인",
                   "인계인수서 내역 작성(부속서 B: 전자문서·데이터세트·종이·시청각·행정박물)"],
    "A.2 기록관리기준 정비": ["구 단위과제별 과제관리카드 종결", "기관코드 정비", "보존기간 미확정 단위과제 점검·확정·시스템 반영",
                       "신규 단위과제 신설 확인·보존기간 책정·영구기관 협의·고시", "신규 단위과제 과제관리카드 생성"],
    "A.3 비전자 인계인수": ["관계기관(기록관·인수기관·관할 영구기관)에 폐지 상황 알림", "인계인수 일정 협의", "유형별 정리 및 이관목록(부속서 C)",
                       "승계기관이 있으면 인계 착수 전 인계계획을 중앙기록물관리기관에 통보(공문)", "이관목록을 붙여 이관 협조 요청(공문) 후 인계인수"],
    "A.4 전자 인계인수": ["기록관·정보화 담당 협업팀 구성", "인계인수 일정 협의", "문서관리카드 완료·결재 완료 처리", "대상 확인(온나라 과제·문서카드·메모보고, RMS, 디지털 시청각, 데이터세트)",
                      "업무관리시스템 일체 → RMS 자체 이관", "인수기관 기록관·정보화 담당 확인", "기록물 인수 H/W 증설", "인계기관 RMS → 인수기관 RMS",
                      "원생산기관 정보 유지 확인", "인계기관 업무관리시스템의 업무 활용 기록물은 '사본' 처리", "검수·인계인수서 작성", "결과 통보(공문)"],
    "A.5 인계인수 후": ["원생산기관 명의 정리·재편철 여부", "미정리 인수기록물 확인·정리계획", "인수 종료 뒤 3개월 안 인계인수서 사본·인수 목록을 중앙기록물관리기관에 통보",
                     "인수 기록물 중 30년 이상은 기산일부터 10년 지나면 영구기관 이관, 10년 이하는 심사·심의 거쳐야 폐기(9.2)"],
}


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def group(f):
    f = re.sub(r"\s", "", f or "")
    if re.search(r"행정박물|박물|관인|현판|깃발|선물", f):
        return "행정박물"
    if re.search(r"시청각|사진|필름|영상|음성|녹음|테이프", f):
        return "시청각"
    return "문서"


def num(v):
    m = re.search(r"\d+", v or "")
    return int(m[0]) if m else 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--type", required=True, choices=list(WHO))
    ap.add_argument("--org", default="", help="한시·폐지기관명")
    ap.add_argument("--to", default="", help="인수기관명(승계기관·통합기관·관할 영구기록물관리기관)")
    ap.add_argument("--closing", default="", help="폐지(활동 종료) 일자")
    ap.add_argument("--done", help="인수 종료일 YYYY-MM-DD — 3개월 안 통보 기한 계산")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("-o", "--out", default="산출/인계인수.xlsx")
    a = ap.parse_args()
    rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    sm = wb.active
    sm.title = "요약"
    sm.append(["폐지 유형", a.type])
    sm.append(["기록물 관리 주체·조치(NAK 15 5장)", WHO[a.type]])
    sm.append(["한시·폐지기관", a.org])
    sm.append(["인수기관", a.to])
    sheets = {
        "문서": ["일련번호", "생산기관", "처리과", "기록물철 제목", "수량", "생산연도", "보존기간", "기록물형태", "공개여부", "비공개사유", "확인할 것"],
        "시청각": ["일련번호", "생산기관", "처리과", "기록물철 제목", "수량", "생산연도", "보존기간", "기록물형태", "세부매체", "공개여부", "비공개사유", "저작권 동의 여부", "확인할 것"],
        "행정박물": ["일련번호", "생산기관", "처리과", "행정박물명", "수량", "생산연도", "유형", "형태", "재질", "크기", "확인할 것"]}
    ws = {}
    for k, h in sheets.items():
        ws[k] = wb.create_sheet(f"이관목록-{k}(부속서 C)")
        ws[k].append(h)
    counts = {"문서(철)": 0, "카드(권/매)": 0, "도면(권/매)": 0, "사진(점)": 0, "시청각 매체": 0, "행정박물": 0}
    issues = 0
    later = []
    for x in rows:
        g = group(x.get("형태"))
        notes = []
        ret = re.sub(r"\s", "", x.get("보존", ""))
        if g != "행정박물" and not ret:
            notes.append("보존기간 없음")
        if re.search(r"비공개|부분", x.get("공개", "")) and not x.get("사유"):
            notes.append("비공개·부분공개인데 사유 없음(정보공개법 제9조①각 호)")
        if g == "시청각" and not x.get("저작권"):
            notes.append("저작권 동의 여부 빈칸")
        if a.type == "폐지-승계없음" and ret in ("1년", "3년", "5년", "10년"):
            notes.append("승계기관 없음 — 10년 이하도 이관 대상, 영구기관 협의 없이 폐기 금지(5.4.2)")
        if a.type == "공공기관전환" and ret in ("30년", "준영구", "영구"):
            m = re.search(r"(19|20)\d{2}", x.get("기산", "") or x.get("연도", ""))
            if m:
                due = int(m[0]) + 10 + (0 if x.get("기산") else 1)
                later.append((x.get("제목", ""), due))
                notes.append(f"{due}년 중 영구기관 이관(기산일부터 10년 지난 다음 연도)")
        issues += bool(notes)
        n = len([r for r in ws[g].iter_rows(min_row=2)]) + 1
        if g == "문서":
            line = [n, x.get("기관", ""), x.get("처리과", ""), x.get("제목", ""), x.get("수량", ""), x.get("연도", ""), x.get("보존", ""), x.get("형태", ""),
                    x.get("공개", ""), x.get("사유", ""), " · ".join(notes)]
            f = x.get("형태", "")
            key = "카드(권/매)" if "카드" in f else "도면(권/매)" if "도면" in f else "문서(철)"
            counts[key] += 1
        elif g == "시청각":
            line = [n, x.get("기관", ""), x.get("처리과", ""), x.get("제목", ""), x.get("수량", ""), x.get("연도", ""), x.get("보존", ""), x.get("형태", ""),
                    x.get("매체", ""), x.get("공개", ""), x.get("사유", ""), x.get("저작권", ""), " · ".join(notes)]
            counts["사진(점)" if "사진" in x.get("형태", "") else "시청각 매체"] += num(x.get("수량")) if "사진" in x.get("형태", "") else 1
        else:
            line = [n, x.get("기관", ""), x.get("처리과", ""), x.get("제목", ""), x.get("수량", ""), x.get("연도", ""), x.get("유형", ""), x.get("형태", ""),
                    x.get("재질", ""), x.get("크기", ""), " · ".join(notes)]
            counts["행정박물"] += num(x.get("수량"))
        ws[g].append(line)
        if notes:
            ws[g].cell(ws[g].max_row, len(line)).fill = red
    hb = wb.create_sheet("인계인수서(부속서 B)")
    hb.append(["기록물 인계인수서"])
    hb.append(["○ 기관명", a.org])
    hb.append(["○ 이관일자", ""])
    hb.append(["○ 이관수량", f"{len(rows)}건(행)"])
    hb.append([f"{a.org or '○○기관'}이 {a.closing or '○○년○○월○○일'}자로 폐지됨에 따라 {a.org or '○○기관'}의 위 기록물을 {a.to or '○○기관'}으로 정히 인계·인수하였음을 확인함."])
    hb.append(["", "기관", "직급", "성명", "(서명)"])
    for r_ in ("인계자", "인수자", "입회자"):
        hb.append([r_, "", "", "", ""])
    hb.append([])
    hb.append(["유형", "계", *counts.keys()])
    hb.append(["수량", sum(counts.values()), *counts.values()])
    ck = wb.create_sheet("체크리스트(부속서 A)")
    ck.append(["단계", "체크 사항", "이행(Y/N)", "미해당"])
    for st, items in CHECK.items():
        for it in items:
            ck.append([st, it, "", ""])
    sc = wb.create_sheet("일정")
    sc.append(["항목", "기한", "근거"])
    if a.type in ("폐지-승계", "한시", "일부승계"):
        sc.append(["인계계획 통보(중앙기록물관리기관)", "인계절차 착수 전", "5.4.1"])
    if a.done:
        d = dt.date.fromisoformat(a.done)
        m = d.month + 3
        due = dt.date(d.year + (m - 1) // 12, (m - 1) % 12 + 1, min(d.day, 28))
        sc.append(["인계인수서 사본·인수 목록 통보(중앙기록물관리기관)", f"{due.isoformat()}까지(인수 종료 {d.isoformat()} + 3개월)", "9.3"])
    else:
        sc.append(["인계인수서 사본·인수 목록 통보", "인수 종료 뒤 3개월 안 — --done으로 종료일을 주면 계산", "9.3"])
    for t_, y in sorted(later, key=lambda z: z[1])[:200]:
        sc.append([f"영구기관 이관: {t_}", f"{y}년 중", "5.6"])
    sm.append(["건수", f"문서 {ws['문서'].max_row - 1} · 시청각 {ws['시청각'].max_row - 1} · 행정박물 {ws['행정박물'].max_row - 1} · 확인할 것 {issues}"])
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
        for col in "ABCDEFGHIJKLM":
            sh.column_dimensions[col].width = 16
        sh.column_dimensions["B"].width = 40 if sh in (sm, ck, sc) else 16
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.type}: " + sm.cell(sm.max_row, 2).value + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
