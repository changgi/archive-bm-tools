#!/usr/bin/env python3
"""업무과정 분석·평가 작업지 — 기록관리기준표를 새로 만들거나 정비할 때, 단위과제마다 '어떤 업무 행위가 어떤 기록을 남겨야 하는지'를 적는다.

근거(국가표준, 원문 유료 — 목차·범위는 ISO 공개 초록과 DIN 공개 목차로 확인):
  KS X ISO TR 26122(ISO/TR 26122:2008) 기록을 위한 업무과정 분석 — 4 범위·참여자, 5 맥락 검토, 6 기능 분석(기능→과정 분해),
     7 순차 분석(7.2 트랜잭션 순서, 7.4 변형, 7.6 규칙, 7.8 다른 과정과의 연결), 8 참여자 검증
  KS X ISO TR 21946(ISO/TR 21946:2018) 기록관리를 위한 평가 — 5.2 범위, 5.3 참여자, 5.5 업무 맥락, 5.6 기술 맥락, 5.7~5.8 기능·순차 분석,
     5.9 행위자, 5.10 업무상 중요 영역, 5.11 기록 요구(업무·법규·사회적 기대), 6.2 요구↔기능 연결, 6.3 위험 평가·처리, 6.4 문서화, 7 모니터링, 8 검토
  공공기록물법 시행령 제25조(기록관리기준표: 단위과제·보존기간·보존방법·보존장소·비치 여부·공개 여부·접근권한), 별표 1(보존기간 책정 기준)

입력: 단위과제 목록(xlsx·csv) — 열(머리글 이름으로 찾음): 기능(대·중·소기능), 단위과제(명), 보존기간, 처리과(부서), 단위과제 설명(있으면)
결과: xlsx
  - 맥락 검토: 범위·참여자·법적 근거·업무 시스템·기술 환경(26122 5 / 21946 5.2~5.6)
  - 기능 분석: 단위과제마다 상위 기능·과정·책임 부서(26122 6)
  - 순차 분석: 단위과제마다 행위(트랜잭션) 1~N 행 — 시작 계기·행위·행위자·남는 기록·시스템·다음 과정(26122 7.2~7.8, 21946 5.9)
  - 기록 요구: 업무 필요·법규·사회적 기대·업무상 중요 여부(21946 5.10~5.11)
  - 위험 평가: 기록이 없거나 늦게 지워질 때의 가능성·영향 1~5 → 점수 수식(21946 6.3), 보존기간이 별표 1 하한보다 짧아 보이는 단서
  - 기록화 설정: record-by-design `sweep`에 바로 쓸 기록화설정.json 초안(단위과제별 하위 폴더)
사람이 채우는 칸은 비워 두고, 목록에서 알 수 있는 것만 채운다. 판단(보존기간 변경·기준표 확정)은 하지 않는다.

예) python appraisal_worksheet.py 단위과제.xlsx --agency 가상시 --steps 4 -o 산출/업무분석.xlsx
"""
import argparse
import csv
import json
import os
import re
import sys

ALIAS = {"기능": ["기능", "대기능", "중기능", "소기능", "기능명"], "과제": ["단위과제", "단위과제명", "과제명", "분류명"], "보존": ["보존기간"],
         "부서": ["처리과", "부서", "부서명", "처리과명"], "설명": ["단위과제 설명", "설명", "단위과제설명", "업무내용"]}
PERIODS = ("영구", "준영구", "30년", "10년", "5년", "3년", "1년")
# 별표 1 단서(보존기간 하한을 시사하는 낱말) — 확정 판정이 아니라 확인 요청
HINTS = [(r"법령|조례|규칙|훈령|예규|고시", "영구", "법령·자치법규 제정·개정(별표 1·준칙에서 영구 여부 확인)"),
         (r"예산|결산|회계", "5년", "예산·결산 — 기관 보존기간 준칙 확인"),
         (r"인사|임용|징계|포상", "준영구", "인사 기록 — 준칙 확인"),
         (r"소송|재판|심판|행정심판", "10년", "쟁송 — 종결 후 기간 확인"),
         (r"계약", "5년", "계약 — 하자담보·감사 기간 확인"),
         (r"토지|지적|등기|건축|인허가", "영구", "재산권·인허가 — 준칙 확인"),
         (r"민원", "3년", "민원 — 처리 결과 기간 확인")]
RANK = {p: i for i, p in enumerate(PERIODS)}


def load(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        rows = [["" if v is None else str(v).strip() for v in r] for r in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)]
    else:
        rows = [[c.strip() for c in r] for r in csv.reader(open(path, encoding="utf-8-sig"))]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:15]) if sum(c.replace(" ", "") in {x.replace(" ", "") for x in names} for c in r) >= 2), 0)
    head = [c.replace(" ", "") for c in rows[hi]]
    idx = {}
    for k, al in ALIAS.items():
        hits = [head.index(n.replace(" ", "")) for n in al if n.replace(" ", "") in head]
        if hits:
            idx[k] = hits if k == "기능" else hits[0]
    out = []
    for r in rows[hi + 1:]:
        if not any(r):
            continue
        g = lambda k: (r[idx[k]] if k in idx and idx[k] < len(r) else "")
        fn = " > ".join(r[i] for i in idx.get("기능", []) if i < len(r) and r[i])
        out.append({"기능": fn, "과제": g("과제"), "보존": g("보존"), "부서": g("부서"), "설명": g("설명")})
    return [x for x in out if x["과제"]]


def main() -> int:
    ap = argparse.ArgumentParser(description="업무과정 분석·평가 작업지(KS X ISO TR 26122·TR 21946)")
    ap.add_argument("listing")
    ap.add_argument("--agency", default="")
    ap.add_argument("--steps", type=int, default=4, help="순차 분석에서 단위과제마다 비워 둘 행위 행 수")
    ap.add_argument("-o", "--out", default="산출/업무분석.xlsx")
    a = ap.parse_args()
    items = load(a.listing)
    if not items:
        print("입력 오류: '단위과제' 열을 찾지 못했습니다.", file=sys.stderr)
        return 2
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    fill = PatternFill("solid", fgColor="FFF4CC")
    bold = Font(bold=True)

    def sheet(title, head, widths, first=False):
        ws = wb.active if first else wb.create_sheet(title)
        ws.title = title
        ws.append(head)
        for c in ws[1]:
            c.font = bold
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        return ws

    cx = sheet("맥락 검토", ["항목", "질문", "답(채움)", "근거"], [16, 60, 50, 26], first=True)
    for row in [("범위", "이번 분석이 다루는 기능·부서·기간은? 빠지는 것은?", "TR 26122 4.3 · TR 21946 5.2"),
                ("참여자", "업무 담당자·기록관리 담당·정보시스템 담당·법무 등 누가 참여하고 누가 검증하는가?", "TR 26122 4.4 · TR 21946 5.3"),
                ("법적 근거", "이 기능을 정한 법령·자치법규·지침과 그 조항은?", "TR 26122 5 · TR 21946 5.11.3"),
                ("업무 맥락", "조직의 임무·조직도·최근 개편(통폐합·이관)은?", "TR 21946 5.5"),
                ("기술 맥락", "기록이 생기는 시스템(전자문서·업무관리·행정정보·웹·SNS)과 보존 포맷은?", "TR 21946 5.6"),
                ("외부 기대", "감사·언론·정보공개 청구에서 이 기능의 기록을 자주 찾는가?", "TR 21946 5.11.4")]:
        cx.append([row[0], row[1], "", row[2]])
    for r in cx.iter_rows(min_row=2, min_col=3, max_col=3):
        r[0].fill = fill
    cx.append([])
    cx.append(["기관", a.agency or "(채움)", "", "시행령 제25조"])

    fa = sheet("기능 분석", ["연번", "기능 경로", "단위과제", "처리과", "현 보존기간", "단위과제 설명", "이 과제를 이루는 과정(채움)", "책임 부서(채움)", "다른 기관과 겹치는 기능(채움)"],
               [6, 34, 30, 14, 10, 40, 40, 16, 24])
    sq = sheet("순차 분석", ["단위과제", "순서", "시작 계기(채움)", "행위(채움)", "행위자·직위(채움)", "남는 기록(채움)", "생산 시스템(채움)", "변형·예외(채움)",
                         "규칙·근거(채움)", "다음 과정·연결(채움)"], [30, 6, 18, 26, 16, 26, 16, 20, 20, 20])
    rq = sheet("기록 요구", ["단위과제", "업무상 필요(채움)", "법규상 요구(채움)", "사회적 기대(채움)", "업무상 중요 영역", "현 보존기간", "확인 단서(별표 1·준칙)"],
               [30, 26, 26, 26, 12, 10, 44])
    rk = sheet("위험 평가", ["단위과제", "위험(채움)", "가능성 1~5", "영향 1~5", "점수", "등급", "처리 방안(채움)"], [30, 40, 10, 10, 8, 8, 40])
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True)
    yn = DataValidation(type="list", formula1='"예,아니오"', allow_blank=True)
    rk.add_data_validation(dv)
    rq.add_data_validation(yn)
    cfg = {"기본": {"agency": a.agency or "(기관명)", "dept": "", "task": "", "period": ""}, "폴더": {}, "제외": ["*.tmp"]}
    short = 0
    for n, x in enumerate(items, 1):
        fa.append([n, x["기능"], x["과제"], x["부서"], x["보존"], x["설명"], "", "", ""])
        for s in range(1, a.steps + 1):
            sq.append([x["과제"], s] + [""] * 8)
        cue = []
        p = next((q for q in PERIODS if q in x["보존"]), "")
        for pat, low, why in HINTS:
            if re.search(pat, x["과제"] + x["설명"]):
                if p and RANK[p] > RANK[low]:
                    cue.append(f"{why}: 현재 {p} — {low} 이상인지 확인")
                    short += 1
                else:
                    cue.append(why)
        if not p:
            cue.append("보존기간이 7종(영구~1년)이 아님")
        rq.append([x["과제"], "", "", "", "", x["보존"], " · ".join(cue)])
        yn.add(f"E{rq.max_row}")
        rk.append([x["과제"], "", None, None, f'=IF(AND(ISNUMBER(C{rk.max_row + 1}),ISNUMBER(D{rk.max_row + 1})),C{rk.max_row + 1}*D{rk.max_row + 1},"")',
                   f'=IF(E{rk.max_row + 1}="","",IF(E{rk.max_row + 1}>=15,"높음",IF(E{rk.max_row + 1}>=8,"보통","낮음")))', ""])
        dv.add(f"C{rk.max_row}:D{rk.max_row}")
        key = re.sub(r'[\\/:*?"<>|]', "_", x["과제"])[:40]
        cfg["폴더"][key] = {"task": x["과제"], **({"dept": x["부서"]} if x["부서"] else {}), **({"period": p} if p else {}),
                           **({"activity": x["설명"][:80]} if x["설명"] else {})}
    for ws in (fa, sq, rq, rk):
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if "(채움)" in str(ws.cell(1, c.column).value) or ws.cell(1, c.column).value in ("가능성 1~5", "영향 1~5", "업무상 중요 영역"):
                    c.fill = fill
                c.alignment = Alignment(wrap_text=True, vertical="top")
    gd = sheet("기준", ["구분", "내용"], [22, 100])
    for row in [("이 작업지의 순서", "맥락 검토 → 기능 분석(기능을 과정으로 나눔) → 순차 분석(과정을 행위 순서로, 변형·규칙·연결까지) → 참여자 검증 → 기록 요구 → 위험 평가 → 문서화·모니터링"),
                ("순차 분석의 쓸모", "'남는 기록' 칸이 비면 그 행위의 증거가 생기지 않는다는 뜻. 생산 시스템에 등록 기능이 없으면 record-by-design으로 생산 시점에 패키지화"),
                ("위험 점수", "가능성×영향(1~25). 15 이상 높음·8 이상 보통 — 기관이 기준을 바꿔도 됨"),
                ("하지 않는 것", "보존기간·기준표 확정(기관 기록관리 절차와 국가기록원 협의로 정함), 법령 해석"),
                ("기록화 설정", "'기록화 설정' 시트의 JSON을 업무 폴더에 기록화설정.json으로 저장하고 단위과제 이름의 하위 폴더에 파일을 두면 capture.py sweep이 그 과제로 패키지화"),
                ("근거", "KS X ISO TR 26122 / ISO/TR 26122:2008, KS X ISO TR 21946 / ISO/TR 21946:2018(목차·범위 확인), 공공기록물법 시행령 제25조·별표 1")]:
        gd.append(row)
    cj = sheet("기록화 설정", ["record-by-design sweep용 기록화설정.json 초안(dept·period 빈 칸은 채움)"], [120])
    for line in json.dumps(cfg, ensure_ascii=False, indent=1).splitlines():
        cj.append([line])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    with open(os.path.splitext(a.out)[0] + "_기록화설정.json", "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=1)
    print(f"단위과제 {len(items)} · 순차 분석 행 {len(items) * a.steps} · 보존기간 확인 단서 {short} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
