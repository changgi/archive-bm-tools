#!/usr/bin/env python3
"""심의 뒤 처분결과 관리와 폐기 집행 점검표: 기록물평가심의회 결과(보존기간 재책정·보류·폐기)가 적힌 평가심의서로
처분결과 등록 목록(다음 평가 연도·이관 대상), 폐기 집행 대조 목록, 매체별 폐기 방법, 위탁 폐기 점검표를 만든다.
근거: NAK 5-1:2014(v2.2) 「기록물 평가·폐기 절차 - 제1부: 기록관용」 7.5~8.4, 부속서 A·D (국가기록원, 2026-11-20 내려받아 확인)

사용:  python disposition_result.py 평가심의서.xlsx [--year 2026] [--cert-count 120] -o 산출/처분결과.xlsx
입력 열(머리글 이름으로 찾음): 관리번호, 제목, 생산부서, 보존기간, 종료연도, 심의 결과(폐기·보존기간 재책정·보류),
       재책정 보존기간, 보류기간(년 또는 만료일), 사유, 기록물형태(전자·비전자), 반출확인
규칙:
  7.5.1 재책정 — 보존기간 변경 이력 관리, 30년·준영구·영구로 재책정되면 이관 연도에 영구기록물관리기관으로 이관
  7.5.2 보류 — 보류기간이 지나거나 사유가 없어지면 그 해 평가 대상에 다시 넣음
  7.6   처분결과 등록: 결과·사유, 재책정 보존기간 또는 보류기간, 폐기 집행 결과(일자·담당자)
  8.1~8.2 폐기 목록 → 서고 반출 → 목록과 대조 확인
  8.3   전자: 물리적 파기(파쇄 조각 0.25mm 이하)·소자·겹쳐쓰기(재사용 시 3회 이상) / 비전자: 기관 안에서 파쇄(민감 기록은 교차 세절) 후 반출, 안 되면 용해·담당자 감독
  8.4   위탁: 보안교육·보안각서, 담당 공무원 동승, 밀폐·잠금 차량, 끝까지 입회, 폐기증명서(수량·방법)
원칙:  목록과 점검표만 만든다. 폐기 집행·시스템 삭제는 하지 않는다. 폐기 금지 기록물(시행령 제54조의2~4)이 섞였는지는 build_dossier --hold로 먼저 거른다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"번호": ["관리번호", "철관리번호", "기록물건등록번호"], "제목": ["제목", "기록물철제목", "기록물건제목"], "부서": ["생산부서", "처리과"],
         "보존": ["보존기간"], "종료": ["종료연도", "종료년도", "생산연도", "생산년도"], "결과": ["심의 결과", "심의결과", "평가결과", "처분"],
         "재책정": ["재책정 보존기간", "재책정보존기간", "변경 보존기간"], "보류": ["보류기간", "보류 기간", "보류만료"],
         "사유": ["사유", "심의 사유", "심의사유", "전문요원 사유"], "형태": ["기록물형태", "형태", "매체"], "반출": ["반출확인", "반출 확인"]}
LONG = {"30년", "준영구", "영구"}


def read(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["심의 대상"] if "심의 대상" in wb.sheetnames else wb.worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def yrs(v):
    s = re.sub(r"\s", "", v or "")
    if s in ("영구", "준영구"):
        return s
    m = re.search(r"(\d+)\s*년", s)
    return int(m[1]) if m else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("--cert-count", type=int, help="대행업체 폐기증명서에 적힌 수량 — 폐기 목록 수와 대조")
    ap.add_argument("-o", "--out", default="산출/처분결과.xlsx")
    a = ap.parse_args()
    rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    rs = wb.active
    rs.title = "처분결과 등록(7.6)"
    rs.append(["관리번호", "제목", "생산부서", "원래 보존기간", "심의 결과", "사유", "재책정 보존기간", "보류기간", "다음 할 일", "다음 평가·이관 연도", "확인할 것"])
    cnt = {"폐기": 0, "보존기간 재책정": 0, "보류": 0, "미정": 0}
    dispose, issues = [], 0
    for x in rows:
        res = re.sub(r"\s", "", x.get("결과", ""))
        kind = "폐기" if "폐기" in res else "보존기간 재책정" if "재책정" in res else "보류" if "보류" in res else "미정"
        cnt[kind] += 1
        end = re.search(r"(19|20)\d{2}", x.get("종료", ""))
        end = int(end[0]) if end else None
        todo, nxt, notes = "", "", []
        if not x.get("사유") and kind != "미정":
            notes.append("사유 없음 — 결과와 사유를 함께 등록(7.6)")
        if kind == "보존기간 재책정":
            ny = yrs(x.get("재책정"))
            if ny is None:
                notes.append("재책정 보존기간 없음")
            elif x.get("재책정", "").replace(" ", "") in LONG:
                todo = "보존기간 변경 이력 등록, 대상 기록물에 재책정 보존기간 표기 · 이관 연도에 영구기록물관리기관으로 이관(7.5.1)"
                nxt = f"{end + 6}년 이관(종료+6, 제40조①)" if end else "종료연도 필요"
                if end and end + 6 <= a.year:
                    nxt += " — 이미 도래"
            else:
                todo = "보존기간 변경 이력 등록, 대상 기록물에 재책정 보존기간 표기(7.5.1)"
                nxt = f"{end + ny + 1}년 평가 대상(종료+{ny}년 만료 다음 해)" if end and isinstance(ny, int) else ""
        elif kind == "보류":
            b = x.get("보류", "")
            m = re.search(r"(20\d{2})[-.]\d{1,2}[-.]\d{1,2}", b)
            n = re.search(r"(\d+)\s*년", b)
            if m:
                nxt = f"{int(m[1])}년(보류 만료 해)부터 다시 평가"
            elif n:
                nxt = f"{a.year + int(n[1])}년 다시 평가"
            else:
                notes.append("보류기간 없음 — 기간이나 보류 사유(소송·민원 등) 소멸 시점을 등록")
            todo = "보류기간 경과 또는 사유 소멸 뒤 그 해 평가 대상에 다시 포함(7.5.2)"
        elif kind == "폐기":
            todo = "폐기 목록 → 반출 → 목록 대조 → 매체별 폐기 → 집행 결과(일자·담당자) 등록(8장)"
            dispose.append(x)
        else:
            notes.append("심의 결과를 읽지 못함(폐기·보존기간 재책정·보류)")
        issues += bool(notes)
        rs.append([x.get("번호", ""), x.get("제목", ""), x.get("부서", ""), x.get("보존", ""), kind if kind != "미정" else x.get("결과", ""),
                   x.get("사유", ""), x.get("재책정", ""), x.get("보류", ""), todo, nxt, " · ".join(notes)])
        if notes:
            rs.cell(rs.max_row, 11).fill = red
    dl = wb.create_sheet("폐기 집행 대조(8.1~8.2)")
    dl.append(["번호", "관리번호", "제목", "생산부서", "형태", "서고 반출 확인", "목록 대조", "폐기 방법", "폐기 일자", "담당자", "확인할 것"])
    elec = paper = 0
    for i, x in enumerate(dispose, 1):
        f = x.get("형태", "")
        is_e = bool(re.search(r"전자", f)) and "비전자" not in f
        elec += is_e
        paper += not is_e
        method = ("물리적 파기(조각 0.25mm 이하)·소자·겹쳐쓰기 중 택 — 기록관리자·시스템·정보보안 담당자 책임 하에" if is_e
                  else "기관 안에서 파쇄(민감 기록은 교차 세절) 후 반출, 불가하면 원형 반출·용해 + 담당자 감독")
        dl.append([i, x.get("번호", ""), x.get("제목", ""), x.get("부서", ""), f or "(형태 없음)", x.get("반출", ""), "", method, "", "",
                   "" if f else "기록물형태(전자·비전자)를 적어야 폐기 방법을 고름"])
    dl.append([])
    dl.append(["", "폐기 집행 전 기관 누리집 등에 심의 결과·폐기 대상 목록을 일정 기간 공개해 의견을 듣는 것이 바람직(7.5.3 비고, 선택)"])
    mt = wb.create_sheet("폐기 방법(부속서 A)")
    mt.append(["구분", "방법", "기준"])
    for line in [("전자", "물리적 파기", "파쇄·천공·소각·용해 — 파쇄 조각 0.25mm 이하"), ("전자", "소자(디가우징)", "공인 전용 디가우저, 매체 자기력보다 강해야 함"),
                 ("전자", "겹쳐쓰기", "난수·0·1로 1·3·7·35회 — 재사용 시 3회 이상 영구삭제 프로그램"),
                 ("전자", "재사용하지 않는 디스크", "디스크마다 물리적 파기 + 소자 함께 권고"),
                 ("전자", "영구삭제 프로그램을 못 쓰는 디스크 어레이", "A→B로 전량 이동 → A 겹쳐쓰기·로우 레벨 포맷 → 폐기 대상 뺀 나머지를 A로 → B 완전 삭제"),
                 ("비전자", "파쇄", "내용 판독 불가하게 세절, 보안·개인정보는 교차 세절"), ("비전자", "용해", "제대로 처리하면 매우 안전"),
                 ("공통", "메타데이터", "폐기 기록물 메타데이터(NAK 8 부속서 F) 남김")]:
        mt.append(list(line))
    vd = wb.create_sheet("위탁 폐기 점검(8.4)")
    vd.append(["점검 항목", "확인", "비고"])
    for line in ["대행업체 담당자 보안교육을 했는가?", "보안각서를 받았는가?", "기관 밖 반출 시 담당 공무원이 이송 차량에 함께 탔는가?",
                 "적재 공간이 밀폐되고 잠금장치가 있는 차량인가?", "완전 삭제·물리적 소멸이 끝날 때까지 직원이 입회해 감독했는가?",
                 "폐기증명서(폐기 수량·방법)를 받았는가?"]:
        vd.append([line, "", ""])
    if a.cert_count is not None:
        ok = a.cert_count == len(dispose)
        vd.append([f"폐기증명서 수량 {a.cert_count} / 폐기 목록 {len(dispose)}", "일치" if ok else "불일치", "" if ok else "차이를 확인해 집행 결과 등록 전 바로잡음"])
        if not ok:
            vd.cell(vd.max_row, 2).fill = red
    ck = wb.create_sheet("체크리스트(부속서 D)")
    ck.append(["영역", "점검 항목", "이행 여부", "근거자료"])
    for line in [("평가정책", "기관 업무·조직 특성을 반영한 평가정책을 수립했는가?"), ("평가정책", "담당자별 역할·책임을 업무분장에 반영했는가?"),
                 ("심의회", "민간 전문가 2명(특수기록관 1명) 이상(기록학·역사학·행정학·문헌정보학 등)을 넣었는가?"),
                 ("평가절차", "전문요원 심사 뒤 기록물평가심의서를 작성했는가?"),
                 ("평가절차", "책정 보존기간과 기준표 보존기간이 다르면 더 긴 쪽을 고려했는가?"), ("평가절차", "심의·처분 결과를 관리하는가?"),
                 ("폐기", "전자기록물을 기록관리자 등의 책임 하에 복구 불가능하게 폐기했는가?"), ("폐기", "비전자기록물은 파쇄·용해로 유출을 막았는가?"),
                 ("폐기", "대행업체 폐기 때 보안각서를 받고 전 과정을 감독했는가?"), ("폐기", "대행업체에서 폐기 증명서류를 받았는가?")]:
        ck.append([*line, "", ""])
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
    for sh, widths in ((rs, (16, 30, 12, 10, 14, 24, 12, 12, 60, 24, 36)), (dl, (6, 16, 30, 12, 10, 12, 10, 60, 12, 10, 30)),
                       (mt, (10, 34, 80)), (vd, (60, 10, 40)), (ck, (12, 80, 10, 20))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"처분결과 {len(rows)}건 · 폐기 {cnt['폐기']}(전자 {elec}·비전자 {paper}) · 재책정 {cnt['보존기간 재책정']} · 보류 {cnt['보류']}"
          + (f" · 결과 미정 {cnt['미정']}" if cnt["미정"] else "") + f" · 확인할 것 {issues} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
