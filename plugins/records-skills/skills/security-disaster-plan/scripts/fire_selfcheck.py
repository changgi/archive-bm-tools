#!/usr/bin/env python3
"""기록관·서고 화재 예방 자체 점검표 — 담당자가 자주(월 1회 권장) 돌아보며 '예/아니오'를 적고, '아니오'는 시정 담당·기한·완료를 추적한다.

근거(항목 구성 참고): 국가기록포털 해외표준 '화재예방 자가 점검 양식'(Fire Safety Self-Inspection Form for Cultural Institutions, NFPA 계열 문화기관용)의
  구성(전 층 공통 → 특수 구역(열람·전시, 서고, 보존처리실, 작업·포장실, 강당·교육장, 식당) → 외부(대피·주변 환경) → 사람(교육·조직))을 따르되,
  문장은 국내 기록관 실정에 맞게 새로 썼다(원문 번역 아님). 법정 기준은 시행령 제60조·별표 6, NAK 11:2025, NAK 2-1:2024(보안·재난관리)가 우선이다.
  자체 점검은 소방 전문기관의 정기 점검과 설비 유지보수를 대신하지 않는다.

사용:
  python fire_selfcheck.py --blank --org 가상시 -o 산출/화재_자체점검표.xlsx            # 빈 점검표(구역 선택: --areas 서고,보존처리실)
  python fire_selfcheck.py --results 채운점검표.xlsx -o 산출/화재_자체점검_결과.xlsx      # '아니오' 집계·시정 추적·지난 점검과 비교(--previous)
"""
import argparse
import datetime as dt
import os
import sys

ITEMS = [
    ("전 층 공통", "비상구·유도등이 켜져 있고 비상조명은 주기적으로 시험한다"),
    ("전 층 공통", "복도·계단·비상구에 물건이 없고 잠겨 있지 않다(보안 장치가 대피를 막지 않음)"),
    ("전 층 공통", "방화문이 자동으로 닫히고 받침 등으로 열어 두지 않았다"),
    ("전 층 공통", "금연이 지켜지고, 청소용품·가연물은 정해진 곳에 정리돼 있다"),
    ("전 층 공통", "임시 배선·멀티탭 연결이 없고, 전열기·커피포트는 금지하거나 자동 차단 제품만 쓴다"),
    ("전 층 공통", "쓰레기를 매일 치우고 건물 밖 닫힌 금속 용기에 둔다"),
    ("전 층 공통", "소화기가 충분하고 종류가 맞으며, 걸이·표지·충압·점검표가 정상이다"),
    ("전 층 공통", "스프링클러 헤드 아래 공간이 비어 있다(적재물이 헤드를 가리지 않음)"),
    ("전 층 공통", "옥내소화전·제어밸브가 표시돼 있고 가려지지 않았으며 밸브 열림을 정기 확인한다"),
    ("전 층 공통", "화재 감지·소화 설비가 작동 상태이고 정기 시험 기록이 있다"),
    ("열람·전시", "목록·소장품을 보호할 응급 복구 용품(비닐·방수포 등)이 바로 쓸 수 있게 있다"),
    ("열람·전시", "전시 구조물·칸막이가 불연·난연 재료이고 소화전·경보기를 가리지 않는다"),
    ("서고", "서고가 방화벽·방화문으로 다른 공간과 구획되고 배관·배선 관통부가 막혀 있다"),
    ("서고", "자동 화재 감지(조기 경보)가 있고 경보가 24시간 감시 지점(종합상황실 등)으로 간다"),
    ("서고", "서고에 맞는 자동 소화 설비(가스계 등)가 있고 소방대가 서고에 접근할 수 있다"),
    ("서고", "서고 안에 사무용품·포장재·종이 박스 같은 불필요한 가연물이 없다"),
    ("서고", "응급 복구 용품(포장 상자·비닐·흡수지 등)이 준비돼 있다(disaster-recovery 참고)"),
    ("보존처리실", "인화성 용제·약품은 표시해 소량만 안전 캐비닛에 두고 안전 용기로 덜어 쓴다"),
    ("보존처리실", "기름·용제 묻은 폐기물은 자동 닫힘 용기에 담아 매일 처리한다"),
    ("보존처리실", "분무·소독 작업 공간의 환기·방폭 전기·자동 소화가 갖춰졌고 출입을 제한한다"),
    ("보존처리실", "작업자가 특수 위험과 대응법을 교육받았고 보호구가 있다"),
    ("작업·포장실", "페인트·시너 등은 환기되는 안전 캐비닛에 적정량만 두고, 가연성 포장재는 닫히는 용기에 둔다"),
    ("작업·포장실", "전동공구는 접지돼 있고 쓰지 않을 때 플러그를 뽑으며, 집진기를 정기적으로 비운다"),
    ("강당·교육장", "수용 인원이 게시되고 지켜지며, 통로에 서거나 앉지 않는다"),
    ("식당·탕비실", "조리기구·후드·배기 덕트를 정기적으로 청소하고(마지막 청소일 기록) 자동 소화 장치가 있다"),
    ("외부", "비상구·피난 경로가 안전한 장소까지 막힘없이 이어진다"),
    ("외부", "건물 주변에 가연물이 쌓여 있지 않고 소방차 진입로·소화전·송수구가 비어 있다"),
    ("사람", "모든 직원이 화재 신고 방법, 대피 때 맡은 역할, 소화기 사용 시점·방법을 안다"),
    ("사람", "화재 예방 책임자가 지정돼 있고 직원 교육 계획이 운영된다"),
    ("사람", "서면 화재 비상 계획이 최신이고 배포돼 있다(security-disaster-plan sd_plan.py)"),
    ("사람", "관할 소방서와 서고 위치·우선 구출 기록물·응급 복구 절차를 공유했고 최근 소방 훈련 날짜를 기록했다"),
]


def main() -> int:
    ap = argparse.ArgumentParser(description="기록관·서고 화재 예방 자체 점검표")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--blank", action="store_true")
    g.add_argument("--results")
    ap.add_argument("--previous", help="지난 점검 결과 xlsx(같은 양식) — 계속 '아니오'인 항목 표시")
    ap.add_argument("--areas", help="쉼표로 구역 선택(예: 전 층 공통,서고,사람). 기본: 전부")
    ap.add_argument("--org", default="")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    areas = [x.strip() for x in a.areas.split(",")] if a.areas else None
    items = [(n, ar, q) for n, (ar, q) in enumerate(ITEMS, 1) if not areas or ar in areas]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검표"
    head = ["번호", "구역", "점검 항목", "예/아니오/해당없음", "메모", "시정 담당", "시정 기한", "시정 완료일"]
    if a.blank:
        ws.append(head)
        dv = DataValidation(type="list", formula1='"예,아니오,해당없음"', allow_blank=True)
        ws.add_data_validation(dv)
        for n, ar, q in items:
            ws.append([n, ar, q, "", "", "", "", ""])
            dv.add(f"D{ws.max_row}")
        ws.append([])
        ws.append(["", "점검일", "", "점검자", "", "검토자", "", ""])
        ws.append(["", "기관", a.org])
    else:
        src = openpyxl.load_workbook(a.results, data_only=True).worksheets[0]
        rows = [r for r in src.iter_rows(min_row=2, values_only=True) if r and isinstance(r[0], int)]
        prev = {}
        if a.previous:
            prev = {r[0]: str(r[3] or "") for r in openpyxl.load_workbook(a.previous, data_only=True).worksheets[0].iter_rows(min_row=2, values_only=True) if r and isinstance(r[0], int)}
        ws.append(head + ["판정"])
        red = PatternFill("solid", fgColor="F8D7DA")
        today = dt.date.today()
        no = open_ = repeat = blank = 0
        for r in rows:
            ans = str(r[3] or "").strip()
            done = r[7]
            due = r[6]
            verdict = ""
            if ans == "아니오":
                no += 1
                if not done:
                    open_ += 1
                    verdict = "시정 필요" + ("" if r[5] else " — 담당 지정") + (" · 기한 지남" if isinstance(due, (dt.date, dt.datetime)) and (due.date() if isinstance(due, dt.datetime) else due) < today else "")
                else:
                    verdict = "시정 완료"
                if prev.get(r[0]) == "아니오":
                    repeat += 1
                    verdict += " · 지난 점검에도 아니오"
            elif not ans:
                blank += 1
                verdict = "미기입"
            ws.append(list(r[:8]) + [verdict])
            if ans == "아니오" and not done:
                for c in ws[ws.max_row]:
                    c.fill = red
        sm = wb.create_sheet("요약")
        sm.append(["항목 수", "아니오", "시정 미완", "연속 아니오", "미기입"])
        sm.append([len(rows), no, open_, repeat, blank])
        sm.append([])
        sm.append(["안내", "'아니오'는 모두 담당 부서·직원에게 넘기고, 관리자는 시정 완료를 확인한다. 자체 점검은 소방 전문 점검·설비 유지보수를 대신하지 않는다."])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.column_dimensions["C"].width = 80
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"점검 항목 {len(items) if a.blank else ws.max_row - 1} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
