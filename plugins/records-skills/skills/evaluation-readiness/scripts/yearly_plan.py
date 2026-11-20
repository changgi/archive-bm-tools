#!/usr/bin/env python3
"""기록관 연간 업무 일정과 기관장 보고용 수준 점검표: NAK 18 부속서 D 월별 업무 일정에 법령·표준의 기한(필수기록물 10/31,
보안대책 연 1회, 정수점검 등)을 더해 월별 표와 캘린더 파일(.ics)을 만들고, 부속서 E 기록관리 수준 체크리스트와
기록관 인력 중 기록물관리 전문요원 비율(최소 1/4)을 점검한다.
근거: NAK 18:2022(v1.1) 「기관장 등 관리자의 기록관리 책임과 역할」 부속서 D·E·F (국가기록원, 2026-11-20 내려받아 확인),
      NAK 2-2:2025 5.4.3(10월 31일), NAK 2-1:2024 5.5(연 1회), 공공기록물법 시행령 제42조①(생산현황 8월 31일)·제63조(평가 1·2·3월)·제78조②·제79조⑤
   — NAK 18 부속서 D는 생산현황 통보를 7월로 적었으나 2026. 8. 25. 개정 시행령이 '매년 8월 31일까지'로 정했으므로 시행령을 따른다.

사용:  python yearly_plan.py --year 2027 [--staff 4 --experts 1] [--ics 산출/기록관리일정_2027.ics] -o 산출/연간업무계획_2027.xlsx
원칙:  일정표와 점검표만 만든다. 캘린더 파일은 담당자가 직접 가져오기(import)한다. 기관 사정에 맞게 날짜를 고쳐 쓴다.
"""
import argparse
import datetime as dt
import os
import sys
import uuid

MONTHLY = {  # NAK 18 부속서 D + 이 킷의 스킬 연결
    1: [("처리과 기록물 정리", "records-register-audit"), ("영구기록물관리기관의 해당 연도 이관 일정·대상 접수", "transfer-schedule"),
        ("해당 연도 기록물관리지침 접수·배포", ""), ("처리과 기록물관리지침 교육", "")],
    2: [("전년도 생산 기록물 정리 감독", "records-register-audit")],
    3: [("처리과 기록물 정리 완료", "records-register-audit"), ("중앙기록물관리기관 기록관리 평가계획 접수", "evaluation-readiness")],
    4: [("전년도 중요 기록물 복제본·보존매체 수록분 목록 제출(생산현황 통보 때 함께 가능)", "production-status-report")],
    5: [("처리과 → 기록관 이관", "transfer-schedule"), ("처리과 → 기록관 생산현황 통보 완료", "production-status-report"),
        ("보존시설·장비·환경 구축 현황 중앙기록물관리기관 통보", "stack-environment")],
    6: [],
    7: [("전년도 기록물 생산현황 영구기록물관리기관 통보 준비(기한 8월 31일, 시행령 제42조① 2026. 8. 25. 개정)", "production-status-report")],
    8: [("중앙기록물관리기관이 이중보존 수록 목록 중 송부 대상·시기 통보", "")],
    9: [],
    10: [("기록관리기준표 변경 단위과제 영구기록물관리기관 협의", "records-schedule-check"),
         ("필수기록물 필수성·보호기간 검토·확정(10월 31일까지)", "vital-records")],
    11: [("다음 연도 이관 일정·대상 영구기록물관리기관에 제출", "transfer-schedule"), ("영구기록물관리기관의 기준표 단위과제 검토 결과 접수", "records-schedule-check")],
    12: [("전자기록생산시스템 정리", ""), ("보존기간 만료 기록물 평가·심의", "retention-review → disposal-dossier")],
}
YEARLY = [("처리과 → 기록관 이관", "transfer-schedule"), ("기록관 → 영구기록물관리기관 이관", "transfer-schedule"), ("폐기(평가 심의 후)", "disposal-dossier"),
          ("공개재분류(5년 주기)", "disclosure-reclass"), ("열람 대응·정보공개 기한", "access-request"), ("정수점검(2년에 1회)·상태검사", "preservation-inspection"),
          ("소속기관 점검(해당 기관)", ""), ("기록관리 평가 준비", "evaluation-readiness"), ("보안대책·재난관리계획 검토(연 1회 이상)", "security-disaster-plan"),
          ("보존매체 점검·간이 검사", "media-check"), ("서고 온습도(매월·분기)", "stack-environment")]
DEADLINES = [((1, 31), "중앙기록물관리기관 기록관리 평가계획 수립·통보(시행령 제63조①) — 지표 확인"),
             ((2, 28), "시·군·구·교육지원청: 전년도 기록관리현황을 시·도(교육청)에 제출(시행령 제63조②)"),
             ((3, 31), "중앙행정기관·시·도(교육청): 전년도 기록관리현황을 중앙기록물관리기관에 제출(시행령 제63조③)"),
             ((8, 31), "기록관·특수기록관: 전년도 기록물 생산현황 관할 영구기록물관리기관 통보 기한(시행령 제42조①)"),
             ((10, 31), "필수기록물 필수성·보호기간 확정 기한(NAK 2-2 5.4.3)")]
LEVEL = {
    "기록관리 정책": ["정책 수립 때 내·외부 환경을 분석했는가?", "정책을 기관장 등 관리자가 채택·승인했는가?", "정책을 조직에 효율적으로 공지하는가?",
                 "정책을 환경·행정 변화에 맞춰 최신으로 유지하는가?", "기관장이 정책 집행을 감독할 관리자를 지정했는가?", "기록물관리 전문요원의 책임이 정의돼 있는가?",
                 "모든 구성원의 기록관리 책임이 정의돼 있는가?", "목표가 측정 가능하고 실현 가능한가?", "목표 달성을 위한 기본계획·세부실천계획이 있는가?"],
    "자원 할당": ["적절한 예산이 지원되는가?", "필요한 인적·물적 자원이 충분한가?", "자격을 갖춘 기록물관리 전문요원이 배치돼 있는가?", "기록관리 인력이 적절히 배치돼 있는가?",
              "보존시설 규모가 적절한가?", "기록관리 장비가 적절한가?"],
    "교육·인식 확산": ["모든 구성원에게 기록관리 교육을 하는가?", "교육 내용이 분야별 업무 특성에 맞는가?", "교육을 정기적·지속적으로 하는가?", "교육 프로그램을 주기적으로 검토·갱신하는가?",
                  "전문요원이 정기적으로 전문교육기관 교육을 이수하는가?", "홍보 프로그램을 개발·운영하는가?"],
    "모니터링·개선": ["정책·절차를 모니터링하는 기능이 있는가?", "모니터링으로 미비사항을 개선했는가?", "관할 공공기관 기록관리도 모니터링·개선했는가?"],
}


def ics_event(d, title, desc):
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    esc = lambda s: s.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\;").replace("\n", "\\n")
    return ["BEGIN:VEVENT", f"UID:{uuid.uuid5(uuid.NAMESPACE_URL, d.isoformat() + title)}@records-skills", f"DTSTAMP:{stamp}",
            f"DTSTART;VALUE=DATE:{d:%Y%m%d}", f"DTEND;VALUE=DATE:{d + dt.timedelta(days=1):%Y%m%d}", f"SUMMARY:{esc(title)}", f"DESCRIPTION:{esc(desc)}", "END:VEVENT"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=dt.date.today().year + 1)
    ap.add_argument("--staff", type=int, help="기록관 인력 수")
    ap.add_argument("--experts", type=int, help="그중 기록물관리 전문요원 수")
    ap.add_argument("--ics", help="캘린더 파일(.ics) 경로 — 매월 1일 할 일과 기한을 종일 일정으로")
    ap.add_argument("-o", "--out", default="")
    a = ap.parse_args()
    out = a.out or f"산출/연간업무계획_{a.year}.xlsx"
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "월별 일정"
    ws.append([f"{a.year}년 기록관 업무 일정(NAK 18 부속서 D + 기한)", "", "", ""])
    ws.append(["월", "할 일", "쓸 스킬", "완료"])
    events = []
    for m in range(1, 13):
        items = MONTHLY[m] or [("(정기 업무 없음 — 연중 업무 진행)", "")]
        for t, sk in items:
            ws.append([f"{m}월", t, sk, ""])
            if not t.startswith("("):
                events.append((dt.date(a.year, m, 1), t, f"{m}월 기록관 업무" + (f" — 스킬: {sk}" if sk else "")))
    for (mm, dd), t in DEADLINES:
        ws.append([f"{mm}월 {dd}일", t, "", ""])
        events.append((dt.date(a.year, mm, dd), "[기한] " + t, t))
    ws.append([])
    ws.append(["연중", "할 일", "쓸 스킬", ""])
    for t, sk in YEARLY:
        ws.append(["연중", t, sk, ""])
    lv = wb.create_sheet("수준 체크리스트")
    lv.append(["영역", "점검 항목(NAK 18 부속서 E — 기관장 보고용)", "이행 여부", "근거자료"])
    for k, items in LEVEL.items():
        for it in items:
            lv.append([k, it, "", ""])
    if a.staff:
        e = a.experts or 0
        need = -(-a.staff // 4)
        lv.append([])
        lv.append(["인력 배치", f"기록물관리기관 정원 {a.staff}명 중 전문요원 {e}명 — 정원의 1/4 이상(1명 미만이면 1명) = 최소 {need}명(시행령 제78조②, NAK 18 부속서 F)",
                   "충족" if e >= need else "미달", ""])
        lv.append(["교육", "전문요원이 아닌 기록물관리기관 종사자는 보직 전 또는 보직 후 6개월 안에 전문교육·훈련 과정 이수(시행령 제79조⑤)", "", ""])
        if e < need:
            lv.cell(lv.max_row, 3).fill = red
    rl = wb.create_sheet("역할(NAK 18)")
    for line in [("기관장", "기록관리 정책 채택·승인, 감독 관리자 지정, 예산·인력·시설 지원, 교육 지원, 평가 결과 반영"),
                 ("기록관장", "기록관 운영 총괄"),
                 ("기록물관리 전문요원", "기본계획, 기준 운영, 평가심의회, 평가·폐기, 공개재분류, 지도·감독, 생산현황 취합·통보, 기록정보 콘텐츠 — 기록관 인력의 1/4 이상"),
                 ("보존", "시설·장비, 인수·이관, 서고·정수점검, 비전자 전자화·보존처리"), ("열람서비스", "간행물·행정자료, 열람·대출, 정보공개 접수"),
                 ("모든 직원", "업무 과정과 결과를 기록으로 생산·등록, 임의 폐기 금지")]:
        rl.append(list(line))
    for sh, widths in ((ws, (10, 80, 34, 8)), (lv, (16, 80, 10, 20)), (rl, (20, 110))):
        sh["A1"].font = Font(bold=True)
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)
    if a.ics:
        lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//records-skills//yearly_plan//KO", "CALSCALE:GREGORIAN", f"X-WR-CALNAME:기록관리 일정 {a.year}"]
        for d, t, desc in events:
            lines += ics_event(d, t, desc)
        lines.append("END:VCALENDAR")
        os.makedirs(os.path.dirname(a.ics) or ".", exist_ok=True)
        with open(a.ics, "w", encoding="utf-8", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")
    print(f"{a.year}년 일정 {len(events)}건" + (f" · 캘린더 {a.ics}" if a.ics else "") + f" → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
