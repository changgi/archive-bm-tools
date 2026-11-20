#!/usr/bin/env python3
"""행정정보 데이터세트 관리대상 선정(A·B·C 유형) 초안: 기관 정보시스템 목록으로 유형 후보와 근거를 달고
「데이터세트 관리대상 선정 협의 서식」 칸 순서로 채운다. 유형 확정은 기록관·관할 영구기록물관리기관 협의로 한다.

사용:  python dataset_screen.py 시스템목록.xlsx --org 기관명 -o 산출/데이터세트_관리대상_선정.xlsx
목록 열(머리글 이름으로 찾음): 시스템명, 시스템개요, 데이터세트내용(생성되는 데이터세트 내용), 부서명, 담당자, 연락처,
       DB저장(예/아니오), 연계표출(예/아니오: 다른 시스템 데이터를 보여 주기만 함), 자체생성(예/아니오: 연계 데이터로 통계 등 별도 생성),
       다른법률(근거 법률명 — 공공기록물법 아닌 다른 법률에 기록관리 특별 규정), 선정유형(이미 정한 값, 있으면 그대로 두고 점검만)
근거:  「2026년 기록물관리 지침」 제2장 5. 행정정보 데이터세트·제3장 제2절(행정정보 데이터세트의 기록관리 기준):
         A 관리대상 — 업무 관련 데이터세트를 생산·보유·활용하는 정보시스템 → 관리기준표 작성
         B 관리제외(데이터세트 기록관리 미대상) — ① DB에 업무내역이 저장되지 않는 단순 중계(공지메일·문자·웹하드)
            ② ICT 운영 기술 시스템(자료전송·화상회의 중계·침입탐지·IP관리) ③ 연계 데이터 표출만(별도 생성 데이터세트가 있으면 A)
            ④ 공공기록물법이 아닌 다른 법률에 특별 규정 — 제외사유 검토 대상
         C 관리제외(별도 기록관리 방안 적용) — 업무관리·기록관리시스템(RMS/AMS)·기관 대표 홈페이지 등.
            다만 업무 데이터세트가 기록관리체계에서 누락되면 A로 포함 가능(예: 홈페이지 ○○게시판)
         B·C는 '선정 제외 사유'를 반드시 구체적으로 적는다. 협의 서식 칸: 연번·기관명(부서명)·담당(이름·연락처)·
         시스템 명세(시스템명·시스템 개요·생성되는 데이터세트 내용)·선정유형(A,B,C)·선정 제외 사유·비고.
       시행령 제34조의3, NAK 35:2020(v1.0). 국가기록원 관할기관은 '행정정보 데이터세트 종합관리시스템'(www.darm.go.kr, 업무망)으로
       대상 선정·관리기준표·협의 — 중앙행정기관은 반드시 활용(지침 참고사항).
원칙:  후보와 근거만 단다. 유형 확정·변경(변경 시 사유 이력관리)은 기록관과 관할 영구기록물관리기관이 한다.
"""
import argparse
import os
import re
import sys

ALIAS = {"시스템명": ["시스템명", "정보시스템명", "시스템"], "개요": ["시스템개요", "시스템 개요", "개요", "주요기능"],
         "데이터": ["데이터세트내용", "생성되는 데이터세트 내용", "데이터세트 내용", "보유데이터"], "부서": ["부서명", "담당부서", "기관명(부서명)"],
         "담당자": ["담당자", "이름", "담당자명"], "연락처": ["연락처", "전화", "전화번호"], "db": ["DB저장", "DB 저장", "업무내역저장"],
         "표출": ["연계표출", "표출만", "연계표출만"], "생성": ["자체생성", "별도생성", "통계생성"], "법률": ["다른법률", "다른 법률", "특별규정법률"],
         "유형": ["선정유형", "선정 유형", "유형"], "사유": ["선정제외사유", "선정 제외 사유", "제외사유"]}
RELAY = r"메일발송|문자전송|문자발송|웹하드|알림톡|팩스|푸시|중계|메신저|쪽지|웹\s*메일|메일\s*시스템"
ICT = r"네트워크|NMS|TMS|트래픽|서버\s*관리|웹\s*필터링|자료전송|화상회의|차단|유해사이트|침입탐지|침입방지|IPS|IDS|IP관리|백신|방화벽|접근제어|NAC|DLP|로그관리|모니터링|패치|망연계|SSO|통합인증|보안"
RECORDS = r"업무관리시스템|온나라|전자문서\s*시스템|기록관리시스템|RMS|AMS|CAMS|전자결재|사내\s*포털|EIP"
WEB = r"홈페이지|누리집|모바일\s*앱|모바일앱|웹\s*사이트"
BOARD = r"게시판|민원|신청|접수|공모|예약"
YES = lambda v: str(v or "").strip().upper() in ("Y", "예", "O", "TRUE", "1", "있음")
NO = lambda v: str(v or "").strip().upper() in ("N", "아니오", "X", "FALSE", "0", "없음")


def suggest(x):
    """(유형, 세부, 근거, 확인할 것)"""
    blob = " ".join(x.get(k, "") for k in ("시스템명", "개요", "데이터"))
    head = " ".join(x.get(k, "") for k in ("시스템명", "개요"))  # 기술 시스템·중계 여부는 이름·개요로 본다(데이터 칸의 '사업모니터링' 등 오판 방지)
    if x.get("법률"):
        return "B", "④", f"다른 법률({x['법률']})에 기록관리 특별 규정", "그 법률 조항이 기록물 보존·폐기를 정하는지 확인(제외사유 검토 대상)"
    if re.search(RECORDS, head, re.I):
        return "C", "", "공공기록물법령의 별도 관리체계(전자기록생산·기록관리시스템)가 적용되는 시스템", "기관 업무 데이터세트가 기록관리체계에서 빠지는 것이 없는지 확인"
    if YES(x.get("표출")):
        if YES(x.get("생성")):
            return "A", "", "연계 데이터를 표출하지만 통계 등 데이터세트를 별도 생성·관리", ""
        return "B", "③", "다른 시스템에서 연계받은 데이터를 표출만 함", "원천 시스템이 A로 관리되는지 확인"
    if re.search(ICT, head, re.I):
        return "B", "②", "ICT 운영 등을 위한 기술적 시스템", "업무 데이터(예: 차단목록 선정 결과)가 쌓이는지 확인"
    if re.search(RELAY, head, re.I) or NO(x.get("db")):
        return "B", "①", "업무내역이 DB로 저장되지 않는 단순 중계 시스템", "발송 이력 등이 업무 기록으로 쓰이는지 확인"
    if re.search(WEB, head, re.I):
        if re.search(BOARD, x.get("시스템명", "")):
            return "A", "", "기관 대표 홈페이지의 게시판 등 업무 데이터세트가 기록관리체계에서 빠지는 경우 A로 포함(2026 지침 C유형 단서)", "그 게시판 데이터가 웹기록물 수집이나 RMS로 관리되는지 확인"
        return "C", "", "기관 대표 홈페이지·모바일 앱은 별도 관리체계(웹기록물)가 적용되는 시스템", "게시판·신청 등 업무 데이터세트가 기록관리체계에서 빠지면 그 부분만 A로 포함 검토"
    if x.get("데이터") or YES(x.get("db")):
        return "A", "", "업무와 관련하여 데이터세트를 생산·보유·활용", "관리기준표 작성(dataset_schedule.py)"
    return "", "", "판단 정보 부족", "시스템 개요·데이터세트 내용·DB 저장 여부를 채울 것"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--org", default="")
    ap.add_argument("-o", "--out", default="산출/데이터세트_관리대상_선정.xlsx")
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    items = [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "선정 협의 서식"
    ws.append(["데이터세트 관리대상 선정 협의 서식(초안) — 「2026년 기록물관리 지침」 서식 칸 순서. 선정유형은 후보이며 협의로 확정"])
    ws.append(["연번", "기관명(부서명)", "담당 이름", "담당 연락처", "시스템명", "시스템 개요", "생성되는 데이터세트 내용", "선정유형(A, B, C)",
               "선정 제외 사유", "비고"])
    wn = wb.create_sheet("판단 근거")
    wn.append(["연번", "시스템명", "후보 유형", "세부", "근거", "확인할 것", "기존 값과 차이"])
    cnt = {"A": 0, "B": 0, "C": 0, "": 0}
    yellow = PatternFill("solid", fgColor="FFF2CC")
    for i, x in enumerate(items, 1):
        typ, sub, why, ask = suggest(x)
        given = x.get("유형", "").strip().upper()[:1]
        final = given if given in ("A", "B", "C") else typ
        reason = x.get("사유", "") or ("" if final == "A" or not final else f"{'B' + sub if final == 'B' and sub else final}: {why}")
        dept = " / ".join(v for v in (a.org, x.get("부서", "")) if v)
        ws.append([i, dept, x.get("담당자", ""), x.get("연락처", ""), x.get("시스템명", ""), x.get("개요", ""), x.get("데이터", ""), final, reason or "",
                   "후보(협의 전)" if final and not given else ("" if final else "판단 정보 부족")])
        r = ws.max_row
        if final in ("B", "C") and not x.get("사유"):
            ws.cell(r, 9).fill = yellow
        if not final:
            ws.cell(r, 8).fill = yellow
        diff = f"기존 '{given}' / 후보 '{typ}'" if given in ("A", "B", "C") and typ and given != typ else ""
        wn.append([i, x.get("시스템명", ""), typ, sub, why, ask, diff])
        cnt[final] = cnt.get(final, 0) + 1
    for c in ws[2]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col, wdt in zip("ABCDEFGHIJ", (5, 18, 8, 14, 18, 36, 36, 10, 40, 12)):
        ws.column_dimensions[col].width = wdt
    gd = wb.create_sheet("유형 기준")
    for line in (["유형", "세부", "형태", "비고"],
                 ["A 관리대상", "", "업무와 관련하여 생산·보유·활용할 목적으로 구축한 데이터세트를 가진 정보시스템", "관리기준표 작성"],
                 ["B 관리제외(데이터세트 기록관리 미대상)", "①", "업무내역이 DB로 저장되지 않는 단순 중계(공지메일발송·문자전송·웹하드 등)", "제외 사유"],
                 ["", "②", "ICT 운영 등을 위한 기술적 시스템(자료전송·화상회의 중계·침입탐지·IP관리 등)", "제외 사유"],
                 ["", "③", "타 시스템 연계 데이터를 표출만(연계 데이터로 통계 등 별도 생성하면 A)", "제외 사유"],
                 ["", "④", "공공기록물법이 아닌 '다른 법률'에 특별 규정", "제외 사유 검토 대상"],
                 ["C 관리제외(별도 기록관리 방안 적용)", "", "업무관리시스템·기록관리시스템(RMS/AMS)·기관 대표 홈페이지 등. 업무 데이터세트가 기록관리체계에서 누락되면 A 포함 가능(홈페이지 ○○게시판 등)", "제외 사유"],
                 ["근거", "", "「2026년 기록물관리 지침」 행정정보 데이터세트 관리대상 유형 구분, 시행령 제34조의3, NAK 35:2020", ""],
                 ["협의", "", "국가기록원 관할기관은 행정정보 데이터세트 종합관리시스템(www.darm.go.kr, 업무망) — 중앙행정기관 필수, 최초 협의요청은 공문 병행", ""]):
        gd.append(line)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"시스템 {len(items)}개 · A {cnt['A']} · B {cnt['B']} · C {cnt['C']} · 판단 정보 부족 {cnt['']} → {a.out}")
    print("B·C의 노란 칸(선정 제외 사유)은 구체적으로 고쳐 쓴 뒤 기록관 검토 → 관할 영구기록물관리기관 협의. A는 dataset_schedule.py로 관리기준표 초안.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
