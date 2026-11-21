#!/usr/bin/env python3
"""간행물 점검: 기관 간행물 목록의 발간등록번호·발간주기·송부(책자 3부+PDF)·생략대상을 국가기록원 지침대로 보고, 서식 7 수치를 낸다.

사용:  python pubs_check.py 간행물목록.xlsx --year 2025 -o 산출/간행물_점검.xlsx
목록 열(머리글 이름으로 찾음): 제목, 발간등록번호, 발행일, 발간형태(책자·전자), 간행물유형(팸플릿·리플릿·소식지 등), 송부일, 송부부수, PDF송부(Y/N), 공개구분, 비밀구분
근거:  공공기록물법 제22조, 시행규칙 제36조(발간등록)·제37조(분류)
       국가기록원 「2026년 기록물관리 지침」 간행물 관리:
         원고 확정 후 인쇄 전에 발간등록번호 신청, 앞표지 좌측 상단에 표기
         발간 주기(번호 뒤 2자리): 단행본 01, 연속간행물 02~15. 연속간행물은 제목·발행기관·발간주기가 바뀌면 새 번호
         책자형태 간행물은 국가기록원에 책자 3부와 같은 내용의 전자파일(PDF) 1점을 함께 송부(관할 기록관·영구기록물관리기관에도 따로 송부)
         송부 기한: 발간 후 15일이 지나기 전까지(시행령 제55조④, 국가기록포털 발간등록 안내 2026-11-20 확인)
         발간등록 생략대상(수첩·팸플릿·리플릿·속보·만화·기관신문·추록, 홍보자료·소식지·안내지·지도, 장학·학습·교육원 교재·교과서, 관보·보도자료 등)은 신청·송부 생략
         비밀기록물을 책자로 만든 것은 간행물이 아니라 비밀기록물 절차
       생산현황 통보서식 서식 7: 생산량·송부량은 부수가 아닌 종(철) 수, 생략대상 제외, 없으면 0
       국가기록원 「2026년도 간행물 발간등록 및 송부지침」(2026. 10.): 붙임1 영구기록물관리기관 기호표(앞 2자리), 붙임2 발간주기 구분기호(뒤 2자리),
         Ⅴ 생략대상 유형(소책자·카드뉴스·기관신문·추록·인쇄물 / 홍보자료·교육자료(교재·교과서·교범·시스템 매뉴얼)·공보자료(관보·보도자료·출장보고서)·
         행사물(작품집·행사진행자료·세미나·공청회·심포지엄·포럼)·가공·편집 간행물), 생략대상 기관(시행령 제3조: 공공기관·지방공사·공단·출자출연기관·특별법인·학교),
         송부 방식 우편 일원화, 연구용역보고서 발행기관 = 발주기관, 여러 권은 권마다 번호, 비공개·부분공개는 신청 때 호·사유(부분공개는 목차·쪽)
원칙:  점검과 수치 계산만 한다. 발간등록 신청·송부는 담당자가 국가기록포털에서 한다.
"""
import argparse
import collections
import os
import re
import sys

NUM = re.compile(r"^(\d{2})-(\d{7})-(\d{6})-(\d{2})$")
# 「2026년도 간행물 발간등록 및 송부지침」 붙임 1·2
ARCH = {"11": "국가기록원", "12": "대통령기록관", "31": "국회", "32": "법원", "33": "헌법재판소", "34": "중앙선거관리위원회", "51": "서울특별시", "52": "부산광역시",
        "53": "대구광역시", "54": "인천광역시", "55": "광주광역시(통합 전 발간분)", "56": "대전광역시", "57": "울산광역시", "58": "세종특별자치시", "59": "전남광주통합특별시",
        "71": "경기도", "72": "강원특별자치도", "73": "충청북도", "74": "충청남도", "75": "경상북도", "76": "경상남도", "77": "전북특별자치도", "78": "전라남도(통합 전 발간분)",
        "79": "제주특별자치도", "84": "충청남도교육청", "86": "경상남도교육청", "91": "청주시", "92": "이천시"}
CYCLE = {1: "단행본", 2: "일간", 3: "주간", 4: "격주간", 5: "순간", 6: "월간", 7: "격월간", 8: "계간", 9: "반년간", 10: "연간", 11: "격년간", 12: "3년간", 13: "5년간",
         14: "부정기", 15: "4년간"}
SKIP_ORG = ["공사", "공단", "출자", "출연", "재단", "진흥원", "대학교", "고등학교", "중학교", "초등학교", "유치원"]
SKIP = ["카드뉴스", "포스터", "전단", "낱장", "작품집", "글모음", "도록", "행사진행", "세미나", "공청회", "심포지엄", "포럼", "출장보고", "연수보고", "교범", "시스템 매뉴얼", "사용자 매뉴얼",
        "수첩", "팸플릿", "팜플렛", "리플릿", "리플렛", "속보", "만화", "웹툰", "기관신문", "신문", "추록", "홍보", "소식지", "안내지", "브로셔",
        "브로슈어", "지도", "장학", "학습교재", "교재", "교과서", "관보", "구보", "군보", "시보", "보도자료"]
TITLE_SKIP = ["카드뉴스", "출장보고서", "연수보고서", "작품집", "행사 진행", "사용자 매뉴얼", "수첩", "팸플릿", "팜플렛", "리플릿", "리플렛", "웹툰", "기관신문", "추록", "소식지", "안내지",
              "브로셔", "브로슈어", "교과서", "관보", "보도자료", "홍보물"]
ALIAS = {"제목": ["제목", "간행물명"], "번호": ["발간등록번호"], "발행일": ["발행일", "발간일"], "형태": ["발간형태", "형태"],
         "유형": ["간행물유형", "유형"], "송부일": ["송부일", "송부일자"], "부수": ["송부부수"], "pdf": ["PDF송부", "전자파일송부", "전자파일"],
         "공개": ["공개구분"], "비밀": ["비밀구분"], "기관": ["발행기관", "기관명"], "사유": ["비공개사유", "공개제한사유"],
         "원고": ["최종본등록", "원고최종본등록", "최종본 전자파일 등록"]}


def _d(v):
    import datetime as _dt
    t = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y%m%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--today", help="송부 기한 판정일(기본: 오늘)")
    ap.add_argument("-o", "--out", default="산출/간행물_점검.xlsx")
    a = ap.parse_args()
    import datetime as _dt
    _today = _dt.date.fromisoformat(a.today) if a.today else _dt.date.today()
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    items = [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]
    by_no = collections.defaultdict(list)
    for n, x in items:
        if x.get("번호"):
            by_no[x["번호"]].append(x.get("제목", ""))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "제목", "발간등록번호", "발간주기", "생략대상 후보", "송부", "판정"])
    produced, sent = set(), set()
    for n, x in items:
        issues = []
        title, no = x.get("제목", ""), x.get("번호", "")
        if x.get("비밀", "") not in ("", "일반", "N", "해당없음"):
            ws.append([n, "(비밀 표시)", "", "", "", "", "간행물이 아니라 비밀기록물 절차로 관리"])
            continue
        kind = x.get("유형", "")
        skip = next((k for k in SKIP if k in kind), "") or next((k for k in TITLE_SKIP if k in title), "")
        cycle = ""
        m = NUM.match(no)
        if no and not m:
            issues.append(f"발간등록번호 형식 '{no}' — NN-NNNNNNN-NNNNNN-NN(예: 11-1311153-000040-10)")
        if m:
            c = int(m[4])
            cycle = f"{CYCLE[c]}({m[4]})" if c in CYCLE else f"주기 코드 {m[4]} — 01~15가 아님(지침 붙임 2)"
            if c not in CYCLE:
                issues.append(f"발간주기 구분기호 {m[4]} — 01 단행본, 02~15 연속간행물(지침 붙임 2)")
            if m[1] not in ARCH:
                issues.append(f"영구기록물관리기관 기호 {m[1]} — 기호표(지침 붙임 1)에 없음")
            if c == 1 and len(by_no[no]) > 1:
                issues.append(f"단행본 번호가 {len(by_no[no])}건에 중복 — 단행본은 새 번호를 받는다")
            if c != 1 and len({re.sub(r'\s|\d{4}년?|제?\d+[호권집]', '', t) for t in by_no[no]}) > 1:
                issues.append("같은 연속간행물 번호에 제목이 다른 간행물 — 제목이 바뀌었으면 변경신청으로 새 번호")
        org = x.get("기관", "")
        org_skip = next((k for k in SKIP_ORG if k in org), "")
        if org_skip and not no:
            skip = skip or f"생략대상 기관 후보({org_skip})"
        if "공개" in idx and x.get("공개", "") in ("비공개", "부분공개") and not re.search(r"[1-8]\s*호", x.get("사유", "") + x.get("공개", "")):
            issues.append("비공개·부분공개 간행물 — 신청 때 정보공개법 제9조①의 호와 사유(부분공개는 목차·쪽)를 적음")
        if "원고" in idx and no and x.get("원고", "").upper() not in ("Y", "O", "예", "완료"):
            issues.append("간행물 최종본 전자파일을 발간 결재 붙임으로 등록했는지 확인(인쇄 전 원고만 결재한 사례 많음)")
        if re.search(r"행정간행물등록번호|정부간행물등록번호", title + no):
            issues.append("'발간등록번호'로 표기(행정·정부간행물등록번호 등 유사 명칭 금지)")
        if not no and not skip:
            issues.append("발간등록번호 없음 — 인쇄 전에 신청했어야 함(생략대상이면 무시)")
        if no and skip:
            issues.append(f"생략대상 후보('{skip}')인데 번호를 받음 — 지침 Ⅴ 생략대상 유형표로 확인. 받았다면 번호를 표기하고 송부 — 생략대상은 송부해도 영구보존 대상이 아닐 수 있음")
        send = ""
        if no:
            is_book = "전자" not in x.get("형태", "") or "책자" in x.get("형태", "")
            copies = int(re.sub(r"\D", "", x.get("부수", "")) or 0)
            pdf = x.get("pdf", "").upper() in ("Y", "O", "예", "완료")
            if is_book:
                send = f"책자 {copies}부·PDF {'있음' if pdf else '없음'}"
                if copies < 3:
                    issues.append(f"책자 {copies}부 송부 — 국가기록원에 3부")
                if not pdf:
                    issues.append("책자와 같은 PDF 1점을 함께 송부·업로드")
            else:
                send = f"전자파일 {'있음' if pdf else '없음'}"
                if not pdf:
                    issues.append("전자간행물은 전자파일 1점을 업로드")
        if no and not skip:
            pd_, sd_ = _d(x.get("발행일")), _d(x.get("송부일"))
            if pd_ and sd_ and (sd_ - pd_).days > 15:
                issues.append(f"발간 후 {(sd_ - pd_).days}일 만에 송부 — 발간 후 15일이 지나기 전까지(시행령 제55조④)")
            elif pd_ and not sd_ and (_today - pd_).days > 15:
                issues.append(f"발간 후 {(_today - pd_).days}일 지났는데 송부일 없음 — 15일 안 3부(관할 기록관·관할 영구기록물관리기관·중앙기록물관리기관)")
        year_ok = str(a.year) in x.get("발행일", "")
        if year_ok and m and not skip:
            produced.add(no)
            if (x.get("송부일") or "").strip():
                sent.add(no)
        ws.append([n, title, no, cycle, skip, send, " · ".join(issues) or "통과"])
    s7 = wb.create_sheet("서식 7 수치")
    s7.append([f"{a.year}년 간행물 생산량(종) ①", "간행물 송부량(종) ②", "기준"])
    produced, sent = len(produced), len(sent)
    s7.append([produced, sent, "발행일이 해당 연도·형식이 맞는 발간등록번호(같은 번호는 1종)·생략대상 제외, 송부량은 송부일이 있는 종 수(부수 아님)"])
    for i, w in enumerate((8, 40, 24, 16, 14, 20, 70)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    bad = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[6] != "통과")
    print(f"간행물 {len(items)}종 · 고칠 것 있는 종 {bad} · 서식 7 생산량 {produced} · 송부량 {sent} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
