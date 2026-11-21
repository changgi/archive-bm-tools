#!/usr/bin/env python3
"""학교기록물 점검 — 유치원·초·중·고·특수학교의 등록 목록으로 사안별·유형별 관리, 이관 시기, 건명 개인정보, 폐교 인계를 본다.

근거: 국가기록원 「2025년 학교기록물 관리 지침」(2025. 1.)
  4.2.4 정리(다음 해 2월 말까지) · 4.2.5 생산현황 통보(5월 31일, 표 1 유형별 대상) · 4.2.6 이관(보존기간 기산일부터 2년 범위 보관,
        K-에듀파인 전자기록물은 1년 단위, 민원·소송 등은 기록관과 협의해 기산일부터 10년 범위 연장, 사본 분리·표시)
  4.2.8 학교는 자체 폐기 불가 — 교육(지원)청 기록관이 주관 · 4.2.10 문서고(이중잠금·제한구역, 20±2℃·50±5%, 100~300lux, 대장 3종)
  4.2.11 열람(30년 이상 비전자 원본은 책임자 입회, 반출 시 반출·반입대장)
  5.1 학적(졸업대장: NEIS 출력 → 비전자 등록, 건명 'YYYY학년도 제N회 졸업대장', 생산년도는 졸업일) · 5.2 시험(평가원안·답안지 건명, 이의신청 뒤 문서고 → 이관)
  5.3 학교연혁지('학교역사기록물관리' 단위과제) · 5.4 학교폭력(같은 학교폭력 단위과제에 일관 편철, 사안 종료 후 이관) · 5.5 보건(보건일지 출력·비전자 등록)
  5.6 급식(수기 급식일지·CP/CCP 일지는 학년도 단위 비전자 등록) · 5.7 동의서·신청서·설문지(개인정보 보호법으로 처분하지 않고 공공기록물법 절차)
  6.2 졸업앨범은 발간 즉시 등록 · 6.4 대장류(여러 해 쓰는 대장은 생산 종료·분철하는 해에 등록, NEIS 대장은 출력하지 않음)
  7 폐교(보유기록물 목록 → 실사·정리 → 인계·인수 계획, 승계학교 유무, 공인은 기록관으로, 생산 당시 학교명·기관코드 유지)
  부록 해석례: 사립유치원은 공공기관(유치원생활기록부 포함), 공립어린이집은 아님, 수기 출석부도 관리, 스캔 첨부해도 비전자 증빙 원본 보존,
             개인정보가 든 문서 제목은 이력을 남기고 수정 가능, 붙임 암호 설정 금지

입력: K-에듀파인 등록 목록(xlsx·csv) — 열(머리글 이름으로 찾음): 건명(제목), 생산일자(등록일자), 단위과제(명), 전자구분(전자·비전자),
      이관일자, 이관연장(Y), 비치(Y), 공개구분, 처분(있으면)
결과: xlsx — 점검 · 이관 시기 · 생산현황 유형(표 1) · 문서고 대장(빈 서식 3종) · [폐교] 인계 계획
판단(보존기간·폐기·공개 결정)은 하지 않는다. 학교는 기록물을 자체 폐기할 수 없다.

예) python school_check.py 등록목록.xlsx --today 2026-03-02 -o 산출/학교기록물_점검.xlsx
    python school_check.py 등록목록.xlsx --closing 2026-08-31 --successor "가상초등학교" -o 산출/폐교_인계.xlsx
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

ALIAS = {"건명": ["건명", "제목", "기록물철제목"], "생산": ["생산일자", "등록일자", "생산일"], "과제": ["단위과제", "단위과제명", "단위과제카드"],
         "전자": ["전자구분", "구분", "전자여부"], "이관": ["이관일자", "이관일"], "연장": ["이관연장", "이관시기연장"], "비치": ["비치", "비치여부"],
         "공개": ["공개구분", "공개여부"], "처분": ["처분", "처리"], "번호": ["문서번호", "등록번호", "생산등록번호"]}
PII = [(re.compile(r"[가-힣]○{1,2}|[가-힣]{1}[*＊]{1,2}"), "가린 이름"), (re.compile(r"\d-\d{1,2}\s*[가-힣]{2,4}"), "학년-반 + 이름"),
       (re.compile(r"\d{6}-?[1-4]\d{6}"), "주민등록번호 형태"), (re.compile(r"01[016789]-?\d{3,4}-?\d{4}"), "휴대전화")]
TYPES = [("회의록", r"회의록|협의회|위원회 회의"), ("시청각기록물", r"졸업앨범|사진|동영상|영상|녹음"), ("행정박물", r"관인|상패|상장|기념품|교기|교패"),
         ("간행물", r"년사|교지|규정집|문집")]


def d(v):
    t = re.sub(r"\s", "", str(v or ""))[:10].rstrip(".")
    for f in ("%Y-%m-%d", "%Y.%m.%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


def load(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        rows = [["" if v is None else str(v).strip() for v in r] for r in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)]
    else:
        rows = [[c.strip() for c in r] for r in csv.reader(open(path, encoding="utf-8-sig"))]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in al if n in head}
    return [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)], idx


def school_year(day):
    return day.year - 1 if day.month <= 2 else day.year


def main() -> int:
    ap = argparse.ArgumentParser(description="학교기록물 점검(2025년 학교기록물 관리 지침)")
    ap.add_argument("listing")
    ap.add_argument("--today")
    ap.add_argument("--closing", help="폐교일 YYYY-MM-DD — 있으면 인계 계획 시트")
    ap.add_argument("--successor", default="", help="승계학교 이름(없으면 비움: 관할 교육(지원)청 기록관으로 전부)")
    ap.add_argument("--per-box", type=int, default=10, help="이송 상자 하나에 넣을 철 수")
    ap.add_argument("-o", "--out", default="산출/학교기록물_점검.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    items, idx = load(a.listing)
    if not items or "건명" not in idx:
        print("입력 오류: '건명'(제목) 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "건명", "생산일자", "단위과제", "확인할 것"])
    tr = wb.create_sheet("이관 시기")
    tr.append(["원본 행", "건명", "전자구분", "보존기간 기산일", "이관할 때", "상태", "근거"])
    counts = {k: 0 for k, _ in TYPES}
    counts["전자·비전자문서"] = 0
    bad = 0
    school_violence_tasks = set()
    for n, x in items:
        t, made, task = x.get("건명", ""), d(x.get("생산")), x.get("과제", "")
        elec = "비전자" not in x.get("전자", "") and x.get("전자", "") != ""
        iss = []
        for rx, why in PII:
            if rx.search(t):
                iss.append(f"건명에 개인정보({why}) — 이력을 남기고 건명 수정(지침 부록 해석례 7)")
        if "졸업대장" in t:
            m = re.search(r"(\d{4})학년도\s*제\s*(\d+)\s*회\s*졸업대장", t)
            if not m:
                iss.append("졸업대장 건명은 'YYYY학년도 제N회 졸업대장'(5.1)")
            elif made and int(m[1]) != school_year(made):
                iss.append(f"졸업대장 학년도 {m[1]} ≠ 생산일자(졸업일) {made} 기준 {school_year(made)}학년도 — 생산년도는 졸업일(5.1)")
            if elec:
                iss.append("졸업대장은 NEIS에서 출력해 비전자로 등록하고 문서등록번호를 표기한 원본으로 관리(5.1)")
        if re.search(r"평가원안|답안지|시험지", t):
            if not re.search(r"\d{4}학년도\s*\d학기\s*\S*(고사|평가)\s*(평가원안|답안지)", t):
                iss.append("시험기록물 건명: '○○○○학년도 ○학기 중간고사 평가원안' / '… 답안지(○학년)'(5.2)")
            if "답안지" in t and not re.search(r"\d\s*학년|\d\s*반", t):
                iss.append("답안지는 시험유형·학년·반별로 묶은 1권을 1건으로 — 건명에 학년(반)(5.2)")
        if re.search(r"학교폭력|학폭", t):
            school_violence_tasks.add(task or "(단위과제 없음)")
            if task and "학교폭력" not in task:
                iss.append(f"학교폭력 기록물이 '{task}'에 편철 — 학교폭력 관련 단위과제에 일관 편철(5.4)")
        if re.search(r"연혁", t) and task and "역사" not in task:
            iss.append("학교연혁지는 '학교역사기록물관리' 단위과제에 편철(5.3)")
        if re.search(r"보건일지|급식일지|CCP|CP\s*일지|중요관리점|일반관리점", t) and elec:
            iss.append("자체 프로그램·수기 일지는 학년도 마감 때 출력해 비전자로 등록(5.5·5.6)")
        if re.search(r"동의서|신청서|설문지", t) and re.search(r"개인정보|파기", x.get("처분", "")):
            iss.append("동의서·신청서·설문지는 개인정보 보호법으로 파기하지 않고 공공기록물법 절차로 보존기간 경과 뒤 폐기(5.7)")
        if re.search(r"폐기", x.get("처분", "")) and not re.search(r"기록관|심의", x.get("처분", "")):
            iss.append("학교는 자체 폐기 불가 — 교육(지원)청 기록관이 평가·폐기 주관(4.2.8)")
        if re.search(r"출석부", t) and re.search(r"폐기|미등록", x.get("처분", "")):
            iss.append("수기 출석부도 공공기록물(해석례 5, 관련 지침 5년)")
        if re.search(r"졸업앨범", t) and made and not x.get("번호"):
            iss.append("졸업앨범은 발간 즉시 등록(6.2) — 등록번호 없음")
        if re.search(r"사본", t):
            iss.append("비전자 철에 사본이 섞이면 분리하거나 사본 표시 후 이관(4.2.6)")
        k = next((k for k, rx in TYPES if re.search(rx, t)), "전자·비전자문서")
        counts[k] += 1
        # 이관 시기(4.2.6)
        if made:
            base = dt.date(made.year + 1, 1, 1)
            ext = re.match(r"(Y|예|연장)", x.get("연장", "").strip())
            keep = x.get("비치", "").strip().upper() in ("Y", "예")
            if keep:
                due, why = None, "비치기록물 — 비치 사유가 끝나면 이관"
            elif ext:
                due, why = dt.date(base.year + 10, 1, 1), "민원·소송 등으로 기록관과 협의해 연장 — 기산일부터 10년 범위(단위과제카드 단위)"
            elif elec:
                due, why = base, "K-에듀파인 전자기록물 — 1년 단위로 전년도 생산분 이관"
            else:
                due, why = dt.date(base.year + 2, 1, 1), "비전자 — 기산일부터 2년 범위 보관 후 기록관 일정에 따라"
            if not d(x.get("이관")):
                st = "보관" if (due is None or due > today) else ("이관 대상" if (today - due).days <= 365 else "이관 늦음")
                tr.append([n, t, "전자" if elec else "비전자", base.isoformat(), due.isoformat() if due else "", st, why])
        if iss:
            bad += 1
        ws.append([n, t, made.isoformat() if made else x.get("생산", ""), task, " · ".join(iss) or "통과"])
    if len(school_violence_tasks) > 1:
        ws.append(["", "(학교폭력 기록물 전체)", "", " / ".join(sorted(school_violence_tasks)), "학교폭력 기록물이 여러 단위과제에 흩어짐 — 하나의 관련 단위과제로 같은 보존기간(5.4)"])
        bad += 1
    pc = wb.create_sheet("생산현황 유형(표 1)")
    pc.append(["유형", "건수(건명 낱말로 추정)", "작성 대상(지침 표 1)"])
    for k, why in [("전자·비전자문서", "문서(대장 포함)·도면·카드"), ("회의록", "각종 회의록"), ("시청각기록물", "사진·필름(졸업앨범 포함), 녹음·동영상"),
                   ("행정박물", "관인·견본·상징·기념·상장상패·사무집기 등"), ("간행물", "학교○○년사·교지·학교규정집 등")]:
        pc.append([k, counts[k], why])
    pc.append(["통보 기한", "매년 5월 31일", "정리(2월 말) 뒤 K-에듀파인으로 관할 기록관에 통보(4.2.5)"])
    for title, head in [("보존문서기록대장", ["연번", "단위과제", "기록물철 제목", "생산연도", "보존기간", "권수", "서가 위치", "비고"]),
                        ("반출·반입대장", ["연번", "반출일", "기록물철 제목", "반출자", "목적", "반입일", "확인자"]),
                        ("문서고 출입대장", ["일자", "입실", "퇴실", "소속·성명", "목적", "동행·확인"])]:
        sh = wb.create_sheet(f"문서고 {title}")
        sh.append(head)
        for c in sh[1]:
            c.font = Font(bold=True)
    if a.closing:
        cl = wb.create_sheet("폐교 인계 계획")
        dest = a.successor or "관할 교육(지원)청 기록관"
        cl.append(["상자", "원본 행", "건명", "생산일자", "인수 기관", "실물 확인", "비고"])
        for i, (n, x) in enumerate(items):
            dest_i = "관할 교육(지원)청 기록관" if re.search(r"(관인|직인|공인)(?!\s*(대장|관리|사용|날인))", x.get("건명", "")) else dest
            note = "공인은 승계학교가 있어도 기록관으로(7.2.1)" if dest_i != dest else ""
            cl.append([f"{i // a.per_box + 1:05d}", n, x.get("건명", ""), x.get("생산", ""), dest_i, "", note])
        ck = wb.create_sheet("폐교 체크리스트")
        ck.append(["단계", "할 일", "근거", "완료"])
        for row in [("사전 준비", "문서고·캐비닛 전부 조사해 보유기록물 목록 작성 → 관할 기록관·승계학교에 송부", "7.1"),
                    ("사전 준비", "인수기관 실사: 미편철·보존기간 미책정 정리, 사본·참고자료 분리", "7.1"),
                    ("계획", f"인계·인수 계획(목록·정리·포장·일정·인력·배차) — 폐교일 {a.closing}", "7.2"),
                    ("인계", "상자 겉면에 일련번호+목록 부착, 실물 일치 확인 후 밀봉, 당일 차량 이송", "7.2.1"),
                    ("인계", "전자기록물: K-에듀파인 과제카드 종료요청 → 인계등록·인계요청 → 승계학교 인수접수", "7.2.1"),
                    ("인계", "공인(관인)은 관할 교육(지원)청 기록관으로", "7.2.1"),
                    ("인수 후", "실물·목록 재확인, 인수결과를 폐교·기록관에 통보", "7.3"),
                    ("인수 후", "생산 당시 학교명·기관코드 유지 — 표지 교체·등록정보 수정·재편철 금지(맥락 보존)", "7.3")]:
            ck.append([*row, ""])
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True)
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["E"].width = 90
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    late = sum(1 for r in tr.iter_rows(min_row=2, values_only=True) if r[5] == "이관 늦음")
    print(f"{len(items)}건 · 확인할 것 {bad} · 이관 늦음 {late}" + (f" · 폐교 인계 상자 {(len(items) - 1) // a.per_box + 1}" if a.closing else "") + f" → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
