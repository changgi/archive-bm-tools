#!/usr/bin/env python3
"""기록물관리기관 보안·재난관리 대책 점검: 보안·재난관리 체크리스트, 위험평가표(발생가능성 × 위험도), 업무 유형별 근무자 안전규칙,
개인별 임무카드, 비상용품 관리대장, 대응·복구 절차 항목, 재난 보고서 서식, (영구기록물관리기관) 전자기록물 RTO·RPO 재난복구 모델을 만든다.
근거: NAK 2-1:2024(v1.2) 「기록물관리기관 보안 및 재난관리 기준」(국가기록원, 2026-11-20 내려받아 확인)

사용:  python sd_plan.py --org 기록관 [--risk 위험평가.xlsx] [--staff 임무카드.xlsx] [--supplies 비상용품.xlsx]
                        [--bia 업무영향.xlsx] [--work 서고관리,상태검사] [--last-review 2025-11-01] [--today 2026-11-20] -o 산출/보안재난.xlsx
입력(모두 선택, 머리글 이름으로 찾음):
  위험평가: 구분, 위협(위협의 종류), 내용, 발생가능성, 위험도(1~5 또는 매우낮음·낮음·보통·높음·매우높음), 대책
  임무카드: 이름, 소속, 평시업무, 재난시반, 재난시업무, 대행, 구내, 휴대폰
  비상용품: 품명, 보존위치, 수량, 단위, 상태, 유통기한
  업무영향(영구): 업무프로세스, RTO(시간), RPO(최종시점·D-1)
규칙:
  5.5   보안대책은 연 1회 이상 검토 → --last-review가 1년 지났으면 표시
  6.4.2 위험평가표 작성(표 1, 부속서 E 위협 12종) — 점수 = 발생가능성 × 위험도, 15점 이상은 점검 강화(6.4.4)
  6.5.1 개인별 임무카드: 비상연락처, 재난 유형별 임무, 같은 조원(부속서 D)
  6.5.5 비상용품: 출입구 근처, 목록 관리, 주기적으로 수량·위치·유통기한 확인(표 5)
  7.1.2 RTO 1: 4시간, 2: 12시간, 3: 24시간, 4: 7일, 5: 7일 넘음 / RPO 1: 최종시점, 2: D-1 → Tier·DR 모델(표 7)
원칙:  계획·점검표의 빈칸과 기한만 계산한다. 위험 점수·임무 배정은 기관이 정한다. 비밀기록물 목록은 열지 않는다.
"""
import argparse
import datetime as dt
import os
import re
import sys

LEVEL = {"매우낮음": 1, "낮음": 2, "보통": 3, "높음": 4, "매우높음": 5}

SEC_CHECK = [  # 부속서 A
    ("보안성 검토", "기관 특성에 따른 보안성 검토를 수행하였는가?"),
    ("인원보안", "출입인원을 식별하고 접근권한을 통제하고 있는가?(출입대장, 외부인 보안서약서·보안교육, 5.3.1)"),
    ("시설보안", "제한구역·통제구역 등 보호구역이 적합하게 설정되어 있는가?(보존·작업공간과 전산시설은 보호구역, 5.3.2)"),
    ("시설보안", "CCTV, 출입카드, 이중 잠금장치 등 보안장비가 운용되고 있는가?"),
    ("기록물보안", "비밀기록물의 보안 상태는 적절한가?(별도 보안 조치)"),
    ("기록물보안", "기록물의 열람·취급은 적절히 통제되는가?(직접 취급은 인가 직원, 위탁 시 감독자 지정, 반입·반출 통제 5.3.3)"),
    ("시스템 보안", "기록관리시스템과 스토리지에 침해 위험은 없는가?(사용자 인증·접속이력, 침입차단, 운영 장소 통제구역, 공중망 전송 금지 5.3.4)"),
    ("보안대책 수립·운용", "보안대책이 수립되어 있는가?(기관장 결재 포함)"),
    ("보안대책 수립·운용", "보안대책이 실행되고 있는가? 연 1회 이상 검토하는가?(5.5)"),
    ("보안대책 수립·운용", "보안대책에 책임과 역할이 분명히 적혀 있는가?"),
    ("보안대책 수립·운용", "보안사고 대응 조직·비상연락체계·유형별 대응 시나리오가 있는가?(5.4)"),
    ("교육·훈련", "직원과 외부 작업인력에게 보안교육을 실시하는가?"),
    ("교육·훈련", "직원들이 보안 지식을 숙지하고 있는가?"),
]
DIS_CHECK = [  # 부속서 B + 6.5·6.6
    ("계획", "재난관리계획을 수립하였는가?(기관장 결재 포함)"),
    ("예방", "기관 특성에 따른 위험평가를 수행하였는가?(위험평가표)"),
    ("예방", "기록물과 시설·장비를 주기적으로 점검하는가?(정수·상태점검, 서고·서가·온습도·공기질·조명, 공조·전기·통신)"),
    ("예방", "위험이 내재된 업무에 근무자 안전규칙을 수립하고 작업 구역에 게시했는가?"),
    ("대비", "비상조직체계와 직원·관련자 비상연락망을 수립했는가? 개인별 임무카드가 있는가?"),
    ("대비", "기록물 대피 우선순위(1~3등급)와 필수기록물을 지정했는가? 중요기록물에 등급별 식별표식을 붙였는가?"),
    ("대비", "영구 중요기록물을 성격이 다른 두 매체 이상으로 이중보존하고 떨어진 곳에 분산했는가?(공공기록물법 제21조)"),
    ("대비", "비상용품을 출입구 근처에 두고 목록·수량·유통기한을 관리하는가?"),
    ("대응", "재난 유형별 대응 절차(직원 행동요령, 기록물·인명 대피경로 도면, 피난 장소, 응급조치)가 있는가?"),
    ("대응", "재난이나 징후를 발견하면 시장·군수·구청장·긴급구조기관에 신고할 비상연락망이 있는가?(6.6.1)"),
    ("복구", "부서별 복구 계획, 자문기관·전문기관 연락처, 임시 작업장, 손상 기록물 이송 방법이 준비되었는가?"),
    ("복구", "손상된 비밀기록물은 일반 기록물과 분리하고 이송 때 비밀취급 인가자가 동행하는가?(6.6.3)"),
    ("사후관리", "반출 기록물은 현장 안정 전까지 재배치하지 않고, 재배치 뒤 정수점검을 하는가?(6.6.4)"),
    ("사후관리", "재난 보고서를 쓰고 재난관리계획을 보완·교육하는가?(6.6.4)"),
    ("교육·훈련", "직원들이 재난 대응 지식을 숙지하고 있는가? 시나리오 훈련(상황 발생→비상조직 편성→계획 실행→종합 보고→종료)을 하는가?"),
]
ELEC_CHECK = [  # 7장(영구기록물관리기관)
    ("7.1.1", "전자기록물 관리·보존·활용 업무 현황을 분석했는가?"),
    ("7.1.2", "위험분석(자연·기술적·인적 위협, 발생가능성·피해 정도 5단계)과 업무영향분석(RTO·RPO)을 했는가?"),
    ("7.1.3", "업무연속성계획과 재난복구계획을 재난관리계획에 반영했는가?"),
    ("7.1.4", "재난복구시스템 구현 가이드라인(원칙·솔루션·운영·발전방향)을 작성했는가?"),
    ("7.4~7.5", "IT시스템 복구 절차서를 작성하고 평소 훈련하는가?"),
]
THREATS = [  # 부속서 E
    ("자연재해", "홍수/태풍", "보존상태 훼손, 정전 등 2차 재해"), ("자연재해", "산불/화재", "시설·설비 파괴"), ("자연재해", "시설물 붕괴", "기록물·시스템 붕괴"),
    ("기술적 재해", "시설 결함", "정전 등으로 시스템 운영 장애"), ("기술적 재해", "보안장비 이상", "무단침입·임의 접근"),
    ("기술적 재해", "통신 인프라 이상", "통신 장애로 시스템 운영 장애"), ("기술적 재해", "업무관리시스템 이상", "업무 장애·전자기록물 보존 이상"),
    ("기술적 재해", "DB/스토리지 이상", "전자기록물 보존 위험"), ("인적 재해", "컴퓨터 바이러스", "시스템 운영·업무 장애"),
    ("인적 재해", "중요 인물의 직무 불능", "대체 불가 인력 공백"), ("인적 재해", "해킹/시스템 공격", "정보보안 위험·시스템 파괴"),
    ("인적 재해", "고의적 파괴", "시스템 완전 파괴·백업 장애"),
]
SAFETY = {  # 표 2
    "공통": ["소화기 사용법 숙지", "작업실 내 흡연·음식물 반입 금지", "전기제품 취급 유의사항 준수", "개인보호구 착용"],
    "인수·등록": ["하역 시 안전화·마스크·미끄럼 방지 장갑", "운반 시 적재함 용량 고려", "상자를 던지거나 떨어뜨리지 않기", "칼·송곳은 주변 작업자 고려", "작업 뒤 장비 정리"],
    "상태검사": ["방진마스크·앞치마·토시·장갑 착용", "장비 사용 시 물·약품 관리", "작업 종료·퇴실 시 기록물은 서고로 입고"],
    "보존상자 제작": ["원지 사용 시 장갑", "장비 가동 중 안전거리", "구급약 비치", "전열기구·인화물질 사용 금지"],
    "제본·편철": ["방진마스크·앞치마·토시·장갑 착용", "칼·송곳은 장갑 착용 후 사용", "구급약 비치", "작업 종료·퇴실 시 기록물은 서고로 입고"],
    "서고관리": ["서고 내 작업은 2시간 이내, 근무 위치를 관리자에게 알림", "호흡기 질환 의심 시 서고 출입 금지", "이상 감지 시 즉시 서고 밖으로",
              "기록물 이동은 북트럭, 쏟아지지 않게", "대량·대형 이동 시 공간 확보", "저온서고는 2인 이상·방한복", "5단 이상은 이동식 계단·사다리",
              "원본 취급 시 마스크·장갑", "전열기구·인화물질 사용 금지"],
    "탈산·소독": ["작업 중 환기시설 가동", "방진마스크 착용, 소화기 위치 확인", "약제는 환기되는 곳에 보관", "가동 중 출입문에 작업 중 푯말", "방출 시 장비 내 압력이 상압인지 확인"],
    "마이크로필름·시청각": ["약품은 사용 뒤 지정 보관함", "약품 처리 시 보호장갑·보호안경", "관계자 외 출입 통제", "현상실은 지정 출입문", "현상 시 롤러에 옷 감김 주의",
                    "폐액은 전문처리업체", "밀폐 공간 작업 뒤 실외 10분 이상 휴식", "질산염 필름은 정전기 방지 장갑·보호안경, 25℃·50% 정도", "작업대에 음식물 금지"],
    "복원·복제": ["화학약품 작업 전후 충분한 환기", "화학약품 목록 작성·관리", "개인보호용구 착용", "약품은 직사광선 피해 전용 보관함", "장비 작동 매뉴얼 작성·관리",
              "칼·송곳은 주변 작업자 고려"],
}
DUTY = {"화재/산불": "피해(예상) 기록물 포장 — 보호장구, 포장용품, 관내 소산지역 확보, 우선등급별 5~10kg 단위 포장, 상자태그(우선등급·서고-서가-열-단·권수)",
        "풍수해": "피해(예상) 기록물 포장 — 화재/산불과 같음", "산사태": "피해(예상) 기록물 포장 — 화재/산불과 같음",
        "장애": "서고 조치 — 항온항습 시간대별 점검, 저온서고 온습도 점검, 휴대용 냉동장비·에어컨·팬 섭외"}
SALIAS = {"이름": ["이름", "성명"], "소속": ["소속", "부서"], "평시": ["평시업무", "평시 업무", "주요업무"], "반": ["재난시반", "재난시 반", "재난시소속"],
          "임무": ["재난시업무", "재난시 업무", "임무"], "대행": ["대행", "업무대행", "조원"], "구내": ["구내", "내선"], "휴대폰": ["휴대폰", "휴대전화"]}
RALIAS = {"구분": ["구분"], "위협": ["위협", "위협의 종류", "재난유형", "재난 유형"], "내용": ["내용", "위협의 내용"], "가능성": ["발생가능성", "발생 가능성"],
          "위험도": ["위험도", "위험규모", "심각성"], "대책": ["대책", "대책 및 실천방안"]}
UALIAS = {"품명": ["품명", "품목"], "위치": ["보존위치", "위치"], "수량": ["수량"], "단위": ["단위"], "상태": ["상태"], "기한": ["유통기한", "사용기한"]}
BALIAS = {"업무": ["업무프로세스", "업무 프로세스", "업무"], "RTO": ["RTO", "RTO(시간)", "복구시간목표"], "RPO": ["RPO", "복구시점목표"]}


def read(path, alias):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in alias.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {}
    for k, al in alias.items():
        for n in al:
            if n in head and k not in idx:
                idx[k] = head.index(n)
    return [{k: (r[i] if i < len(r) else "") for k, i in idx.items()} for r in rows[hi + 1:] if any(r)]


def lvl(v):
    s = re.sub(r"\s", "", str(v or ""))
    if s in LEVEL:
        return LEVEL[s]
    m = re.search(r"[1-5]", s)
    return int(m[0]) if m else None


def day(v):
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(v or ""))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def tier(rto_h, rpo):
    last = not re.search(r"D-?1|하루|1일", rpo or "")
    if rto_h is None:
        return "", "", "RTO(시간)를 적어야 함"
    rt = 1 if rto_h <= 4 else 2 if rto_h <= 12 else 3 if rto_h <= 24 else 4 if rto_h <= 168 else 5
    if rt == 1 and last:
        return "Tier 1", "HOT", ""
    if rt == 2 and last:
        return "Tier 2", "HOT", ""
    if rt <= 3 and last:
        return "Tier 3", "WARM", "" if rt == 3 else "RTO는 짧지만 표 7 예시 범위 밖 — 기관이 확인"
    if rt <= 3:
        return "Tier 4", "WARM", ""
    if rt == 4:
        return "Tier 5", "COLD", ""
    return "Tier 6", "없음", ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--org", choices=["기록관", "특수기록관", "영구"], default="기록관")
    ap.add_argument("--risk")
    ap.add_argument("--staff")
    ap.add_argument("--supplies")
    ap.add_argument("--bia", help="영구기록물관리기관: 업무프로세스·RTO(시간)·RPO")
    ap.add_argument("--work", default="", help="안전규칙을 뽑을 업무 유형(쉼표) — 비우면 전부")
    ap.add_argument("--last-review", help="보안대책 마지막 검토일 YYYY-MM-DD")
    ap.add_argument("--today")
    ap.add_argument("-o", "--out", default="산출/보안재난.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    summary = []

    sc = wb.active
    sc.title = "보안 체크리스트"
    sc.append(["영역", "점검 항목(NAK 2-1 부속서 A)", "이행여부", "근거자료"])
    for r in SEC_CHECK:
        sc.append([*r, "", ""])
    if a.last_review:
        lr = day(a.last_review)
        if lr:
            due = dt.date(lr.year + 1, lr.month, min(lr.day, 28))
            late = today > due
            sc.append(["검토 주기", f"보안대책 마지막 검토 {lr.isoformat()} → 다음 검토 {due.isoformat()}(연 1회 이상, 5.5)", "지남" if late else "", ""])
            if late:
                sc.cell(sc.max_row, 3).fill = red
                summary.append("보안대책 연 1회 검토 지남")
    dc = wb.create_sheet("재난관리 체크리스트")
    dc.append(["영역", "점검 항목(NAK 2-1 부속서 B·6장)", "이행여부", "근거자료"])
    for r in DIS_CHECK:
        dc.append([*r, "", ""])
    if a.org == "영구":
        dc.append([])
        dc.append(["전자기록물(7장)", "영구기록물관리기관 전자기록물 재난관리(공공기록물법 제30조②)", "", ""])
        for r in ELEC_CHECK:
            dc.append([*r, "", ""])

    rk = wb.create_sheet("위험평가표")
    rk.append(["구분", "위협(재난 유형)", "내용", "발생가능성(1~5)", "위험도(1~5)", "점수", "판정", "대책 및 실천방안", "확인할 것"])
    rows = read(a.risk, RALIAS) if a.risk else [{"구분": g, "위협": t, "내용": c} for g, t, c in THREATS]
    high = blank = 0
    scored = []
    for r in rows:
        p, s = lvl(r.get("가능성")), lvl(r.get("위험도"))
        sc_ = p * s if p and s else None
        verdict = "" if sc_ is None else "높음 — 점검 강화·대책 우선(6.4.4)" if sc_ >= 15 else "보통" if sc_ >= 8 else "낮음"
        note = []
        if sc_ is None:
            note.append("발생가능성·위험도를 1~5(매우낮음~매우높음)로")
            blank += 1
        elif sc_ >= 15 and not r.get("대책"):
            note.append("높은 위험인데 대책 없음")
        high += bool(sc_ and sc_ >= 15)
        scored.append((-(sc_ or 0), [r.get("구분", ""), r.get("위협", ""), r.get("내용", ""), p or "", s or "", sc_ or "", verdict, r.get("대책", ""), " · ".join(note)]))
    for _, line in sorted(scored, key=lambda x: x[0]):
        rk.append(line)
        if line[5] and line[5] >= 15:
            rk.cell(rk.max_row, 7).fill = red
    rk.append([])
    rk.append(["기준", "점수 = 발생가능성 × 위험도(각 1 매우낮음 ~ 5 매우높음). 15 이상 높음, 8~14 보통, 7 이하 낮음 — 이 킷의 구간. 위험 요소 예: 상습침수지역, 인근 화학공장, 산사태, 안전진단 취약점"])
    summary.append(f"위험 높음 {high}" + (f" · 미평가 {blank}" if blank else ""))

    sf = wb.create_sheet("안전규칙(표 2)")
    sf.append(["업무 유형", "근무자 안전규칙 — 해당 업무 구역에 게시(6.4.3)"])
    want = [w.strip() for w in a.work.split(",") if w.strip()]
    for k, rules in SAFETY.items():
        if want and k != "공통" and not any(w in k for w in want):
            continue
        for x in rules:
            sf.append([k, x])

    mc = wb.create_sheet("임무카드(부속서 D)")
    mc.append(["이름", "소속", "평시 업무", "재난 시 반", "재난 시 업무", "업무 대행·조원", "구내", "휴대폰", "확인할 것"])
    nstaff = nmiss = 0
    if a.staff:
        for r in read(a.staff, SALIAS):
            nstaff += 1
            miss = [n for k, n in (("반", "재난 시 반"), ("임무", "재난 시 업무"), ("대행", "대행·조원"), ("휴대폰", "휴대폰")) if not r.get(k)]
            nmiss += bool(miss)
            mc.append([r.get(k, "") for k in ("이름", "소속", "평시", "반", "임무", "대행", "구내", "휴대폰")] + [("빈칸: " + ", ".join(miss)) if miss else ""])
    else:
        mc.append(["(--staff로 직원 목록을 주면 채움)"])
    mc.append([])
    mc.append(["카드에 넣을 연락처(6.5.1)", "직원, 시설·장비 유지보수 업체, 소방서·병원·경찰·가스·전력·기상, 보존·복원 자문 전문가, 긴급 이송 장비·업체"])
    for k, v in DUTY.items():
        mc.append([f"유형별 임무 예시 — {k}", v])
    if a.staff:
        summary.append(f"임무카드 {nstaff}명 · 빈칸 있는 카드 {nmiss}")

    su = wb.create_sheet("비상용품 대장(표 5)")
    su.append([f"점검일: {today.isoformat()}", "확인자:"])
    su.append(["연번", "품명", "보존위치", "수량", "단위", "상태", "유통기한", "확인할 것"])
    nbad = 0
    if a.supplies:
        for i, r in enumerate(read(a.supplies, UALIAS), 1):
            note = []
            q = re.search(r"\d+", r.get("수량", ""))
            if not q or int(q[0]) == 0:
                note.append("수량 없음")
            if not r.get("위치"):
                note.append("보존위치 없음(출입구 근처 권장)")
            ex = day(r.get("기한"))
            if ex and ex < today:
                note.append("유통기한 지남")
            elif ex and (ex - today).days <= 30:
                note.append(f"유통기한 {(ex - today).days}일 남음")
            nbad += bool(note)
            su.append([i, r.get("품명", ""), r.get("위치", ""), r.get("수량", ""), r.get("단위", ""), r.get("상태", ""), r.get("기한", ""), " · ".join(note)])
        summary.append(f"비상용품 확인할 것 {nbad}")
    else:
        for i, x in enumerate(["보호복·방진마스크·장갑·장화", "지퍼백·비닐봉지(동결 포장)", "간지·부직포 롤지·흡수지", "플라스틱 운반상자·판지상자", "제습기·선풍기",
                               "HEPA 진공청소기", "위험경고테이프", "필기구·라벨·메모용지", "휴대용 발전기(비상전원)", "구급약"], 1):
            su.append([i, x, "", "", "", "", "", ""])
        su.append(["", "목록 예시 — 실제 품목은 disaster-recovery '물품·장비(부속서 A)' 참고. 출입구 근처, 주기적으로 수량·위치·유통기한 확인(6.5.5)"])

    pr = wb.create_sheet("대응·복구 계획 항목")
    pr.append(["단계", "재난대책에 들어가야 할 항목(NAK 2-1)", "있음/없음"])
    for st_, items in (("대응(6.5.6)", ["직원 행동요령", "기록물·인명 대피경로(도면 포함), 피난 장소, 대피 절차", "재난 유형별 기록물 응급조치"]),
                       ("복구(6.5.6)", ["업무부서별 복구 활동 계획", "분야별 전문가·전문기관 연락처", "손상 기록물 (임시)작업장 확보 계획", "손상 기록물 이송 방법·시나리오"]),
                       ("교육(6.5.7)", ["비상조직 구성·역할", "조직·유형·개인별 대응 행동절차", "비상용품 용도·사용법", "비상시 집결지·구난 경로", "기록물 응급조치·복구"]),
                       ("대응 원칙(6.6.2)", ["인명 보호·구난이 먼저, 그다음 기록물 구난", "구난 장소는 피해 우려 없는 다른 서고 우선, 외부는 복수 후보지", "기록물 위치 이동·응급조치 내역을 모두 기록"])):
        for x in items:
            pr.append([st_, x, ""])
    rp = wb.create_sheet("재난 보고서(6.6.4)")
    for k in ["재난 발생 일시·기간", "위치", "재난의 특성·원인", "인명 영향", "기록물 영향(유실·손상, 재배치 뒤 정수점검 결과)", "업무 영향", "시설 영향",
              "재난 대응 경과(신고·전파·구난·응급조치·기록물 이동 이력)", "향후 대비책(재난관리계획 보완·교육)"]:
        rp.append([k, ""])

    if a.org == "영구":
        bi = wb.create_sheet("RTO·RPO(표 7)")
        bi.append(["업무 프로세스", "RTO(시간)", "RPO", "복구 수준", "DR 모델", "확인할 것"])
        if a.bia:
            for r in read(a.bia, BALIAS):
                m = re.search(r"[\d.]+", r.get("RTO", ""))
                h = float(m[0]) * (24 if re.search(r"일", r.get("RTO", "")) else 1) if m else None
                t_, mdl, note = tier(h, r.get("RPO", ""))
                bi.append([r.get("업무", ""), "" if h is None else f"{h:g}", r.get("RPO", ""), t_, mdl, note])
        else:
            bi.append(["(--bia로 업무 프로세스별 RTO·RPO를 주면 Tier·DR 모델을 붙임)"])
        bi.append([])
        bi.append(["기준", "RTO 1: 4시간 · 2: 12시간 · 3: 24시간 · 4: 7일 · 5: 7일 넘음 / RPO 1: 최종시점 · 2: D-1 → Tier 1·2 HOT, 3·4 WARM, 5 COLD, 6 없음(표 7 예시)"])
        bi.append(["DR 모델", "Mirror 즉시(Active-Active) · Hot 4시간 안(Active-Standby, 실시간 미러링) · Warm 수일~수주(중요 자원만, 수시간~1일 백업) · Cold(표 8)"])

    for sh, widths in ((sc, (18, 90, 10, 20)), (dc, (14, 100, 10, 20)), (rk, (12, 20, 30, 14, 12, 8, 30, 30, 30)), (sf, (20, 80)),
                       (mc, (24, 14, 16, 14, 24, 20, 10, 14, 30)), (su, (20, 30, 16, 8, 8, 10, 12, 30)), (pr, (16, 80, 10)), (rp, (50, 60))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{a.org} 보안·재난관리 점검표 · " + " · ".join(summary) + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
