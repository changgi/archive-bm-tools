#!/usr/bin/env python3
"""재난 피해 종이기록물 응급복구 분류: 피해 기록물 목록으로 응급복구 순서와 48시간 기한(곰팡이 발아)을 계산하고,
냉동동결·곰팡이 분리·습식 클리닝 불가 재질을 표시하며, 포장 라벨 정보와 응급복구 체크리스트를 만든다.
재질별 건조 방법(표 6·8), 부속서 D 상태 체크리스트, 부속서 E 작업공간 점검, 부속서 A 물품 목록도 만든다.
종이 외 매체(마이크로필름·영화필름·사진·자기디스크·테이프·도면·액자)는 NAK 2-1 부속서 C의 시간(48·72시간)과 조치를 쓴다.
평소(대비)에는 보유 목록으로 대피 우선순위 후보(1~3등급)를 만든다.
근거: NAK 38:2024(v1.0) 「재난 피해 종이기록물 응급복구 절차」(국가기록원, 2026-11-20 내려받아 확인), NAK 2-1:2024(v1.2) 대피 우선순위 예시

사용:  python disaster_triage.py 피해목록.xlsx --found "2026-07-15 09:00" [--now "2026-07-15 18:00"] -o 산출/응급복구.xlsx
       python disaster_triage.py 보유목록.xlsx --prepare -o 산출/대피우선순위.xlsx      (평소 대비)
피해목록 열(머리글 이름으로 찾음): 관리번호, 제목, 보존기간, 대피등급(1~3), 재질(낱장·서적·코팅·특수지·가죽 등), 제본(유·무),
       훼손유형(완전침수·부분침수·습기·화재·그을림·탄화·오염·곰팡이·파손), 위치, 발견일시, 비고(수용성 잉크 등)
규칙(NAK 38):
  6.2 응급복구 대상 구분: 훼손 유형·정도(완전 침수 > 부분 침수 > 습기 노출), 48시간 안 응급복구 가능 여부, 재질별 건조·냉동동결 가능 여부
  6.3.2 작업공간 상대습도 60% 이하, 수해 기록물은 48시간이 지나면 곰팡이 발생 가능
  6.5 포장 때 기본정보(등급·재질·훼손상태·곰팡이 유무·이력·위치)를 함께 넣고, 곰팡이 기록물은 투명 지퍼백 + 별도 상자
  6.6.2.3 48시간 안 응급복구가 어려우면 즉시 냉동동결, 곰팡이가 이미 생긴 것도 즉시 동결. 포장지에 기록물명·수습장소·날짜
  부속서 C 코팅지·가죽 장정본·수용성 잉크 기록물은 습식 클리닝 불가
  표 2 대피 우선순위: 1등급 사료적 가치 높음·대체 불가·복구 불가 / 2등급 손실 시 심각하나 비용 들여 대체 가능 / 3등급 쉽게 대체 가능
원칙:  분류·목록만 만든다. 실제 처리 방법(세척·건조·동결)은 보존처리 담당자가 기록물 상태를 보고 정한다. 화재 피해 기록물은 환기되는 곳에서.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"번호": ["관리번호", "철분류번호", "등록번호"], "제목": ["제목", "철제목"], "보존": ["보존기간"], "등급": ["대피등급", "우선순위", "등급"],
         "재질": ["재질", "종이재질"], "제본": ["제본", "제본상태"], "훼손": ["훼손유형", "피해유형", "훼손상태"], "위치": ["위치", "수습장소", "서고"],
         "발견": ["발견일시", "피해일시"], "비고": ["비고"]}
SEVERITY = [("부분", 2), ("완전", 3), ("침수", 3), ("젖", 3), ("탄", 3), ("습기", 1), ("그을", 1), ("오염", 1), ("파손", 1)]  # 앞에서 먼저 맞는 것
# NAK 2-1:2024 부속서 C 피해기록물 응급조치 요령 — 종이 외 매체(시간, 조치, 동결 가능)
MEDIA = [
    (r"마이크로필름", 72, "72시간 안 세척·건조(동결건조·공기건조) · 보관상자에 깨끗한 물을 채워 임시 보관 후 재세척·건조", False),
    (r"영화필름|필름캔", 72, "72시간 안 세척·건조 · 필름캔을 차가운 물로 채운 뒤 재세척·건조", False),
    (r"사진|프린트|네거티브|슬라이드", 72, "72시간 안 동결 또는 공기·동결건조(진공건조 금지) · 컬러사진 → 프린트 → 네거티브·슬라이드 순으로 구난 · 유제층 접촉 금지, 찬물에서 유지", True),
    (r"자기디스크|디스크|광디스크|하드", 48, "48시간 안 공기건조 또는 동결건조(-18℃ 이상) · 표면 접촉·충격 금지 · 플라스틱 상자에 세워 보관", True),
    (r"테이프|카세트|릴", 48, "공기건조 또는 동결건조(-10℃ 이상) · 릴·카세트 가장자리 충격 금지 · 플라스틱·판지 상자에 세워 보관", True),
    (r"도면", 48, "48시간 안 동결 또는 건조 · 낱장 분리 금지 · 서랍 속 그대로, 판판한 상자로 포장", True),
    (r"액자", 48, "48시간 안 동결 또는 공기건조 · 분해해 도면처럼 처리", True),
]


def media_rule(mat, note):
    m = mat + " " + note
    return next(((h, t, fz) for rx, h, t, fz in MEDIA if re.search(rx, m)), None)


KIT = [("개인보호장비", [("보호복", "침수·오염 기록물 처리"), ("방진마스크", "분진 차단"), ("안전화(고무장화 등)", "침수·오염 기록물 처리"),
                       ("장갑(고무·니트릴·라텍스)", "침수·오염 기록물 처리"), ("방수앞치마·팔토시", "침수·오염 기록물 처리"), ("위험경고테이프", "피해 수습지역 보안")]),
       ("세척·복원 소도구", [("습식용 트레이", "습식 세척"), ("밀폐형 양동이", "습식 세척"), ("플라스틱 비이커", "습식 세척"), ("부직포 롤지", "간지·수습"),
                        ("스프레이", "습식 세척"), ("붓", "세척"), ("비닐시트(방수포)", "작업대 덮개·건조"), ("핀셋", "응급복원"), ("대나무 헤라", "응급복원"),
                        ("가위", "응급복원"), ("보존용 중성테이프", "응급복원")]),
       ("건조·보관", [("건조용 흡수포", "건조"), ("깨끗한 타올", "건조·포장"), ("지퍼백", "냉동동결 포장"), ("플라스틱 운반상자", "취급·이동"), ("보관상자", "물품 보관·이동")]),
       ("피해상황 기록", [("메모용지", "피해지역 작업"), ("필기구", "분류·동결 확인")]),
       ("기타", [("HEPA 필터 진공청소기", "곰팡이·분진"), ("선풍기", "공기 순환"), ("제습기", "상대습도 60% 이하")])]
NO_WET = re.compile(r"코팅|가죽|수용성|특수지|아트지")


def drying(mat, note, dmg):
    """표 6·표 8(NAK 38) 재질별 건조 방법 — 참고값. (건조 방법 코드 a/b/c, 설명)"""
    m = mat + " " + note
    if "가죽" in m:
        return "a", "자연건조 권장(가죽 장정본은 동결건조 때 변형 가능, 표 8 비고)"
    if re.search(r"코팅|아트지|특수지", m):
        return "b", "진공동결건조가 가장 좋음 · 동결이 어려우면 간지를 끼워 분리해 흡수지 위에 눕혀 자연건조 · 진공건조 절대 금지(점착)"
    if "서적" in m:
        t = "세울 수 있으면 북엔드·플라스틱 상자로 받쳐 세워 자연건조 · 대량이면 진공동결건조(진공건조는 변형이 커 재복원 필요)"
    else:
        t = "흡수지 위에 펼쳐 자연건조, 젖은 흡수지는 자주 교체 · 좁으면 약 20장씩 겹쳐 자주 뒤집기 · 대량이면 진공동결건조"
    if "수용성" in m:
        t += " · 수용성 잉크 — 진공건조는 번짐 위험"
    return "a", t + " · 드라이기·열풍기·장시간 직사광선 금지, 건조 공간 20℃ 이하·상대습도 60% 이하"


def when(s):
    s = (s or "").strip().replace("T", " ").replace(".", "-").replace("/", "-")
    for f in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head and k not in idx:
                idx[k] = head.index(n)
    return [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]


def grade_hint(ret):
    r = re.sub(r"\s", "", ret or "")
    if r in ("영구", "준영구"):
        return 1, "영구·준영구 — 대체 불가 가능성 높음(후보)"
    if r in ("30년", "10년"):
        return 2, "중장기 보존(후보)"
    if r:
        return 3, "단기 보존(후보)"
    return "", "보존기간 없음"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--found", help="피해 발생·발견 시각(목록에 발견일시가 없을 때) 예: '2026-07-15 09:00'")
    ap.add_argument("--now", help="판정 시각(기본: 지금)")
    ap.add_argument("--prepare", action="store_true", help="평소 대비: 보유 목록으로 대피 우선순위 후보만 만듦")
    ap.add_argument("--room", help="작업공간 현재 온도,상대습도 예: '23,68' — 20℃·60% 넘으면 경고(부속서 E)")
    ap.add_argument("-o", "--out", default="산출/응급복구.xlsx")
    a = ap.parse_args()
    items = read(a.listing)
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    ws = wb.active
    if a.prepare:
        ws.title = "대피 우선순위"
        ws.append(["관리번호", "제목", "보존기간", "지정 등급", "등급 후보", "후보 근거", "확인할 것"])
        c = {1: 0, 2: 0, 3: 0}
        for n, x in items:
            g, why = grade_hint(x.get("보존"))
            given = re.search(r"[123]", x.get("등급", ""))
            if isinstance(g, int):
                c[int(given[0]) if given else g] += 1
            ws.append([x.get("번호", ""), x.get("제목", ""), x.get("보존", ""), given[0] if given else "", g, why,
                       "" if given else "담당자가 등급 지정(사료적 가치·대체 가능성·복구 가능성)"])
        g = wb.create_sheet("기준")
        for line in [["1등급", "사료적 가치가 높은 기록물, 손실되면 대체 불가능, 복구 불가능하거나 어려운 기록물"],
                     ["2등급", "손실되면 심각한 문제가 생기고 비용이 많이 들지만 대체 가능한 기록물"],
                     ["3등급", "손실되어도 문제가 없고 쉽게 대체 가능한 기록물"],
                     ["근거", "NAK 38:2024 표 2(NAK 2-1:2024 예시) — 보유 목록을 반드시 확보하고 우선 대피할 수 있게 관리(5.2)"],
                     ["주의", "보존기간으로 붙인 '등급 후보'는 출발점일 뿐 — 필수기록물(NAK 2-2)과 사료적 가치를 함께 봄"]]:
            g.append(line)
        os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
        wb.save(a.out)
        print(f"보유 {len(items)}건 · 1등급 {c[1]} · 2등급 {c[2]} · 3등급 {c[3]}(지정 또는 후보) → {a.out}")
        return 0

    now = when(a.now) if a.now else dt.datetime.now()
    found_all = when(a.found)
    ws.title = "응급복구 순서"
    ws.append(["순서", "관리번호", "제목", "대피등급", "훼손", "심각도", "48시간 기한", "남은 시간", "조치", "주의", "위치", "건조 방법(참고, 표 6·8)"])
    rows, frozen, mold, nowet = [], 0, 0, 0
    for n, x in items:
        dmg = x.get("훼손", "")
        sev = next((s for k, s in SEVERITY if k in dmg), 0)
        g = re.search(r"[123]", x.get("등급", ""))
        g = int(g[0]) if g else grade_hint(x.get("보존"))[0] or 3
        f = when(x.get("발견")) or found_all
        mr = media_rule(x.get("재질", ""), x.get("비고", ""))
        hours = mr[0] if mr else (0 if "수용성" in x.get("재질", "") + x.get("비고", "") else 48)
        deadline = f + dt.timedelta(hours=hours) if f else None
        left = (deadline - now).total_seconds() / 3600 if deadline else None
        acts, warns = [], []
        has_mold = "곰팡이" in dmg
        wet = any(k in dmg for k in ("침수", "젖", "습기", "수해"))
        if wet and mr:
            acts.append(mr[1])
            if mr[2] and left is not None and left <= 0:
                acts.append(f"즉시 냉동동결({hours}시간 지남)")
                frozen += 1
            elif not mr[2] and left is not None and left <= 0:
                acts.append(f"{hours}시간 지남 — 물속 임시 보관을 유지하고 전문기관에 바로 의뢰(동결 대상 아님)")
            if has_mold:
                acts.append("곰팡이 — 별도 상자로 분리 표시")
                mold += 1
        elif has_mold:
            acts.append("즉시 냉동동결(확산 방지) · 투명 지퍼백 + 별도 상자로 분리 표시")
            mold += 1
            frozen += 1
        elif wet and (left is None or left <= 0 or sev >= 3 and left < 12):
            acts.append(("즉시 냉동동결 또는 건조(수용성 잉크 — 즉시 조치, 선풍기 바람 자제)" if hours == 0 else "즉시 냉동동결(48시간 안 응급복구 어려움)")
                        if (left is not None and left <= 0) or sev >= 3 else "냉동동결 검토")
            frozen += "즉시" in acts[-1]
        elif wet:
            acts.append("즉시 동결 또는 건조(수용성 잉크 — 선풍기 바람은 번짐 우려, 철 사이 백지 끼워 판지상자 포장)" if hours == 0
                        else "48시간 안 현장 응급복구(세척·건조) — 못 하면 냉동동결 · 낱장으로 떼어 내지 말고 철 사이에 간지를 끼워 판지상자 포장")
            if re.search(r"코팅", x.get("재질", "")):
                acts.append("코팅지 — 즉시 포장, 젖은 채로 보관 후 동결(건조)")
        if any(k in dmg for k in ("화재", "그을", "탄")):
            acts.append("운반도구 사용(직접 취급 금지) · 환기되는 공간에서 작업")
            if not wet:
                acts.append("젖지 않은 화재 피해 — 수습 후 응급복원처리 절차(7장)로(6.6.1)")
        if any(k in dmg for k in ("오염", "토사", "그을", "재")):
            acts.append("건식 세척(붓·면봉) 먼저, 안 되면 습식(분무·침지) — 보호장비·환기(6.6.2.1)" if not NO_WET.search(x.get("재질", "") + " " + x.get("비고", "")) else "건식 세척(붓·면봉)만 — 보호장비·환기(6.6.2.1)")
        if NO_WET.search(x.get("재질", "") + " " + x.get("비고", "")):
            warns.append("습식 클리닝 불가 재질(코팅지·가죽 장정본·수용성 잉크 등)")
            nowet += 1
        if "서적" in x.get("재질", "") or x.get("제본", "") in ("유", "있음", "Y"):
            warns.append("서적류 — 책등을 아래로, 크기 순으로, 상자를 가득 채우지 말 것")
        if not f:
            warns.append("발견 시각 없음 — 48시간 기한을 계산할 수 없음")
        dry = ("a", "매체별 조치 칸 참고(NAK 2-1 부속서 C)") if mr else drying(x.get("재질", ""), x.get("비고", ""), dmg)
        rows.append((g, -sev, deadline or dt.datetime.max, n, x, sev, deadline, left, acts, warns, dry))
    rows.sort(key=lambda r: (0 if "곰팡이" in r[4].get("훼손", "") else 1, r[0], r[1], r[2]))
    for i, (g, _, _, n, x, sev, deadline, left, acts, warns, dry) in enumerate(rows, 1):
        ws.append([i, x.get("번호", ""), x.get("제목", ""), g, x.get("훼손", ""), sev, deadline.strftime("%Y-%m-%d %H:%M") if deadline else "",
                   "" if left is None else (f"{left:.0f}시간" if left > 0 else "지남"), " · ".join(acts), " · ".join(warns), x.get("위치", ""), dry[1]])
        if left is not None and left <= 12:
            ws.cell(ws.max_row, 8).fill = red
    lab = wb.create_sheet("포장 라벨")
    lab.append(["관리번호", "기록물명", "등급", "재질", "훼손상태", "곰팡이 유무", "이력(수습 일시)", "위치(수습장소)", "동결 여부"])
    for g, _, _, n, x, sev, deadline, left, acts, warns, dry in rows:
        lab.append([x.get("번호", ""), x.get("제목", ""), g, x.get("재질", ""), x.get("훼손", ""), "유" if "곰팡이" in x.get("훼손", "") else "무",
                    now.strftime("%Y-%m-%d %H:%M"), x.get("위치", ""), "동결" if any("냉동동결(" in s or "즉시 냉동동결" in s for s in acts) else ""])
    lab.append([])
    lab.append(["냉동동결 포장(6.6.2.3)", "표면 수분 제거 → 간지 삽입 → 투명 지퍼백·비닐봉지에 소분 밀봉 → 겉면에 기록물명·수습장소·날짜 → 냉동고에 겹치지 말고 눕혀서"])
    lab.append(["응급복구 후 보관(6.7)", "중성폴더로 포장 → 별도 보존상자 → 상자에 기록물명·수량, '취급주의' 표시 → 별도 관리"])
    st = wb.create_sheet("상태 체크리스트(부속서 D)")
    st.append(["관리번호", "종이재질(①낱장 ②서적 ③코팅·특수지 ④기타)", "훼손유형(A수침 B화재 C오염 D파손 E곰팡이 F기타)", "건조방법(a자연 b진공동결 c진공)",
               "냉동동결(○/×)", "제본(유/무)", "보존기간(1영구 2준영구 3한시)", "훼손등급(1양호 2보통 3심각)"])
    for g, _, _, n, x, sev, deadline, left, acts, warns, dry in rows:
        m, d = x.get("재질", ""), x.get("훼손", "")
        mat = "③" if re.search(r"코팅|특수지|아트지", m) else "②" if "서적" in m or "책" in m else "①" if re.search(r"낱장|문서|종이", m) else "④"
        kinds = "".join(c for c, rx in (("A", r"침수|젖|습기|수해"), ("B", r"화재|그을|탄"), ("C", r"오염|토사|오물"), ("D", r"파손|찢|결실"), ("E", r"곰팡이")) if re.search(rx, d)) or "F"
        ret = re.sub(r"\s", "", x.get("보존", ""))
        grade = 3 if "E" in kinds or sev >= 3 else 2 if sev == 2 else 1
        st.append([x.get("번호", ""), mat, kinds, dry[0], "○" if any("냉동동결(" in s_ for s_ in acts) else "×",
                   "유" if "서적" in m or x.get("제본", "") in ("유", "있음", "Y") else "무",
                   "1" if ret == "영구" else "2" if ret == "준영구" else "3" if ret else "", grade])
    env = wb.create_sheet("작업공간(부속서 E)")
    env.append(["점검 항목", "확인", "비고"])
    room_note = ""
    if a.room:
        try:
            rt, rh = (float(v) for v in a.room.split(","))
            bad = [w for w, c in ((f"온도 {rt:g}℃ > 20℃", rt > 20), (f"상대습도 {rh:g}% > 60%", rh > 60)) if c]
            room_note = ("부적합: " + ", ".join(bad) + " — 48시간 넘게 두면 곰팡이 생장 가능성 높음, 제습기·에어컨 24시간·선풍기") if bad else f"적합: {rt:g}℃ · {rh:g}%"
        except ValueError:
            room_note = f"읽지 못함: {a.room} (예: 23,68)"
    for line in ["온도·습도를 조절할 수 있는 깨끗하고 건조하며 안전한 공간인가", "주변에 안전띠를 둘러 작업공간임을 표시했는가",
                 "창문에 UV 차단 필름 또는 블라인드로 직사광선을 막았는가", f"온도 20℃ 이하·상대습도 60% 이하로 유지하는가{(' — ' + room_note) if room_note else ''}",
                 "항온항습이 어려우면 여름 에어컨 24시간·겨울 난방 조절이 되는가", "공조시스템(없으면 선풍기)으로 공기가 순환되는가",
                 "비상전원용 휴대용 발전기를 준비했는가", "출입을 통제하고 작업자 안전 점검을 했는가(6.3.1)", "퇴실 전 정리·청소하는가"]:
        env.append([line, "부적합" if room_note.startswith("부적합") and "20℃" in line else "", ""])
    kit = wb.create_sheet("물품·장비(부속서 A)")
    kit.append(["구분", "품목", "용도", "보유"])
    for grp, things in KIT:
        for it, use in things:
            kit.append([grp, it, use, ""])
    kit.append(["보관", "이동용 상자(응급키트)", "겉면에 내용물 목록과 '재난 피해 응급복구 물품' 표시 — 재난 복구에만 사용(5.4)", ""])
    ck = wb.create_sheet("체크리스트(부속서 C)")
    ck.append(["영역", "점검 항목", "확인", "비고"])
    for line in [["예방·대비", "재난 대비 종이기록물 우선순위를 지정해 관리하는가?"], ["예방·대비", "직원들이 응급복구 지식을 숙지하고 있는가?"],
                 ["예방·대비", "응급복구 물품·장비를 구비하고 있는가?(이동용 상자, '재난 피해 응급복구 물품' 표시)"],
                 ["대응", "응급복구 작업공간을 마련했는가?(보안, 상대습도 60% 이하, 환기)"], ["대응", "응급복구 대상 우선순위를 정했는가? 손상 유형·피해 정도를 조사했는가?"],
                 ["대응", "비축 물품 외 추가 물품·장비를 확보했는가?"], ["대응", "취급·이동용 물품(보존용지·산모아지·화학지 등)을 준비했는가?"],
                 ["복구", "피해 유형별 응급복구 방법을 확인했는가?"], ["복구", "물 사용 가능·불가능 기록물을 구분해 오염을 제거했는가?(코팅지·가죽·수용성 잉크는 습식 불가)"],
                 ["복구", "페이지 펼침 가능·불가능을 파악해 건조 방법에 따라 건조했는가?"], ["복구", "48시간 안 응급복구가 불가능한 기록물은 냉동동결했는가?"],
                 ["복구", "냉동동결 포장지에 기록물명·수습장소·날짜를 적었는가?"], ["복구", "곰팡이가 생긴 기록물을 별도 공간으로 분리했는가?"],
                 ["보관", "응급복구가 끝난 기록물을 중성폴더·별도 보존상자에 넣고 '취급주의'를 표시했는가?(6.7)"],
                 ["사후관리", "반출 기록물은 현장이 안정될 때까지 재배치하지 않았는가? 재배치 뒤 정수점검을 했는가?(NAK 2-1 6.6.4)"],
                 ["사후관리", "기록물 위치 이동·응급조치 내역을 모두 기록했는가? 재난 보고서를 썼는가?(NAK 2-1 6.6.2·6.6.4)"],
                 ["보안", "손상된 비밀기록물은 일반 기록물과 분리하고, 이송 때 비밀취급 인가자가 동행했는가?(NAK 2-1 6.6.3)"]]:
        ck.append([*line, "", ""])
    for sh, widths in ((ws, (6, 16, 34, 8, 18, 8, 17, 10, 50, 50, 14, 70)), (st, (16, 22, 22, 16, 10, 8, 14, 14)), (env, (90, 10, 30)), (kit, (16, 30, 50, 8)), (lab, (16, 34, 6, 12, 18, 10, 17, 14, 8)), (ck, (12, 80, 8, 30))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"피해 {len(items)}건 · 즉시 냉동동결 {frozen} · 곰팡이 분리 {mold} · 습식 클리닝 불가 {nowet} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
