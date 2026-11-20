#!/usr/bin/env python3
"""재난 피해 종이기록물 응급복구 분류: 피해 기록물 목록으로 응급복구 순서와 48시간 기한(곰팡이 발아)을 계산하고,
냉동동결·곰팡이 분리·습식 클리닝 불가 재질을 표시하며, 포장 라벨 정보와 응급복구 체크리스트를 만든다.
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
NO_WET = re.compile(r"코팅|가죽|수용성|특수지|아트지")


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
    ws.append(["순서", "관리번호", "제목", "대피등급", "훼손", "심각도", "48시간 기한", "남은 시간", "조치", "주의", "위치"])
    rows, frozen, mold, nowet = [], 0, 0, 0
    for n, x in items:
        dmg = x.get("훼손", "")
        sev = next((s for k, s in SEVERITY if k in dmg), 0)
        g = re.search(r"[123]", x.get("등급", ""))
        g = int(g[0]) if g else grade_hint(x.get("보존"))[0] or 3
        f = when(x.get("발견")) or found_all
        deadline = f + dt.timedelta(hours=48) if f else None
        left = (deadline - now).total_seconds() / 3600 if deadline else None
        acts, warns = [], []
        has_mold = "곰팡이" in dmg
        wet = any(k in dmg for k in ("침수", "젖", "습기", "수해"))
        if has_mold:
            acts.append("즉시 냉동동결(확산 방지) · 투명 지퍼백 + 별도 상자로 분리 표시")
            mold += 1
            frozen += 1
        elif wet and (left is None or left <= 0 or sev >= 3 and left < 12):
            acts.append("즉시 냉동동결(48시간 안 응급복구 어려움)" if (left is not None and left <= 0) or sev >= 3 else "냉동동결 검토")
            frozen += "즉시" in acts[-1]
        elif wet:
            acts.append("48시간 안 현장 응급복구(세척·건조) — 못 하면 냉동동결")
        if any(k in dmg for k in ("화재", "그을", "탄")):
            acts.append("운반도구 사용(직접 취급 금지) · 환기되는 공간에서 작업")
        if NO_WET.search(x.get("재질", "") + " " + x.get("비고", "")):
            warns.append("습식 클리닝 불가 재질(코팅지·가죽 장정본·수용성 잉크 등)")
            nowet += 1
        if "서적" in x.get("재질", "") or x.get("제본", "") in ("유", "있음", "Y"):
            warns.append("서적류 — 책등을 아래로, 크기 순으로, 상자를 가득 채우지 말 것")
        if not f:
            warns.append("발견 시각 없음 — 48시간 기한을 계산할 수 없음")
        rows.append((g, -sev, deadline or dt.datetime.max, n, x, sev, deadline, left, acts, warns))
    rows.sort(key=lambda r: (0 if "곰팡이" in r[4].get("훼손", "") else 1, r[0], r[1], r[2]))
    for i, (g, _, _, n, x, sev, deadline, left, acts, warns) in enumerate(rows, 1):
        ws.append([i, x.get("번호", ""), x.get("제목", ""), g, x.get("훼손", ""), sev, deadline.strftime("%Y-%m-%d %H:%M") if deadline else "",
                   "" if left is None else (f"{left:.0f}시간" if left > 0 else "지남"), " · ".join(acts), " · ".join(warns), x.get("위치", "")])
        if left is not None and left <= 12:
            ws.cell(ws.max_row, 8).fill = red
    lab = wb.create_sheet("포장 라벨")
    lab.append(["관리번호", "기록물명", "등급", "재질", "훼손상태", "곰팡이 유무", "이력(수습 일시)", "위치(수습장소)", "동결 여부"])
    for g, _, _, n, x, sev, deadline, left, acts, warns in rows:
        lab.append([x.get("번호", ""), x.get("제목", ""), g, x.get("재질", ""), x.get("훼손", ""), "유" if "곰팡이" in x.get("훼손", "") else "무",
                    now.strftime("%Y-%m-%d %H:%M"), x.get("위치", ""), "동결" if any("냉동동결(" in s or "즉시 냉동동결" in s for s in acts) else ""])
    ck = wb.create_sheet("체크리스트(부속서 C)")
    ck.append(["영역", "점검 항목", "확인", "비고"])
    for line in [["예방·대비", "재난 대비 종이기록물 우선순위를 지정해 관리하는가?"], ["예방·대비", "직원들이 응급복구 지식을 숙지하고 있는가?"],
                 ["예방·대비", "응급복구 물품·장비를 구비하고 있는가?(이동용 상자, '재난 피해 응급복구 물품' 표시)"],
                 ["대응", "응급복구 작업공간을 마련했는가?(보안, 상대습도 60% 이하, 환기)"], ["대응", "응급복구 대상 우선순위를 정했는가? 손상 유형·피해 정도를 조사했는가?"],
                 ["대응", "비축 물품 외 추가 물품·장비를 확보했는가?"], ["대응", "취급·이동용 물품(보존용지·산모아지·화학지 등)을 준비했는가?"],
                 ["복구", "피해 유형별 응급복구 방법을 확인했는가?"], ["복구", "물 사용 가능·불가능 기록물을 구분해 오염을 제거했는가?(코팅지·가죽·수용성 잉크는 습식 불가)"],
                 ["복구", "페이지 펼침 가능·불가능을 파악해 건조 방법에 따라 건조했는가?"], ["복구", "48시간 안 응급복구가 불가능한 기록물은 냉동동결했는가?"],
                 ["복구", "냉동동결 포장지에 기록물명·수습장소·날짜를 적었는가?"], ["복구", "곰팡이가 생긴 기록물을 별도 공간으로 분리했는가?"]]:
        ck.append([*line, "", ""])
    for sh, widths in ((ws, (6, 16, 34, 8, 18, 8, 17, 10, 50, 50, 14)), (lab, (16, 34, 6, 12, 18, 10, 17, 14, 8)), (ck, (12, 80, 8, 30))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"피해 {len(items)}건 · 즉시 냉동동결 {frozen} · 곰팡이 분리 {mold} · 습식 클리닝 불가 {nowet} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
