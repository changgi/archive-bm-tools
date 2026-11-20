#!/usr/bin/env python3
"""기관 통폐합 때 두 기관의 기록관리표(단위과제 목록)를 맞대어 통합 기록관리표 초안을 만든다.

  merge_schedule.py 기관A_단위과제.xlsx 기관B_단위과제.xlsx [기관C_단위과제.xlsx …] [--names 가청,나청,다청] [--type 중앙행정기관] -o 산출/통합_기록관리표.xlsx

입력 열(머리글 이름으로 찾음, 없는 열은 비워 둠): 부서(부서명) · 대기능 · 중기능 · 소기능 · 기능유형 · 단위과제(명) · 단위과제ID ·
  단위과제설명 · 보존기간 · 보존기간책정사유 · 단위과제카드 · 업무담당자 · 보존기간확정여부 · 부서폐지여부

결과 시트
  사전 정비 대상   단위과제카드 없음·업무담당자 누락·보존기간 미확정·같은 기관 안 중복·폐지 부서 → '종료 대상' 또는 '확정 먼저'
  단위과제비교표   두 기관의 비슷한 단위과제를 한 줄에 놓고 사무분장(부서)·설명·보존기간·책정사유·공통업무 준칙을 나란히,
                   문제유형(②중복 ③명칭 ⑤기능유형 ⑥보존기간 오책정 ⑦보존기간 비대칭)과 개선안, 통합 후보를 적음
  통합 기록관리표(초안)  비교표에서 고른 통합 후보 + 한 기관에만 있는 과제(현행 유지 후보)
  부서 점검표      부서별 교차 검증 체크리스트(업무 누락·중복, 명칭·분류, 설명·보존기간)
  절차·근거

근거: 공공기록물법 제25조(폐지기관의 기록물관리)·시행령 제58조, NAK 4:2025(v2.3) 기록관리표 작성 및 관리 절차(보존기간 협의),
      국가기록원 연구보고서 「기록물 평가·분류체계 개편 및 기관 통폐합 관련 분류체계 정비 방안 연구」(2026) 6.1
      '기관 통폐합에 따른 기록관리표 정비 가이드라인'(사전 준비 → 분류체계 정비 → 후속 조치, 표 86~90)
원칙: 이름 유사도로 후보를 짝지을 뿐이다. 통합 여부·보존기간은 추진단(기록관·조직관리부서·처리과)이 정하고, 신설·변경 보존기간은
      관할 영구기록물관리기관과 협의한다. 시스템(BRM·온나라·RMS) 반영과 고시는 하지 않는다.
"""
import argparse
import csv
import difflib
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
COMMON = os.path.join(HERE, "..", "..", "retention-review", "assets", "common", "common_task_criteria.csv")
RANK = {"영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "2년": 2.5, "1년": 2}
ALIAS = {"부서": ["부서", "부서명", "처리과", "처리과명"], "대기능": ["대기능"], "중기능": ["중기능"], "소기능": ["소기능"],
         "기능유형": ["기능유형"], "과제": ["단위과제", "단위과제명"], "ID": ["단위과제ID", "단위과제코드", "단위과제번호"],
         "설명": ["단위과제설명", "단위과제 설명", "업무설명"], "보존": ["보존기간"], "사유": ["보존기간책정사유", "책정사유"],
         "카드": ["단위과제카드", "단위과제카드명", "과제카드"], "담당": ["업무담당자", "담당자"],
         "확정": ["보존기간확정여부", "확정여부"], "폐지": ["부서폐지여부", "폐지여부"]}


def key(s):
    s = re.sub(r"^\((?:기관|처리과)공통\)\s*", "", s or "")
    return re.sub(r"[\s·ㆍ,()（）\-/]|및|관리$|업무$|운영$", "", s)


def norm_head(h):
    return re.sub(r"\s", "", str(h or ""))


def read(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(path, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    names = {norm_head(n) for v in ALIAS.values() for n in v}
    hi = max(range(min(10, len(rows))), key=lambda i: sum(norm_head(c) in names for c in rows[i]))
    head = [norm_head(h) for h in rows[hi]]
    idx = {k: next((head.index(norm_head(n)) for n in al if norm_head(n) in head), None) for k, al in ALIAS.items()}
    out = []
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        g = {k: (r[i].strip() if i is not None and i < len(r) else "") for k, i in idx.items()}
        if g["과제"]:
            g["보존"] = re.sub(r"\s", "", g["보존"])
            g["행"] = n
            out.append(g)
    return out, idx


def precheck(name, rows, idx):
    seen, flags = {}, []
    for g in rows:
        seen.setdefault((g["부서"], key(g["과제"])), []).append(g)
    for g in rows:
        why = []
        if idx["카드"] is not None and not g["카드"]:
            why.append("단위과제카드 없음 → 종료 대상 검토")
        if idx["담당"] is not None and not g["담당"]:
            why.append("업무담당자 누락")
        if g["확정"] and g["확정"] not in ("Y", "예", "확정", "O"):
            why.append("보존기간 미확정 → 영구기록물관리기관 협의로 먼저 확정")
        if not g["보존"]:
            why.append("보존기간 빈칸")
        if len(seen[(g["부서"], key(g["과제"]))]) > 1:
            why.append("같은 부서 안 중복 단위과제(②) → 하나만 남기고 종료")
        if g["폐지"] in ("Y", "예", "폐지", "O"):
            why.append("폐지 부서 — 업무 승계 부서 확인")
        if why:
            flags.append([name, g["행"], g["부서"], g["과제"], g["보존"], " · ".join(why)])
    return flags


def main() -> int:
    ap = argparse.ArgumentParser(description="기관 통폐합 단위과제비교표·통합 기록관리표 초안")
    ap.add_argument("files", nargs="+", help="기관별 단위과제 목록 2개 이상(행정통합 등 3개 이상도 가능)")
    ap.add_argument("--names", default="", help="기관 이름(쉼표, 파일 순서대로)")
    ap.add_argument("--type", default="", help="공통업무 준칙 기관유형(중앙행정기관·지방자치단체·교육(행정)기관·공공기관·학교·대학)")
    ap.add_argument("--threshold", type=float, default=0.8)
    ap.add_argument("-o", "--out", default="산출/통합_기록관리표.xlsx")
    a = ap.parse_args()
    if len(a.files) < 2:
        print("입력 오류: 기관 단위과제 목록을 2개 이상 주세요.", file=sys.stderr)
        return 2
    names = [x.strip() for x in a.names.split(",")] if a.names else []
    names = (names + [f"인계기관{chr(65 + k)}" for k in range(len(a.files))])[:len(a.files)]
    data = [read(f) for f in a.files]
    common = []
    if a.type and os.path.exists(COMMON):
        with open(COMMON, encoding="utf-8-sig") as f:
            common = [c for c in csv.DictReader(f) if c["기관유형"] == a.type]
        if not common:
            print(f"입력 오류: --type '{a.type}'에 맞는 공통업무 준칙이 없습니다.", file=sys.stderr)
            return 2
    ckey = {key(c["단위과제명"]): c for c in common}

    def rule_of(name):
        k = key(name)
        if k in ckey:
            return ckey[k]
        best = max(common, key=lambda c: difflib.SequenceMatcher(None, k, key(c["단위과제명"])).ratio(), default=None)
        if best and difflib.SequenceMatcher(None, k, key(best["단위과제명"])).ratio() >= 0.85:
            return best
        return None

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = openpyxl.Workbook()
    pre = wb.active
    pre.title = "사전 정비 대상"
    pre.append(["기관", "원본 행", "부서", "단위과제", "보존기간", "정비할 것"])
    for nm, (rows_, idx_) in zip(names, data):
        for r in precheck(nm, rows_, idx_):
            pre.append(r)

    # 첫 기관 목록을 기준으로 묶음(group)을 만들고, 다음 기관 과제를 가장 비슷한 묶음에 붙인다(같은 기관 안 중복은 건너뜀)
    groups = []  # {기관명: 행}
    for nm, (rows_, _) in zip(names, data):
        seen_k = set()
        for g in rows_:
            k = key(g["과제"])
            if k in seen_k:
                continue  # 같은 기관 안 중복은 사전 정비 대상에서 종료
            seen_k.add(k)
            best, score = None, 0.0
            for G in groups:
                if nm in G["m"]:
                    continue
                sc = max(1.0 if key(x["과제"]) == k else difflib.SequenceMatcher(None, k, key(x["과제"])).ratio() for x in G["m"].values())
                if sc > score:
                    best, score = G, sc
            if best and score >= a.threshold:
                best["m"][nm] = g
                best["s"] = min(best["s"], score)
            else:
                groups.append({"m": {nm: g}, "s": 1.0})

    cmp_ = wb.create_sheet("단위과제비교표")
    cmp_.append([f"{nm} {c}" for nm in names for c in ("부서", "단위과제", "설명", "보존기간", "책정사유")]
                + ["공통업무 준칙", "준칙 보존기간", "유사도", "문제유형", "개선안", "통합 후보 단위과제", "통합 후보 보존기간"])
    merged, flagged, n_pair, n_only = [], 0, 0, 0
    for G in groups:
        mem = G["m"]
        cells = []
        for nm in names:
            g = mem.get(nm)
            cells += [g["부서"], g["과제"], g["설명"][:200], g["보존"], g["사유"][:120]] if g else ["", "", "", "", ""]
        xs = list(mem.values())
        rule = next((r for r in (rule_of(x["과제"]) for x in xs) if r), None)
        if len(mem) == 1:
            n_only += 1
            org, g = next(iter(mem.items()))
            cmp_.append(cells + [(rule or {}).get("단위과제명", ""), (rule or {}).get("보존기간", ""), "", f"{org}에만 있음(단독 업무)",
                                 "단독 업무 수행부서는 현행 유지 · 폐지 부서면 제외 · 업무가 통합 뒤 사라지면 종료", g["과제"], g["보존"]])
            merged.append([g["부서"], g["과제"], g["보존"], g["설명"], f"{org} 단독(현행 유지 후보)"])
            continue
        n_pair += 1
        probs, fixes = [], []
        if len({x["과제"] for x in xs}) > 1:
            probs.append("③ 명칭 다름")
            fixes.append("③ 단위과제 명칭 변경(통일)")
        ft = [x["기능유형"] for x in xs if x["기능유형"]]
        if len(set(ft)) > 1:
            probs.append(f"⑤ 기능유형 다름({'/'.join(ft)})")
            fixes.append("⑤ 기능유형 변경")
        ps = [x["보존"] for x in xs if x["보존"]]
        if len(set(ps)) > 1:
            probs.append(f"⑦ 보존기간 비대칭({'/'.join(ps)})")
            fixes.append("⑦ 유사업무 보존기간 일원화")
        name = max((x["과제"] for x in xs), key=len)
        per = max((p for p in ps if p in RANK), key=RANK.get, default=ps[0] if ps else "")
        if rule:
            name, rp = rule["단위과제명"], rule["보존기간"]
            low = [f"{o}({x['보존']})" for o, x in mem.items() if x["보존"] in RANK and rp in RANK and RANK[x["보존"]] < RANK[rp]]
            if low:
                probs.append(f"⑥ {'·'.join(low)}이 준칙({rp})보다 짧음")
                fixes.append("⑥-2 책정기준 반영해 보존기간 변경")
            per = rp or per
            fixes.append("공통업무는 준칙의 명칭·설명·보존기간·책정사유 준용")
        if probs:
            flagged += 1
        merged.append([" / ".join(x["부서"] for x in xs), name, per, (rule or {}).get("업무설명", "")[:300] or next((x["설명"] for x in xs if x["설명"]), ""),
                       f"{len(mem)}개 기관 통합({'·'.join(mem)})"])
        cmp_.append(cells + [(rule or {}).get("단위과제명", ""), (rule or {}).get("보존기간", ""), f"{G['s']:.2f}",
                             " · ".join(probs) or "같은 과제", " · ".join(fixes) or "현행 유지", name, per])
    tb = wb.create_sheet("통합 기록관리표(초안)")
    tb.append(["부서(통합 전)", "단위과제", "보존기간", "단위과제 설명", "출처", "부서명(통합 후)", "정책분야", "정책영역", "대기능", "중기능", "소기능", "기능유형", "보존기간 책정사유"])
    for m in merged:
        tb.append(m)
    ck = wb.create_sheet("부서 점검표")
    ck.append(["구분", "점검항목", "확인(O/X)", "의견"])
    for row in (["업무 누락 및 중복", "우리 부서의 사무분장에 명시된 업무가 누락 없이 기록관리표에 모두 반영되어 있습니까?"],
                ["업무 누락 및 중복", "통폐합으로 다른 부서로 기능이 넘어간 업무가 여전히 우리 부서 과제로 배정되어 있지 않습니까?"],
                ["업무 누락 및 중복", "명칭만 다르고 실제 수행하는 내용이 같은 중복 단위과제는 없습니까?"],
                ["명칭 및 분류", "단위과제 명칭만으로도 어떤 업무인지 바로 알 수 있습니까?"],
                ["명칭 및 분류", "분류된 대·중·소기능이 실제 우리 부서의 설립 목적·성격과 맞습니까?"],
                ["명칭 및 분류", "부서의 고유 핵심업무가 처리과공통 과제(서무·복무 등)로 잘못 분류되어 있지 않습니까?"],
                ["설명 및 보존기간", "단위과제 설명이 업무 대상과 행위 범위(계획·회의·결과보고 등)를 정확히 담고 있습니까?"],
                ["설명 및 보존기간", "핵심 고유업무인데 보존기간이 단기(3년·5년 등)로 과소 책정된 과제는 없습니까?"]):
        ck.append(row + ["", ""])
    st = wb.create_sheet("절차·근거")
    for row in (["1. 사전 준비", "추진단 구성(기록관: 지침 이행·국가기록원 협의 / 조직관리부서(BRM): 조직개편안·사무분장 공유 / 처리과: 정비안 검토), 업무기능 승계 내역 파악, 부서별 단위과제 목록 추출, '사전 정비 대상' 시트의 미사용·중복 과제 종료와 보존기간 미확정 과제 확정, 정비 기간 중 단위과제 신설 자제"],
                ["2. 분류체계 정비", "통폐합 전·후 조직 분석(조직명이 아니라 기능 중심) → 유사 업무 수행부서 매칭(유사: 통폐합 / 단독: 현행 유지 / 신설: 사무분장으로 신설 / 폐지: 제외) → 단위과제비교표 → 통합 기록관리표 → 부서 교차 검증(부서 점검표) → 이견은 기록물관리 전문요원이 조정 → 심의"],
                ["3. 후속 조치", "BRM 단위과제 등록(대·중·소기능 매핑)과 온나라 연계 확인, 신설·변경 단위과제 보존기간을 관할 영구기록물관리기관과 협의(NAK 4: 기록관 검토요청 → 영구기관 회신 → 반영), 고시, 모니터링"],
                ["단위과제 설명 작성", "[업무 대상 및 목적]을(를) 위한 [구체적 업무 행위·범위·절차 수립·회의 개최·결과 보고 등]에 관한 업무"],
                ["문제유형(표 89)", "① 업무 누락(신설) ② 중복 운영(종료) ③ 명칭 다름(변경) ④ 크기 과대·과소(분리·통폐합) ⑤ 기능유형 오분류 ⑥ 보존기간 오책정(30년 이상 상향·기준 반영) ⑦ 보존기간 비대칭(일원화)"],
                ["근거", "공공기록물법 제25조·시행령 제58조, NAK 4:2025(v2.3), 국가기록원 「기록물 평가·분류체계 개편 및 기관 통폐합 관련 분류체계 정비 방안 연구」(2026) 6.1"],
                ["하지 않는 것", "시스템 등록·고시·공문 발송, 보존기간 확정(협의 전 후보일 뿐)"]):
        st.append(row)
    hdr = PatternFill("solid", fgColor="DDEBF7")
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font, c.fill = Font(bold=True), hdr
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width = 18
    for c in cmp_["N"][1:]:
        if c.value and c.value != "같은 과제" and "단독" not in str(c.value):
            c.fill = PatternFill("solid", fgColor="FFF2CC")
    st.column_dimensions["B"].width = 120
    for r in st.iter_rows():
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(" · ".join(f"{nm} {len(d[0])}과제" for nm, d in zip(names, data)) + f" · 묶은 과제 {n_pair} · 문제유형 있는 묶음 {flagged} · 단독 {n_only} · 사전 정비 {pre.max_row - 1} → {a.out}")
    return 1 if flagged or pre.max_row > 1 else 0


if __name__ == "__main__":
    sys.exit(main())
