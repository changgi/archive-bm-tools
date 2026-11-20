#!/usr/bin/env python3
"""필수기록물 선별·보호: 단위과제(또는 단위업무·기록물철)의 필수성 평가표를 만들고 판정하며,
보호기간과 보존기간을 대조하고, 필수기록물 관리목록표와 이중·분산 보호 점검을 만든다.
근거: NAK 2-2:2025(v1.2) 「재난시 기능연속성 유지를 위한 필수기록물 선별 및 보호절차」(국가기록원, 2026-11-20 내려받아 확인)

사용:  python vital_records.py 단위과제표.xlsx [--files 기록물철목록.xlsx] [--year 2026] -o 산출/필수기록물.xlsx
단위과제표 열(머리글 이름으로 찾음): 분류체계(기능분류), 단위과제명(또는 단위업무명·기록물철명), 업무설명, 보존기간,
       기능성, 고유성, 증거성(0·1·3·5 또는 없음·낮음·중간·높음), 보호기간, 판단사유
기록물철목록 열(있으면 관리목록표): 단위과제명, 철명(제목), 관리번호, 생산부서, 생산년도, 생산매체, 보존기간, 보존위치,
       이중매체, 이중매체관리번호, 보호위치, 보호책임, 보호기간(만료일), 분산장소(;로 여럿), 원격거리km
판정(표 1·표 2): 한 기준에서라도 '높음(5)'이거나, 두 항목 이상 '중간(3)' 이상이면 필수성 인정
  기능성 5 = 즉각 이용 또는 기능 상실 후 1일 안 복구에 이용 / 3 = 7일 안 / 1 = 복구에 활용 / 0 = 보조·참고
  고유성 5 = 법·규정상 다른 활동으로 대체 불가 / 3 = 공동 수행해야 목적 달성 / 1 = 내부 지원 / 0 = 외부 지원 등
  증거성 5 = 권리·의무를 유일하게 신설·변경·해지 / 3 = 대조·확인 / 1 = 단순 열람·발급 / 0 = 보조
보호기간(5.5): 보존기간 이하. 기능성 높음 → 정보 갱신 주기, 고유성 높음 → 업무 반복 단위, 증거성 높음 → 효력 존속기간, 그중 가장 긴 기간
일정(5.4.3): 기록관은 매년 10월 31일까지 필수성·보호기간을 검토·수정해 영구기록물관리기관과 협의·확정, 11월 1일 이후 신설·변경은 즉시 협의
보호(6장): 이중보호(복본) + 분산보호 — 분산 장소 두 곳 이상, 원격분산은 원본에서 일반적으로 10km 이상, 보호계획은 최소 1년에 한 번 검토,
  수록 매체는 주기적으로 점검. 비상대응 기록물은 현지분산(접근성), 권한보호 기록물은 원격분산
재난 뒤:  --damage 피해목록 → 피해 평가표(표 16 5단계·표 17), 업무복구 목록(필수·피해 3 이상, 기능성 높은 업무부터), 이용 대장(7.2 승인·반납·통보)
원칙:  평가 점수는 처리과가 매기고 기록관이 검토한다. 이 스크립트는 점수가 비면 낱말로 '유형 후보'만 붙이고 판정하지 않는다.
"""
import argparse
import datetime as dt
import os
import re
import sys

ALIAS = {"분류": ["분류체계", "기능분류", "분류"], "과제": ["단위과제명", "단위과제", "단위업무명", "단위업무", "기록물철명"],
         "설명": ["업무설명", "해당업무설명", "단위과제설명"], "보존": ["보존기간"], "기능성": ["기능성"], "고유성": ["고유성"], "증거성": ["증거성"],
         "보호": ["보호기간"], "사유": ["판단사유", "필수성판단사유"]}
FALIAS = {"과제": ["단위과제명", "단위과제"], "철": ["철명", "철제목", "건명", "제목"], "번호": ["관리번호", "철분류번호", "등록번호"],
          "부서": ["생산부서", "처리과"], "연도": ["생산년도", "생산연도"], "매체": ["생산매체", "기록물형태"], "보존": ["보존기간"],
          "위치": ["보존위치"], "책임": ["보존책임"], "이중": ["이중매체"], "이중번호": ["이중매체관리번호"], "보호위치": ["보호위치"],
          "보호책임": ["보호책임"], "보호": ["보호기간", "보호기간(만료일)", "보호만료일"], "분산": ["분산장소", "분산보호장소"], "거리": ["원격거리km", "원격거리"]}
WORD = {"높음": 5, "중간": 3, "낮음": 1, "없음": 0}
RET = {"영구": 999, "준영구": 998, "30년": 30, "10년": 10, "5년": 5, "3년": 3, "1년": 1}
# 4.2 유형 예시 — 점수가 비었을 때 '유형 후보' 표시용(판정에 쓰지 않음)
TYPES = [("비상대응", r"비상|재난|안전관리|민방위|연락망|시설도면|도면|보안|대피"),
         ("업무재개", r"표준운영|매뉴얼|업무계획|직제|조직|직무|법령|규정|시스템\s*운영|계약.*진행"),
         ("조직 권한보호", r"계약|협약|협정|자산|재산|회계|결산|예산|인사|급여|특허|상표|소송"),
         ("이해관계자 권한보호", r"주민등록|등록원부|등기|자동차등록|연금|사회보장|판결|병적|토지|지적|가족관계|인허가|허가|면허")]


def score(v):
    s = re.sub(r"\s", "", str(v or ""))
    if not s:
        return None
    if s in WORD:
        return WORD[s]
    m = re.search(r"\d", s)
    if m and int(m[0]) in (0, 1, 3, 5):
        return int(m[0])
    return "?"


def years(v):
    s = re.sub(r"\s", "", str(v or ""))
    if s in RET:
        return RET[s]
    m = re.search(r"(\d+)\s*년", s)
    return int(m[1]) if m else None


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
    return idx, [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("tasks")
    ap.add_argument("--files", help="기록물철 목록 — 필수 단위과제의 철로 관리목록표를 만듦")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("--today", help="일정 판정일(기본: 오늘)")
    ap.add_argument("--damage", help="재난 뒤 피해 조사 목록(관리번호, 피해정도 1~5 또는 경미·낮음·중간·높음·심각, 피해내용) — 피해 평가표(표 17)와 업무복구 목록(7.4)")
    ap.add_argument("--set", action="append", default=[], help="피해 평가표 머리 칸: 발생일=… 피해장소=… 피해원인=… 피해규모=… 조사일=… 조사자=… 확인자=…")
    ap.add_argument("-o", "--out", default="산출/필수기록물.xlsx")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    idx, tasks = read(a.tasks, ALIAS)
    if "과제" not in idx:
        print("단위과제명(또는 단위업무명·기록물철명) 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    hl = PatternFill("solid", fgColor="E2EFDA")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "필수성 평가표"
    ws.append(["표 1 단위과제 필수성 평가표(NAK 2-2:2025 5.3)"])
    ws["A1"].font = Font(bold=True)
    ws.append(["분류체계", "단위과제명", "해당업무 설명", "기능성", "고유성", "증거성", "종합판단", "필수성 판단사유", "보존기간", "보호기간", "확인할 것", "유형 후보(참고)"])
    vital, issues_n, todo, fscore, allfiles = [], 0, 0, {}, {}
    for n, x in tasks:
        sc = {k: score(x.get(k)) for k in ("기능성", "고유성", "증거성")}
        notes = []
        if any(v == "?" for v in sc.values()):
            notes.append("점수는 0·1·3·5(없음·낮음·중간·높음) 중 하나")
        vals = [v for v in sc.values() if isinstance(v, int)]
        if len(vals) < 3:
            verdict = "평가 필요"
            todo += 1
        else:
            ok = any(v >= 5 for v in vals) or sum(v >= 3 for v in vals) >= 2
            verdict = "필수" if ok else "해당 없음"
        ry, py = years(x.get("보존")), years(x.get("보호"))
        if verdict == "필수":
            if not x.get("보호"):
                notes.append("보호기간 없음 — 기능성 높음은 갱신 주기, 고유성 높음은 업무 반복 단위, 증거성 높음은 효력 존속기간 중 가장 긴 기간(5.5)")
            elif ry and py and ry < 900 and py > ry:
                notes.append(f"보호기간 {x.get('보호')}이 보존기간 {x.get('보존')}보다 김 — 보존기간 이하(5.5)")
            if not x.get("사유"):
                notes.append("필수성 판단사유 없음")
        cand = [t for t, rx in TYPES if re.search(rx, (x.get("과제", "") + " " + x.get("설명", "")))]
        issues_n += bool(notes)
        ws.append([x.get("분류", ""), x.get("과제", ""), x.get("설명", ""), *(sc[k] if sc[k] is not None else "" for k in ("기능성", "고유성", "증거성")),
                   verdict, x.get("사유", ""), x.get("보존", ""), x.get("보호", ""), " · ".join(notes), "·".join(cand)])
        if verdict == "필수":
            ws.cell(ws.max_row, 7).fill = hl
            vital.append(x.get("과제", ""))
            fscore[x.get("과제", "")] = sc["기능성"] if isinstance(sc["기능성"], int) else 0
    lst = wb.create_sheet("관리목록표")
    lst.append(["표 4 필수기록물 관리목록표(NAK 2-2:2025 5.6) — 기본 항목 | 보호 항목"])
    lst.append(["단위과제명", "철명(건명)", "관리번호", "생산부서", "생산년도", "생산매체", "보존기간", "보존위치", "보존책임",
                "이중매체", "이중매체 관리번호", "보호위치", "보호책임", "보호기간(만료일)", "확인할 것"])
    prot = wb.create_sheet("보호 점검")
    prot.append(["관리번호", "철명", "이중보호(복본)", "분산 장소 수", "원격분산(10km 이상)", "보호기간 만료", "확인할 것"])
    nfiles = 0
    if a.files:
        fidx, files = read(a.files, FALIAS)
        vset = set(vital)
        for n, f in files:
            if f.get("번호"):
                allfiles[f["번호"]] = f
            if f.get("과제") not in vset:
                continue
            nfiles += 1
            lst.append([f.get(k, "") for k in ("과제", "철", "번호", "부서", "연도", "매체", "보존", "위치", "책임", "이중", "이중번호", "보호위치", "보호책임", "보호")] + [""])
            pn = []
            dual = bool(f.get("이중"))
            places = [p for p in re.split(r"[;,/]", f.get("분산", "")) if p.strip()]
            km = re.search(r"[\d.]+", f.get("거리", ""))
            km = float(km[0]) if km else None
            if not dual:
                pn.append("복본(이중매체) 없음 — 형태별 이중화 방법(표 12: 마이크로필름·스캐닝 후 광매체·보존스토리지 등)")
            if len(places) < 2:
                pn.append("분산 장소 두 곳 이상이 바람직(6.4.3)")
            if km is not None and km < 10:
                pn.append(f"원격분산 {km:g}km — 일반적으로 10km 이상(같은 재해를 받지 않는 거리)")
            exp = None
            m = re.search(r"(\d{4})[-.](\d{1,2})[-.](\d{1,2})", f.get("보호", ""))
            if m:
                exp = dt.date(int(m[1]), int(m[2]), int(m[3]))
                if exp < today:
                    pn.append("보호기간 만료 — 복본은 이중·보존 매체로 활용하거나 보존기간에 따라 관리(6.4.3)")
            prot.append([f.get("번호", ""), f.get("철", ""), "있음" if dual else "없음", len(places), "" if km is None else ("예" if km >= 10 else "아니오"),
                         exp.isoformat() if exp else "", " · ".join(pn)])
            lst.cell(lst.max_row, 15).value = " · ".join(pn)
    if not nfiles:
        lst.append(["(--files로 기록물철 목록을 주면 필수 단위과제의 철을 채움)"])
    ndmg = nrec = 0
    if a.damage:
        DIDX = {"번호": ["관리번호", "철분류번호", "등록번호"], "정도": ["피해정도", "피해등급"], "내용": ["피해내용", "피해상태", "비고"],
                "과제": ["단위과제명", "단위과제"], "철": ["철명", "기록물철", "제목"], "매체": ["원본매체", "생산매체"]}
        _, dmg = read(a.damage, DIDX)
        head = dict(kv.split("=", 1) for kv in a.set if "=" in kv)
        dv = wb.create_sheet("피해 평가표")
        dv.append(["기록물 피해 평가표(NAK 2-2:2025 표 17)"])
        dv["A1"].font = Font(bold=True)
        keys = ["발생일", "피해장소", "피해원인", "피해규모", "조사일", "조사자", "확인자"]
        dv.append(keys)
        dv.append([head.get(k, "") for k in keys])
        dv.append([])
        dv.append(["단위과제", "기록물철", "관리번호", "원본매체", "피해정도", "피해내용", "필수성", "사본매체", "복본위치", "확인할 것"])
        DLV = {"심각": 5, "높음": 4, "중간": 3, "낮음": 2, "경미": 1}
        DTXT = {5: "심각", 4: "높음", 3: "중간", 2: "낮음", 1: "경미"}
        rec = []
        for n, d in dmg:
            ndmg += 1
            f = allfiles.get(d.get("번호", ""), {})
            raw = re.sub(r"\s", "", d.get("정도", ""))
            lv = next((v for k, v in DLV.items() if k in raw), None)
            if lv is None:
                m = re.search(r"[1-5]", raw)
                lv = int(m[0]) if m else None
            task = f.get("과제") or d.get("과제", "")
            is_vital = task in fscore
            copy_ = f.get("이중", "")
            notes = []
            if lv is None:
                notes.append("피해정도는 1~5(경미·낮음·중간·높음·심각)")
            if not f:
                notes.append("기록물철 목록(--files)에 없는 관리번호 — 미리 만든 목록과 대조(7.4.1)")
            if lv and lv >= 4 and not copy_:
                notes.append("현재 기술로 복구 불가한 손상인데 복본 없음")
            if lv and lv >= 3 and is_vital:
                if copy_:
                    rec.append((-fscore.get(task, 0), -lv, task, f.get("철", "") or d.get("철", ""), d.get("번호", ""), copy_, f.get("보호위치", "")))
                else:
                    rec.append((-fscore.get(task, 0), -lv, task, f.get("철", "") or d.get("철", ""), d.get("번호", ""), "", ""))
            dv.append([task, f.get("철", "") or d.get("철", ""), d.get("번호", ""), f.get("매체", "") or d.get("매체", ""),
                       f"{DTXT[lv]}({lv})" if lv else d.get("정도", ""), d.get("내용", ""), "필수" if is_vital else "", copy_, f.get("보호위치", ""), " · ".join(notes)])
        dv.append([])
        for lv in (5, 4, 3, 2, 1):
            dv.append([f"{DTXT[lv]}({lv})", ["기록매체 특성을 완전히 잃어 활용 불가, 현재 기술로 복구 불가", "절반 이상 해독 가능하나 손상 부분은 복구 불가",
                                            "일부 소실, 해독 가능, 손상 부분 복구 가능", "기록 부분 직접 손상 없음, 보존처리 필요", "기록 부분 직접 손상 없음, 보존처리 불필요"][5 - lv]])
        rv = wb.create_sheet("업무복구 목록")
        rv.append(["복구 순위", "단위과제", "기록물철", "관리번호", "이용할 복본", "복본 위치", "조치"])
        for i, r in enumerate(sorted(rec), 1):
            nrec += 1
            rv.append([i, r[2], r[3], r[4], r[5], r[6], "보호기관 장의 승인 후 복본으로 이용(책임자·목적·대상·기간), 이용 뒤 반납·내역 통보(7.2)" if r[5]
                       else ("복본 없음·원본은 현재 기술로 복구 불가 — 다른 기관 보유분·관련 기록 확인, 복구 뒤 보호계획 재검토" if -r[1] >= 4
                             else "복본 없음 — 원본 응급복구(disaster-recovery) 후 이용, 복구 뒤 복본 제작")])
        rv.append([])
        rv.append(["순위 기준", "필수 과제 중 피해 '중간(3)' 이상 → 기능성 점수가 높은 업무부터, 같으면 피해가 큰 것부터(7.4.2)"])
        us = wb.create_sheet("이용 대장")
        us.append(["관리번호", "이용기관", "이용 책임자", "목적", "대상", "기간", "보호기관 승인일", "이용 매체(복본)", "반납일", "이용 내역 통보일"])
        for r in sorted(rec):
            us.append([r[4], "", "", "", "", "", "", r[5] or "", "", ""])
    sch = wb.create_sheet("일정·기준")
    oct31 = dt.date(a.year, 10, 31)
    for line in [
        ["기록관 검토·확정 기한", oct31.isoformat(), "매년 10월 31일까지 필수성·보호기간 검토·수정 → 영구기록물관리기관 협의·확정(5.4.3)",
         "지남" if today > oct31 else f"{(oct31 - today).days}일 남음"],
        ["11월 1일 이후 신설·변경", "", "즉시 영구기록물관리기관과 협의해 확정(5.4.3)", ""],
        ["보호계획 검토", "", "최소 1년에 한 번 이상 검토·갱신(6.3.1)", ""],
        ["판정 규칙", "", "한 기준이라도 높음(5) 또는 두 항목 이상 중간(3) 이상 → 필수(5.3)", ""],
        ["보호기간", "", "보존기간 이하, 기능성→갱신 주기 / 고유성→업무 반복 단위 / 증거성→효력 존속기간 중 가장 긴 기간(5.5)", ""],
        ["역할", "", "처리과: 분류·평가·보호기간 책정 / 기록관: 검토·수정, 기록관리기준표 반영·고시 / 영구기록물관리기관: 기준 제시·확정(표 3)", ""],
        ["기능연속성계획", "", "계획 수립 기관은 기록관이 계획수립 부서와 협력(「재난 및 안전관리 기본법」 제25조의4⑤)", ""],
        ["단위과제 미사용 기관", "", "단위업무 또는 기록물철 단위로 같은 평가표(표 5)", ""],
    ]:
        sch.append(line)
    for sh, widths in ((ws, (16, 24, 40, 8, 8, 8, 10, 30, 10, 12, 60, 20)), (lst, (20, 30, 16, 12, 8, 8, 8, 12, 10, 10, 14, 14, 10, 14, 60)),
                       (prot, (16, 30, 12, 10, 14, 12, 70)), (sch, (24, 12, 90, 14))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"단위과제 {len(tasks)} · 필수 {len(vital)} · 평가 필요 {todo} · 확인할 것 {issues_n} · 관리목록 철 {nfiles}"
          + (f" · 피해 {ndmg}철 · 업무복구 대상 {nrec}" if a.damage else "") + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
