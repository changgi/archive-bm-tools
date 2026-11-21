#!/usr/bin/env python3
"""행정박물 관리대상 선별: 후보 목록으로 「행정박물 관리지침」(2021) 표 4 선별 체크리스트(7기준 × 처리과·기록관 5점)를 만들고,
점수가 채워져 있으면 합산(70점 만점, 36점 이상 관리대상)과 확인할 것(공직자선물 신고·관인 폐기공고·증서류·대형 박물·훼손)을 적는다.

사용:  python artifact_select.py 후보목록.xlsx [--cut 36] -o 산출/행정박물_선별.xlsx
목록 열(머리글 이름으로 찾음): 행정박물명, 생산일자, 생산부서, 유형, 수량, 비고(상태·가액 등)
  점수 열(있으면 합산): 처리과_중요도·처리과_인물·처리과_상징성·처리과_예술성·처리과_희소성·처리과_활용성·처리과_보존가능성,
                      기록관_… 같은 7개 — 머리글에 '처리과'/'기록관'과 기준 낱말이 들어 있으면 된다. 1~5점.
근거:  국가기록원 「행정박물 관리지침」(2021. 12.) 4.1 처리과 작성·기록관 검토 확정, 4.2 선별기준 7개(행사·사업의 중요도, 주요 인물 행적 관련성,
       상징성, 예술성, 희소성, 활용성, 보존가능성 — 생산 당시 원형을 간직한 것이 원칙), 4.3 표 4 체크리스트(처리과·기록관 합산 70점 중 36점 이상),
       부록 3 자주 묻는 질문(관인 폐기공고문 첨부 이관, 증서류는 일반문서, 공직자선물 미화 100달러·10만원 이상 신고 — 공직자윤리법 시행령 제29조,
       표지석·대형 현수막은 탁본·사진으로 갈음 가능, 미지정 행정박물 폐기는 전문요원 심사·평가심의회 — 시행령 제57조⑥·제43조)
원칙:  점수는 처리과와 기록관이 매긴다. 이 스크립트는 빈 체크리스트를 만들고 합산·확인만 한다. 기관 세부기준(--cut)은 기관이 정한다.
"""
import argparse
import re
import sys

CRIT = [("중요도", "행사‧사업의 중요도", "공공기관의 국내·외 행사, 사업, 정책 등의 중요도"),
        ("인물", "주요 인물 행적 관련성", "주요 직위자(기관장 등)의 업무수행 행적 등과 관련된 정도"),
        ("상징", "상징성", "공공기관의 변천, 고유업무의 수행과정을 상징적으로 보여주는 정도"),
        ("예술", "예술성", "내용 및 형태가 심미적, 예술적 가치를 지니는 정도"),
        ("희소", "희소성", "공공기관만의 유일하거나 희소성을 갖는 정도"),
        ("활용", "활용성", "홍보, 교육, 전시, 학술연구 등 활용(가능) 정도"),
        ("보존", "보존가능성", "훼손정도, 보존·수리·복원과 관련된 공간·비용·인력을 감안한 보존가능성")]
ALIAS = {"이름": ["행정박물명", "명칭", "박물명"], "일자": ["생산일자", "취득일자", "수령일"], "부서": ["생산부서", "처리과", "부서"],
         "유형": ["유형", "행정박물유형"], "수량": ["수량", "수량(점)"], "비고": ["비고", "상태", "특이사항", "내용요약"]}
NOTES = [(r"선물|증정|기증", "공직자선물이면 소속기관장에게 신고(미화 100달러·한화 10만원 이상), 영구보존 필요 선물은 영구기록물관리기관으로 이관(공직자윤리법 시행령 제29조) — 유형 '그 밖의 유형(공직자선물)'"),
         (r"관인|직인|인장|청인|도장", "활용이 끝난 관인은 관보 폐기공고 후 인영이 보이는 폐기공고문을 첨부해 기록관으로 이관"),
         (r"협약서|양해각서|MOU|증서|계약서", "증서 원본은 일반문서로 관리(형태가 다르면 분리등록). 함께 받은 형상기록물만 행정박물로 등록 가능"),
         (r"표지석|현수막|대형|조형물|간판", "탁본·사진 등으로 내용과 실물 모습을 남겨 보존 목적을 달성하면 실물 관리 불필요 — 가용자원·비용·가치 종합 판단"),
         (r"파손|훼손|노후|부식|깨짐|곰팡이", "수리·복원이 불가능해 원래 가치를 잃었으면 제외. 중요도가 높으면 손상 실물도 선정 가능"),
         (r"복제|모조|재제작|레플리카", "생산 당시 원형을 간직한 것을 선정하는 것이 원칙 — 복제품인지 확인")]


def find(head, key):
    for name in ALIAS[key]:
        for i, h in enumerate(head):
            if name in h.replace(" ", ""):
                return i
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="행정박물 관리대상 선별 체크리스트")
    ap.add_argument("list")
    ap.add_argument("--cut", type=int, default=36, help="관리대상 합산 점수 기준(지침 36/70, 기관 세부기준으로 바꿀 수 있음)")
    ap.add_argument("-o", "--out", default="산출/행정박물_선별.xlsx")
    a = ap.parse_args()
    import openpyxl
    from openpyxl.styles import Alignment, Font
    try:
        ws0 = openpyxl.load_workbook(a.list, read_only=True, data_only=True).active
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    except Exception as e:
        print(f"입력 오류: 목록을 읽을 수 없습니다({e})", file=sys.stderr)
        return 2
    rows = [r for r in rows if any(r)]
    if not rows:
        print("입력 오류: 빈 목록", file=sys.stderr)
        return 2
    head = [h.replace(" ", "") for h in rows[0]]
    col = {k: find(head, k) for k in ALIAS}
    if col["이름"] is None:
        print("입력 오류: '행정박물명' 열이 필요합니다.", file=sys.stderr)
        return 2
    score_col = {}
    for who in ("처리과", "기록관"):
        for key, *_ in CRIT:
            j = next((i for i, h in enumerate(head) if who in h and key in h), None)
            if j is not None:
                score_col[(who, key)] = j
    get = lambda r, k: (r[col[k]] if col[k] is not None and col[k] < len(r) else "")
    wb = openpyxl.Workbook()
    res = wb.active
    res.title = "선별 결과"
    res.append(["번호", "행정박물명", "생산일자", "생산부서", "유형", "수량", "처리과(35)", "기록관(35)", "합산(70)", "판정", "확인할 것"])
    n_sel = n_part = 0
    for i, r in enumerate(rows[1:], 1):
        tot = {}
        bad = []
        for who in ("처리과", "기록관"):
            vals = []
            for key, label, _ in CRIT:
                j = score_col.get((who, key))
                v = r[j].strip() if j is not None and j < len(r) else ""
                if v:
                    if re.fullmatch(r"[1-5](\.0)?", v):
                        vals.append(int(float(v)))
                    else:
                        bad.append(f"{who} {label} '{v}'(1~5)")
            tot[who] = sum(vals) if len(vals) == len(CRIT) else (None if not vals else f"{sum(vals)}(미완 {len(vals)}/7)")
        both = [tot[w] for w in ("처리과", "기록관")]
        if all(isinstance(x, int) for x in both):
            s = sum(both)
            verdict = "관리대상" if s >= a.cut else "비대상"
            n_sel += verdict == "관리대상"
        else:
            s, verdict = "", "점수 입력 필요" if not any(both) else "점수 미완"
            n_part += 1
        text = " ".join(get(r, k) for k in ("이름", "유형", "비고"))
        notes = [msg for rx, msg in NOTES if re.search(rx, text)] + bad
        res.append([i, get(r, "이름"), get(r, "일자"), get(r, "부서"), get(r, "유형"), get(r, "수량"),
                    tot["처리과"] if tot["처리과"] is not None else "", tot["기록관"] if tot["기록관"] is not None else "", s, verdict, " / ".join(notes)])
    # 표 4 체크리스트(후보마다 한 덩어리, 인쇄·서명용)
    ck = wb.create_sheet("선별 체크리스트(표4)")
    line = 1
    for i, r in enumerate(rows[1:], 1):
        ck.cell(line, 1, f"행정박물 관리대상 선별 체크리스트 ({i})").font = Font(bold=True, size=13)
        ck.append([])
        ck.append(["검토일자", "", "처리과", "(서명)", "검토일자", "", "기록관", "(서명)"])
        ck.append(["행정박물명", get(r, "이름"), "", "", "생산일자", get(r, "일자")])
        ck.append(["생산부서", get(r, "부서"), "", "", "행정박물 유형", get(r, "유형"), "수량", get(r, "수량")])
        ck.append(["선별기준", "선별내용", "처리과(5·4·3·2·1)", "기록관(5·4·3·2·1)"])
        for key, label, desc in CRIT:
            pj, gj = score_col.get(("처리과", key)), score_col.get(("기록관", key))
            ck.append([label, desc, r[pj] if pj is not None and pj < len(r) else "", r[gj] if gj is not None and gj < len(r) else ""])
        ck.append(["합산점수(기록관)", f"합산점수가 {a.cut}점 이상이면 관리대상으로 선별, 미만이면 선별하지 않음"])
        ck.append(["종합의견", ""])
        ck.append(["검토결과", "관리대상(    )    비대상(    )"])
        ck.append([])
        line = ck.max_row + 2
    for c, w in zip("ABCDEFGH", (20, 52, 18, 18, 14, 16, 10, 10)):
        ck.column_dimensions[c].width = w
    for c, w in zip("ABCDEFGHIJK", (6, 30, 12, 14, 14, 6, 10, 10, 10, 12, 90)):
        res.column_dimensions[c].width = w
    for c in res["K"]:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    g = wb.create_sheet("기준")
    g.append(["선별기준", "설명(행정박물 관리지침 4.2)"])
    for row in (["원칙", "생산 당시의 원형을 간직한 행정박물을 관리대상으로 선정"],
                ["① 행사와 사업의 중요도", "기관 설립·국제대회·국내 주요 행사·대규모 사업과 관련해 그 내용을 보여주는 것(설립 기념품, 고유업무 기획사업 물품 등)"],
                ["② 주요 인물의 행적 관련성", "기관장·직원이 업무 수행에서 생산·사용한 것(주요 직위자가 받은 외국 선물, 사용한 집기·명패·유니폼 등)"],
                ["③ 상징성", "기관 변천·고유업무 수행을 상징적으로 보여주는 것(직무 관련 의복류, 현판·깃발·배지 등)"],
                ["④ 예술성", "심미적·예술적 가치가 높아 보존이 필요한 것"],
                ["⑤ 희소성", "다른 기관에서 생산되지 않는 유일·희소한 것(특정 행사 기념주화 등)"],
                ["⑥ 활용성", "홍보·대국민 교육·전시·학술연구에 쓸 수 있는 것"],
                ["⑦ 보존가능성", "공간·비용 등 가용능력 고려. 단 그 때문에 가치 있는 것이 빠지지 않게. 복원 불가로 가치를 잃은 것은 제외(중요도 높으면 예외)"],
                ["절차", "처리과 작성 → 기록관 의견 반영·검토·확정 → 처리과 통보 → 관리카드(표 5, 기본정보 11항목) 등록 — artifacts_check.py로 점검"],
                ["근거", "국가기록원 「행정박물 관리지침」(2021. 12.) 4장·부록 3, 공공기록물법 시행령 제57조·별표 4·5"]):
        g.append(row)
    g.column_dimensions["A"].width = 26
    g.column_dimensions["B"].width = 110
    wb.save(a.out)
    print(f"후보 {len(rows) - 1}점 · 관리대상 {n_sel} · 점수 입력·미완 {n_part} (기준 {a.cut}/70) → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
