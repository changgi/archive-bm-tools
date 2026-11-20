#!/usr/bin/env python3
"""과거 공문서 분류번호로 당시 보존기간 찾기: 1999년 이전에 생산된 기록물의 옛 분류번호(5자리)와 제목으로
생산 당시 「공문서 분류번호 및 보존기간표」(1985·1992·1997)의 기능명·단위사안·보존기간 후보를 찾고, 지금 책정된 보존기간과 견준다.
쓰는 때: 기존(비전자) 기록물 정리, 보존기간 재평가, 이관·폐기 심의자료에 '생산 당시 기준'을 적어야 할 때.

근거: 국가기록원 기록관리 자료실 > 기록물분류기준 > 공문서분류 및 보존기간표(2026-11-20 내려받아 확인)
      1985·1992 엑셀(단위사안별 보존기간 열), 1997 엑셀(「공문서분류번호 구성체계 및 사용방법」:
      [ ] 생산 문서 보존기간(영·30·20·10·5·3·1), ( ) 접수 문서 보존기간, [☆] 조기호표 적용, [＃]·[◎] 해당 분류 시작부분 설명)
      → assets/legacy/gongmunseo_legacy.csv(1964·1979·1985·1992·1997, 52,632행)
      1964 「공문서종별보존기간책정기준표」(HWP, 분류번호 3~4자리)는 본문 표를 글자 위치로 옮기고 한자 옆에 한글 독음을 붙였다(4,930행, 1964~1978 생산분에 적용).
      1979 「공문서보존기간종별책정기준표」(총리령 제223호 별표, HWP 3.0)는 rhwp 엔진(@rhwp/core, 이음오피스가 쓰는 오픈소스 HWP 엔진)으로 읽어 옮겼다
      (8,389행, 1979. 7. 1.~1984 생산분, 영구(갑)·영구(을)·준영구 구분). 정부처무규정(1949~1955)은 규정 본문이라 표로 옮기지 않았다.
      2000년 이후(기록물분류기준표)는 기관별 공개본(국가기록원 같은 메뉴 '기록물분류기준표' 탭)을 본다.

사용:  python legacy_lookup.py 목록.xlsx -o 산출/과거분류_보존기간.xlsx [--edition 1985|1992|1997]
목록 열(머리글 이름으로 찾음): 분류번호(옛 공문서분류번호), 생산연도, 제목(철 제목), 보존기간(지금 책정된 값, 있으면)
원칙:  후보를 보여 줄 뿐이다. 판정은 담당자가 당시 기준과 지금 기준(시행령 별표 1·기관 기준표)을 함께 보고 한다.
"""
import argparse
import csv
import os
import re
import sys

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "legacy", "gongmunseo_legacy.csv")
RANK = {"영구(갑)": 9, "영구(을)": 9, "영구": 9, "준영구": 8, "30년": 7, "20년": 6, "10년": 5, "5년": 4, "3년": 3, "2년": 2.5, "1년": 2, "6개월": 1, "연말 폐기": 0.5, "폐기": 0.5}
ALIAS = {"코드": ["분류번호", "공문서분류번호", "옛분류번호", "구분류번호", "기능분류번호"], "연도": ["생산연도", "생산년도", "연도"],
         "제목": ["제목", "철제목", "기록물철제목"], "보존": ["보존기간", "현보존기간", "책정보존기간"], "번호": ["관리번호", "철번호", "등록번호"]}


def edition(year):
    if year is None:
        return None
    if year < 1964:
        return "1964년 이전"
    if year < 1979:
        return "1964"
    if year < 1985:
        return "1979"
    if year < 1992:
        return "1985"
    if year < 1997:
        return "1992"
    if year < 2000:
        return "1997"
    return "2000년 이후"


def tokens(s):
    return {w for w in re.findall(r"[가-힣一-鿿A-Za-z0-9]{2,}", s or "") if w not in ("관계", "문서", "관계문서", "관한", "사항", "일반", "기타")}


def overlap(tt, item):
    """제목 낱말이 단위사안에 들어 있거나 그 반대면 점수(붙여 쓴 옛 표기 '기본정책방침관계문서' 대응)"""
    it = re.sub(r"\s", "", item)
    return sum(len(w) for w in tt if w in it) + sum(len(w) for w in tokens(item) if any(w in x for x in tt))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("list")
    ap.add_argument("--edition", choices=["1964", "1979", "1985", "1992", "1997"], help="판을 정해 찾기(기본: 생산연도로 고름)")
    ap.add_argument("-o", "--out", default="산출/과거분류_보존기간.xlsx")
    a = ap.parse_args()
    table = {}
    with open(ASSET, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            table.setdefault((r["판"], r["분류번호"]), []).append(r)
    if a.list.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.list, encoding="utf-8-sig")))
    head = [re.sub(r"\s", "", h) for h in rows[0]]
    idx = {k: next((head.index(x) for x in al if x in head), None) for k, al in ALIAS.items()}
    if idx["코드"] is None:
        print("입력 오류: 분류번호 열이 필요합니다.", file=sys.stderr)
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대조"
    ws.append(["원본 행", "관리번호", "분류번호", "생산연도", "제목", "적용 판", "당시 기능", "후보 단위사안(당시 보존기간)", "당시 최장", "지금 보존기간", "판정"])
    flagged = 0
    for n, r in enumerate(rows[1:], 2):
        if not any(r):
            continue
        g = lambda k, r=r: (r[idx[k]] if idx[k] is not None and idx[k] < len(r) else "").strip()
        code = re.sub(r"\D", "", g("코드"))[:5]
        y = re.search(r"(19|20)\d{2}", g("연도"))
        ed = a.edition or edition(int(y[0]) if y else None)
        cur = re.sub(r"\s", "", g("보존"))
        if ed in (None, "1964년 이전", "2000년 이후"):
            msg = {None: "생산연도 없음 — --edition으로 판을 정함", "1964년 이전": "1964년 이전 — 정부처무규정(1949 대통령훈령 제1호) 제69조: 갑종 영구·을종 10년·병종 3년·정종 1년(완결 다음 해부터) — 철 표지의 보존종별 확인",
                   "2000년 이후": "2000년 이후 — 기관 기록물분류기준표(2012 공개본)·기록관리기준표 확인"}[ed]
            ws.append([n, g("번호"), g("코드"), g("연도"), g("제목"), ed or "", "", "", "", cur, msg])
            continue
        if ed in ("1964", "1979"):
            code = re.sub(r"\D", "", g("코드"))
        elif len(code) < 5:
            code = code.ljust(5, "0")
        items = table.get((ed, code), [])
        if not items:
            ws.append([n, g("번호"), g("코드"), g("연도"), g("제목"), ed, "", "", "", cur, f"{ed}판에 분류번호 {code} 없음 — 번호 확인"])
            flagged += 1
            continue
        fn = items[0]["기능"]
        cand = [i for i in items if i["단위사안"]]
        tt = tokens(g("제목"))
        scored = sorted(cand, key=lambda i: -overlap(tt, i["단위사안"]))
        hit = [i for i in scored if overlap(tt, i["단위사안"])][:5] or scored[:5]
        show = " / ".join(f"{i['단위사안'][:30]}({i['보존기간']})" for i in hit)
        pers = [i["보존기간"] for i in cand if i["보존기간"] in RANK]
        longest = max(pers, key=RANK.get) if pers else (cand[0]["보존기간"] if cand else items[0]["보존기간"])
        best = next((i["보존기간"] for i in hit if overlap(tt, i["단위사안"]) and i["보존기간"] in RANK), "")
        judge = "제목과 맞는 단위사안 없음 — 후보 목록에서 고름" if not best else f"제목이 맞는 단위사안의 당시 보존기간 {best}"
        if cur in RANK and best in RANK and RANK[cur] < RANK[best]:
            judge += f" — 지금 보존기간({cur})이 당시 기준보다 짧음, 재평가 때 확인"
            flagged += 1
        if any(i["보존기간"] == "조기호표" for i in hit):
            judge += " · [☆] 조기호표 적용 항목 있음"
        ws.append([n, g("번호"), g("코드"), g("연도"), g("제목"), ed, fn, show, longest, cur, judge])
    s = wb.create_sheet("판 구분")
    for row in (["생산연도", "적용 판", "비고"], ["~1963", "정부처무규정(1949~1955)", "이 스크립트 범위 밖"],
                ["1964~1978", "공문서종별보존기간책정기준표 1964", "한자 표기, 분류번호 3~4자리"],
                ["1979~1984", "공문서보존기간종별책정기준표 1979(총리령 제223호 별표)", "영구(갑)·(을), 준영구는 규칙 제2조③의 하한 기간(아래)"],
                ["1979 준영구", "인사발령 발령 후 15년 · 훈령·예규·방침·기준 폐지 후 5년 · 정책·계획 완료 후 3년 · 허가·인가 취소·만료 후 · 인사기록 사망 후 5년 · 기타 기관장 인정 기간(10년 미만이면 10년), 10년 지나면 3년마다 점검", "규칙 제2조③④"],
                ["1985~1991", "공문서 분류번호 및 보존기간표 1985", ""], ["1992~1996", "공문서 분류번호 및 보존기간표 1992", "조기호표 별도"],
                ["1997~1999", "공문서 분류번호 및 보존기간표 1997", "[ ] 생산·( ) 접수·[☆] 조기호표"], ["2000~", "기록물분류기준표(기관별)", "이 스크립트 범위 밖"],
                ["주의", "판의 시행 시점과 기관의 실제 적용 시점이 다를 수 있다. 철 표지의 분류번호·보존기간 표시를 먼저 본다.", ""]):
        s.append(row)
    for i, w in enumerate((8, 14, 10, 8, 30, 8, 18, 70, 8, 10, 50)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"{len(rows) - 1}행 · 확인 필요 {flagged} → {a.out}")
    return 1 if flagged else 0


if __name__ == "__main__":
    sys.exit(main())
