#!/usr/bin/env python3
"""개별 법령 법정 보존기간 점검: 기관 기록관리기준표의 단위과제를 개별 법령이 최소 보존기간을 정한 기록 종류와
낱말로 맞춰, 기준표 보존기간이 법정 기간보다 짧으면 '법정 최소 미달'로 표시하고 맞출 보존기간을 제안한다.

근거(law.go.kr 2026-11-21 원문 확인) → assets/law/statutory_minimum.csv (44종 — 최소 보존 40·파기 상한 4, 기관 기준표 대조용으로 고른 것)
  - 「근로기준법」 제42조, 시행령 제22조: 근로자 명부·근로계약서·임금대장·고용·해고·퇴직·승급·휴가·서면 합의 3년(기산일 따로)
  - 「산업안전보건법」 제164조, 시행규칙 제241조: 선임 서류·산업재해 기록·유해성 조사 3년, 위원회 회의록 2년,
    작업환경측정·건강진단 5년(고시 물질 30년), 기관석면조사 3년
  - 「국세기본법」 제85조의3②: 장부·증거서류 법정신고기한부터 5년(역외거래 7년)
  - 「소방시설 설치 및 관리에 관한 법률 시행규칙」 제23조④: 자체점검 실시결과 보고서 점검이 끝난 날부터 2년
  - 「의료법 시행규칙」 제15조①: 진료기록부 등 9종
  - 공공기관 공통: 「장애인복지법」 제25조의2(인식개선교육 자료 3년)·「장애인고용촉진법」 제5조의2, 「개인정보의 안전성
    확보조치 기준」 제8조(접속기록 1년, 대량·고유식별·민감정보 2년)·제5조(접근 권한 내역 3년), 「재난안전법 시행령」 제76조
    (재난상황 기록 5년), 「자연재해대책법 시행령」 제33조의3(재해대장 5년), 「공직자윤리법 시행령」 제36조의3(재산등록 자료
    퇴직 후 10년), 「공무원 제안 규정」 제23조(미채택 제안 2년), 「실내공기질 관리법」 제12조(측정결과 10년), 「영유아보육법」
    제15조의4(어린이집 CCTV 60일)·시행규칙 제9조의7(열람대장 3년), 「사학기관 재무·회계 규칙」 제53조의2(장부·증빙 5년), 「보조금 관리에 관한 법률」 시행령 제10조의2·「지방자치단체 보조금 관리에 관한 법률」 제16조③(보조사업 자료 5년)
전체 목록 → assets/law/statutory_catalog.csv (10,084건: 법령 1350건·법령 별표·서식 314건·행정규칙 3695건·자치법규 246개 지자체 4725건)
  국가법령정보센터 '조문내용' 검색(현행 법령·행정규칙)으로 '년'·숫자 표현 28가지를 모으고, 법령 원문 2,761개, 별표·서식 PDF 9,730개·한글 파일 815개, 행정규칙 원문 2,228개, 자치법규 원문 4,378개에서
  조문을 뽑았다(검색 결과의 기간 문장 651건 중 637건 확보). 별표·서식과 '(검색 문구)' 행은 검색 결과 문구(일부)만 담았으니 원문을 연다.
  기준표와 이름이 닮은 법정 기록은 '전체 목록 후보' 시트로 따로 보여 준다(판정 아님). `--find 낱말`로 목록을 찾는다.
  자치법규(조례·규칙)는 지자체마다 달라 `--org 가평군`처럼 기관을 주면 그 지자체 것만 후보로 넣는다(주지 않으면 뺀다).사용:  python statutory_check.py 기준표.xlsx [--law 근로기준법,산업안전보건법] [--org 가평군] -o 산출/법정보존기간_점검.xlsx
       python statutory_check.py --list   (법정 기간 표 보기)
기준표 열(머리글 이름으로 찾음): 단위과제명(단위과제), 보존기간, (선택) 단위과제코드, 단위과제 설명(업무설명·설명)
성격:  대부분 '최소 보존'(그 기간 이상 둔다)이고, '지나면 파기(상한)'(사회보장정보·보조금관리정보 등 N년이 지나면 파기)는
       기준표가 더 길면 '확인 필요'로만 적는다. 공공기록물로 등록된 기록은 공공기록물법 평가·폐기 절차를 따르기 때문이다.
원칙:  법정 기간은 '최소'다. 기준표가 더 길면 그대로 둔다. 공공기록물 보존기간은 1·3·5·10·30년·준영구·영구라
       법정 2년이면 3년, 7년이면 10년처럼 그 이상인 가장 짧은 구분을 제안한다(시행령 제26조①).
       단위과제명으로 맞추면 '법정 최소 미달', 설명 낱말로만 맞추면 '확인 필요'로 낮춰 표시한다.
       낱말로 고른 후보일 뿐이니 '적용 대상'(예: 근로기준법은 공무원이 아닌 근로자)과 기산일을 담당자가 확인한다.
       공공기록물은 법정 기간이 지나도 공공기록물법 평가·폐기 절차로만 폐기한다.
"""
import argparse
import csv
import os
import re
import sys

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "law", "statutory_minimum.csv")
CLASSES = [1, 3, 5, 10, 30]


def load(laws=None):
    with open(ASSET, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if laws:
        rows = [r for r in rows if any(l.strip() and l.strip() in r["법령"] for l in laws)]
    return rows


def years(p):
    """보존기간 글자를 연수로(60일·6개월은 소수): 영구·준영구는 아주 큰 값, 모르면 None."""
    p = re.sub(r"\s|간$|동안$|이상$|까지$", "", re.sub(r"\s", "", p or ""))
    if p.startswith("준영구"):
        return 999
    if p.startswith("영구"):
        return 1000
    m = re.match(r"(\d+)(년|개월|일|시간)", p)
    return int(m.group(1)) / {"년": 1, "개월": 12, "일": 365, "시간": 8760}[m.group(2)] if m else None


def suggest(n):
    return next((f"{c}년" for c in CLASSES if c >= n), "준영구")


CATALOG = os.path.join(os.path.dirname(ASSET), "statutory_catalog.csv")
STOP = {"관리", "운영", "업무", "일반", "지원", "계획", "추진", "사업", "관련", "처리", "시행", "기관", "총괄", "기타", "대장", "자료", "결과"}


def load_catalog():
    if not os.path.exists(CATALOG):
        return []
    with open(CATALOG, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def tokens(name):
    return [w for w in re.split(r"[\s·ㆍ,()/\-]+|및|등", name or "") if len(w) >= 3 and w not in STOP]


def catalog_hits(cat, name, limit=3, org=None, all_rows=False):
    """단위과제명의 낱말(3글자 이상)이 조문 제목·기록 문구에 들어 있는 법정 기록. 많이 겹칠수록 앞.
    자치법규는 org(지자체 이름 일부)를 줄 때 그 기관 것만 넣는다."""
    tk = tokens(name)
    if not tk:
        return []
    scored = []
    for c in cat:
        if not all_rows and c.get("판정") == "의무 아님 가능성":  # 보존 의무 판정기(retention_judge.py)가 걸러 낸 행
            continue
        if c["출처"].startswith("자치법규") and not (org and (org in c.get("기관", "") or org in c["법령"])):
            continue
        hay = c["조문"] + " " + c["기록(문구)"] + " " + c["문장"][:160] + ("" if c["출처"] == "법령" else " " + c["법령"])
        sc = sum(len(w) for w in tk if w in hay)
        if sc:
            scored.append((sc, c))
    scored.sort(key=lambda x: -x[0])
    return [c for _, c in scored[:limit]]


ITEMS = os.path.join(os.path.dirname(ASSET), "statutory_items.csv")


def load_items():
    if not os.path.exists(ITEMS):
        return []
    with open(ITEMS, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def items_hits(items, name, limit=4):
    """위임 추적기(retention_tree.py)가 호·항·단서 단위로 나누고 시행령·시행규칙·고시까지 따라간 행 가운데
    단위과제명 낱말이 기록·참조 조문·조문 제목에 든 것."""
    tk = tokens(name)
    if not tk:
        return []
    scored = []
    for c in items:
        if not c.get("보존기간") or c.get("판정") == "의무 아님 가능성":
            continue
        hay = c["기록(항목)"] + " " + c.get("참조 조문", "") + " " + c["조문"]
        sc = sum(len(w) for w in tk if w in hay)
        if sc:
            scored.append((sc, c))
    scored.sort(key=lambda x: -x[0])
    return [c for _, c in scored[:limit]]


def match(rows, name, desc=""):
    """단위과제명에서 먼저, 없으면 설명에서 법정 기록 종류를 찾는다. 같은 과제에 여러 종류면 모두 준다."""
    ok = [r for r in rows if not (r.get("제외 낱말") and re.search(r["제외 낱말"], (name or "") + " " + (desc or "")))]
    hit = [r for r in ok if re.search(r["낱말"], name or "")]
    if hit:
        return hit, "단위과제명"
    return [r for r in ok if desc and re.search(r["낱말"], desc)], "설명"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table", nargs="?")
    ap.add_argument("--law", help="쉼표로 나눈 법령 이름 일부(기본: 전부)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--find", help="전체 목록(statutory_catalog.csv)에서 낱말로 찾기 — 법령명·조문·기록·문장")
    ap.add_argument("--no-catalog", action="store_true", help="'전체 목록 후보' 시트를 만들지 않음")
    ap.add_argument("--all-catalog", action="store_true", help="판정기가 '의무 아님 가능성'으로 본 행도 후보에 넣음")
    ap.add_argument("--org", help="자치법규 후보를 넣을 지자체 이름(예: 가평군·서울특별시 마포구)")
    ap.add_argument("-o", "--out", default="산출/법정보존기간_점검.xlsx")
    a = ap.parse_args()
    rows = load(a.law.split(",") if a.law else None)
    if a.find:
        cat = load_catalog()
        kw = [k for k in re.split(r"\s+", a.find) if k]
        hits = [c for c in cat if all(k in " ".join(c.values()) for k in kw) and (not a.org or not c["출처"].startswith("자치법규") or a.org in c.get("기관", "") + c["법령"])]
        for c in hits[:200]:
            print(f"[{c['출처']}{'·' + c['기관'] if c.get('기관') else ''}] {c['법령']} {c['조문'][:40]} | {c['보존기간']} | {(c['기록(문구)'] or c['문장'])[-70:]} | {c['원문']}")
        print(f"전체 목록 {len(cat)}건 중 {len(hits)}건" + (" (앞 200건만 보임)" if len(hits) > 200 else ""))
        return 0
    if a.list or not a.table:
        for r in rows:
            print(f"{r['법령']} {r['조항']} | {r['기록 종류']} | {r['최소 보존기간']} | 기산 {r['기산'] or '-'} | {r['적용 대상']}")
        print(f"법정 보존기간 {len(rows)}종 (law.go.kr 2026-11-21 확인)")
        return 0
    if a.table.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        data = [["" if v is None else str(v).strip() for v in r] for r in load_workbook(a.table, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        data = list(csv.reader(open(a.table, encoding="utf-8-sig")))
    data = data or [[]]
    hi = max(range(min(10, len(data))), key=lambda i: sum(bool(c) for c in data[i]))
    head = [re.sub(r"\s", "", h) for h in data[hi]]
    find = lambda *ns: next((head.index(n) for n in ns if n in head), None)
    ni, pi = find("단위과제명", "단위과제"), find("보존기간")
    ci, di = find("단위과제코드", "코드", "단위과제번호"), find("단위과제설명", "업무설명", "설명")
    if ni is None or pi is None:
        print("입력 오류: 단위과제명·보존기간 열이 필요합니다.", file=sys.stderr)
        return 2
    cell = lambda r, i: r[i] if i is not None and i < len(r) else ""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "단위과제코드", "단위과제명", "기준표 보존기간", "법령", "조항", "기록 종류", "법정 기간", "기산", "적용 대상",
               "판정", "제안 보존기간", "비고"])
    short, maybe, checked, over = 0, 0, 0, 0
    for n, r in enumerate(data[hi + 1:], hi + 2):
        name, per = cell(r, ni), re.sub(r"\s", "", cell(r, pi))
        if not name:
            continue
        hits, via = match(rows, name, cell(r, di))
        if not hits:
            continue
        checked += 1
        y = years(per)
        for h in hits:
            need = years(h["최소 보존기간"])
            if y is None:
                judge, sug = "보존기간 표기 확인", ""
            elif h.get("성격", "").startswith("지나면"):
                judge, sug = ("파기 기한(상한)보다 김 — 확인 필요(그 정보·자료는 기한이 지나면 파기. 공공기록물로 등록된 기록은 공공기록물법 절차)", "") if y > need else ("상한 안", "")
                over += y > need
            elif y < need and via == "설명":
                judge, sug = "확인 필요 — 설명 낱말로만 맞춤(해당 기록이 이 과제에 철해지는지 확인)", suggest(need)
                maybe += 1
            elif y < need:
                judge, sug = "법정 최소 미달 — 기준표 보존기간을 늘림", suggest(need)
                short += 1
            else:
                judge, sug = "충족", ""
            ws.append([n, cell(r, ci), name, per, h["법령"], h["조항"], h["기록 종류"], h["최소 보존기간"], h["기산"], h["적용 대상"],
                       judge, sug, h["비고"]])
    cand = 0
    if not a.no_catalog:
        cat = load_catalog()
        cs = wb.create_sheet("전체 목록 후보")
        cs.append(["원본 행", "단위과제명", "기준표 보존기간", "출처", "법령·행정규칙·자치법규", "기관", "조문", "법정 기간", "성격", "기록(문구)", "문장(일부)", "원문", "비교"])
        for n, r in enumerate(data[hi + 1:], hi + 2):
            name, per = cell(r, ni), re.sub(r"\s", "", cell(r, pi))
            if not name:
                continue
            for c in catalog_hits(cat, name, org=a.org, all_rows=a.all_catalog):
                ny, y = years(c["보존기간"]), years(per)
                if c.get("성격", "").startswith("지나면"):
                    cmp = "파기 기한(상한)보다 김 — 원문 확인" if ny and y and y > ny else "상한 안"
                else:
                    cmp = "법정 기간이 더 김 — 같은 기록인지 원문 확인" if ny and y and ny > y else "기준표가 같거나 김"
                cand += cmp.startswith("법정")
                cs.append([n, name, per, c["출처"], c["법령"], c.get("기관", ""), c["조문"], c["보존기간"], c.get("성격", ""), c["기록(문구)"], c["문장"][:200], c["원문"], cmp])
    icand = 0
    if not a.no_catalog:
        items = load_items()
        if items:
            isht = wb.create_sheet("호별·위임 추적 후보")
            isht.append(["원본 행", "단위과제명", "기준표 보존기간", "법령", "조문", "항·호", "기록(항목)", "참조 조문", "법정 기간", "조건(단서)", "의무자",
                         "근거 경로(법률→시행령·시행규칙→고시)", "상태", "비교"])
            for n, r in enumerate(data[hi + 1:], hi + 2):
                name, per = cell(r, ni), re.sub(r"\s", "", cell(r, pi))
                if not name:
                    continue
                for c in items_hits(items, name):
                    ny, y = years(re.sub(r"\(.*\)|예외.*", "", c["보존기간"])), years(per)
                    if "지나면 파기" in c["보존기간"]:
                        cmp = "파기 기한(상한)보다 김 — 원문 확인" if ny and y and y > ny else "상한 안"
                    else:
                        cmp = "법정 기간이 더 김 — 같은 기록인지 원문 확인" if ny and y and ny > y else "기준표가 같거나 김"
                    icand += cmp.startswith("법정")
                    isht.append([n, name, per, c["법령"], c["조문"], c["항"] + (c["호"] + ("호" if c["호"].isdigit() else "목") if c["호"] else ""),
                                 c["기록(항목)"], c.get("참조 조문", ""), c["보존기간"], c.get("조건", ""), c.get("의무자", ""), c["근거 경로"], c["상태"], cmp])
    tb = wb.create_sheet("법정 보존기간")
    tb.append(["법령", "조항", "시행", "기록 종류", "법정 기간", "성격", "기산", "적용 대상", "찾는 낱말", "제외 낱말", "비고"])
    for r in rows:
        tb.append([r["법령"], r["조항"], r["시행"], r["기록 종류"], r["최소 보존기간"], r.get("성격", "최소 보존"), r["기산"], r["적용 대상"], r["낱말"].replace("\\s*", ""),
                   r.get("제외 낱말", ""), r["비고"]])
    s = wb.create_sheet("요약")
    for row in (["법정 보존기간 종류", len(rows)], ["법정 기록에 해당한 단위과제", checked], ["법정 최소 미달", short], ["확인 필요(설명 낱말로만 맞춤)", maybe], ["파기 기한(상한)보다 긴 과제 — 확인 필요", over],
                ["전체 목록 후보 중 법정 기간이 더 긴 것(판정 아님)", cand],
                ["호별·위임 추적 후보 중 법정 기간이 더 긴 것(판정 아님)", icand],
                ["확인일", "law.go.kr 2026-11-21 원문(조문은 시행일 열 참조)"],
                ["주의", "법정 기간은 최소다. 적용 대상·기산일을 확인한다. 공공기록물 보존기간 구분(1·3·5·10·30년·준영구·영구)으로 올려 맞추고, "
                         "법정 기간이 지나도 공공기록물법 평가·폐기 절차로만 폐기한다. 법령이 바뀌면 records-qa-gate 근거 변화 감시가 알린다."]):
        s.append(row)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"법정 보존기간 {len(rows)}종 · 해당 단위과제 {checked} · 법정 최소 미달 {short} · 확인 필요 {maybe} · 파기 상한 초과 {over} · 전체 목록 후보(더 김) {cand} → {a.out}")
    return 1 if short else 0


if __name__ == "__main__":
    sys.exit(main())
