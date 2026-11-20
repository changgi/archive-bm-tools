#!/usr/bin/env python3
"""기록물 정리 점검: 등록정보와 실물 점검표를 대조해 시행규칙 제14조의 확인 사항을 표로 낸다.

사용:  python register_audit.py 등록정보.xlsx 실물점검표.xlsx [--key 철분류번호|등록번호] -o 산출/정리점검.xlsx
등록정보 열: 철분류번호, 등록번호, 제목, 생산연도, 보존기간, 단위과제, 쪽수, 공개구분, 비공개사유, 접근권한, 비밀구분, 기록물형태
실물점검표 열: 상자번호, 철분류번호(또는 등록번호), 실물쪽수, 등록번호표기(Y/N), 제목(선택), 비고
근거:  시행규칙 제14조① <개정 2026. 9. 11.> 정리 때 확인할 것
         1. 전자기록생산시스템에 등록되지 않은 기록물의 유무 → '미등록 실물'
         2. 등록정보와 실제 기록물 상태의 일치 여부        → '실물 없음', '쪽수 불일치'
         3. 공개여부, 접근권한 및 비밀여부의 적절성        → '공개·비밀 점검'
         4. 생산등록번호 또는 접수등록번호의 표기 유무     → '등록번호 표기 없음'
       시행규칙 제14조② <신설 2026. 9. 11.> 기록물철을 생산연도별·보존기간별로 구분해 보존상자에 넣는다 → '상자 혼재'
       시행령 제32조⑤ 비전자기록물은 단위과제별로 보존상자에 넣어 이관한다 → '상자 혼재(단위과제)'
       시행규칙 제9조③ 철당 100매 이내, 분권은 같은 제목·분류번호에 권호수만 다르게 → '철 분량'
       시행규칙 제6조·제8조② 등록정보 수정은 기록물관리책임자 확인, 수정일자·내용·사유를 전자적으로 관리 → '수정 이력'
       시행령 제24조① 다음 해 2월 말일까지 공개여부·접근권한 재분류, 분류·편철 확정
원칙:  점검 결과만 낸다. 등록정보 수정은 처리과가 시스템에서 한다. 비밀 표시 행은 제목을 옮기지 않는다.
"""
import argparse
import collections
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "pii-redact", "scripts"))
try:
    from pii_scan import scan_text
except ImportError:
    scan_text = None

ALIAS = {"철분류번호": ["철분류번호", "기록물철분류번호", "분류번호", "철번호"], "등록번호": ["등록번호", "생산등록번호", "접수등록번호"],
         "제목": ["제목", "철제목", "기록물철제목"], "생산연도": ["생산연도", "생산년도"], "보존기간": ["보존기간"],
         "단위과제": ["단위과제", "단위과제명"], "쪽수": ["쪽수", "면수"], "실물쪽수": ["실물쪽수", "실제쪽수"],
         "공개구분": ["공개구분", "공개여부"], "비공개사유": ["비공개사유", "비공개 사유"], "접근권한": ["접근권한", "접근범위"],
         "비밀": ["비밀구분", "비밀여부"], "형태": ["기록물형태", "형태"], "상자": ["상자번호", "보존상자", "상자"],
         "표기": ["등록번호표기", "번호표기", "등록번호 표기"], "비고": ["비고"], "권호": ["권호수", "권호"],
         "수정일자": ["수정일자", "수정일"], "수정내용": ["수정내용"], "수정사유": ["수정사유"]}
SECRET_RX = re.compile(r"[ⅠⅡⅢ1-3]\s*급\s*비밀|대외비|비밀\s*(문서|기록)")
EMPTY = ("", "일반", "N", "아니오", "해당없음", "없음")


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
            if n in head:
                idx[k] = head.index(n)
                break
    return [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)], idx


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("register")
    ap.add_argument("physical")
    ap.add_argument("--key", choices=["철분류번호", "등록번호"], default="철분류번호")
    ap.add_argument("-o", "--out", default="산출/정리점검.xlsx")
    a = ap.parse_args()
    reg, ri = read(a.register)
    phy, pi = read(a.physical)
    if a.key not in ri or a.key not in pi:
        print(f"입력 오류: 두 표 모두 '{a.key}' 열이 필요합니다.")
        return 2
    K = a.key
    regmap = collections.defaultdict(list)
    for n, r in reg:
        regmap[r[K]].append((n, r))
    physmap = collections.defaultdict(list)
    for n, r in phy:
        physmap[r[K]].append((n, r))
    issues = []   # (구분, 근거, 키, 상자, 내용)

    def add(kind, basis, key, box, msg):
        issues.append((kind, basis, key, box, msg))

    def secret(r):
        return r.get("비밀", "") not in EMPTY or bool(SECRET_RX.search(r.get("제목", "")))

    for key, ps in physmap.items():
        box = ps[0][1].get("상자", "")
        if not key:
            add("미등록 실물", "제14조①1", "", box, f"실물점검표 {ps[0][0]}행: 번호가 없음")
            continue
        if key not in regmap:
            add("미등록 실물", "제14조①1", key, box, "실물은 있는데 등록정보에 없음")
            continue
        r = regmap[key][0][1]
        if "실물쪽수" in pi and "쪽수" in ri:
            reg_p = sum(int(x[1]["쪽수"]) for x in regmap[key] if x[1].get("쪽수", "").isdigit()) if K == "철분류번호" else (int(r["쪽수"]) if r.get("쪽수", "").isdigit() else None)
            phy_p = sum(int(x[1]["실물쪽수"]) for x in ps if x[1].get("실물쪽수", "").isdigit())
            if reg_p is not None and phy_p and reg_p != phy_p and "전자" != r.get("형태", ""):
                add("쪽수 불일치", "제14조①2", key, box, f"등록 {reg_p}쪽 · 실물 {phy_p}쪽")
        if "표기" in pi and any(x[1].get("표기", "").upper() in ("N", "아니오", "없음", "X") for x in ps):
            add("등록번호 표기 없음", "제14조①4", key, box, "실물에 생산·접수등록번호가 표기되지 않음")
    for key, rs in regmap.items():
        r = rs[0][1]
        if key and key not in physmap and r.get("형태", "") != "전자":
            add("실물 없음", "제14조①2", key, "", "등록정보에는 있는데 실물점검표에 없음")
        for n, x in rs:
            disc = x.get("공개구분", "").replace(" ", "")
            label = key if not secret(x) else f"{key} (비밀 표시)"
            if secret(x) and x.get("비밀", "") in EMPTY:
                add("공개·비밀 점검", "제14조①3", label, "", "제목에 비밀 표시가 있는데 비밀구분이 비어 있음")
            if "공개구분" in ri and disc not in ("공개", "부분공개", "비공개"):
                add("공개·비밀 점검", "제14조①3·시행령 제27조", label, "", f"공개구분 '{x.get('공개구분', '')}' — 건 단위 공개여부 미구분")
            if disc in ("비공개", "부분공개") and "비공개사유" in ri and not re.search(r"[1-8]\s*호", x.get("비공개사유", "")):
                add("공개·비밀 점검", "제14조①3·정보공개법 제9조①", label, "", "비공개인데 사유에 호(1~8호)가 없음")
            if disc == "공개" and scan_text and not secret(x):
                hits = scan_text(x.get("제목", ""))
                if hits:
                    add("공개·비밀 점검", "제14조①3·정보공개법 제9조①6호", label, "", "공개인데 제목에 개인정보 후보: " + ", ".join(h[0] for h in hits))
            if disc in ("비공개", "부분공개") and "접근권한" in ri and not x.get("접근권한", ""):
                add("공개·비밀 점검", "제14조①3", label, "", "비공개인데 접근권한이 비어 있음")
    # 철당 100매 이내(시행규칙 제9조③): 같은 철분류번호·권호수의 비전자 쪽수 합
    if "철분류번호" in ri and "쪽수" in ri:
        vol = collections.defaultdict(int)
        titles = collections.defaultdict(set)
        for n, x in reg:
            if x.get("형태", "") == "전자" or not x.get("철분류번호"):
                continue
            k = (x["철분류번호"], x.get("권호", ""))
            vol[k] += int(x["쪽수"]) if x.get("쪽수", "").isdigit() else 0
            if "권호" in ri and x.get("권호"):
                titles[x["철분류번호"]].add(re.sub(r"\s|\(\d+\)$", "", x.get("제목", "")) if K == "철분류번호" else "")
        for (no, v), p_ in sorted(vol.items()):
            if p_ > 100:
                add("철 분량", "시행규칙 제9조③", no + (f" ({v}권)" if v else ""), "", f"{p_}매 — 철당 100매 이내, 넘으면 같은 제목·분류번호로 분권하고 권호수만 다르게")
        for no, ts in titles.items():
            ts.discard("")
            if len(ts) > 1:
                add("철 분량", "시행규칙 제9조③", no, "", "같은 분류번호의 권마다 제목이 다름 — 분권은 같은 제목에 권호수만 다르게")
    # 등록정보 수정 이력(시행규칙 제6조·제8조②): 수정일자·수정내용·수정사유를 함께 관리
    trio = [c for c in ("수정일자", "수정내용", "수정사유") if c in ri]
    if trio:
        for n, x in reg:
            filled = [c for c in ("수정일자", "수정내용", "수정사유") if x.get(c)]
            if filled and len(filled) < 3:
                miss = [c for c in ("수정일자", "수정내용", "수정사유") if not x.get(c)]
                add("수정 이력", "시행규칙 제6조·제8조②", x.get(K, ""), "", "수정 기록에 " + "·".join(miss) + " 없음 — 기록물관리책임자 확인 아래 세 가지를 함께 남김")
    boxes = collections.defaultdict(list)
    for n, p in phy:
        if p.get("상자") and p[K] in regmap:
            boxes[p["상자"]].append(regmap[p[K]][0][1])
    for b, rs in sorted(boxes.items()):
        for col, label, basis in (("생산연도", "생산연도", "제14조②"), ("보존기간", "보존기간", "제14조②"), ("단위과제", "단위과제", "시행령 제32조⑤")):
            vals = sorted({x.get(col, "") for x in rs if x.get(col, "")})
            if len(vals) > 1:
                add("상자 혼재", basis, "", b, f"{label}{'가' if label == '단위과제' else '이'} 섞임: {', '.join(vals)}")

    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "요약"
    s.append(["기록물 정리 점검(시행규칙 제14조)"])
    s["A1"].font = Font(bold=True, size=13)
    s.append(["대조 기준", K])
    s.append(["등록정보", len(reg)])
    s.append(["실물점검표", len(phy)])
    c = collections.Counter(i[0] for i in issues)
    for k in ("미등록 실물", "실물 없음", "쪽수 불일치", "등록번호 표기 없음", "공개·비밀 점검", "상자 혼재", "철 분량", "수정 이력"):
        s.append([k, c.get(k, 0)])
    s.append([])
    s.append(["정리 기한", "다음 해 2월 말일까지 공개여부·접근권한 재분류, 분류·편철 확정(시행령 제24조①)"])
    d = wb.create_sheet("점검 결과")
    d.append(["구분", "근거", K, "상자번호", "내용"])
    for i in issues:
        d.append(list(i))
    for sh, wd in ((s, (22, 70)), (d, (18, 26, 26, 10, 60))):
        for j, x in enumerate(wd):
            sh.column_dimensions[chr(65 + j)].width = x
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"등록 {len(reg)} · 실물 {len(phy)} · " + " · ".join(f"{k} {v}" for k, v in c.items()) + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
