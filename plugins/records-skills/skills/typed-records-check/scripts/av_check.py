#!/usr/bin/env python3
"""시청각기록물 등록 점검: 등록 목록의 유형·쪽수·내용요약(생산정보)이 국가기록원 지침대로 채워졌는지 본다.

사용:  python av_check.py 시청각등록목록.xlsx [--files 원본폴더/] -o 산출/시청각_점검.xlsx
목록 열(머리글 이름으로 찾음): 등록번호, 제목, 유형(등록구분), 기록물형태, 쪽수, 내용요약, 등록일자, 과제카드명, 공개여부
근거:  공공기록물법 제23조, 시행령 제19조(시청각기록물 생산), 시행규칙 제4조 제3·4호
       국가기록원 「2026년 기록물관리 지침」 시청각기록물 등록:
         유형은 '사진·필름류' 또는 '녹음·동영상류'로 선택해야 기록물형태·내용요약 칸이 생긴다
         등록일자(시행일자)는 촬영일자와 다르므로 내용요약에 생산일자를 반드시 넣는다
         쪽수는 디지털파일이면 파일 개수, 아날로그면 매·롤·점 수
         내용요약에는 생산일자, 주요내용, 장소, 주요 인물 성명·직위(및 위치정보)를 넣는다
         동영상은 개요·시간별 촬영 세부사항을 담은 설명문을 따로 작성한다
       기록관리 평가지표 '시청각기록물 관리의 적절성'은 맥락정보 포함 여부를 본다.
원칙:  표시를 찾는 점검이다. 인물·장소가 맞는지는 사람이 확인한다. --files를 주면 등록번호로 시작하는 파일 수와 쪽수를 비교한다.
"""
import argparse
import os
import re
import sys

ALIAS = {"번호": ["등록번호", "생산등록번호"], "제목": ["제목", "건제목"], "유형": ["유형", "등록구분", "기록물유형", "유형구분"],
         "형태": ["기록물형태", "매체형태"], "쪽수": ["쪽수", "수량", "파일수"], "요약": ["내용요약"], "등록일": ["등록일자", "생산(접수)등록일시", "등록일시"],
         "과제": ["과제카드명", "단위과제명"], "공개": ["공개여부", "공개구분"]}
DATE = re.compile(r"(19|20)\d{2}\s*[.\-/년]\s*\d{1,2}\s*[.\-/월]\s*\d{1,2}")
PLACE = re.compile(r"장소|에서|회의실|청사|강당|센터|현장|[가-힣]+(시|군|구|동|읍|면|로|길)\b")
PERSON = re.compile(r"(장관|차관|청장|시장|군수|구청장|도지사|교육감|위원장|원장|처장|국장|과장|실장|대표|의원|총장|대통령|총리)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--files", help="원본 파일 폴더(파일명이 등록번호로 시작)")
    ap.add_argument("-o", "--out", default="산출/시청각_점검.xlsx")
    a = ap.parse_args()
    import openpyxl
    ws0 = openpyxl.load_workbook(a.listing, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws0.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 3), 0)
    head = rows[hi]
    idx = {}
    for k, al in ALIAS.items():
        for n in al:
            if n in head:
                idx[k] = head.index(n)
                break
    files = []
    if a.files:
        for root, _, fs in os.walk(a.files):
            files += fs
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검"
    ws.append(["원본 행", "등록번호", "제목", "유형", "생산일자", "장소", "주요 인물", "쪽수 확인", "판정"])
    bad = 0
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        if not any(r):
            continue
        g = lambda k: r[idx[k]] if k in idx and idx[k] < len(r) else ""
        typ, summ = g("유형"), g("요약")
        issues = []
        tnorm = re.sub(r"[\s·/ㆍ]", "", typ)
        if tnorm not in ("사진필름류", "녹음동영상류"):
            issues.append(f"유형 '{typ}' — 사진·필름류 / 녹음·동영상류 중 하나로")
        d = "○" if DATE.search(summ) else "✕"
        pl = "○" if PLACE.search(summ) else "✕"
        pe = "○" if PERSON.search(summ) else "△"
        if not summ:
            issues.append("내용요약 없음(생산정보 필수)")
        else:
            if d == "✕":
                issues.append("내용요약에 생산(촬영)일자 없음")
            if pl == "✕":
                issues.append("내용요약에 장소 없음")
            if pe == "△":
                issues.append("주요 인물 성명·직위 표시를 못 찾음(인물이 없으면 무시)")
        cnt = ""
        if a.files and g("번호"):
            k = sum(1 for f in files if f.startswith(g("번호")))
            if g("쪽수").isdigit() and k and int(g("쪽수")) != k:
                cnt = f"쪽수 {g('쪽수')} ≠ 파일 {k}"
                issues.append(cnt)
            elif not k:
                cnt = "파일 못 찾음"
            else:
                cnt = "일치"
        if "녹음동영상" in tnorm and not re.search(r"설명문|촬영\s*세부|시간별", summ):
            issues.append("동영상: 개요·시간별 촬영 세부사항 설명문을 따로 작성했는지 확인")
        bad += bool([x for x in issues if "무시" not in x and "확인" not in x])
        ws.append([n, g("번호"), g("제목"), typ, d, pl, pe, cnt, " · ".join(issues) or "통과"])
    for i, w in enumerate((8, 16, 30, 14, 8, 8, 8, 16, 70)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"시청각 {ws.max_row - 1}건 · 고칠 것 있는 건 {bad} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
