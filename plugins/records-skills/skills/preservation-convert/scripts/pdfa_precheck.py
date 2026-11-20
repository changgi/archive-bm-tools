#!/usr/bin/env python3
"""PDF/A-1b 사전 점검: veraPDF가 없을 때 NAK 30:2022(v1.1) 「문서유형 전자기록물 보존포맷 기술규격: PDF/A-1b 기반의 포맷」
6장 세부규격에서 자주 걸리는 항목을 pikepdf로 본다. 공식 적합성 판정이 아니다(최종은 veraPDF).
근거: NAK 30:2022 6.1 파일 구조(암호화·LZW·내장 파일·OCProperties), 6.2 그래픽(OutputIntent GTS_PDFA1), 6.3 폰트(모두 내장),
      6.4 투명도(SMask·Transparency 그룹·CA/ca<1 금지), 6.5 주석(FileAttachment·Sound·Movie 금지), 6.6 행위(JavaScript·Launch 등 금지),
      6.7 메타데이터(XMP pdfaid:part=1, conformance=B), 6.9 대화식 폼 (국가기록원, 2026-11-20 내려받아 확인)

사용:  python pdfa_precheck.py 보존포맷폴더/ -o 산출/PDFA사전점검.csv
원칙:  읽기만 한다. '사전 점검 통과'는 '적합'이 아니다 — veraPDF로 공식 검증한다.
"""
import argparse
import csv
import os
import sys

import pikepdf

BAD_ACTIONS = {"/JavaScript", "/Launch", "/Sound", "/Movie", "/ResetForm", "/ImportData", "/Hide", "/SetOCGState", "/Rendition", "/Trans", "/GoTo3DView"}
BAD_ANNOTS = {"/FileAttachment", "/Sound", "/Movie", "/Screen", "/3D", "/RichMedia"}


def walk_objects(pdf):
    for obj in pdf.objects:
        yield obj


def check(path):
    issues = []
    try:
        pdf = pikepdf.open(path)
    except pikepdf.PasswordError:
        return ["암호화(6.1.3 Encrypt 금지)"], ""
    with pdf:
        if pdf.is_encrypted:
            issues.append("암호화(6.1.3 Encrypt 금지)")
        root = pdf.Root
        version = str(pdf.pdf_version)
        # 메타데이터
        try:
            meta = pdf.open_metadata()
            part, conf = meta.get("pdfaid:part"), meta.get("pdfaid:conformance")
            if str(part) != "1" or str(conf).upper() != "B":
                issues.append(f"XMP PDF/A 선언 part={part} conformance={conf} — part=1, conformance=B(6.7)")
        except Exception:
            issues.append("XMP 메타데이터를 읽지 못함(6.7)")
        # OutputIntent
        oi = root.get("/OutputIntents")
        if not oi or not any(str(x.get("/S")) == "/GTS_PDFA1" and "/DestOutputProfile" in x for x in oi):
            issues.append("OutputIntent GTS_PDFA1(ICC 프로파일) 없음(6.2.2)")
        names = root.get("/Names")
        if names is not None and "/EmbeddedFiles" in names:
            issues.append("내장 파일 EmbeddedFiles(6.1.11)")
        if names is not None and "/JavaScript" in names:
            issues.append("문서 수준 JavaScript(6.6)")
        if "/OCProperties" in root:
            issues.append("선택적 콘텐츠 OCProperties(6.1.12)")
        if "/AA" in root:
            issues.append("문서 추가 행위 AA(6.6.2)")
        oa = root.get("/OpenAction")
        if isinstance(oa, pikepdf.Dictionary) and str(oa.get("/S")) in BAD_ACTIONS:
            issues.append(f"열 때 실행 행위 {oa.get('/S')}(6.6)")
        if "/AcroForm" in root and root.AcroForm.get("/NeedAppearances") is True:
            issues.append("대화식 폼 NeedAppearances true(6.9)")
        lzw = trans = act = 0
        for obj in walk_objects(pdf):
            if isinstance(obj, pikepdf.Stream):
                f = obj.get("/Filter")
                fl = [str(x) for x in f] if isinstance(f, pikepdf.Array) else [str(f)] if f is not None else []
                if "/LZWDecode" in fl:
                    lzw += 1
                if obj.get("/SMask") is not None and str(obj.get("/Subtype")) == "/Image":
                    trans += 1
            if isinstance(obj, (pikepdf.Dictionary, pikepdf.Stream)):
                g = obj.get("/Group")
                if isinstance(g, pikepdf.Dictionary) and str(g.get("/S")) == "/Transparency":
                    trans += 1
                if str(obj.get("/Type")) == "/ExtGState":
                    for k in ("/CA", "/ca"):
                        v = obj.get(k)
                        if v is not None and float(v) < 1.0:
                            trans += 1
                    sm = obj.get("/SMask")
                    if sm is not None and str(sm) != "/None":
                        trans += 1
                if str(obj.get("/Type")) == "/Action" or "/S" in obj and str(obj.get("/S")) in BAD_ACTIONS:
                    if str(obj.get("/S")) in BAD_ACTIONS:
                        act += 1
        if lzw:
            issues.append(f"LZWDecode 필터 {lzw}곳(6.1.10)")
        if trans:
            issues.append(f"투명도 사용 {trans}곳(6.4)")
        if act:
            issues.append(f"금지된 행위(JavaScript·Launch 등) {act}곳(6.6)")
        unemb, bad_annot, subset = set(), set(), set()
        for page in pdf.pages:
            res = page.obj.get("/Resources")
            fonts = res.get("/Font") if res is not None else None
            if fonts is not None:
                for _, font in fonts.items():
                    targets = [font]
                    if str(font.get("/Subtype")) == "/Type0" and "/DescendantFonts" in font:
                        targets = list(font.DescendantFonts)
                    for ft in targets:
                        if str(ft.get("/Subtype")) == "/Type3":
                            continue
                        fd = ft.get("/FontDescriptor")
                        if fd is None or not any(k in fd for k in ("/FontFile", "/FontFile2", "/FontFile3")):
                            unemb.add(str(font.get("/BaseFont", "?")))
                        elif "+" in str(ft.get("/BaseFont", "")):
                            st = str(ft.get("/Subtype"))
                            if st.startswith("/CIDFontType") and "/CIDSet" not in fd:
                                subset.add(f"{ft.get('/BaseFont')}(CIDSet)")
                            elif st in ("/Type1", "/MMType1") and "/CharSet" not in fd:
                                subset.add(f"{ft.get('/BaseFont')}(CharSet)")
            for an in page.obj.get("/Annots", []) or []:
                st = str(an.get("/Subtype"))
                if st in BAD_ANNOTS:
                    bad_annot.add(st)
        if unemb:
            issues.append("내장 안 된 폰트: " + ", ".join(sorted(unemb)[:5]) + "(6.3.4)")
        if subset:
            issues.append("부분집합 폰트에 CharSet/CIDSet 없음: " + ", ".join(sorted(subset)[:3]) + "(6.3.5) — veraPDF로 확인, TrueType 한글 글꼴로 다시 변환 권장")
        if bad_annot:
            issues.append("금지된 주석 " + ", ".join(sorted(bad_annot)) + "(6.5)")
        return issues, version


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--out", default="산출/PDFA사전점검.csv")
    a = ap.parse_args()
    rows = []
    paths = [a.folder] if os.path.isfile(a.folder) else [os.path.join(d, f) for d, _, fs in os.walk(a.folder) for f in sorted(fs) if f.lower().endswith(".pdf")]
    for p in paths:
        try:
            iss, ver = check(p)
        except Exception as e:
            iss, ver = [f"읽기 실패: {str(e)[:80]}"], ""
        rows.append([os.path.relpath(p, a.folder) if os.path.isdir(a.folder) else os.path.basename(p), ver, len(iss), " · ".join(iss),
                     "사전 점검 통과(공식 검증 전)" if not iss else "고칠 것 있음"])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["파일", "PDF 버전", "문제 수", "문제(NAK 30 조항)", "판정"])
        w.writerows(rows)
    bad = sum(1 for r in rows if r[2])
    print(f"{len(rows)}개 · 고칠 것 있음 {bad} · 사전 점검 통과 {len(rows) - bad}(veraPDF 공식 검증 전) → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
