#!/usr/bin/env python3
"""보존포맷 변환: 문서·표·발표 파일을 PDF/A-1b로 바꾸고, 변환 결과를 점검·기록한다.

사용:  python to_pdfa.py 원본폴더/ 산출/보존포맷/ [--verapdf]
근거:  NAK 30:2022(v1.1) 문서유형 전자기록물 보존포맷 기술규격 — PDF/A-1b 기반,
       NAK 37:2022(v1.0) 전자기록물 보존포맷 선정기준
방식:  LibreOffice headless 변환(SelectPdfVersion=1 → PDF/A-1) 뒤
       ① XMP의 pdfaid:part=1, conformance=B 선언 ② 모든 글꼴 임베딩 ③ 암호화 없음 을 점검하고,
       veraPDF가 설치돼 있으면 공식 적합성 검증까지 한다(최종 판정은 veraPDF).
       변환 뒤 문서정보(Info)의 작성·수정 일시를 XMP 값에 맞춘다(PDF/A-1b 6.7.3 — LibreOffice는 원본 작성일을 XMP에,
       변환 시각을 Info에 넣어 어긋난다. --no-fix-dates로 끔). Type1(CFF) 글꼴 부분집합에 CharSet이 없으면(6.3.5) 알린다
       — LibreOffice가 OTF/CFF 한글 글꼴(Noto CJK 등)을 이렇게 넣으므로 TrueType 한글 글꼴(맑은 고딕·나눔고딕 TTF 등)을 쓴다.
검증:  2026-11-20 veraPDF 1.30.4(greenfield)로 확인 — 한글 TrueType 글꼴 + 날짜 맞춤이면 PDF/A-1b PASS,
       Noto CJK(OTF) 대체 글꼴이면 6.3.5-2 FAIL, 날짜를 맞추지 않으면 6.7.3-1 FAIL.
기록:  변환이력.csv — 원본·결과의 SHA-256, 점검 결과, KST 시각(진본성 입증 자료)
원칙:  원본은 읽기만 하고, 결과 폴더에 같은 이름이 있으면 덮어쓰지 않는다.
한계:  HWP/HWPX는 LibreOffice로 서식이 보장되지 않아 '한컴 변환 필요'로 표시만 한다. 이미 PDF인 파일은 변환하지 않고 점검만 한다.
"""
import argparse
import csv
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

KST = dt.timezone(dt.timedelta(hours=9))
FILTERS = {
    "writer_pdf_Export": {".doc", ".docx", ".odt", ".rtf", ".txt", ".html", ".htm"},
    "calc_pdf_Export": {".xls", ".xlsx", ".ods", ".csv"},
    "impress_pdf_Export": {".ppt", ".pptx", ".odp"},
}
HANCOM = {".hwp", ".hwpx", ".hwt"}


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def convert(src, outdir, flt):
    opts = json.dumps({"SelectPdfVersion": {"type": "long", "value": "1"}})
    with tempfile.TemporaryDirectory() as profile:  # 사용자 설정과 섞이지 않게 임시 프로필
        cmd = [soffice(), f"-env:UserInstallation=file://{profile}", "--headless", "--norestore",
               "--convert-to", f"pdf:{flt}:{opts}", "--outdir", outdir, src]
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    return out if os.path.exists(out) else None, (p.stdout + p.stderr).strip()[-200:]


def check_pdf(path):
    """(PDF/A 선언, 글꼴 임베딩, 암호화) 점검 결과를 돌려준다."""
    from pypdf import PdfReader
    r = PdfReader(path)
    if r.is_encrypted:
        return "", "확인 불가", "암호화됨"
    meta = r.trailer["/Root"].get("/Metadata")
    xmp = meta.get_object().get_data().decode("utf-8", "replace") if meta else ""
    part = re.search(r"pdfaid:part(?:>|=\")\s*(\d)", xmp)
    conf = re.search(r"pdfaid:conformance(?:>|=\")\s*([ABUab])", xmp)
    decl = f"PDF/A-{part.group(1)}{conf.group(1).lower()}" if part and conf else "선언 없음"
    missing, nocharset = set(), set()
    for page in r.pages:
        fonts = (page.get("/Resources") or {}).get("/Font") or {}
        for ref in fonts.values():
            f = ref.get_object()
            fd = f.get("/FontDescriptor")
            if fd is None and "/DescendantFonts" in f:
                fd = f["/DescendantFonts"][0].get_object().get("/FontDescriptor")
            fd = fd.get_object() if fd is not None else {}
            name = str(f.get("/BaseFont", "?"))
            if not any(k in fd for k in ("/FontFile", "/FontFile2", "/FontFile3")) and f.get("/Subtype") != "/Type3":
                missing.add(name)
            elif f.get("/Subtype") == "/Type1" and "+" in name and "/CharSet" not in fd:
                nocharset.add(name)
    if missing:
        font = "미임베딩: " + ", ".join(sorted(missing))[:120]
    elif nocharset:
        font = "CharSet 없음(6.3.5): " + ", ".join(sorted(nocharset))[:120]
    else:
        font = "모두 임베딩"
    return decl, font, "없음"


def fix_dates(path):
    """Info의 CreationDate·ModDate를 XMP CreateDate·ModifyDate에 맞춘다(PDF/A-1 6.7.3). 고쳤으면 설명을 돌려준다."""
    try:
        import pikepdf
        from pikepdf.models.metadata import encode_pdf_date, decode_pdf_date
    except ImportError:
        return "pikepdf 없음 — 날짜 맞춤 못 함"
    changed = []
    with pikepdf.open(path, allow_overwriting_input=True) as pdf:
        with pdf.open_metadata(set_pikepdf_as_editor=False, update_docinfo=False) as meta:
            for xk, ik in (("xmp:CreateDate", "/CreationDate"), ("xmp:ModifyDate", "/ModDate")):
                xv = meta.get(xk)
                if not xv or ik not in pdf.docinfo:
                    continue
                try:
                    want = dt.datetime.fromisoformat(str(xv).replace("Z", "+00:00"))
                except ValueError:
                    want = None  # 원본에 날짜가 없어 XMP가 0000년 등으로 들어간 경우
                try:
                    have = decode_pdf_date(str(pdf.docinfo[ik]).rstrip("'"))
                except Exception:
                    have = None
                if want is None and have is not None:
                    meta[xk] = have.isoformat()  # XMP를 Info 값으로
                    changed.append(xk)
                    for other in ("xmp:ModifyDate", "xmp:MetadataDate"):  # 같은 0000년 값이 남지 않게
                        if str(meta.get(other, "")).startswith("0000"):
                            meta[other] = have.isoformat()
                elif want is not None and have != want:
                    pdf.docinfo[ik] = encode_pdf_date(want)
                    changed.append(ik.strip("/"))
        if changed:
            pdf.save(path)
    return ("날짜 맞춤(" + "·".join(changed) + ")") if changed else ""


def verapdf(path):
    """veraPDF PDF/A-1b 검증. 한글 파일명을 Java가 못 읽는 환경이 있어 영문 이름 임시 사본으로 검증한다."""
    exe = shutil.which("verapdf")
    if not exe:
        return "미설치"
    env = dict(os.environ, LC_ALL="C.UTF-8", LANG="C.UTF-8")
    with tempfile.TemporaryDirectory() as td:
        tmp = os.path.join(td, "check.pdf")
        shutil.copyfile(path, tmp)
        p = subprocess.run([exe, "--flavour", "1b", "--format", "text", tmp], capture_output=True, text=True, timeout=600, env=env)
        if re.search(r"^PASS\b", p.stdout, re.M):
            return "적합"
        if re.search(r"^FAIL\b", p.stdout, re.M):
            x = subprocess.run([exe, "--flavour", "1b", "--format", "xml", tmp], capture_output=True, text=True, timeout=600, env=env).stdout
            rules = sorted({f"{c}-{t}" for c, t in re.findall(r'<rule [^>]*clause="([^"]+)"[^>]*testNumber="(\d+)"[^>]*status="failed"', x)}
                           | {f"{c}-{t}" for t, c in re.findall(r'<rule [^>]*testNumber="(\d+)"[^>]*clause="([^"]+)"[^>]*status="failed"', x)})
            return "부적합" + (f"({', '.join(rules[:6])})" if rules else "")
    err = [l for l in (p.stdout + p.stderr).splitlines() if l.strip() and "JAVA_OPTIONS" not in l and "_JAVA_" not in l]
    return "검증 오류: " + (err[-1][:80] if err else f"종료 코드 {p.returncode}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--verapdf", action="store_true", help="veraPDF로 공식 검증(설치돼 있어야 함)")
    ap.add_argument("--no-fix-dates", action="store_true", help="Info 날짜를 XMP에 맞추지 않는다")
    a = ap.parse_args()
    if not soffice():
        print("LibreOffice(soffice)가 없습니다. 설치 후 다시 실행하세요.", file=sys.stderr)
        return 2
    os.makedirs(a.out, exist_ok=True)
    log, n_ok, n_bad = [], 0, 0
    for root, _, fs in os.walk(a.src):
        for fn in sorted(fs):
            src = os.path.join(root, fn)
            rel = os.path.relpath(src, a.src)
            ext = os.path.splitext(fn)[1].lower()
            outdir = os.path.join(a.out, os.path.dirname(rel))
            os.makedirs(outdir, exist_ok=True)
            now = dt.datetime.now(KST).isoformat(timespec="seconds")
            row = {"원본": rel, "원본 sha256": sha256(src), "결과": "", "결과 sha256": "", "PDF/A 선언": "",
                   "글꼴": "", "암호화": "", "veraPDF": "", "판정": "", "비고": "", "시각(KST)": now}
            if ext in HANCOM:
                row.update(판정="보류", 비고="한컴 변환 필요 — 한글에서 PDF/A로 저장하거나 기관 변환 도구 사용")
            elif ext == ".pdf":
                decl, font, enc = check_pdf(src)
                row.update({"PDF/A 선언": decl, "글꼴": font, "암호화": enc, "판정": "점검만", "비고": "이미 PDF — 변환하지 않음"})
                if a.verapdf:
                    row["veraPDF"] = verapdf(src)
            else:
                flt = next((f for f, exts in FILTERS.items() if ext in exts), None)
                target = os.path.join(outdir, os.path.splitext(fn)[0] + ".pdf")
                if flt is None:
                    row.update(판정="대상 아님", 비고="문서·표·발표 파일이 아님")
                elif os.path.exists(target):
                    row.update(판정="건너뜀", 비고="결과 파일이 이미 있음(덮어쓰지 않음)")
                else:
                    out, msg = convert(src, outdir, flt)
                    if not out:
                        row.update(판정="실패", 비고=msg)
                    else:
                        note = "" if a.no_fix_dates else fix_dates(out)
                        decl, font, enc = check_pdf(out)
                        row.update({"결과": os.path.relpath(out, a.out), "결과 sha256": sha256(out),
                                    "PDF/A 선언": decl, "글꼴": font, "암호화": enc})
                        if a.verapdf:
                            row["veraPDF"] = verapdf(out)
                        good = decl == "PDF/A-1b" and font == "모두 임베딩" and enc == "없음" and (not a.verapdf or row["veraPDF"] == "적합")
                        row["판정"] = "통과" if good else "확인 필요"
                        tips = [note] if note else []
                        if font.startswith("CharSet 없음") or "6.3.5" in row["veraPDF"]:
                            tips.append("OTF/CFF 대체 글꼴 — TrueType 한글 글꼴(맑은 고딕·나눔고딕 TTF 등)을 설치·지정하고 다시 변환")
                        if row["veraPDF"].startswith("검증 오류"):
                            tips.append("veraPDF 실행 문제 — 결과 판정 아님")
                        row["비고"] = " · ".join(tips)
            n_ok += row["판정"] == "통과"
            n_bad += row["판정"] in ("실패", "확인 필요")
            log.append(row)
    path = os.path.join(a.out, "변환이력.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(log[0].keys()) if log else ["원본"])
        w.writeheader()
        w.writerows(log)
    print(f"{len(log)}개 · 통과 {n_ok} · 실패·확인 필요 {n_bad} → {path}")
    return 1 if n_bad else 0


if __name__ == "__main__":
    sys.exit(main())
