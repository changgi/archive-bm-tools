#!/usr/bin/env python3
"""파일 포맷 식별: 확장자가 아니라 내용으로 포맷을 판정하고, 확장자와 어긋나거나 보존에 위험한 파일을 표시한다.

사용:  python identify.py 폴더/ -o 산출/포맷식별.csv
판정 순서
  1. Siegfried(sf)가 설치돼 있으면 PRONOM ID(예: fmt/354 = PDF/A-1b)를 함께 기록
  2. Magika(딥러닝 식별, AIC와 같은 엔진)가 있으면 그 결과
  3. 둘 다 없으면 파일 첫 바이트(시그니처)로 최소 판정
위험 표시: 확장자 불일치, 실행 파일, 암호화·손상 의심(0바이트), 장기보존 검토가 필요한 포맷
NAK 37:2025(v2.0) 「전자기록물 보존포맷 선정기준」 부속서 평가결과(국가기록원, 2026-11-20 내려받아 확인)로 '보존포맷·수용가능포맷·권고 외'를 붙인다.
  (디지털화 결과물의 JPEG 등은 NAK 26 디지털화 기준을 따로 본다 — 여기 표시는 전자기록물 보존포맷 관점)
"""
import argparse
import csv
import json
import os
import shutil
import subprocess
import sys

SIGS = [
    (b"%PDF", "pdf"), (b"PK\x03\x04", "zip-container"), (b"\xd0\xcf\x11\xe0", "ole2(hwp·doc·xls)"),
    (b"\xff\xd8\xff", "jpeg"), (b"\x89PNG", "png"), (b"II*\x00", "tiff"), (b"MM\x00*", "tiff"),
    (b"GIF8", "gif"), (b"MZ", "exe"), (b"RIFF", "riff(wav·avi)"), (b"ID3", "mp3"),
]
ZIP_FAMILY = {"hwpx", "docx", "xlsx", "pptx", "odt", "ods", "zip"}
REVIEW = {"hwp": "독자 포맷 — 보존포맷 변환 검토", "doc": "구형 오피스", "xls": "구형 오피스", "ppt": "구형 오피스",
          "gif": "색 제한", "bmp": "비압축 독자", "exe": "실행 파일 — 반입 금지 검토"}


NAK37 = {  # 확장자 → (유형, 등급)
    **{e: ("텍스트", "보존포맷") for e in ("pdf", "docx", "odt", "epub")}, "hwpx": ("텍스트", "수용가능포맷"),
    **{e: ("프리젠테이션", "보존포맷") for e in ("pptx", "odp")}, **{e: ("스프레드시트", "보존포맷") for e in ("xlsx", "ods")},
    **{e: ("이미지", "보존포맷") for e in ("tif", "tiff", "png")},
    **{e: ("오디오", "보존포맷") for e in ("wav", "bwf", "mp3", "aac", "m4a")}, **{e: ("오디오", "수용가능포맷") for e in ("flac", "wma")},
    **{e: ("비디오(컨테이너)", "보존포맷") for e in ("mpg", "mpeg", "mp4", "mxf")}, **{e: ("비디오(컨테이너)", "수용가능포맷") for e in ("avi", "mkv", "mov", "wmv")},
    **{e: ("데이터베이스·구조화", "보존포맷") for e in ("csv", "json", "xml")}, "siard": ("데이터베이스", "수용가능포맷"),
}
EVALUATED_NOT = {"hwp": "텍스트", "doc": "텍스트", "txt": "텍스트", "ppt": "프리젠테이션", "show": "프리젠테이션", "pps": "프리젠테이션", "pptm": "프리젠테이션",
                 "xls": "스프레드시트", "cell": "스프레드시트", "xlk": "스프레드시트", "asc": "구조화"}


def nak37(ext, path, sig):
    if ext == "pdf" and sig == "pdf":
        try:
            with open(path, "rb") as f:
                head = f.read(200000)
            m = __import__("re").search(rb"pdfaid:part[=>\"']+\s*(\d)", head)
            if m and m.group(1) in (b"1", b"2"):
                return f"보존포맷(텍스트, PDF/A-{m.group(1).decode()})"
        except OSError:
            pass
    if ext in NAK37:
        t, g = NAK37[ext]
        return f"{g}({t})"
    if ext in EVALUATED_NOT:
        return f"권고 외({EVALUATED_NOT[ext]} 평가 대상이었으나 권고 아님) — 보존포맷 변환 검토"
    return ""


DRM = [(rb"SCD000[2-4]", "소프트캠프"), (rb"Document SAFER", "마크애니"), (rb"(?i)fasoo", "파수")]  # 국가기록원 2018 연구 헤더 분석


def drm(path):
    try:
        with open(path, "rb") as f:
            b = f.read(4096)
    except OSError:
        return ""
    return next((v for rx, v in DRM if __import__("re").search(rx, b)), "")


def head_sig(path):
    try:
        with open(path, "rb") as f:
            b = f.read(16)
    except OSError:
        return "읽기 실패"
    if not b:
        return "빈 파일"
    for sig, name in SIGS:
        if b.startswith(sig):
            return name
    return "알 수 없음"


def siegfried(folder):
    if not shutil.which("sf"):
        return {}
    try:
        out = subprocess.run(["sf", "-json", folder], capture_output=True, text=True, timeout=3600).stdout
        res = {}
        for f in json.loads(out).get("files", []):
            m = (f.get("matches") or [{}])[0]
            res[os.path.abspath(f["filename"])] = (m.get("id", ""), m.get("format", ""))
        return res
    except Exception:
        return {}


def magika_fn():
    try:
        from magika import Magika
        m = Magika()
        return lambda p: m.identify_path(p).output.label
    except Exception:
        return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--out", default="산출/포맷식별.csv")
    a = ap.parse_args()
    sf = siegfried(a.folder)
    mg = magika_fn()
    rows, flagged = [], 0
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            p = os.path.join(root, fn)
            ext = os.path.splitext(fn)[1].lower().lstrip(".")
            sig = head_sig(p)
            label = mg(p) if mg else ""
            puid, fmt = sf.get(os.path.abspath(p), ("", ""))
            flags = []
            if sig == "빈 파일":
                flags.append("0바이트")
            if sig == "exe" or label in ("pebin", "elf"):
                flags.append("실행 파일")
            if ext in ZIP_FAMILY and sig not in ("zip-container", "빈 파일"):
                flags.append("확장자 불일치(zip 계열이 아님)")
            if ext == "pdf" and sig != "pdf":
                flags.append("확장자 불일치(PDF 아님)")
            if ext in ("jpg", "jpeg") and sig != "jpeg":
                flags.append("확장자 불일치(JPEG 아님)")
            d = drm(p)
            if d:
                flags.append(f"문서 DRM 의심({d}) — 내용 열람 불가, 이관 전 해제(transfer-schedule encrypted_scan.py 목록)")
            if ext in REVIEW:
                flags.append(REVIEW[ext])
            flagged += bool(flags)
            rows.append([os.path.relpath(p, a.folder), ext, sig, label, puid, fmt, nak37(ext, p, sig), " · ".join(flags)])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["파일", "확장자", "시그니처", "Magika", "PRONOM ID", "Siegfried 포맷", "NAK 37 권고(2025)", "표시"])
        w.writerows(rows)
    engines = ", ".join(e for e, ok in (("Siegfried", bool(sf)), ("Magika", bool(mg))) if ok) or "시그니처만"
    print(f"{len(rows)}개 식별({engines}) · 표시 {flagged}건 → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
