#!/usr/bin/env python3
"""스캔 품질 점검: 스캔 이미지의 해상도(DPI)·기울기·빈 면·밝기·크기를 점검해 재스캔 후보를 표시한다.

사용:  python scan_qc.py 스캔폴더/ [--min-dpi 300] [--max-skew 1.0] -o 산출/스캔품질.csv
대상:  .tif .tiff .jpg .jpeg .png .bmp (PDF는 이미지로 뽑은 뒤 점검)
기준값은 기관의 스캔 지침에 맞게 바꾼다(기본 300dpi·기울기 1도).
NAK 26:2023(v2.1) 「기록물 디지털화 기준」: --kind 문서|사진|필름|항공필름|박물 --quality FHD|2K|4K|8K 를 주면
  최소 해상도를 표 1·3·5·7·12 값으로 쓰고(문서 165/175/330, 인화사진 320/345/640, 사진필름 1360/1450/2710,
  항공필름 4K 2710·8K 5420, 행정박물 2D 촬영 600ppi), 비트심도(흑백 1비트 금지 — 최소 그레이 8비트, 컬러 24비트),
  파일포맷·압축(TIFF 무압축·LZW·RLE·JPEG / JPEG / JPEG2000 / 문서는 PDF-A), JPEG 품질 70% 이상(양자화표로 추정)을 본다.
  --doc-* 옵션으로 디지털화 과정 기록화 항목(4.2.2: 목적·대상·작업일·수행자·장치·소프트웨어·압축·해시)을 이력에 넣는다.
결과의 '판정'은 후보일 뿐이고, 재스캔 여부는 사람이 이미지를 보고 정한다.
"""
import argparse
import csv
import os
import sys

import numpy as np
from PIL import Image

EXTS = {".tif", ".tiff", ".jpg", ".jpeg", ".png", ".bmp", ".jp2"}
MASTER = {".tif", ".tiff", ".png", ".jp2"}  # 보존용 마스터 권장 포맷(TIFF·PNG·JPEG 2000, 문서형은 PDF/A)
Image.MAX_IMAGE_PIXELS = 400_000_000
# NAK 26:2023 최소 해상도(ppi) — 품질 단계별
NAK26_PPI = {"문서": {"FHD": 165, "2K": 175, "4K": 330}, "사진": {"FHD": 320, "2K": 345, "4K": 640},
             "필름": {"FHD": 1360, "2K": 1450, "4K": 2710}, "항공필름": {"4K": 2710, "8K": 5420}, "박물": {"FHD": 600, "2K": 600, "4K": 600}}
NAK26_FMT = {"문서": {".tif", ".tiff", ".jpg", ".jpeg", ".jp2", ".pdf"}, "사진": {".tif", ".tiff", ".jpg", ".jpeg"},
             "필름": {".tif", ".tiff", ".jpg", ".jpeg"}, "항공필름": {".tif", ".tiff", ".jpg", ".jpeg"}, "박물": {".tif", ".tiff", ".jpg", ".jpeg"}}
TIFF_OK = {"raw": "무압축", "tiff_lzw": "무손실(LZW)", "packbits": "무손실(RLE)", "tiff_jpeg": "손실(JPEG)", "jpeg": "손실(JPEG)"}
STD_LUMA = [16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55, 14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62,
            18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92, 49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99]


def jpeg_quality(im):
    """libjpeg 방식 양자화표에서 품질(%)을 거꾸로 추정한다. 추정값이므로 참고용."""
    q = getattr(im, "quantization", None)
    if not q or 0 not in q:
        return None
    t = list(q[0])
    if len(t) != 64:
        return None
    scale = sum(t) * 100.0 / sum(STD_LUMA)
    qual = (200 - scale) / 2 if scale <= 100 else 5000 / scale
    return max(1, min(100, round(qual)))


def nak26(im, path, kind, quality):
    """(비트심도, 압축, JPEG 품질, 표시 목록)"""
    flags, soft = [], []
    mode = im.mode
    bits = {"1": "1bit", "L": "8bit 그레이", "P": "8bit 팔레트", "RGB": "24bit 컬러", "RGBA": "32bit(RGBA)", "CMYK": "32bit CMYK",
            "I;16": "16bit 그레이", "I;16B": "16bit 그레이", "I": "32bit", "LA": "8bit 그레이+알파"}.get(mode, mode)
    ext = os.path.splitext(path)[1].lower()
    comp = ""
    if ext in (".tif", ".tiff"):
        c = im.info.get("compression", "raw")
        comp = TIFF_OK.get(c, c)
    elif ext in (".jpg", ".jpeg"):
        comp = "손실(JPEG)"
    elif ext == ".jp2":
        comp = "JPEG2000"
    jq = jpeg_quality(im) if ext in (".jpg", ".jpeg") or comp == "손실(JPEG)" else None
    if kind:
        if mode == "1":
            flags.append("1비트 흑백 — 최소 그레이 8비트(NAK 26)")
        elif mode == "P":
            flags.append("팔레트 8비트 — 그레이 8비트 또는 컬러 24비트로")
        if ext not in NAK26_FMT[kind]:
            soft.append(f"{ext} 포맷 — NAK 26 목록 밖(국제표준 인정 포맷이면 가능, 5 ④)")
        if ext in (".tif", ".tiff") and comp not in TIFF_OK.values():
            soft.append(f"TIFF 압축 {comp} — 목록 밖(무압축·LZW·RLE·JPEG)")
        if jq is not None and jq < 70:
            flags.append(f"JPEG 품질 약 {jq}% < 70%")
    return bits, comp, jq, flags, soft


def estimate_skew(gray: np.ndarray, max_angle=5.0, step=0.25) -> float:
    """글줄 방향으로 투영했을 때 행 합의 분산이 가장 큰 각도를 기울기로 본다."""
    img = Image.fromarray(gray)
    if img.width > 1000:
        img = img.resize((1000, int(img.height * 1000 / img.width)))
    ink = (np.asarray(img) < 128).astype(np.uint8) * 255
    base = Image.fromarray(ink)
    best, best_score = 0.0, -1.0
    for ang in np.arange(-max_angle, max_angle + 1e-9, step):
        prof = np.asarray(base.rotate(ang, resample=Image.NEAREST, fillcolor=0)).sum(axis=1).astype(float)
        score = prof.var()
        if score > best_score:
            best, best_score = float(ang), score
    return -best  # 이미지를 best만큼 돌리면 바로 선다 → 원래 기울기는 반대 부호


def check(path, min_dpi, max_skew, kind=None, quality=None):
    im = Image.open(path)
    bits, comp, jq, nflags, soft = nak26(im, path, kind, quality)
    dpi = im.info.get("dpi")
    dpi = round(float(dpi[0])) if dpi else None
    gray = np.asarray(im.convert("L"))
    h, w = gray.shape
    mean, std = float(gray.mean()), float(gray.std())
    dark_ratio = float((gray < 128).mean())
    flags = []
    if dpi is None:
        flags.append("DPI 정보 없음")
    elif dpi < min_dpi:
        flags.append(f"DPI {dpi} < {min_dpi}" + (f"(NAK 26 {kind} {quality})" if kind else ""))
    flags += nflags
    if dark_ratio < 0.002 or std < 3:
        flags.append("빈 면 의심")
        skew = 0.0
    else:
        skew = estimate_skew(gray)
        if abs(skew) > max_skew:
            flags.append(f"기울기 {skew:+.1f}°")
    if mean < 60:
        flags.append("너무 어두움")
    if min(w, h) < 1000:
        flags.append("화소 수 적음")
    master = "적합" if os.path.splitext(path)[1].lower() in MASTER else "이용본 포맷(마스터는 TIFF·PNG·JPEG 2000 권장)"
    return [w, h, dpi or "", im.mode, round(mean, 1), round(skew, 2), " · ".join(flags + soft), master, bits, comp, "" if jq is None else jq,
            "재스캔 후보" if flags else "포맷 확인(변환 후보)" if soft else "통과"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--min-dpi", type=int, default=300)
    ap.add_argument("--max-skew", type=float, default=1.0)
    ap.add_argument("--kind", choices=list(NAK26_PPI), help="NAK 26 기준 적용: 문서(지도·도면 포함)·사진(인화사진)·필름(사진필름)·항공필름·박물(행정박물 2D 촬영)")
    ap.add_argument("--quality", default="", help="NAK 26 품질 단계 FHD·2K·4K(항공필름은 4K·8K) — 기본: 가장 낮은 단계")
    ap.add_argument("--doc-purpose", default="", help="디지털화 목적(4.2.2)")
    ap.add_argument("--doc-operator", default="", help="담당자/작업자")
    ap.add_argument("--doc-device", default="", help="장치(모델·제조사)")
    ap.add_argument("--doc-software", default="", help="소프트웨어")
    ap.add_argument("-o", "--out", default="산출/스캔품질.csv")
    ap.add_argument("--history", help="전자화 이력 CSV도 쓴다(시행령 제29조③: 해시값·변환 시점·변환 과정)")
    ap.add_argument("--process", default="", help="변환 과정 설명(장비·소프트웨어·해상도·담당 등) — 이력에 그대로 적는다")
    a = ap.parse_args()
    if a.kind:
        table = NAK26_PPI[a.kind]
        if not a.quality:
            a.quality = next(iter(table))
        if a.quality not in table:
            print(f"--quality는 {', '.join(table)} 중 하나", file=sys.stderr)
            return 2
        a.min_dpi = table[a.quality]
    rows = []
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            if os.path.splitext(fn)[1].lower() not in EXTS:
                continue
            p = os.path.join(root, fn)
            try:
                rows.append([os.path.relpath(p, a.folder), *check(p, a.min_dpi, a.max_skew, a.kind, a.quality)])
            except Exception as e:
                rows.append([os.path.relpath(p, a.folder), "", "", "", "", "", "", f"읽기 실패: {str(e)[:80]}", "", "", "", "", "확인 필요"])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        wr = csv.writer(f)
        wr.writerow(["파일", "가로(px)", "세로(px)", "DPI", "색 방식", "평균 밝기", "기울기(°)", "표시", "보존 포맷", "비트심도", "압축", "JPEG 품질(추정 %)", "판정"])
        wr.writerows(rows)
    if a.history:
        import datetime as _dt
        import hashlib
        kst = _dt.timezone(_dt.timedelta(hours=9))
        with open(a.history, "w", encoding="utf-8-sig", newline="") as f:
            wr = csv.writer(f)
            wr.writerow(["파일", "해시함수", "SHA-256", "파일 수정 시각(KST)", "이력 작성 시각(KST)", "변환 과정", "디지털화 목적", "담당자·작업자",
                         "장치", "소프트웨어", "압축방법", "적용 기준", "품질 판정", "행정전자서명"])
            now = _dt.datetime.now(kst).isoformat(timespec="seconds")
            for r in rows:
                fp = os.path.join(a.folder, r[0])
                h = hashlib.sha256(open(fp, "rb").read()).hexdigest()
                mt = _dt.datetime.fromtimestamp(os.path.getmtime(fp), kst).isoformat(timespec="seconds")
                std = f"NAK 26 {a.kind} {a.quality} ({a.min_dpi}ppi 이상)" if a.kind else f"{a.min_dpi}dpi 이상(기관 지침)"
                wr.writerow([r[0], "SHA-256", h, mt, now, a.process, a.doc_purpose, a.doc_operator, a.doc_device, a.doc_software, r[10], std, r[-1],
                             "기록관리시스템에서 부여(이 도구는 하지 않음)"])
        print(f"전자화 이력 {len(rows)}건 → {a.history}  (시행령 제29조③)")
    bad = sum(1 for r in rows if r[-1] not in ("통과", "포맷 확인(변환 후보)"))
    fmt = sum(1 for r in rows if r[-1] == "포맷 확인(변환 후보)")
    print(f"{len(rows)}장 점검 · 재스캔 후보 {bad}장" + (f" · 포맷 확인 {fmt}장" if fmt else "") + f" → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
