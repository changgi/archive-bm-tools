#!/usr/bin/env python3
"""스캔 품질 점검: 스캔 이미지의 해상도(DPI)·기울기·빈 면·밝기·크기를 점검해 재스캔 후보를 표시한다.

사용:  python scan_qc.py 스캔폴더/ [--min-dpi 300] [--max-skew 1.0] -o 산출/스캔품질.csv
대상:  .tif .tiff .jpg .jpeg .png .bmp (PDF는 이미지로 뽑은 뒤 점검)
기준값은 기관의 스캔 지침에 맞게 바꾼다(기본 300dpi·기울기 1도).
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


def check(path, min_dpi, max_skew):
    im = Image.open(path)
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
        flags.append(f"DPI {dpi} < {min_dpi}")
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
    return [w, h, dpi or "", im.mode, round(mean, 1), round(skew, 2), " · ".join(flags), master, "재스캔 후보" if flags else "통과"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--min-dpi", type=int, default=300)
    ap.add_argument("--max-skew", type=float, default=1.0)
    ap.add_argument("-o", "--out", default="산출/스캔품질.csv")
    ap.add_argument("--history", help="전자화 이력 CSV도 쓴다(시행령 제29조③: 해시값·변환 시점·변환 과정)")
    ap.add_argument("--process", default="", help="변환 과정 설명(장비·소프트웨어·해상도·담당 등) — 이력에 그대로 적는다")
    a = ap.parse_args()
    rows = []
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            if os.path.splitext(fn)[1].lower() not in EXTS:
                continue
            p = os.path.join(root, fn)
            try:
                rows.append([os.path.relpath(p, a.folder), *check(p, a.min_dpi, a.max_skew)])
            except Exception as e:
                rows.append([os.path.relpath(p, a.folder), "", "", "", "", "", "", f"읽기 실패: {str(e)[:80]}", "", "확인 필요"])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        wr = csv.writer(f)
        wr.writerow(["파일", "가로(px)", "세로(px)", "DPI", "색 방식", "평균 밝기", "기울기(°)", "표시", "보존 포맷", "판정"])
        wr.writerows(rows)
    if a.history:
        import datetime as _dt
        import hashlib
        kst = _dt.timezone(_dt.timedelta(hours=9))
        with open(a.history, "w", encoding="utf-8-sig", newline="") as f:
            wr = csv.writer(f)
            wr.writerow(["파일", "SHA-256", "파일 수정 시각(KST)", "이력 작성 시각(KST)", "변환 과정", "품질 판정", "행정전자서명"])
            now = _dt.datetime.now(kst).isoformat(timespec="seconds")
            for r in rows:
                fp = os.path.join(a.folder, r[0])
                h = hashlib.sha256(open(fp, "rb").read()).hexdigest()
                mt = _dt.datetime.fromtimestamp(os.path.getmtime(fp), kst).isoformat(timespec="seconds")
                wr.writerow([r[0], h, mt, now, a.process, r[-1], "기록관리시스템에서 부여(이 도구는 하지 않음)"])
        print(f"전자화 이력 {len(rows)}건 → {a.history}  (시행령 제29조③)")
    bad = sum(1 for r in rows if r[-1] != "통과")
    print(f"{len(rows)}장 점검 · 재스캔 후보 {bad}장 → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
