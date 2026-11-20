#!/usr/bin/env python3
"""시청각 디지털화 결과 점검: 영상·음성 파일을 ffprobe로 읽어 NAK 26:2023(v2.1) 「기록물 디지털화 기준」 표 9~11의
해상도·코덱·파일포맷·샘플링·전체 비트레이트에 맞는지 본다.

사용:  python av_qc.py 디지털화폴더/ --kind 비디오HD -o 산출/시청각디지털화.csv
종류(--kind):
  영화35   35㎜ 영화필름: 4K(4096×2160), 무손실 MKV/FFV1/FLAC 또는 손실 MP4/H.264 20Mbps 이상(DPX 프레임 스캔도 가능)
  영화16   16㎜·8㎜ 영화필름: 2K(2048×1080), 손실이면 10Mbps 이상
  비디오SD 비디오테이프·DVD SD: 720×480, 손실이면 2.5Mbps 이상
  비디오HD 비디오테이프 HD: 1920×1080, 손실(H.264·H.265)이면 10Mbps 이상
  음성     오디오테이프·오디오CD: 24bit·48kHz, FLAC(무손실) 또는 AAC(손실)
공통:  음성 24bit·48kHz. 영상 서브샘플링 4:2:2(yuv422) 권장 — 4:2:0이면 '확인'으로 표시.
원칙:  ffprobe 값만 본다. 재생 품질(잡음·끊김)은 사람이 보고 듣는다. ffprobe가 없으면 설치를 안내하고 멈춘다.
"""
import argparse
import csv
import json
import os
import shutil
import subprocess
import sys

KIND = {"영화35": ((4096, 2160), 20.0), "영화16": ((2048, 1080), 10.0), "비디오SD": ((720, 480), 2.5), "비디오HD": ((1920, 1080), 10.0), "음성": (None, None)}
EXTS = {".mkv", ".mp4", ".mov", ".avi", ".mxf", ".mpg", ".mpeg", ".ts", ".flac", ".aac", ".m4a", ".wav", ".mp3", ".dpx"}


def probe(path):
    r = subprocess.run(["ffprobe", "-v", "error", "-print_format", "json", "-show_format", "-show_streams", path], capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError(r.stderr.strip()[:120] or "ffprobe 실패")
    return json.loads(r.stdout)


def check(path, kind):
    info = probe(path)
    fmt = info.get("format", {})
    container = fmt.get("format_name", "")
    br = float(fmt.get("bit_rate") or 0) / 1e6
    vs = next((s for s in info.get("streams", []) if s.get("codec_type") == "video" and not s.get("disposition", {}).get("attached_pic")), None)
    au = next((s for s in info.get("streams", []) if s.get("codec_type") == "audio"), None)
    need, min_br = KIND[kind]
    hard, soft = [], []
    vcodec = vs.get("codec_name", "") if vs else ""
    acodec = au.get("codec_name", "") if au else ""
    ext = os.path.splitext(path)[1].lower()
    lossless = vcodec == "ffv1" or (kind == "음성" and acodec == "flac")
    if kind != "음성":
        if not vs:
            hard.append("영상 스트림 없음")
        else:
            w, h = int(vs.get("width", 0)), int(vs.get("height", 0))
            if need and (w < need[0] or h < need[1]):
                hard.append(f"해상도 {w}×{h} < {need[0]}×{need[1]}")
            if vcodec == "ffv1":
                if "matroska" not in container:
                    soft.append("FFV1은 MKV(Matroska)에 담음")
            elif vcodec in ("h264", "hevc"):
                if "mp4" not in container:
                    soft.append(f"{vcodec.upper()}는 MP4에 담음")
                if min_br and br and br < min_br:
                    hard.append(f"전체 비트레이트 {br:.1f}Mbps < {min_br:g}Mbps")
                if vcodec == "hevc" and "10" not in (vs.get("pix_fmt", "") + str(vs.get("bits_per_raw_sample", ""))):
                    soft.append("H.265는 10bit 권장")
            elif ext != ".dpx":
                hard.append(f"영상 코덱 {vcodec or '알 수 없음'} — 목록 밖(FFV1·H.264·H.265, 필름은 DPX)")
            pf = vs.get("pix_fmt", "")
            if pf and "422" not in pf and "444" not in pf and not pf.startswith(("rgb", "gbr", "bgr")):
                soft.append(f"서브샘플링 {pf} — 4:2:2 권장")
    if au:
        sr = int(au.get("sample_rate") or 0)
        bits = int(au.get("bits_per_raw_sample") or au.get("bits_per_sample") or 0)
        if sr and sr < 48000:
            hard.append(f"샘플링 {sr / 1000:g}kHz < 48kHz")
        if acodec in ("flac", "pcm_s24le", "pcm_s24be", "pcm_s32le") and bits and bits < 24:
            hard.append(f"음성 {bits}bit < 24bit")
        if acodec not in ("flac", "aac") and not acodec.startswith("pcm"):
            soft.append(f"음성 코덱 {acodec} — 목록 밖(FLAC·AAC)")
        if acodec.startswith("pcm"):
            soft.append("무압축 PCM(WAV) — 목록은 FLAC·AAC, 변환 후보")
        if kind != "음성" and vcodec == "ffv1" and acodec != "flac":
            soft.append("무손실(FFV1)이면 음성은 FLAC")
    elif kind == "음성":
        hard.append("음성 스트림 없음")
    if kind == "음성" and vs and not vs.get("disposition", {}).get("attached_pic"):
        soft.append("음성 기록인데 영상 스트림 있음")
    res = f"{vs.get('width')}×{vs.get('height')}" if vs else ""
    verdict = "재작업 후보" if hard else "확인" if soft else "통과"
    return [container, vcodec, res, (vs or {}).get("pix_fmt", ""), acodec, (au or {}).get("sample_rate", ""),
            (au or {}).get("bits_per_raw_sample", "") if au else "", f"{br:.1f}" if br else "", "무손실" if lossless else "손실" if vcodec or acodec else "",
            " · ".join(hard + soft), verdict]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--kind", required=True, choices=list(KIND))
    ap.add_argument("-o", "--out", default="산출/시청각디지털화.csv")
    a = ap.parse_args()
    if not shutil.which("ffprobe"):
        print("ffprobe가 없습니다. FFmpeg를 설치한 뒤 다시 실행하세요(담당자 허락 후).", file=sys.stderr)
        return 2
    rows = []
    for root, _, fs in os.walk(a.folder):
        for fn in sorted(fs):
            if os.path.splitext(fn)[1].lower() not in EXTS:
                continue
            p = os.path.join(root, fn)
            try:
                rows.append([os.path.relpath(p, a.folder), *check(p, a.kind)])
            except Exception as e:
                rows.append([os.path.relpath(p, a.folder)] + [""] * 9 + [f"읽기 실패: {str(e)[:80]}", "확인 필요"])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        wr = csv.writer(f)
        wr.writerow(["파일", "컨테이너", "영상 코덱", "해상도", "화소 형식", "음성 코덱", "샘플링(Hz)", "음성 비트", "전체 비트레이트(Mbps)", "압축",
                     "표시", "판정"])
        wr.writerows(rows)
    bad = sum(r[-1] == "재작업 후보" for r in rows)
    print(f"{len(rows)}개 점검({a.kind}, NAK 26) · 재작업 후보 {bad} · 확인 {sum(r[-1] == '확인' for r in rows)} → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
