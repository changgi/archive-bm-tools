#!/usr/bin/env python3
"""사진 등록대상 선별 보조: 행사·현장 사진 폴더에서 연속촬영(연사) 묶음과 화질 문제를 찾아 등록 후보를 추린다.

사용:  python photo_select.py 사진폴더/ [--burst-sec 3] [--min-mp 2] -o 산출/사진_선별.xlsx
대상:  .jpg .jpeg .png .tif .tiff (하위 폴더 포함)
근거:  국가기록원 「시청각기록물 관리지침」(2024. 12., 표준·지침·매뉴얼 > 지침) 3.2 등록대상 선정 기준
         - 사진·필름류는 촬영물 중 보존대상으로 적합한 작품을 선정한 뒤 생산등록번호를 부여(보존대상 선별 직후 등록)
         - 원본 우선, 전자파일은 고해상도·대용량 작품
         - 색감: 노출 과다·부족으로 실물보다 희거나 검게 표현된 작품, 의도적으로 흑백·세피아로 바꾼 작품 제외
         - 연속촬영물(연사)은 장면별 대표 사진 선정
         - 불특정 일반인의 얼굴 정면 컷은 지양(사생활·초상권) — 이 스크립트는 얼굴을 판별하지 않으니 사람이 본다
원칙:  화소·밝기·채도·촬영시각·작은 축소 이미지 비교만 한다. 구도·인물·장소 식별성은 사람이 판단한다. 원본 파일은 바꾸지 않는다.
"""
import argparse
import datetime as dt
import os
import sys

EXT = (".jpg", ".jpeg", ".png", ".tif", ".tiff")


def exif_time(im):
    try:
        ex = im.getexif()
        v = ex.get(36867) or ex.get_ifd(0x8769).get(36867) or ex.get(306)
        return dt.datetime.strptime(str(v)[:19], "%Y:%m:%d %H:%M:%S") if v else None
    except Exception:
        return None


def dhash(im, size=8):
    import numpy as np
    g = np.asarray(im.convert("L").resize((size + 1, size)), dtype=np.int16)
    bits = (g[:, :-1] > g[:, 1:]).flatten()
    return sum(1 << i for i, b in enumerate(bits) if b)


def main() -> int:
    ap = argparse.ArgumentParser(description="사진 등록대상 선별 보조(연사 묶음·노출·흑백 변환·저해상도)")
    ap.add_argument("folder")
    ap.add_argument("--burst-sec", type=float, default=3.0, help="이 초 안에 찍힌 비슷한 사진을 연사로 묶음")
    ap.add_argument("--min-mp", type=float, default=2.0, help="이보다 작은 화소(백만)는 저해상도로 표시")
    ap.add_argument("-o", "--out", default="산출/사진_선별.xlsx")
    a = ap.parse_args()
    try:
        import numpy as np
        from PIL import Image
    except ModuleNotFoundError:
        print("Pillow·numpy 꾸러미가 필요합니다. 담당자 허락 후 `python -m pip install pillow numpy`", file=sys.stderr)
        return 2
    rows = []
    for root, _, files in os.walk(a.folder):
        for fn in sorted(files):
            if not fn.lower().endswith(EXT):
                continue
            p = os.path.join(root, fn)
            try:
                with Image.open(p) as im:
                    w, h = im.size
                    t = exif_time(im)
                    sm = im.convert("RGB")
                    sm.thumbnail((256, 256))
                    arr = np.asarray(sm, dtype=np.float32)
                    hsv = np.asarray(sm.convert("HSV"), dtype=np.float32)
                    hsh = dhash(sm)
            except Exception as e:
                rows.append({"파일": os.path.relpath(p, a.folder), "오류": f"열 수 없음({type(e).__name__})"})
                continue
            lum = arr.mean(axis=2)
            mean, dark, bright = float(lum.mean()), float((lum < 16).mean()), float((lum > 240).mean())
            sat = float(hsv[..., 1].mean())
            r, g, b = (float(arr[..., k].mean()) for k in range(3))
            notes = []
            if mean > 215 or bright > 0.45:
                notes.append("노출 과다(희게 날아감)")
            if mean < 40 or dark > 0.55:
                notes.append("노출 부족(검게 뭉개짐)")
            if sat < 8:
                notes.append("흑백 변환 의심")
            elif sat < 40 and r > g > b and (r - b) > 18:
                notes.append("세피아 변환 의심")
            mp = w * h / 1e6
            if mp < a.min_mp:
                notes.append(f"저해상도({mp:.1f}MP)")
            rows.append({"파일": os.path.relpath(p, a.folder), "촬영일시": t, "가로": w, "세로": h, "화소(MP)": round(mp, 1), "평균밝기": round(mean),
                         "채도": round(sat), "바이트": os.path.getsize(p), "hash": hsh, "문제": notes, "오류": ""})
    ok = [x for x in rows if not x.get("오류")]
    ok.sort(key=lambda x: (x["촬영일시"] or dt.datetime.max, x["파일"]))
    gid, groups = 0, {}
    prev = None
    for x in ok:
        same = False
        if prev is not None:
            close = x["촬영일시"] and prev["촬영일시"] and (x["촬영일시"] - prev["촬영일시"]).total_seconds() <= a.burst_sec
            similar = bin(x["hash"] ^ prev["hash"]).count("1") <= 12
            same = bool(close and similar) or (not x["촬영일시"] and bin(x["hash"] ^ prev["hash"]).count("1") <= 6)
        if not same:
            gid += 1
        x["묶음"] = gid
        groups.setdefault(gid, []).append(x)
        prev = x
    for g, xs in groups.items():
        good = [x for x in xs if not x["문제"]] or xs
        best = max(good, key=lambda x: (x["화소(MP)"], -abs(x["평균밝기"] - 128), x["바이트"]))
        for x in xs:
            if len(xs) == 1:
                x["판정"] = "등록 후보" if not x["문제"] else "제외 검토 — " + "·".join(x["문제"])
            elif x is best:
                x["판정"] = f"연사 {len(xs)}장 중 대표 후보" + (" (문제 있음: " + "·".join(x["문제"]) + ")" if x["문제"] else "")
            else:
                x["판정"] = "연사 — 대표만 등록 검토" + (" · " + "·".join(x["문제"]) if x["문제"] else "")
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "선별"
    ws.append(["파일", "촬영일시", "가로", "세로", "화소(MP)", "평균밝기(0~255)", "채도", "연사 묶음", "판정", "사람이 볼 것"])
    for x in ok:
        ws.append([x["파일"], x["촬영일시"].strftime("%Y-%m-%d %H:%M:%S") if x["촬영일시"] else "", x["가로"], x["세로"], x["화소(MP)"], x["평균밝기"],
                   x["채도"], x["묶음"], x["판정"], "구도·비율, 주요 인물 식별성(눈 감음 등), 장소·행사 식별성, 일반인 얼굴 정면 컷"])
    for x in rows:
        if x.get("오류"):
            ws.append([x["파일"], "", "", "", "", "", "", "", x["오류"], ""])
    s = wb.create_sheet("요약·근거")
    n_burst = sum(1 for xs in groups.values() if len(xs) > 1)
    cand = sum(1 for x in ok if x["판정"] == "등록 후보" or "대표 후보" in x["판정"])
    for row in (["사진", len(rows)], ["연사 묶음", n_burst], ["등록 후보(단독·연사 대표)", cand],
                ["문제 있는 사진", sum(1 for x in ok if x["문제"])],
                ["근거", "국가기록원 「시청각기록물 관리지침」(2024) 3.2 등록대상 선정 기준 — 원본 우선, 고해상도, 노출 과다·부족·흑백/세피아 변환 제외, 연사는 장면별 대표"],
                ["등록", "선별한 사진은 바로 전자기록생산시스템에 사진·필름류로 등록하고, 내용요약에 촬영일자·장소·주요 인물 성명·직위를 넣는다(av_check.py로 점검)"],
                ["주의", "밝기·채도 기준은 자동 추정이다. 야간·실내·흑백 원본 필름 스캔본은 정상일 수 있으니 사람이 확인한다."]):
        s.append(row)
    for col, wdt in zip("ABCDEFGHIJ", (40, 20, 7, 7, 9, 12, 7, 9, 40, 60)):
        ws.column_dimensions[col].width = wdt
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"사진 {len(rows)} · 연사 묶음 {n_burst} · 등록 후보 {cand} · 문제 {sum(1 for x in ok if x['문제'])} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
