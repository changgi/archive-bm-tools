#!/usr/bin/env python3
"""SIARD 2.1 이관 파일 구조 점검: 행정정보 데이터세트를 SIARD로 내보낸 파일이 NAK 35:2020(v1.0) 「행정정보 데이터세트 관리」
부속서 B(SIARD 2.1, eCH-0165 한글 번역)의 필수(M) 요건 중 파일·폴더 구조로 확인할 수 있는 것을 지키는지 본다.
근거: NAK 35:2020 부속서 B — G_4.1-1~5(ZIP·deflate/비압축·암호화 금지·.siard 확장자), P_4.2-1~6(content/·header/만,
      schema/table 폴더, header/siardversion/2.1/, header/metadata.xml·metadata.xsd, 이름 규칙), P_4.3-1·2(metadata.xml과 content 구조·열 수 일치)
      (국가기록원, 2026-11-20 내려받아 확인). SIARD는 NAK 37:2025 데이터베이스형 '수용가능포맷'이다(보존포맷은 CSV·JSON·XML).

사용:  python siard_check.py 데이터세트.siard -o 산출/SIARD점검.csv
원칙:  읽기만 한다. 테이블 값 자체(개인정보 포함 가능)는 열지 않고 구조·이름·열 수만 본다.
"""
import argparse
import csv
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]*(\.[A-Za-z0-9_]+)?$")


def local(tag):
    return tag.rsplit("}", 1)[-1]


def check(path):
    probs = []  # (요건, 대상, 내용, 구분)
    if not path.lower().endswith(".siard"):
        probs.append(("G_4.1-5", os.path.basename(path), "확장자가 .siard가 아님", "M"))
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return probs + [("G_4.1-1", os.path.basename(path), "ZIP으로 열리지 않음", "M")], {}
    infos = z.infolist()
    names = [i.filename for i in infos]
    for i in infos:
        if i.flag_bits & 0x1:
            probs.append(("G_4.1-3", i.filename, "암호로 보호·암호화된 항목", "M"))
        if i.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            probs.append(("G_4.1-2", i.filename, f"압축 방식 {i.compress_type} — 비압축 또는 deflate만", "M"))
    tops = {n.split("/", 1)[0] for n in names if n}
    for t in sorted(tops - {"content", "header"}):
        probs.append(("P_4.2-1", t, "최상위에 content/·header/ 밖의 폴더·파일", "M"))
    if not any(n.startswith("header/siardversion/2.1/") or n == "header/siardversion/2.1" for n in names):
        probs.append(("P_4.2-4", "header/siardversion/2.1/", "버전 표시 빈 폴더가 없음", "M"))
    for f in ("header/metadata.xml", "header/metadata.xsd"):
        if f not in names:
            probs.append(("P_4.2-5", f, "없음", "M"))
    for n in names:
        if not n.startswith("content/"):
            continue
        for part in [p for p in n.split("/") if p]:
            if not NAME.match(part):
                probs.append(("P_4.2-6", n, f"이름 규칙 위반('{part}': 영문자로 시작, 영문·숫자·_만)", "M"))
                break
            if len(part) > 20:
                probs.append(("P_4.2-6 권고", n, f"이름이 20자 초과('{part}')", "권고"))
                break
    # content 구조: content/<schema>/<table>/<table>.xml(.xsd)
    content = {}
    for n in names:
        parts = [p for p in n.split("/") if p]
        if len(parts) >= 3 and parts[0] == "content":
            content.setdefault(parts[1], set()).add(parts[2])
    for sch, tabs in content.items():
        for tb in tabs:
            for ext in (".xml", ".xsd"):
                f = f"content/{sch}/{tb}/{tb}{ext}"
                if f not in names:
                    probs.append(("P_4.2-3", f, "테이블 폴더에 같은 이름의 데이터 파일이 없음", "M"))
    info = {"스키마": len(content), "테이블": sum(len(v) for v in content.values())}
    if "header/metadata.xml" in names:
        try:
            root = ET.fromstring(z.read("header/metadata.xml"))
        except ET.ParseError as e:
            return probs + [("P_4.3-1", "header/metadata.xml", f"XML 파싱 실패: {e}", "M")], info
        declared = {}
        for s in (e for e in root.iter() if local(e.tag) == "schema"):
            sf = next((c.text or "" for c in s if local(c.tag) == "folder"), "").strip()
            for t in (e for e in s.iter() if local(e.tag) == "table"):
                tf = next((c.text or "" for c in t if local(c.tag) == "folder"), "").strip()
                cols = sum(1 for e in t.iter() if local(e.tag) == "column")
                declared[(sf, tf)] = cols
        found = {(s, t) for s, ts in content.items() for t in ts}
        for k in sorted(set(declared) - found):
            probs.append(("P_4.3-1", f"content/{k[0]}/{k[1]}/", "metadata.xml에는 있으나 content에 없음", "M"))
        for k in sorted(found - set(declared)):
            probs.append(("P_4.3-1", f"content/{k[0]}/{k[1]}/", "content에는 있으나 metadata.xml에 없음", "M"))
        for (sf, tf), cols in declared.items():
            xsd = f"content/{sf}/{tf}/{tf}.xsd"
            if xsd in names:
                try:
                    x = ET.fromstring(z.read(xsd))
                    row = next((e for e in x.iter() if local(e.tag) == "element" and e.get("name") == "row"), None)
                    n = sum(1 for e in row.iter() if local(e.tag) == "element" and e is not row) if row is not None else None
                    if n is not None and n != cols:
                        probs.append(("P_4.3-2", xsd, f"열 수 불일치(metadata.xml {cols} · xsd {n})", "M"))
                except ET.ParseError:
                    probs.append(("P_4.3-2", xsd, "xsd 파싱 실패", "M"))
    return probs, info


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("siard")
    ap.add_argument("-o", "--out", default="산출/SIARD점검.csv")
    a = ap.parse_args()
    probs, info = check(a.siard)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["요건", "대상", "내용", "구분"])
        w.writerows(probs)
        if not probs:
            w.writerow(["-", os.path.basename(a.siard), f"구조 요건 통과(스키마 {info.get('스키마', 0)} · 테이블 {info.get('테이블', 0)})", "통과"])
    m = sum(1 for p in probs if p[3] == "M")
    print(f"SIARD 스키마 {info.get('스키마', 0)} · 테이블 {info.get('테이블', 0)} · 필수 위반 {m} · 권고 {len(probs) - m} → {a.out}")
    return 1 if m else 0


if __name__ == "__main__":
    sys.exit(main())
