#!/usr/bin/env python3
"""손상 전자기록물 1차 분류: 이관·인수한 전자파일이 열리지 않는 원인을 국가기록원 손상 코드로 나누고 복구 가능성을 붙인다.
근거: 국가기록원 「파일손상 전자기록물 복구 프로토타입 연구」(2018) 표 18 손상 코드별 설명·복구 가능 여부(O 원상 복구 가능 · △ 텍스트 추출 가능성 · X 불가),
      표 25 손상파일 구분 결과(샘플 4,724개), 「손상파일 검사·복구 프로토타입 적용방안 연구」(2019).
      여기서 보는 것은 파일 앞뒤 바이트와 구조만으로 판단할 수 있는 코드다:
        S001 원본 장애·삭제·바이러스 안내문만 남은 파일 · B002 HML(HWPML)인데 XML 형식 오류 · B003 10바이트 이하
        B007 HWP(OLE)인데 'HWP Document File' 없음 · B015 OLE 파일 뒤 잘림(섹터 크기와 안 맞음) · B020 HWP 3.0 헤더(30+128바이트)를 못 읽음
        E001 내용 전체가 Base64 · E002 XML 형태의 .hox(핸디소프트) · E003 '#Beginning_of_MetaData_Area_DestinyFS'로 시작
        E004 '=?euc-kr?B?'로 시작(Base64 메일 인코딩) · E005 문서 DRM(소프트캠프·마크애니·파수 헤더) · E006 HC20 · E007 FSPA
      연구 밖이지만 같은 방식으로 더 보는 것(코드 앞에 '추가'): PDF 끝 %%EOF 없음(뒤 잘림 의심), ZIP 계열(docx·xlsx·pptx·hwpx) 압축 구조 손상.

사용:  python damage_check.py 폴더/ -o 산출/손상파일_분류.csv
원칙:  읽기만 한다. 복구는 하지 않는다(국가기록원 복구 엔진·전문 도구로). 바이러스 안내문 파일은 백신 검사 결과를 함께 확인한다.
"""
import argparse
import base64
import csv
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

S001_TEXT = ["원문서을 확인할 수 없습니다", "원문서를 확인할 수 없습니다", "데이터가 손상 또는 손실되어", "전자문서시스템 원본 장애", "0x80004005",
             "바이러스 감염", "원본이 손상되어 존재하지 않는 문서", "첨부파일 오류", "삭제된 첨부파일", "오류 404"]
DRM = [(rb"SCD000[2-4]", "소프트캠프"), (rb"Document SAFER", "마크애니"), (rb"(?i)fasoo", "파수")]
CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIPS = {"docx", "xlsx", "pptx", "hwpx", "odt", "ods", "odp"}
RECOVER = {"S001": "X", "B002": "O", "B003": "X", "B007": "△", "B015": "O", "B020": "X", "E001": "O", "E002": "X", "E003": "O", "E004": "O",
           "E005": "X", "E006": "X", "E007": "O"}


def as_text(b):
    if b[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return b.decode("utf-16", "replace")
    for enc in ("utf-8", "cp949"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return ""


def classify(path):
    """[(코드, 설명)]"""
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    size = os.path.getsize(path)
    if size <= 10:
        return [("B003", f"파일 크기 {size}바이트 — 문서로 취급하기 어려움")]
    with open(path, "rb") as f:
        head = f.read(65536)
        f.seek(max(0, size - 2048))
        tail = f.read()
    out = []
    for rx, v in DRM:
        if re.search(rx, head[:4096]):
            return [("E005", f"문서 DRM({v}) — 복호화 협의 필요(transfer-schedule encrypted_scan.py)")]
    for sig, code, desc in ((b"#Beginning_of_MetaData_Area_DestinyFS", "E003", "DestinyFS 메타데이터 영역으로 시작 — 내부에 원파일이 남아 있음"),
                            (b"=?euc-kr?B?", "E004", "메일 Base64 인코딩(파일명+내용)으로 시작"), (b"HC20", "E006", "HC20으로 시작"), (b"FSPA", "E007", "FSPA로 시작")):
        if head.startswith(sig):
            return [(code, desc)]
    if size < 3 * 1024 * 1024 and not head.startswith((CFB, b"PK\x03\x04", b"%PDF", b"\x89PNG", b"\xff\xd8")):
        txt = as_text(open(path, "rb").read())
        hit = next((k for k in S001_TEXT if k in txt), None)
        if hit and len(txt) < 4000:
            return [("S001", f"안내문만 남은 파일('{hit}') — 원본 장애·삭제·감염")]
        s = re.sub(r"\s", "", txt)
        if len(s) >= 64 and re.fullmatch(r"[A-Za-z0-9+/]+=*", s) and len(s) % 4 == 0:
            try:
                base64.b64decode(s, validate=True)
                return [("E001", "내용 전체가 Base64 — 풀면 원파일일 수 있음")]
            except ValueError:
                pass
        if ext == "hox" and txt.lstrip().startswith("<"):
            return [("E002", "XML 형태의 .hox(핸디소프트) — 전용 변환기 필요")]
        if ext == "hml":
            try:
                ET.fromstring(open(path, "rb").read())
            except ET.ParseError as e:
                return [("B002", f"HML(HWPML)인데 XML 형식 오류({str(e)[:60]})")]
    if head.startswith(CFB):
        shift = int.from_bytes(head[30:32], "little") if len(head) >= 32 else 9
        sector = 1 << shift if shift in (9, 12) else 512
        if (size - 512) % sector:
            out.append(("B015", f"OLE 파일 크기가 섹터({sector}바이트) 단위가 아님 — 뒤가 잘린 형태 의심"))
        if ext == "hwp" and b"HWP Document File" not in open(path, "rb").read(min(size, 4 * 1024 * 1024)):
            out.append(("B007", "HWP(OLE)인데 'HWP Document File' 서명이 없음"))
    elif ext == "hwp" and head.startswith(b"HWP Document File V3"):
        if size < 30 + 128:
            out.append(("B020", "HWP 3.0 파일헤더(30바이트 인식정보+128바이트 문서정보)를 다 읽을 수 없음"))
    elif ext == "pdf" and head.startswith(b"%PDF") and b"%%EOF" not in tail:
        out.append(("추가-PDF", "파일 끝에 %%EOF가 없음 — 뒤 잘림 의심(복구 도구로 확인)"))
    elif ext in ZIPS:
        if not head.startswith(b"PK"):
            out.append(("추가-ZIP", f".{ext}인데 ZIP 서명이 없음"))
        else:
            try:
                with zipfile.ZipFile(path) as z:
                    bad = z.testzip()
                if bad:
                    out.append(("추가-ZIP", f"압축 항목 CRC 오류({bad})"))
            except zipfile.BadZipFile:
                out.append(("추가-ZIP", "ZIP 구조 손상(중앙 디렉터리 없음 — 뒤 잘림 의심)"))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--out", default="산출/손상파일_분류.csv")
    a = ap.parse_args()
    files = [a.folder] if os.path.isfile(a.folder) else [os.path.join(d, f) for d, _, fs in os.walk(a.folder) for f in sorted(fs)]
    rows, counts = [], {}
    for p in files:
        try:
            res = classify(p)
        except OSError as e:
            res = [("U001", f"읽기 오류 {e.__class__.__name__}")]
        for code, desc in res:
            counts[code] = counts.get(code, 0) + 1
            rows.append([os.path.relpath(p, a.folder) if os.path.isdir(a.folder) else os.path.basename(p), code, desc,
                         RECOVER.get(code, "확인")])
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["파일", "손상 코드", "설명", "복구 가능(O 원상·△ 텍스트·X 불가)"])
        w.writerows(rows)
    print(f"{len(files)}개 검사 · 손상 의심 {len({r[0] for r in rows})}개 " + " ".join(f"{k}:{v}" for k, v in sorted(counts.items())) + f" → {a.out}")
    return 1 if rows else 0


if __name__ == "__main__":
    sys.exit(main())
