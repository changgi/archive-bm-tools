#!/usr/bin/env python3
"""암호화 기록물 찾기: 이관할 폴더에서 암호가 걸린 파일을 찾아 '암호화된 기록물 목록'을 만든다.
시행령 제31조의2 [시행 2027. 1. 1.] — 생산·접수 과정에서 암호화해 관리하던 기록물을 이관할 때는
암호화된 기록물의 목록과 그 해제에 필요한 정보를 첨부한다(기록관→영구기록물관리기관도 같음).

사용:  python encrypted_scan.py 이관폴더 [--list 이관목록.xlsx] -o 산출/암호화목록.xlsx
찾는 것(파일 내용으로 판단, 비밀번호를 풀지 않음):
  PDF       /Encrypt 사전
  OOXML     암호화된 docx·xlsx·pptx는 OLE(CFB) 파일 안에 EncryptionInfo·EncryptedPackage 스트림
  HWP 5     FileHeader 속성 비트 1(암호 설정)
  HWPX·ODF  META-INF/manifest.xml 의 encryption-data
  ZIP       항목의 암호화 플래그(general purpose bit 0)
  옛 DOC·XLS  확실히 판단하지 못하면 '확인 필요'
원칙:  목록만 만든다. 비밀번호·해제 정보는 이 파일에 적지 말고 기관의 보안 절차로 따로 전달한다.
       파일을 열거나 고치지 않는다(읽기만).
"""
import argparse
import os
import sys
import zipfile

CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
U16 = lambda s: s.encode("utf-16-le")


def read_head_tail(path, n=2_000_000):
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        if size <= 2 * n:
            return f.read()
        head = f.read(n)
        f.seek(size - n)
        return head + f.read()


def check(path):
    """(판정, 근거) — 판정: '암호화' / '아님' / '확인 필요'"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if os.path.getsize(path) == 0:
            return "확인 필요", "0바이트"
        with open(path, "rb") as f:
            sig = f.read(8)
        if sig.startswith(b"%PDF"):
            data = read_head_tail(path)
            return ("암호화", "PDF /Encrypt") if b"/Encrypt" in data else ("아님", "PDF")
        if sig == CFB:
            data = read_head_tail(path, 8_000_000)
            if U16("EncryptionInfo") in data or U16("EncryptedPackage") in data:
                return "암호화", "OOXML 암호(EncryptionInfo)"
            i = data.find(b"HWP Document File")
            if i >= 0:
                prop = int.from_bytes(data[i + 36:i + 40], "little") if len(data) >= i + 40 else 0
                return ("암호화", "HWP 암호 설정") if prop & 0x2 else ("아님", "HWP")
            if ext in (".doc", ".xls", ".ppt"):
                if U16("EncryptedSummary") in data or b"Microsoft Enhanced Cryptographic" in data or b"Microsoft Base Cryptographic" in data:
                    return "암호화", "옛 Office RC4 암호"
                return "확인 필요", "옛 Office 형식 — 열어 보아 확인"
            return "아님", "OLE"
        if sig.startswith(b"PK"):
            with zipfile.ZipFile(path) as z:
                infos = z.infolist()
                if any(i.flag_bits & 0x1 for i in infos):
                    return "암호화", "ZIP 항목 암호"
                for name in ("META-INF/manifest.xml",):
                    if name in z.namelist():
                        if b"encryption-data" in z.read(name):
                            return "암호화", "manifest encryption-data(HWPX·ODF)"
            return "아님", "ZIP 기반"
        return "아님", "그 밖의 형식"
    except zipfile.BadZipFile:
        return "확인 필요", "ZIP 손상"
    except OSError as e:
        return "확인 필요", f"읽기 오류 {e.__class__.__name__}"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--list", help="이관목록(파일명 또는 경로 열이 있으면 철분류번호·제목을 붙임)")
    ap.add_argument("-o", "--out", default="산출/암호화목록.xlsx")
    a = ap.parse_args()
    skip = ("비밀", "secret", "classified")
    found = []
    total = 0
    for root, dirs, files in os.walk(a.folder):
        dirs[:] = [d for d in dirs if not any(k in d.lower() for k in skip)]
        for fn in sorted(files):
            if fn.startswith(("~$", ".")):
                continue
            total += 1
            p = os.path.join(root, fn)
            verdict, why = check(p)
            if verdict != "아님":
                found.append([os.path.relpath(p, a.folder), verdict, why, os.path.getsize(p)])
    meta = {}
    if a.list:
        import openpyxl
        ws = openpyxl.load_workbook(a.list, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
        head = rows[0] if rows else []
        fi = next((head.index(c) for c in ("파일명", "파일경로", "경로", "파일") if c in head), None)
        ci = next((head.index(c) for c in ("철분류번호", "등록번호", "관리번호") if c in head), None)
        ti = next((head.index(c) for c in ("철제목", "제목") if c in head), None)
        if fi is not None:
            for r in rows[1:]:
                if fi < len(r) and r[fi]:
                    meta[os.path.basename(r[fi])] = (r[ci] if ci is not None and ci < len(r) else "", r[ti] if ti is not None and ti < len(r) else "")
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "암호화된 기록물 목록"
    ws.append(["연번", "파일(이관 폴더 기준)", "철분류번호·등록번호", "제목", "판정", "근거", "크기(바이트)", "해제 정보 전달 방법(비밀번호는 적지 않음)"])
    for i, (rel, v, why, size) in enumerate(found, 1):
        no, title = meta.get(os.path.basename(rel), ("", ""))
        ws.append([i, rel, no, title, v, why, size, ""])
    if not found:
        ws.append(["", "해당 없음"])
    g = wb.create_sheet("안내")
    for line in [["근거", "공공기록물 관리에 관한 법률 시행령 제31조의2 [시행 2027. 1. 1.]"],
                 ["할 일", "암호화된 기록물의 목록과 그 해제에 필요한 정보를 첨부해 이관(기록관→영구기록물관리기관도 같음)"],
                 ["주의", "비밀번호·키를 이 파일이나 메일 본문에 적지 말고 기관 보안 절차(별도 매체·대면 등)로 전달"],
                 ["주의", "'확인 필요'는 형식상 판단하지 못한 것 — 파일을 열어 보아 확인"],
                 ["제외", "경로에 비밀·secret·classified가 든 폴더는 검사하지 않음(비밀기록물은 담당자가 따로)"]]:
        g.append(line)
    for i, w in enumerate((6, 50, 22, 36, 10, 30, 12, 40)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    enc = sum(1 for f in found if f[1] == "암호화")
    print(f"파일 {total}개 · 암호화 {enc} · 확인 필요 {len(found) - enc} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
