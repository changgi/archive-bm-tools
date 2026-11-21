#!/usr/bin/env python3
"""생산 시점 기록화(Records by design): 업무 파일 하나를 '맥락·증거성을 갖춘 기록 패키지'로 만들고,
그 패키지의 메타데이터 하나에서 여러 매개(등록대장 행·공개 데이터·한 장 요약 HTML·카드 문구·연결 데이터)를 다시 만들고,
나중에 고정성(해시)을 다시 확인해 이력에 쌓는다.

사용:
  python capture.py capture 원본파일 --task "단위과제명" --agency 기관명 --dept 부서 --person 담당자 --role 기안자
                    --period 10년 --access 공개|부분공개|비공개 [--reason "6호 …"] [--activity "업무 행위"] [--purpose "왜 만들었나"]
                    [--relation "참조:선행기록식별자"] [--id 기록식별자] [--title 제목] -o 산출/기록패키지
  python capture.py verify 산출/기록패키지/<식별자>          # 해시 다시 확인, 이력에 '고정성 점검' 추가
  python capture.py derive 산출/기록패키지/<식별자>          # 메타데이터에서 매개 파일 다시 만들기(원본은 건드리지 않음)
  python capture.py batch 목록.xlsx [--task … --agency … --dept … --period …] -o 산출/기록패키지   # 목록 한 행 = 파일 하나
  python capture.py catalog 산출/기록패키지 [--verify] -o 산출/기록목록   # 목록.csv·공개데이터.jsonl·연결그래프.jsonld·목록.html

패키지 구조(BagIt 1.0, RFC 8493 꼴 — 다른 기관·시스템으로 옮겨도 스스로 설명되는 상자):
  bagit.txt · bag-info.txt · manifest-sha256.txt · tagmanifest-sha256.txt
  data/원본/<파일>                 원본 그대로(바이트 같음)
  metadata.json                   NAK 8:2022 요소 번호·이름으로 쓴 메타데이터(필수 상위 요소 + 23 무결성 + 21 이력 + 22 관계)
  context.jsonld                  RiC-O 1.0 어휘로 쓴 맥락(기록-생산자-업무활동-단위과제-관계) — 연결 데이터
  events.jsonl                    PREMIS 사건 꼴 이력(생산·등록·패키지화·고정성 점검·파생) — 한 줄에 한 사건, 덧붙이기만
  매개/등록대장.csv · 매개/공개데이터.json · 매개/요약.html · 매개/카드.txt
근거:  공공기록물법 시행령 제4조(진본성·무결성·신뢰성·이용가능성), 제20조(전자기록물 생산 시 등록), NAK 8:2022(v2.3) 기록관리 메타데이터 표준
       (필수 상위 12요소, 23 무결성, 21 이력, 22 관계), ICA Records in Contexts RiC-O 1.0, PREMIS 3.0 사건, BagIt RFC 8493, ISO 15489-1:2016(맥락·증거)
원칙:  원본은 복사만 하고 바꾸지 않는다. 비밀 표시가 보이는 파일은 패키지로 만들지 않는다. 비공개·부분공개 기록은 공개 매개에 본문을 넣지 않고,
       개인정보 후보가 있으면 공개데이터에 넣지 않고 알린다. 메타데이터 값 가운데 자동 추출한 것은 '후보'로 표시한다(담당자 확인 전).
"""
import argparse
import csv
import datetime as dt
import hashlib
import html
import json
import os
import re
import shutil
import sys
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "pii-redact", "scripts"))
sys.path.insert(0, os.path.join(HERE, "..", "..", "metadata-autofill", "scripts"))
KST = dt.timezone(dt.timedelta(hours=9))
SECRET = re.compile(r"(Ⅰ|Ⅱ|Ⅲ|I|II|III|1|2|3)\s*급\s*비밀|대외비|비밀\s*등급|SECRET|CONFIDENTIAL", re.I)
PERIODS = ("1년", "3년", "5년", "10년", "30년", "준영구", "영구")
ROLES = ("기안자", "검토자", "협조자", "결재자", "접수자", "생산자", "작성자", "보고자")
FORMATS = {".pdf": "PDF", ".hwpx": "HWPX", ".hwp": "HWP", ".docx": "DOCX", ".xlsx": "XLSX", ".csv": "CSV", ".txt": "TXT",
           ".md": "Markdown", ".jpg": "JPEG", ".jpeg": "JPEG", ".png": "PNG", ".tif": "TIFF", ".tiff": "TIFF", ".mp4": "MP4", ".wav": "WAV"}
OPEN_FORMATS = {"PDF", "HWPX", "CSV", "TXT", "Markdown", "PNG", "TIFF", "WAV", "XLSX", "DOCX"}


def now():
    return dt.datetime.now(KST).replace(microsecond=0).isoformat()


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def read_text(path):
    try:
        from pii_scan import read_units
        return "\n".join(t for _, t in read_units(path))
    except Exception:
        return ""


def event(pkg, kind, detail, outcome="성공", agent="records-skills/record-by-design"):
    e = {"eventIdentifier": str(uuid.uuid4()), "eventType": kind, "eventDateTime": now(), "eventDetail": detail,
         "eventOutcome": outcome, "linkingAgent": agent}
    with open(os.path.join(pkg, "events.jsonl"), "a", encoding="utf-8") as f:
        f.write(json.dumps(e, ensure_ascii=False) + "\n")
    return e


def write_bag(pkg, meta):
    with open(os.path.join(pkg, "bagit.txt"), "w", encoding="utf-8") as f:
        f.write("BagIt-Version: 1.0\nTag-File-Character-Encoding: UTF-8\n")
    payload = []
    for root, _, files in os.walk(os.path.join(pkg, "data")):
        for n in sorted(files):
            p = os.path.join(root, n)
            payload.append((sha256(p), os.path.relpath(p, pkg).replace(os.sep, "/"), os.path.getsize(p)))
    with open(os.path.join(pkg, "manifest-sha256.txt"), "w", encoding="utf-8") as f:
        for h, rel, _ in payload:
            f.write(f"{h}  {rel}\n")
    info = {"Source-Organization": meta["1.2 기관명"], "Organization-Address": meta.get("1.3 부서명", ""),
            "External-Identifier": meta["3.1 기본식별자"], "External-Description": meta["4.1 제목"],
            "Bagging-Date": dt.datetime.now(KST).date().isoformat(), "Payload-Oxum": f"{sum(s for *_, s in payload)}.{len(payload)}",
            "Bag-Software-Agent": "records-skills record-by-design"}
    with open(os.path.join(pkg, "bag-info.txt"), "w", encoding="utf-8") as f:
        for k, v in info.items():
            f.write(f"{k}: {str(v).replace(chr(10), ' ')}\n")
    write_tags(pkg)


def write_tags(pkg):
    tags = [n for n in ("bagit.txt", "bag-info.txt", "manifest-sha256.txt", "metadata.json", "context.jsonld", "events.jsonl") if os.path.exists(os.path.join(pkg, n))]
    for root, _, files in os.walk(os.path.join(pkg, "매개")):
        tags += [os.path.relpath(os.path.join(root, n), pkg).replace(os.sep, "/") for n in sorted(files)]
    with open(os.path.join(pkg, "tagmanifest-sha256.txt"), "w", encoding="utf-8") as f:
        for rel in tags:
            f.write(f"{sha256(os.path.join(pkg, rel))}  {rel}\n")


def context_ld(meta):
    rid = f"urn:records:{meta['3.1 기본식별자']}"
    g = {"@context": {"rico": "https://www.ica.org/standards/RiC/ontology#", "nak8": "urn:nak8:", "dcterms": "http://purl.org/dc/terms/"},
         "@id": rid, "@type": "rico:Record", "rico:title": meta["4.1 제목"], "rico:identifier": meta["3.1 기본식별자"],
         "rico:hasCreator": {"@id": f"urn:agent:{meta['1.2 기관명']}/{meta.get('1.3 부서명', '')}", "@type": "rico:CorporateBody",
                             "rico:name": f"{meta['1.2 기관명']} {meta.get('1.3 부서명', '')}".strip()},
         "rico:isOrWasIncludedIn": {"@id": f"urn:task:{meta['12.2.2 분류명']}", "@type": "rico:RecordSet", "rico:name": meta["12.2.2 분류명"],
                                    "rico:hasRecordSetType": "단위과제"},
         "rico:creationDate": meta["13.1 생산일시"], "rico:hasContentOfSomeType": meta["8.1 기록유형"],
         "rico:conditionsOfAccess": meta["17.3.1 공개구분"], "nak8:보존기간": meta["15.1 보존기간"],
         "rico:hasInstantiation": {"@type": "rico:Instantiation", "rico:hasExtent": f"{meta['11.1 용량']} {meta['11.2 단위']}",
                                   "dcterms:format": meta["9.1 포맷명"], "nak8:무결성체크값": meta["23.2 무결성체크값"]}}
    if meta.get("1.6 개인명"):
        g["rico:hasOrHadParticipant"] = {"@type": "rico:Person", "rico:name": meta["1.6 개인명"], "nak8:생산자유형": meta["1.1 생산자유형"]}
    if meta.get("업무활동"):
        g["rico:isAssociatedWithActivity"] = {"@type": "rico:Activity", "rico:name": meta["업무활동"], "rico:history": meta.get("생산목적", "")}
    rel = []
    for r in meta.get("22 관계", []):
        rel.append({"rico:relationType": r["22.1 관계유형"], "@id": f"urn:records:{r['22.2 관계대상식별자']}"})
    if rel:
        g["rico:isRelatedTo"] = rel
    return g


def derive(pkg):
    meta = json.load(open(os.path.join(pkg, "metadata.json"), encoding="utf-8"))
    d = os.path.join(pkg, "매개")
    os.makedirs(d, exist_ok=True)
    open_ = meta["17.3.1 공개구분"] == "공개" and not meta.get("_개인정보후보")
    with open(os.path.join(d, "등록대장.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        cols = ["3.1 기본식별자", "4.1 제목", "1.2 기관명", "1.3 부서명", "1.6 개인명", "13.1 생산일시", "12.2.2 분류명", "15.1 보존기간",
                "17.3.1 공개구분", "17.3.2 공개제한사유", "9.1 포맷명", "11.1 용량", "23.2 무결성체크값"]
        w.writerow(cols)
        w.writerow([meta.get(c, "") for c in cols])
    pub = {k: meta[k] for k in ("3.1 기본식별자", "4.1 제목", "1.2 기관명", "1.3 부서명", "13.1 생산일시", "12.2.2 분류명", "15.1 보존기간", "17.3.1 공개구분", "9.1 포맷명")}
    pub["설명"] = meta.get("5.2 기술내용", "") if open_ else "(비공개·부분공개 또는 개인정보 후보가 있어 본문 요약을 싣지 않음)"
    pub["무결성(SHA-256)"] = meta["23.2 무결성체크값"]
    json.dump(pub, open(os.path.join(d, "공개데이터.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    e = html.escape
    rows = "".join(f"<tr><th>{e(k.split(' ', 1)[-1])}</th><td>{e(str(v))}</td></tr>" for k, v in pub.items())
    page = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{e(meta['4.1 제목'])}</title><style>body{{font:16px/1.6 system-ui,'Malgun Gothic',sans-serif;max-width:720px;margin:24px auto;padding:0 16px;color:#1d1a16;background:#fffdf8}}
h1{{font-size:28px;line-height:1.2}}table{{border-collapse:collapse;width:100%}}th,td{{text-align:left;border-bottom:1px solid #ddd;padding:8px;vertical-align:top}}th{{width:32%;color:#5b5248}}
.note{{font-size:13px;color:#5b5248}}</style></head><body><p class="note">기록 패키지 요약 · {e(meta['1.2 기관명'])}</p><h1>{e(meta['4.1 제목'])}</h1>
<table>{rows}</table><p class="note">이 요약은 metadata.json에서 만든 파생본입니다. 원본은 패키지의 data/원본에 있고, 위 SHA-256으로 같은지 확인할 수 있습니다.</p></body></html>"""
    open(os.path.join(d, "요약.html"), "w", encoding="utf-8").write(page)
    card = [f"📁 {meta['4.1 제목']}", f"{meta['1.2 기관명']} {meta.get('1.3 부서명', '')} · {meta['13.1 생산일시'][:10]}",
            f"단위과제 {meta['12.2.2 분류명']} · 보존 {meta['15.1 보존기간']} · {meta['17.3.1 공개구분']}"]
    if open_ and meta.get("5.2 기술내용"):
        card.append(meta["5.2 기술내용"][:120])
    card.append(f"원본 확인값 {meta['23.2 무결성체크값'][:16]}…")
    open(os.path.join(d, "카드.txt"), "w", encoding="utf-8").write("\n".join(card) + "\n")
    json.dump(context_ld(meta), open(os.path.join(pkg, "context.jsonld"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return open_


def cmd_capture(a):
    src = a.file
    if not os.path.isfile(src):
        print(f"입력 오류: 파일이 없습니다 — {src}", file=sys.stderr)
        return 2
    text = read_text(src)
    if SECRET.search(os.path.basename(src)) or SECRET.search(text[:5000]):
        print("중단: 비밀 표시가 보입니다. 비밀기록물은 이 도구로 패키지화하지 않습니다 — 담당자가 비밀 절차로 처리하세요.", file=sys.stderr)
        return 3
    if a.period not in PERIODS:
        print(f"입력 오류: --period는 {'·'.join(PERIODS)} 가운데 하나", file=sys.stderr)
        return 2
    if a.access != "공개" and not a.reason:
        print("입력 오류: 부분공개·비공개는 --reason(정보공개법 제9조제1항 몇 호와 구체 사유)이 필요합니다.", file=sys.stderr)
        return 2
    cand = {}
    try:
        from extract_candidates import candidates
        c, _ = candidates(src)
        cand = {k: v[0][0] for k, v in c.items() if v}
    except Exception:
        pass
    title = a.title or cand.get("제목") or os.path.splitext(os.path.basename(src))[0]
    made = a.date or cand.get("생산일자") or dt.datetime.fromtimestamp(os.path.getmtime(src), KST).date().isoformat()
    rid = a.id or f"{dt.datetime.now(KST):%Y%m%d}-{uuid.uuid4().hex[:8]}"
    if re.search(r"[\\/:*?\"<>|]", rid):
        print("입력 오류: --id에 파일 이름에 쓸 수 없는 문자가 있습니다.", file=sys.stderr)
        return 2
    pkg = os.path.join(a.out, rid)
    if os.path.exists(pkg):
        print(f"입력 오류: 이미 있는 패키지입니다(덮어쓰지 않음) — {pkg}", file=sys.stderr)
        return 2
    os.makedirs(os.path.join(pkg, "data", "원본"))
    dst = os.path.join(pkg, "data", "원본", os.path.basename(src))
    shutil.copy2(src, dst)
    h = sha256(dst)
    if h != sha256(src):
        print("중단: 복사본 해시가 원본과 다릅니다.", file=sys.stderr)
        return 3
    ext = os.path.splitext(src)[1].lower()
    fmt = FORMATS.get(ext, ext.lstrip(".").upper() or "알 수 없음")
    pii = []
    try:
        from pii_scan import scan_text
        pii = scan_text(text)
    except Exception:
        pass
    auto = [k for k, used in (("제목", not a.title and "제목" in cand), ("생산일시", not a.date and "생산일자" in cand)) if used]
    meta = {"1.1 생산자유형": a.role, "1.2 기관명": a.agency, "1.3 부서명": a.dept, "1.6 개인명": a.person or "",
            "2 기록계층": "건", "3.1 기본식별자": rid, "4.1 제목": title, "5.1 기술유형": "요약" if a.summary else "",
            "5.2 기술내용": a.summary or "", "7 전자기록물여부": "전자", "8.1 기록유형": a.kind,
            "9.1 포맷명": fmt, "11.1 용량": os.path.getsize(dst), "11.2 단위": "Byte", "12.1 분류체계유형": "기록관리기준표",
            "12.2.2 분류명": a.task, "13.1 생산일시": made, "15.1 보존기간": a.period, "16 보존장소": a.place,
            "17.3.1 공개구분": a.access, "17.3.2 공개제한사유": a.reason or "", "23.1 무결성체크법": "SHA-256", "23.2 무결성체크값": h,
            "업무활동": a.activity or "", "생산목적": a.purpose or "", "22 관계": [], "_자동추출후보": auto, "_개인정보후보": len(pii)}
    for r in a.relation or []:
        t, _, target = r.partition(":")
        if target:
            meta["22 관계"].append({"22.1 관계유형": t.strip(), "22.2 관계대상식별자": target.strip()})
    json.dump(meta, open(os.path.join(pkg, "metadata.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    event(pkg, "생산", f"원본 {os.path.basename(src)} · 생산일시 {made}")
    event(pkg, "등록", f"기본식별자 {rid} · 단위과제 {a.task} · 보존기간 {a.period}")
    event(pkg, "메시지 다이제스트 계산", f"SHA-256 {h}")
    open_ = derive(pkg)
    event(pkg, "파생", "매개 4종(등록대장.csv·공개데이터.json·요약.html·카드.txt)과 context.jsonld를 metadata.json에서 만듦")
    write_bag(pkg, meta)
    issues = []
    for k in ("1.2 기관명", "1.3 부서명", "12.2.2 분류명", "15.1 보존기간"):
        if not meta[k]:
            issues.append(f"{k} 비어 있음")
    if a.role in ROLES[:5] and not a.person:
        issues.append("생산자유형이 기안자·검토자·협조자·결재자·접수자이면 1.6 개인명 필수(NAK 8 부속서 D)")
    if auto:
        issues.append(f"자동 추출 후보로 채운 요소: {', '.join(auto)} — 담당자 확인")
    if pii:
        issues.append(f"개인정보 후보 {len(pii)}건 — 공개 매개에서 본문 요약 뺌")
    if fmt not in OPEN_FORMATS:
        issues.append(f"포맷 {fmt}: 장기보존 권고 포맷이 아님 — preservation-convert로 보존 포맷 사본 검토")
    with open(os.path.join(pkg, "점검.md"), "w", encoding="utf-8") as f:
        f.write(f"# 기록 패키지 점검 — {rid}\n\n- 원본 SHA-256: `{h}` (복사본과 일치)\n- 공개 매개 본문 포함: {'예' if open_ else '아니오'}\n")
        f.write("".join(f"- 확인: {x}\n" for x in issues) or "- 확인할 것 없음\n")
        f.write("\n나중에 `python capture.py verify <패키지>`로 고정성을 다시 확인하면 events.jsonl에 사건이 쌓입니다.\n")
    print(f"패키지 {pkg} · SHA-256 {h[:12]}… · 확인 {len(issues)}건")
    return 1 if issues else 0


def cmd_verify(a):
    pkg = a.package
    bad = []
    for name in ("manifest-sha256.txt", "tagmanifest-sha256.txt"):
        p = os.path.join(pkg, name)
        if not os.path.exists(p):
            bad.append(f"{name} 없음")
            continue
        for line in open(p, encoding="utf-8"):
            h, _, rel = line.strip().partition("  ")
            f = os.path.join(pkg, rel)
            if name.startswith("tag") and rel == "events.jsonl":
                continue  # 이력은 덧붙이므로 tagmanifest를 다시 쓴다
            if not os.path.exists(f):
                bad.append(f"없어짐: {rel}")
            elif sha256(f) != h:
                bad.append(f"바뀜: {rel}")
    event(pkg, "고정성 점검", "; ".join(bad) or "모든 파일 해시 일치", "실패" if bad else "성공")
    if not bad:
        write_tags(pkg)  # 이력이 늘었으므로 태그 해시만 새로 씀(원본 해시 목록은 그대로)
    print("고정성 " + ("실패: " + "; ".join(bad) if bad else "통과"))
    return 1 if bad else 0


def cmd_derive(a):
    open_ = derive(a.package)
    meta = json.load(open(os.path.join(a.package, "metadata.json"), encoding="utf-8"))
    event(a.package, "파생", "매개를 metadata.json에서 다시 만듦")
    write_tags(a.package)
    print(f"매개 다시 만듦(공개 본문 {'포함' if open_ else '제외'})")
    return 0


BATCH_COLS = {"file": ["파일", "원본파일", "파일명"], "task": ["단위과제", "단위과제명", "분류명"], "agency": ["기관", "기관명"], "dept": ["부서", "부서명", "처리과"],
              "person": ["담당자", "개인명", "기안자"], "role": ["역할", "생산자유형"], "period": ["보존기간"], "access": ["공개구분", "공개여부"],
              "reason": ["공개제한사유", "비공개사유", "사유"], "activity": ["업무활동", "업무"], "purpose": ["생산목적", "목적"], "summary": ["요약", "기술내용"],
              "relation": ["관계", "관련기록"], "id": ["식별자", "기록식별자", "관리번호"], "title": ["제목"], "date": ["생산일자", "생산일시"], "kind": ["기록유형"]}


def cmd_batch(a):
    """목록(CSV·XLSX) 한 행 = 파일 하나. 파일 경로는 목록이 있는 폴더 기준."""
    import argparse as _ap
    if a.list.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        rows = [["" if v is None else str(v).strip() for v in r] for r in openpyxl.load_workbook(a.list, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(a.list, encoding="utf-8-sig")))
    rows = [r for r in rows if any(r)]
    head = [h.replace(" ", "") for h in rows[0]]
    idx = {k: next((head.index(n) for n in ns if n in head), None) for k, ns in BATCH_COLS.items()}
    if idx["file"] is None:
        print("입력 오류: '파일' 열이 필요합니다.", file=sys.stderr)
        return 2
    base = os.path.dirname(os.path.abspath(a.list))
    results = []
    for n, r in enumerate(rows[1:], 2):
        g = lambda k, d="": (r[idx[k]].strip() if idx[k] is not None and idx[k] < len(r) and r[idx[k]].strip() else d)
        ns = _ap.Namespace(file=os.path.join(base, g("file")), task=g("task", a.task or ""), agency=g("agency", a.agency or ""), dept=g("dept", a.dept or ""),
                           person=g("person"), role=g("role", "기안자"), period=g("period", a.period or ""), access=g("access", "공개"), reason=g("reason"),
                           kind=g("kind", "일반문서"), place="처리과", activity=g("activity"), purpose=g("purpose"), summary=g("summary"),
                           relation=[x for x in re.split(r"[;\n]", g("relation")) if x.strip()] or None, id=g("id") or None, title=g("title") or None,
                           date=g("date") or None, out=a.out)
        if ns.role not in ROLES:
            ns.role = "생산자"
        if not (ns.task and ns.agency and ns.dept and ns.period):
            results.append([n, g("file"), "", "건너뜀", "단위과제·기관·부서·보존기간 중 빈 칸(열 또는 --task 등 기본값)"])
            continue
        import io, contextlib
        buf_o, buf_e = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(buf_o), contextlib.redirect_stderr(buf_e):
            rc = cmd_capture(ns)
        msg = (buf_o.getvalue() + buf_e.getvalue()).strip().replace("\n", " ")
        state = {0: "만듦", 1: "만듦(확인 있음)", 2: "입력 오류", 3: "거부"}.get(rc, str(rc))
        pid = re.search(r"패키지 (\S+)", msg)
        results.append([n, g("file"), pid.group(1) if pid else "", state, msg[:300]])
    os.makedirs(a.out, exist_ok=True)
    with open(os.path.join(a.out, "일괄_결과.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["목록 행", "파일", "패키지", "결과", "메시지"])
        w.writerows(results)
    made = sum(1 for x in results if x[3].startswith("만듦"))
    print(f"{len(results)}행 · 패키지 {made} · 거부·오류·건너뜀 {len(results) - made} → {os.path.join(a.out, '일괄_결과.csv')}")
    return 0 if made == len(results) else 1


def cmd_catalog(a):
    """패키지 폴더 전체에서 목록·공개데이터·연결 그래프·검색 페이지를 만든다(매개 확산). --verify면 고정성도 다시 본다."""
    pkgs = sorted(d for d in (os.path.join(a.root, x) for x in os.listdir(a.root)) if os.path.isfile(os.path.join(d, "metadata.json")))
    os.makedirs(a.out, exist_ok=True)
    metas, graph, fixity = [], [], {}
    for p in pkgs:
        m = json.load(open(os.path.join(p, "metadata.json"), encoding="utf-8"))
        metas.append(m)
        ld = os.path.join(p, "context.jsonld")
        if os.path.exists(ld):
            g = json.load(open(ld, encoding="utf-8"))
            g.pop("@context", None)
            graph.append(g)
        if a.verify:
            import io, contextlib, argparse as _ap
            with contextlib.redirect_stdout(io.StringIO()):
                fixity[m["3.1 기본식별자"]] = "통과" if cmd_verify(_ap.Namespace(package=p)) == 0 else "실패"
    cols = ["3.1 기본식별자", "4.1 제목", "1.2 기관명", "1.3 부서명", "13.1 생산일시", "12.2.2 분류명", "15.1 보존기간", "17.3.1 공개구분", "9.1 포맷명", "23.2 무결성체크값"]
    with open(os.path.join(a.out, "기록목록.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols + ["관계", "고정성"])
        for m in metas:
            rel = "; ".join(f"{x['22.1 관계유형']}:{x['22.2 관계대상식별자']}" for x in m.get("22 관계", []))
            w.writerow([m.get(c, "") for c in cols] + [rel, fixity.get(m["3.1 기본식별자"], "")])
    pub = [m for m in metas if m.get("17.3.1 공개구분") == "공개" and not m.get("_개인정보후보")]
    with open(os.path.join(a.out, "공개데이터.jsonl"), "w", encoding="utf-8") as f:
        for m in pub:
            f.write(json.dumps({k: m.get(k, "") for k in cols[:-1]} | {"설명": m.get("5.2 기술내용", ""), "SHA-256": m["23.2 무결성체크값"]}, ensure_ascii=False) + "\n")
    json.dump({"@context": {"rico": "https://www.ica.org/standards/RiC/ontology#", "nak8": "urn:nak8:", "dcterms": "http://purl.org/dc/terms/"}, "@graph": graph},
              open(os.path.join(a.out, "연결그래프.jsonld"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    ids = {m["3.1 기본식별자"] for m in metas}
    e = html.escape
    rows = []
    for m in sorted(metas, key=lambda x: x.get("13.1 생산일시", "")):
        rid = m["3.1 기본식별자"]
        rel = " ".join(f'<a href="#{e(x["22.2 관계대상식별자"])}">{e(x["22.1 관계유형"])}: {e(x["22.2 관계대상식별자"])}</a>' if x["22.2 관계대상식별자"] in ids
                       else f'{e(x["22.1 관계유형"])}: {e(x["22.2 관계대상식별자"])}' for x in m.get("22 관계", []))
        open_ = m in pub
        rows.append(f'<tr id="{e(rid)}" data-q="{e((m["4.1 제목"] + m["12.2.2 분류명"] + m.get("1.3 부서명", "")).lower())}"><td data-l="생산일">{e(m.get("13.1 생산일시", "")[:10])}</td>'
                    f'<td data-l="제목"><b>{e(m["4.1 제목"])}</b><br><small>{e(rid)} · {e(m.get("1.3 부서명", ""))}</small>'
                    + (f'<br><small>{e(m.get("5.2 기술내용", ""))}</small>' if open_ and m.get("5.2 기술내용") else "")
                    + f'</td><td data-l="단위과제">{e(m["12.2.2 분류명"])}</td><td data-l="보존">{e(m["15.1 보존기간"])}</td><td data-l="공개">{e(m["17.3.1 공개구분"])}</td><td data-l="관계">{rel}</td>'
                    f'<td data-l="SHA-256·고정성"><code title="{e(m["23.2 무결성체크값"])}">{e(m["23.2 무결성체크값"][:12])}…</code> {e(fixity.get(rid, ""))}</td></tr>')
    page = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>기록 패키지 목록</title>
<style>body{{font:15px/1.55 system-ui,'Malgun Gothic',sans-serif;margin:24px auto;max-width:1100px;padding:0 16px;color:#1d1a16;background:#fffdf8}}
table{{border-collapse:collapse;width:100%}}th,td{{border-bottom:1px solid #ddd;padding:8px;text-align:left;vertical-align:top}}th{{color:#5b5248;font-size:13px}}
input{{font:inherit;font-size:16px;padding:8px 12px;width:min(100%,360px);margin:8px 0 16px}}small{{color:#5b5248}}tr:target{{background:#fff1c9}}
@media(max-width:640px){{thead{{display:none}}tr{{display:block;border-bottom:1px solid #ddd;padding:10px 0}}td{{display:grid;grid-template-columns:6.5em 1fr;gap:8px;border:0;padding:3px 0}}td::before{{content:attr(data-l);color:#5b5248;font-size:12px}}td:empty{{display:none}}}}</style></head><body>
<h1>기록 패키지 목록</h1><p><small>{len(metas)}건 · 공개 {len(pub)}건 · 각 행은 패키지의 metadata.json에서 만들었습니다. 관계 링크를 누르면 이어진 기록으로 갑니다.</small></p>
<label>찾기 <input id="q" type="search" placeholder="제목·단위과제·부서"></label>
<table><thead><tr><th>생산일</th><th>제목</th><th>단위과제</th><th>보존</th><th>공개</th><th>관계</th><th>SHA-256·고정성</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
<script>document.getElementById('q').addEventListener('input',e=>{{const v=e.target.value.toLowerCase();for(const r of document.querySelectorAll('tbody tr'))r.hidden=v&&!r.dataset.q.includes(v)}})</script></body></html>"""
    open(os.path.join(a.out, "목록.html"), "w", encoding="utf-8").write(page)
    bad = sum(1 for v in fixity.values() if v == "실패")
    print(f"패키지 {len(metas)} · 공개데이터 {len(pub)} · 그래프 노드 {len(graph)}" + (f" · 고정성 실패 {bad}" if a.verify else "") + f" → {a.out}")
    return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser(description="생산 시점 기록화 — 맥락·증거성을 갖춘 기록 패키지")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("capture")
    c.add_argument("file")
    c.add_argument("--task", required=True, help="단위과제명(12.2.2 분류명)")
    c.add_argument("--agency", required=True)
    c.add_argument("--dept", required=True)
    c.add_argument("--person", default="")
    c.add_argument("--role", default="기안자", choices=ROLES)
    c.add_argument("--period", required=True)
    c.add_argument("--access", default="공개", choices=["공개", "부분공개", "비공개"])
    c.add_argument("--reason", default="")
    c.add_argument("--kind", default="일반문서")
    c.add_argument("--place", default="처리과")
    c.add_argument("--activity", default="", help="이 기록을 만든 업무 행위(맥락)")
    c.add_argument("--purpose", default="", help="생산 목적")
    c.add_argument("--summary", default="", help="기술내용(요약)")
    c.add_argument("--relation", action="append", help="관계유형:대상식별자 (여러 번)")
    c.add_argument("--id")
    c.add_argument("--title")
    c.add_argument("--date", help="생산일시 YYYY-MM-DD")
    c.add_argument("-o", "--out", default="산출/기록패키지")
    v = sub.add_parser("verify")
    v.add_argument("package")
    d = sub.add_parser("derive")
    d.add_argument("package")
    b = sub.add_parser("batch", help="목록 한 행 = 파일 하나로 여러 패키지")
    b.add_argument("list")
    for k in ("task", "agency", "dept", "period"):
        b.add_argument(f"--{k}", help="목록에 빈 칸일 때 쓸 기본값")
    b.add_argument("-o", "--out", default="산출/기록패키지")
    g = sub.add_parser("catalog", help="패키지 폴더 전체로 목록·공개데이터·연결 그래프·검색 페이지")
    g.add_argument("root")
    g.add_argument("--verify", action="store_true")
    g.add_argument("-o", "--out", default="산출/기록목록")
    a = ap.parse_args()
    return {"capture": cmd_capture, "verify": cmd_verify, "derive": cmd_derive, "batch": cmd_batch, "catalog": cmd_catalog}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
