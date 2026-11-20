#!/usr/bin/env python3
"""장기보존패키지(NEO2) 구조 점검: 이관·보존용 NEO XML 파일이 NAK 31-1:2022(v2.3) 「전자기록물 장기보존패키지 기술규격 - 제1부
XML로 포맷화된 방식(NEO2)」의 필수 요소를 갖췄는지, 컴포넌트 데이터가 Base64로 풀리고 크기가 맞는지, 전자서명 블록이 있는지 본다.
근거: NAK 31-1:2022 6장 요소 정의(필수 여부) (국가기록원, 2026-11-20 내려받아 확인). 요소 이름 철자(Souce… 등)는 원문 그대로.

NEO3(NAK 31-2:2022 v1.1 「제2부: 디렉토리로 구조화된 방식」 부속서 A 필수 조건): 폴더나 .zip을 주면 최상위 Metadata.xml·PackagingInformation.xml과
  각 XSD(Metadata.xsd·PackagingInformation.xsd), ZIP이면 첫 항목 'mimetype'(비압축, 내용 application/neo3+zip),
  PackagingInformation.xml의 PackageManifest/ManifestItem(Path·HashValue·HashAlgorithm)대로 파일이 있고 해시가 맞는지 본다.

사용:  python neo_check.py NEO폴더/ -o 산출/NEO점검.csv
       python neo_check.py 패키지.zip(또는 NEO3 폴더) -o 산출/NEO3점검.csv
원칙:  읽기만 한다. 서명·인증서의 유효성(장기검증, NAK 32-1)은 검증하지 않는다 — 기록관리시스템·전용 검증기로 한다.
       필수 요소 점검은 '부모 요소가 있을 때 그 안에 있어야 할 요소'를 본다(해당시 필수 요소는 선택으로 정의돼 있어 보지 않음).
"""
import argparse
import base64
import csv
import os
import re
import sys
import xml.etree.ElementTree as ET

ALWAYS = ["NAKEncapsulatedObject", "NEOFormatDescription", "Version", "SignedObject", "ObjectCreationDate", "ObjectContent", "Record", "RecordMetadata"]
CHILD = {  # 부모가 있으면 있어야 할 자식(필수)
    "CreatorCon": ["CreatorType"], "RecordIdentifierCon": ["FileID", "RecordID"], "SystemID": ["SystemIDType", "SystemIDValue"],
    "TitleCon": ["MainTitle"], "AlternativeTitleCon": ["AlternativeTitleType", "AlternativeTitleWords"], "DescriptionCon": ["DescriptionType", "DescriptionText"],
    "SubjectCon": ["SubjectType", "SubjectWords"], "TypeCon": ["RecordType"], "ExtentCon": ["ExtentType", "Size", "Unit"],
    "StructureCon": ["StructureType"], "StructureInfoCon": ["TargetObjectID", "TargetObjectInfo"], "ClassificationCon": ["ClassificationType"],
    "ClassificationValueCon": ["ClassificationID", "ClassificationName", "ClassificationLevel"],
    "CreationHistoryCon": ["RecordCreationSystem", "CreationType"], "CreationSectionCon": ["CreationSectionName"],
    "DateTimeCon": ["DateTimeCreated", "DateTimeClosed"], "RetentionPeriodCon": ["RetentionPeriod", "PreservationPlace"],
    "SecurityCon": ["SecurityLevel", "MandateofSecurityLevel", "ProtectionPeriod"], "ExternalAccessCon": ["ExternalAccessControl"],
    "ManagementHistoryCon": ["EventType", "EventDateTime"], "EventAgentCon": ["EventCorporateName", "EventCorporateCode", "EventPersonName"],
    "ChangedElementCon": ["ChangedElementName", "ChangedElementValue"], "PreservationHistoryCon": ["PreservationActionType", "PreservationActionDate"],
    "PreservationActionAgentCon": ["PreservationCorporateName", "PreservationCorporateCode"], "RelationCon": ["RelationType", "RelationItemID"],
    "ComponentMetadata": ["ComponentTitle", "ComponentType"], "ComponentRelationCon": ["ComponentRelationType", "ComponentRelationItemID"],
    "EncodingMetadata": ["SourceFileIdentifier", "SouceFileFormatName", "SourceFileExtent"], "Encoding": ["ComponentData"],
    "SignatureBlock": ["SignatureFormatDescription", "SignatureDate", "Signer", "Signature"], "CertificateBlock": ["Certificate"],
    "SignatureAlgorithm": ["SignatureAlgorithmIdentifier"],
}


def local(tag):
    return tag.rsplit("}", 1)[-1].split(":")[-1]


def check(path):
    issues, info = [], {}
    try:
        tree = ET.parse(path)
    except ET.ParseError as e:
        return [f"XML 오류: {str(e)[:80]}"], info
    root = tree.getroot()
    if local(root.tag) != "NAKEncapsulatedObject":
        issues.append(f"루트 요소 {local(root.tag)} — neo:NAKEncapsulatedObject여야 함")
    names = {local(e.tag) for e in root.iter()}
    for n in ALWAYS:
        if n not in names:
            issues.append(f"필수 요소 없음: neo:{n}")
    miss = set()
    for e in root.iter():
        req = CHILD.get(local(e.tag))
        if req:
            have = {local(c.tag) for c in e.iter()}
            for r in req:
                if r not in have:
                    miss.add(f"{local(e.tag)}→{r}")
    if miss:
        issues.append("부모 안에 필수 요소 없음: " + ", ".join(sorted(miss)[:8]) + (" 외" if len(miss) > 8 else ""))
    comps = [e for e in root.iter() if local(e.tag) == "Component"]
    info["컴포넌트"] = len(comps)
    if not comps:
        issues.append("Component(본문·첨부 파일) 없음")
    bad_b64 = size_diff = 0
    for c in comps:
        data = next((e for e in c.iter() if local(e.tag) == "ComponentData"), None)
        ext = next((e for e in c.iter() if local(e.tag) == "SourceFileExtent"), None)
        if data is None or not (data.text or "").strip():
            continue
        try:
            raw = base64.b64decode(re.sub(r"\s", "", data.text), validate=True)
        except Exception:
            bad_b64 += 1
            continue
        if ext is not None and (ext.text or "").strip().isdigit() and int(ext.text.strip()) != len(raw):
            size_diff += 1
    if bad_b64:
        issues.append(f"ComponentData Base64 해독 실패 {bad_b64}개")
    if size_diff:
        issues.append(f"SourceFileExtent와 실제 크기가 다른 컴포넌트 {size_diff}개")
    if "SignatureBlock" not in names:
        issues.append("전자서명 블록(SignatureBlock) 없음 — 진본성 확인 불가")
    v = next((e.text for e in root.iter() if local(e.tag) == "Version"), "")
    info["버전"] = (v or "").strip()
    return issues, info


NEO3_TOP = ("Metadata.xml", "PackagingInformation.xml", "Metadata.xsd", "PackagingInformation.xsd")
HASHES = {"MD5": "md5", "SHA1": "sha1", "SHA2-224": "sha224", "SHA2-256": "sha256", "SHA2-384": "sha384", "SHA2-512": "sha512"}
WEAK = {"MD2", "MD4", "MD5", "SHA1"}


def neo3(src):
    """NEO3 패키지(폴더 또는 zip) 점검 → [(대상, 문제)]"""
    import hashlib
    import zipfile
    probs, files = [], {}
    if os.path.isfile(src):
        try:
            z = zipfile.ZipFile(src)
        except zipfile.BadZipFile:
            return [("(패키지)", "ZIP으로 열리지 않음")], 0
        infos = z.infolist()
        if not infos or infos[0].filename != "mimetype":
            probs.append(("mimetype", "ZIP 첫 항목이 'mimetype'이 아님(부속서 A 4.1)"))
        mi = next((i for i in infos if i.filename == "mimetype"), None)
        if mi:
            if mi.compress_type != zipfile.ZIP_STORED:
                probs.append(("mimetype", "mimetype이 압축돼 있음 — 비압축으로 넣어야 함(4.1)"))
            if z.read(mi).decode("utf-8", "replace").strip() != "application/neo3+zip":
                probs.append(("mimetype", "내용이 application/neo3+zip이 아님(4.2)"))
        for i in infos:
            if not i.is_dir():
                files[i.filename] = (lambda n=i.filename: z.read(n))
    else:
        for d, _, fs in os.walk(src):
            for f in fs:
                full = os.path.join(d, f)
                files[os.path.relpath(full, src).replace(os.sep, "/")] = (lambda q=full: open(q, "rb").read())
    for t in NEO3_TOP:
        if t not in files:
            probs.append((t, "최상위에 없음(부속서 A 1~3: 파일 이름·위치 고정)"))
    items = 0
    if "PackagingInformation.xml" in files:
        try:
            root = ET.fromstring(files["PackagingInformation.xml"]())
        except ET.ParseError as e:
            return probs + [("PackagingInformation.xml", f"XML 파싱 실패: {e}")], 0
        ver = next((el.text or "" for el in root.iter() if local(el.tag) == "Version"), "")
        if ver.strip() != "3.0":
            probs.append(("PackagingInformation.xml", f"Version이 3.0이 아님({ver.strip() or '없음'})"))
        default_alg = root.get("HashAlgorithm", "")
        listed = set()
        for it in (el for el in root.iter() if local(el.tag) == "ManifestItem"):
            items += 1
            kids = {local(c.tag): (c.text or "").strip() for c in it}
            path, hv, alg = kids.get("Path", ""), kids.get("HashValue", ""), it.get("HashAlgorithm") or default_alg
            if not path:
                if kids.get("UniqueID"):
                    continue  # 외부(원격) 파일 — 패키지 안에서 확인 불가
                probs.append((f"ManifestItem {items}", "Path·UniqueID가 모두 없음"))
                continue
            path = path.lstrip("./").replace("\\", "/")
            listed.add(path)
            if path not in files:
                probs.append((path, "목록에 있으나 패키지에 없음"))
                continue
            if not hv:
                probs.append((path, "HashValue 없음"))
            elif alg not in HASHES:
                probs.append((path, f"HashAlgorithm '{alg or '미표기'}' — 해시 대조 불가"))
            else:
                h = hashlib.new(HASHES[alg], files[path]())
                if hv.lower() not in (h.hexdigest(), base64.b64encode(h.digest()).decode().lower()):
                    probs.append((path, f"해시 불일치({alg})"))
            if alg in WEAK:
                probs.append((path, f"약한 해시 알고리즘 {alg} — SHA2 계열 권장(검토)"))
        for f in files:
            if f not in listed and f not in NEO3_TOP and f != "mimetype":
                probs.append((f, "PackageManifest 목록에 없는 파일(검토)"))
    return probs, items


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--out", default="산출/NEO점검.csv")
    a = ap.parse_args()
    is3 = (os.path.isfile(a.folder) and a.folder.lower().endswith((".zip", ".neo3"))) or (
        os.path.isdir(a.folder) and "PackagingInformation.xml" in os.listdir(a.folder))
    if is3:
        probs, items = neo3(a.folder)
        os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
        hard = [p for p in probs if "(검토)" not in p[1]]
        with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["대상", "문제", "구분"])
            for tgt, msg in probs:
                w.writerow([tgt, msg, "검토" if "(검토)" in msg else "고칠 것"])
            if not probs:
                w.writerow(["(패키지)", f"NEO3 필수 조건 통과 · 목록 {items}건 해시 일치", "통과"])
        print(f"NEO3 목록 {items}건 · 고칠 것 {len(hard)} · 검토 {len(probs) - len(hard)} → {a.out}")
        return 1 if hard else 0
    files = [a.folder] if os.path.isfile(a.folder) else [os.path.join(d, f) for d, _, fs in os.walk(a.folder) for f in sorted(fs) if f.lower().endswith((".xml", ".neo"))]
    rows = []
    for p in files:
        if os.path.basename(p) in ("Metadata.xml", "PackagingInformation.xml"):
            continue
        iss, info = check(p)
        rows.append([os.path.relpath(p, a.folder) if os.path.isdir(a.folder) else os.path.basename(p), info.get("버전", ""), info.get("컴포넌트", ""),
                     len(iss), " · ".join(iss), "통과" if not iss else "고칠 것 있음"])
    top = os.listdir(a.folder) if os.path.isdir(a.folder) else []
    note = ""
    if top and any(x in top for x in ("Metadata.xml", "PackagingInformation.xml")) and not all(x in top for x in ("Metadata.xml", "PackagingInformation.xml")):
        note = "압축 패키지는 최상위에 Metadata.xml과 PackagingInformation.xml이 함께 있어야 함"
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["파일", "NEO 버전", "컴포넌트 수", "문제 수", "문제", "판정"])
        w.writerows(rows)
        if note:
            w.writerow(["(패키지)", "", "", 1, note, "고칠 것 있음"])
    bad = sum(1 for r in rows if r[3]) + bool(note)
    print(f"NEO {len(rows)}개 · 고칠 것 있음 {bad} → {a.out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
