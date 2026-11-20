#!/usr/bin/env python3
"""이메일 기록 평가 보조: MBOX 파일이나 EML 폴더를 읽어 메시지 목록·상대 집계·첨부 목록·개인정보·비밀 표시를 정리한다.

사용:  python email_report.py 보관함.mbox -o 산출/이메일평가.xlsx
       python email_report.py eml폴더/      -o 산출/이메일평가.xlsx
산출 시트: 요약 · 메시지 목록 · 주고받은 상대 · 첨부 목록
원칙:  본문은 결과에 옮기지 않는다(메타데이터와 탐지 결과만). 개인정보 값은 가려서 센다.
       어떤 메일이 기록물인지, 무엇을 남길지는 사람이 정한다 — 이 표는 평가를 돕는 목록이다.
참고:  ePADD(이메일 평가·정리·공개 도구)의 평가 모듈이 하는 일 중 목록화·상대 분석·민감정보 선별 부분을 가볍게 옮긴 것.
"""
import argparse
import collections
import email
import email.policy
import hashlib
import mailbox
import os
import re
import sys
from email.utils import getaddresses, parsedate_to_datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "pii-redact", "scripts"))
try:
    from pii_scan import scan_text
except ImportError:
    scan_text = None

SECRET_RX = re.compile(r"(Ⅰ|Ⅱ|Ⅲ|I{1,3}|[123])\s?급\s?비밀|대외비")


def messages(src):
    if os.path.isdir(src):
        for root, _, fs in os.walk(src):
            for fn in sorted(fs):
                if fn.lower().endswith(".eml"):
                    with open(os.path.join(root, fn), "rb") as f:
                        yield os.path.relpath(os.path.join(root, fn), src), email.message_from_binary_file(f, policy=email.policy.default)
    else:
        for i, m in enumerate(mailbox.mbox(src), 1):
            yield f"#{i}", email.message_from_bytes(m.as_bytes(), policy=email.policy.default)


def body_text(msg):
    parts = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart" or part.get_filename():
            continue
        if part.get_content_type() in ("text/plain", "text/html"):
            try:
                t = part.get_content()
            except Exception:
                t = (part.get_payload(decode=True) or b"").decode("utf-8", "replace")
            if part.get_content_type() == "text/html":
                t = re.sub(r"<[^>]+>", " ", t)
            parts.append(t)
    return "\n".join(parts)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--roles", help="역할기반(캡스톤) 목록 csv/xlsx: 주소, 직위[, 구분] — 구분이 비면 '주요 직위'로 봄")
    ap.add_argument("-o", "--out", default="산출/이메일평가.xlsx")
    a = ap.parse_args()
    roles = {}
    if a.roles:
        if a.roles.lower().endswith((".xlsx", ".xlsm")):
            import openpyxl as _ox
            rr = [["" if v is None else str(v).strip() for v in r] for r in _ox.load_workbook(a.roles, read_only=True, data_only=True).active.iter_rows(values_only=True)]
        else:
            import csv
            rr = list(csv.reader(open(a.roles, encoding="utf-8-sig")))
        h = [re.sub(r"\s", "", x) for x in rr[0]]
        ai = next((h.index(x) for x in ("주소", "이메일", "메일주소", "email") if x in h), 0)
        pi = next((h.index(x) for x in ("직위", "직급", "역할") if x in h), None)
        gi = next((h.index(x) for x in ("구분", "분류") if x in h), None)
        for r in rr[1:]:
            if len(r) > ai and r[ai]:
                roles[r[ai].strip().lower()] = ((r[pi] if pi is not None and pi < len(r) else ""), (r[gi] if gi is not None and gi < len(r) and r[gi] else "주요 직위"))
    rows, atts, people = [], [], collections.Counter()
    dates = []
    for where, msg in messages(a.src):
        try:
            d = parsedate_to_datetime(msg["Date"]) if msg["Date"] else None
        except Exception:
            d = None
        if d:
            dates.append(d)
        frm = getaddresses([str(msg.get("From", ""))])
        to = getaddresses([str(msg.get("To", "")), str(msg.get("Cc", ""))])
        for _, addr in frm + to:
            if addr:
                people[addr.lower()] += 1
        text = body_text(msg)
        subj = str(msg.get("Subject", ""))
        hits = scan_text(subj + "\n" + text) if scan_text else []
        secret = bool(SECRET_RX.search(subj + text))
        n_att = 0
        for part in msg.walk():
            fn = part.get_filename()
            if fn:
                data = part.get_payload(decode=True) or b""
                atts.append([where, subj[:60], fn, len(data), hashlib.sha256(data).hexdigest()])
                n_att += 1
        role = ""
        if roles:
            hit = [(addr.lower(), roles[addr.lower()]) for _, addr in frm + to if addr and addr.lower() in roles]
            if hit:
                role = f"역할기반 보존 후보 — {hit[0][1][0] or hit[0][0]}({hit[0][1][1]})" + (" 발신" if hit[0][0] in {x[1].lower() for x in frm} else " 수신")
            else:
                role = "일반 계정 — 기관 기준으로 판단"
        rows.append([where, d.isoformat(timespec="minutes") if d else "", ", ".join(x[1] for x in frm),
                     ", ".join(x[1] for x in to)[:200], subj, n_att,
                     ", ".join(sorted({h[0] for h in hits})), "비밀 표시 — 담당자 확인" if secret else "",
                     str(msg.get("Message-ID", ""))] + ([role] if roles else []))
    import openpyxl
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "요약"
    s.append(["메시지", len(rows)])
    s.append(["기간", f"{min(dates):%Y-%m-%d} ~ {max(dates):%Y-%m-%d}" if dates else "날짜 없음"])
    s.append(["주고받은 주소", len(people)])
    s.append(["첨부", len(atts)])
    s.append(["개인정보 후보가 있는 메시지", sum(1 for r in rows if r[6])])
    s.append(["비밀 표시 메시지", sum(1 for r in rows if r[7])])
    s.append([])
    s.append(["주의", "본문은 옮기지 않았습니다. 기록물 여부·보존 판단은 사람이 합니다."])
    if roles:
        s.append(["역할기반 보존 후보 메시지", sum(1 for r in rows if r[9].startswith("역할기반"))])
        s.append(["역할기반(캡스톤) 참고", "미국 NARA GRS 6.1(이메일·전자메시지 역할기반 관리): 기관장·차관급 등 지정한 주요 직위 계정의 메일은 영구, 그 밖의 계정은 한시로 관리하는 방식. "
                  "국내 법정 기준이 아니며 국가기록원 「전자기록 등 기타 특수유형 기록관리 체계 구축을 위한 국가기록원 기록물 평가정책 연구」(2024)가 소개한 참고 방식이다. 보존기간은 기관 기록관리표·평가 절차로 정한다."])
    m = wb.create_sheet("메시지 목록")
    m.append(["위치", "날짜", "보낸 사람", "받는 사람(참조 포함)", "제목", "첨부 수", "개인정보 후보", "비밀", "Message-ID"] + (["역할기반 구분"] if roles else []))
    for r in rows:
        m.append(r)
    p = wb.create_sheet("주고받은 상대")
    p.append(["주소", "등장 횟수"])
    for addr, n in people.most_common():
        p.append([addr, n])
    t = wb.create_sheet("첨부 목록")
    t.append(["메시지", "제목", "파일명", "바이트", "sha256"])
    for x in atts:
        t.append(x)
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(f"메시지 {len(rows)} · 상대 {len(people)} · 첨부 {len(atts)} · 개인정보 후보 메시지 {sum(1 for r in rows if r[6])} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
