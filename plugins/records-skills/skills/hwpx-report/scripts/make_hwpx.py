#!/usr/bin/env python3
"""보고서·공문 HWPX 만들기: 마크다운 본문을 한글(HWPX) 문서로 만들거나, 기관 서식(.hwpx)의 자리표시를 채운다.

사용:
  새 문서   python make_hwpx.py --md 보고.md -o 산출/보고.hwpx [--header "가상시 기록관"] [--page-number]
  서식 채움 python make_hwpx.py --template 기관서식.hwpx --values 값.xlsx [--md 본문.md] -o 산출/공문.hwpx
값 파일:  2열(키, 값) xlsx·csv 또는 JSON 객체.
          키 '제목'      → 서식 안 {{제목}} 을 값으로 바꾼다.
          키 '칸:수신자' → 서식 표에서 '수신자' 칸의 오른쪽 칸에 값을 넣는다.
마크다운: 제목(#~####), 문단, -/1. 목록, | 표 |, **굵게**. (python-hwpx text.append_markdown 범위)
점검:    저장 후 다시 열어 ① 스키마 검사(validate) ② 남은 {{…}} ③ 원문 줄이 모두 들어갔는지 ④ 개인정보 후보를 본다.
원칙:    결과 파일이 이미 있으면 덮어쓰지 않는다. 결재·발송은 사람이 한다. 한글 프로그램에서 열어 모양을 확인하기 전에는 완성이라 하지 않는다.
필요:    pip install python-hwpx (Apache-2.0)
"""
import argparse
import csv
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning)
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "pii-redact", "scripts"))
try:
    from pii_scan import scan_text
except ImportError:
    scan_text = None


def load_values(path):
    if not path:
        return {}
    if path.lower().endswith(".json"):
        return {str(k): str(v) for k, v in json.load(open(path, encoding="utf-8")).items()}
    if path.lower().endswith(".csv"):
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    else:
        import openpyxl
        rows = [["" if c is None else str(c) for c in r] for r in openpyxl.load_workbook(path, data_only=True).active.iter_rows(values_only=True)]
    if rows and rows[0][:2] in (["키", "값"], ["key", "value"]):
        rows = rows[1:]
    return {r[0].strip(): (r[1] if len(r) > 1 else "") for r in rows if r and r[0].strip()}


def md_lines(md):
    """본문 확인용: 마크다운 표시를 걷어낸 줄들."""
    out = []
    for line in md.splitlines():
        s = line.strip()
        if not s or re.fullmatch(r"\|?\s*:?-{3,}.*", s):
            continue
        s = re.sub(r"^#{1,6}\s*|^[-*]\s+|^\d+\.\s+|\*\*|\*", "", s)
        if s.startswith("|"):
            out += [c.strip() for c in s.strip("|").split("|") if c.strip()]
        else:
            out.append(s)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--md")
    ap.add_argument("--template")
    ap.add_argument("--values")
    ap.add_argument("--header")
    ap.add_argument("--page-number", action="store_true")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    if not a.md and not a.template:
        print("입력 오류: --md 또는 --template 중 하나는 있어야 합니다.")
        return 2
    if os.path.exists(a.out):
        print(f"덮어쓰지 않음: {a.out} 이(가) 이미 있습니다.")
        return 2
    try:
        from hwpx.document import HwpxDocument
    except ImportError:
        print("python-hwpx가 없습니다: pip install python-hwpx")
        return 2

    doc = HwpxDocument.open(a.template) if a.template else HwpxDocument.new()
    vals = load_values(a.values)
    filled, missing_label = 0, []
    for k, v in vals.items():
        if k.startswith("칸:"):
            label = k[2:].strip()
            hits = doc.tables.find_cell_by_label(label).get("matches", [])
            if not hits:
                missing_label.append(label)
                continue
            tables = doc.tables.all if isinstance(doc.tables.all, list) else list(doc.tables.all())
            for h in hits:
                t = tables[h["table_index"]]
                t.set_cell_text(h["target_cell"]["row"], h["target_cell"]["col"], v)
                filled += 1
        else:
            filled += doc.replace_text_in_runs("{{" + k + "}}", v)
    md = open(a.md, encoding="utf-8").read() if a.md else ""
    if md:
        doc.text.append_markdown(md)
    if a.header:
        doc.set_header_text(a.header)
    if a.page_number:
        doc.page.set_page_number()
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    doc.save_to_path(a.out)

    chk = HwpxDocument.open(a.out)
    rep = chk.validate()
    plain = chk.text.plain()
    flat = re.sub(r"\s+", "", plain)
    left = sorted(set(re.findall(r"\{\{[^}]+\}\}", plain)))
    lost = [l for l in md_lines(md) if re.sub(r"\s+", "", l) not in flat]
    pii = scan_text(plain) if scan_text else []
    report = [f"# HWPX 점검 — {os.path.basename(a.out)}", "",
              f"- 스키마 검사(python-hwpx validate): {'통과' if rep.ok else '실패'}" + ("" if rep.ok else f" — {[str(i) for i in rep.issues][:5]}"),
              f"- 채운 값: {filled}개" + (f", 서식에 없는 칸: {', '.join(missing_label)}" if missing_label else ""),
              f"- 남은 자리표시: {', '.join(left) if left else '없음'}",
              f"- 본문 누락 줄: {len(lost)}" + ("".join(f"\n  - {l[:60]}" for l in lost[:10]) if lost else ""),
              f"- 개인정보 후보: {len(pii)}" + ("".join(f"\n  - {p[0]} {p[1]}" for p in pii[:10]) if pii else ""),
              "- 모양 확인: 한글 프로그램에서 열어 보기 전에는 완성이 아님(이 도구는 글자·구조만 점검)"]
    rp = os.path.splitext(a.out)[0] + "_점검.md"
    open(rp, "w", encoding="utf-8").write("\n".join(report) + "\n")
    ok = rep.ok and not left and not lost
    print(f"{'통과' if ok else '확인 필요'} · 값 {filled} · 남은 자리표시 {len(left)} · 누락 {len(lost)} · 개인정보 후보 {len(pii)} → {a.out}, {rp}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
