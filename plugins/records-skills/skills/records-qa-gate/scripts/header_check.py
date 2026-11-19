#!/usr/bin/env python3
"""열 맞춤 점검: 기관 시스템에서 내보낸 목록의 '머리글 행만' 읽어, 킷의 각 스크립트가 어느 열을 무엇으로 알아보는지 보여 준다.
띄어쓰기·괄호·별표·'년도/연도' 차이로 못 알아보는 열은 '비슷함'으로 알리고, --fix면 머리글만 바꾼 사본을 만든다.

사용:  python header_check.py 내보낸목록.xlsx [--script production_status] [--fix -o 사본.xlsx]
       --script를 빼면 ALIAS를 가진 모든 스크립트에 대해 본다(스크립트 이름 일부만 써도 됨).
       --map "처리과명=생산부서,철명=기록물철제목" 처럼 직접 짝을 지어 줄 수도 있다(--fix와 함께).
원칙:  자료 행은 읽지 않고 화면에도 내지 않는다(머리글 행과 그 위 몇 줄만). 원본은 고치지 않는다. --fix 사본도 자료는 그대로 복사.
       비밀 폴더의 파일은 다루지 않는다.
"""
import argparse
import ast
import glob
import os
import re
import sys

SK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def norm(s):
    s = str(s or "")
    s = re.sub(r"[\s\*·∙ㆍ_\-/:：]", "", s)
    s = re.sub(r"[①-⑳]", "", s)
    s = s.replace("년도", "연도").replace("（", "(").replace("）", ")").replace("[", "(").replace("]", ")")
    return s.lower()


def aliases():
    """각 스크립트 원문에서 ALIAS = {...} 사전만 꺼낸다(실행하지 않음)."""
    out = {}
    for p in sorted(glob.glob(os.path.join(SK, "*", "scripts", "*.py"))):
        try:
            tree = ast.parse(open(p, encoding="utf-8").read())
        except SyntaxError:
            continue
        for node in tree.body:
            if isinstance(node, ast.Assign) and any(getattr(t, "id", "") == "ALIAS" for t in node.targets):
                try:
                    val = ast.literal_eval(node.value)
                except ValueError:
                    continue
                if isinstance(val, dict):
                    out[os.path.splitext(os.path.basename(p))[0]] = {k: list(v) for k, v in val.items() if isinstance(v, (list, tuple))}
    return out


def read_head(path):
    if path.lower().endswith(".csv"):
        import csv
        with open(path, encoding="utf-8-sig", newline="") as f:
            return [r for _, r in zip(range(10), csv.reader(f))]
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    return [["" if v is None else str(v) for v in r] for _, r in zip(range(10), ws.iter_rows(values_only=True))]


def match(head, alias):
    """필드마다 (판정, 열 이름, 받아들이는 이름)."""
    exact = {h.strip(): i for i, h in enumerate(head)}
    nh = {norm(h): h for h in head if h.strip()}
    res = {}
    for field, names in alias.items():
        hit = next((n for n in names if n in exact), None)
        if hit:
            res[field] = ("맞음", hit, hit)
            continue
        near = next(((nh[norm(n)], n) for n in names if norm(n) in nh), None)
        if near:
            res[field] = ("비슷함", near[0], near[1])
            continue
        part = next(((h, n) for n in names for h in head if h.strip() and len(norm(n)) >= 3 and norm(n) in norm(h)), None)
        res[field] = ("후보", part[0], part[1]) if part else ("없음", "", names[0])
    return res


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("listing")
    ap.add_argument("--script", help="스크립트 이름(일부)")
    ap.add_argument("--map", help="'기관열=킷열,…' 직접 짝짓기")
    ap.add_argument("--fix", action="store_true", help="머리글만 바꾼 사본을 만든다")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    if re.search(r"(^|/)(비밀|secret|classified)(/|$)", os.path.abspath(a.listing)):
        print("멈춤: 비밀 폴더의 파일은 다루지 않습니다.")
        return 2
    al = aliases()
    if a.script:
        al = {k: v for k, v in al.items() if a.script in k}
        if not al:
            print(f"입력 오류: '{a.script}'에 해당하는 스크립트가 없습니다.")
            return 2
    rows = read_head(a.listing)
    every = {n for v in al.values() for ns in v.values() for n in ns}
    hi = max(range(len(rows)), key=lambda i: sum(1 for c in rows[i] if norm(c) in {norm(n) for n in every}), default=0)
    head = rows[hi] if rows else []
    print(f"머리글 행: {hi + 1}행 · 열 {sum(1 for h in head if h.strip())}개 (자료 행은 읽지 않음)")
    rename = {}
    for name, alias in al.items():
        res = match(head, alias)
        ok = sum(1 for r in res.values() if r[0] == "맞음")
        print(f"\n[{name}] 알아봄 {ok}/{len(res)}")
        for field, (st, col, want) in res.items():
            if st == "맞음":
                continue
            if st in ("비슷함", "후보"):
                print(f"  {st}: '{col.strip()}' → '{want}' ({field})" + ("" if st == "비슷함" else " — 맞으면 --map으로 짝지어 주세요"))
                if st == "비슷함":
                    rename.setdefault(col, want)
            else:
                print(f"  없음: {field} — 받아들이는 이름 예: '{want}'")
    if a.map:
        for pair in a.map.split(","):
            if "=" in pair:
                k, v = (s.strip() for s in pair.split("=", 1))
                src = next((h for h in head if h.strip() == k), None)
                if src is None:
                    print(f"  --map: '{k}' 열이 없음")
                else:
                    rename[src] = v
    if not a.fix:
        if rename:
            print(f"\n머리글을 바꾸면 알아볼 열 {len(rename)}개 — --fix -o 사본.xlsx 로 사본을 만들 수 있습니다.")
        return 0
    if not a.out:
        print("입력 오류: --fix에는 -o 사본 경로가 필요합니다.")
        return 2
    if os.path.abspath(a.out) == os.path.abspath(a.listing) or os.path.exists(a.out):
        print("멈춤: 원본이나 이미 있는 파일에는 쓰지 않습니다.")
        return 2
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    if a.listing.lower().endswith(".csv"):
        import csv
        with open(a.listing, encoding="utf-8-sig", newline="") as f, open(a.out, "w", encoding="utf-8-sig", newline="") as g:
            w = csv.writer(g)
            for i, r in enumerate(csv.reader(f)):
                w.writerow([rename.get(c, c) for c in r] if i == hi else r)
    else:
        import openpyxl
        wb = openpyxl.load_workbook(a.listing)
        ws = wb.worksheets[0]
        for c in ws[hi + 1]:
            if c.value is not None and str(c.value) in rename:
                c.value = rename[str(c.value)]
        wb.save(a.out)
    print(f"\n머리글 {len(rename)}개를 바꾼 사본 → {a.out} (자료는 그대로)")
    for k, v in rename.items():
        print(f"  '{k.strip()}' → '{v}'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
