#!/usr/bin/env python3
"""기록물 목록 작성·디지털화 사업 검수: 색인목록(철·건)의 면표시 연속성, 분권 철의 쪽 이어짐, 보존기간별 편철 순서,
필수 항목·자리수, 쪽수와 스캔 이미지 수 대조, 표본 추출(최소 3%·작업자별 일일), 품질 수준(99.9%) 판정을 한다.
근거: NAK 23:2017(v1.2) 「기록물 목록 작성 및 디지털화 작업 지침」 5.3.6·5.4·5.5·5.6.2·5.7 (국가기록원, 2026-11-20 내려받아 확인)

사용:  python listing_qa.py 색인목록.xlsx [--images 스캔폴더] [--sample 3] [--seed 1] [--checked 2000 --errors 1] -o 산출/목록검수.xlsx
입력 열(머리글 이름으로 찾음): 철관리번호(레이블번호), 철제목, 권호, 보존기간, 등록번호, 건제목, 생산일자, 시작쪽, 끝쪽(또는 쪽수), 작업자, 작업일
규칙:
  면표시(5.3.6): 철마다 1쪽부터 빈틈·겹침 없이, 분권 철은 앞 권 마지막 쪽 다음 번호로 시작(표지·색인목록 제외)
  편철 순서(표 2): 영구·준영구·30년은 오래된 문서가 위(앞 쪽), 10년 이하는 최근 문서가 위
  이미지(5.6.2): 건의 쪽수와 스캔 이미지 수가 다르면 재스캔 — --images 폴더 아래 '등록번호' 이름의 폴더 또는 '등록번호_'로 시작하는 파일을 셈
  표본(5.4.2·5.5.3·5.7.1): 전체에서 최소 3% 권장, 작업자·작업일이 있으면 작업자별 일일 표본
  품질 수준(5.7.2): 통상 99.9% — --checked(검사 건수)와 --errors(오류 건수)로 판정
원칙:  검수 보조 목록이다. 오류는 검증 담당자가 직접 고치지 않고 작업자가 고치게 한다(5.4.1). 기록물 원본은 열지 않는다.
"""
import argparse
import os
import random
import re
import sys
from collections import defaultdict

ALIAS = {"철": ["철관리번호", "레이블번호", "철분류번호", "기록물철관리번호"], "철제목": ["철제목", "기록물철제목", "철명"], "권": ["권호", "권호수", "권"],
         "보존": ["보존기간"], "건": ["등록번호", "기록물건등록번호", "건번호"], "건제목": ["건제목", "제목", "기록물건제목"], "일자": ["생산일자", "시행일자", "접수일자"],
         "시작": ["시작쪽", "시작 쪽", "시작면"], "끝": ["끝쪽", "끝 쪽", "끝면"], "쪽수": ["쪽수", "면수"], "작업자": ["작업자", "입력자", "등록 담당자"],
         "작업일": ["작업일", "입력일", "등록일"]}
NEED = ["철", "철제목", "보존", "건", "건제목", "일자"]
LONG = ("영구", "준영구", "30년")


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    names = {n for v in ALIAS.values() for n in v}
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c in names for c in r) >= 2), 0)
    head = rows[hi]
    idx = {k: head.index(n) for k, al in ALIAS.items() for n in reversed(al) if n in head}
    return idx, [(n, {k: (r[i] if i < len(r) else "") for k, i in idx.items()}) for n, r in enumerate(rows[hi + 1:], hi + 2) if any(r)]


def num(v):
    m = re.search(r"\d+", v or "")
    return int(m[0]) if m else None


def day(v):
    m = re.search(r"(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})", v or "")
    return (int(m[1]), int(m[2]), int(m[3])) if m else None


def count_images(root):
    cnt = defaultdict(int)
    exts = {".tif", ".tiff", ".jpg", ".jpeg", ".png", ".jp2", ".pdf"}
    for d, _, fs in os.walk(root):
        base = os.path.basename(d)
        for f in fs:
            if os.path.splitext(f)[1].lower() not in exts:
                continue
            if d != root:
                cnt[base] += 1
            m = re.match(r"(.+?)_\d+", os.path.splitext(f)[0])
            if m and d == root:
                cnt[m[1]] += 1
    return cnt


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("table")
    ap.add_argument("--images", help="스캔 이미지 폴더(등록번호별 폴더 또는 '등록번호_쪽' 파일)")
    ap.add_argument("--sample", type=float, default=3.0, help="표본 비율 %%(최소 3%% 권장)")
    ap.add_argument("--seed", type=int, default=1, help="표본 추출 난수 시드(재현용)")
    ap.add_argument("--maxlen", type=int, default=0, help="제목 자리수 상한(시스템 입력 한도) — 0이면 점검 안 함")
    ap.add_argument("--checked", type=int, help="검사한 건수")
    ap.add_argument("--errors", type=int, help="검사에서 나온 오류 건수")
    ap.add_argument("-o", "--out", default="산출/목록검수.xlsx")
    a = ap.parse_args()
    idx, rows = read(a.table)
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "검수 결과"
    ws.append(["행", "철관리번호", "권호", "등록번호", "구분", "내용", "고칠 사람"])
    issues = []
    for n, x in rows:
        miss = [k for k in NEED if k in idx and not x.get(k)]
        miss += [k for k in NEED if k not in idx and k in ("철", "건")]
        if miss:
            issues.append((n, x, "필수 항목", "빈 칸: " + ", ".join(miss)))
        if a.maxlen:
            for k in ("철제목", "건제목"):
                if len(x.get(k, "")) > a.maxlen:
                    issues.append((n, x, "자리수", f"{k} {len(x[k])}자 > {a.maxlen}자 — 예외처리 요청(5.8)"))
    # 철·권별 묶기
    vols = defaultdict(list)
    for n, x in rows:
        vols[(x.get("철", ""), num(x.get("권")) or 1)].append((n, x))
    by_file = defaultdict(list)
    for (f, v), items in vols.items():
        by_file[f].append((v, items))
    for f, vlist in by_file.items():
        vlist.sort(key=lambda z: z[0])
        expect = 1
        for v, items in vlist:
            spans = []
            for n, x in items:
                s, e = num(x.get("시작")), num(x.get("끝"))
                if s is not None and e is None and num(x.get("쪽수")):
                    e = s + num(x.get("쪽수")) - 1
                if s is None:
                    continue
                spans.append((s, e if e is not None else s, n, x))
            spans.sort(key=lambda z: z[0])
            if spans and spans[0][0] != expect:
                issues.append((spans[0][2], spans[0][3], "면표시", f"{v}권 첫 쪽 {spans[0][0]} — {expect}쪽부터 시작해야 함"
                               + ("(분권: 앞 권 마지막 쪽 다음 번호)" if v > 1 else "")))
            for (s1, e1, n1, x1), (s2, e2, n2, x2) in zip(spans, spans[1:]):
                if s2 > e1 + 1:
                    issues.append((n2, x2, "면표시", f"{e1 + 1}~{s2 - 1}쪽 빠짐(건 유실 또는 번호 누락)"))
                elif s2 <= e1:
                    issues.append((n2, x2, "면표시", f"{s2}쪽이 앞 건({e1}쪽까지)과 겹침"))
            for s, e, n, x in spans:
                if e < s:
                    issues.append((n, x, "면표시", f"끝쪽 {e} < 시작쪽 {s}"))
            if spans:
                expect = max(e for _, e, _, _ in spans) + 1
            # 편철 순서
            ret = re.sub(r"\s", "", items[0][1].get("보존", ""))
            dated = [(s, day(x.get("일자")), n, x) for s, e, n, x in spans if day(x.get("일자"))]
            if len(dated) >= 2 and ret:
                asc = ret in LONG
                bad = [(d1, d2, n2, x2) for (_, d1, _, _), (_, d2, n2, x2) in zip(dated, dated[1:]) if (d2 < d1 if asc else d2 > d1)]
                if bad:
                    d1, d2, n2, x2 = bad[0]
                    issues.append((n2, x2, "편철 순서", ("영구·준영구·30년은 오래된 문서가 위(앞)" if asc else "10년 이하는 최근 문서가 위(앞)")
                                   + f" — {'-'.join(map(str, d1))} 다음에 {'-'.join(map(str, d2))}" + (f" 외 {len(bad) - 1}곳" if len(bad) > 1 else "")))
    imgs = count_images(a.images) if a.images else None
    nimg = 0
    if imgs is not None:
        for n, x in rows:
            s, e = num(x.get("시작")), num(x.get("끝"))
            pages = num(x.get("쪽수")) or ((e - s + 1) if s is not None and e is not None else None)
            got = imgs.get(x.get("건", ""), 0)
            if pages is not None and got != pages:
                nimg += 1
                issues.append((n, x, "이미지 수", f"쪽수 {pages} · 이미지 {got} — 재스캔 또는 이미지 연결 확인(5.6.2)"))
    for n, x, kind, msg in issues:
        ws.append([n, x.get("철", ""), x.get("권", ""), x.get("건", ""), kind, msg,
                   "색인목록 등록 담당자" if kind in ("필수 항목", "자리수") else "디지털화 담당자" if kind == "이미지 수" else "기록물정리 작업자"])
        ws.cell(ws.max_row, 5).fill = red
    # 표본
    sp = wb.create_sheet("표본")
    sp.append(["작업자", "작업일", "행", "철관리번호", "등록번호", "건제목", "검사 결과(정상/오류)", "오류 내용"])
    rnd = random.Random(a.seed)
    groups = defaultdict(list)
    for n, x in rows:
        groups[(x.get("작업자", ""), x.get("작업일", ""))].append((n, x))
    picked = 0
    for (w, d), items in sorted(groups.items()):
        k = max(1, -(-len(items) * a.sample // 100))
        for n, x in sorted(rnd.sample(items, min(int(k), len(items)))):
            sp.append([w, d, n, x.get("철", ""), x.get("건", ""), x.get("건제목", ""), "", ""])
            picked += 1
    sp.append([])
    sp.append([f"표본 {picked}건 / 전체 {len(rows)}건 ({picked / max(len(rows), 1):.1%}) — 최소 3% 권장(5.7.1), 작업자·작업일별로 뽑음(5.4.2·5.5.3), 시드 {a.seed}"])
    q = wb.create_sheet("품질 수준")
    q.append(["항목", "값", "판정"])
    if a.checked:
        e = a.errors or 0
        acc = 1 - e / a.checked
        q.append(["검사 건수", a.checked, ""])
        q.append(["오류 건수", e, ""])
        q.append(["정확도", f"{acc:.3%}", "충족" if acc >= 0.999 else "미달 — 재작업 지시·시정조치 확인 후 교정확인서(5.7.2)"])
    else:
        q.append(["정확도", "--checked·--errors를 주면 계산", "품질 수준 통상 99.9%"])
    q.append(["이 목록의 자동 점검 오류", len(issues), "검사 전에 먼저 고칠 것"])
    fm = wb.create_sheet("서식 목록(부속서)")
    fm.append(["서식", "쓰임"])
    for line in [("A.1 기록물 반출·반입서", "작업장으로 반출·반입 — records-forms carry 서식으로 만들 수 있음"), ("A.2 기록물 인수·인계서", "공정 사이·기관 인계"),
                 ("A.3 예외처리 요청서 / A.4 협의서", "반출 불가, 중첩 부착물, 판독 어려운 수기 목록, 자리수 초과, 규격 외 스캔(5.8)"),
                 ("A.5 교정확인서", "결함 발견 → 재작업·시정 확인"), ("A.6 검사확인서", "표본·전수 검사 결과"),
                 ("보안 실태점검(표 1)", "인적·장비·네트워크·기록물 4개 유형, 월 1회 이상(4.2)")]:
        fm.append(list(line))
    for sh in wb.worksheets:
        sh["A1"].font = Font(bold=True)
    for sh, widths in ((ws, (6, 16, 6, 18, 10, 70, 18)), (sp, (12, 12, 6, 16, 18, 40, 16, 30)), (q, (24, 30, 60)), (fm, (30, 80))):
        for i, w in enumerate(widths):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    kinds = defaultdict(int)
    for _, _, k, _ in issues:
        kinds[k] += 1
    print(f"{len(rows)}건 · 오류 {len(issues)}(" + ", ".join(f"{k} {v}" for k, v in kinds.items()) + f") · 표본 {picked} → {a.out}")
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
