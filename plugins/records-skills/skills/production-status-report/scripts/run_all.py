#!/usr/bin/env python3
"""생산현황 통보 준비 한 번에: 입력 폴더의 목록 파일을 찾아 점검 스크립트를 차례로 돌리고,
국가기록원 통보서식 서식 1~9(비밀 10·11 제외)를 채우고, 통보 전에 고칠 것을 한 파일로 모은다.

사용:  python run_all.py 입력폴더 --year 2025 [--deadline 2026-05-31] [--list-scope long|all|none] [--fix-headers] -o 산출/통보준비_2025
입력폴더에 둘 파일(이름에 아래 낱말이 들어 있으면 찾음, .xlsx):
  등록정보   (필수) 서식 1·2, 생산현황 집계·점검          → production_status.py, fill_official_form.py
  시청각     서식 3·4, 시청각 등록 점검(시청각원본 폴더가 있으면 쪽수=파일 수) → av_check.py
  회의목록   서식 8                                       회의체   서식 9 (--bodies)
  회의록     폴더: 회의록 파일 구성요소 8가지·생산의무 회의 대조 → minutes_check.py
  행정박물   서식 5·6, 관리카드·이관 기한 점검             → artifacts_check.py
  간행물     서식 7, 발간등록번호·송부 점검                 → pubs_check.py
  통보서식   (선택) 국가기록원 원본 서식. 없으면 킷에 든 2024년 서식을 쓴다.
비밀:  비밀기록물은 이 명령에 넣지 않는다. 이름에 '비밀'이 든 파일·폴더는 건너뛰고 알린다.
       서식 10·11은 담당자가 비밀 전용 장비에서 classified-records-check와 fill_official_form.py --secret으로 따로 채운다.
결과:  산출 폴더에 각 점검 결과 xlsx, 생산현황_통보서식_<연도>.xlsx, 통보전_확인.md(단계별 실행·고칠 것 상위 항목).
원칙:  입력 파일은 읽기만 한다. 산출 폴더가 이미 있고 비어 있지 않으면 멈춘다. 제출은 사람이 한다.
"""
import argparse
import glob
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.path.dirname(os.path.dirname(HERE))
S = lambda skill, name: os.path.join(SK, skill, "scripts", name)
FORM = os.path.join(SK, "production-status-report", "assets", "2024년도기록물생산현황통보서식.xlsx")


def find(folder, word):
    hits = sorted(p for p in glob.glob(os.path.join(folder, f"*{word}*.xlsx")) if "비밀" not in os.path.basename(p) and not os.path.basename(p).startswith("~$"))
    return hits[0] if hits else None


def issues_of(path, sheet, col_name, ok=("통과",)):
    """결과 파일의 점검 시트에서 '통과'가 아닌 행 수와 앞 5개."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None, []
    ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return 0, []
    head = [str(c or "") for c in rows[0]]
    ci = next((i for i, h in enumerate(head) if h == col_name), len(head) - 1)
    bad = [r for r in rows[1:] if r and r[ci] not in (None, "", *ok)]
    return len(bad), [" · ".join(str(c) for c in r[:3] if c not in (None, ""))[:60] + " — " + str(r[ci])[:110] for r in bad[:5]]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--deadline")
    ap.add_argument("--list-scope", choices=["long", "all", "none"], default="long")
    ap.add_argument("--fix-headers", action="store_true", help="'비슷함' 머리글만 고친 사본을 산출 폴더에 만들어 그것으로 돌림(원본은 그대로)")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    if "비밀" in os.path.abspath(a.folder) or "비밀" in os.path.abspath(a.out):
        print("멈춤: 비밀기록물 폴더는 이 명령으로 다루지 않습니다(classified-records-check를 비밀 전용 장비에서).")
        return 2
    if os.path.isdir(a.out) and os.listdir(a.out):
        print(f"멈춤: 산출 폴더 '{a.out}'가 비어 있지 않음 — 덮어쓰지 않습니다.")
        return 2
    os.makedirs(a.out, exist_ok=True)
    f = {k: find(a.folder, k) for k in ("등록정보", "시청각", "회의목록", "회의체", "행정박물", "간행물", "통보서식")}
    minutes_dir = next((p for p in glob.glob(os.path.join(a.folder, "*회의록*")) if os.path.isdir(p) and "비밀" not in p), None)
    av_dir = next((p for p in glob.glob(os.path.join(a.folder, "*시청각원본*")) if os.path.isdir(p)), None)
    skipped_secret = [os.path.basename(p) for p in glob.glob(os.path.join(a.folder, "*비밀*"))]
    if not f["등록정보"]:
        print("멈춤: 입력 폴더에 '등록정보' 파일(.xlsx)이 없습니다.")
        return 2
    O = lambda name: os.path.join(a.out, name)
    steps, report = [], []
    heads = []  # 열 이름 점검(머리글만 읽음)
    for key, script in (("등록정보", "production_status"), ("등록정보", "fill_official_form"), ("시청각", "av_check"),
                        ("행정박물", "artifacts_check"), ("간행물", "pubs_check")):
        if f[key]:
            hp = subprocess.run([sys.executable, S("records-qa-gate", "header_check.py"), f[key], "--script", script], capture_output=True, text=True)
            near = [l.strip() for l in hp.stdout.splitlines() if l.strip().startswith(("비슷함:", "후보:"))]
            if near and a.fix_headers and any(l.startswith("비슷함:") for l in near):
                fixed = O(f"열맞춤_{os.path.basename(f[key])}")
                if not os.path.exists(fixed):
                    subprocess.run([sys.executable, S("records-qa-gate", "header_check.py"), f[key], "--fix", "-o", fixed], capture_output=True, text=True)
                if os.path.exists(fixed):
                    heads.append((os.path.basename(f[key]), script, [f"머리글을 고친 사본({os.path.basename(fixed)})으로 돌림"] + [l for l in near if l.startswith("후보:")]))
                    f[key] = fixed
                    continue
            if near:
                heads.append((os.path.basename(f[key]), script, near))

    def run(label, args, result=None, sheet=None, col=None, ok=("통과",)):
        p = subprocess.run([sys.executable, *args], capture_output=True, text=True)
        line = (p.stdout.strip().splitlines() or [""])[0]
        n, top = (issues_of(result, sheet, col, ok) if result and os.path.exists(result) and sheet else (None, []))
        steps.append((label, "실행" if p.returncode == 0 else f"실패({p.returncode})", line or p.stderr.strip()[-200:], n, top))
        return p

    run("생산현황 집계·점검", [S("production-status-report", "production_status.py"), f["등록정보"], "--year", str(a.year),
        *(["--deadline", a.deadline] if a.deadline else []), "-o", O(f"생산현황_{a.year}.xlsx")],
        O(f"생산현황_{a.year}.xlsx"), "고칠 것", "내용", ok=())
    if f["시청각"]:
        run("시청각 등록 점검", [S("typed-records-check", "av_check.py"), f["시청각"], *(["--files", av_dir] if av_dir else []), "-o", O("시청각_점검.xlsx")],
            O("시청각_점검.xlsx"), "점검", "판정")
    if minutes_dir:
        run("회의록 구성요소 점검", [S("typed-records-check", "minutes_check.py"), minutes_dir, *(["--meetings", f["회의목록"]] if f["회의목록"] else []),
            "-o", O("회의록_점검.xlsx")], O("회의록_점검.xlsx"), "구성 요소", "판정", ok=("구성 요소 갖춤",))
    if f["행정박물"]:
        run("행정박물 점검", [S("admin-artifacts-check", "artifacts_check.py"), f["행정박물"], "--asof", f"{a.year}-12-31", "-o", O("행정박물_점검.xlsx")],
            O("행정박물_점검.xlsx"), "점검", "판정")
    if f["간행물"]:
        run("간행물 점검", [S("publications-check", "pubs_check.py"), f["간행물"], "--year", str(a.year), "-o", O("간행물_점검.xlsx")],
            O("간행물_점검.xlsx"), "점검", "판정")
    form_out = O(f"생산현황_통보서식_{a.year}.xlsx")
    args = [S("production-status-report", "fill_official_form.py"), f["등록정보"], "--form", f["통보서식"] or FORM, "--year", str(a.year),
            "--list-scope", a.list_scope, "-o", form_out]
    if f["시청각"]:
        args += ["--av", f["시청각"]]
    if f["회의목록"]:
        args += ["--meetings", f["회의목록"]] + (["--bodies", f["회의체"]] if f["회의체"] else [])
    if os.path.exists(O("행정박물_점검.xlsx")):
        args += ["--artifacts", O("행정박물_점검.xlsx")]
    if os.path.exists(O("간행물_점검.xlsx")):
        args += ["--pubs", O("간행물_점검.xlsx")]
    p = run("통보서식 채우기", args)
    checks = [l.strip()[3:].strip() for l in p.stdout.splitlines() if l.strip().startswith("확인:")]

    filled = ["1", "2-1·2-2" if a.list_scope != "none" else ""] + (["3·4"] if f["시청각"] else []) + (["5·6"] if f["행정박물"] else []) \
        + (["7"] if f["간행물"] else []) + (["8"] if f["회의목록"] else []) + (["9"] if f["회의목록"] and f["회의체"] else [])
    missing = {"3·4": "시청각", "5·6": "행정박물", "7": "간행물", "8": "회의목록", "9": "회의체"}
    report.append(f"# {a.year}년 생산분 생산현황 통보 전 확인\n")
    report.append(f"- 입력 폴더: `{a.folder}`" + (f" · 통보 기한(기관 확인값): {a.deadline}" if a.deadline else " · 통보 기한: 기관에서 확인해 적을 것"))
    report.append(f"- 채운 서식: {', '.join(x for x in filled if x)}  → `{os.path.basename(form_out)}`")
    gaps = [f"서식 {k}('{v}' 파일 없음)" for k, v in missing.items() if k not in filled]
    report.append("- 비운 서식: " + (", ".join(gaps) if gaps else "없음") + " · 서식 10·11(비밀)은 이 명령 밖에서 담당자가 비밀 전용 장비로")
    if skipped_secret:
        report.append(f"- 건너뛴 비밀 관련 항목 {len(skipped_secret)}개(이름은 적지 않음)")
    if heads:
        report.append("\n## 열 이름 확인(머리글만 봄)\n")
        report.append("아래 열은 띄어쓰기·괄호 차이 등으로 스크립트가 못 알아볼 수 있습니다. "
                      "'비슷함'은 `--fix-headers`를 붙여 다시 돌리면 산출 폴더에 머리글만 고친 사본을 만들어 씁니다. "
                      "'후보'는 `header_check.py <파일> --map \"기관열=킷열\" --fix -o <사본>`으로 직접 짝지은 사본을 만들어 입력 폴더의 그 파일과 바꿔 넣으세요.\n")
        for fn, sc, near in heads:
            report.append(f"- `{fn}` ({sc}): " + " / ".join(near[:6]))
    report.append("\n## 단계별 결과\n")
    report.append("| 단계 | 실행 | 결과 | 고칠 것·확인 |\n|---|---|---|---|")
    for label, st, line, n, _ in steps:
        report.append(f"| {label} | {st} | {line.replace('|', '/')[:120]} | {'' if n is None else n} |")
    for label, st, line, n, top in steps:
        if top:
            report.append(f"\n### {label} — 앞 {len(top)}개\n")
            report += [f"- {t}" for t in top]
    if checks:
        report.append(f"\n### 통보서식 채우며 나온 확인 {len(checks)}개\n")
        report += [f"- {c}" for c in checks[:30]]
    report.append("\n## 제출 전에 사람이 할 일\n")
    report.append("1. 위의 고칠 것을 처리과에 돌려 등록정보를 고친 뒤 이 명령을 새 산출 폴더로 다시 돌린다(시행규칙 제21조: 정리 결과 반영 후 작성).")
    report.append("2. 통보서식을 엑셀에서 열어 총계·소계 수식 값을 확인한다. 서식 변경·시트 삭제는 하지 않는다.")
    report.append("3. 기관 유형에 따라 서식 2 제출 여부를 정한다(지자체·교육청·기타공공기관은 제외).")
    report.append("4. 시스템 연계 기관은 RMS→CAMS로, 아니면 공문으로 제출한다.")
    open(O("통보전_확인.md"), "w", encoding="utf-8").write("\n".join(report) + "\n")
    fails = [s for s in steps if not s[1].startswith("실행")]
    total = sum(s[3] or 0 for s in steps)
    print(f"{len(steps)}단계 · 실패 {len(fails)} · 고칠 것·확인 {total + len(checks)} · 서식 {', '.join(x for x in filled if x)} → {a.out}/통보전_확인.md")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
