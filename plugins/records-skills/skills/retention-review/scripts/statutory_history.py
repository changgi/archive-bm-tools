#!/usr/bin/env python3
"""법정 보존기간 연혁: 점검 표(statutory_minimum.csv)의 법령마다 국가법령정보센터 '연혁' 목록의 모든 판본 원문을 받아,
그 기록 종류(낱말)의 보존기간이 언제부터 언제까지 몇 년이었는지 표로 만든다. 옛 기록(예: 2012년 생산)을 평가할 때
'생산 당시 법정 기간'을 확인하는 데 쓴다.

사용:  python statutory_history.py --build [-o ../assets/law/statutory_history.csv] [--law 근로기준법] [--workers 3]
       python statutory_history.py --find 휴가 [--year 2012]      # 표에서 찾기(인터넷 없음)
원칙:  외부 API·비용 없음. 법령(법률·시행령·시행규칙)만 본다 — 행정규칙(고시 등)은 연혁 화면이 달라 아직 넣지 않음.
       조문 번호는 판본마다 바뀔 수 있어 번호가 아니라 점검 표의 '낱말'로 같은 기록을 찾는다. 원문은 3개씩 동시에, 요청마다 0.7초 쉼.
"""
import argparse
import csv
import importlib.util
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
MINIMUM = os.path.join(HERE, "..", "assets", "law", "statutory_minimum.csv")
OUT = os.path.join(HERE, "..", "assets", "law", "statutory_history.csv")
BASE = "https://www.law.go.kr/LSW/"
COLS = ["기록 종류", "법령", "판본 시행일", "다음 판본 시행일", "조문", "보존기간", "성격", "공포", "비고", "문장"]


def _mod(name):
    sp = importlib.util.spec_from_file_location(name, os.path.join(HERE, name + ".py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


def get(url):
    err = None
    for i in range(5):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=90) as r:
                return r.read().decode("utf-8", "ignore")
        except Exception as e:  # noqa: BLE001
            err = e
            time.sleep(4 + 3 * i)
    raise err


def law_names(row):
    """점검 표 한 행이 가리키는 법령 이름들(법률 + '시행령 제N조'면 시행령 …)."""
    base = re.sub(r"\s*(시행령|시행규칙)$", "", row["법령"]).strip()
    names = {row["법령"]}
    for kind in ("시행령", "시행규칙"):
        if kind in row["조항"]:
            names.add(f"{base} {kind}")
    return sorted(names)


def find_current(name):
    """현행 법령의 lsiSeq·lsId."""
    h = get(f"{BASE}lsSc.do?menuId=1&subMenuId=15&tabMenuId=81&query={urllib.parse.quote(name)}")
    for seq in dict.fromkeys(re.findall(r"(?:lawInfoPop|lsViewWideAll)\('(\d+)'", h)):
        p = get(f"{BASE}lsInfoP.do?lsiSeq={seq}")
        lid = re.search(r'id="lsId"[^>]*value="([^"]*)"', p) or re.search(r'value="([^"]*)"[^>]*id="lsId"', p)
        nm = re.search(r'data-id="history">\s*([^\[<]+?)\s*\[시행', p)
        if lid and nm and re.sub(r"\s", "", nm.group(1)) == re.sub(r"\s", "", name):
            return seq, lid.group(1)
    return None, None


def versions(seq, lsid):
    h = get(f"{BASE}lsHstListR.do?lsiSeq={seq}&lsId={lsid}&chrClsCd=010202&ancYnChk=0")
    ko = h.split('data-id="ENG"')[0] if 'data-id="ENG">' in h else h
    out = []
    for m in re.finditer(r"lsViewLsHst2\('(\d+)',\s*'(\d{8})',\s*'([^']*)',\s*'(\d{8})'", ko):
        out.append({"seq": m.group(1), "공포": m.group(2), "시행": m.group(4)})
    seen, uniq = set(), []
    for v in out:
        if v["seq"] not in seen:
            seen.add(v["seq"])
            uniq.append(v)
    return sorted(uniq, key=lambda v: v["시행"])


CACHE = None  # --cache: 받은 판본 원문을 두는 폴더(다시 만들 때 다시 받지 않음)


def body(v):
    fn = os.path.join(CACHE, f"{v['seq']}_{v['시행']}.txt") if CACHE else None
    if fn and os.path.exists(fn):
        with open(fn, encoding="utf-8") as f:
            return f.read()
    t = get(f"{BASE}lsInfoR.do?lsiSeq={v['seq']}&chrClsCd=010202&urlMode=lsInfoP&efYd={v['시행']}&ancYnChk=0")
    time.sleep(0.7)
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
    if fn:
        with open(fn, "w", encoding="utf-8") as f:
            f.write(t)
    return t


def ymd(s):
    return f"{s[:4]}-{s[4:6]}-{s[6:]}" if s and len(s) == 8 else ""


def build(a):
    b = _mod("statutory_catalog_build")
    with open(MINIMUM, encoding="utf-8-sig") as f:
        mins = [r for r in csv.DictReader(f) if not r["법령"].endswith(")")]  # 고시 등 행정규칙은 뺌
    if a.law:
        mins = [r for r in mins if any(x in r["법령"] for x in a.law.split(","))]
    fams = {}
    for r in mins:
        for n in law_names(r):
            fams.setdefault(n, []).append(r)
    rows, log, got = [], [], {}
    for name in fams:
        seq, lsid = find_current(name)
        if not seq:
            log.append(f"못 찾음: {name}")
            continue
        vs = versions(seq, lsid)
        with ThreadPoolExecutor(max_workers=max(1, min(a.workers, 4))) as ex:
            got[name] = list(ex.map(lambda v: (v, body(v)), vs))
        print(f"  {name}: 판본 {len(vs)}개 받음", flush=True)
        log.append(f"{name}: 판본 {len(vs)}개")

    def sub_has(law, date, jo, pat):
        """법률 조문이 '대통령령·부령으로 정하는 서류'를 보존하라고 할 때, 그 날 시행 중인 시행령·시행규칙에서
        '법 {jo}'를 가리키는 조문이 그 기록(낱말)을 대상으로 드는지."""
        for kind in ("시행령", "시행규칙"):
            sub = [(v, t) for v, t in got.get(f"{law} {kind}", []) if v["시행"] <= date]
            if not sub:
                continue
            t = sub[-1][1]
            ms = list(b.ART.finditer(t))
            jp = jo.replace("제", "제\\s*")
            for i, m in enumerate(ms):
                art = t[m.start(): ms[i + 1].start() if i + 1 < len(ms) else len(t)]
                if re.search(r"법\s*" + jp + r"(?!의)", art) and re.search(pat, art):
                    return True
        return False

    for name, recs in fams.items():
        vs = [v for v, _ in got.get(name, [])]
        for i, (v, t) in enumerate(got.get(name, [])):
            nxt = vs[i + 1]["시행"] if i + 1 < len(vs) else ""
            ext = b.extract(f"{name} [시행 {v['시행'][:4]}. {int(v['시행'][4:6])}. {int(v['시행'][6:])}.]", t)
            for rec in recs:
                for e in ext:
                    rec_txt = (e.get("기록(문구)") or "").strip()
                    # 한 조문에 여러 기록·기간을 늘어놓은 경우(의료법 시행규칙 제15조 등) 기간 바로 앞 문구로만 맞춤
                    hay = rec_txt if len(rec_txt) >= 2 else e["문장"]
                    direct = re.search(rec["낱말"], hay)
                    deleg = (not direct and re.search(r"(대통령령|부령|총리령)(으로|이)\s*정하는", e["문장"])
                             and sub_has(name, v["시행"], e["조문"].split("(")[0], rec["낱말"]))
                    if (direct or deleg) and not (rec.get("제외 낱말") and re.search(rec["제외 낱말"], hay)):
                        rows.append({"기록 종류": rec["기록 종류"], "법령": name, "판본 시행일": ymd(v["시행"]), "다음 판본 시행일": ymd(nxt),
                                     "조문": e["조문"], "보존기간": e["보존기간"], "성격": e.get("성격", ""), "공포": ymd(v["공포"]), "문장": e["문장"][:300]})
    # 같은 기록 종류·법령·조문·기간이 이어지는 판본은 한 줄로 합침(언제부터 언제까지)
    rows.sort(key=lambda r: (r["기록 종류"], r["법령"], r["판본 시행일"]))
    merged = []
    for r in rows:
        last = next((m for m in reversed(merged) if m["기록 종류"] == r["기록 종류"] and m["법령"] == r["법령"]
                     and m["조문"].split("(")[0] == r["조문"].split("(")[0] and m["보존기간"] == r["보존기간"]), None)
        if last and last["다음 판본 시행일"] >= r["판본 시행일"] or (last and last["_end"] == r["판본 시행일"]):
            last["다음 판본 시행일"] = r["다음 판본 시행일"]
            last["_end"] = r["다음 판본 시행일"]
            continue
        r["_end"] = r["다음 판본 시행일"]
        merged.append(r)
    multi = {}
    for m in merged:
        k = (m["기록 종류"], m["법령"], m["조문"].split("(")[0])
        multi.setdefault(k, set()).add(m["보존기간"])
    today = time.strftime("%Y-%m-%d")
    for m in merged:
        notes = []
        if len(multi[(m["기록 종류"], m["법령"], m["조문"].split("(")[0])]) > 1:
            notes.append("같은 조문에 기간이 여럿 — 호·단서별로 원문 확인")
        if m["판본 시행일"] > today:
            notes.append("시행 예정")
        m["비고"] = "; ".join(notes)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(merged)
    print("\n".join(log))
    print(f"연혁 {len(merged)}줄(판본별 {len(rows)}건) → {a.out}")
    return 0


def find(a):
    with open(a.out, encoding="utf-8-sig") as f:
        R = list(csv.DictReader(f))
    hit = [r for r in R if a.find in r["기록 종류"] or a.find in r["문장"] or a.find in r["법령"]]
    if a.year:
        y = f"{a.year}-12-31"
        hit = [r for r in hit if r["판본 시행일"] <= y and (not r["다음 판본 시행일"] or r["다음 판본 시행일"] > f"{a.year}-01-01")]
    for r in hit:
        print(f"{r['기록 종류']}\t{r['법령']} {r['조문']}\t{r['판본 시행일']} ~ {r['다음 판본 시행일'] or '현행'}\t{r['보존기간']}" + (f"\t({r['비고']})" if r.get("비고") else ""))
    print(f"{len(hit)}줄")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--build", action="store_true")
    ap.add_argument("--find")
    ap.add_argument("--year", type=int)
    ap.add_argument("--law", help="쉼표로 나눈 법령 이름 일부(--build 범위)")
    ap.add_argument("--workers", type=int, default=3)
    ap.add_argument("--cache", default="산출/법령연혁원문", help="받은 판본 원문을 두는 폴더(다시 만들 때 재사용)")
    ap.add_argument("-o", "--out", default=OUT)
    a = ap.parse_args()
    if a.build:
        global CACHE
        CACHE = a.cache
        os.makedirs(CACHE, exist_ok=True)
        return build(a)
    if a.find:
        return find(a)
    ap.error("--build 또는 --find")


if __name__ == "__main__":
    sys.exit(main())
