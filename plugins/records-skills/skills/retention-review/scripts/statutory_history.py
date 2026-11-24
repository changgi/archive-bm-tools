#!/usr/bin/env python3
"""법정 보존기간 연혁: 점검 표(statutory_minimum.csv)의 법령마다 국가법령정보센터 '연혁' 목록의 모든 판본 원문을 받아,
그 기록 종류(낱말)의 보존기간이 언제부터 언제까지 몇 년이었는지 표로 만든다. 옛 기록(예: 2012년 생산)을 평가할 때
'생산 당시 법정 기간'을 확인하는 데 쓴다.

사용:  python statutory_history.py --build [-o ../assets/law/statutory_history.csv] [--law 근로기준법] [--workers 3]
       python statutory_history.py --find 휴가 [--year 2012]      # 표에서 찾기(인터넷 없음)
원칙:  외부 API·비용 없음. 법령(법률·시행령·시행규칙)과 행정규칙(고시 등, admRulHstListR 연혁 목록)을 본다.
       판본마다 보존 조문 위임 추적기(retention_tree.py)로 항·호·단서를 나누고 그날 시행 중이던 시행령·시행규칙·(현행) 고시까지 잇는다.
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
ITEMS = os.path.join(HERE, "..", "assets", "law", "statutory_items.csv")  # 위임 추적 목록 — 사슬에 걸린 고시 이름
OUT = os.path.join(HERE, "..", "assets", "law", "statutory_history.csv")
BASE = "https://www.law.go.kr/LSW/"
COLS = ["기록 종류", "법령", "판본 시행일", "다음 판본 시행일", "조문", "보존기간", "조건", "성격", "공포", "비고", "근거 경로", "문장"]


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


def find_admin(name):
    """현행 행정규칙(고시·훈령 등)의 admRulSeq."""
    h = get(f"{BASE}admRulSc.do?menuId=5&subMenuId=41&tabMenuId=183&query={urllib.parse.quote(name)}")
    for seq in dict.fromkeys(re.findall(r"lawInfPop\('(\d+)'", h)):
        vs = admin_versions(seq)
        if vs and re.sub(r"\s|\(.*?\)", "", vs[-1]["이름"]) == re.sub(r"\s|\(.*?\)", "", name):
            return seq
    return None


def admin_versions(seq):
    h = get(f"{BASE}admRulHstListR.do?admRulSeq={seq}")
    out = []
    for m in re.finditer(r"admRulViewHst\('Y','(\d+)'\)[^>]*>\s*\d+\.\s*([^<]+)<br\s*/?>\s*<div[^>]*>\s*\[시행\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\]\s*\[[^\],]*,\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.", h):
        out.append({"seq": m.group(1), "이름": m.group(2).strip(), "시행": f"{m.group(3)}{int(m.group(4)):02d}{int(m.group(5)):02d}",
                    "공포": f"{m.group(6)}{int(m.group(7)):02d}{int(m.group(8)):02d}", "행정규칙": True})
    return sorted(out, key=lambda v: v["시행"])


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
    t = get(f"{BASE}admRulLsInfoR.do?admRulSeq={v['seq']}" if v.get("행정규칙") else
            f"{BASE}lsInfoR.do?lsiSeq={v['seq']}&chrClsCd=010202&urlMode=lsInfoP&efYd={v['시행']}&ancYnChk=0")
    time.sleep(0.7)
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
    if fn:
        with open(fn, "w", encoding="utf-8") as f:
            f.write(t)
    return t


def _jk(jo):
    """합칠 때 쓰는 조문 열쇠: 제목 빼고, 호 없이 '①'만 붙은 것은 항 없는 것과 같게(판본마다 항이 생기고 없어짐)."""
    k = re.sub(r"\(.*?\)", "", jo)
    return re.sub(r"①$", "", k)


def _ck(c):
    """조건 비교: 판본마다 바뀐 띄어쓰기·가운뎃점(·ㆍ)·꼬리 글자는 같은 것으로."""
    c = re.sub(r"[가-힣]{2,10}(?:부장관|처장|청장|위원장)", "", c)  # 정부조직 개편으로 바뀐 부처 이름('노동부장관'→'고용노동부장관')
    return re.sub(r"[\s·ㆍ,.]", "", c)[:80]


def ymd(s):
    return f"{s[:4]}-{s[4:6]}-{s[6:]}" if s and len(s) == 8 else ""


def build(a):
    b = _mod("statutory_catalog_build")
    with open(MINIMUM, encoding="utf-8-sig") as f:
        mins = list(csv.DictReader(f))
    if a.law:
        mins = [r for r in mins if any(x in r["법령"] for x in a.law.split(","))]
    fams = {}
    for r in mins:
        for n in law_names(r):
            fams.setdefault(n, []).append(r)
    rows, log, got = [], [], {}
    for name in fams:
        if name.endswith(")"):  # '개인정보의 안전성 확보조치 기준(개인정보보호위원회고시)' — 행정규칙 연혁
            seq = find_admin(re.sub(r"\([^)]*\)$", "", name).strip())
            vs = admin_versions(seq) if seq else []
        else:
            seq, lsid = find_current(name)
            vs = versions(seq, lsid) if seq else []
        if not vs:
            log.append(f"못 찾음: {name}")
            continue
        with ThreadPoolExecutor(max_workers=max(1, min(a.workers, 4))) as ex:
            got[name] = list(ex.map(lambda v: (v, body(v)), vs))
        print(f"  {name}: 판본 {len(vs)}개 받음", flush=True)
        log.append(f"{name}: 판본 {len(vs)}개")

    # 위임 사슬에 걸린 고시(위임 추적 목록에서 그 법령부터 이어진 고시) — 고시도 판본마다 받아 그날 시행 중인 판본으로 잇는다
    gosi_of, ggot = {}, {}
    if os.path.exists(ITEMS) and not a.no_gosi_history:
        with open(ITEMS, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if r.get("위임") == "고시" and r.get("상태") == "위임 추적":
                    root = re.sub(r"\s*(?:시행령|시행규칙)$", "", r["근거 경로"].split(" → ")[0].rsplit(" ", 1)[0])
                    gosi_of.setdefault(root, set()).add(r["법령"])
    for name in fams:
        for g in sorted(gosi_of.get(re.sub(r"\s*(?:시행령|시행규칙)$", "", name), ())):
            if g in ggot:
                continue
            gseq = find_admin(g)
            gvs = admin_versions(gseq) if gseq else []
            if not gvs:
                log.append(f"고시 못 찾음: {g}")
                ggot[g] = []
                continue
            with ThreadPoolExecutor(max_workers=max(1, min(a.workers, 4))) as ex:
                ggot[g] = list(ex.map(lambda v: (v, body(v)), gvs))
            print(f"  {g}(고시): 판본 {len(gvs)}개 받음", flush=True)
            log.append(f"{g}(고시): 판본 {len(gvs)}개 — {', '.join(sorted({v['이름'] for v in gvs}))}")
    rt = _mod("retention_tree")
    admins = {}
    if a.gosi_cache and os.path.isdir(a.gosi_cache):  # 위임 추적기가 받아 둔 현행 고시(고시 연혁은 아직 따라가지 않음)
        for fn in os.listdir(a.gosi_cache):
            if fn.startswith("행정규칙_"):
                t = open(os.path.join(a.gosi_cache, fn), encoding="utf-8").read()
                nm = re.search(r"([^\[\]{};]{2,80}?)\s*\[시행[^\]]*\]\s*\[([^\]\s]+)", t)
                if nm and nm.group(1).strip() not in ggot:
                    admins[nm.group(1).strip()] = f"[{nm.group(2)}] " + t
    for name, recs in fams.items():
        seqs = list(got.get(name, []))
        mine = [g for g in sorted(gosi_of.get(re.sub(r"\s*(?:시행령|시행규칙)$", "", name), ())) if ggot.get(g)]
        if seqs and mine:  # 고시가 바뀐 날도 판본 경계로(그날 시행 중인 법령 판본과 함께)
            base_seqs = list(seqs)  # 법령 판본만(덧붙인 고시 날짜 항목은 고르지 않음)
            have = {v["시행"] for v, _ in seqs}
            for g in mine:
                for gv, _ in ggot[g]:
                    d = gv["시행"]
                    if d in have or d < seqs[0][0]["시행"]:
                        continue
                    lv, lt = [x for x in base_seqs if x[0]["시행"] <= d][-1]
                    seqs.append((dict(lv, 시행=d, 공포=gv["공포"], _gosi=g), lt))
                    have.add(d)
            seqs.sort(key=lambda x: x[0]["시행"])
        vs = [v for v, _ in seqs]
        for i, (v, t) in enumerate(seqs):
            nxt = vs[i + 1]["시행"] if i + 1 < len(vs) else ""
            date = v["시행"]
            laws_d = {}
            for n, lst in got.items():  # 그 날 시행 중이던 법률·시행령·시행규칙 판본끼리 위임을 잇는다
                sub = [tt for vv, tt in lst if vv["시행"] <= date]
                if sub:
                    laws_d[n] = sub[-1]
            laws_d[name] = t
            adm_d = dict(admins)
            for g in mine:
                cur = [(gv, gt) for gv, gt in ggot[g] if gv["시행"] <= date]
                if cur:
                    gt = cur[-1][1]
                    kd = re.search(r"\[시행[^\]]*\]\s*\[([^\]\s]+)", gt)  # '[고용노동부고시 제2021-…호' — 고시 찾기가 기관을 앞머리에서 봄
                    adm_d[cur[-1][0]["이름"]] = (f"[{kd.group(1)}] " if kd else "") + gt
            corpus = rt.Corpus(laws_d, adm_d, live=False, live_laws=False)
            items = []
            for jo, title, abody in corpus.arts(name):
                if rt.KEEPV.search(abody) or rt.EVENT.search(abody):
                    items += [x for x in rt.analyze(corpus, name, jo, title, abody) if x["법령"] == name or x.get("_via")]
            for rec in recs:
                pat = rec["낱말"]
                # 그 조문에 이 기록이 호·항으로 따로 적혀 있으면 그 문구로만, 아니면 조문 제목도 봄
                own = {x["조문"] for x in items if re.search(pat, x["기록(항목)"] + " " + x.get("참조 조문", ""))}
                for x in items:
                    hay = x["기록(항목)"] + " " + x.get("참조 조문", "") + ("" if x["조문"] in own else " " + x["조문"])
                    if not x["보존기간"] or not re.search(pat, hay):
                        continue
                    if rec.get("제외 낱말") and re.search(rec["제외 낱말"], hay):
                        continue
                    jo_lbl = x["조문"] + (x["항"] or "") + (x["호"] + "호" if x["호"] and x["호"].isdigit() else (x["호"] + "목" if x["호"] else ""))
                    rows.append({"기록 종류": rec["기록 종류"], "법령": x["법령"], "판본 시행일": ymd(v["시행"]), "다음 판본 시행일": ymd(nxt),
                                 "조문": jo_lbl, "보존기간": x["보존기간"], "성격": "", "공포": ymd(v["공포"]),
                                 "근거 경로": x["근거 경로"], "조건": x.get("조건", ""), "문장": (x["기록(항목)"] + " ⟵ " + x["문장"])[:300]})
    # 같은 기록 종류·법령·조문·기간이 이어지는 판본은 한 줄로 합침(언제부터 언제까지)
    rows.sort(key=lambda r: (r["기록 종류"], r["법령"], r["판본 시행일"]))
    merged = []
    for r in rows:
        last = next((m for m in reversed(merged) if m["기록 종류"] == r["기록 종류"] and m["법령"] == r["법령"]
                     and _jk(m["조문"]) == _jk(r["조문"]) and m["보존기간"] == r["보존기간"] and _ck(m.get("조건", "")) == _ck(r.get("조건", ""))), None)
        if last and last["_end"] == "":  # 이미 '현행'까지 이어진 줄
            continue
        if last and last["다음 판본 시행일"] >= r["판본 시행일"] or (last and last["_end"] == r["판본 시행일"]):
            last["다음 판본 시행일"] = r["다음 판본 시행일"]
            last["_end"] = r["다음 판본 시행일"]
            continue
        r["_end"] = r["다음 판본 시행일"]
        merged.append(r)
    multi = {}
    for m in merged:  # 같은 때(기간이 겹치는 판본)에 같은 조문·항·호가 서로 다른 기간을 가질 때만
        k = (m["기록 종류"], m["법령"], _jk(m["조문"]))
        for o in multi.get(k, []):
            if o["보존기간"] != m["보존기간"] and not (o.get("조건") or m.get("조건")) and o["판본 시행일"] < (m["다음 판본 시행일"] or "9999") and m["판본 시행일"] < (o["다음 판본 시행일"] or "9999"):
                o["_multi"] = m["_multi"] = True
        multi.setdefault(k, []).append(m)
    today = time.strftime("%Y-%m-%d")
    for m in merged:
        notes = []
        if m.get("_multi"):
            notes.append("같은 때 같은 조문·항·호에 기간이 여럿 — 단서·조건 확인")
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
    hit = [r for r in R if a.find in r["기록 종류"]] or [r for r in R if a.find in r["문장"] or a.find in r["법령"]]
    if a.year:
        y = f"{a.year}-12-31"
        hit = [r for r in hit if r["판본 시행일"] <= y and (not r["다음 판본 시행일"] or r["다음 판본 시행일"] > f"{a.year}-01-01")]
    for r in hit:
        print(f"{r['기록 종류']}\t{r['법령']} {r['조문']}\t{r['판본 시행일']} ~ {r['다음 판본 시행일'] or '현행'}\t{r['보존기간']}"
              + (f"\t[{r['조건'][:60]}]" if r.get("조건") else "") + (f"\t({r['비고']})" if r.get("비고") else ""))
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
    ap.add_argument("--no-gosi-history", action="store_true", help="고시 판본은 받지 않고 현행 고시만(예전 방식)")
    ap.add_argument("--gosi-cache", default="산출/위임추적원문", help="위임 추적기(retention_tree.py)가 받아 둔 고시 원문 폴더")
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
