#!/usr/bin/env python3
"""법정 보존기간 목록 갱신 감시: 국가법령정보센터 '최신법령'·'최신행정규칙' 목록에서 최근 공포·발령된 것 가운데
전체 목록(assets/law/statutory_catalog.csv)에 있는 법령·행정규칙과 새로 제정된 것을 골라 원문을 다시 받고,
보존·보관 기간 조문을 다시 뽑아 이전 목록과 비교한다(생김·바뀜·사라짐). --apply면 그 법령의 행을 새것으로 바꾼다.

사용:  python statutory_watch.py [--days 30] [--report 산출/법정보존기간_변화.md] [--apply] [--max 60]
       (--offline 목록.json: 시험용 — {"최신": [...], "본문": {lsiSeq: 텍스트}})
빠르게(외부 API·비용 없음): 이미 읽은 판본(lsiSeq·admRulSeq, assets/law/statutory_seen.json)은 다시 받지 않고(--refresh로 무시),
       원문은 동시에 3개씩(--workers, 최대 4) 받으며 각 요청 뒤에 0.7초 쉰다. 새로 뽑은 행에는 보존 의무 판정기(retention_judge.py) 판정을 붙인다.
원칙: 개정된 법령의 별표 원문은 다시 받지 않는다(조문만). 자치법규는 보지 않는다.
       점검 표(statutory_minimum.csv)에 있는 법령이 바뀌면 '점검 표 확인'으로 따로 알린다 — 점검 표는 사람이 고친다.
종료:  0 바뀐 것 없음 · 1 보존기간 조문이 생기거나 바뀜 · 2 입력 오류·받기 실패
"""
import argparse
import csv
import datetime as dt
import html
import importlib.util
import json
import os
import re
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CATALOG = os.path.join(HERE, "..", "assets", "law", "statutory_catalog.csv")
MINIMUM = os.path.join(HERE, "..", "assets", "law", "statutory_minimum.csv")
ITEMS = os.path.join(HERE, "..", "assets", "law", "statutory_items.csv")  # 위임 추적기(retention_tree.py) 결과 — 바뀐 법령은 그 법령에서 시작하는 행을 다시 추적
SEEN = os.path.join(HERE, "..", "assets", "law", "statutory_seen.json")  # 이미 읽은 판본(법령 lsiSeq·행정규칙 admRulSeq) — 같은 판본은 다시 받지 않음
BASE = "https://www.law.go.kr/LSW/"
LISTS = {"법령": "lsSc.do?menuId=1&subMenuId=23&tabMenuId=121", "행정규칙": "admRulSc.do?menuId=5&subMenuId=45&tabMenuId=203"}
DELAY = 0.7


def _mod(name):
    sp = importlib.util.spec_from_file_location(name, os.path.join(HERE, name + ".py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


def judge_mod():
    sp = importlib.util.spec_from_file_location("retention_judge", os.path.join(HERE, "retention_judge.py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


def builder():
    sp = importlib.util.spec_from_file_location("statutory_catalog_build", os.path.join(HERE, "statutory_catalog_build.py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


def get(url):
    err = None
    for i in range(5):
        try:
            return urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=120).read().decode("utf-8", "ignore")
        except Exception as e:
            err = e
            time.sleep(4 + 3 * i)
    raise err


def text(h):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", h))).strip()


def ymd(s):
    m = re.search(r"(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.", s or "")
    return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def recent(kind, since, pages=30):
    """최신 목록을 앞에서부터 읽어 공포(발령)일이 since 이후인 것만."""
    out = []
    for pg in range(1, pages + 1):
        t = get(f"{BASE}{LISTS[kind]}&outmax=150&pg={pg}")
        trs = [tr for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", t, re.S) if "lawInfoPop" in tr or "lawInfPop" in tr]
        if not trs:
            break
        old = 0
        for tr in trs:
            row = text(tr)
            a = re.search(r"(?:lawInfoPop|lawInfPop)\('(\d+)'", tr)
            dates = re.findall(r"\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.", row)
            link = re.search(r"<a[^>]*(?:lawInfoPop|lawInfPop)[^>]*>(.*?)</a>", tr, re.S)  # 이름은 링크 글자에서
            name = text(link.group(1)) if link else ""
            if kind == "법령":
                ann, eff = ymd(dates[0] if dates else ""), ymd(dates[1] if len(dates) > 1 else "")
            else:
                ann = ymd(re.search(r"발령\s*(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.)", row).group(1)) if "발령" in row else (ymd(dates[0]) if dates else None)
                eff = ymd(re.search(r"시행\s*(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.)", row).group(1)) if re.search(r"시행\s*\d{4}", row) else None
            kindtxt = "제정" if "제정" in row else ("폐지" if "폐지" in row else "개정")
            if ann and ann < since:
                old += 1
                continue
            rk = re.search(r"(\S*(?:고시|훈령|예규|지침|규정|공고|규칙))\s*제\s*[\d-]+호", row) if kind == "행정규칙" else None
            out.append({"종류": kind, "이름": name, "규칙종류": rk.group(1) if rk else "", "seq": a.group(1) if a else "", "공포": ann.isoformat() if ann else "",
                        "시행": eff.isoformat() if eff else "", "구분": kindtxt})
        if old == len(trs):
            break
        time.sleep(DELAY)
    return out


def body(item):
    if item["종류"] == "법령":
        ef = item["시행"].replace("-", "")
        b = get(f"{BASE}lsInfoR.do?lsiSeq={item['seq']}&chrClsCd=010202&urlMode=lsInfoP&efYd={ef}&ancYnChk=0")
    else:
        b = get(f"{BASE}admRulLsInfoR.do?admRulSeq={item['seq']}")
    if len(b) < 2000:
        raise RuntimeError("원문이 비었습니다")
    return text(re.sub(r"<script.*?</script>|<style.*?</style>", " ", b, flags=re.S))


def key(r):
    """비교 열쇠 — 화면마다 다른 띄어쓰기('제84조제5항'·'제84조 제5항')는 무시."""
    return (r["조문"].split("(")[0], r["보존기간"], r.get("성격", "") or "최소 보존", re.sub(r"\s", "", r.get("기록(문구)") or "")[-30:])


def main() -> int:
    global DELAY
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=30, help="며칠 전 공포·발령분부터 볼지(기본 30일 — 3주 감시와 겹치게)")
    ap.add_argument("--report", default="산출/법정보존기간_변화.md")
    ap.add_argument("--apply", action="store_true", help="목록(statutory_catalog.csv)의 해당 법령 행을 새 추출로 바꿈")
    ap.add_argument("--catalog", default=CATALOG)
    ap.add_argument("--items", default=ITEMS, help="위임 추적 목록(없으면 건너뜀)")
    ap.add_argument("--max", type=int, default=200, help="원문을 받을 최대 개수")
    ap.add_argument("--delay", type=float, default=0.7)
    ap.add_argument("--workers", type=int, default=3, help="원문을 동시에 받을 개수(기본 3 — 사이트에 부담을 주지 않게 4 이하 권장)")
    ap.add_argument("--seen", default=SEEN, help="이미 읽은 판본 기록 파일")
    ap.add_argument("--refresh", action="store_true", help="판본 기록을 무시하고 모두 다시 받음")
    ap.add_argument("--offline", help="(시험용) {'최신': [...], '본문': {seq: 텍스트}} JSON")
    a = ap.parse_args()
    DELAY = a.delay
    b = builder()
    jd = judge_mod()
    try:
        seen = {} if a.offline and a.seen == SEEN else json.load(open(a.seen, encoding="utf-8"))
    except (OSError, ValueError):
        seen = {}
    sk = lambda x: f"{x['종류']}|{x['이름']}|{x.get('규칙종류', '')}"
    with open(a.catalog, encoding="utf-8-sig") as f:
        cat = list(csv.DictReader(f))
    cols = list(cat[0].keys()) if cat else []
    with open(MINIMUM, encoding="utf-8-sig") as f:
        curated = {r["법령"].split("(")[0].strip() for r in csv.DictReader(f)}
    names = {r["법령"] for r in cat if r["출처"] in ("법령", "행정규칙")}
    try:  # 시험(--offline)에서는 --items를 따로 줄 때만 씀 — 실제 목록을 건드리지 않음
        with open(a.items, encoding="utf-8-sig") as f:
            items_rows = [] if a.offline and a.items == ITEMS else list(csv.DictReader(f))
    except OSError:
        items_rows = []
    root = lambda r: r["근거 경로"].split(" → ")[0].rsplit(" ", 1)[0]
    for r in items_rows:  # 위임으로 이어진 시행령·시행규칙·고시가 바뀌어도 다시 읽음
        if r.get("출처") in ("법령", "행정규칙"):
            names.add(r["법령"])

    def same(r, x, src):  # 같은 법령·규칙인가 — 행정규칙은 이름이 같은 다른 기관 규칙이 있어 종류(○○청예규 등)까지 봄
        return r["법령"] == x["이름"] and r["출처"] == src and (src == "법령" or not x.get("규칙종류") or not r.get("종류") or r["종류"] == x["규칙종류"])
    since = dt.date.today() - dt.timedelta(days=a.days)
    off = json.load(open(a.offline, encoding="utf-8")) if a.offline else None
    try:
        if off:
            items = off["최신"]
        else:  # 최신법령·최신행정규칙 목록을 동시에 받음
            from concurrent.futures import ThreadPoolExecutor
            with ThreadPoolExecutor(max_workers=2) as ex:
                fl, fa = ex.submit(recent, "법령", since), ex.submit(recent, "행정규칙", since)
                items = fl.result() + fa.result()
    except Exception as e:
        print(f"받기 실패: {e}", file=sys.stderr)
        return 2
    allpick = [x for x in items if x["이름"] in names or (x["구분"] == "제정" and x["종류"] == "법령")]
    latest = {}  # 같은 법령이 기간 안에 두 번 시행되면 마지막 시행 판본만 본다(시행 예정 판본은 따로 모두 알림)
    today = dt.date.today().isoformat()
    for x in allpick:
        k = sk(x)
        if x["시행"] and x["시행"] > today:
            continue
        if k not in latest or (x.get("seq", "").isdigit() and int(x["seq"]) > int(latest[k].get("seq") or 0)):
            latest[k] = x
    allpick = [x for x in allpick if (x["시행"] and x["시행"] > today) or latest[sk(x)] is x]
    same_ver = [x for x in allpick if not a.refresh and x.get("seq") and seen.get(sk(x)) == x["seq"]]
    pick = [x for x in allpick if x not in same_ver][: a.max]
    changes, newrows, fail = [], {}, []
    def fetch(x):  # 원문 받기만 동시에 하고(무료·추가 비용 없음), 추출·비교는 아래에서 차례로
        if off:
            return off["본문"][x["seq"]]
        t = body(x)
        time.sleep(DELAY)
        return t
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=max(1, min(a.workers, 4))) as ex:
        futs = [(x, ex.submit(fetch, x)) for x in pick]
    for x, fu in futs:
        try:
            t = fu.result()
        except Exception as e:
            fail.append(f"{x['이름']}: {e}")
            continue
        label = f"{x['이름']} [시행 {x['시행'].replace('-', '. ')}.]" if x["시행"] else x["이름"]
        rows = b.extract(label, t)
        for r in rows:
            r["판정"] = jd.judge(r.get("문장", ""), r.get("성격", ""))[1]
            if x["종류"] == "행정규칙":
                jo = r["조문"].split("(")[0]
                r.update({"출처": "행정규칙", "종류": x.get("규칙종류", ""), "원문": f"https://www.law.go.kr/행정규칙/{x['이름'].replace(' ', '')}" + ("" if jo == "본문" else f"/{jo}")})
        src = "법령" if x["종류"] == "법령" else "행정규칙"
        before = {key(r): r for r in cat if same(r, x, src)}
        after = {key(r): r for r in rows}
        add = [after[k] for k in after if k not in before]
        gone = [before[k] for k in before if k not in after]
        future = bool(x["시행"]) and x["시행"] > dt.date.today().isoformat()
        if add or gone:
            changes.append((x, add, gone, future))
        else:
            seen[sk(x)] = x["seq"]  # 바뀐 것 없음(시행 예정 판본 포함) → 이 판본은 다음부터 건너뜀
        if not future:  # 시행 전인 개정은 알리기만 하고 목록은 바꾸지 않음(목록은 현행 기준)
            newrows[(src, x["이름"], x.get("규칙종류", ""))] = (x, rows)
    lines = [f"# 법정 보존기간 목록 변화 — {dt.date.today().isoformat()}", "",
             f"최신법령·최신행정규칙 {len(items)}건(공포·발령 {since.isoformat()} 이후: 법령 {sum(x['종류'] == '법령' for x in items)}·행정규칙 {sum(x['종류'] == '행정규칙' for x in items)}) 중 "
             f"목록에 있거나 새로 제정된 {len(allpick)}건 가운데 이미 읽은 판본 {len(same_ver)}건은 건너뛰고 {len(pick)}건의 원문을 다시 읽음. "
             f"보존기간 조문이 바뀐 법령 {len(changes)}건" + (f", 못 받음 {len(fail)}건" if fail else "") + ".", ""]
    for x, add, gone, future in changes:
        mark = " ⚠️ 점검 표(statutory_minimum.csv)에 있는 법령 — 점검 표도 확인" if x["이름"] in curated else ""
        lines.append(f"## {x['이름']} ({x['구분']}, 공포 {x['공포']} · 시행 {x['시행']}{' — 시행 예정: 목록은 시행일 뒤에 바꿈' if future else ''}){mark}")
        for r in add:
            lines.append(f"- 🆕 {r['조문']} {r['보존기간']}({r.get('성격', '')}) [판정: {r.get('판정', '')}] — {(r.get('기록(문구)') or r['문장'])[-80:]}")
        for r in gone:
            lines.append(f"- ➖ {r['조문']} {r['보존기간']} — {(r.get('기록(문구)') or r['문장'])[-80:]}")
        lines.append("")
    for f_ in fail:
        lines.append(f"- 못 받음: {f_}")
    os.makedirs(os.path.dirname(a.report) or ".", exist_ok=True)
    open(a.report, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    now = [c for c in changes if not c[3]]
    # 위임 추적 목록: 다시 읽은(시행 중) 법령과, 그 법령을 거쳐 가는 행의 출발 법령을 다시 추적
    item_changes = []
    if items_rows and newrows:
        rt = _mod("retention_tree")
        fresh = {x["이름"]: t for x, fu in futs if not (x["시행"] and x["시행"] > today) for t in [fu.result()] if not fu.exception()} if not off else \
                {x["이름"]: off["본문"][x["seq"]] for x in pick if not (x["시행"] and x["시행"] > today)}
        corpus = rt.Corpus(dict(fresh), {}, live=not off, live_laws=not off)
        roots = set(fresh) | {root(r) for r in items_rows if any(n in r["근거 경로"] for n in fresh)}
        old_by_root = {}
        for r in items_rows:
            old_by_root.setdefault(root(r), []).append(r)
        new_items = []
        ik = lambda r: (r["법령"], r["조문"].split("(")[0], r["항"], r["호"], re.sub(r"\s", "", r["보존기간"]), re.sub(r"\s", "", r["기록(항목)"])[:30])  # 띄어쓰기 차이(원문 받는 곳마다 다름)는 같은 것으로
        all_old = {ik(r) for r in items_rows}
        regen = set()
        for law in sorted(roots):
            if not corpus.arts(law):  # 원문을 못 받으면 그 출발 법령의 행은 그대로 둠
                continue
            rows_ = []
            for jo, title, abody in corpus.arts(law):
                if rt.KEEPV.search(abody) or rt.EVENT.search(abody):
                    rows_ += rt.analyze(corpus, law, jo, title, abody)
            rows_ = rt.add_judgement([r for r in rows_ if len(re.sub(r"[\s\d.\[\]]", "", r["기록(항목)"])) >= 2])
            src = next((o.get("출처") for o in old_by_root.get(law, [])), "법령")
            for r in rows_:
                r.setdefault("출처", src if r["법령"] == law else ("행정규칙" if r.get("위임") == "고시" else "법령"))
                r.setdefault("기관", "")
            ob, nb = {ik(r) for r in old_by_root.get(law, [])}, {ik(r) for r in rows_}
            regen |= nb
            new_items.append((law, rows_, ob, nb))
        for law, rows_, ob, nb in new_items:  # 다른 출발 법령에서 이미 있던 행(같은 위임 사슬)은 새것으로 치지 않음
            add_, gone_ = sorted(nb - all_old), sorted(ob - regen)
            if add_ or gone_:
                item_changes.append((law, add_, gone_))
        if item_changes:
            lines.append("## 위임 추적(호·항·단서, 시행령·시행규칙·고시) 변화")
            for law, add_, gone_ in item_changes:
                lines.append(f"- {law}: 🆕 {len(add_)} · ➖ {len(gone_)}" + "".join(f"\n  - 🆕 {k[1]}{k[2]}{k[3]} {k[4]} — {k[5]}" for k in add_[:8])
                             + "".join(f"\n  - ➖ {k[1]}{k[2]}{k[3]} {k[4]} — {k[5]}" for k in gone_[:8]))
            lines.append("")
            open(a.report, "w", encoding="utf-8").write("\n".join(lines) + "\n")
        if a.apply and item_changes:
            changed = {law for law, _, _ in item_changes}
            keep_i = [r for r in items_rows if root(r) not in changed]
            have = {ik(r) for r in keep_i}
            for law, rows_, _, _ in new_items:
                if law in changed:
                    for r in rows_:
                        if ik(r) not in have:
                            have.add(ik(r))
                            keep_i.append(r)
            old_byl = {ik(r): r.get("별표", "") for r in items_rows if r.get("별표")}  # 다시 추적한 행도 붙여 둔 별표(statutory_byl.py)는 이어받음
            for r in keep_i:
                if not r.get("별표") and ik(r) in old_byl:
                    r["별표"] = old_byl[ik(r)]
            icols = list(items_rows[0].keys())
            with open(a.items, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=icols, extrasaction="ignore")
                w.writeheader()
                w.writerows(keep_i)
    if a.apply and now:
        hit = [(k, v) for k, v in newrows.items() if any(c[0] is v[0] for c in now)]
        keep = [r for r in cat if not any(same(r, x, k[0]) for k, (x, _) in hit)]
        for k, (x, rows) in hit:
            keep += [{c: b.RRN.sub("******-*******", str(r.get(c, "") or "")) for c in cols} for r in rows]
        with open(a.catalog, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            w.writerows(keep)
        for k, (x, _) in hit:
            seen[sk(x)] = x["seq"]
    if not off or a.seen != SEEN:
        with open(a.seen, "w", encoding="utf-8") as f:
            json.dump(dict(sorted(seen.items())), f, ensure_ascii=False, indent=0)
    print(f"최신 {len(items)} · 같은 판본 건너뜀 {len(same_ver)} · 다시 읽음 {len(pick)} · 바뀐 법령 {len(changes)}(시행 예정 {sum(c[3] for c in changes)})"
          + f" · 위임 추적 바뀐 법령 {len(item_changes)}" + (" · 목록 고침" if a.apply and (now or item_changes) else "") + f" → {a.report}")
    return 2 if fail and not changes else (1 if changes or item_changes else 0)


if __name__ == "__main__":
    sys.exit(main())
