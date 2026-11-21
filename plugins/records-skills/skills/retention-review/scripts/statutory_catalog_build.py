#!/usr/bin/env python3
"""법정 보존기간 전체 목록(assets/law/statutory_catalog.csv) 다시 만들기 — 국가법령정보센터(www.law.go.kr) 원문에서.

순서(2026-11-21에 처음 만든 방법 그대로):
  1. 조문내용 검색: 현행 법령(lsSc.do 조문내용 탭)과 현행 행정규칙(admRulSc.do 조문내용 탭)을 보존·보관 기간 문구
     14가지로 검색해 법령·조문·문구·일련번호를 모은다(한 쪽 150건, 쪽이 빌 때까지).
  2. 법령 원문: 검색에 걸린 법령마다 본문(lsInfoR.do, 시행일 efYd)을 받아 조문 단위로 나누고, 부칙을 잘라 낸다.
  3. 추출: 'N년(간·동안·이상) … 보존·보관·비치·유지·저장·관리하여야/할 것', '보존기간은 … N년으로 한다',
     보존 의무 조문의 나열형 '1. 기록 : N년'을 뽑는다. 날짜(2015년·1월 1일), 기한(N일 이내), 벌칙·과태료·처분 문장,
     '최근 N년간', 'N년 이상이 지난', 주식·부동산 보유기간은 뺀다.
  4. 행정규칙은 원문(admRulLsInfoR.do)을 받아 같은 방법으로 뽑고, 원문에서 못 찾은 조문만 검색 문구로 담는다.
     자치법규(ordinSc.do 조문내용 탭, 현행 조례·규칙)도 기간 문구가 걸린 것은 원문(ordinInfoR.do)에서 뽑는다.
     별표·서식은 검색 결과 문구(일부)만 담는다
     (원문 주소·지자체 함께). 주민등록번호 모양 문자열은 가린다.
사용:  python statutory_catalog_build.py -o 산출/statutory_catalog.csv [--work 산출/법령원문] [--delay 0.7] [--limit N]
       (--work 폴더에 받은 원문을 두어 다시 돌릴 때 이어서 받는다. 법령 1,800여 개라 30분~1시간 걸린다.)
주의:  국가법령정보센터에 부담을 주지 않게 한 번에 하나씩, 요청 사이를 띄운다. 오픈API(DRF)는 등록한 서버 주소만 받으므로
       쓰지 않는다. 결과는 검토용 '후보 목록'이다 — 적용 대상·기산일은 원문을 열어 확인한다.
"""
import argparse
import csv
import html
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

BASE = "https://www.law.go.kr/LSW/"
QUERIES = ["년간 보존", "년간 보관", "년 동안 보존", "년 동안 보관", "년 이상 보존", "년 이상 보관", "개월간 보존", "개월간 보관",
           "보존기간", "보관기간", "보유기간", "년간 비치", "영구 보존", "년간 기록"]
RRN = re.compile(r"(?<!\d)\d{6}-?[1-4]\d{6}(?!\d)")
NUM = r"(?:\d+)"
PERIOD = re.compile(rf"(영구(?:히)?|준영구|{NUM}\s*(?:년|개월|일|시간)(?:\s*(?:간|동안|이상|이내|까지))?)")
VERB = re.compile(r"(보존|보관|비치|유지|저장|관리)(?:하여야|해야|하여야만|하고|한다|할 수|할 것|하도록|하여 두어야|해 두어야|하며|기간)")
ART = re.compile(r"(?<![가-힣])(제\d+조(?:의\d+)?)\s*\(([^)]{1,60})\)")
NOISE = re.compile(r"행정처분|처분기준|과태료|벌칙|영업정지|업무정지|아니한 경우|않은 경우|위반")
SNIP = re.compile(r"((?:\d+)\s*(?:년|개월)\s*(?:간|동안|이상)?|영구(?:히)?|준영구)\s*(?:[^.│┃]{0,25}?)(보존|보관|비치)(?:하여야|해야|할 것|ㆍ관리하여야|하고|한다)")
DELAY = 0.7


def get(url):
    err = None
    for i in range(6):
        try:
            return urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=120).read().decode("utf-8", "ignore")
        except Exception as e:  # 연결이 끊기면 조금 쉬고 다시
            err = e
            time.sleep(5 + 3 * i)
    raise err


def text(h):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", h))).strip()


def search_laws(q):
    out, pg, law = [], 1, None
    while pg <= 80:
        t = get(f"{BASE}lsSc.do?menuId=1&subMenuId=15&tabMenuId=85&tabMenuIdx=3&outmax=150&pg={pg}&query=" + urllib.parse.quote(q))
        i = t.find("<!-- lsScListR -->")
        got = 0
        for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", t[i:t.find("</table>", i)], re.S):
            tds = [text(x) for x in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)]
            pops = re.findall(r"lawInfoPop\(([^)]*)\)", tr)
            if not tds or not pops:
                continue
            a = [x.strip().strip("'") for x in pops[-1].split(",")]
            if len(tds) >= 4:
                law, jo, snip = tds[1], tds[2], tds[3]
            else:
                jo, snip = tds[-2], tds[-1]
            out.append({"법령": law, "조문": jo, "문구": snip, "lsiSeq": a[0], "bylSeq": a[9] if len(a) > 9 else ""})
            got += 1
        if not got:
            break
        pg += 1
        time.sleep(DELAY)
    return out


def search_rules(q):
    out, pg = [], 1
    while pg <= 60:
        t = get(f"{BASE}admRulSc.do?menuId=5&subMenuId=41&tabMenuId=187&tabMenuIdx=3&outmax=150&pg={pg}&query=" + urllib.parse.quote(q))
        i = t.find("<tbody", t.find("검색결과 목록"))
        got = 0
        for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", t[i:t.find("</table>", i)], re.S):
            pops = re.findall(r"lawInfPop\(([^)]*)\)", tr)
            if not pops:
                continue
            a = [x.strip().strip("'") for x in pops[-1].split(",")]
            parts = re.split(r'"\s*target="_blank">', re.sub(r"<[^>]+>", "", re.sub(r"\s+", " ", html.unescape(tr))))
            rest = parts[2].strip() if len(parts) > 2 else ""
            jo, _, snip = rest.partition("...")
            out.append({"규칙": parts[1].strip() if len(parts) > 1 else "", "조문": jo.strip(), "문구": snip.strip(), "admSeq": a[0]})
            got += 1
        if not got:
            break
        pg += 1
        time.sleep(DELAY)
    return out


ORD_QUERIES = ["년간 보존", "년간 보관", "년 동안 보존", "년 동안 보관", "년 이상 보존", "년 이상 보관", "영구 보존", "보존기간", "보관기간"]


def search_ordin(q):
    out, pg, cur, org = [], 1, None, ""
    while pg <= 250:
        t = get(f"{BASE}ordinSc.do?menuId=3&subMenuId=27&tabMenuId=143&tabMenuIdx=3&outmax=150&pg={pg}&query=" + urllib.parse.quote(q))
        got = 0
        for tr in re.findall(r'<tr class="dt">(.*?)</tr>', t, re.S):
            tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
            tit = [x for x in tds if 'id="list"' in x]
            if tit:
                m = re.search(r"ordinInfPop\('(\d+)'", tit[0])
                cur, i = (m.group(1) if m else "", text(tit[0])), tds.index(tit[0])
                org, rest = (text(tds[i + 1]) if i + 1 < len(tds) else ""), tds[i + 2:]
            else:
                rest = [x for x in tds if "ordinInfPop" in x]
            if len(rest) >= 2 and cur:
                out.append({"자치법규": cur[1], "seq": cur[0], "기관": org, "조문": text(rest[-2]), "문구": text(rest[-1])})
                got += 1
        if not got:
            break
        pg += 1
        time.sleep(DELAY)
    return out


def years(p):
    if "준영구" in p:
        return 999
    if "영구" in p:
        return 1000
    m = re.match(rf"({NUM})\s*(년|개월|일|시간)", p)
    return int(m.group(1)) / {"년": 1, "개월": 12, "일": 365, "시간": 8760}[m.group(2)] if m else None


def kind(name):
    if "시행규칙" in name:
        return "시행규칙"
    if "시행령" in name:
        return "시행령"
    if re.search(r"규칙|규정|령\s*\[", name) and "법률" not in name:
        return "대통령령·부령 등"
    return "법률"


def extract(name, t):
    law = re.sub(r"\s*\[.*", "", name).strip()
    shi = re.search(r"\[시행 ([^\]]+)\]", name)
    st = t.find("제1조(")
    if st < 0:
        return []
    t = t[st:]
    cut = re.search(r"부\s*칙\s*(?:\([^)]{0,80}\)\s*)?<[^>]{0,40}제\s*[\d-]+호", t)
    t = t[:cut.start()] if cut else t
    rows, seen = [], set()
    ms = list(ART.finditer(t))

    def add(jo, title, rec, per, base, subj, sent):
        rows.append({"출처": "법령", "법령": law, "종류": kind(name), "시행": shi.group(1) if shi else "", "조문": f"{jo}({title})", "의무자": subj,
                     "기록(문구)": rec[-120:], "보존기간": re.sub(r"\s+|(간|동안)$", "", per), "기산(문구)": base, "문장": sent[:400],
                     "원문": f"https://www.law.go.kr/법령/{law.replace(' ', '')}/{jo}"})

    for i, m in enumerate(ms):
        jo, title = m.group(1), m.group(2).strip()
        body = t[m.end(): ms[i + 1].start() if i + 1 < len(ms) else len(t)]
        if (jo, title) in seen:
            continue
        seen.add((jo, title))
        if re.search(r"(보존|보관)(하여야|해야|기간)", body):
            for lm in re.finditer(rf"(?:(?<=\s)|^)(\d+\.|[가-하]\.)\s*([^:：.]{{2,90}}?)\s*[:：]\s*({PERIOD.pattern})", body):
                if years(lm.group(3)) is None or re.match(r"\d{4}\s*년", lm.group(3)) or (jo, lm.group(3), lm.group(2)[-60:]) in seen:
                    continue
                seen.add((jo, lm.group(3), lm.group(2)[-60:]))
                add(jo, title, lm.group(2).strip(), lm.group(3), "", "", body[max(0, lm.start() - 120): lm.end() + 20])
        for s in [x.strip() for x in re.split(r"(?<=다\.)\s+|(?=[①-⑳])", body) if x.strip()]:
            s = s[:900]
            for pm in PERIOD.finditer(s):
                p = pm.group(1)
                tail = re.sub(r"\([^)]{0,60}\)", "", s[pm.end(): pm.end() + 110])[:45]
                head = s[max(0, pm.start() - 50): pm.start()]
                if re.search(r"최근\s*$", s[max(0, pm.start() - 6): pm.start()]) or re.match(r"\s*이상이?\s*(지난|경과|된|근무|재직|거주)", s[pm.end() - 2: pm.end() + 12]):
                    continue
                form = bool(re.search(r"(보존|보관|보유)\s*기간[은는이가]?\s*[^.]{0,80}$", head) and re.match(r"\s*(?:으로|로)\s*(?:한다|하되|하며|정한다)", tail))
                if not (VERB.search(tail) or form) or re.match(r"\s*(?:이내에|안에|까지\s*(?:신고|제출|통보))", tail):
                    continue
                if re.search(r"(징역|벌금|과태료|자격정지|금고|정지|이내의 기간)", s[max(0, pm.start() - 30): pm.end() + 10]):
                    continue
                if p.replace(" ", "").endswith("이내") or (re.search(r"(주식|주택|토지|지분|출자|자산|부동산)", s[max(0, pm.start() - 60): pm.start()])
                                                         and not re.search(r"(서류|기록|장부|자료|대장|증빙)", s[max(0, pm.start() - 80): pm.end() + 20])):
                    continue
                y = years(p)
                if y is None or re.match(r"\d{4}\s*년", p) or (re.search(r"\d\s*(일|시간)$", p.strip()) and not re.search(r"(간|동안|이상)$", p)):
                    continue
                pre = s[max(0, pm.start() - 140): pm.start()]
                rec = re.split(r"(?:한다|된다|하고|하며)[.,]?\s|[①-⑳]|\d+\.\s", pre)[-1].strip(" ,")
                if (jo, p, rec[-60:]) in seen:
                    continue
                seen.add((jo, p, rec[-60:]))
                base = re.search(r"([^,.]{0,40}(?:부터|한 날|종료|끝난|완료|후))\s*$", pre)
                subj = re.match(r"([^,]{1,60}?(?:는|은))\s", s)
                add(jo, title, rec, p, base.group(1).strip() if base else "", subj.group(1) if subj else "", s)
    return rows


def main() -> int:
    global DELAY
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default="산출/statutory_catalog.csv")
    ap.add_argument("--work", default="산출/법령원문", help="받은 법령 원문을 두는 폴더(이어 받기)")
    ap.add_argument("--delay", type=float, default=0.7)
    ap.add_argument("--limit", type=int, help="(시험용) 법령 원문을 앞에서 N개만")
    a = ap.parse_args()
    DELAY = a.delay
    os.makedirs(a.work, exist_ok=True)
    hits, rules = {}, {}
    for q in QUERIES:
        for r in search_laws(q):
            hits.setdefault((r["lsiSeq"], r["조문"], r["bylSeq"]), r)
        for r in search_rules(q):
            rules.setdefault((r["admSeq"], r["조문"]), r)
        print(f"검색 '{q}': 법령 조문 {len(hits)} · 행정규칙 조문 {len(rules)}", file=sys.stderr)
    ords = {}
    for q in ORD_QUERIES:
        for r in search_ordin(q):
            ords.setdefault((r["seq"], r["조문"]), r)
        print(f"자치법규 검색 '{q}': 조문 {len(ords)}", file=sys.stderr)
    names = {}
    for r in hits.values():
        names.setdefault(r["lsiSeq"], r["법령"] or "")
    out = []
    for n, (seq, name) in enumerate(sorted(names.items())):
        if a.limit and n >= a.limit:
            break
        fn = os.path.join(a.work, f"{seq}.txt")
        if not (os.path.exists(fn) and os.path.getsize(fn) > 1000):
            m = re.search(r"\[시행 (\d{4})\. (\d{1,2})\. (\d{1,2})\.\]", name)
            ef = f"{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}" if m else ""
            try:
                b = get(f"{BASE}lsInfoR.do?lsiSeq={seq}&chrClsCd=010202&urlMode=lsInfoP&efYd={ef}&ancYnChk=0")
            except Exception as e:
                print(f"못 받음 {seq} {name[:30]}: {e}", file=sys.stderr)
                continue
            if len(b) < 3000:
                continue
            open(fn, "w", encoding="utf-8").write(text(re.sub(r"<script.*?</script>|<style.*?</style>", " ", b, flags=re.S)))
            time.sleep(DELAY)
        out += extract(name, open(fn, encoding="utf-8").read())
    seen = set()
    for r in hits.values():  # 별표·서식: 검색 문구
        if not r["bylSeq"] or NOISE.search(r["조문"]):
            continue
        law = re.sub(r"\s*\[.*", "", r["법령"] or "").strip()
        for m in SNIP.finditer(r["문구"]):
            ctx = r["문구"][max(0, m.start() - 80): m.end() + 10]
            k = (law, r["조문"], m.group(1), ctx[:40])
            if NOISE.search(ctx) or k in seen:
                continue
            seen.add(k)
            shi = re.search(r"\[시행 ([^\]]+)\]", r["법령"] or "")
            out.append({"출처": "별표·서식", "법령": law, "종류": "", "시행": shi.group(1) if shi else "", "조문": r["조문"][:80], "의무자": "", "기록(문구)": "",
                        "보존기간": re.sub(r"\s|(간|동안)$", "", m.group(1)), "기산(문구)": "", "문장": ctx.strip(" .│┃"), "원문": "https://www.law.go.kr/법령/" + law.replace(" ", "")})
    rname = {}
    for r in rules.values():
        n = r["규칙"]
        if n and "target" not in n and '"' not in n and not n.startswith("제") and len(n) >= 4:
            rname.setdefault(r["admSeq"], n)
    got = set()
    for n, (seq, full) in enumerate(sorted(rname.items())):  # 행정규칙 원문(admRulLsInfoR.do)에서 조문 추출
        if a.limit and n >= a.limit:
            break
        fn = os.path.join(a.work, f"adm_{seq}.txt")
        if not (os.path.exists(fn) and os.path.getsize(fn) > 800):
            try:
                b = get(f"{BASE}admRulLsInfoR.do?admRulSeq={seq}")
            except Exception as e:
                print(f"못 받음 행정규칙 {seq}: {e}", file=sys.stderr)
                continue
            if len(b) < 2000:
                continue
            open(fn, "w", encoding="utf-8").write(text(re.sub(r"<script.*?</script>|<style.*?</style>", " ", b, flags=re.S)))
            time.sleep(DELAY)
        rule = re.sub(r"\s*\[시행.*", "", full).strip()
        kd = re.search(r"\]\s*\[([^\]]*?)\s*제[\d-]", full)
        shi = re.search(r"\[시행 ([^\]]+)\]", full)
        for r in extract(f"{rule} [시행 {shi.group(1) if shi else ''}]", open(fn, encoding="utf-8").read()):
            jo = r["조문"].split("(")[0]
            r.update({"출처": "행정규칙", "종류": kd.group(1) if kd else "", "원문": f"https://www.law.go.kr/행정규칙/{rule.replace(' ', '')}/{jo}"})
            out.append(r)
            got.add((seq, jo))
    for r in rules.values():  # 원문에서 못 찾은 조문은 검색 문구로
        full = rname.get(r["admSeq"])
        jm = re.match(r"(제\d+조(?:의\d+)?)", r["조문"])
        if not full or (jm and (r["admSeq"], jm.group(1)) in got):
            continue
        rule = re.sub(r"\s*\[시행.*", "", full).strip()
        shi = re.search(r"\[시행 ([^\]]+)\]", full)
        kd = re.search(r"\]\s*\[([^\]]*?)\s*제[\d-]", full)
        for m in SNIP.finditer(r["문구"]):
            ctx = r["문구"][max(0, m.start() - 90): m.end() + 10]
            k = (rule, r["조문"][:40], m.group(1), ctx[:40])
            if NOISE.search(ctx) or k in seen:
                continue
            seen.add(k)
            out.append({"출처": "행정규칙(검색 문구)", "법령": rule, "종류": kd.group(1) if kd else "", "시행": shi.group(1) if shi else "", "조문": r["조문"][:80], "의무자": "",
                        "기록(문구)": "", "보존기간": re.sub(r"\s|(간|동안)$", "", m.group(1)), "기산(문구)": "", "문장": ctx.strip(" ."),
                        "원문": "https://www.law.go.kr/행정규칙/" + rule.replace(" ", "")})
    oget = set()
    seqs = sorted({r["seq"] for r in ords.values() if r["seq"] and SNIP.search(r["문구"])})
    for n, seq in enumerate(seqs):  # 자치법규 원문(ordinInfoR.do) — 기간 문구가 걸린 것만
        if a.limit and n >= a.limit:
            break
        v = next(r for r in ords.values() if r["seq"] == seq)
        fn = os.path.join(a.work, f"ord_{seq}.txt")
        if not (os.path.exists(fn) and os.path.getsize(fn) > 800):
            try:
                b = get(f"{BASE}ordinInfoR.do?ordinSeq={seq}")
            except Exception as e:
                print(f"못 받음 자치법규 {seq}: {e}", file=sys.stderr)
                continue
            if len(b) < 2000:
                continue
            open(fn, "w", encoding="utf-8").write(text(re.sub(r"<script.*?</script>|<style.*?</style>", " ", b, flags=re.S)))
            time.sleep(DELAY)
        name = re.sub(r"\s*\[시행.*", "", v["자치법규"]).strip()
        kd = re.search(r"\]\[(조례|규칙|훈령|예규|규정)", v["자치법규"])
        for r in extract(v["자치법규"].replace("][", "] ["), open(fn, encoding="utf-8").read()):
            r.update({"출처": "자치법규", "법령": name, "종류": kd.group(1) if kd else "", "기관": v["기관"], "원문": "https://www.law.go.kr/자치법규/" + name.replace(" ", "")})
            out.append(r)
            oget.add((seq, r["조문"].split("(")[0]))
    for r in ords.values():  # 원문에서 못 찾은 조문은 검색 문구로
        jm = re.match(r"(제\d+조(?:의\d+)?)", r["조문"])
        if jm and (r["seq"], jm.group(1)) in oget:
            continue
        name = re.sub(r"\s*\[시행.*", "", r["자치법규"]).strip()
        shi = re.search(r"\[시행 ([^\]]+)\]", r["자치법규"])
        kd = re.search(r"\]\[(조례|규칙|훈령|예규|규정)", r["자치법규"])
        for m in SNIP.finditer(r["문구"]):
            if re.search(r"최근\s*$", r["문구"][max(0, m.start() - 8): m.start()]):
                continue
            ctx = r["문구"][max(0, m.start() - 90): m.end() + 10]
            k = (r["seq"], r["조문"], m.group(1), ctx[:40])
            if NOISE.search(ctx) or k in seen:
                continue
            seen.add(k)
            out.append({"출처": "자치법규(검색 문구)", "법령": name, "종류": kd.group(1) if kd else "", "시행": shi.group(1) if shi else "", "기관": r["기관"],
                        "조문": r["조문"][:80], "의무자": "", "기록(문구)": "", "보존기간": re.sub(r"\s|(간|동안)$", "", m.group(1)), "기산(문구)": "",
                        "문장": ctx.strip(" ."), "원문": "https://www.law.go.kr/자치법규/" + name.replace(" ", "")})
    cols = ["출처", "법령", "종류", "시행", "기관", "조문", "의무자", "기록(문구)", "보존기간", "기산(문구)", "문장", "원문"]
    for r in out:
        for c in cols:
            r[c] = RRN.sub("******-*******", (r.get(c) or "").strip())
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out)
    by = {}
    for r in out:
        by[r["출처"]] = by.get(r["출처"], 0) + 1
    print(f"법정 보존기간 전체 목록 {len(out)}건 {json.dumps(by, ensure_ascii=False)} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
