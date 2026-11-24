#!/usr/bin/env python3
"""보존 조문 위임 추적기 — 한 조문 안에서 항·호·단서마다 보존기간이 다른 경우를 호 단위로 나누고,
'○○령으로 정하는 서류'·'보존기간을 연장할 수 있다'·'장관이 정하여 고시하는 물질'처럼 아래 법령에 맡긴 부분은
법률 → 시행령·시행규칙 → 고시(행정규칙)까지 끝까지 따라가 그 법령의 해당 조문이 정한 기록과 기간을 붙인다.

예) 산업안전보건법 제164조① "다음 각 호의 서류를 3년(제2호의 경우 2년을 말한다) 동안 보존 … 다만, 고용노동부령으로 정하는 바에 따라 연장"
    → 1호 선임 서류 3년 · 2호 회의록 2년 · 3호 '고용노동부령으로 정하는 사항을 적은 서류' 3년(→ 시행규칙에서 그 호를 가리키는 조문의 사항)
    · 단서 → 시행규칙 제241조① 작업환경측정 결과 5년, 다만 '고용노동부장관이 정하여 고시하는 물질' 30년(→ 그 고시)

사용:  python retention_tree.py --law "산업안전보건법" --jo 제164조                  # 원문을 받아 그 조문을 끝까지 추적(인터넷)
       python retention_tree.py --bodies 본문폴더 --names 이름표.json [--adm 행정규칙폴더 --adm-names 이름표.json] -o statutory_items.csv
                                                                              # 받아 둔 원문 전체에 적용(인터넷 없음)
출력 칸: 법령·조문·항·호(목)·기록(항목)·보존기간·의무자·조건·근거 경로·위임·상태
상태: '확정'(기간이 그 문장에 있음) · '위임 추적'(아래 법령에서 찾아 붙임 — 법률이 기간 없이 '보존하여야'만 하고 시행령·시행규칙이 기간을 정한 숨은 위임 포함) · '위임 못 찾음'(가리키는 조문을 못 찾음 — 원문 확인) · '방법만 위임(기간 없음)'(서식·장소·작성 방법만 맡김) · '사건 기한'(○○할 때까지)
외부 API·비용 없음.
"""
import argparse
import csv
import importlib.util
import json
import os
import re
import sys
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
CIRC = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
NUM = r"\d+"
PERIOD = re.compile(r"(영구(?:히)?(?!\s*(?:기록|적|성|물관리|보존시설|보존기록물관리))|준영구|(?<!\d)(?:\d{1,3}\s*(?:년|개월)|\d{1,4}\s*일)(?!\s*\d+\s*월)(?:\s*(?:간|동안|이상|이내))?)")  # '2019년'·'1월 1일' 같은 날짜는 뺌
LAPSE = re.compile(r"(\d+\s*(?:년|개월|일))\s*(?:(?:이|가)\s*(?:지나면|지난|경과하면|경과한)[^.]{0,40}?(?:파기|삭제|폐기)|(?:을|를)\s*초과하여\s*(?:보유|보존|보관)할\s*수\s*없)")
KEEPV = re.compile(r"(?:파기|삭제)하여야|(보존|보관|비치|보유|유지|저장)(?:\s*[ㆍ·]\s*(?:관리|유지|보관|보존|활용))?(?:하여야|해야|하여야만|하고|한다|할 수|할 것|하도록|하여 두어야|해 두어야|하며|하게|기간)|보존기간")
EVENT = re.compile(r"(?:을|를)\s*((?:그\s)?[^,.()을를]{2,50}?(?:때|날|시점)\s*까지)\s*(?:보존|보관)")
DELEG = re.compile(r"(대통령령|총리령|[가-힣]{1,12}부령|[가-힣]{1,12}(?:부|처|청|위원회)\s*(?:장관|장|위원장)?(?:이|가)\s*정하여\s*고시하는|[가-힣]{1,12}(?:부|처|청|위원회)의?\s*고시|조례|규칙)\s*(?:으로|로)?\s*정하는")
RECW = re.compile(r"기록|서류|장부|자료|대장|문서|정보|명부|일지|영상|파일|사본|증빙|결과표|회의록|내역|데이터|원부|카드|전표|계약서|신청서|보고서|대장|부본|필름|녹음|녹화|로그|설계도")
DELEG_GOSI = re.compile(r"([가-힣]{1,12}(?:부|처|청|위원회))\s*(?:장관|장|위원장)?(?:이|가)\s*(?:정하여\s*)?고시하는")
EXTEND = re.compile(r"다만[^.]{0,40}?(대통령령|총리령|[가-힣]{1,12}부령)\s*(?:으로|로)\s*정하는\s*바에\s*따라[^.]{0,30}?(?:보존기간|기간)을?\s*(?:연장|늘릴|달리)")


def _builder():
    sp = importlib.util.spec_from_file_location("statutory_catalog_build", os.path.join(HERE, "statutory_catalog_build.py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


ART = re.compile(r"(?<![가-힣])(제\d+(?:-\d+)?조(?:의\d+)?)\s*\(([^)]{1,60})\)")


def ref_prefix(src):
    """아래 법령이 위 법령을 부르는 말: 법률·조례 → '법'·'조례', 시행령(○○령) → '영', 시행규칙 → '규칙'."""
    if not src or src.endswith("조례") or not re.search(r"(?:령|규칙|규정)$", src):
        return r"(?:법|조례)"
    if src.endswith("규칙"):
        return r"(?:규칙|시행규칙)"
    return r"(?:영|시행령)"


def articles(text):
    """본문 → [(조, 제목, 조문 본문)] (부칙 앞까지)."""
    t = re.sub(r"\s+", " ", text)
    t = re.sub(r"(\d)\s*연간", r"\1년간", t)
    first = re.search(r"(?<![가-힣])제\s*1(?:-1)?\s*조(?:의\d+)?\s*\(", t)
    if first:
        t = t[first.start():]
    cut = re.search(r"부\s*칙\s*(?:\([^)]{0,80}\)\s*)?<[^>]{0,40}제\s*[\d-]+호|부\s*칙\s*<", t)
    if cut:
        t = t[:cut.start()]
    t = re.sub(r"&lt;[^&]{0,60}&gt;|<[^>]{0,60}>", " ", t)  # <개정 2019. 1. 15.> 같은 꼬리표
    # '제3조(정의)에 따라'·'제5조(…)의 규정'은 조문을 가리키는 말이라 경계가 아님 — 다만 '제20조(기록의 보존) 의료기관의 장은'처럼 본문이 '의'로 시작하는 것은 경계
    ms = [m for m in ART.finditer(t) if not re.match(r"\s*(?:에서|에|의|부터|까지|및|또는|·)(?![가-힣])", t[m.end():m.end() + 8])]
    out = []
    for i, m in enumerate(ms):
        body = t[m.end(): ms[i + 1].start() if i + 1 < len(ms) else len(t)].strip()
        out.append((m.group(1), m.group(2), body))
    return out


def hangs(body):
    """조문 본문 → [(항 번호 '①'…, 본문)]; 항이 없으면 [('', 본문)]."""
    idx = [i for i, c in enumerate(body) if c in CIRC]
    # 항 번호는 차례대로 나와야 함(본문 속 '제1항' 인용과 구별)
    seq, want = [], 0
    for i in idx:
        if CIRC.index(body[i]) == want:
            seq.append(i)
            want += 1
    if not seq:
        return [("", body)]
    out = [("", body[:seq[0]].strip())] if body[:seq[0]].strip() else []
    for k, i in enumerate(seq):
        out.append((body[i], body[i + 1: seq[k + 1] if k + 1 < len(seq) else len(body)].strip()))
    return out


def hos(text):
    """항 본문 → (머리 문장, [(호 번호, 호 본문)]); 호는 1. 2. 3. 이 차례대로 나와야 함."""
    cands = [m for m in re.finditer(r"(?:(?<=\s)|^)(\d{1,2})\.\s", text)]
    seq, want = [], 1
    for m in cands:
        if int(m.group(1)) == want:
            seq.append(m)
            want += 1
    if not seq:
        return text, []
    head = text[:seq[0].start()].strip()
    items = []
    for k, m in enumerate(seq):
        items.append((m.group(1), text[m.end(): seq[k + 1].start() if k + 1 < len(seq) else len(text)].strip()))
    return head, items


def norm_period(p):
    p = re.sub(r"\s+", "", p)
    p = re.sub(r"(간|동안)$", "", p)
    return p.replace("영구히", "영구")


def overrides(head):
    """'3년(제2호의 경우 2년을 말한다)', '다만, 제3호 및 제4호의 서류는 5년' → {호: 기간}."""
    out = {}
    for m in re.finditer(r"\(([^()]{0,80})\)|다만,?([^.]{0,200})", head):
        seg = m.group(1) or m.group(2) or ""
        per = PERIOD.search(seg)
        if not per:
            continue
        nums = re.findall(r"제\s*(\d+)\s*호", seg[:per.start()])
        rng = re.search(r"제\s*(\d+)\s*호부터\s*제\s*(\d+)\s*호까지", seg[:per.start()])
        if rng:
            nums += [str(i) for i in range(int(rng.group(1)), int(rng.group(2)) + 1)]
        for n in nums:
            out[n] = norm_period(per.group(1))
    return out


def clauses(text):
    """'…3년 동안 보존하여야 하고, …2년 동안 보존하여야 하며, …' → 보존 의무별 마디."""
    parts = re.split(r"(?:(?<=하여야\s하고)|(?<=하여야\s하며)|(?<=해야\s하고)|(?<=해야\s하며)|(?<=보존하고)|(?<=보관하고)|(?<=보존하며)|(?<=보관하며)),?\s", text)
    return [p.strip() for p in parts if p.strip()]


def main_period(clause):
    """괄호 속 예외를 뺀 마디의 대표 기간(보존 동사 바로 앞의 기간을 우선)."""
    bare = re.sub(r"\([^()]{0,80}\)", " ", clause)
    bare = re.split(r"다만,?", bare)[0]
    # '보존기간이 30년 이상인 기록물'처럼 보존기간 구분을 가리키는 말은 의무 기간이 아님
    bare = re.sub(r"보존\s*기간(?:이|은)?\s*(?:영구|준영구|\d+\s*년)\s*(?:이상|이하|미만|초과)?\s*(?:인|으로\s*정하여진|으로\s*분류된|에\s*해당하는)", " ", bare)
    lp = LAPSE.search(bare)
    if lp:  # 'N년이 지나면 파기' — 그보다 오래 두면 안 되는 상한
        return norm_period(lp.group(1)) + "(지나면 파기)"
    cands = []
    for pm in PERIOD.finditer(bare):
        tail = bare[pm.end(): pm.end() + 25]
        if re.search(r"(보존|보관|비치|보유|유지|저장|관리하여야|관리해야)", tail):
            cands.append(pm)
    if not cands:
        for m in re.finditer(r"(?:보존|보관)\s*(?:\([^)]*\))?\s*기간\s*(?:은|을|는|이)?\s*(?:각각\s*)?", bare):
            pm = PERIOD.search(bare, m.end(), m.end() + 12)
            if pm:
                cands.append(pm)
    if cands:
        return norm_period(cands[0].group(1))
    dm = re.search(r"기간[”\"’']?\s*이?란[^.]{0,100}?(" + PERIOD.pattern + r")", bare)  # '“해양수산부령으로 정하는 기간”이란 … 2년까지를 말한다'
    if dm:
        return norm_period(dm.group(1))
    ev = EVENT.search(bare)
    if ev:
        return ev.group(1).strip()
    return ""


SUBJ = re.compile(r"(?:^|\s)(?:[가-힣ㆍ·]{0,30}?(?:사업주|자|기관|소유주등|업자|지도사|사람|장|처리자|개설자|의료인|사용자|공무원|위원회|법인|단체|사업자|학교|센터|본부|청|부|시설|대표|운영자|관리자|설치자|보험자|약사|공단|협회|조합|회사|은행))(?:은|는|이|가)\s")


def subject_object(clause):
    """마디에서 '…를/을 N년 동안 보존' 또는 '…는 보존기간을 N년' 앞의 대상(기록) 문구."""
    bare = re.sub(r"\s+", " ", re.sub(r"\([^()]{0,120}\)", " ", clause))
    if not bare.lstrip().startswith("다만"):
        bare = re.split(r"\s다만,?\s", bare)[0]  # 본문만 — 단서는 따로 다룸
    bare = re.sub(r"^(?:법|영|규칙)?\s*제\s*\d+조[^,]{0,60}?(?:에\s*따라|에서|에\s*따른)\s*", "", bare)
    m = re.search(r"([^.]{0,220}?)\s*(?:은|는)\s*(?:그\s*)?(?:보존|보관)\s*기간", bare) \
        or re.search(r"([^.]{0,220}?)\s*(?:을|를)\s*(?:" + PERIOD.pattern + r"|(?:그\s)?[^,.을를]{0,50}?(?:때|날)\s*까지)", bare) \
        or re.search(r"([^.]{0,220}?)\s*(?:은|는)\s*(?:" + PERIOD.pattern + r")\s*(?:간|동안)?\s*(?:보존|보관)", bare) \
        or re.search(r"([^.]{0,220}?)\s*(?:을|를|은|는)\s*[^.]{0,40}?" + LAPSE.pattern, bare)
    obj = m.group(1) if m else bare[:120]
    sub = SUBJ.search(obj)  # 첫 주어(의무자) 뒤부터가 기록 — '사업주는 … 근로자가 제출한 서류'
    if sub:
        obj = obj[sub.end():]
    obj = re.sub(r"^(?:다만,?\s*|그\s+)", "", obj.strip(" ,"))
    if sub and re.fullmatch(r"(?:정보|서류|자료|기록|장부|문서)", obj):  # '보조금관리정보를 보유한 자는 그 정보를' → 보조금관리정보
        hm = re.match(r"\s*([^\s]+?)(?:을|를)\s", bare[: sub.end()])
        if hm:
            obj = hm.group(1)
    return obj.strip()


def deleg_of(text):
    m = DELEG.search(text)
    if not m:
        return ""
    return re.sub(r"\s+", "", m.group(1))


def _hist():
    sp = importlib.util.spec_from_file_location("statutory_history", os.path.join(HERE, "statutory_history.py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


class Corpus:
    """현행 법령·행정규칙 원문 모음(이름 → 본문). 위임 대상 찾기에 씀. live=True면 없는 법령·고시를 국가법령정보센터에서 받아 옴."""

    def __init__(self, laws, admins=None, live=False, cache=None, date=None, live_laws=True):
        self.laws = laws          # {이름(시행일 꼬리 뺀): 본문}
        self.admins = admins or {}
        self._arts = {}
        self.live = live
        self.cache = cache
        self.date = date or __import__("time").strftime("%Y%m%d")
        self._miss = set()
        self._gq = {}
        self.live_laws = live_laws

    def _cached(self, key, fn):
        if self.cache:
            os.makedirs(self.cache, exist_ok=True)
            p = os.path.join(self.cache, re.sub(r"[^\w가-힣.-]", "_", key) + ".txt")
            if os.path.exists(p):
                return open(p, encoding="utf-8").read()
            t = fn()
            if t:
                open(p, "w", encoding="utf-8").write(t)
            return t
        return fn()

    def law(self, name):
        if name in self.laws or not self.live or not self.live_laws or name in self._miss:
            return self.laws.get(name, "")
        h = _hist()
        if re.search(r"(조례|규칙)(?:\s*시행규칙)?$", name) and re.match(r"\S+(?:특별시|광역시|특별자치시|도|특별자치도)\b", name):
            t = self._cached("자치법규_" + name, lambda: self._fetch_ordin(name))
            if t:
                self.laws[name] = t
            else:
                self._miss.add(name)
            return t

        def fetch():
            seq, lsid = h.find_current(name)
            if not seq:
                return ""
            vs = [v for v in h.versions(seq, lsid) if v["시행"] <= self.date]
            return h.body(vs[-1]) if vs else ""
        t = self._cached("법령_" + name + "_" + self.date, fetch)
        if t:
            self.laws[name] = t
        else:
            self._miss.add(name)
        return t

    def _fetch_ordin(self, name):
        """자치법규 이름으로 찾아 원문(ordinInfoR)을 받음 — 조례가 '규칙으로 정하는' 시행규칙을 따라갈 때."""
        b = _builder()
        q = " ".join(name.split()[1:]) if len(name.split()) > 2 else name
        h = b.get(f"{b.BASE}ordinSc.do?menuId=3&subMenuId=27&tabMenuId=139&query={urllib.parse.quote(q)}")
        key = re.sub(r"\s", "", name)
        for m in re.finditer(r"ordinInfPop\('(\d+)'[^>]*>\s*([^<]{2,120})", h):
            if re.sub(r"\s", "", m.group(2)) == key:
                t = b.get(f"{b.BASE}ordinInfoR.do?ordinSeq={m.group(1)}")
                return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
        return ""

    def purpose_index(self):
        """'이 영은 「소방공무원법」에서 위임된 사항…'처럼 제1조(목적)에서 위임 법률을 밝힌 하위 법령 색인(이름이 '○○법 시행령'이 아닌 것: 임용령·복무규정 등)."""
        if not hasattr(self, "_pidx"):
            self._pidx = {}
            for n, t in self.laws.items():
                m = re.search(r"제\s*1\s*조\s*\(목적\)(.{0,400})", t[:60000])
                if not m:
                    continue
                for mm in list(re.finditer(r"「([^」]{2,60})」\s*(?:\([^)]{0,40}\))?\s*(?:제\s*\d+조[^.「」]{0,400}?)?(?:에서|에\s*따라|에\s*따른|의\s*시행|및\s*같은\s*법)", m.group(1)))[:1]:  # 목적의 첫 법률만 — 그 법령의 '법'은 이것
                    x = mm.group(1)
                    # 이름이 닮은 것만(소방공무원법 → 소방공무원 임용령) — 목적에 다른 법을 곁들여 든 것은 뺌
                    a_, b_ = re.sub(r"\s|법$", "", x), re.sub(r"\s", "", n)
                    lm = __import__("difflib").SequenceMatcher(None, a_, b_).find_longest_match(0, len(a_), 0, len(b_))
                    if x != n and lm.size >= 3:
                        self._pidx.setdefault(x, []).append(n)
        return self._pidx

    def source_of(self, phrase, prefer=()):
        """행정규칙·자치법규가 법률 문구를 그대로 옮겨 적은 경우('안전ㆍ보건상의 조치 사항으로서 고용노동부령으로 정하는 사항을 적은 서류')
        그 문구가 있는 법률 조문(법령, 조, 항, 호)을 찾음 — 위임은 그 법률에서 따라간다."""
        key = re.sub(r"[\s「」ㆍ·,]", "", phrase or "")
        if len(key) < (16 if prefer else 20):
            return None
        if not hasattr(self, "_norm"):
            self._norm = {n: re.sub(r"[\s「」ㆍ·,]", "", t) for n, t in self.laws.items() if not re.search(r"(조례|규칙)$", n) or n.endswith("시행규칙")}
        order = [n for n in self._norm if strip_name(n) in prefer or any(n.startswith(p + " ") for p in prefer)] if prefer else []
        for n in order + [n for n in self._norm if n not in order and len(key) >= 20]:
            t = self._norm[n]
            if key in t:
                for jo, title, body in self.arts(n):
                    for hn, ht in hangs(body):
                        if key in re.sub(r"[\s「」ㆍ·,]", "", ht):
                            head, items = hos(ht)
                            ho = next((no for no, it in items if key in re.sub(r"[\s「」ㆍ·,]", "", it)), "")
                            return n, jo, hn, ho
        return None

    def source_like(self, text, per, prefer):
        """옛 문구로 옮겨 적어 글자가 그대로 맞지 않는 경우('고시하는 발암성 확인물질' ↔ 현행 '고시하는 물질'):
        그 규정이 근거로 든 법률 집안(법·시행령·시행규칙)에서 '고시하는'과 같은 보존기간이 함께 있는 항 가운데 낱말이 가장 많이 겹치는 것."""
        if not prefer or not per:
            return None
        words = set(re.findall(r"[가-힣]{2,}", text)) - {"고용노동부장관이", "정하여", "고시하는", "가운데", "서류", "보존하여야", "한다"}
        pn = re.sub(r"\s", "", per)
        best, score = None, 1
        for n in self._norm:
            if not (strip_name(n) in prefer or any(n.startswith(p + " ") for p in prefer)):
                continue
            for jo, title, body in self.arts(n):
                if "고시하는" not in body:
                    continue
                for hn, ht in hangs(body):
                    if "고시하는" in ht and pn in re.sub(r"\s", "", ht):
                        sc = len(words & set(re.findall(r"[가-힣]{2,}", ht)))
                        if sc > score:
                            best, score = (n, jo, hn, ""), sc
        return best

    def arts(self, name):
        if name not in self._arts:
            src = self.laws.get(name) or self.admins.get(name) or self.law(name) or ""
            self._arts[name] = articles(src) if src else []
        return self._arts[name]

    def base(self, name):
        return re.sub(r"\s*(시행령|시행규칙)$", "", name)

    def subs(self, name, kind):
        """'산업안전보건법' + '고용노동부령' → '산업안전보건법 시행규칙' 등 위임받는 법령 이름 후보."""
        b = self.base(name)
        if kind in ("조례", "규칙") or name.endswith("조례"):  # 'OO시 △△ 조례' → 'OO시 △△ 조례 시행규칙'
            return [c for c in (f"{name} 시행규칙", re.sub(r"\s*조례$", " 조례 시행규칙", name)) if c in self.laws]
        if kind == "대통령령":
            cand = [f"{b} 시행령"]
        elif kind.endswith("부령") or kind == "총리령":
            cand = [f"{b} 시행규칙", f"{b} 시행령"]
        else:
            cand = [f"{b} 시행령", f"{b} 시행규칙"]
        out = [c for c in cand if c != name and (c in self.laws or self.law(c))]
        for c in self.purpose_index().get(b, []):  # 소방공무원법 → 소방공무원 임용령·소방공무원 징계령 …
            is_decree = bool(re.search(r"(?:령|규정)$", c)) and not c.endswith("규칙")
            if c not in out and ((kind == "대통령령" and is_decree) or (kind != "대통령령" and (c.endswith("규칙") or not is_decree)) or kind == "규칙"):
                out.append(c)
        return out

    def find_refs(self, sub, jo, hang="", ho="", proviso=False, src=""):
        """sub 법령에서 '법 제164조제1항제3호'·'법 제164조제1항 단서'를 가리키는 조문·항.
        src(위임한 법령)가 시행령이면 시행규칙은 '영 제N조', 시행규칙이면 '규칙 제N조'로 가리킨다."""
        j = jo.replace("제", "제\\s*").replace("조", "\\s*조")
        pat = ref_prefix(src) + r"\s*" + j + r"(?!의)"  # 시행령·시행규칙은 '법 제N조', 조례 시행규칙은 '조례 제N조'로 가리킴
        if hang:
            pat += r"\s*제\s*" + str(CIRC.index(hang) + 1) + r"\s*항"
        if ho:
            pat += r"\s*제\s*" + ho + r"\s*호"
        elif proviso:
            pat += r"\s*(?:각\s*호\s*외의\s*부분\s*)?단서"
        rx = re.compile(pat)
        out = []
        for ajo, title, body in self.arts(sub):
            for hn, ht in hangs(body):
                if rx.search(ht):
                    out.append((ajo, title, hn, ht))
        return out

    def live_gosi(self, org, law, jo, hang):
        """국가법령정보센터 행정규칙 조문 검색으로 후보를 찾아 원문을 받아 둠."""
        if not self.live:
            return
        b = _builder()
        short = "규칙" if law.endswith("시행규칙") else ("영" if law.endswith("시행령") else "법")
        hn = f"제{CIRC.index(hang) + 1}항" if hang else ""
        qs = [f"{short} {jo}{hn}", f"{law.split()[-1]} {jo}{hn}", f"{jo}{hn}"]
        for q in qs:
            if q in self._gq:
                continue
            try:  # 첫 쪽(150건)만 — 넓은 검색어로 수십 쪽을 넘기지 않음
                _get = b.get
                pages = {"n": 0}

                def one_page(u, _g=_get):
                    pages["n"] += 1
                    return _g(u) if pages["n"] == 1 else ""
                b.get = one_page
                try:
                    res = b.search_rules(q)
                finally:
                    b.get = _get
            except Exception:  # noqa: BLE001
                res = []
            self._gq[q] = res
            for r in res:
                nm = r.get("규칙", "")
                kind = re.search(r"\[시행[^\]]*\]\s*\[([^\]\s]+)", nm)  # [고용노동부고시 제2020-44호 …]
                if not kind or (org and org not in kind.group(1)):
                    continue
                seq = r.get("admSeq")
                name = strip_name(nm)
                if not seq or name in self.admins:
                    continue
                t = self._cached("행정규칙_" + seq, lambda: re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", b.get(f"{b.BASE}admRulLsInfoR.do?admRulSeq={seq}"))))
                self.admins[name] = t
            if any(org in n for n in self.admins):
                break

    def find_gosi(self, org, refs, hang=""):
        """'고용노동부장관이 정하여 고시하는 물질' → 그 법령 조문(refs: ['산업안전보건법 시행규칙 제241조', …])을 근거로 든 행정규칙.
        고시는 법령 이름을 줄여 씀('「산업안전보건법 시행규칙」(이하 "규칙"이라 한다)') → 줄임말도 봄."""
        for r in refs:
            law, jo = r.rsplit(" ", 1)
            self.live_gosi(org, law, jo, hang)
        hits = []
        for name, body in self.admins.items():
            if org and org not in name and org not in body[:600]:
                continue
            for r in refs:
                law, jo = r.rsplit(" ", 1)
                aliases = [law] + [m.group(2) for m in re.finditer(r"「([^」]+)」\s*\(이하\s*[\"“]([^\"”]{1,10})[\"”]", body) if m.group(1).replace(" ", "") == law.replace(" ", "")]
                jp = jo.replace("제", "제\\s*").replace("조", "\\s*조")
                hp = (r"\s*제\s*" + str(CIRC.index(hang) + 1) + r"\s*항") if hang else ""
                rx = re.compile(r"(?:" + "|".join(re.escape(x).replace("\\ ", "\\s*") for x in aliases) + r")[」\s]*(?:[^.「」]{0,80}?[,·ㆍ및]\s*)?" + jp + hp + r"(?!의)")  # '규칙 제186조제1항제2호, …, 제241조제1항 단서' 
                for ajo, title, abody in articles(body):
                    if rx.search(abody) and re.search(r"고시하는", abody):
                        hits.append((name, ajo, title, abody[:600]))
            
            if org and org[:2] not in name and org not in body[:400]:
                continue
        return hits


def analyze(corpus, law, jo, title, body, depth=0, path=None, seen=None, hang=None):
    """한 조문 → 항목 행 목록(위임은 끝까지 따라감). hang: 위임받은 한 항만 줄 때 그 항 번호."""
    path = path or [f"{law} {jo}"]
    seen = seen or set()
    rows = []
    for hn, htext in ([(hang, body)] if hang is not None else hangs(body)):
        if not (KEEPV.search(htext) or EVENT.search(htext)):
            continue
        head, items = hos(htext)
        ov = overrides(head)
        for cl in clauses(head if items else htext):
            if not (KEEPV.search(cl) or EVENT.search(cl)):
                continue
            strong = re.search(r"(?:" + PERIOD.pattern + r")\s*(?:간|동안)?\s*(?:보존|보관)", cl)
            if not strong and not RECW.search(cl) and not (items and any(RECW.search(it) for _, it in items)):
                continue  # '사내유보금을 보유' 같은 기록 아닌 '보유·유지'
            per = main_period(cl)
            obj = subject_object(cl)
            here = path + [f"{hn}" if hn else ""]
            pro = re.split(r"\s다만,?\s", " " + cl, maxsplit=1)
            pro_txt = pro[1] if len(pro) > 1 else ""
            permissive = re.search(r"초과하여|(?:보유|보존|보관)할\s*수\s*있다|연장할\s*수", pro_txt)
            pro_per = main_period(pro_txt) if pro_txt and not EXTEND.search("다만, " + pro_txt) and not permissive else ""
            if items and re.search(r"다음\s*각\s*호", cl):
                hidden = False
                if not per and not deleg_of(cl) and depth < 3 and HIDDEN_KEEP.search(cl) and not EVENT.search(cl) and not re.search(r"(?:령|규칙|규정)$", law) and not any(item_period(it) for _, it in items):
                    hid = implied(corpus, law, jo, hn, obj, depth, path + [f"{hn}"], seen)  # '다음 각 호의 사항을 기록 및 보존' — 기간은 하위 법령에
                    if hid:
                        per, hidden = hid[0]["보존기간"], True
                        rows += hid
                if not per and not deleg_of(cl) and not any(item_period(it) or deleg_of(it) for _, it in items):
                    continue  # 기간도 위임도 없는 '다음 각 호의 장부를 비치' 따위
                in_pro = bool(pro_txt and re.search(r"다음\s*각\s*호", pro_txt) and pro_per)
                if in_pro:  # '…1년 이상 보관. 다만, 다음 각 호의 경우에는 2년 이상' — 본문 1행 + 각 호는 단서 기간
                    obj = subject_object(cl)
                    rows.append(row(law, jo, title, hn, "", obj, per, cl, path + [f"{hn}"], "", "확정"))
                for n, it in items:
                    p = ov.get(n, pro_per if in_pro else per)
                    k0 = len(rows)
                    rows += item_rows(corpus, law, jo, title, hn, n, it, p, cl, depth, path, seen)
                    if hidden and len(rows) > k0:
                        rows[k0]["조건"] = "법률은 기간을 정하지 않음 — 하위 법령이 정함"
                        if rows[k0]["상태"] == "확정":
                            rows[k0]["상태"] = "위임 추적"
                    if in_pro:
                        rows[-1]["조건"] = "단서: " + it[:120]
                        rows[-1]["기록(항목)"] = f"{obj} — 해당하는 경우: {it[:150]}"
            else:
                d = deleg_of(obj)
                if not per and not d:
                    if depth < 3 and HIDDEN_KEEP.search(cl) and RECW.search(cl) and not EVENT.search(cl) and not re.search(r"(?:령|규칙|규정)$", law):
                        hid = implied(corpus, law, jo, hn, obj, depth, path, seen)
                        if hid:
                            rows.append(dict(row(law, jo, title, hn, "", obj, hid[0]["보존기간"], cl, path, "", "위임 추적"), 조건="법률은 기간을 정하지 않음 — 하위 법령이 정함"))
                            rows += hid
                    continue
                rows.append(row(law, jo, title, hn, "", obj, per, cl, path, d, "확정" if per and not d else ("사건 기한" if per and not PERIOD.search(per) else "")))
                if d and depth < 4:
                    rows += follow(corpus, law, jo, hn, "", d, per, depth, path, seen, rows[-1])
                if pro_per and pro_per != per and not DELEG_GOSI.search(pro_txt):  # 단서가 다른 기간(고시 단서는 아래에서 따로)
                    rows.append(dict(row(law, jo, title, hn, "", obj, pro_per, cl, path + ["단서"], "", "확정"), 조건="단서: " + pro_txt[:160]))
        exc = re.search(r"다만,?[^.]{0,80}?(대통령령|총리령|[가-힣]{1,12}부령)\s*(?:으로|로)\s*정하는\s*([^.]{0,60}?)(?:은|는)\s*([^.]{0,80}?)(?:보유|보존|보관)할\s*수\s*있다", htext)
        if exc and depth < 4:  # '다만, 대통령령으로 정하는 자료는 5년을 초과하여 … 보유할 수 있다' → 그 자료와 예외 기간
            rows += follow(corpus, law, jo, hn, "", re.sub(r"\s+", "", exc.group(1)), f"예외 — {exc.group(3).strip()}", depth, path + ["단서"], seen, None, proviso=True)
        ext = EXTEND.search(htext)
        if ext and depth < 4:
            rows += follow(corpus, law, jo, hn, "", re.sub(r"\s+", "", ext.group(1)), "", depth, path, seen, None, proviso=True)
        for g in DELEG_GOSI.finditer(htext):
            if depth < 4:
                st = htext.rfind("다만", 0, g.start())
                st = st if st >= 0 else max(htext.rfind(". ", 0, g.start()) + 1, 0)
                en = htext.find(". ", g.end())
                cl = htext[st: en if en > 0 else len(htext)]
                # 보존기간을 정하는 문장 안의 '고시하는 ○○'만 따라감('고시하는 금액·지정 기준'은 뺌)
                if PERIOD.search(cl) and KEEPV.search(cl) and RECW.search(cl):
                    main = re.split(r"다만,?", htext)[0]
                    rows += follow_gosi(corpus, law, jo, hn, g.group(1), cl, depth, path, seen, title, parent_obj=subject_object(main))
    # 기간이 없는 행은 버림(보존 의무와 무관한 위임 조문을 따라간 것) — 위임을 못 찾은 보존 조문만 남겨 사람 확인
    rows = [r for r in rows if r["보존기간"] or r["상태"] in ("위임 못 찾음", "방법만 위임(기간 없음)")]
    for r in rows:
        if "참조 조문" not in r:
            r["참조 조문"] = cross_refs(corpus, r["법령"], r["기록(항목)"])
    return rows


def cross_refs(corpus, law, text):
    """기록 문구 속 '제24조제3항 및 제75조제4항에 따른 회의록'의 조문 제목을 찾아 붙임(산업안전보건위원회·협의체).
    '법 제N조'는 위 법률, 「다른 법」 제N조는 건너뜀."""
    out = []
    for m in re.finditer(r"(「[^」]{1,40}」\s*)?(법\s*)?(?<![가-힣])제\s*(\d+)\s*조(?:\s*의\s*(\d+))?", text or ""):
        if m.group(1):
            continue
        target = corpus.base(law) if (m.group(2) and law != corpus.base(law)) else law
        jo = f"제{m.group(3)}조" + (f"의{m.group(4)}" if m.group(4) else "")
        t = next((ti for j, ti, _ in corpus.arts(target) if j == jo), "")
        if t:
            out.append(f"{'법 ' if target != law else ''}{jo}({t})")
    return " · ".join(dict.fromkeys(out))


DUTY = re.compile(r"((?:[^\s,]+\s){0,4}?[^\s,]*(?:사업주|자|기관|소유주등|업자|지도사|사람|장|처리자|개설자|의료인|사용자|공무원|위원회|법인|단체|사업자|학교|센터|본부|공단|협회|조합|회사|은행|운영자|관리자|설치자|보험자|약사))(?:은|는|이|가)\s")


def duty_holder(sent):
    bare = re.sub(r"\([^()]{0,120}\)", " ", sent or "")
    bare = re.sub(r"^(?:법|영|규칙)?\s*제\s*\d+조[^,]{0,60}?(?:에\s*따라|에서|에\s*따른)\s*", "", bare)
    m = DUTY.search(bare)
    if not m:
        return ""
    h = re.sub(r"^.*?(?:에\s*따라|라\s*한다\)|,)\s*", "", m.group(1))
    return h.strip()


def row(law, jo, title, hn, ho, rec, per, sent, path, deleg, status):
    return {"법령": law, "조문": f"{jo}({title})", "항": hn, "호": ho, "기록(항목)": rec[:200], "보존기간": per,
            "의무자": duty_holder(sent), "조건": "", "근거 경로": " → ".join(p for p in path if p), "위임": deleg,
            "상태": status or ("위임 추적" if deleg else "확인 필요"), "문장": sent[:300]}


def item_period(it):
    """호 안에 제 기간이 있으면('보관물 관리부: 10년', '…의 경우: 5년')."""
    m = re.search(r"[:：]\s*(?:([^:：.]{1,20}?(?:부터|이후|후))\s*)?(" + PERIOD.pattern + r")" + NOTCYCLE, it)  # '원본: 발급일부터 5년'
    if m:
        return norm_period(m.group(2)) + (f"({m.group(1).strip()})" if m.group(1) else "")
    m = re.search(r"(" + PERIOD.pattern + r")\s*(?:간|동안)?\s*(?:보존|보관)", it)
    return norm_period(m.group(1)) if m else ""


NOTCYCLE = r"(?!\s*(?:마다|에\s*(?:1|한)\s*(?:회|번)|(?:이|가)\s*(?:지난|경과)(?!\s*(?:한\s*)?(?:날|때)\s*까지)|주기|이상\s*경과|이내에))"  # '2년마다 점검'·'5년이 지난'은 보존기간 아님


MOK = "가나다라마바사아자차카타파하"


def moks(it):
    """호 본문 → (호 제목, [(목, 목 본문)]); 목은 가. 나. 다. 가 차례대로 나와야 함('1. 영업에 관한 자료 가. 투자권유 관련 자료: 10년 나. …')."""
    cands = list(re.finditer(r"(?:(?<=\s)|^)([" + MOK + r"])\.\s", it))
    seq, k = [], 0
    for m in cands:
        if k < len(MOK) and m.group(1) == MOK[k]:
            seq.append(m)
            k += 1
    if len(seq) < 2:
        return it, []
    return it[:seq[0].start()].strip(), [(m.group(1), it[m.end(): seq[i + 1].start() if i + 1 < len(seq) else len(it)].strip()) for i, m in enumerate(seq)]


def item_rows(corpus, law, jo, title, hn, n, it, per, head, depth, path, seen):
    ttl, mk = moks(it)
    if mk and any(item_period(x) for _, x in mk):  # 목마다 기간 → 목별 행
        out = []
        for g, x in mk:
            out += item_rows(corpus, law, jo, title, hn, f"{n}{g}", (f"{ttl} — " if ttl else "") + x, per, head, depth, path, seen)
        return out
    own = item_period(it)
    per = own or per
    if own:  # '환자 명부 : 5년' → 기록은 '환자 명부'
        it = re.sub(r"\s*[:：]\s*(?:" + PERIOD.pattern + r").*$", "", it).strip() or it
    d = deleg_of(it)
    st = "사건 기한" if per and not PERIOD.search(per) else ("확정" if per and not d else "")
    r = row(law, jo, title, hn, n, it, per, head, path + [f"{hn}{n}호"], d, st)
    out = [r]
    if d and depth < 4:
        out += follow(corpus, law, jo, hn, n, d, per, depth, path + [f"{hn}{n}호"], seen, r)
    return out


HIDDEN_KEEP = re.compile(r"(?:보존|보관)(?:\s*[ㆍ·]\s*(?:관리|유지))?\s*(?:하여야|해야)|(?:기록|작성)\s*및\s*보존")


def implied(corpus, law, jo, hn, obj, depth, path, seen):
    """법률은 '…을 기록ㆍ보존하여야 한다'고만 하고 기간도 위임 문구도 없는데, 시행령·시행규칙이 '법 제N조제M항에 따른 … 기록을 3년간 보존'처럼
    기간을 정한 경우(숨은 위임). 같은 조의 다른 항이 '제1항에 따른 …'으로 이어받아 기간을 정한 것도 찾는다."""
    out, b = [], corpus.base(law)
    subs_ = []
    for k in ("대통령령", "총리령"):
        for x in corpus.subs(law, k):
            if x not in subs_ and x != law:
                subs_.append(x)
    for sub in subs_:
        refs = corpus.find_refs(sub, jo, hn, src=law) or (corpus.find_refs(sub, jo, src=law) if hn else [])
        for ajo, title, ahn, atext in refs:
            cands = [(ahn, atext)]
            body = next((bb for j, t, bb in corpus.arts(sub) if j == ajo), "")
            if ahn:  # 같은 조에서 '제1항에 따른 …'으로 이 항을 이어받는 항
                k = CIRC.index(ahn) + 1
                cands += [(h2, t2) for h2, t2 in hangs(body) if h2 != ahn and re.search(r"(?<![조\d])제\s*" + str(k) + r"\s*항(?:\s*각\s*호)?에\s*따(?:른|라)", t2)]
            for h2, t2 in cands:
                if not (PERIOD.search(t2) and KEEPV.search(t2)):
                    continue
                key = (sub, ajo, h2, "implied")
                if key in seen:
                    continue
                seen.add(key)
                got = analyze(corpus, sub, ajo, title, t2, depth + 1, path + [f"{sub} {ajo}{h2}"], seen, hang=h2)
                for r in got:
                    if r["보존기간"]:
                        r["조건"] = (r.get("조건", "") + " · " if r.get("조건") else "") + "법률은 기간을 정하지 않음 — 하위 법령이 정함"
                out += [r for r in got if r["보존기간"]]
    return out


def follow(corpus, law, jo, hn, ho, kind, per, depth, path, seen, parent, proviso=False):
    """위임 따라가기: 아래 법령에서 '법 {jo}{항}{호|단서}'를 가리키는 조문을 찾아 분석."""
    out = []
    subs = corpus.subs(law, kind)
    found = False
    for sub in subs:
        if found and (kind.endswith("부령") or kind == "총리령") and not sub.endswith("규칙"):
            continue  # 부령 위임은 시행규칙에서 찾았으면 시행령까지 가지 않음(다른 문장의 대통령령 위임을 가로채지 않게)
        refs = corpus.find_refs(sub, jo, hn, ho, proviso, src=law)
        if not refs and hn:  # 항이 하나뿐인 조문은 '법 제30조'처럼 항 없이 가리키기도 함
            refs = corpus.find_refs(sub, jo, "", ho, proviso, src=law)
        whole = False
        if not refs and ho:  # 호를 따로 가리키지 않고 항 단위로 받은 경우('법 제23조의2제1항에 따라 … 정보는 별표 3의8과 같다')
            j = jo.replace("제", "제\\s*").replace("조", "\\s*조")
            refs = [x for x in corpus.find_refs(sub, jo, hn, "", proviso, src=law)
                    if not re.search(ref_prefix(law) + r"\s*" + j + r"(?!의)\s*(?:제\s*\d+\s*항\s*)?(?:제\s*\d+\s*호|각\s*호\s*외의\s*부분)", x[3])
                    and re.search(r"기록|보존|보관|대장|장부|서식|별표|사항", x[3])]
            whole = bool(refs)
        for ajo, title, ahn, atext in refs:
            key = (sub, ajo, ahn)
            if key in seen:
                found = found or proviso  # 단서(보존기간 연장)가 가리키는 조문을 다른 길(각 호·본문)로 이미 따라감
                continue
            seen.add(key)
            found = True
            sub_rows = analyze(corpus, sub, ajo, title, atext, depth + 1, path + [f"{sub} {ajo}{ahn}"], seen, hang=ahn)
            if not sub_rows:  # 위임받은 조문이 기간 없이 '…란 다음 각 호의 서류를 말한다'만 정함 → 위 기간을 그대로
                head, items = hos(atext)
                for n, it in (items or [("", re.sub(r"^.*?(?:란|이란)\s*", "", atext)[:200])]):
                    ip = item_period(it) if items else ""  # '"○○령으로 정하는 기간"이란 … 1. 원본: 발급일부터 5년'
                    out.append(row(sub, ajo, title, ahn, n, it, ip or per, atext, path + [f"{sub} {ajo}{ahn}{n + '호' if n else ''}"], "",
                                   "위임 추적" if (ip or per) else "확인 필요"))
                    if whole:
                        out[-1]["조건"] = f"{ho}호를 따로 가리키지 않고 항 전체로 정함 — 별표·서식 확인"
                    out[-1]["의무자"] = parent["의무자"] if parent else ""
                    out[-1]["_via"] = "inherit"
            else:
                out += sub_rows
    copied = law in corpus.admins or bool(re.match(r"\S+(?:특별시|광역시|특별자치시|도|특별자치도)\s.*(?:조례|규칙)$", law))
    if not found and parent is not None and copied and not subs_from_source(parent):  # 법률을 옮겨 적은 행정규칙·자치법규만
        ph = parent.get("기록(항목)", "")
        dm = DELEG.search(ph)
        own = corpus.admins.get(law) or corpus.laws.get(law) or ""
        prefer = set(re.findall(r"「([^」]{2,60})」", own))  # 그 규정이 근거로 든 법률을 먼저 — 옛 문구('안전ㆍ보건상의 조치 사항')는 위임 문구부터만 맞춤
        src = None
        for back in ((8, 5, 2, 0) if dm else (0,)):
            src = corpus.source_of(ph[max(0, dm.start() - back):] if dm else ph, prefer)
            if src:
                break
        if src and src[0] != law:  # 법률 문구를 옮겨 적은 행정규칙·자치법규 → 그 법률에서 위임을 따라감
            s_law, s_jo, s_hn, s_ho = src
            sub_out = follow(corpus, s_law, s_jo, s_hn, s_ho, kind, per, depth + 1,
                             path + [f"(옮겨 적은 법률) {s_law} {s_jo}{s_hn}{s_ho + '호' if s_ho else ''}"], seen, None)
            if sub_out:
                out += sub_out
                found = True
                parent["_src"] = 1
    if not found and parent is not None and not parent.get("보존기간"):
        # 아래 법령이 '법 제N조'를 가리키지 않고 기록 이름으로만 정한 경우('인사기록') — 위임 추정으로 붙이고 표시
        nm = re.findall(r"[가-힣]{1,8}(?:기록|장부|서류|대장|명부|일지)", parent.get("기록(항목)", ""))
        noun = max(nm, key=len) if nm else ""
        if len(noun) >= 3:
            for sub in subs:
                for ajo, title, abody in corpus.arts(sub):
                    if noun in abody and KEEPV.search(abody) and PERIOD.search(abody):
                        guess = [r for r in analyze(corpus, sub, ajo, title, abody, depth + 1, path + [f"{sub} {ajo}(추정)"], seen) if noun in r["기록(항목)"] + r["문장"]]
                        for r in guess:
                            r["상태"] = "위임 추정"
                        out += guess
                        found = found or bool(guess)
    if parent is not None:
        manner = re.search(r"정하는\s*바에\s*따라", parent.get("기록(항목)", ""))
        rec = parent.get("기록(항목)", "")
        # '규칙으로 정하는 서식에 따라 장부를 비치'·'대법원규칙으로 정하는 장소에 보관'·'○○령으로 정하는 바에 따라 기록' — 기간이 아니라 서식·장소·방법만 맡김
        form_only = not parent["보존기간"] and re.search(r"정하는\s*(?:서식|양식|장소|관리대장|신청서|신고서|요청서|입찰에\s*관한\s*서류|소득공제신청서|지원\s*신청서)|정하는\s*바에\s*따라\s*(?:[^.]{0,40}?)(?:작성|기록|비치|복사|발급|첨부|유지|관리)", rec) \
            and not re.search(r"정하는\s*기간|보존기간", rec)
        parent["상태"] = ("위임 추정" if any(r.get("상태") == "위임 추정" for r in out) else "위임 추적") if found else ("확정(방법만 위임)" if manner and parent["보존기간"] else ("방법만 위임(기간 없음)" if form_only else "위임 못 찾음"))
        if found and not parent["보존기간"]:  # '○○령으로 정하는 기간 동안 보존' → 아래 법령이 정한 기간을 위로
            pp = next((r["보존기간"] for r in out if r["보존기간"]), "")
            parent["보존기간"] = pp
    elif not found and proviso:
        out.append(row(law, jo, "", hn, "", "보존기간 연장(단서)", "", "", path + ["단서"], kind, "위임 못 찾음"))
    return out


def subs_from_source(parent):
    return parent.get("_src")


def org_head(org):
    """'고용노동부' → '고용노동부장관이', '식품의약품안전처' → '식품의약품안전처장이', '금융위원회' → '금융위원회가'."""
    if org.endswith("부"):
        return org + "장관이"
    if org.endswith(("처", "청")):
        return org + "장이"
    return org + "가"


def follow_gosi(corpus, law, jo, hn, org, text, depth, path, seen, title="", parent_obj=""):
    refs = [f"{law} {jo}"]
    hits = corpus.find_gosi(org, refs, hn)
    seg_jo = jo
    if not hits:  # 행정규칙이 법령 문구를 옮겨 적은 경우 → 그 법령 조문을 근거로 든 고시
        g = re.search(r"[가-힣]{1,12}(?:부|처|청|위원회)\s*(?:장관|장|위원장)?(?:이|가)\s*(?:정하여\s*)?고시하는[^.]{0,24}", text)
        src = corpus.source_of(g.group(0)) if g else None
        if not src and g:
            own = corpus.admins.get(law) or corpus.laws.get(law) or ""
            pm = re.search(r"고시하는[^.]{0,80}?" + PERIOD.pattern, text)
            src = corpus.source_like(text, norm_period(pm.group(1)) if pm else "", set(re.findall(r"「([^」]{2,60})」", own)))
        if src and src[0] != law:
            hits = corpus.find_gosi(org, [f"{src[0]} {src[1]}"], src[2])
            if hits:
                seg_jo = src[1]  # 고시 안에서 찾을 조문 번호는 옮겨 적은 법령의 것(규칙 제241조)
                path = path + [f"(옮겨 적은 법령) {src[0]} {src[1]}{src[2]}"]
    out = []
    m = re.search(r"(?:정하여\s*)?고시하는\s*([^.,]{1,30}?)(?:에|을|를|이|의|은|는)", text)
    what = m.group(1) if m else "대상"
    per_m = re.search(r"고시하는[^.]{0,80}?" + PERIOD.pattern, text)
    per = norm_period(per_m.group(1)) if per_m else ""
    if not hits:
        out.append(dict(row(law, jo, title, hn, "", (f"{parent_obj} 가운데 " if parent_obj else "") + f"{org_head(org)} 정하여 고시하는 {what}", per, text,
                            path + [f"{org} 고시(못 찾음)"], "고시", "위임 못 찾음"), 조건="단서: " + re.sub(r"^다만,?\s*", "", text)[:160]))
    for name, ajo, title, abody in hits[:3]:
        key = (name, ajo)
        if key in seen:
            continue
        seen.add(key)
        # 고시 조문에서 그 법령 조문을 가리키는 호·목만 떼어 '고시하는 물질' 목록으로
        seg = abody
        jm = re.search(seg_jo.replace("제", "제\\s*").replace("조", "\\s*조"), abody)
        if jm:
            seg = abody[jm.start():]
        head, items = hos(seg)
        if not items:  # '3. 규칙 제241조제2항의 … 물질이란 다음 각 목' 처럼 목(가. 나.)으로 나열
            items = [(m.group(1), m.group(2).strip()) for m in re.finditer(r"(?:(?<=\s)|^)([가나다라마바사아자차카타파하])\.\s(.{2,150}?)(?=\s[가나다라마바사아자차카타파하]\.\s|\s\d{1,2}\.\s|\s제\s*\d+\s*조|$)", seg[:900])]
        base = subject_object(text) or f"{org_head(org)} 정하여 고시하는 {what}"
        if parent_obj and parent_obj not in base:  # '작업환경측정 결과를 기록한 서류' 가운데 '고시하는 물질에 대한 기록이 포함된 서류'
            base = f"{parent_obj} 가운데 {base}"
        if items:
            for n, it in items[:30]:
                if re.search(r"고시하는|이란|란\s", it) and len(it) > 120:
                    continue
                it = re.sub(r"\s제\s*\d+\s*조(?:의\d+)?\s*$", "", re.split(r"\s제\s*\d+\s*조(?:의\d+)?\s*(?:\(|<|&lt;)", it)[0])
                out.append(dict(row(name, ajo, title, "", n, f"{base} — 고시 대상: {it[:150]}", per, seg[:300], path + [f"「{name}」 {ajo} {n}"], "고시", "위임 추적"), 조건="단서: " + re.sub(r"^다만,?\s*", "", text)[:120] + " — 고시 대상: " + it[:80], _via="gosi"))
        else:
            out.append(dict(row(name, ajo, title, "", "", f"{base} — 「{name}」 {ajo}({title})", per, abody, path + [f"「{name}」 {ajo}"], "고시", "위임 추적"), 조건="단서: " + re.sub(r"^다만,?\s*", "", text)[:160], _via="gosi"))
    return out


def add_judgement(rows):
    """보존 의무 판정기(retention_judge.py, 외부 API 없음)로 '판정' 칸을 매김. 위임받은 사항 나열은 문장만으로 판정하기 어려워 '확인 필요' 이상."""
    sp = importlib.util.spec_from_file_location("retention_judge", os.path.join(HERE, "retention_judge.py"))
    jd = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(jd)
    for r in rows:
        r["판정"] = jd.judge(r.get("문장") or r["기록(항목)"], "상한" if "지나면 파기" in r["보존기간"] else "")[1]
        if r["상태"] == "위임 추적" and r["판정"] == "의무 아님 가능성":
            r["판정"] = "확인 필요"
    return rows


def strip_name(n):
    return re.sub(r"\s*\[시행[^\]]*\].*$", "", n).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--bodies", action="append", help="법령 원문 폴더(여러 번 줄 수 있음)")
    ap.add_argument("--names", action="append", help="--bodies와 짝인 이름표(json: 일련번호 → 이름)")
    ap.add_argument("--ord", help="자치법규 원문 폴더")
    ap.add_argument("--ord-names")
    ap.add_argument("--ord-orgs", help="자치법규 일련번호 → 지자체(json)")
    ap.add_argument("--adm")
    ap.add_argument("--adm-names")
    ap.add_argument("--law")
    ap.add_argument("--jo")
    ap.add_argument("--only", help="이 파일(한 줄에 하나)에 적힌 법령만 분석 — 원문 묶음은 그대로 둔 채 일부만 다시 돌릴 때")
    ap.add_argument("-o", "--out")
    ap.add_argument("--live", action="store_true", help="없는 법령·고시를 국가법령정보센터에서 받아 옴(--law만 주면 자동)")
    ap.add_argument("--cache", default="산출/위임추적원문", help="받은 원문을 두는 폴더")
    a = ap.parse_args()
    laws, admins, origin, orgs = {}, {}, {}, {}
    for bd, nmf in zip(a.bodies or [], a.names or []):
        names = json.load(open(nmf, encoding="utf-8"))
        for seq, nm in names.items():
            nm = nm["이름"] if isinstance(nm, dict) else nm
            fn = os.path.join(bd, f"{seq}.txt")
            if os.path.exists(fn) and strip_name(nm) not in laws:
                laws[strip_name(nm)] = open(fn, encoding="utf-8", errors="ignore").read()
                origin[strip_name(nm)] = "법령"
    if a.ord and a.ord_names:
        names = json.load(open(a.ord_names, encoding="utf-8"))
        og = json.load(open(a.ord_orgs, encoding="utf-8")) if a.ord_orgs else {}
        for seq, nm in names.items():
            fn = os.path.join(a.ord, f"{seq}.txt")
            if os.path.exists(fn):
                laws[strip_name(nm)] = open(fn, encoding="utf-8", errors="ignore").read()
                origin[strip_name(nm)] = "자치법규"
                orgs[strip_name(nm)] = og.get(seq, "")
    if a.adm and a.adm_names:
        names = json.load(open(a.adm_names, encoding="utf-8"))
        for seq, nm in names.items():
            fn = os.path.join(a.adm, f"{seq}.txt")
            if os.path.exists(fn):
                admins[strip_name(nm)] = open(fn, encoding="utf-8", errors="ignore").read()
                origin[strip_name(nm)] = "행정규칙"
    corpus = Corpus(laws, admins, live=a.live or not a.bodies, cache=a.cache)
    rows = []
    targets = [a.law] if a.law else list(laws) + list(admins)
    if a.only:
        want = {strip_name(x.strip()) for x in open(a.only, encoding="utf-8") if x.strip()}
        targets = [t for t in targets if t in want]
    if a.law:
        corpus.law(a.law)
    for k, law in enumerate(targets):
        if not a.law and k % 200 == 0:
            print(f"  {k}/{len(targets)} 법령 …", flush=True)
        for jo, title, body in corpus.arts(law):
            if a.jo and jo != a.jo:
                continue
            if not (KEEPV.search(body) or EVENT.search(body)):
                continue
            rows += analyze(corpus, law, jo, title, body)
    seen_k, uniq = set(), []
    for r in rows:  # 같은 행이 여러 길(제 조문·다른 법령의 위임)로 나오면 하나만
        k = (r["법령"], r["조문"], r["항"], r["호"], r["기록(항목)"], r["보존기간"], r.get("조건", ""))
        if k not in seen_k:
            seen_k.add(k)
            uniq.append(r)
    rows = add_judgement([r for r in uniq if len(re.sub(r"[\s\d.\[\]]", "", r["기록(항목)"])) >= 2])
    for r in rows:
        r["출처"] = origin.get(r["법령"], "행정규칙" if r.get("위임") == "고시" else "법령")
        r["기관"] = orgs.get(r["법령"], "")
    cols = ["출처", "기관", "법령", "조문", "항", "호", "기록(항목)", "참조 조문", "보존기간", "의무자", "조건", "근거 경로", "위임", "상태", "판정", "문장"]
    if a.out:
        with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)
        import collections
        print(f"{len(rows)}행 → {a.out}: " + " · ".join(f"{k} {v}" for k, v in collections.Counter(r["상태"] for r in rows).most_common()))
    else:
        for r in rows:
            print(f"{r['법령']} {r['조문']} {r['항']}{r['호'] + '호' if r['호'] else ''}\t{r['보존기간'] or '—'}\t[{r['상태']}]\t{r['기록(항목)'][:70]}\t⟵ {r['근거 경로']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
