#!/usr/bin/env python3
"""위임 추적 목록(statutory_items.csv)에서 '별표 N'을 가리키는 행에 그 별표의 제목·요지·원문 주소를 붙인다.

예: 농약관리법 시행규칙 제24조의4 ① '기록ㆍ보존해야 하는 판매ㆍ구매 정보는 별표 3의8과 같다'
    → [별표 3의8] 제조업자 등의 기록ㆍ보존 대상 정보 — 1. 제조업자 가. 구매자 이름ㆍ주소ㆍ연락처 …

어디의 별표인지: '「산업안전보건기준에 관한 규칙」 별표 12' → 그 법령, '영 별표 1' → 시행령, '규칙 별표' → 시행규칙,
'안전보건규칙 별표 12'처럼 줄인 이름은 목록 문장 속 「…」(이하 "안전보건규칙")에서, 아무 말 없으면 그 행의 법령.

받는 길(국가법령정보센터 화면, 외부 API·비용 없음):
  법령 lsBylInfoPLinkR.do → lsBylInfoR.do(별표 목록) → lsBylContentsInfoR.do(PDF 파일 번호) → flDownload.do(PDF)
  행정규칙·자치법규는 admRulByl…·ordinByl… 같은 꼴. PDF는 pdftotext로 글자만 뽑는다(받은 파일은 실행하지 않음).

사용:
  python statutory_byl.py                      # 붙일 행과 별표 제목만 보여 줌(받기는 함, 고치지 않음)
  python statutory_byl.py --apply              # '별표' 칸을 채워 목록을 고침
  python statutory_byl.py --law 농약관리법 --apply
"""
from __future__ import annotations

import argparse
import csv
import html
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ITEMS = os.path.join(HERE, "..", "assets", "law", "statutory_items.csv")
BASE = "https://www.law.go.kr/LSW/"
REF = re.compile(r"(?:「([^」]{2,60})」\s*|(?<![가-힣])(영|규칙|법|시행령|시행규칙)\s+|([가-힣]{2,12}(?:규칙|규정|고시|기준))\s+)?별표\s*(\d+(?:의\d+)?)(?!\s*의\s*\d)")
KINDS = {  # 종류 → (목록 화면, 목록 조각, 본문 조각, 일련번호 이름)
    "법령": ("lsBylInfoPLinkR.do", "lsBylInfoR.do", "lsBylContentsInfoR.do", "lsiSeq"),
    "행정규칙": ("admRulBylInfoPLinkR.do", "admRulBylInfoR.do", "admRulBylContentsInfoR.do", "admRulSeq"),
    "자치법규": ("ordinBylInfoPLinkR.do", "ordinBylInfoR.do", "ordinBylContentsInfoR.do", "ordinSeq"),
}
DELAY = 0.4


def _mod(name):
    sp = importlib.util.spec_from_file_location(name, os.path.join(HERE, name + ".py"))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


def http(path, data=None, raw=False):
    err = None
    for i in range(4):
        try:
            req = urllib.request.Request(BASE + path, data=urllib.parse.urlencode(data).encode() if data is not None else None,
                                         headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=90) as r:
                b = r.read()
            time.sleep(DELAY)
            return b if raw else b.decode("utf-8", "ignore")
        except Exception as e:  # noqa: BLE001
            err = e
            time.sleep(3 + 3 * i)
    raise err


def text(h):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", h))).strip()


def aliases(rows):
    """목록 문장 속 「산업안전보건기준에 관한 규칙」(이하 "안전보건규칙"이라 한다) → {'안전보건규칙': '산업안전보건기준에 관한 규칙'}."""
    out = {}
    for r in rows:
        for m in re.finditer(r"「([^」]{2,60})」\s*\(\s*이하\s*[“\"]([^”\"]{2,20})[”\"]", r.get("문장", "") + " " + r.get("기록(항목)", "")):
            out.setdefault(m.group(2), m.group(1))
    return out


SHORT = {"안전보건규칙": "산업안전보건기준에 관한 규칙"}  # 고시·행정규칙이 정의 없이 쓰는 줄인 이름(법령 원문의 「…」(이하 "…") 그대로)


def target(r, al):
    """그 행이 가리키는 (법령 이름, 별표 번호) 목록."""
    law = r["법령"]
    base = re.sub(r"\s*(?:시행령|시행규칙)$", "", law)
    out = []
    for m in REF.finditer(r["기록(항목)"] + " " + r.get("조건", "")):
        full, short, named, no = m.groups()
        if full:
            name = full
        elif short in ("영", "시행령"):
            name = f"{base} 시행령"
        elif short in ("규칙", "시행규칙") and re.search(r"(?:법|시행령|령)$", law):  # 법률·시행령이 '규칙 별표'라 하면 시행규칙(규칙·규정 자신이면 그 규칙)
            name = f"{base} 시행규칙"
        elif short == "법":
            name = base
        elif named and named in al:
            name = al[named]
        elif named and named not in law and named in SHORT:
            name = SHORT[named]
        else:
            name = law
        if (name, no) not in out:
            out.append((name, no))
    return out


def hwp_text(d, stem):
    """받은 한글 파일 바이트 → 글자(실행하지 않고 글자만). HWP 5는 hwp5txt가 있을 때만."""
    import zipfile
    import io
    if d[:2] == b"PK":
        try:
            with zipfile.ZipFile(io.BytesIO(d)) as z:
                return "\n".join(" ".join(re.findall(r"<hp:t[^>]*>([^<]*)</hp:t>", z.read(n).decode("utf-8", "ignore")))
                                 for n in sorted(z.namelist()) if re.match(r"Contents/section\d+\.xml", n))
        except Exception:  # noqa: BLE001
            return ""
    if d[:4] == b"\xd0\xcf\x11\xe0" and shutil.which("hwp5txt"):
        p = stem + ".hwp"
        open(p, "wb").write(d)
        try:
            r = subprocess.run(["hwp5txt", p], timeout=180, capture_output=True)
            return r.stdout.decode("utf-8", "ignore")
        except Exception:  # noqa: BLE001
            return ""
        finally:
            os.remove(p)
    return ""


class Byl:
    def __init__(self, cache):
        self.cache = cache
        os.makedirs(cache, exist_ok=True)
        self.idx_path = os.path.join(cache, "목록.json")
        self.idx = json.load(open(self.idx_path, encoding="utf-8")) if os.path.exists(self.idx_path) else {}
        self.hist = None

    def save(self):
        json.dump(self.idx, open(self.idx_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    def seq(self, name, kind_hint=""):
        """이름 → (종류, 일련번호)."""
        k = f"seq|{name}"
        if k in self.idx:
            return tuple(self.idx[k]) if self.idx[k] else (None, None)
        self.hist = self.hist or _mod("statutory_history")
        got = (None, None)
        order = ["자치법규"] if re.search(r"(?:조례|조례 시행규칙)$", name) or kind_hint == "자치법규" else (["행정규칙", "법령"] if kind_hint == "행정규칙" else ["법령", "행정규칙"])
        for kind in order:
            try:
                if kind == "법령":
                    s, _ = self.hist.find_current(name)
                elif kind == "행정규칙":
                    s = self.hist.find_admin(name)
                else:
                    s = None
                    parts = name.split()
                    for q in dict.fromkeys([" ".join(parts[k:]) for k in range(1, min(len(parts), 4))] + [name]):
                        h = http(f"ordinSc.do?menuId=3&subMenuId=27&tabMenuId=139&outmax=150&query={urllib.parse.quote(q)}")
                        # 검색 결과는 '가평군 <span>자체감사 규칙</span>'처럼 지자체와 이름이 갈라져 있음 — 링크 글자 전체로 맞춤
                        s = next((m.group(1) for m in re.finditer(r"ordinInfPop\('(\d+)'[^>]*>(.{2,300}?)</a>", h, re.S)
                                  if re.sub(r"\s", "", text(m.group(2))) == re.sub(r"\s", "", name)), None)
                        if s:
                            break
            except Exception:  # noqa: BLE001
                s = None
            if s:
                got = (kind, s)
                break
        self.idx[k] = list(got) if got[0] else None
        return got

    def options(self, kind, seq):
        """별표 목록: [(옵션 값, 제목)]."""
        k = f"opt|{kind}|{seq}"
        if k in self.idx:
            return self.idx[k]
        page, lst, _, sk = KINDS[kind]
        p = http(f"{page}?{sk}={seq}")
        d = re.search(r'id="bylSeq"[^>]*value="(\d*)"', p)
        opts = []
        if d and d.group(1):
            t = http(lst, {"bylSeq": d.group(1), sk: seq, "vSct": "", "efYd": ""})
            opts = [(v, text(ti)) for v, ti in re.findall(r'<option value="([^"]+)"[^>]*>([^<]{2,200})</option>', t)]
        self.idx[k] = opts
        return opts

    def pdf_text(self, kind, seq, opt):
        """그 별표의 PDF를 받아 글자로(없으면 '')."""
        v = opt.split(",")
        k = f"txt|{kind}|{v[0]}"
        if k in self.idx:
            f = self.idx[k]
            return (open(os.path.join(self.cache, f), encoding="utf-8", errors="ignore").read(), self.idx.get(f"fl|{kind}|{v[0]}", "")) if f else ("", "")
        _, _, cont, sk = KINDS[kind]
        prm = {"bylSeq": v[0], "bylNo": v[1] if len(v) > 1 else "", "bylBrNo": v[2] if len(v) > 2 else "", "bylClsCd": v[3] if len(v) > 3 else "", "vSct": ""}
        if kind == "법령":
            prm.update(lsiSeq=v[0], bylEfYd=v[4] if len(v) > 4 else "")
        else:
            prm[sk] = seq
        c = http(cont, prm)
        fl = re.search(r'id="pdfFlSeq"[^>]*value="(\d+)"', c)
        han = re.search(r'id="hanFlSeq"[^>]*value="(\d+)"', c)
        out = ""
        if fl and shutil.which("pdftotext"):
            d = http(f"flDownload.do?flSeq={fl.group(1)}", raw=True)
            if d.startswith(b"%PDF"):
                pdf = os.path.join(self.cache, f"{fl.group(1)}.pdf")
                txt = os.path.join(self.cache, f"{fl.group(1)}.txt")
                open(pdf, "wb").write(d)
                subprocess.run(["pdftotext", "-layout", pdf, txt], timeout=120, capture_output=True)
                os.remove(pdf)
                if os.path.exists(txt):
                    out = open(txt, encoding="utf-8", errors="ignore").read()
                    self.idx[k] = os.path.basename(txt)
        if not out and han:  # PDF가 없으면(자치법규 등) 한글 파일 — HWP는 hwp5txt, HWPX는 zip 안의 글자
            d = http(f"flDownload.do?flSeq={han.group(1)}", raw=True)
            out = hwp_text(d, os.path.join(self.cache, han.group(1)))
            if out:
                open(os.path.join(self.cache, f"{han.group(1)}.txt"), "w", encoding="utf-8").write(out)
                self.idx[k] = f"{han.group(1)}.txt"
            fl = fl or han
        if not out:
            self.idx[k] = ""
        self.idx[f"fl|{kind}|{v[0]}"] = fl.group(1) if fl else ""
        return out, (fl.group(1) if fl else "")

    def find(self, name, no, kind_hint=""):
        """그 별표 → {법령, 제목, 본문, 주소}; 법령을 못 찾으면 None, 별표가 등록돼 있지 않으면 {'없음': 사유}."""
        kind, seq = self.seq(name, kind_hint)
        if not seq:
            return None
        want = f"[별표 {no}]"
        opts = self.options(kind, seq)
        if opts and not any(ti.startswith("[별표") for _, ti in opts):
            return {"없음": "별표 등록 없음(별지 서식만)"}
        for v, ti in opts:
            if ti.replace(" ]", "]").startswith(want) or (no in ("", "1") and ti.startswith("[별표 ]")):
                body, fl = self.pdf_text(kind, seq, v)
                return {"법령": name, "제목": ti, "본문": body, "주소": f"{BASE}flDownload.do?flSeq={fl}" if fl else ""}
        return None


def marked(body, mark):
    """별표 속 표시가 붙은 항목 이름(조건까지): '7) 2-니트로톨루엔(…; 88-72-2)(특별관리물질)' → '2-니트로톨루엔',
    '24) 크롬[7440-47-3] 및 그 화합물(…)(6가크롬 화합물만 특별관리물질)' → '크롬 및 그 화합물(6가크롬 화합물만)',
    '10) 수은 …(특별관리물질. 다만, 아릴화합물 및 알킬화합물은 특별관리물질에서 제외한다)' → '수은 및 그 화합물(다만, 아릴화합물 및 알킬화합물은 제외)'.
    줄이 넘어가도 번호 단위로 읽고, '1)부터 123)까지의 물질을 … 함유한 혼합물'은 뺀다."""
    names = []
    body = re.sub(r"\s*".join(mark), mark, body)  # PDF 줄바꿈으로 '특별관리 물질'이 된 것
    for chunk in re.split(r"\n\s*(?=\d{1,3}\)\s)", "\n" + body):
        one = re.sub(r"\s+", " ", re.sub(r"(?<=[가-힣])[ \t]*\n\s*(?=[가-힣])", "", chunk)).strip()  # 줄바꿈으로 갈린 낱말('화합 물만')은 붙임
        if mark not in one or re.match(r"\d{1,3}\)\s*\d{1,3}\)\s*부터|\d{1,3}\)\s*부터", one) or "까지의 물질" in one[:40]:
            continue
        m = re.match(r"\d{1,3}\)\s*(.+)", one)
        if not m:
            continue
        rest = m.group(1)
        parens = [(x.start(), x.group(1)) for x in re.finditer(r"\(((?:[^()]|\([^()]*\))*)\)", rest)]
        hit = next(((st, x) for st, x in parens if mark in x), None)
        if not hit:
            continue
        head = rest[:hit[0]]
        head = re.sub(r"\[[\d\-, 등]+\]", "", head)  # '[7440-47-3]' CAS 번호
        head = re.split(r"\s*[(\[](?=[A-Za-z0-9][^)\]]{3,})", head)[0].strip()  # '(Benzene; 71-43-2)'·'[Benzo(a)pyrene; …]' 영문
        note = hit[1].strip()
        note = "" if note == mark else re.sub(r"\s*(?:은|는)?\s*" + mark + r"(?:에서)?\s*", " ", re.sub(r"^" + mark + r"\.\s*", "", note)).strip()
        note = re.sub(r"\s+(제외한다|제외)$", " 제외", note).strip(" .")
        n = head + (f"({note})" if note else "")
        if head and n not in names:
            names.append(n)
    return names


def summary(b, limit=600, want=""):
    """'[별표 3의8] 제조업자 등의 기록ㆍ보존 대상 정보(제24조의4제1항 관련) — 1. 제조업자 가. … (PDF 주소)'."""
    t = re.sub(r"\s+", " ", b["본문"])
    t = re.sub(r"^■?\s*" + re.escape(b["법령"]) + r"\s*\[별표[^\]]*\]\s*(?:<[^>]{0,40}>)?\s*", "", t)
    ttl = re.sub(r"^\[별표\s*[\d의 ]*\]\s*", "", b["제목"])
    t = t.replace(ttl, "", 1).strip() if ttl and ttl[:12] in t[:200] else t
    t = t[:limit] + ("…" if len(t) > limit else "")
    if want and "특별관리물질" in want:  # '관리대상 유해물질 중 특별관리물질' → 별표에서 (특별관리물질) 표시가 붙은 것만
        sp = marked(b["본문"], "특별관리물질")
        if sp:
            t = f"특별관리물질 {len(sp)}종: " + ", ".join(sp)
    return f"{b['제목']} — {t}" + (f" ({b['주소']})" if b["주소"] else "")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--items", default=ITEMS)
    ap.add_argument("--cache", default="산출/별표원문", help="받은 별표 글자·목록 캐시(받은 PDF는 글자만 남기고 지움)")
    ap.add_argument("--law", help="이 법령(또는 이 이름이 근거 경로에 있는) 행만")
    ap.add_argument("--apply", action="store_true", help="'별표' 칸을 채워 목록을 고침")
    ap.add_argument("--limit", type=int, default=600, help="요지 글자 수")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.items, encoding="utf-8-sig")))
    if not rows:
        print("목록이 비었음")
        return 1
    al = aliases(rows)
    byl = Byl(a.cache)
    done = miss = 0
    memo = {}
    for r in rows:
        if a.law and a.law not in r["법령"] and a.law not in r["근거 경로"]:
            continue
        tg = target(r, al)
        if not tg:
            continue
        got, why = [], []
        for name, no in tg:
            key = (name, no)
            if key not in memo:
                try:
                    memo[key] = byl.find(name, no, "행정규칙" if r.get("출처") == "행정규칙" and name == r["법령"] else ("자치법규" if r.get("출처") == "자치법규" and name == r["법령"] else ""))
                except Exception as e:  # noqa: BLE001
                    print(f"  ⚠ {name} 별표 {no}: {e}", file=sys.stderr)
                    memo[key] = None
            b = memo[key]
            if b and "없음" in b:
                why.append(f"{name} 별표 {no}: {b['없음']}")
            elif b:
                got.append(summary(b, a.limit, r["기록(항목)"] + " " + r.get("조건", "")))
        if got:
            done += 1
            r["별표"] = " ‖ ".join(got)
            print(f"✓ {r['법령']} {r['조문'].split('(')[0]} {r['항']}{r['호']} → {' · '.join(g.split(' — ')[0] for g in got)}")
        else:
            miss += 1
            print(f"· {r['법령']} {r['조문'].split('(')[0]} {r['항']}{r['호']} → " + ("; ".join(why) if why else f"별표 {', '.join(n + ' ' + no for n, no in tg)} 못 받음"))
    byl.save()
    print(f"별표 붙임 {done}행 · 못 받음 {miss}행")
    if a.apply:
        cols = list(rows[0].keys())
        if "별표" not in cols:
            cols.insert(cols.index("문장") if "문장" in cols else len(cols), "별표")
        with open(a.items, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                r.setdefault("별표", "")
                w.writerow(r)
        print(f"→ {a.items}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
