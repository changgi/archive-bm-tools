#!/usr/bin/env python3
"""근거 자료 변화 감시 — 스킬이 기대는 누리집(국가기록포털 표준·지침·매뉴얼·분류기준·업무안내 자료, 법제처 입법예고, 국가법령정보센터)을
주기적으로 열어 '지난번에 본 것'과 비교하고, 새 게시물·새 첨부·법령 시행일 변화를 보고서로 남긴다.

하는 일
  1. sources.json(같은 폴더의 ../references/source_watch_sources.json)의 각 주소를 내려받는다(첨부 파일은 받지 않음, 쪽 HTML만).
  2. 쪽마다 지문을 만든다: 첨부 파일 이름(up_filename), 게시물 번호(board_seq·ogLmPp 번호)와 제목, 법령 '[시행 …]'·'제N호' 표기, 본문 글자 해시.
  3. 기준선(../references/source_baseline.json)과 비교해 새로 생긴 것·사라진 것·본문만 바뀐 것을 나눈다.
  4. 보고서(Markdown)와 JSON을 쓰고, --update면 기준선을 새로 쓴다.
  종료 코드: 0 변화 없음 · 1 변화 있음 · 2 입력 오류 · 3 모두 못 받음(네트워크)

하지 않는 것: 첨부 파일 내려받기·스킬 자동 수정. 변화는 사람이(또는 Claude 세션이) 읽고 스킬에 반영한다 — 보고서의 '반영할 곳' 칸이 출발점.

예) python source_watch.py --report 산출/근거변화.md            # 비교만
    python source_watch.py --update --report 산출/근거변화.md   # 비교 후 기준선 갱신(저장소 작업 흐름용)
"""
import argparse
import datetime as dt
import hashlib
import html
import json
import os
import re
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, "..", "references")
KST = dt.timezone(dt.timedelta(hours=9))
UA = "Mozilla/5.0 (records-skills source_watch; +https://github.com/changgi/archive-bm-tools)"


def fetch(url, timeout=40, tries=3):
    import time
    for i in range(tries):
        try:
            return _fetch(url, timeout)
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def _fetch(url, timeout):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ko"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read(5_000_000)
        cs = r.headers.get_content_charset() or "utf-8"
    return raw.decode(cs, errors="replace")


def clean(t):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t))).strip()


def fingerprint(page, kind):
    """쪽에서 비교할 항목(키→이름)과 본문 해시."""
    items = {}
    for fn in re.findall(r"up_filename=([^\"'&<>]+)", page):
        name = html.unescape(urllib.request.unquote(fn.replace("+", " "))).strip()
        if name:
            items[f"첨부:{name}"] = name
    for hx in re.findall(r"up_filepath=([0-9A-Fa-f]{16,})(?![0-9A-Fa-f])(?![^\"'<>]*up_filename=[^\"'&<>]+)", page):
        try:  # 파일 이름이 경로(16진수, EUC-KR)에만 있는 첨부
            name = bytes.fromhex(hx).decode("euc-kr", "replace").replace("\\", "/").rsplit("/", 1)[-1].strip()
        except ValueError:
            continue
        if name:
            items.setdefault(f"첨부:{name}", name)
    for m in re.finditer(r"<a\b([^>]*)>(.*?)</a>", page, re.S | re.I):
        attrs, text = m.group(1), clean(m.group(2))
        title = re.search(r'title="([^"]*)"', attrs)
        tt = html.unescape(title.group(1)).strip() if title else ""
        label = (tt if len(tt) > 3 and not re.search(r"상세|이동|바로가기|다운로드|새창|열기", tt) else text or tt)[:200]
        for pat, tag in ((r"board_seq=(\d+)", "게시물"), (r"ogLmPp/(\d+)(?!/myOpn)", "입법예고"), (r"bg_no=(\d+)", "간행물"), (r"[?&](?:amp;)?seq=(\d{3,})", "게시물"),
                         (r"fn_\w*[Dd]etail\w*\(\s*'?(\d{3,})", "게시물")):
            g = re.search(pat, attrs)
            if g and label and "myOpn" not in attrs:
                items.setdefault(f"{tag}:{g.group(1)}", label)

    if kind == "KS":
        for code in set(re.findall(r"KS [A-Z] (?:ISO |IEC )?(?:TR )?\d{3,5}(?:-\d)?(?::\d{4})?", clean(page))):
            items[f"KS:{code}"] = code
    if kind == "법령":
        for s in sorted(set(re.findall(r"\[시행 [^\]]{6,40}\]", page)) | set(re.findall(r"제\d{2,5}호, \d{4}\. ?\d{1,2}\. ?\d{1,2}\.[^\]<]{0,12}", page))):
            items[f"법령:{s}"] = s
    body = page
    m = re.search(r'id="(contents|container)"', page)
    if m:
        body = page[m.start():]
    body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", body, flags=re.S | re.I)
    text = clean(body)
    text = re.sub(r"(조회수?|조회|방문자|접속자)\s*[:：]?\s*[\d,]+", "", text)
    text = re.sub(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?", "", text)
    # 순서 바뀜·조회수 같은 숫자만 바뀐 것은 변화로 보지 않음: 숫자만인 낱말을 빼고 낱말 집합으로 해시(같은 길이·다른 순서 오탐 방지)
    words = sorted(set(w for w in text.split() if not re.fullmatch(r"[\d,.:()/-]+", w)))
    return items, hashlib.sha256(" ".join(words).encode()).hexdigest(), len(text)


def main() -> int:
    ap = argparse.ArgumentParser(description="근거 자료 변화 감시")
    ap.add_argument("--sources", default=os.path.join(REF, "source_watch_sources.json"))
    ap.add_argument("--baseline", default=os.path.join(REF, "source_baseline.json"))
    ap.add_argument("--update", action="store_true", help="비교 뒤 기준선을 새로 씀")
    ap.add_argument("--report", default="산출/근거변화.md")
    ap.add_argument("--only", help="이름에 이 글자가 든 출처만")
    ap.add_argument("--offline", help="(시험용) 출처 이름→HTML 파일 경로 JSON — 내려받지 않고 파일을 읽음")
    a = ap.parse_args()
    try:
        srcs = json.load(open(a.sources, encoding="utf-8"))["출처"]
    except Exception as e:
        print(f"입력 오류: 출처 목록을 못 읽음 — {e}", file=sys.stderr)
        return 2
    base = json.load(open(a.baseline, encoding="utf-8")) if os.path.exists(a.baseline) else {"출처": {}}
    offline = json.load(open(a.offline, encoding="utf-8")) if a.offline else {}
    now = dt.datetime.now(KST)
    result, failed, changed = {}, [], []
    for s in srcs:
        if a.only and a.only not in s["이름"]:
            continue
        try:
            page = open(offline[s["이름"]], encoding="utf-8", errors="replace").read() if s["이름"] in offline else fetch(s["주소"])
        except Exception as e:
            failed.append((s, f"{type(e).__name__}: {str(e)[:120]}"))
            if s["이름"] in base["출처"]:
                result[s["이름"]] = base["출처"][s["이름"]]
            continue
        items, h, n = fingerprint(page, s.get("종류", "목록"))
        old = base["출처"].get(s["이름"])
        rec = {"주소": s["주소"], "항목": items, "본문해시": h, "본문길이": n, "확인": now.isoformat(timespec="seconds"),
               "바뀐날": old.get("바뀐날", "") if old else now.isoformat(timespec="seconds")}
        if old:
            new = {k: v for k, v in items.items() if k not in old["항목"]}
            gone = {k: v for k, v in old["항목"].items() if k not in items}
            body = old["본문해시"] != h
            if new or gone or (body and s.get("본문비교", True)):
                rec["바뀐날"] = now.isoformat(timespec="seconds")
                changed.append((s, new, gone, body and not new and not gone, old.get("본문길이", 0), n))
        else:
            changed.append((s, {}, {}, False, 0, n))
        result[s["이름"]] = rec
    lines = [f"# 근거 자료 변화 보고 — {now:%Y-%m-%d %H:%M} KST", "",
             f"출처 {len(srcs)} · 변화 {len(changed)} · 못 받음 {len(failed)} · 기준선 {'갱신함' if a.update else '그대로'}", ""]
    if changed:
        lines += ["| 출처 | 새로 생김 | 사라짐 | 본문만 | 반영할 곳 |", "|---|---|---|---|---|"]
        for s, new, gone, body, n0, n1 in changed:
            first = s["이름"] not in base["출처"]
            lines.append(f"| [{s['이름']}]({s['주소']}) | {'(처음 기록)' if first else len(new)} | {'' if first else len(gone)} | "
                         f"{'예(' + str(n0) + '→' + str(n1) + '자)' if body else ''} | {s.get('반영', '')} |")
        for s, new, gone, body, *_ in changed:
            if s["이름"] not in base["출처"] or not (new or gone):
                continue
            lines += ["", f"## {s['이름']}", ""]
            lines += [f"- 🆕 {v}" for v in list(new.values())[:60]] + ([f"- …외 {len(new) - 60}건"] if len(new) > 60 else [])
            lines += [f"- ➖ {v}" for v in list(gone.values())[:30]]
        lines += ["", "다음 할 일: 🆕 항목의 게시물·첨부를 열어 읽고, '반영할 곳' 스킬의 스크립트·references를 고친 뒤 시험을 돌린다. "
                  "법령은 국가법령정보센터 원문(별표·별지 포함)으로 확인한다. 비밀 관련 자료도 공개 게시물이면 읽되, 실제 비밀기록물은 열지 않는다."]
    else:
        lines.append("변화 없음.")
    if failed:
        lines += ["", "## 못 받은 출처(변화로 보지 않음 — 다음 실행에서 다시 봄)", ""] + [f"- {s['이름']}: {e}" for s, e in failed]
    os.makedirs(os.path.dirname(a.report) or ".", exist_ok=True)
    open(a.report, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    json.dump({"시각": now.isoformat(timespec="seconds"), "변화": [{"출처": s["이름"], "새로": list(n.values()), "사라짐": list(g.values()), "본문만": b}
                                                              for s, n, g, b, *_ in changed], "못받음": [s["이름"] for s, _ in failed]},
              open(os.path.splitext(a.report)[0] + ".json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if a.update:
        keep = {k: v for k, v in base["출처"].items() if k not in result and any(k == x["이름"] for x in srcs)}  # --only로 건너뛴 출처는 그대로
        json.dump({"설명": "source_watch.py 기준선 — 직접 고치지 말 것", "출처": {**keep, **result}}, open(a.baseline, "w", encoding="utf-8"), ensure_ascii=False, indent=1, sort_keys=True)
    real = [c for c in changed if c[0]["이름"] in base["출처"]]
    print(f"출처 {len(srcs)} · 변화 {len(real)} · 처음 기록 {len(changed) - len(real)} · 못 받음 {len(failed)} → {a.report}")
    if failed and len(failed) == len(srcs):
        return 3
    return 1 if real else 0


if __name__ == "__main__":
    sys.exit(main())
