#!/usr/bin/env python3
"""웹기록물 수집 준비 점검: 기관 누리집이 국가기록원 웹 수집 로봇(원격 수집)에 잘 수집·재현될지 보고,
기관이 직접 만든 WARC 파일의 무결성(WARC-Block/Payload-Digest)과 SNS·블로그 채널 목록을 점검한다.

사용:  python web_readiness.py --url https://www.example.go.kr [--max-pages 40] [--delay 0.5] -o 산출/웹기록물_점검.xlsx
       python web_readiness.py --warc 수집본.warc.gz -o 산출/웹기록물_점검.xlsx
       python web_readiness.py --channels 채널목록.xlsx -o ...   (열: 채널명, 주소, 운영부서, 개설일, 종료일, 수집일)
       세 가지를 함께 써도 된다. 누리집 점검은 같은 호스트 안에서만, 정해진 쪽수만, 간격을 두고 읽는다(읽기만 함).
근거(2026-11-20 확인):
  공공기록물법 시행령 제2조 10호(웹기록물 정의: 웹사이트·블로그·SNS의 기록정보와 운영·구축 관리정보),
    제46조의2(2025.5.20. 신설: 영구기록물관리기관이 협의해 수집, 기술적 보호조치 해제 등 협조 요청, 중복·재현 불가는 심의 후 폐기)
  국가기록원 「2026년 기록물관리 지침」 웹기록물(시행령 제46조의2 반영):
    수집 범위 대표 URL 이하 포괄 수집, 주기 연 1회(폐지·한시기관은 폐지 약 3개월 전부터 1회 이상, 조직개편 시),
    보존형식 WARC(ISO 28500)·K-WARC, 기관 협조: 개인정보 등 일부 페이지만 로봇 배제, 누리집 전체 차단 금지
    (공공데이터법 제3조③), 수집 로봇 IP 예외처리, 사이트맵(누리집 지도) 페이지를 만들어 메인화면에 연결,
    동적 기술(자바스크립트 링크·마우스오버 메뉴·무한 스크롤·클릭 전환·자동 리다이렉션)은 수집·재현 불가 가능
원칙:  점검만 한다. robots.txt·누리집 수정, 수집 로봇 IP 예외 등록, 국가기록원 협의는 담당자가 한다.
"""
import argparse
import base64
import collections
import hashlib
import html.parser
import io
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.robotparser

UA = "records-skills-web-check/1.0 (+self-check before archival capture)"
CRAWLERS = ["heritrix", "archive.org_bot", "ia_archiver", "*"]


class Links(html.parser.HTMLParser):
    def __init__(self):
        super().__init__()
        self.links, self.js, self.onclick, self.meta_refresh, self.scripts, self.text = [], 0, 0, "", 0, 0
        self.anchors, self._a = [], None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if any(k.startswith("on") for k in a):
            self.onclick += 1
        if tag == "a":
            self._a = [a.get("href") or "", ""]
        if tag == "script":
            self.scripts += 1
        if tag == "meta" and (a.get("http-equiv") or "").lower() == "refresh":
            self.meta_refresh = a.get("content", "")
        for k in ("href", "src"):
            v = (a.get(k) or "").strip()
            if not v:
                continue
            if v.lower().startswith("javascript:") or (k == "href" and v in ("#", "#none", "#!")):
                self.js += 1
            elif not v.startswith(("mailto:", "tel:", "data:", "#")):
                self.links.append(v)

    def handle_endtag(self, tag):
        if tag == "a" and self._a:
            self.anchors.append(tuple(self._a))
            self._a = None

    def handle_data(self, d):
        self.text += len(d.strip())
        if self._a is not None:
            self._a[1] += d


def fetch(url, timeout=15):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.geturl(), r.headers.get("Content-Type", ""), r.read(3_000_000)
    except urllib.error.HTTPError as e:
        return e.code, url, "", b""
    except Exception as e:  # 연결 실패
        return 0, url, str(e)[:80], b""


def site_check(start, max_pages, delay):
    p = urllib.parse.urlsplit(start)
    root = f"{p.scheme}://{p.netloc}"
    findings, pages = [], []
    st, _, _, body = fetch(root + "/robots.txt")
    rtxt = body.decode("utf-8", "replace") if st == 200 else ""
    rp = urllib.robotparser.RobotFileParser()
    rp.parse(rtxt.splitlines())
    blocked = [c for c in CRAWLERS if not rp.can_fetch(c, start)]
    if blocked:
        findings.append(("robots.txt 전체 차단", "고칠 것", f"대표 URL이 {', '.join(blocked)}에 막힘",
                         "개인정보 등 일부 경로만 Disallow, 누리집 전체 차단 금지(공공데이터법 제3조③, 지침 8)"))
    else:
        dis = [l.split(":", 1)[1].strip() for l in rtxt.splitlines() if l.lower().startswith("disallow:") and l.split(":", 1)[1].strip()]
        findings.append(("robots.txt", "통과" if st == 200 else "확인", f"{'있음' if st == 200 else f'없음({st})'} · 제외 경로 {len(dis)}개" +
                         (f" ({', '.join(dis[:5])})" if dis else ""), "제외 경로가 비공개 페이지에 한정되는지 확인"))
    sm = [l.split(":", 1)[1].strip() for l in rtxt.splitlines() if l.lower().startswith("sitemap:")]
    st2, _, _, b2 = fetch(sm[0] if sm else root + "/sitemap.xml")
    xml_urls = re.findall(rb"<loc>\s*([^<\s]+)\s*</loc>", b2) if st2 == 200 else []
    seen, queue, sitemap_page = set(), [start], ""
    host_links = collections.Counter()
    while queue and len(pages) < max_pages:
        u = urllib.parse.urldefrag(queue.pop(0))[0]
        if u in seen:
            continue
        seen.add(u)
        if not rp.can_fetch(UA, u) and pages:
            pages.append([u, "robots 제외", "", "", "", "", ""])
            continue
        code, final, ctype, body = fetch(u)
        note = []
        if final.rstrip("/") != u.rstrip("/"):
            note.append(f"리다이렉트 → {final[:80]}")
        if "html" not in ctype.lower() or not body:
            pages.append([u, code, "", "", "", "", " · ".join(note) or ctype[:30]])
            time.sleep(delay)
            continue
        lp = Links()
        try:
            lp.feed(body.decode("utf-8", "replace"))
        except Exception:
            pass
        same = []
        for href in lp.links:
            v = urllib.parse.urljoin(final, href)
            q = urllib.parse.urlsplit(v)
            if q.netloc == p.netloc and q.scheme in ("http", "https"):
                same.append(v)
            elif q.netloc:
                host_links[q.netloc] += 1
        if lp.meta_refresh:
            note.append(f"meta refresh '{lp.meta_refresh[:40]}'")
        if lp.js >= 3:
            note.append(f"자바스크립트 링크 {lp.js}개")
        if len(same) < 3 and lp.scripts >= 3:
            note.append("정적 링크가 거의 없음 — 스크립트로 그리는 화면일 수 있음")
        if len(pages) == 0 and not sitemap_page:
            hit = next((h for h, t in lp.anchors if re.search(r"사이트\s*맵|누리집\s*지도|sitemap", t + " " + h, re.I) and not h.endswith(".xml")), "")
            sitemap_page = urllib.parse.urljoin(final, hit) if hit else ""
        pages.append([u, code, len(same), lp.js, lp.onclick, lp.scripts, " · ".join(note)])
        queue.extend(v for v in same if v not in seen and not re.search(r"\.(pdf|hwp|hwpx|zip|jpg|png|gif|mp4|xlsx?|docx?)$", v, re.I))
        time.sleep(delay)
    findings.append(("사이트맵", "통과" if (xml_urls or sitemap_page) else "고칠 것",
                     f"sitemap.xml {len(xml_urls)}개 URL" + (f" · 메인에 사이트맵 링크 {sitemap_page}" if sitemap_page else " · 메인화면에 사이트맵(누리집 지도) 링크 없음"),
                     "누리집 지도(대-중-소메뉴, 정확한 URL)를 만들어 메인화면에 연결(지침 8)"))
    dyn = [r for r in pages if isinstance(r[3], int) and (r[3] >= 3 or "스크립트로" in r[6])]
    findings.append(("동적 링크", "확인" if dyn else "통과", f"자바스크립트 링크·스크립트 화면 의심 {len(dyn)}/{len(pages)}쪽",
                     "메뉴·게시판 이동은 실제 주소(href)로, 클릭·마우스오버·스크롤만으로 나오는 콘텐츠는 수집·재현이 안 될 수 있음"))
    redir = [r for r in pages if "리다이렉트" in str(r[6]) or "meta refresh" in str(r[6])]
    if redir:
        findings.append(("자동 리다이렉션", "확인", f"{len(redir)}쪽", "자동 이동 주소는 재현 오류가 날 수 있음 — 최종 주소를 직접 링크"))
    bad = [r for r in pages if isinstance(r[1], int) and (r[1] == 0 or r[1] >= 400)]
    if bad:
        findings.append(("응답 오류", "확인", f"{len(bad)}쪽 (예: {bad[0][0][:60]} → {bad[0][1]})",
                         "연결 실패가 많으면 방화벽·DDoS 차단일 수 있음 — 수집 로봇 IP 예외처리 요청(지침 8)"))
    if host_links:
        findings.append(("외부 호스트", "참고", ", ".join(f"{h}({n})" for h, n in host_links.most_common(6)),
                         "외부 호스트 자원(동영상·지도·SNS)은 대표 URL 이하 수집 범위 밖일 수 있음"))
    return findings, pages


def warc_records(path):
    raw = open(path, "rb").read()
    if raw[:2] == b"\x1f\x8b":
        out, data = [], raw
        while data:
            d = __import__("zlib").decompressobj(31)
            out.append(d.decompress(data))
            data = d.unused_data
        raw = b"".join(out)
    f, recs = io.BytesIO(raw), []
    while True:
        line = f.readline()
        if not line:
            break
        if not line.startswith(b"WARC/"):
            continue
        hdr = {}
        while True:
            h = f.readline()
            if h in (b"\r\n", b"\n", b""):
                break
            k, _, v = h.decode("utf-8", "replace").partition(":")
            hdr[k.strip().lower()] = v.strip()
        n = int(hdr.get("content-length", "0") or 0)
        recs.append((line.strip().decode(), hdr, f.read(n)))
    return recs


def digest_ok(spec, data):
    algo, _, val = spec.partition(":")
    algo = algo.lower().replace("-", "")
    if algo not in ("sha1", "sha256", "md5"):
        return None
    h = hashlib.new(algo, data).digest()
    return val.upper() in (base64.b32encode(h).decode(), h.hex().upper(), base64.b64encode(h).decode().upper())


def warc_check(path):
    rows, cnt = [], collections.Counter()
    for ver, h, block in warc_records(path):
        t = h.get("warc-type", "")
        cnt[t] += 1
        issues = []
        if "warc-record-id" not in h or "warc-date" not in h:
            issues.append("필수 머리글(WARC-Record-ID·WARC-Date) 없음")
        if h.get("warc-block-digest"):
            ok = digest_ok(h["warc-block-digest"], block)
            if ok is False:
                issues.append("Block-Digest 불일치")
        if h.get("warc-payload-digest") and t == "response":
            payload = block.split(b"\r\n\r\n", 1)[1] if b"\r\n\r\n" in block else b""
            ok = digest_ok(h["warc-payload-digest"], payload)
            if ok is False:
                issues.append("Payload-Digest 불일치")
        if t == "response":
            m = re.match(rb"HTTP/\S+\s+(\d{3})", block)
            if m and int(m[1]) >= 400:
                issues.append(f"HTTP {m[1].decode()} 응답을 수집")
        rows.append([t, h.get("warc-target-uri", "").strip("<>"), h.get("warc-date", ""), len(block), " · ".join(issues) or "통과"])
    return rows, cnt


def channels_check(path, asof):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    head, out = rows[0], []
    g = lambda r, k: r[head.index(k)] if k in head and head.index(k) < len(r) else ""
    for r in rows[1:]:
        if not any(r):
            continue
        issues = []
        if not g(r, "운영부서"):
            issues.append("운영부서 없음")
        if not g(r, "주소").startswith("http"):
            issues.append("주소 없음")
        end, got = g(r, "종료일")[:10], g(r, "수집일")[:10]
        if end and (not got or got > end):
            issues.append("운영 종료 채널인데 종료 전 수집 기록 없음 — 폐지·종료 전 수집(지침: 폐지기관은 약 3개월 전부터)")
        if not end and got and got[:4] < str(asof - 1):
            issues.append(f"마지막 수집 {got[:4]}년 — 연 1회 주기 확인")
        out.append([g(r, "채널명"), g(r, "주소"), g(r, "운영부서"), end, got, " · ".join(issues) or "통과"])
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--url")
    ap.add_argument("--max-pages", type=int, default=40)
    ap.add_argument("--delay", type=float, default=0.5)
    ap.add_argument("--warc")
    ap.add_argument("--channels")
    ap.add_argument("--year", type=int, default=time.localtime().tm_year)
    ap.add_argument("-o", "--out", default="산출/웹기록물_점검.xlsx")
    a = ap.parse_args()
    if not (a.url or a.warc or a.channels):
        print("입력 오류: --url, --warc, --channels 중 하나 이상")
        return 2
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "점검 결과"
    ws.append(["항목", "상태", "내용", "조치·근거"])
    msg = []
    if a.url:
        findings, pages = site_check(a.url, a.max_pages, a.delay)
        for f in findings:
            ws.append(list(f))
        wp = wb.create_sheet("페이지")
        wp.append(["URL", "응답", "같은 호스트 링크", "자바스크립트 링크", "on* 이벤트", "script 수", "메모"])
        for r in pages:
            wp.append(r)
        msg.append(f"누리집 {len(pages)}쪽 · 고칠 것 {sum(1 for f in findings if f[1] == '고칠 것')} · 확인 {sum(1 for f in findings if f[1] == '확인')}")
    if a.warc:
        rows, cnt = warc_check(a.warc)
        ww = wb.create_sheet("WARC")
        ww.append(["WARC-Type", "Target-URI", "WARC-Date", "블록 바이트", "판정"])
        for r in rows:
            ww.append(r)
        bad = sum(1 for r in rows if r[4] != "통과")
        ws.append(["WARC 무결성", "통과" if not bad else "고칠 것", f"레코드 {len(rows)}개({', '.join(f'{k} {v}' for k, v in cnt.items())}) · 문제 {bad}",
                   "Digest 불일치는 재수집, 오류 응답만 담긴 쪽은 재현 불가 — 보존형식 WARC(ISO 28500)"])
        msg.append(f"WARC 레코드 {len(rows)} · 문제 {bad}")
    if a.channels:
        rows = channels_check(a.channels, a.year)
        wc = wb.create_sheet("SNS·블로그 채널")
        wc.append(["채널명", "주소", "운영부서", "종료일", "수집일", "판정"])
        for r in rows:
            wc.append(r)
        bad = sum(1 for r in rows if r[5] != "통과")
        ws.append(["SNS·블로그 채널", "통과" if not bad else "확인", f"{len(rows)}개 중 {bad}개 확인", "SNS·블로그도 웹기록물(시행령 제2조 10호)"])
        msg.append(f"채널 {len(rows)} · 확인 {bad}")
    for i, w in enumerate((20, 10, 70, 80)):
        ws.column_dimensions[chr(65 + i)].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    print(" · ".join(msg) + f" → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
