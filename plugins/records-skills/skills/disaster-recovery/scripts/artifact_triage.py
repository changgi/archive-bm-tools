#!/usr/bin/env python3
"""재난 피해 행정박물 응급조치 분류: 피해 행정박물 목록으로 재질별 응급조치 순서(즉시·24시간 / 후순위),
직물 48시간 기한과 동결 처리, 수해·화재별 세척·물기 제거·건조 요령, 상태조사표(참고 4), 현장 체크리스트(참고 3),
응급복구 키트 물품(참고 2)을 만든다. 평소에는 보유 목록으로 구조 우선순위(참고 5, 1~6등급·등외)를 매긴다.

사용:  python artifact_triage.py 피해박물.xlsx --found "2026-07-15 09:00" [--now "2026-07-15 18:00"] -o 산출/행정박물_응급조치.xlsx
       python artifact_triage.py 보유박물.xlsx --prepare -o 산출/행정박물_구조우선순위.xlsx
피해목록 열(머리글 이름으로 찾음): 관리번호, 명칭, 재질(목재·직물·피혁·서화·골각·패각·토기·금속·도자기·유리·석재·플라스틱…), 형태(현판·의복·가구·액자…),
       피해(수해·화재·수해+화재), 손상유형(A 젖음·B 그을림·C 오염·D 파손·E 곰팡이, 'A,B'처럼), 훼손도(1~3), 수량, 위치, 비고(자개·염색·기존훼손 등)
보유목록 열(--prepare): 관리번호, 명칭, 보존기간, 생산연도, 생산기관, 보존가치(높음·보통·낮음), 상징성(높음·보통·낮음), 특별기준(문화재·세계기록유산 등)
근거:  국가기록원 「재난 피해 행정박물 응급조치 실무 지침」(2024. 12., 복원관리과)
         2 작업공간: 곰팡이 생장 억제를 위해 20℃ 이하·습도 50% 이하, 직사광선 차단, 화재 피해는 환기되는 실내
         5 분류: 우선순위 응급조치(즉시 또는 24시간 내) 목재·직물·피혁·서화/회화·골각/패각·토기 / 후순위 금속·도자기·석재·플라스틱,
           직물은 48시간 안 응급조치 가능 여부 판단(넘으면 곰팡이), 표면 장식 목재·염색 직물·기존 훼손·오래된 것 먼저
         6 수해: 깨끗한 물로 약하게 세척(고압·문지르기 금지) → 눌러서 물기 제거 → 그늘에서 서서히 건조(드라이기·히터 금지, 습도 10%씩)
         7 화재: 그을음·재 제거에 물·수분 사용 금지, 헤파필터 진공청소기(노즐에 망사, 낮은 세기)
         참고 5 구조 우선순위(「국가기록원 기록물 재난관리 매뉴얼」 2023): 보존기간·생산시기·생산기관·보존가치·상징성 각 가·나·다 → 1~6등급·등외
원칙:  분류·목록만 만든다. 서화·회화·액자류와 실제 처리 방법은 복원 전문가와 상의한다. 박물에 직접 쓰지 않고 포장재에 식별정보를 쓴다.
"""
import argparse
import datetime as dt
import re
import sys

ALIAS = {"번호": ["관리번호", "박물번호", "등록번호", "번호"], "명칭": ["명칭", "박물명", "제목", "기록물명"], "재질": ["재질", "주요재질"],
         "형태": ["형태", "유형"], "피해": ["피해", "재난", "피해유형"], "손상": ["손상유형", "손상"], "훼손도": ["훼손도", "정도"],
         "수량": ["수량", "피해규모"], "위치": ["위치", "수습장소", "보관위치"], "비고": ["비고", "특이사항"],
         "보존": ["보존기간"], "연도": ["생산연도", "생산년도", "연도"], "기관": ["생산기관", "기관"], "가치": ["보존가치"],
         "상징": ["상징성"], "특별": ["특별기준", "지정"]}
FIRST = ["목재", "나무", "직물", "섬유", "의복", "피혁", "가죽", "서화", "회화", "그림", "액자", "골각", "뼈", "패각", "자개", "토기", "테라코타"]
LATER = ["금속", "도자기", "자기", "유리", "석재", "돌", "플라스틱", "합성수지"]
DAMAGE = {"A": "젖음", "B": "불에 타거나 그을림", "C": "오염", "D": "파손(찢김·손실·균열)", "E": "곰팡이"}
WET = {"목재": "나뭇결 방향으로 세척. 완전 침수 목재는 흡수재로 싸고 PE필름을 덮어 급속 건조 방지. 수평 바닥·받침대(팔렛트)에서 건조, 서랍·문을 조금 열어 통풍",
       "직물": "48시간 안 처리 불가면 지퍼백 포장 후 동결(직물·접힘면 사이 간지). 비틀거나 당기지 말고 눌러 물기 제거, 접힘은 물기가 빠진 뒤 펼침. 겹치지 말고 평평하게 건조, 어깨·소매에 흡수재 보충",
       "피혁": "비틀거나 당기지 말고 눌러 물기 제거. 겹치지 말고 평평하게 건조",
       "서화": "무리하게 처리하지 말고 복원 전문가와 먼저 상의",
       "도자기": "물줄기로도 손상될 수 있으니 미세 분무기로 약하게. 내부 오염물은 조심스럽게 기울여 제거. 문진·완충재로 넘어짐 방지",
       "금속": "표면이 매끄러워 흙 입자 상태로 문지르면 잔긁힘 — 물로 흘려 씻고 눌러 말림",
       "플라스틱": "표면이 매끄러워 흙 입자 상태로 문지르면 잔긁힘 — 물로 흘려 씻고 눌러 말림",
       "석재": "깨끗한 물로 흘려 씻고 흡수재로 눌러 물기 제거, 그늘에서 서서히 건조",
       "토기": "깨지기 쉬움 — 미세 분무기로 약하게, 완충재로 지지해 건조",
       "골각": "흡수재로 눌러 물기 제거, 급속 건조 금지(균열)"}
GROUP = {"나무": "목재", "섬유": "직물", "의복": "직물", "가죽": "피혁", "회화": "서화", "그림": "서화", "액자": "서화", "뼈": "골각", "패각": "골각",
         "자개": "목재", "테라코타": "토기", "자기": "도자기", "유리": "도자기", "돌": "석재", "합성수지": "플라스틱"}
KIT = [("작업자 보호", "방진복·장갑(일회용·면·고무·니트릴)·방진마스크·고글·방수 앞치마·방수 팔토시·고무장화"),
       ("운반·공간 준비", "완충포장재·바닥시트·필름·운반 상자·이동카트(모서리 완충재), 위험경고 테이프"),
       ("세척·오염 제거", "미세분무기·비이커·습식용 트레이·양동이·비닐시트·붓·무선 진공청소기(헤파필터)·블로워·스폰지·나무 주걱"),
       ("물기 제거·건조", "흡수지·흡수티슈·면타올·스폰지·무선 공기순환기·이동식 제습기"),
       ("포장·보관", "PE필름(랩)·지퍼백·얇은 간지·플라스틱 상자·실리카겔·보드지 파티션·타이벡"),
       ("복원용 소도구", "핀셋·가위·헤라·임시 보수테이프·넓은 리본끈"),
       ("기록", "라벨용지·포스트잇·노트·필기구(박물에 직접 쓰지 않음)"),
       ("동결처리", "지퍼백·산모아지(간지) — 48시간 안 처리 못 하는 직물")]
CHECK = [("대피 및 작업공간", "기록물 대피공간 및 작업공간이 마련되었는가?(20℃ 이하·습도 50% 이하, 직사광선 차단, 재질별 구역)"),
         ("응급조치 물품", "물품(도구 및 소모품 등)이 충분히 구비되어 있는가?"),
         ("상태조사", "박물의 주요 재질 조사를 했는가?"), ("상태조사", "긴급 응급조치 필요한 박물이 있는가?"),
         ("우선순위 선정", "응급조치 대상의 우선순위를 선정했는가?"),
         ("응급조치", "곰팡이가 발생한 기록물이 있다면, 별도 공간으로 분리하였는가?"),
         ("응급조치", "48시간 내 응급조치 불가한 직물류일 경우, 동결처리를 진행했는가?"),
         ("응급조치", "동결처리 과정에서 포장지에 기록물 정보(기록물명, 수습장소, 날짜 등)를 기입했는가?"),
         ("응급조치", "오염물을 제거했는가?"), ("응급조치", "젖은 박물의 물기를 제거했는가?"), ("응급조치", "건조 과정을 충분히 수행했는가?"),
         ("보관", "안전하게 포장하고 박물 식별정보를 기록했는가?('취급주의' 표기, 포개지 않음, 곰팡이·형태 변형 주기 관찰)")]
# 참고 5: (보존기간, 생산연도, 생산기관, 보존가치, 상징성) → 등급
RULES = [("가가가가가", "1"), ("가가나가가", "2"), ("나나가나나", "3"), ("나나나나나", "4"), ("다가가다다", "5"), ("다나나다다", "6"), ("다다나다다", "등외")]
PRIOR_ORG = ("국무회의", "외교부", "국정원", "국가정보원", "국방부", "통일부")
COLOR = {"1": "적색", "2": "황색", "3": "녹색"}


def find(head, key):
    for name in ALIAS[key]:
        for i, h in enumerate(head):
            if name in h.replace(" ", ""):
                return i
    return None


def read(path):
    import openpyxl
    if path.lower().endswith((".csv", ".txt")):
        import csv
        with open(path, encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh))
    else:
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    rows = [r for r in rows if any(r)]
    head = [h.replace(" ", "") for h in rows[0]]
    col = {k: find(head, k) for k in ALIAS}
    get = lambda r, k: (r[col[k]] if col[k] is not None and col[k] < len(r) else "").strip()
    return [{k: get(r, k) for k in ALIAS} for r in rows[1:]]


def material_group(m):
    for w in FIRST + LATER:
        if w in m:
            return GROUP.get(w, w)
    return ""


def grade_parts(r, year):
    p = r["보존"].replace(" ", "")
    a = "가" if p == "영구" else "나" if p == "준영구" else "다" if p else ""
    y = re.search(r"(1[89]\d\d|20\d\d)", r["연도"])
    b = "" if not y else "가" if int(y.group(1)) <= 1960 else "나" if int(y.group(1)) <= 1980 else "다"
    org = r["기관"]
    c = "" if not org else "가" if any(o in org for o in PRIOR_ORG) else "나" if re.search(r"(부|처|청|위원회)$", org) else "다"
    lvl = lambda v: "가" if "높" in v else "나" if "보통" in v else "다" if "낮" in v else ""
    return "".join(x or "?" for x in (a, b, c, lvl(r["가치"]), lvl(r["상징"])))


def rank(code, special):
    if special:
        return "1", "특별기준(국가지정문화재·세계기록유산·원장 지정) — 최우선"
    hit = next((g for k, g in RULES if k == code), None)
    if hit:
        return hit, "선정기준표와 일치"
    known = [i for i, ch in enumerate(code) if ch != "?"]
    best = min(RULES, key=lambda kg: (sum(1 for i in known if kg[0][i] != code[i]), RULES.index(kg)))
    miss = "·".join(n for n, ch in zip(("보존기간", "생산연도", "생산기관", "보존가치", "상징성"), code) if ch == "?")
    note = f"표 밖 조합 — 가장 가까운 행 {best[0]}({best[1]}{'등급' if best[1].isdigit() else ''})을 후보로 둠, 기록관 자문회의에서 확정"
    return best[1], note + (f" (빈 값: {miss})" if miss else "")


def parse_time(s):
    for f in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y.%m.%d %H:%M", "%Y%m%d%H%M"):
        try:
            return dt.datetime.strptime(s.strip(), f)
        except ValueError:
            pass
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="재난 피해 행정박물 응급조치 분류")
    ap.add_argument("list")
    ap.add_argument("--found", help="피해 발견(구조) 일시 YYYY-MM-DD HH:MM")
    ap.add_argument("--now", help="지금 일시(기본: 현재)")
    ap.add_argument("--prepare", action="store_true", help="평소: 보유 목록으로 구조 우선순위(참고 5)")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("-o", "--out", default="산출/행정박물_응급조치.xlsx")
    a = ap.parse_args()
    import openpyxl
    try:
        rows = read(a.list)
    except Exception as e:
        print(f"입력 오류: 목록을 읽을 수 없습니다({e})", file=sys.stderr)
        return 2
    wb = openpyxl.Workbook()
    if a.prepare:
        ws = wb.active
        ws.title = "구조 우선순위"
        ws.append(["관리번호", "명칭", "보존기간", "생산연도", "생산기관", "보존가치", "상징성", "기준 조합(가·나·다)", "행정박물 우선순위", "표식", "근거·확인"])
        out = []
        for r in rows:
            code = grade_parts(r, a.year)
            g, note = rank(code, r["특별"])
            out.append((int(g) if g.isdigit() else 9, [r["번호"], r["명칭"], r["보존"], r["연도"], r["기관"], r["가치"], r["상징"], code, g, COLOR.get(g, ""), note]))
        for _, x in sorted(out, key=lambda t: t[0]):
            ws.append(x)
        n = wb.create_sheet("선정기준")
        n.append(["보존기간·생산연도·생산기관·보존가치·상징성", "우선순위(종이·시청각·행정박물 같음)"])
        for k, g in RULES:
            n.append([k, g])
        n.append(["보존기간", "영구(가)·준영구(나)·30년 이하(다)"])
        n.append(["생산시기", "1960년 이전(가)·1961~1980년(나)·1981년 이후(다)"])
        n.append(["생산기관", "우선기관(가: 국무회의·외교부·국정원·국방부·통일부)·중앙행정기관(나)·기타(다) — 이름 끝 부·처·청·위원회를 중앙행정기관 후보로 봄"])
        n.append(["보존가치·상징성", "기록관별 자문회의 등으로 등급 구분"])
        n.append(["표식(안)", "1등급 적색·2등급 황색·3등급 녹색"])
        n.append(["근거", "「재난 피해 행정박물 응급조치 실무 지침」(2024. 12.) 참고 5, 출처 「국가기록원 기록물 재난관리 매뉴얼」(2023)"])
        wb.save(a.out)
        print(f"구조 우선순위 {len(out)}점 → {a.out}")
        return 0
    found = parse_time(a.found) if a.found else None
    now = parse_time(a.now) if a.now else dt.datetime.now()
    ws = wb.active
    ws.title = "응급조치 순서"
    ws.append(["순서", "관리번호", "명칭", "재질", "형태", "피해", "손상유형", "훼손도", "수량", "응급조치 군", "기한", "조치", "주의", "위치"])
    out = []
    for r in rows:
        grp = material_group(r["재질"] + r["형태"])
        dmg = sorted(set(re.findall(r"[A-E]", r["손상"].upper()))) or (["A"] if "수해" in r["피해"] else ["B"] if "화재" in r["피해"] else [])
        fire, wet = "B" in dmg or "화재" in r["피해"], "A" in dmg or "수해" in r["피해"]
        first = grp in ("목재", "직물", "피혁", "서화", "골각", "토기")
        extra = bool(re.search(r"자개|금속\s*장식|장식|염색|염료|기존\s*훼손|부식|파손|오래", r["비고"] + r["재질"]))
        tier = "우선(즉시·24시간 내)" if first else "후순위" if grp else "재질 확인"
        if re.search(r"종이|증서|상장|한지", r["재질"]) and grp in ("", "서화"):
            tier = "우선(즉시·24시간 내)" if grp else "종이 — 종이기록물 응급복구(disaster_triage.py)"
        due, deadline = "", None
        if first and found:
            deadline = found + dt.timedelta(hours=24)
            due = deadline.strftime("%m-%d %H:%M까지")
        acts, warn = [], []
        if grp == "직물" and wet and found:
            d48 = found + dt.timedelta(hours=48)
            left = (d48 - now).total_seconds() / 3600
            due = f"48시간 {d48.strftime('%m-%d %H:%M')} (남은 {left:.0f}시간)"
            if left <= 0:
                warn.append("48시간 지남 — 곰팡이 생장 가능, 동결처리 후 순차 처리")
        if "E" in dmg:
            acts.append("곰팡이 — 다른 공간·이동식 후드로 분리, 건조 상태에서 깨끗한 붓으로 제거(붓 자주 교체)")
        if fire:
            acts.append("그을음·재: 물·수분 절대 쓰지 않음, 헤파필터 진공청소기 낮은 세기(노즐에 망사, 박물에 닿지 않게), 접힌 직물은 겉 → 안")
            warn.append("유해가스 잔류 — 환기되는 실내, 방진복·니트릴 장갑·마스크·고글")
        if wet:
            acts.append("세척: 깨끗한 물로 약하게(고압·문지르기 금지) → 물기: 흡수재로 눌러서 → 건조: 그늘·공기순환, 드라이기·히터·직사광선 금지, 습도 50% 이하로 10%씩")
            if grp in WET:
                acts.append(WET[grp])
        if fire and wet:
            warn.append("소화수 피해 — 수해 응급조치(6장)를 따르고 그을음은 건조 뒤 진공")
        if "D" in dmg:
            acts.append("파손: 조각 모두 수거해 개별 완충 포장 후 지퍼백 / 균열: PE필름 임시 고정·보수테이프 / 느슨한 연결: 리본끈·랩으로 임시 고정")
        if grp == "서화":
            warn.append("서화·회화·액자류 — 전문가와 먼저 상의")
        if not r["번호"]:
            warn.append("식별정보 없음 — 포장재·이름표에 기록(박물에 직접 쓰지 않음)")
        if extra:
            warn.append("같은 군 안에서 먼저(표면 장식·염색·기존 훼손·오래된 것)")
        dg = int(r["훼손도"]) if r["훼손도"].isdigit() else 0
        key = (0 if first else 1 if grp else 2, 0 if extra else 1, -dg, 0 if "E" in dmg else 1)
        out.append((key, [r["번호"], r["명칭"], r["재질"], r["형태"], r["피해"], ",".join(dmg), r["훼손도"], r["수량"], tier, due,
                          " / ".join(acts) or "상태 확인 후 결정", " / ".join(warn), r["위치"]]))
    out.sort(key=lambda t: t[0])
    for i, (_, x) in enumerate(out, 1):
        ws.append([i] + x)
    s = wb.create_sheet("상태조사표(참고4)")
    s.append(["재질/형태", "피해 규모(수량)", "손상 유형/훼손도", "응급조치 필요 여부", "긴급", "비고"])
    for _, x in out:
        codes = ",".join(f"{c}{x[6] or ''}" for c in x[5].split(",") if c)
        s.append([f"{x[2]}/{x[3]}", x[7], codes, "네" if x[5] else "확인", "긴급" if x[8].startswith("우선") else "", x[11]])
    s.append([])
    s.append(["손상 유형", "A 젖음 · B 불에 타거나 그을림 · C 오염 · D 파손(찢김·손실·균열 등) · E 곰팡이", "훼손도", "1·2·3"])
    c = wb.create_sheet("현장 체크리스트(참고3)")
    c.append(["영역", "점검 항목", "이행 여부", "근거 자료"])
    for area, q in CHECK:
        c.append([area, q, "", ""])
    k = wb.create_sheet("응급복구 키트(참고2)")
    k.append(["구분", "물품"])
    for row in KIT:
        k.append(list(row))
    g = wb.create_sheet("근거")
    for line in ("「재난 피해 행정박물 응급조치 실무 지침」(국가기록원 복원관리과, 2024. 12.) 2~8장·참고 2~5",
                 "적용근거: 「공공기록물 관리에 관한 법률」 제30조(기록물 보안 및 재난 대책) 제1항",
                 "종이 재질 행정박물(증서·상장 등)은 종이기록물 응급복구(disaster_triage.py, NAK 38:2024)를 쓴다.",
                 "응급조치는 훼손 심화·멸실을 막는 최소한의 임시 조치다. 처리 방법은 보존처리·복원 담당자가 정한다.",
                 f"발견 일시 {a.found or '(안 줌 — 기한 계산 안 함)'}, 기준 시각 {now.strftime('%Y-%m-%d %H:%M')}"):
        g.append([line])
    wb.save(a.out)
    print(f"행정박물 {len(out)}점: 우선 {sum(1 for _, x in out if x[8].startswith('우선'))} · 후순위 {sum(1 for _, x in out if x[8] == '후순위')} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
