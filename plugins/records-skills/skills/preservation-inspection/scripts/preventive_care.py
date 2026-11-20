#!/usr/bin/env python3
"""보존 취약 종이기록물 예방조치 계획: 상태검사·정리 목록으로 철마다 취약 유형(재질·구조·편철 재료·훼손 심화)을 가리고,
처리방향·조치 단위(면/철)·필요한 보존 용품과 수량, 복원 연계 대상을 정한다. 작업 뒤 후속관리 환경 기준도 붙인다.

사용:  python preventive_care.py 기록물목록.xlsx -o 산출/예방조치_계획.xlsx
목록 열(머리글 이름으로 찾음): 관리번호(철분류번호), 철제목, 매수(쪽수), 재질(감열지·팩스·청사진·사진·트레이싱지…), 규격(혼합·A3·대형…),
       편철(클립·스테이플러·핀·고무줄…), 상태(접힘·노출·바스라짐·파손·곰팡이·변색…), 해당 면(예: 3-5,8), 비고
근거:  국가기록원 「보존 취약 종이기록물 예방조치 지침」(2024. 11.)
         2.1 대상: ① 재질 취약(감열지·감광지·사진·트레이싱지) ② 구조적 취약(다규격 혼합철·거대 편철 500매 초과·편철 재료보다 큰 대형 문서·접힌 문서/도면)
                   ③ 편철 재료 취약(클립·금속핀 등 금속철, 고무줄) ④ 훼손 심화(바스라짐·심한 파손)
         2.3.3 표 1·3·5·7: 감열지·감광지·사진 → 기록면 위 간지(면 단위, 비슷하거나 얇은 두께) / 트레이싱지 → 보호 폴더
                   다규격 혼합철·거대 편철 → 편철 개선(보호 보드·벨크로, 철 단위) / 대형 문서 → 앞뒤 보호 보드 / 접힌 문서·도면 → 보호 폴더(더 두꺼운 용지)
                   금속철 → 전용 도구로 제거, 필요하면 부식방지 클립 / 고무줄 → 제거, 필요하면 벨크로
                   훼손 심화 → 보존용 봉투 밀봉(지지가 약하면 보존상자) + 복원 연계
         간지·폴더에는 처리한 면(범위)을 연필로 적는다. 기존 보존 봉투·상자와 서가 배치를 유지하고, 부피가 늘면 투입량을 조절한다.
         3 후속관리: 시행령 별표 6 — 온도 20±2℃·습도 50±5%, 보존서고 조명 100~300럭스 등
         참고: 일반문서류 철당 100매 이내 원칙(공공기록물법 시행규칙 제9조제3항)
원칙:  계획·목록만 만든다. 원본 취급과 조치는 보존 담당자가 한다. 복원·탈산 같은 직접 처리는 restoration_plan.py로 넘긴다.
"""
import argparse
import re
import sys

ALIAS = {"번호": ["관리번호", "철분류번호", "분류번호", "번호"], "제목": ["철제목", "제목", "기록물철제목"], "매수": ["매수", "쪽수", "기록물쪽수", "면수"],
         "재질": ["재질", "용지"], "규격": ["규격", "크기"], "편철": ["편철", "편철재료", "철끈"], "상태": ["상태", "훼손", "상태검사"],
         "면": ["해당면", "해당쪽", "면", "페이지"], "비고": ["비고", "특이사항"]}
RULES = [
    # (유형, 대상, 정규식, 대상 열, 처리, 단위, 용품, 두께)
    ("재질 취약", "감열지", r"감열|팩스|팩시밀리|영수증|FAX", "재질", "기록면 위에 간지 배치(기록 내용 보호)", "면", "간지", "기록물과 비슷하거나 얇게"),
    ("재질 취약", "감광지", r"감광|청사진|습식\s*복사|블루프린트", "재질", "기록면 위에 간지 배치", "면", "간지", "기록물과 비슷하거나 얇게"),
    ("재질 취약", "사진(부착)", r"사진|인화", "재질", "기록면 위에 간지 배치", "면", "간지", "기록물과 비슷하거나 얇게"),
    ("재질 취약", "트레이싱지", r"트레이싱|기름종이|반투명", "재질", "보호 폴더 사이에 끼움(인접 기록물 산성 전이 방지), 많으면 한 폴더에 일괄", "면", "보호 폴더", "간지와 비슷하게"),
    ("구조적 취약", "다규격 혼합철", r"혼합|다규격|여러\s*규격", "규격", "편철 개선 — 벨크로 고정 + 보호 보드 지지", "철", "보호 보드·벨크로", "기록물보다 두껍고 단단하게"),
    ("구조적 취약", "대형 문서", r"대형|노출|A2|A1|A0|B3|B2", "규격 상태", "철 앞뒤 보호 보드(노출 부분은 기록물 크기에 맞춤)", "철", "보호 보드 2장", "기록물보다 두껍고 단단하게"),
    ("구조적 취약", "접힌 문서·도면", r"접힌|접힘|접음|접혀", "상태 규격 재질", "접힌 모서리를 감싸는 보호 폴더", "면", "보호 폴더(두꺼운 보존 용지)", "기록물보다 두껍게"),
    ("편철 재료 취약", "금속철", r"클립|스테이플|호치키스|핀|금속|철침|압정|녹", "편철 상태", "전용 도구로 금속철 제거(긁힘·찢김 주의), 필요하면 부식방지 보존용 클립", "철", "금속철 제거 도구·보존용 클립", ""),
    ("편철 재료 취약", "고무줄", r"고무줄|고무밴드|밴드", "편철 상태", "고무줄 제거(탄력·오염 주의), 필요하면 벨크로", "철", "벨크로", ""),
    ("훼손 심화", "훼손 심화", r"바스라|바스러|부스러|파편|심한\s*파손|조각|곰팡이|건조화|산화|취화", "상태", "보존용 봉투(비닐백) 밀봉 분리, 지지가 약하면 보존상자 → 복원 연계", "철·면", "보존용 봉투(+보존상자)", ""),
]
ENV = [("온도", "20±2℃"), ("습도", "50±5%"), ("미세먼지(PM-10)", "50㎍/㎥"), ("이산화황(SO2)", "0.05ppm 이하"), ("산화질소(NOx)", "0.05ppm 이하"),
       ("포름알데히드(HCHO)", "120㎍/㎥"), ("휘발성유기화합물(VOC)", "400㎍/㎥"), ("조명", "보존서고 100~300럭스(자외선 차단등), 전시관 50~200럭스"),
       ("취급", "정돈된 작업환경, 장갑 또는 깨끗한 손, 마스크")]


def find(head, key):
    for name in ALIAS[key]:
        for i, h in enumerate(head):
            if name in h.replace(" ", ""):
                return i
    return None


def read(path):
    if path.lower().endswith((".csv", ".txt")):
        import csv
        with open(path, encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh))
    else:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    rows = [r for r in rows if any(r)]
    head = [h.replace(" ", "") for h in rows[0]]
    col = {k: find(head, k) for k in ALIAS}
    get = lambda r, k: (r[col[k]] if col[k] is not None and col[k] < len(r) else "").strip()
    return [{k: get(r, k) for k in ALIAS} for r in rows[1:]], [k for k, v in col.items() if v is None]


def pages(spec):
    n = 0
    for part in re.split(r"[,\s]+", spec or ""):
        m = re.fullmatch(r"(\d+)\s*[-~]\s*(\d+)", part)
        if m:
            n += abs(int(m.group(2)) - int(m.group(1))) + 1
        elif part.isdigit():
            n += 1
    return n


def main() -> int:
    ap = argparse.ArgumentParser(description="보존 취약 종이기록물 예방조치 계획")
    ap.add_argument("list")
    ap.add_argument("--giant", type=int, default=500, help="거대 편철 기준 매수(지침 500)")
    ap.add_argument("-o", "--out", default="산출/예방조치_계획.xlsx")
    a = ap.parse_args()
    import openpyxl
    try:
        rows, missing = read(a.list)
    except Exception as e:
        print(f"입력 오류: 목록을 읽을 수 없습니다({e})", file=sys.stderr)
        return 2
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "예방조치 계획"
    ws.append(["관리번호", "철제목", "매수", "취약 유형", "대상", "처리방향", "조치 단위", "해당 면", "보존 용품", "용지 두께", "복원 연계", "확인"])
    supply, kinds, plain, restore = {}, {}, 0, []
    for r in rows:
        hits = []
        for typ, tgt, rx, cols, act, unit, item, thick in RULES:
            text = " ".join(r[c] for c in cols.split()) + " " + r["비고"]
            if re.search(rx, text, re.I):
                hits.append((typ, tgt, act, unit, item, thick))
        n = int(re.sub(r"\D", "", r["매수"]) or 0)
        if n > a.giant:
            hits.append(("구조적 취약", "거대 편철", f"편철 개선 — 벨크로 고정 + 보호 보드 지지({n}매 > {a.giant})", "철", "보호 보드·벨크로", "기록물보다 두껍고 단단하게"))
        note = []
        if 100 < n <= a.giant:
            note.append(f"{n}매 — 일반문서류 철당 100매 이내 원칙(시행규칙 제9조제3항), 분철 검토")
        if not hits:
            plain += 1
            continue
        pg = pages(r["면"])
        for typ, tgt, act, unit, item, thick in hits:
            link = "복원·탈산 등 후속조치 연계(restoration_plan.py)" if typ == "훼손 심화" else ""
            if link:
                restore.append(r["번호"] or r["제목"])
            chk = list(note)
            if unit == "면" and not r["면"]:
                chk.append("해당 면을 적으면 용품 수량을 셈 — 간지·폴더에 처리 면을 연필로 기입")
            if typ == "재질 취약" and r["상태"] and re.search(r"휘발|흐림|안\s*보|탈색", r["상태"]):
                chk.append("내용이 이미 흐려짐 — 사본(스캔) 확보 검토")
            ws.append([r["번호"], r["제목"], n or "", typ, tgt, act, unit, r["면"], item, thick, link, " / ".join(chk)])
            kinds[typ] = kinds.get(typ, 0) + 1
            for it in item.replace("(+보존상자)", "·보존상자(필요시)").split("·"):
                it = it.strip()
                if not it:
                    continue
                q = 2 if "2장" in it else (pg or 1) if unit == "면" else 1
                it = it.replace(" 2장", "")
                supply[it] = supply.get(it, 0) + q
    s = wb.create_sheet("용품 준비")
    s.append(["보존 용품", "필요 수량(추정)", "비고"])
    for it, q in sorted(supply.items(), key=lambda x: -x[1]):
        s.append([it, q, "면 단위는 '해당 면' 수, 없으면 철당 1로 셈 — 규격 사전 조사 후 재단"])
    s.append(["공통 도구", "", "재단기·커터·커팅보드, 연필(기입용), 유해가스 흡착시트(필요시)"])
    t = wb.create_sheet("요약")
    t.append(["구분", "값"])
    t.append(["목록 철 수", len(rows)])
    t.append(["예방조치 대상 아님", plain])
    for k, v in kinds.items():
        t.append([k, v])
    t.append(["복원 연계", len(set(restore))])
    if missing:
        t.append(["못 찾은 열", ", ".join(missing)])
    t.append(["절차", "사전 준비(재료 재단·작업환경) → 처리방향 설정 → 예방조치 실시 → 후속관리"])
    t.append(["원칙", "기존 보존 봉투·상자와 서가 배치를 유지하고, 간지·폴더 투입으로 부피가 늘면 투입량을 조절한다."])
    t.append(["근거", "국가기록원 「보존 취약 종이기록물 예방조치 지침」(2024. 11.) 2.1·2.3·3, 표 1·3·5·7"])
    e = wb.create_sheet("후속관리 환경")
    e.append(["구분", "기준(시행령 별표 6, 지침 표 9)"])
    for row in ENV:
        e.append(list(row))
    wb.save(a.out)
    print(f"철 {len(rows)} · 예방조치 {sum(kinds.values())}건({', '.join(f'{k} {v}' for k, v in kinds.items())}) · 복원 연계 {len(set(restore))} → {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
