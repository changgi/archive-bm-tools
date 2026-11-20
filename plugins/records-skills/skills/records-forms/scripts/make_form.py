#!/usr/bin/env python3
"""법정 서식 작성: 공공기록물법 시행규칙 별지 서식을 칸 구성 그대로 엑셀로 채운다. 빈 필수 칸은 '빈칸 점검' 시트에 모은다.

사용:  python make_form.py <서식> [--list 목록.xlsx] [--info 기관정보.json] [--set 항목=값 ...] -o 산출/서식.xlsx
       python make_form.py --forms            (서식 이름과 칸 보기)
서식 이름(별지 번호):
  index      제1호   색인목록(일반문서류) — 목록의 쪽수로 쪽표시(1-3, 4-9 …)를 철마다 이어서 계산
  media      제4호   전자매체 수록계획서 — 분량·파일크기 합계
  treatment  제6호   기록물보존처리서
  carry      제8호   기록물 반출·반입서(철 단위·건 단위) — 보존기간 30년 이상을 서고 밖으로 내면 표시(시행규칙 제32조)
  restore    제9호   기록물복원처리서
  pubreg     제11호  간행물 발간등록번호 부여 신청
  hold       제11호의2 기록물의 폐기 금지 요청서     hold-release 제11호의3 해제 요청서     hold-extend 제11호의4 연장 요청서
  annotate   제11호의5 기록물 부기 또는 정정 요청서
  microfilm  제5호   마이크로필름 촬영계획서
  limited    제14호  비공개기록물 제한적 열람신청서 / review 제14호 2. 비공개기록물 열람 재심의 요청서(주민등록번호는 파일에 적지 않음)
  extend     제1호의2 기록물 이관 기간 연장·일정 연기 목록 (2027. 1. 1. 시행, 시행규칙 제20조 — 처리과→기록관, 기록관→영구기록물관리기관)
근거: 공공기록물법 시행규칙 별지 서식(law.go.kr 서식 보기로 칸 확인, 2026-11-20) — 제9조②(색인목록), 제25조(전자매체),
      제30조(보존처리), 제32조(반출 제한), 제34조(복원), 제36조(발간등록), 제37조의2(폐기 금지), 제38조의2(부기·정정, 2026. 9. 11. 개정)
원칙:  서식을 채우기만 한다. 제출·결재·직인은 사람이 한다. 한글(HWPX)로 바꿀 때는 hwpx-report를 쓴다.
"""
import argparse
import json
import os
import re
import sys

REQ_ORG = ["기관명", "소재지", "대표자 성명", "담당부서명", "담당부서장명", "담당부서장 전화번호", "담당자명", "담당자 전화번호"]
FORMS = {
    "index": {"no": "제1호", "title": "색인목록", "basis": "시행규칙 제9조② — 철 맨 위 별지 2 표지, 그다음 색인목록, 그다음 문서",
              "fields": [], "per": "철",
              "cols": ["일련번호", "일자", "등록번호", "제목", "보낸 기관", "받은 기관", "쪽표시", "전자문서여부", "비고"],
              "src": {"일자": ["일자", "생산일자", "접수일자", "등록일자"], "등록번호": ["등록번호", "문서번호", "생산등록번호"], "제목": ["제목", "건제목"],
                      "보낸 기관": ["보낸기관", "보낸 기관", "발신기관", "발신"], "받은 기관": ["받은기관", "받은 기관", "수신기관", "수신"],
                      "쪽수": ["쪽수", "면수", "매수"], "전자문서여부": ["전자문서여부", "전자여부", "기록물형태"], "비고": ["비고"],
                      "철번호": ["철분류번호", "분류번호", "기록물철분류번호"], "철제목": ["철제목", "기록물철제목"], "생산연도": ["생산연도", "생산년도"]}},
    "extend": {"no": "제1호의2", "title": "기록물 이관 기간 연장·일정 연기 목록",
               "basis": "시행규칙 제20조(2026. 9. 11. 전문개정, 2027. 1. 1. 시행) — 처리과: 영 제32조①단서 기간 연장 협의 시 기록관·특수기록관(영 제10조② 공공기관은 지정 부서)에, "
                        "기록관: 영 제40조③ 일정 연기(제시받은 날부터 30일 안 협의, 최대 1년)·제41조③ 시기 연장 시 영구기록물관리기관에 첨부",
               "fields": ["기관명"],
               "cols": ["일련번호", "처리과기관코드", "기록물철분류번호", "문서등록번호", "기록물철(건) 제목", "생산연도", "이관희망시기", "사유"],
               "src": {"처리과기관코드": ["처리과기관코드", "처리과코드", "기관코드"], "기록물철분류번호": ["기록물철분류번호", "철분류번호", "분류번호"],
                       "문서등록번호": ["문서등록번호", "등록번호", "생산등록번호"], "기록물철(건) 제목": ["기록물철(건) 제목", "철제목", "건제목", "제목"],
                       "생산연도": ["생산연도", "생산년도"], "이관희망시기": ["이관희망시기", "이관희망연도", "희망시기"], "사유": ["사유", "연장사유", "연기사유"]}},
    "media": {"no": "제4호", "title": "전자매체 수록계획서", "basis": "시행규칙 제25조", "fields": ["매체번호", "수록일자", "담당자"],
              "cols": ["분류번호", "기록물철 제목", "분량", "파일크기"],
              "src": {"분류번호": ["철분류번호", "분류번호"], "기록물철 제목": ["철제목", "기록물철제목", "제목"], "분량": ["분량", "건수", "쪽수"],
                      "파일크기": ["파일크기", "크기", "용량"]}},
    "treatment": {"no": "제6호", "title": "기록물보존처리서", "basis": "시행규칙 제30조", "fields": ["제 호", "처리구분", "일자", "담당자"],
                  "cols": ["일련번호", "분류번호", "제목", "기록물형태", "수량", "비고"],
                  "src": {"분류번호": ["철분류번호", "분류번호", "등록번호"], "제목": ["철제목", "제목"], "기록물형태": ["기록물형태", "형태", "유형"],
                          "수량": ["수량", "권수", "매수"], "비고": ["비고"]}},
    "microfilm": {"no": "제5호", "title": "마이크로필름 촬영계획서", "basis": "시행규칙 제26조 <2026. 9. 11.> — 촬영 전에 시스템으로 작성, 촬영 시작부분에 별표 13 시작표지와 이 계획서를 넣고 계획서 순서대로 수록",
                  "fields": ["필름번호", "촬영일자", "사용필름", "필름규격", "촬영자", "확인"],
                  "cols": ["분류번호", "제목", "세부 목록", "분량", "촬영 위치(시작)", "촬영 위치(끝)", "검사결과", "확인"],
                  "src": {"분류번호": ["철분류번호", "분류번호"], "제목": ["철제목", "기록물철제목", "제목"], "세부 목록": ["세부목록", "건제목"],
                          "분량": ["분량", "쪽수", "매수"], "촬영 위치(시작)": ["시작", "촬영시작", "시작프레임"], "촬영 위치(끝)": ["끝", "촬영끝", "끝프레임"],
                          "검사결과": ["검사결과"]}},
    "limited": {"no": "제14호", "title": "비공개기록물 제한적 열람신청서", "basis": "법 제37조①, 시행령 제73조①~③ — 영구기록물관리기관의 장은 10일 안 결정(생산기관 의견조회·기록물공개심의회 심의 필요 시 10일 범위 연장), 목적 증명 자료 요청 가능",
                "fields": ["신청인 구분(당사자/대리인)", "성명(기관·단체명)", "소속", "전화번호", "주소", "전자우편", "열람신청 기록물", "신청목적", "신청목적 내 사용 동의(동의/동의하지 않음)"],
                "cols": []},
    "review": {"no": "제14호", "title": "비공개기록물 열람 재심의 요청서", "basis": "시행령 제73조④·⑤ — 결정에 불복하면 7일 안 재심의 요청, 영구기록물관리기관은 7일 안 기록물공개심의회에서 재결정",
               "fields": ["신청인 구분(당사자/대리인)", "성명(기관·단체명)", "전화번호", "주소", "전자우편", "열람신청 기록물", "재심의 요청 취지 및 이유"],
               "cols": []},
    "carry": {"no": "제8호", "title": "기록물 반출·반입서", "basis": "시행규칙 제32조 — 보존기간 30년 이상은 서고 밖 반출 원칙 금지, 기관 밖은 기관장 승인 + 이 서식을 서고관리책임자에게",
              "fields": ["제 호", "구분(반출/반입)", "일자", "소속기관명", "직급", "성명", "반출기간", "사유"],
              "cols": ["일련번호", "분류번호", "제목", "기록물형태", "수량", "확인"],
              "src": {"분류번호": ["철분류번호", "분류번호", "등록번호"], "제목": ["철제목", "제목"], "기록물형태": ["기록물형태", "형태"],
                      "수량": ["수량", "권수"], "보존기간": ["보존기간"]}},
    "restore": {"no": "제9호", "title": "기록물복원처리서", "basis": "시행규칙 제34조① — 복원 과정과 전후 상태도 기록으로 남김", "fields": ["제 호", "담당자"],
                "cols": ["일련번호", "분류번호", "건번호", "쪽수", "생산연도", "제목", "복원처리내용", "작업자", "비고"],
                "src": {"분류번호": ["철분류번호", "분류번호"], "건번호": ["건번호", "등록번호"], "쪽수": ["쪽수", "복원쪽"], "생산연도": ["생산연도"],
                        "제목": ["제목", "철제목"], "복원처리내용": ["복원처리내용", "처리내용"], "작업자": ["작업자"], "비고": ["비고"]}},
    "pubreg": {"no": "제11호", "title": "간행물의 발간등록번호부여의 신청", "basis": "시행규칙 제36조, 시행령 제55조 — 원고 완성 뒤 인쇄 결재 전 신청, 발간 후 15일 안 3부 송부",
               "fields": ["제목", "신청일", "발행기관", "주관부서", "발간주기", "용도", "발간매체(1 책자·2 파일·3 책자+파일·4 그 밖)", "전화번호", "팩스번호", "내용요약"],
               "cols": []},
    "hold": {"no": "제11호의2", "title": "기록물의 폐기 금지 요청서", "basis": "법 제27조의3①, 시행령 제54조의2①1·4호, 시행규칙 제37조의2① → 국가기록원장",
             "fields": REQ_ORG + ["폐기 금지 요청 이유", "폐기 금지 대상 기관", "폐기 금지 대상 기록물", "폐기 금지 요청 기간"], "cols": []},
    "hold-release": {"no": "제11호의3", "title": "기록물의 폐기 금지 해제 요청서", "basis": "시행령 제54조의3②, 시행규칙 제37조의2② → 국가기록원장",
                     "fields": REQ_ORG + ["폐기 금지 해제 요청 이유", "폐기 금지 해제 대상 기관", "폐기 금지 해제 대상 기록물", "폐기 금지 해제 요청일"], "cols": []},
    "hold-extend": {"no": "제11호의4", "title": "기록물의 폐기 금지 연장 요청서", "basis": "시행령 제54조의3④, 시행규칙 제37조의2③ → 국가기록원장",
                    "fields": REQ_ORG + ["폐기 금지 연장 요청 이유", "폐기 금지 연장 대상 기관", "폐기 금지 연장 대상 기록물", "폐기 금지 연장 요청 기간"], "cols": []},
    "annotate": {"no": "제11호의5", "title": "기록물 부기 또는 정정 요청서", "basis": "법 제30조의3, 시행령 제62조의3, 시행규칙 제38조의2 → 그 기록물을 보유한 기록물관리기관의 장, 사유 증명 서류 첨부. 처리 방법: 국가기록원훈령 제224호",
                 "fields": ["기관명", "소재지", "기관장 성명", "담당부서명", "담당부서장명", "담당부서장 전화번호", "담당자명", "담당자 전화번호",
                            "부기 또는 정정의 대상 기록물", "부기 또는 정정의 사유 및 내용"], "cols": []},
}
FOOT = {"limited": "※ 주민등록번호는 이 파일에 적지 않고 출력한 서식에 신청인이 직접 적는다. 신청 목적 외 용도로 쓰면 공공기록물법 제51조제4호(3년 이하 징역 또는 2천만원 이하 벌금). 처리: 작성 → 접수 → 검토 → 결정 통지(10일 이내, 10일 범위 연장 가능).",
        "review": "※ 주민등록번호는 이 파일에 적지 않고 출력한 서식에 신청인이 직접 적는다. 시행령 제73조제4항·제5항에 따른 재심의 요청.",
"hold": "붙임 서류: 기록물의 폐기 금지 시행에 참고가 될 만한 자료가 있는 경우에 첨부합니다.",
        "hold-release": "붙임 서류: 기록물의 폐기 금지 해제에 참고가 될 만한 자료가 있는 경우에 첨부합니다.",
        "hold-extend": "붙임 서류: 기록물의 폐기 금지 연장에 참고가 될 만한 자료가 있는 경우에 첨부합니다.",
        "annotate": "첨부 서류: 부기 또는 정정의 사유를 증명할 수 있는 서류"}
TO = {"limited": "(접수기관의 장) 귀하", "review": "(접수기관의 장) 귀하", "hold": "국가기록원장 귀하", "hold-release": "국가기록원장 귀하", "hold-extend": "국가기록원장 귀하", "annotate": "기록물관리기관의 장 귀하"}


def read_rows(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    hi = max(range(min(10, len(rows))), key=lambda i: sum(bool(c) for c in rows[i])) if rows else 0
    return rows[hi], [r for r in rows[hi + 1:] if any(r)]


def pick(head, row, names):
    for n in names:
        if n in head:
            i = head.index(n)
            return row[i] if i < len(row) else ""
    return ""


def pages(v):
    m = re.search(r"\d+", str(v or ""))
    return int(m[0]) if m else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("form", nargs="?")
    ap.add_argument("--list")
    ap.add_argument("--info", help="기관·담당자 정보 JSON({\"기관명\": …})")
    ap.add_argument("--set", action="append", default=[], help="항목=값 (여러 번)")
    ap.add_argument("--forms", action="store_true")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    if a.forms or not a.form:
        for k, f in FORMS.items():
            print(f"{k:13} 별지 {f['no']:8} {f['title']} — 칸: {', '.join(f['fields'] + f['cols'])}")
        return 0
    if a.form not in FORMS:
        print(f"서식 이름 '{a.form}' — --forms로 목록을 보세요.", file=sys.stderr)
        return 2
    F = FORMS[a.form]
    info = {}
    if a.info:
        info.update(json.load(open(a.info, encoding="utf-8")))
    for kv in a.set:
        if "=" in kv:
            k, v = kv.split("=", 1)
            info[k.strip()] = v.strip()
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "서식"
    ws.append([f"■ 공공기록물 관리에 관한 법률 시행규칙 [별지 {F['no']}서식]"])
    ws.append([F["title"]])
    ws["A2"].font = Font(size=16, bold=True)
    missing = []
    for k in F["fields"]:
        v = info.get(k, "")
        ws.append([k, v])
        ws.cell(ws.max_row, 1).border = ws.cell(ws.max_row, 2).border = box
        if not v:
            missing.append(["머리 칸", k, "비어 있음"])
    warn = []
    rows_out = 0
    if F["cols"]:
        if not a.list:
            missing.append(["목록", "--list", "목록 파일이 없어 표를 비워 둠"])
            ws.append([])
            ws.append(F["cols"])
        else:
            head, rows = read_rows(a.list)
            src = F.get("src", {})
            if a.form == "index":
                groups = {}
                for r in rows:
                    key = (pick(head, r, src["철번호"]), pick(head, r, src["철제목"]), pick(head, r, src["생산연도"]))
                    groups.setdefault(key, []).append(r)
                for (no, title, yr), lst in groups.items():
                    ws.append([])
                    ws.append(["1. 일반문서류", f"기록물철 제목: {title}", f"분류번호: {no}", f"(생산연도: {yr})"])
                    ws.append(F["cols"])
                    for c in range(1, len(F["cols"]) + 1):
                        ws.cell(ws.max_row, c).font = Font(bold=True)
                        ws.cell(ws.max_row, c).border = box
                    p = 0
                    for i, r in enumerate(lst, 1):
                        n = pages(pick(head, r, src["쪽수"]))
                        mark = f"{p + 1}-{p + n}" if n > 1 else (str(p + 1) if n == 1 else "")
                        if not n:
                            warn.append(["행", f"{no} {i}번", "쪽수 없음 — 쪽표시를 비움"])
                        p += n
                        el = pick(head, r, src["전자문서여부"])
                        el = "전자" if el and "비전자" not in el and el not in ("N", "아니오") else ("비전자" if el else "")
                        ws.append([i, pick(head, r, src["일자"]), pick(head, r, src["등록번호"]), pick(head, r, src["제목"]),
                                   pick(head, r, src["보낸 기관"]), pick(head, r, src["받은 기관"]), mark, el, pick(head, r, src["비고"])])
                        for c in range(1, len(F["cols"]) + 1):
                            ws.cell(ws.max_row, c).border = box
                        rows_out += 1
                    if p > 100:
                        warn.append(["철", no or title, f"{p}매 — 철당 100매 이내 원칙(시행규칙 제9조③), 넘으면 같은 제목·분류번호로 분권하고 권호수만 다르게"])
            else:
                ws.append([])
                ws.append(F["cols"])
                for c in range(1, len(F["cols"]) + 1):
                    ws.cell(ws.max_row, c).font = Font(bold=True)
                    ws.cell(ws.max_row, c).border = box
                tot_q, tot_size = 0, 0.0
                for i, r in enumerate(rows, 1):
                    vals = []
                    for c in F["cols"]:
                        if c == "일련번호":
                            vals.append(i)
                        elif c == "확인":
                            vals.append("")
                        else:
                            vals.append(pick(head, r, src.get(c, [c])))
                    ws.append(vals)
                    for c in range(1, len(F["cols"]) + 1):
                        ws.cell(ws.max_row, c).border = box
                    rows_out += 1
                    if a.form == "media":
                        tot_q += pages(pick(head, r, src["분량"]))
                        m = re.search(r"([\d.]+)\s*(KB|MB|GB|TB)?", pick(head, r, src["파일크기"]), re.I)
                        if m:
                            tot_size += float(m[1]) * {"KB": 1 / 1024, "MB": 1, "GB": 1024, "TB": 1024 * 1024}.get((m[2] or "MB").upper(), 1)
                    if a.form == "extend":
                        tag = f"{i}번 {pick(head, r, src['기록물철분류번호'])}"
                        for c in ("처리과기관코드", "기록물철분류번호", "기록물철(건) 제목", "생산연도", "이관희망시기", "사유"):
                            if not pick(head, r, src[c]):
                                missing.append(["행", tag, f"{c} 비어 있음"])
                        y, h = re.search(r"\d{4}", pick(head, r, src["생산연도"])), re.search(r"\d{4}", pick(head, r, src["이관희망시기"]))
                        if y and h and int(h[0]) <= int(y[0]):
                            warn.append(["행", tag, "이관희망시기가 생산연도보다 늦지 않음"])
                        if len(re.sub(r"\s", "", pick(head, r, src["사유"]))) and len(re.sub(r"\s", "", pick(head, r, src["사유"]))) < 8:
                            warn.append(["행", tag, "사유가 짧음 — 업무 활용 등 구체적 사유를 적음"])
                    if a.form == "carry":
                        ret = re.sub(r"\s", "", pick(head, r, src["보존기간"]))
                        if ret.startswith(("영구", "준영구", "30년")):
                            warn.append(["행", f"{i}번 {pick(head, r, src['분류번호'])}", f"보존기간 {ret} — 서고 밖 반출은 원칙 금지, 업무 사용 반출은 이력 관리, 기관 밖 반출은 기관장 승인(시행규칙 제32조)"])
                if a.form == "extend":
                    ws.append([])
                    ws.append(["* 문서등록번호는 '건' 단위로 이관하는 경우에만 적는다(서식 주석)."])
                if a.form == "media":
                    ws.append([])
                    ws.append(["수록량(합계)", tot_q, "전체파일크기(MB)", round(tot_size, 1)])
    else:
        ws.append([])
        if a.form in TO:
            verb = {"limited": "비공개기록물의 제한적 열람을", "review": "비공개 결정에 대한 재심의를", "hold": "폐기 금지를", "hold-release": "폐기 금지 해제를", "hold-extend": "폐기 금지 연장을", "annotate": "부기 또는 정정을"}[a.form]
            if a.form in ("limited", "review"):
                ws.append(["          년      월      일"])
                ws.append(["신청인                (서명 또는 인)"])
                ws.append(["대리인                (서명 또는 인)"])
            else:
                ws.append([f"위와 같이 기록물의 {verb} 요청하오니 {'처리' if a.form == 'annotate' else '조치'}하여 주시기 바랍니다."])
                ws.append(["          년      월      일"])
                ws.append([f"요청기관명 {info.get('기관명', '')}  (직인)"])
            ws.append([TO[a.form]])
            ws.append([FOOT[a.form]])
            if a.form == "annotate":
                ws.append([])
                ws.append(["처리 방법(받은 기록물관리기관 참고) — 국가기록원 훈령 제224호 「기록물의 부기 또는 정정 처리 기준」 제2조"])
                ws.append(["① 소관 법령에 처리 방법이 있으면 그 방법으로 부기·정정"])
                ws.append(["② 없으면 「행정업무의 운영 및 혁신에 관한 규정」 제17조 준용 — 결재받은 문서는 재작성해 결재. "
                           "종이문서의 명백한 오류 등 경미한 사항은 원안 글자를 알 수 있게 글자 중앙에 가로 두 선을 긋고 고친 사람이 서명·날인(같은 규정 시행규칙 제14조)"])
                ws.append(["③ 보존매체에 수록된 기록물이면 보존매체에도 부기·정정 내용 반영"])
                ws.append(["※ 훈령 제1조는 '시행령 제62조의2'를 근거로 적었으나 현행 시행령에서 부기·정정은 제62조의3이다(제62조의2는 보존·복원 지원, law.go.kr 2026-11-20 확인)"])
                ws.append(["※ 훈령 제3조 재검토 기한(2026. 9. 20.)이 지났으므로 개정 여부를 국가기록원 소관법령 게시판에서 확인"])
    ws.append([])
    ws.append([f"근거: {F['basis']}"])
    ws.column_dimensions["A"].width = 24
    for col, w in zip("BCDEFGHI", (40, 22, 40, 18, 18, 12, 12, 20)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="center")
    chk = wb.create_sheet("빈칸 점검")
    chk.append(["구분", "항목", "내용"])
    for r in missing + warn:
        chk.append(r)
    if not missing and not warn:
        chk.append(["", "", "빈칸·경고 없음"])
    out = a.out or f"산출/별지{F['no']}_{F['title'].replace(' ', '')}.xlsx"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)
    print(f"별지 {F['no']} {F['title']} · 표 {rows_out}행 · 빈칸 {len(missing)} · 경고 {len(warn)} → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
