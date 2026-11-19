# records-skills (Claude 플러그인)

설치: `claude plugin marketplace add changgi/archive-bm-tools` → `claude plugin install records-skills@changgi-archive-bm-tools` (비공개 저장소라 저장소 접근 권한이 있는 GitHub 계정으로 로그인해 있어야 한다)

기관 작업 폴더에서 한 번: `templates/CLAUDE.md`와 `templates/.claude/settings.json`을 작업 폴더에 복사한다. 플러그인은 거부 규칙(permissions)을 대신 걸 수 없으므로 이 복사가 필요하다. 훅(비밀·원본 보호, 처리 기록)은 플러그인이 건다.

이용 조건: 코드 MIT, 문서·설명 CC BY 4.0.

---

# records-skills — 기록관리 AI 스타터 킷 (v2.12.1)

기록관리 AI 에코시스템 구조도의 **1단계(사고를 막는 것)** 전체와 **2단계 스킬 4종(공개 재분류·보존기간 검토·스캔 품질·폐기 심의자료)**과 **3단계 보존포맷 변환**, **인수 일괄 점검·이메일 평가**를 실제로 돌아가게 만든 묶음입니다.

v2.12.1 변경: 저장소 게시 준비 — 파이썬 3.10·3.11에서 문법 오류가 나던 f-string(식 안의 역슬래시, fill_official_form.py) 고침, pyflakes 경고 7건 정리, 시험 파일 쓰기에 encoding='utf-8'(한국어 Windows cp949). 시험 48개.

v2.12 변경: veraPDF 1.30.4(공식 배포)로 preservation-convert 결과를 실제 검증 — 세 가지를 찾아 고침. ① 한글 파일명을 veraPDF(Java)가 못 읽어 '부적합'으로 잘못 기록되던 것 → 영문 임시 사본·UTF-8로 검증, '검증 오류'를 판정과 구분, 실패 조항 표시. ② 6.7.3 문서정보·XMP 날짜 불일치 → 자동 맞춤(--no-fix-dates). ③ 6.3.5 OTF(CFF) 한글 대체 글꼴의 CharSet 누락 → 기본 점검에서 잡고 TrueType 한글 글꼴 안내. TrueType 한글 글꼴 + 날짜 맞춤이면 PASS 확인. 시험 48개(veraPDF 있으면 49개).

v2.11 변경: 스킬·스크립트의 지침 근거를 「2026년 기록물관리 지침」으로 바꿈 — 간행물(책자 3부·생략대상)·행정박물(재질 13종·기본정보 11개)·회의록 구성요소·시청각 내용요약(생산일자)·웹기록물(제46조의2)·비밀기록물이 2026년판에도 같음을 본문에서 확인. 개인정보가 포함된 기록물관리는 2026년판 부록 1(2024년판 부록 7)로 고침. 2026년판에 남은 개정 전 문구는 records-qa-gate/references/guideline-2024-notes.md에 따로 둠. 시험 48개.

v2.10 변경: 「2026년 기록물관리 지침」 평가·공개재분류 대조 — disclosure-reclass: 제1호 비공개는 근거 법령 조항(「…」·제○조) 없으면 '사유 결함', `--transfer`로 이관 전 공개여부 재분류 필요·생략 가능(최근 5년) 구분. disposal-dossier: 평가심의서·의견조회표에 '의견 작성 안내'(폐기·재책정·보류+사유, 검토사항 유관법령·업무활용·책정기준, 처리과 자체 폐기 금지). 시험 48개.

v2.9 변경: 「2026년 기록물관리 지침」의 새 절차 두 가지 — `dataset-schedule/scripts/dataset_screen.py`(행정정보 데이터세트 관리대상 A·B①~④·C 유형 후보와 '선정 협의 서식', 종합관리시스템 darm.go.kr 안내)와 `transfer-reconcile/scripts/integrated_transfer.py`(통합(전자+비전자) 이관: ~2015년 비전자문서 건, 붙임(별첨) 표기 예시로 비전자 붙임(별송) 후보, 분리등록 첨부물, 건·붙임 색인목록 서식). 오프라인 이관 인수·인계 확인서 항목을 transfer-reconcile에 적음. 시험 47개.

v2.8 변경: 2026년 고시·지침 반영 — 「2026년 기록물관리 지침」, NAK 39:2026(v1.0) 보존기간 책정표준: 공통업무 제정(안)(고시 제2026-8호 예고), NAK 17:2026(v1.4) 비밀기록물 개정(안), NAK 10:2026(v1.5) 기록관 운영(고시 제2026-7호 확정)을 국가기록원에서 받아 대조. retention-review `--common-basis nak39안`(부속서 A~F 1,333행, 예고안 자체 불일치 4건 메모)과 현행 준칙 차이표. 2026 지침에 남은 개정 전 문구(비밀 이관 10년, 비치기록물 10년, 패키지 변환 11년)는 법령·개정 표준을 따르도록 메모. preservation-convert에 표준보존패키지 변환 시기(기산일+6년 전, NAK 10:2026). evaluation-readiness: 중앙·특별지방·국공립대는 2027년 격년 평가. 시험 46개.

v2.7 변경: `preservation-inspection/scripts/inspection_form.py` — 시행규칙 별지 제7호서식(2026.9.11. 개정, law.go.kr에서 칸 확인) 「보존기록물 정수점검ㆍ상태검사서」 작업지를 올해 대상만 채워 만들고(`form`), 작성한 작업지의 점검·검사 연도와 검사후등급을 보존목록 사본에 옮기며 3등급 항목을 사후조치 후보로 적는다(`apply`). 시험 45개.

v2.6 변경: `records-qa-gate/scripts/header_check.py` — 기관 시스템 내보내기 목록의 머리글만 읽어 킷 스크립트(ALIAS를 가진 것 모두)가 어느 열을 알아보는지, 띄어쓰기·괄호·별표·년도/연도 차이로 못 알아보는 열('비슷함')과 겹치는 열('후보')을 알리고, 머리글만 바꾼 사본을 만든다(`--fix`, `--map`). `run_all.py`는 '열 이름 확인' 절을 쓰고 `--fix-headers`로 그 사본을 쓴다. 지침 등록사항·서식 2-1의 공식 항목명(생산부서(처리과)명·생산시작년도·기록물철제목·단위업무명/단위과제명 등)을 별칭에 더함. 시험 44개.

v2.5 변경: `production-status-report/scripts/run_all.py` — 입력 폴더 하나로 생산현황·시청각·회의록·행정박물·간행물 점검을 돌리고 통보서식 서식 1~9를 채운 뒤 `통보전_확인.md`(단계별 실행·고칠 것·비운 서식)를 만든다. 비밀 파일은 건너뛰고 이름도 적지 않으며, 산출 폴더 덮어쓰기는 거부. 시험 43개.

v2.4 변경: `fill_official_form.py --bodies 회의체목록.xlsx`로 서식 9(회의록 생산목록)를 채운다 — 회의체마다 한 줄, 미개최도 0, 회의를 회의체로 묶고 회의호수로 서식 8 분류, 약칭 근거명칭·구성일 형식·서면 무록·지정회의(시행령 제18조②) 속기록 부족·성명 기재 점검. 이로써 생산현황 통보서식 서식 1~11을 모두 도구로 채운다(제출은 사람). 시험 42개.

v2.3 변경: 새 스킬 `classified-records-check`(비밀기록물 예고문·보호기간·보존기간·사본 제외, 이관 시기를 개정 시행령 제68조②대로, 서식 10 생산·재분류 통계와 서식 11 생산 후 3년 지난 목록). Claude는 비밀 목록을 열지 않고(훅이 `비밀` 폴더 실행까지 차단) 담당자가 비밀 전용 장비에서 실행한다. `fill_official_form.py --secret`. 이로써 통보서식 서식 1~11 중 9(회의록 생산목록)만 사람이 쓴다. 시험 41개.

v2.2 변경: 새 스킬 `web-records-check`(누리집 robots.txt 전체 차단·사이트맵·자바스크립트 링크·자동 리다이렉션·응답 오류, 기관 WARC의 Block/Payload-Digest 대조, SNS·블로그 채널 목록). 시행령 제2조 10호·제46조의2(2025.5.20. 신설) 확인. 시험 40개.

v2.1 변경: 새 스킬 `publications-check`(발간등록번호 형식·발간주기 01/02~15·단행본 번호 중복·생략대상 후보·국가기록원 송부 책자 3부+PDF, 서식 7 생산량·송부량을 종 수로)와 `admin-artifacts-check`(행정박물 관리카드 기본정보 11항목·등록번호·별표 4 유형 7·재질 13종·관인 폐기공고번호·공직자선물 4항목, 이관대상의 이관 기한을 시행령 별표 5대로, 서식 5·6 값). `fill_official_form.py`에 `--artifacts`(서식 5·6)·`--pubs`(서식 7). 시행령 별표 4·5와 제57조는 law.go.kr에서 2026-11-19 확인. 시험 39개.

v2.0 변경: 국가기록원 「2024년 기록물관리 지침」·「기록물관리 지침(공통 매뉴얼)」(2022) 반영. 새 스킬 `typed-records-check`(회의록 구성 요소 8가지·생산의무 회의 대조, 시청각 등록 생산정보·쪽수 점검). pii-redact·disposal-dossier에 '등록된 기록물은 개인정보 보유기간이 지나도 보존기간까지·평가 거쳐 폐기' 규칙. 지침과 현행 법령이 다른 곳 정리(`records-qa-gate/references/guideline-2024-notes.md`). 시험 38개.

v1.9 변경: 공식 생산현황 통보서식의 서식 3·4(시청각, `--av`)와 서식 8(회의록 생산현황, `--meetings`)도 채움, 생산시스템구분 코드(01~06) 점검, 시험 37개. 2026년 평가지표 설명자료는 포털에 공공기관·교육(지원)청판만 있음을 확인.

v1.8 변경: 국가기록포털(www.archives.go.kr) 업무안내ㆍ자료 22쪽·첨부 209건·게시판 47건·발간자료 842건 전수 조사, 공개 파일 12건을 받아 반영(목록: `skills/records-qa-gate/references/archives-go-kr-inventory-2026-11-20.md`).
- 새 스킬 3종: `task-suggest`(공통업무 보존기간 준칙 1,329개 단위과제로 단위과제·보존기간 후보), `evaluation-readiness`(2026년('25년 실적) 기록관리 평가지표별 증빙 점검), `dataset-schedule`(행정정보 데이터세트 관리기준표 초안)
- production-status-report: 국가기록원 「2024년도 기록물 생산현황 통보서식」 원본에 서식 1~11(3~11은 입력 파일을 줄 때)를 채우는 `fill_official_form.py`(서식은 assets에 그대로)
- retention-review: `--common 중앙행정기관|지방자치단체|공공기관|교육(행정)기관|학교|대학`으로 공통업무 준칙 대조
- 시험 36개

v1.7 변경: 국가법령정보센터(www.law.go.kr) 원문 대조 반영 — 시행령(2026. 8. 25. 실질 개정)·시행규칙(2026. 9. 11.)·정보공개법·국가기록원고시 제2025-7호.
- 새 스킬 4종: `records-register-audit`(시행규칙 제14조 정리 점검), `hwpx-report`(마크다운→HWPX, 기관 서식 채움), `preservation-inspection`(별표 14 정수점검·상태검사 주기), `records-schedule-check`(NAK 4:2025 기록관리기준표 관리항목 점검·고시 대상)
- 고친 것: 30년 이상 이관을 기산일부터 **5년** 경과 다음 연도로(시행령 제40조①, 개정 전 10년은 `--transfer-years 10`), 영구·준영구가 이관 대상에서 빠지던 버그
- 추가: 이관목록 별지 제2호서식, 평가심의서 별지 제10호서식 시트, 폐기 금지 목록 제외(제54조의4), 비공개 연장 신청 기한(제72조③), 전자화 이력(제29조③), 정보공개 20일 경과 이의신청·제3자 요청 시 30일 간격
- 참고자료: 시행령 별표 1 전문, 법령 확인 기록(`skills/records-qa-gate/references/law-check-2026-11-20.md`)
- 시험 32개

v1.6 변경: `access-request`(정보공개 결정·연장·공개실시·이의신청 기한을 공휴일 반영해 계산, 비공개 사유·연장 통지·제3자 의견 기간 점검), `production-status-report`(전년도 생산현황 부서×보존기간 집계, 통보 전 고칠 것) 추가, 시험 25개.

v1.5 변경: `metadata-autofill`(본문에서 제목·생산기관·생산일자·문서번호·수신자 후보를 근거 위치와 함께 제안, 기존 목록 빈칸·불일치 표시) 추가, Claude 플러그인 배치 도구 `build_plugin.py` 추가(훅이 플러그인 설정값을 읽음), 시험 23개.

v1.4 변경: `preservation-convert`(PDF/A-1b 변환·점검·변환이력, NAK 30:2022), `email-appraisal`(MBOX·EML 목록·상대·첨부·개인정보·비밀 표시) 추가, scan-qc에 마스터 포맷 열 추가, 시험 22개.
v1.3 변경: `disposal-dossier`(평가심의서·부서별 의견조회·HWPX 안건), `transfer-triage`(포맷·중복·개인정보 한 장 보고서) 추가, 시험 20개. python-hwpx는 Apache-2.0으로 확인.
v1.2 변경: `retention-review`(기준표 대조·만료·이관 대상), `scan-qc`(DPI·기울기·빈 면) 추가, 시험 18개.
v1.1 변경: 비밀기록물 차단(훅·검수 게이트·권한 규칙), `disclosure-reclass` 추가.
Claude Code 저장소 최상위에 그대로 두면 하네스(권한·Hooks)가 걸리고, `skills/`의 스킬을 Claude가 골라 씁니다.

## 들어 있는 것

| 구분 | 경로 | 하는 일 |
| --- | --- | --- |
| 지침 | `CLAUDE.md` | 용어·원칙·끝나면 검수 |
| 권한 | `.claude/settings.json` | 원본·보존 폴더 쓰기 금지, rm·curl·git push 금지, 패키징은 승인 후 |
| 훅 | `hooks/pre_tool_guard.py` | 실행 직전 검사: 원본 쓰기·삭제·외부 전송이면 차단, `비밀/` 폴더는 읽기·검색까지 차단(종료 코드 2) |
| 훅 | `hooks/post_tool_log.py` | 처리 이력을 `logs/processing_log.jsonl`에 KST 시각·SHA-256과 함께 기록 |
| 스킬 | `skills/records-qa-gate` | 비밀 차단 + 4대 속성 검수 게이트 — 실패 0건일 때만 내보냄 |
| 스킬 | `skills/retention-review` | 기준표 대조(짧음·김·없음·허용 값), 만료 지남·올해·내년, 30년 이상 이관 대상 — 개정안 5종은 `--scheme draft` |
| 스킬 | `skills/disposal-dossier` | 평가심의서(의견 칸 선택 목록)·생산부서별 의견조회표·심의안건.hwpx, 검토 신호(소송·감사·재산 등) |
| 스킬 | `skills/preservation-convert` | LibreOffice로 PDF/A-1b 변환 → PDF/A 선언·글꼴 임베딩·암호화 점검(+veraPDF 선택), 원본·결과 SHA-256 변환이력 |
| 스킬 | `skills/publications-check` | 간행물 목록 → 발간등록번호·발간주기·중복·생략대상 후보·송부(책자 3부+PDF), 서식 7 생산량·송부량(종) |
| 스킬 | `skills/admin-artifacts-check` | 행정박물 목록 → 관리카드 11항목·등록번호·유형·재질·관인/공직자선물 항목, 이관 기한(별표 5), 서식 5·6 |
| 스킬 | `skills/web-records-check` | 누리집 URL·WARC·채널 목록 → 웹 수집 준비(robots·사이트맵·동적 링크), WARC Digest, SNS 채널 |
| 스킬 | `skills/classified-records-check` | (담당자가 비밀 전용 장비에서 실행) 비밀 목록 → 예고문·보존·보호기간, 이관 시기, 서식 10·11 |
| 스킬 | `skills/typed-records-check` | 회의록 폴더 → 구성 요소 8가지·생산의무 회의 대조 / 시청각 등록 목록 → 유형·생산일자·장소·인물·쪽수 |
| 스킬 | `skills/task-suggest` | 제목 → 공통업무 보존기간 준칙(1,329개)·기관 기준표에서 단위과제·보존기간 후보 3개 |
| 스킬 | `skills/evaluation-readiness` | 증빙 목록 → 2026년 기록관리 평가지표별 준비·형식 결함·선택지표 규칙 |
| 스킬 | `skills/dataset-schedule` | 시스템 스키마 → 행정정보 데이터세트 관리기준표 6개 영역 초안 |
| 스킬 | `skills/records-register-audit` | 등록정보 ↔ 실물 점검표 → 미등록·실물 없음·쪽수·번호 표기·공개/비밀·상자 혼재(시행규칙 제14조) |
| 스킬 | `skills/records-schedule-check` | 기록관리기준표 → 관리항목 7개 점검(NAK 4:2025), 지난해와 비교한 고시 대상 |
| 스킬 | `skills/preservation-inspection` | 보존목록 → 정수점검·상태검사 다음 연도·지남(시행규칙 별표 14) |
| 스킬 | `skills/hwpx-report` | 마크다운 → HWPX, 기관 서식 {{자리표시}}·칸 채움, 스키마·누락 점검 |
| 스킬 | `skills/access-request` | 정보공개 청구 목록 → 결정기한(10일·연장 10일)·공개실시·이의신청(7일) 기한, 초과·임박, 비공개 사유 결함 — 청구인 정보는 옮기지 않음 |
| 스킬 | `skills/production-status-report` | 등록정보 목록 → 전년도 생산현황 집계(부서×보존기간·형태·유형), 통보 전 고칠 것 — 비밀은 수만 셈 |
| 스킬 | `skills/metadata-autofill` | txt·docx·hwpx·pdf 본문 → 메타데이터 후보 + 근거(파일·쪽·줄), 빈칸 채울 후보·불일치 — 값은 사람이 확정 |
| 스킬 | `skills/email-appraisal` | MBOX·EML → 메시지 목록·주고받은 상대·첨부(해시)·개인정보 후보·비밀 표시, 본문은 옮기지 않음 |
| 스킬 | `skills/transfer-triage` | 인수 폴더 한 장 보고서: 확장자 분포·포맷 위험·중복·개인정보 후보 |
| 스킬 | `skills/scan-qc` | 스캔 이미지 DPI·기울기·빈 면·밝기·화소 수 점검, 재스캔 후보 |
| 스킬 | `skills/disclosure-reclass` | 공개 재분류 후보(30년 경과·5년 주기·6호 유예·이력 없음·사유 결함·비밀 제외) |
| 스킬 | `skills/pii-redact` | 개인정보 후보 탐지(값은 가려서 기록) |
| 스킬 | `skills/transfer-reconcile` | 이관 목록 ↔ 파일 대조, BagIt 패키지 생성·검증 |
| 스킬 | `skills/docno-extract` | 문서번호 추출, `~` 범위 낱개로 펼치기 |
| 스킬 | `skills/fixity-check` | SHA-256 기준선·정기 대조 (AIC가 있으면 AIC 우선) |
| 스킬 | `skills/format-identify` | 내용 기반 포맷 식별(Siegfried·Magika·시그니처), 위험 표시 |
| 시험 | `tests/run_tests.py` | 가상 목록으로 전체 22개 시험 |

## 설치

```bash
pip install openpyxl pypdf pillow numpy   # 필수
pip install magika                    # 선택: 포맷 식별 정확도
pip install holidays                  # 선택: 정보공개 기한에 한국 공휴일 반영
pip install python-hwpx               # 선택: 심의안건 HWPX 생성 (Apache-2.0, Python 3.10+)
# 보존포맷 변환: LibreOffice 필요, veraPDF가 있으면 --verapdf로 공식 검증
# 선택: Siegfried(sf) 설치 시 PRONOM ID까지 기록 — https://github.com/richardlehane/siegfried
python tests/run_tests.py             # 38/38 통과 확인
python build_plugin.py ../plugin-build   # 플러그인 마켓플레이스 배치(게시 안 함)
claude plugin validate ../plugin-build  # 형식 검증
```

폴더 약속: 원본은 `원본/`, 보존본은 `보존/`, 비밀기록물은 `비밀/`(AI가 열지 않음), 모든 결과는 `산출/`. 폴더 이름은 환경변수 `RECORDS_PROTECTED`·`RECORDS_SECRET`로 바꿀 수 있습니다.

## 쓰는 순서 (이관 예)

1. `docno-extract`로 문서번호를 펼친다.
2. `transfer-reconcile`로 목록과 파일을 대조한다. 불일치가 있으면 멈춘다.
3. 일치하면 `make_bag.py`로 패키지를 만들고 `--verify`로 검증한다.
4. `pii-redact`로 공개 대상에 개인정보 후보가 없는지 본다.
5. `records-qa-gate`를 `--source`·`--manifest`와 함께 돌려 통과하면 사람에게 넘긴다.

## 쓰는 순서 (생산현황 통보 예)

1. 처리과에서 받은 등록정보·시청각·회의·행정박물·간행물 목록을 입력 폴더 하나에 모은다(비밀은 넣지 않는다). 처음 쓰는 내보내기 파일이면 `header_check.py`로 머리글부터 본다.
2. `python skills/production-status-report/scripts/run_all.py 입력폴더 --year 2025 -o 산출/통보준비_1차`
3. `통보전_확인.md`의 고칠 것을 처리과에 돌려보내고, 고친 목록으로 `-o 산출/통보준비_2차`처럼 새 폴더에 다시 돌린다.
4. 통보서식을 엑셀에서 열어 합계를 확인하고, 비밀 담당자가 따로 만든 서식 10·11과 함께 사람이 제출한다.

## 하지 않는 것

비밀기록물은 열지도, 요약하지도 않습니다. 비밀 구분 열에 값이 있거나 'Ⅱ급 비밀'·'대외비' 표시가 있으면 검수 게이트가 멈춥니다.


폐기 실행, 외부 게시, 공개 범위 변경, 메일 발송은 스크립트가 하지 않습니다. 목록·심의자료·초안까지만 만들고 결정은 사람이 합니다.

## 알아 둘 한계

- 개인정보 탐지는 패턴이 있는 정보만 찾습니다. 이름·주소·병력은 사람이 확인해야 합니다.
- 공개 재분류의 '30년 경과'는 '기준 연도 − 생산연도 > 30'으로 계산합니다. 기관 지침의 계산법과 맞는지 확인하세요.
- 2020년 10월 이후 발급 주민등록번호는 뒷자리 검증식이 맞지 않을 수 있어, 불일치도 후보로 남깁니다.
- 훅은 Claude Code의 PreToolUse·PostToolUse 입력(JSON: `tool_name`, `tool_input`) 기준으로 작성했습니다. 기관 PC에서 실제 세션으로 한 번 시험한 뒤 쓰세요.
- 보존포맷 변환의 '통과'는 PDF/A-1b 선언·글꼴 임베딩·암호화 없음을 본 기본 점검입니다. 공식 적합성은 veraPDF로 확인하세요. HWP/HWPX는 변환하지 않고 '한컴 변환 필요'로 남깁니다.
- 심의안건.hwpx는 python-hwpx로 만들고 다시 열어 내용을 확인했지만, 한글 프로그램에서의 서식은 직접 열어 확인하세요.
- `extract_docno.py`는 기존 스크립트를 스킬로 옮기며 새로 작성한 판입니다. 원본 스크립트가 있으면 같은 목록으로 결과를 대조하세요.

제작: Song Chang-gi · chgsong@naver.com · https://github.com/changgi
