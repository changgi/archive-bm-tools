<div align="center">

<img src=".github/assets/banner.svg" alt="archive-bm-tools — 기록관리 담당자를 위한 Claude 스킬 28종" width="100%">

**공공기관 기록관리 담당자를 위한 Claude 플러그인 마켓플레이스**

생산현황 통보부터 이관·보존·공개·평가·폐기까지 —<br>
반복되는 점검과 서식 작성은 스크립트가, 판단과 제출은 사람이.

![version](https://img.shields.io/badge/records--skills-v2.12.1-1f6feb)
![skills](https://img.shields.io/badge/스킬-28종-2da44e)
![tests](https://img.shields.io/badge/시험-48%2F48%20통과-2da44e)
![python](https://img.shields.io/badge/Python-3.10%2B-3776ab?logo=python&logoColor=white)
![law](https://img.shields.io/badge/근거-시행령%202026.8.25.-8250df)
![license](https://img.shields.io/badge/코드-MIT-lightgrey)
![docs](https://img.shields.io/badge/문서-CC%20BY%204.0-lightgrey)

[빠른 시작](#-빠른-시작) · [스킬 28종](#-스킬-28종) · [안전장치](#-안전장치) · [근거](#-근거-법령과-자료) · [시험](#-시험)

</div>

---

## ✨ 무엇을 하나요

기록관리 담당자의 일 년은 **반복 점검**과 **서식 채우기**로 가득합니다. 등록정보를 부서별·보존기간별로 세고, 회의록 구성 요소를 하나하나 확인하고, 이관 목록과 상자를 맞춰 보고, 보존기간이 끝난 기록물로 평가심의서를 씁니다.

이 플러그인을 설치하면 Claude Code에게 **"올해 생산현황 통보서 만들어줘"** 라고 말하는 것만으로 이 일을 시작할 수 있습니다. 병원의 검사실이 혈액을 분석하고 의사가 진단하듯, **스크립트는 점검·후보·초안을 만들고 담당자는 확인하고 결정합니다.**

| 이전 | 이후 |
|---|---|
| 엑셀 피벗으로 부서×보존기간 집계, 서식 1~11에 손으로 옮겨 적기 | 입력 폴더 하나 → 통보서식 서식 1~11 자동 채움 + `통보전_확인.md` |
| 이관 목록과 실물 상자를 눈으로 대조 | 빠진 것·남는 것·이름 다른 것을 시트로 |
| 비공개 기록물 중 재분류 대상 찾기 | 올해 심의 대상과 사유 결함 후보를 자동 추출 |
| PDF/A 변환 후 "됐겠지" | veraPDF로 적합·부적합(조항)까지 확인 |

---

## 🧭 기록물 생애주기와 스킬

```mermaid
flowchart LR
    A["📝 생산·등록<br/>10종"] --> B["📦 이관·인수<br/>3종"]
    B --> C["🛡️ 보존<br/>4종"]
    C --> D["🔓 공개·활용<br/>3종"]
    D --> E["⚖️ 평가·폐기<br/>3종"]
    N["🌐 새 유형<br/>기록 3종"] -.-> A
    R["📊 기관평가<br/>보고 2종"] -.-> E
```

---

## 🚀 빠른 시작

> 비공개 저장소입니다. 이 저장소에 접근할 수 있는 GitHub 계정으로 로그인해 있어야 합니다.

**1. 설치** (Claude Code 안에서)

```text
/plugin marketplace add changgi/archive-bm-tools
/plugin install records-skills@changgi-archive-bm-tools
```

**2. 기관 폴더 이름 맞추기** (선택)

```text
/plugin configure records-skills@changgi-archive-bm-tools
```

비워 두면 기본값을 씁니다 — 비밀 폴더 `비밀,secret,classified`, 원본 폴더 `원본,보존,originals,preserved`.

**3. 작업 폴더 준비** (한 번만)

`plugins/records-skills/templates/CLAUDE.md`와 `templates/.claude/settings.json`을 기관 작업 폴더에 복사합니다. 플러그인은 거부 규칙(permissions)을 대신 걸 수 없어서 이 복사가 필요합니다.

**4. 말로 시작하기**

```text
입력 폴더에 등록정보 목록이랑 회의록 넣어 놨어. 올해 생산현황 통보서 만들어줘.
이관 목록.xlsx랑 상자 목록 대조해서 빠진 거 찾아줘.
비공개 기록물 목록에서 올해 재분류 심의 대상 뽑아줘.
이 폴더 문서들 PDF/A로 바꾸고 veraPDF로 확인해줘.
```

**업데이트:** `/plugin marketplace update changgi-archive-bm-tools`

---

## 🧰 스킬 28종

<details open>
<summary><b>📝 생산·등록 (10)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [production-status-report](plugins/records-skills/skills/production-status-report) | 생산현황 통보서식 서식 1~11 채우기, 입력 폴더 일괄 실행(`run_all.py`) |
| [records-register-audit](plugins/records-skills/skills/records-register-audit) | 정리 기간 등록정보와 실물 점검표 대조(시행규칙 제14조) |
| [records-qa-gate](plugins/records-skills/skills/records-qa-gate) | 내보내기 전 목록·메타데이터·이관 패키지 품질 관문, 열 이름 맞춤(`header_check`) |
| [metadata-autofill](plugins/records-skills/skills/metadata-autofill) | 본문에서 제목·생산기관·생산일자·문서번호 후보를 근거 위치와 함께 |
| [docno-extract](plugins/records-skills/skills/docno-extract) | 문서번호 추출, 물결표(~) 범위를 낱개로 펼치기 |
| [task-suggest](plugins/records-skills/skills/task-suggest) | 제목으로 공통업무 단위과제·보존기간 후보 |
| [records-schedule-check](plugins/records-skills/skills/records-schedule-check) | 기록관리기준표 관리항목 점검(NAK 4:2025) |
| [typed-records-check](plugins/records-skills/skills/typed-records-check) | 회의록 구성 요소 8가지, 시청각 생산정보 |
| [publications-check](plugins/records-skills/skills/publications-check) | 발간등록번호·발간주기·송부 책자 3부 |
| [admin-artifacts-check](plugins/records-skills/skills/admin-artifacts-check) | 행정박물 기본정보 11항목·유형 7·재질 13종·이관 기한(별표 4·5) |

</details>

<details>
<summary><b>📦 이관·인수 (3)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [transfer-reconcile](plugins/records-skills/skills/transfer-reconcile) | 이관 목록과 파일·상자 대조, 통합(전자+비전자) 이관 붙임 후보 |
| [transfer-triage](plugins/records-skills/skills/transfer-triage) | 인수·기증 폴더 첫 진단 — 포맷 분포·중복·위험 파일 |
| [format-identify](plugins/records-skills/skills/format-identify) | 내용으로 실제 포맷 식별, 확장자 불일치 |

</details>

<details>
<summary><b>🛡️ 보존 (4)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [preservation-convert](plugins/records-skills/skills/preservation-convert) | PDF/A-1b 변환, veraPDF 검증, 날짜 맞춤 |
| [preservation-inspection](plugins/records-skills/skills/preservation-inspection) | 정수점검·상태검사 주기, 별지 제7호서식 작업지 |
| [fixity-check](plugins/records-skills/skills/fixity-check) | SHA-256 기준선과 정기 대조 |
| [scan-qc](plugins/records-skills/skills/scan-qc) | 스캔 이미지 DPI·기울기·빈 면·밝기 |

</details>

<details>
<summary><b>🔓 공개·활용 (3)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [access-request](plugins/records-skills/skills/access-request) | 정보공개 결정·연장·이의신청 기한(공휴일 반영) |
| [disclosure-reclass](plugins/records-skills/skills/disclosure-reclass) | 공개 재분류 심의 대상, 비공개 사유 결함, 이관 전 재분류 |
| [pii-redact](plugins/records-skills/skills/pii-redact) | 주민등록번호·전화번호 등 개인정보 후보와 부분공개용 가림 목록 |

</details>

<details>
<summary><b>⚖️ 평가·폐기 (3)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [retention-review](plugins/records-skills/skills/retention-review) | 보존기간 대조 — 기관 기준표·공통업무 준칙·NAK 39 제정(안) |
| [disposal-dossier](plugins/records-skills/skills/disposal-dossier) | 기록물평가심의서·의견조회표(별지 제10호) |
| [email-appraisal](plugins/records-skills/skills/email-appraisal) | MBOX·EML 이메일 기록 평가 |

</details>

<details>
<summary><b>🌐 새 유형 기록 (3)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [web-records-check](plugins/records-skills/skills/web-records-check) | 누리집 수집 준비도(robots·사이트맵·리다이렉션), WARC 대조 |
| [dataset-schedule](plugins/records-skills/skills/dataset-schedule) | 행정정보 데이터세트 A·B①~④·C 유형 후보, 선정 협의 서식 |
| [classified-records-check](plugins/records-skills/skills/classified-records-check) | 비밀기록물 서식 10·11 — **담당자가 비밀 전용 장비에서만 실행** |

</details>

<details>
<summary><b>📊 기관평가·보고 (2)</b></summary>

| 스킬 | 하는 일 |
|---|---|
| [evaluation-readiness](plugins/records-skills/skills/evaluation-readiness) | 기록관리 기관평가 지표별 증빙 점검 |
| [hwpx-report](plugins/records-skills/skills/hwpx-report) | 마크다운 → 한글(HWPX) 보고서, 기관 서식 채움 |

</details>

---

## 🔒 안전장치

```mermaid
flowchart LR
    U["👤 담당자 요청"] --> C["🤖 Claude"]
    C --> H{"🚧 훅 검사"}
    H -- "비밀 폴더" --> X["⛔ 차단"]
    H -- "원본 쓰기" --> X
    H -- "통과" --> S["⚙️ 스크립트"]
    S --> O["📄 점검·후보·초안<br/>+ 확인 목록"]
    O --> P["✅ 담당자 확인·제출"]
    S -.-> L["🧾 처리 기록"]
```

- **비밀기록물은 Claude가 열지 않습니다.** `비밀` 폴더의 읽기와 명령 실행을 훅이 막고, 비밀 관련 스킬은 담당자가 비밀 전용 장비에서 직접 실행합니다.
- **원본은 읽기만 합니다.** 결과는 언제나 새 파일로 씁니다. 산출 폴더가 비어 있지 않으면 덮어쓰지 않습니다.
- **결정은 사람이 합니다.** 폐기·이관·공개 결정과 제출은 스크립트가 하지 않습니다.
- **모든 처리를 기록합니다.** 어떤 도구가 언제 무엇을 했는지 로그로 남습니다.

---

## 📚 근거 법령과 자료

| 근거 | 확인 |
|---|---|
| 공공기록물 관리에 관한 법률·시행령(2026. 8. 25. 개정)·시행규칙(2026. 9. 11. 개정) | 국가법령정보센터(law.go.kr) 원문 |
| 「2026년 기록물관리 지침」 | 국가기록원(archives.go.kr) |
| 생산현황 통보서식, 공통업무 보존기간 준칙 | 국가기록원 공개 자료 |
| NAK 10:2026(v1.5) 확정 · NAK 17:2026(v1.4)안 · NAK 39:2026(v1.0) 제정(안) | 국가기록원 고시·예고 |

지침에 남은 개정 전 문구(예: 비밀기록물 이관 10년)는 **개정 법령과 표준을 따릅니다.** 차이는 [`guideline-2024-notes.md`](plugins/records-skills/skills/records-qa-gate/references/guideline-2024-notes.md)에 정리했습니다.

---

## 🧪 시험

```bash
python plugins/records-skills/tests/run_tests.py
```

48개(veraPDF가 PATH에 있으면 49개). 필요 패키지: `openpyxl` `pypdf` `pikepdf` `pillow` `numpy` (선택: `python-docx` `python-hwpx` `holidays`, LibreOffice). Python 3.11·3.12·3.13에서 확인했고, [GitHub Actions](.github/workflows/test.yml)로 3.10~3.13을 돌립니다.

---

## 📄 이용 조건

코드는 [MIT](LICENSE), 문서·설명은 CC BY 4.0입니다. 플러그인에 든 국가기록원 서식·준칙 표는 공공저작물(저작권법 제24조의2)로, 출처를 밝혀 씁니다.

<div align="center">

만든 사람 **Song Chang-gi** · [chgsong@naver.com](mailto:chgsong@naver.com) · [github.com/changgi](https://github.com/changgi)

자세한 변경 내역은 [records-skills README](plugins/records-skills/README.md)에 있습니다.

</div>
