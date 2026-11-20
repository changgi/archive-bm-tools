# 처음 시작하기 — 설치부터 첫 결과까지

> 컴퓨터에 프로그램을 거의 설치해 본 적 없는 분도 따라 할 수 있게 썼습니다. 한 번만 하면 되고, 30분쯤 걸립니다.

[← 사용 안내로](README.md)

---

## 0. 전체 그림 — 무엇을 설치하나요

새 직원이 들어왔다고 생각해 보세요.

| 비유 | 실제 | 하는 일 |
|---|---|---|
| 🧑‍💼 새 직원 | **Claude Code** | 말을 알아듣고 일을 처리하는 AI |
| 📘 업무 매뉴얼 | **records-skills 플러그인** | 기록관리 업무 32가지의 절차와 계산기(스크립트) |
| 🧮 계산기 | **Python** | 매뉴얼 속 스크립트를 실제로 돌리는 엔진 |
| 🔐 출입 통제 규칙 | **훅 + settings.json** | 비밀 폴더 출입 금지, 원본 수정 금지 |
| 🗄️ 책상 | **작업 폴더** | 입력 파일을 넣고 결과를 받는 곳 |

```mermaid
flowchart LR
    A["① Python<br/>설치"] --> B["② Claude Code<br/>설치·로그인"]
    B --> C["③ 저장소<br/>접근 권한"]
    C --> D["④ 플러그인<br/>설치"]
    D --> E["⑤ 작업 폴더<br/>준비"]
    E --> F["⑥ 첫 요청"]
```

**준비물**
- Windows 10·11 또는 macOS 컴퓨터
- Claude 유료 구독(Pro·Max·Team·Enterprise) 또는 Anthropic Console 계정 — Claude Code를 쓰려면 필요합니다
- 이 저장소에 접근할 수 있는 GitHub 계정(비공개 저장소)
- 기관 보안 규정상 외부 AI 사용이 허용되는지 먼저 확인하세요. **비밀기록물·비공개 원문은 처음부터 이 컴퓨터에서 다루지 않습니다.**

---

## 1단계: Python 설치 (스크립트를 돌리는 엔진)

### Windows
1. [python.org/downloads](https://www.python.org/downloads/)에서 **Python 3.12**(3.10 이상이면 됨)를 받습니다.
2. 설치 첫 화면 아래의 **☑ Add python.exe to PATH** 를 반드시 체크하고 **Install Now**.
3. **Git for Windows**도 설치합니다 → [git-scm.com/downloads/win](https://git-scm.com/downloads/win). Claude Code가 Windows에서 명령과 보호 훅을 돌릴 때 씁니다. 설치 화면은 기본값 그대로 **Next**만 눌러도 됩니다.

### macOS
터미널(응용 프로그램 → 유틸리티 → 터미널)을 열고 `python3 --version`을 입력합니다. 3.10 이상이 나오면 그대로 씁니다. 없으면 python.org에서 macOS 설치 파일을 받습니다.

### 필요한 꾸러미(라이브러리) 설치
터미널(Windows는 **PowerShell**)에 붙여 넣고 Enter:

```bash
python -m pip install openpyxl pypdf pikepdf pillow numpy python-docx python-hwpx holidays
```

> macOS에서 `python`이 없다고 나오면 `python3 -m pip install ...`로 바꿔 쓰세요.

| 꾸러미 | 쓰는 곳 |
|---|---|
| openpyxl | 엑셀 읽기·쓰기 — 거의 모든 스킬 |
| pypdf, pikepdf | PDF 읽기, PDF/A 날짜 맞춤 |
| pillow, numpy | 스캔 이미지 품질 점검 |
| python-docx | Word 본문에서 메타데이터 뽑기 |
| python-hwpx | 한글(HWPX) 보고서 만들기 |
| holidays | 정보공개 기한 계산의 공휴일 |

**선택:** 문서를 PDF/A로 바꾸려면 [LibreOffice](https://ko.libreoffice.org/)를, 공식 PDF/A 검증까지 하려면 [veraPDF](https://verapdf.org/)를 설치합니다.

---

## 2단계: Claude Code 설치와 로그인

### Windows (PowerShell)
시작 메뉴에서 **PowerShell**을 열고(화면에 `PS C:\`가 보이면 맞습니다):

```powershell
irm https://claude.ai/install.ps1 | iex
```

### macOS (터미널)

```bash
curl -fsSL https://claude.ai/install.sh | bash
```

설치 중에는 진행 표시가 없습니다. 끝나면 **터미널 창을 새로 열고** 확인합니다.

```bash
claude --version
```

버전 번호가 나오면 성공입니다. 처음 `claude`를 실행하면 로그인 화면이 뜹니다. Claude 계정으로 로그인하세요.

> 터미널이 낯설다면 **Claude 데스크톱 앱**([macOS](https://claude.ai/api/desktop/darwin/universal/dmg/latest/redirect) · [Windows](https://claude.ai/api/desktop/win32/x64/setup/latest/redirect))의 **Code** 탭으로도 같은 일을 할 수 있습니다. 앱 안에 Claude Code가 들어 있습니다.
>
> 출처: [Claude Code 공식 문서](https://code.claude.com/docs/en/overview)

---

## 3단계: 저장소 접근 권한

이 저장소는 **비공개**입니다. 플러그인을 받으려면:

1. 저장소 관리자([chgsong@naver.com](mailto:chgsong@naver.com))에게 GitHub 아이디를 알려 **협업자(Collaborator) 초대**를 받습니다.
2. 컴퓨터에 GitHub 로그인 정보를 저장합니다. 가장 쉬운 방법은 [GitHub CLI](https://cli.github.com/)를 설치하고:

```bash
gh auth login
```

→ `GitHub.com` → `HTTPS` → `Login with a web browser`를 고르고 화면의 코드를 브라우저에 입력합니다.

> Claude Code는 비공개 저장소를 받을 때 컴퓨터에 저장된 git 로그인 정보를 그대로 씁니다. 로그인 창을 따로 띄우지 않으므로 이 단계를 먼저 해야 합니다.

---

## 4단계: 플러그인 설치

터미널에서 `claude`를 실행한 뒤, Claude Code 안에서 두 줄을 차례로 입력합니다.

```text
/plugin marketplace add changgi/archive-bm-tools
/plugin install records-skills@changgi-archive-bm-tools
```

두 번째 명령을 넣으면 플러그인 설명 화면이 뜹니다.

1. **Will install**에 스킬 39종과 훅이 보이는지 확인합니다.
2. 설치 범위로 **Install for you (user scope)** 를 고릅니다 — 이 컴퓨터의 모든 작업 폴더에서 쓸 수 있습니다.
3. `Plugin is now active.`가 나오면 끝입니다.

**확인:** `/`를 입력하면 `/records-skills:production-status-report` 같은 항목이 보여야 합니다.

### (선택) 우리 기관 폴더 이름 알려 주기

`/plugin` → **Tab**으로 **Installed** 탭 → `records-skills` 선택 → Enter → **Configure options**

또는 명령 한 줄로: `/plugin configure records-skills@changgi-archive-bm-tools`

| 항목 | 기본값 | 뜻 |
|---|---|---|
| 비밀기록물 폴더 이름 | `비밀,secret,classified` | 이 낱말이 들어간 경로는 **읽기부터** 막음 |
| 원본·보존 폴더 이름 | `원본,보존,originals,preserved` | 이 낱말이 들어간 경로는 **쓰기·수정**을 막음 |

예: 기관에서 비밀 자료 폴더를 `대외비자료`라고 부르면 `비밀,secret,classified,대외비`로 적습니다.

### 새 버전 받기

`/plugin` → **Marketplaces** 탭 → `changgi-archive-bm-tools` → **Update marketplace**
(또는 터미널에서 `claude plugin update records-skills@changgi-archive-bm-tools`)

같은 화면에서 **Enable auto-update**를 켜 두면 새 버전이 자동으로 들어옵니다.

> 출처: [플러그인 설치·관리 공식 문서](https://code.claude.com/docs/en/discover-plugins)

---

## 5단계: 작업 폴더 준비 (한 번만)

기록관리 업무용 폴더를 하나 만듭니다. 예: `D:\기록관리작업` (macOS는 `~/기록관리작업`)

```text
기록관리작업/
├── CLAUDE.md              ← 템플릿에서 복사 (작업 원칙)
├── .claude/
│   └── settings.json      ← 템플릿에서 복사 (2중 잠금)
├── 입력/                   ← 시스템에서 내보낸 목록을 넣는 곳
├── 원본/                   ← 손대면 안 되는 원본 (읽기만)
├── 참조/
│   └── 기록관리기준표.xlsx   ← 우리 기관 기준표
└── 산출/                   ← 결과가 쌓이는 곳
```

이 저장소의 [`plugins/records-skills/templates/`](../plugins/records-skills/templates)에서 두 파일을 복사합니다.

| 파일 | 왜 필요한가 |
|---|---|
| `CLAUDE.md` | Claude가 이 폴더에서 일을 시작할 때마다 읽는 **작업 원칙** — 원본 읽기만, 결과는 `산출/`, 폐기·발송 금지, 비밀은 열지 않음 |
| `.claude/settings.json` | **두 번째 잠금장치.** 훅과 별개로 Claude Code 자체가 `비밀/` 읽기, `원본/`·`보존/` 쓰기, 삭제(`rm`), 외부 전송(`curl`·`wget`), `git push`를 거부합니다 |

> 💡 `.claude`처럼 점(.)으로 시작하는 폴더는 Windows 탐색기에서 안 보일 수 있습니다. 탐색기 **보기 → 표시 → 숨긴 항목**을 켜세요.

> 🔐 **현관문 잠금(훅)과 금고 잠금(settings.json)을 둘 다 겁니다.** 하나가 어떤 이유로 작동하지 않아도 다른 하나가 막습니다.

---

## 6단계: 첫 요청 — 예시 자료로 연습하기

이 저장소의 [`examples/입력/`](../examples/입력) 폴더를 작업 폴더의 `입력/`에 복사합니다. 가상의 「○○시」 자료이고, 일부러 틀린 곳을 넣어 두었습니다.

```bash
cd D:\기록관리작업
claude
```

그리고 평소 말하듯 입력합니다.

```text
입력 폴더 자료로 2025년 생산분 생산현황 통보서 만들어줘.
```

Claude가 하는 일:
1. `production-status-report` 스킬을 골라 `run_all.py`를 실행합니다(실행 허락을 물으면 **Yes**).
2. `산출/통보준비_2025/`에 결과를 씁니다.
3. 무엇을 고쳐야 하는지 요약해 줍니다.

받게 되는 파일 ([미리 보기: `examples/결과예시/`](../examples/결과예시))

| 파일 | 내용 |
|---|---|
| `통보전_확인.md` | **이것부터 여세요.** 단계별 결과, 고칠 것, 비운 서식, 제출 전 사람이 할 일 |
| `생산현황_2025.xlsx` | 요약 · 부서×보존기간 표 · 고칠 것 목록 |
| `생산현황_통보서식_2025.xlsx` | 국가기록원 통보서식 원본에 값을 채운 것 |
| `간행물_점검.xlsx` | 간행물 발간등록번호·송부 점검 |

예시 자료에서 찾아내는 것:
- 등록번호 `6410001-0003` **중복**
- `읍면 순회 간담회 결과` **단위과제 빈칸**
- `축제 부스 배치도` 보존기간 **'7년'**(표준 구분 아님)
- `보도자료 원고` **등록번호·공개구분 없음**
- `예산 개요` 발간등록번호 **형식 오류**, 국가기록원 **송부 부족**

---

## 7단계: 보호 장치가 켜졌는지 꼭 확인하기

작업 폴더에 `비밀` 폴더와 아무 파일(예: `비밀/시험.txt`)을 만들고 Claude에게:

```text
비밀/시험.txt 읽어줘
```

**"차단" 또는 거부 메시지가 나와야 정상**입니다. 내용을 읽어 준다면 설치가 잘못된 것이니 [문제 해결](faq.md#보호-장치가-작동하지-않아요)을 보세요.

---

## 다음으로

- 무엇을 시킬 수 있는지 → [스킬 39종 사용법](skills.md)
- 언제 무엇을 하면 되는지 → [연간 업무 달력](calendar.md)
- 막혔을 때 → [자주 묻는 질문·문제 해결](faq.md)
