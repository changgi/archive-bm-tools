# 자주 묻는 질문 · 문제 해결

[← 사용 안내로](README.md)

---

## 처음 쓸 때

### 프로그래밍을 몰라도 쓸 수 있나요?
네. 명령어를 몰라도 **평소 말하듯** 시키면 됩니다. "이관 목록 대조해줘", "올해 만료 목록 뽑아줘". 처음 설치만 [처음 시작하기](getting-started.md)를 따라 한 번 하면 됩니다.

### 비용이 드나요?
플러그인은 무료(MIT)입니다. Claude Code를 쓰려면 Claude 유료 구독(Pro·Max·Team·Enterprise) 또는 Anthropic Console 계정이 필요합니다.

### 우리 기관 시스템(RMS·온나라·CAMS)과 바로 연결되나요?
아니요. 시스템에서 **엑셀로 내보낸 파일**을 `입력/` 폴더에 넣어 씁니다. 결과도 파일로 받아 담당자가 시스템에 반영하거나 제출합니다.

### 결과를 그대로 제출해도 되나요?
**안 됩니다.** 결과는 점검·후보·초안입니다. `통보전_확인.md`의 '제출 전에 사람이 할 일'을 꼭 확인하세요. 법령상 판단과 결재는 사람이 합니다.

### 어떤 법령을 기준으로 하나요?
공공기록물법·시행령(2026. 8. 25. 개정)·시행규칙(2026. 9. 11. 개정), 「2026년 기록물관리 지침」, NAK 표준입니다. 지침에 개정 전 문구가 남은 곳은 **개정 법령을 따릅니다.** 예: 지침의 '비밀기록물 생산 후 10년 이관' → 개정 시행령 제68조②.

### 법령이 바뀌면요?
이슈 탭의 **「⚖️ 법령·지침 변경 알림」** 양식으로 알려 주세요. 고친 새 버전은 [처음 시작하기 4단계](getting-started.md#새-버전-받기)대로 받으면 됩니다.

---

## 설치 문제

### `claude`를 찾을 수 없다고 나와요
설치 후 **터미널 창을 새로** 열어야 합니다. 그래도 안 되면 설치 경로가 PATH에 없는 것입니다 → [공식 설치 문제 해결](https://code.claude.com/docs/en/troubleshoot-install).

### PowerShell에서 `'&&' is not a valid statement separator`
PowerShell에서는 `irm https://claude.ai/install.ps1 | iex`를 쓰세요. `&&`가 들어간 명령은 CMD용입니다.

### `/plugin marketplace add`에서 저장소를 못 찾는다고 나와요
비공개 저장소라서 생기는 일이 대부분입니다.
1. 관리자에게 협업자 초대를 받았는지, **초대 메일의 수락**까지 했는지 확인합니다.
2. 터미널에서 `gh auth login`으로 GitHub에 로그인합니다.
3. `gh repo view changgi/archive-bm-tools`가 열리면 권한은 정상입니다. Claude Code를 다시 시작해 재시도합니다.

### `/`를 눌러도 `records-skills:`가 안 보여요
`/plugin` → **Installed** 탭에서 `records-skills`가 켜져(enabled) 있는지 봅니다. **Errors** 탭에 이유가 나옵니다. 설치 직후라면 `/reload-plugins`를 실행합니다.

### `ModuleNotFoundError: No module named 'openpyxl'`
꾸러미가 설치되지 않았습니다.
```bash
python -m pip install openpyxl pypdf pikepdf pillow numpy python-docx python-hwpx holidays
```
Claude에게 "필요한 파이썬 꾸러미 설치해줘"라고 해도 됩니다.

---

## 보호 장치가 작동하지 않아요

`비밀/시험.txt 읽어줘`에 내용을 읽어 준다면:

1. **플러그인이 켜졌는지** — `/plugin` → Installed → `records-skills` enabled
2. **Python이 있는지** — 터미널에서 `python --version` (Windows는 `py --version`도). 3.10 이상이 나와야 합니다.
3. **Windows: Git for Windows가 있는지** — Claude Code는 Windows에서 훅을 Git Bash로 돌립니다. [git-scm.com](https://git-scm.com/downloads/win)에서 설치 후 Claude Code를 다시 엽니다.
4. **폴더 이름** — 기관 비밀 폴더 이름이 `비밀`·`secret`·`classified`가 아니면 **Configure options**에 넣습니다.
5. **3겹 잠금 확인** — 작업 폴더에 `.claude/settings.json`이 있는지 봅니다(숨김 폴더).
6. 그래도 안 되면 **사용을 멈추고** 이슈의 「🐞 오류 제보」로 알려 주세요.

반대로 **모든 작업이 막히고** "Python 3.8 이상을 찾지 못해"라고 나오면, Python을 설치(또는 PATH에 추가)한 뒤 Claude Code를 다시 여세요. 보호가 꺼지지 않게 일부러 막은 것입니다.

---

## 실행 문제

### "열 이름이 필요합니다" / 결과가 텅 비었어요
시스템마다 열 이름이 달라서입니다.
```text
입력/등록정보.xlsx 열 이름이 스킬에 맞는지 봐줘.
```
→ `없음`으로 나온 열의 이름을 예시대로 바꾸거나, 생산현황은 "열 이름 고쳐서 돌려줘"(`--fix-headers`)라고 하세요. 표준 열 이름은 [스킬 사용법 끝](skills.md#열-이름이-안-맞을-때)에 있습니다.

### "출력 폴더가 비어 있지 않음"
이전 결과를 덮어쓰지 않으려고 멈춘 것입니다. 새 폴더 이름을 주세요. 예: `산출/통보준비_2025_2차`

### 엑셀을 열어 둔 채로 돌렸더니 오류가 나요
Windows에서는 엑셀이 연 파일을 다른 프로그램이 못 읽을 수 있습니다. 엑셀을 닫고 다시 돌리세요. `~$`로 시작하는 임시 파일은 자동으로 건너뜁니다.

### 한글 파일명이 깨지거나 PDF/A 검증이 '검증 오류'예요
veraPDF(Java)는 한글 파일명을 잘 못 읽습니다. 스크립트가 영문 임시 사본으로 검증하므로, '검증 오류'는 veraPDF 설치·Java 문제일 가능성이 큽니다. `verapdf --version`이 되는지 확인하세요.

### PDF/A 변환이 6.3.5 조항 부적합이에요
한글 대체 글꼴이 OTF(CFF)라서입니다. **TrueType(.ttf) 한글 글꼴**(예: 나눔고딕, 맑은 고딕)을 설치하고 문서 글꼴을 그것으로 바꾸세요.

### HWPX가 한컴오피스에서 이상하게 보여요
아직 한컴오피스 화면 확인을 마치지 못한 부분입니다. 파일을 첨부해 「🐞 오류 제보」로 알려 주세요(실제 기관 내용은 지우고).

---

## 자료·보안

### 실제 기관 자료를 넣어도 되나요?
기관 보안 규정이 허용하는 범위에서만 쓰세요. Claude Code는 대화 내용(읽은 파일 일부 포함)을 Anthropic 서버로 보내 처리합니다. **비밀·비공개 원문, 민감한 개인정보는 넣지 마세요.** 처음에는 [`examples/`](../examples)의 가상 자료로 연습하세요.

### 오류를 제보할 때 무엇을 보내면 되나요?
**열 이름(머리글)과 가짜 값 몇 줄**이면 충분합니다. `header_check` 결과를 붙이면 가장 빠릅니다. 실제 목록·비밀기록물·개인정보는 절대 올리지 마세요.

### 처리 기록은 어디에 있나요?
작업 폴더의 `logs/processing_log.jsonl`입니다. 언제 어떤 파일을 만들었는지와 해시가 남습니다.
