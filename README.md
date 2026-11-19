# archive-bm-tools

공공기관 기록관리 담당자용 Claude 플러그인 마켓플레이스입니다.

| 플러그인 | 하는 일 |
|---|---|
| [records-skills](plugins/records-skills) | 기록관리 스킬 28종 — 생산현황 통보서식 서식 1~11 채우기, 이관 대조·통합(전자+비전자) 이관 준비, 보존기간 점검(공통업무 준칙·NAK 39 예고안), 공개 재분류 후보, 평가·폐기 심의자료, PDF/A 보존포맷 변환(veraPDF 확인), 행정정보 데이터세트, 웹기록물, 행정박물, 간행물, 정수점검·상태검사, 비밀기록물(담당자 전용 장비) 점검 — 와 비밀·원본 폴더 보호 훅 |

## 설치 (Claude Code)

비공개 저장소이므로 이 저장소에 접근할 수 있는 GitHub 계정으로 로그인해 있어야 합니다.

```
/plugin marketplace add changgi/archive-bm-tools
/plugin install records-skills@changgi-archive-bm-tools
```

새 버전: `/plugin marketplace update changgi-archive-bm-tools`

기관 작업 폴더에서 한 번: `plugins/records-skills/templates/CLAUDE.md`와 `templates/.claude/settings.json`을 작업 폴더에 복사합니다(플러그인은 거부 규칙을 대신 걸 수 없음). 훅(비밀·원본 보호, 처리 기록)은 플러그인이 겁니다.

## 근거와 원칙

- 공공기록물법·시행령(2026.8.25. 개정)·시행규칙(2026.9.11. 개정)은 국가법령정보센터(law.go.kr) 원문, 「2026년 기록물관리 지침」·NAK 표준·통보서식은 국가기록원(archives.go.kr) 공개 자료로 대조했습니다. 지침에 남은 개정 전 문구는 법령을 따릅니다.
- 스크립트는 점검·후보·초안만 만듭니다. 폐기·이관·공개 결정, 제출, 비밀기록물 처리는 사람이 합니다. 원본은 읽기만 합니다.
- 비밀기록물 목록은 Claude가 열지 않습니다. 담당자가 비밀 전용 장비에서 직접 실행합니다.

## 시험

```
python plugins/records-skills/tests/run_tests.py
```

48개(veraPDF가 PATH에 있으면 49개). 필요 패키지: openpyxl, python-docx, pypdf, pikepdf, python-hwpx, holidays, LibreOffice.

## 이용 조건

코드 MIT(LICENSE), 문서·설명 CC BY 4.0. 플러그인에 든 국가기록원 서식·준칙 표는 공공저작물(저작권법 제24조의2)로, 출처를 밝혀 씁니다.

만든 사람: Song Chang-gi · chgsong@naver.com · https://github.com/changgi
