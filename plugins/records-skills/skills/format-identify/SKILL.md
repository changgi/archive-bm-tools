---
name: format-identify
description: 기록물 파일의 실제 포맷을 내용으로 식별하고 확장자 불일치·실행 파일·0바이트·보존포맷 검토 대상을 표시할 때 사용. "포맷 식별", "파일 형식 확인", "PRONOM", "이 파일 진짜 PDF야?" 요청에서 트리거.
---

# 포맷 식별

1. `python "${CLAUDE_PLUGIN_ROOT}"/skills/format-identify/scripts/identify.py 폴더/ -o 산출/포맷식별.csv`
2. 결과의 '표시' 열이 있는 파일만 사용자에게 요약한다. 어떤 엔진(Siegfried·Magika·시그니처)으로 판정했는지 함께 말한다.
3. PDF/A 여부의 최종 판정은 veraPDF 검증 결과로 한다(이 스크립트는 후보만 표시).
