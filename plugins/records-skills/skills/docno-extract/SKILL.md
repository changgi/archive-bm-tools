---
name: docno-extract
description: 기록물 목록 엑셀의 문서번호 열에서 문서번호를 추출하고 '기획행정담당관-351-92~94' 같은 물결표 범위를 낱개로 펼칠 때 사용. "문서번호 추출", "범위 펼쳐줘", "문서번호 전체" 요청에서 트리거.
---

# 문서번호 추출

1. 문서번호가 든 열 이름과 식별 열(철관리번호 등)을 확인한다.
2. `python "${CLAUDE_PLUGIN_ROOT}"/skills/docno-extract/scripts/extract_docno.py 목록.xlsx --col 구기록물문서번호 --id-col 철관리번호 -o 산출/문서번호.xlsx`
3. '비고' 열에 '범위 해석 불가'·'펼치지 않음'이 있으면 그 행을 사용자에게 따로 보여 준다.
