---
name: records-qa-gate
description: 기록물 산출 목록·메타데이터·이관 패키지를 내보내기 직전에 기록의 4대 속성(진본성·무결성·신뢰성·이용가능성)으로 검수할 때 사용. "검수해줘", "내보내도 돼?", "공개구분 확인", "비공개 사유 점검", "최종 점검" 요청과, 다른 기록관리 스킬이 끝난 뒤 항상 실행.
---

# 검수 게이트

1. 산출 목록(xlsx·csv)을 정한다. 원본 목록이 있으면 `--source`, 이관 패키지가 있으면 `--manifest`를 함께 준다.
2. 실행:
   `python "${CLAUDE_PLUGIN_ROOT}"/skills/records-qa-gate/scripts/qa_gate.py 산출/목록.xlsx --source 원본/이관목록.xlsx --key 관리번호 --manifest 산출/bag/manifest-sha256.txt -o 산출/검수결과.md`
3. 종료 코드 0(통과)일 때만 사람에게 넘긴다. 1(실패)이면 검수결과.md의 실패 항목을 고친 뒤 다시 실행한다.
4. '건너뛴 점검'이 있으면 사용자에게 그대로 알린다. 건너뛴 점검을 통과로 말하지 않는다.

판정 기준은 references/four-properties.md.

## 열 맞춤 점검 (기관 시스템 내보내기 목록을 처음 쓸 때)
`python "${CLAUDE_PLUGIN_ROOT}"/skills/records-qa-gate/scripts/header_check.py 내보낸목록.xlsx [--script production_status]`
- 머리글 행만 읽고 자료 행은 읽지도 출력하지도 않는다. 킷의 스크립트(ALIAS를 가진 모든 스크립트)마다 어느 열을 알아보는지, '비슷함'(띄어쓰기·괄호·별표·년도/연도 차이), '후보'(이름 일부가 겹침), '없음'을 보여 준다.
- `--fix -o 사본.xlsx`: '비슷함' 머리글만 바꾼 사본(자료 그대로, 원본은 안 고침). '후보'는 담당자가 확인해 `--map "기관열=킷열"`로 짝지어 준다.
- 사용자에게 자료를 보내 달라고 하지 말고, 이 명령의 출력(머리글 이름만)을 받아 짝을 정한다.
- 비밀 폴더의 파일은 다루지 않는다.
