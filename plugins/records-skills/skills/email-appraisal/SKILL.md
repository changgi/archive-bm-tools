---
name: email-appraisal
description: MBOX 파일이나 EML 폴더로 받은 이메일 기록을 평가할 때 메시지 목록·주고받은 상대·첨부 목록·개인정보 후보·비밀 표시를 정리하는 데 사용. "이메일 기록 정리", "메일함 평가", "MBOX 분석", "퇴직자 메일 인수", "이메일 기증 자료" 요청에서 트리거.
---

# 이메일 기록 평가 보조

1. `python "${CLAUDE_PLUGIN_ROOT}"/skills/email-appraisal/scripts/email_report.py 보관함.mbox -o 산출/이메일평가.xlsx`
2. 결과에는 본문이 없다. 본문을 요약하거나 옮겨 적지 않는다.
3. '비밀' 표시가 있는 메시지는 열지 않고 담당자에게 넘긴다.
4. 기록물로 남길 메일, 보존기간, 공개 여부는 사람이 정한다. 상대 집계는 업무 관계를 파악하는 참고자료로만 쓴다.
5. 첨부 파일은 목록과 해시만 만든다. 첨부를 보존하려면 따로 꺼내 format-identify·fixity-check를 거친다.

참고: ePADD(평가·정리·검색·제공 4개 모듈의 이메일 아카이브 도구), InterPARES 3 이메일 관리·보존 모델(EMPM).
