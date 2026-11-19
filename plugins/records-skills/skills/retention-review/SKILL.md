---
name: retention-review
description: 기록물 목록의 보존기간을 기관 기록관리기준표와 대조하고, 만료 지남·올해·내년 만료 목록과 보존기간 30년 이상 기록의 영구기록물관리기관 이관 대상을 추릴 때 사용. "보존기간 검토", "기준표 대조", "만료 예정 기록", "폐기 대상 목록 준비", "이관 대상 연도" 요청에서 트리거.
---

# 보존기간 검토

1. 기관 기록관리기준표(단위과제코드·단위과제명·보존기간)를 받는다. 없으면 assets/sample_criteria.csv 형식을 보여 주고 요청한다. 견본으로 실제 판단하지 않는다.
2. `python "${CLAUDE_PLUGIN_ROOT}"/skills/retention-review/scripts/check_retention.py 목록.xlsx --criteria 기준표.csv --year 2026 -o 산출/보존기간_검토.xlsx`
   이관연도는 현행 시행령 제40조①(기산일부터 5년 경과 다음 연도)으로 계산한다. 개정 전 규정으로 보려면 `--transfer-years 10`. 입법예고안 5종 값이 섞인 목록이면 `--scheme draft`(현행 시행령은 7종 유지).
   보존기간을 바꾸자고 제안할 때는 references/byeolpyo1-retention-criteria.md의 별표 1 해당 호를 근거로 든다.
3. '기준표보다 짧음'은 조기 폐기 위험이므로 맨 먼저 보고한다.
4. '만료 지남·올해 만료' 목록은 disposal-dossier(심의자료)의 입력이 된다. 폐기를 제안하거나 실행하지 않는다.
5. 끝나면 records-qa-gate.

## 공통업무 기준: 준칙 → NAK 39
- 기관 기준표 대신 `--common 기관유형`을 쓰면 현행 공통업무 보존기간 준칙(assets/common/common_task_criteria.csv)으로 대조한다.
- NAK 39:2026(v1.0) 「기록물의 보존기간 책정표준: 공통업무」가 준칙을 표준으로 바꾸는 중이다(국가기록원고시 제2026-8호 예고, 2026-11-02, 확정 전). `--common-basis nak39안`이면 예고안 부속서 A~F(1,333행)로 대조하고 '확정 고시 전' 알림을 낸다.
- 예고안과 현행 준칙 차이(공공기관 2건·대학 4건 보존기간 변경, 교육·학교 단위과제 추가·삭제, 명칭 '기록관리표')는 references/nak39-draft-diff.md. 사용자에게 바뀌는 과제를 먼저 알린다. 확정 고시가 나오면 확정본으로 자료를 바꾸고 기본값을 옮긴다.

계산 규칙과 근거는 references/retention-rules.md.
