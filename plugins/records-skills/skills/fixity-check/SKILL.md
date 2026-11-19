---
name: fixity-check
description: 보존 파일의 SHA-256 기준선을 만들고 정기적으로 대조해 변경·사라짐·새 파일을 찾을 때 사용. "무결성 검사", "해시 대조", "기준선 만들어줘", "변조 확인" 요청에서 트리거.
---

# 무결성 점검

- AIC가 설치돼 있으면 `aic-baseline` / `aic-check`를 우선 쓴다.
- 없으면: `python "${CLAUDE_PLUGIN_ROOT}"/skills/fixity-check/scripts/fixity.py baseline 보존/ -o 산출/기준선.csv` → 이후 `fixity.py verify 보존/ 산출/기준선.csv`
- '변경'·'사라짐'이 나오면 복구 방법은 사람이 정한다. 파일을 되돌리거나 지우지 않는다.
