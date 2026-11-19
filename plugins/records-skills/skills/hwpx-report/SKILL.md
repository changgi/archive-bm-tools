---
name: hwpx-report
description: 마크다운 보고서를 한글 HWPX 문서로 만들거나, 기관 공문·보고 서식(.hwpx)의 {{자리표시}}와 표 칸을 값으로 채우고, 스키마·남은 자리표시·본문 누락·개인정보 후보를 점검할 때 사용. "한글로 만들어줘", "hwpx로", "공문 서식 채워줘", "보고서 한글 파일" 요청에서 트리거.
---

# HWPX 보고서·공문

1. 새 문서: `python "${CLAUDE_PLUGIN_ROOT}"/skills/hwpx-report/scripts/make_hwpx.py --md 보고.md -o 산출/보고.hwpx --header "기관명" --page-number`
2. 기관 서식: `python "${CLAUDE_PLUGIN_ROOT}"/skills/hwpx-report/scripts/make_hwpx.py --template 서식.hwpx --values 값.xlsx [--md 본문.md] -o 산출/공문.hwpx`
   - 값 파일 2열(키, 값). `제목` → `{{제목}}` 치환, `칸:수신자` → '수신자' 칸 오른쪽 칸에 입력.
3. `*_점검.md`를 같이 보고한다. 스키마 실패·남은 자리표시·본문 누락이 있으면 끝난 것이 아니다.
4. 이 도구는 글자와 구조만 확인한다. 한글 프로그램에서 열어 모양을 확인하기 전에는 완성이라고 하지 않는다.
5. 결재·발송은 하지 않는다. 결과 파일이 이미 있으면 덮어쓰지 않는다.

공문 구성은 「행정업무의 운영 및 혁신에 관한 규정」의 기안문 구성(두문·본문·결문)에 맞춘다. 기관 서식이 있으면 그것을 우선한다.
필요: `pip install python-hwpx` (Apache-2.0)
