---
name: preservation-convert
description: 문서·표·발표 파일을 PDF/A-1b 보존포맷으로 변환하고 PDF/A 선언·글꼴 임베딩·암호화·veraPDF 적합성을 점검해 변환 이력을 남길 때 사용. "보존포맷 변환", "PDF/A로 바꿔줘", "장기보존 준비", "PDF/A 맞는지 확인" 요청에서 트리거.
---

# 보존포맷 변환

1. LibreOffice(soffice)가 있는지 확인한다. veraPDF가 있으면 `--verapdf`로 공식 검증까지 한다.
2. `python "${CLAUDE_PLUGIN_ROOT}"/skills/preservation-convert/scripts/to_pdfa.py 원본폴더/ 산출/보존포맷/ [--verapdf]`
3. 변환이력.csv의 판정을 요약한다. '확인 필요'(선언 없음·글꼴 미임베딩·veraPDF 부적합)는 원인을 함께 알린다.
4. veraPDF가 없으면 사전 점검: `python "${CLAUDE_PLUGIN_ROOT}"/skills/preservation-convert/scripts/pdfa_precheck.py 산출/보존포맷/ -o 산출/PDFA사전점검.csv`
   - NAK 30:2022 6장에서 자주 걸리는 것: 암호화, LZW, 내장 파일, OCProperties, OutputIntent(GTS_PDFA1), 폰트 미내장·부분집합 CharSet/CIDSet, 투명도, 금지 주석(FileAttachment·Sound·Movie), JavaScript·Launch 행위, XMP pdfaid(part=1, conformance=B).
   - '사전 점검 통과'는 공식 적합이 아니다.
5. veraPDF 없이 나온 '통과'는 "기본 점검 통과, 공식 검증 전"이라고 구분해 말한다. veraPDF 결과는 '적합'·'부적합(조항)'·'검증 오류'로 나뉘며, '검증 오류'는 판정이 아니다.
   - veraPDF 설치(공식 배포 software.verapdf.org의 verapdf-gf-installer.zip, Java 11+): 압축을 풀고 `java -jar verapdf-izpack-installer-*.jar -options 설정.properties`(내용 `INSTALL_PATH=/설치경로`)로 무인 설치한 뒤 PATH에 넣는다. 2026-11-20 veraPDF 1.30.4로 이 스크립트 결과를 검증했다.
   - 한글 파일명은 Java 환경에 따라 veraPDF가 못 읽어 '파일 없음'이 나므로, 스크립트가 영문 이름 임시 사본으로 검증한다.
4-1. 한글 글꼴: LibreOffice가 문서 글꼴을 못 찾아 Noto CJK 같은 OTF(CFF) 글꼴로 대체하면 Type1 부분집합에 CharSet이 빠져 PDF/A-1b 6.3.5를 어긴다. 글꼴 칸에 'CharSet 없음(6.3.5)'이 나오면 변환하는 PC에 원문서 글꼴(맑은 고딕·바탕·나눔고딕 등 TrueType)을 설치하거나 원문서 글꼴을 TrueType으로 바꿔 다시 변환한다. 확인한 결과 TrueType 한글 글꼴이면 PASS, OTF 대체면 FAIL.
4-2. 날짜: LibreOffice는 원본 작성일을 XMP에, 변환 시각을 문서정보에 넣어 6.7.3을 어긴다. 스크립트가 문서정보 날짜를 XMP에 맞춘다(원본에 날짜가 없어 XMP가 0000년이면 반대로). `--no-fix-dates`로 끌 수 있고, 맞춘 사실은 비고에 남는다. 변환이력의 결과 sha256은 맞춘 뒤의 값이다.
5. HWP/HWPX는 '한컴 변환 필요'로 남긴다. 억지로 변환하지 않는다.
6. 원본은 그대로 두고 보존본은 산출 폴더에만 쓴다. 끝나면 fixity-check로 보존본 기준선을 만든다.
7. 변환 시기: 기록관이 보존기간 30년 이상 기록물을 5년 이상 보존하게 되면 보존기간 기산일부터 6년이 지나기 전에 표준보존패키지(보존포맷 변환 포함)로 변환한다(NAK 10:2026(v1.5) 점검표 4-1, 개정 시행령 이관시기 5년과 맞춤). 기산일이 다음 해 1월 1일이므로 기한은 '종료 연도 + 6년'의 12월 31일까지. 「2026년 기록물관리 지침」의 '10년 이상 보존 → 11년 경과 전'은 개정 전 문구다. 다만 부칙 제17조①로 종전 이관 규정(기산일부터 10년)을 따르는 기록물(시행 당시 기산일+5년 경과, 종료연도 2020년 이전)은 기록관 보존 기간도 그만큼 길어지므로 변환 기한을 관할 영구기록물관리기관과 확인한다.

근거: NAK 30:2022(v1.1) 문서유형 보존포맷(PDF/A-1b), NAK 37:2025(v2.0) 보존포맷 선정기준(국가기록포털 표준 목록 2026-11-20 확인), NAK 10:2026(v1.5).
