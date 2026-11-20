---
name: format-identify
description: 기록물 파일의 실제 포맷을 내용으로 식별하고 확장자 불일치·실행 파일·0바이트·보존포맷 검토 대상을 표시할 때 사용. "포맷 식별", "파일 형식 확인", "PRONOM", "이 파일 진짜 PDF야?" 요청에서 트리거.
---

# 포맷 식별

1. `python "${CLAUDE_PLUGIN_ROOT}"/skills/format-identify/scripts/identify.py 폴더/ -o 산출/포맷식별.csv`
2. 결과의 '표시' 열이 있는 파일만 사용자에게 요약한다. 어떤 엔진(Siegfried·Magika·시그니처)으로 판정했는지 함께 말한다.
3. PDF/A 여부의 최종 판정은 veraPDF 검증 결과로 한다(이 스크립트는 후보만 표시).
5. 열리지 않는 파일: `python "${CLAUDE_PLUGIN_ROOT}"/skills/format-identify/scripts/damage_check.py 폴더/ -o 산출/손상파일_분류.csv`
   - 국가기록원 「파일손상 전자기록물 복구 프로토타입 연구」(2018)의 손상 코드로 나누고, 표 18의 복구 가능 여부를 붙인다(O 원상 복구·△ 텍스트 추출·X 불가).
   - 보는 코드: S001(원본 장애·삭제·감염 안내문만 남음), B002(HML XML 오류), B003(10바이트 이하), B007(HWP 서명 없음), B015(OLE 뒤 잘림), B020(HWP 3.0 헤더), E001(전체 Base64), E002(.hox), E003(DestinyFS), E004(메일 인코딩), E005(DRM), E006(HC20), E007(FSPA).
   - 연구 밖이지만 같은 방식으로 PDF 끝 %%EOF 없음, ZIP 계열 구조·CRC 손상도 '추가-'로 표시한다.
   - 복구는 하지 않는다. 연구에서는 복구 가능 코드 2,324개 중 약 78%를 복구 엔진으로 되살렸다. O·△ 파일은 원본 보존 후 국가기록원 복구 엔진·전문 도구에 맡긴다.
4. 'NAK 37 권고(2025)' 열: NAK 37:2025(v2.0) 보존포맷 선정기준 평가결과로 보존포맷·수용가능포맷·권고 외를 붙인다.
   - 보존포맷: 텍스트 PDF·PDF/A-1·2·DOCX·ODT·EPUB / 발표 PDF·PPTX·ODP / 표 XLSX·ODS / 이미지 TIF·PNG / 오디오 WAV(BWF)·MP3·AAC / 비디오 MPG·MP4·MXF(코덱 AVC) / DB·구조화 CSV·JSON·XML
   - 수용가능포맷: HWPX, FLAC·WMA, AVI·MKV·MOV·WMV(코덱 FFV1), SIARD
   - 권고 외(평가했으나 권고 아님): HWP·DOC·TXT·PPT·SHOW·PPS·PPTM·XLS·CELL·XLK·ASC → preservation-convert로 변환 검토
   - 비디오는 컨테이너만 확장자로 본다. 코덱(AVC 권고)은 ffprobe(scan-qc의 av_qc.py)로 확인한다.
