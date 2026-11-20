---
name: transfer-reconcile
description: 이관 목록(엑셀)과 실제 파일·상자 목록을 대조해 빠진 것·남는 것·중복을 보고하고, 일치할 때 BagIt 이관 패키지를 만들 때 사용. "이관 목록 대조", "실물 일치 검사", "이관 준비", "패키지 만들어줘" 요청에서 트리거.
---

# 이관 대조와 패키지

1. 목록의 머리글 행과 파일명 열을 확인한다(기본 열 이름: 파일명, 관리번호).
2. `python "${CLAUDE_PLUGIN_ROOT}"/skills/transfer-reconcile/scripts/reconcile_transfer.py 이관목록.xlsx 파일폴더/ --col 파일명 -o 산출/이관대조.xlsx`
3. 불일치가 1건이라도 있으면 여기서 멈추고 '불일치' 시트를 사용자에게 보여 준다. 처리 방법은 사람이 정한다.
4. 불일치 0건일 때만 `make_bag.py 파일폴더/ 산출/bag_이름/ --info Source-Organization=○○` 를 실행한다(설정상 실행 전 승인 요청).
5. `make_bag.py --verify 산출/bag_이름/`로 검증한 뒤 records-qa-gate를 `--manifest`와 함께 실행한다.

원본 폴더는 읽기만 한다. 패키지는 복사본으로 만든다.

## 이관목록 서식 만들기 (시행규칙 별지 제2호서식)
`python "${CLAUDE_PLUGIN_ROOT}"/skills/transfer-reconcile/scripts/make_transfer_list.py 등록정보.xlsx --org "기관명(처리과기관코드)" --year 2026 -o 산출/이관목록.xlsx`
- '1. 기록물철 단위'(분류번호·제목·건수·쪽수·형태·비고)와 '2. 건 단위'(등록번호·제목·쪽수·형태·철 분류번호·공개구분·공개제한 쪽표시·비고) 두 시트. 쪽수는 종이기록물만 적는다.
- 철은 단위과제별로 묶어 정렬한다(시행령 제32조⑤ 단위과제별 보존상자).
- 비밀 표시 행은 넣지 않는다(비밀기록물 이관은 시행령 제68조 별도 절차). 인계·인수 서명은 사람이 한다.

## 통합(전자+비전자) 이관 준비 (「2026년 기록물관리 지침」 제3장 참고 1)
`python "${CLAUDE_PLUGIN_ROOT}"/skills/transfer-reconcile/scripts/integrated_transfer.py 건목록.xlsx [--until 2015] -o 산출/통합이관_준비.xlsx`
- 기록관리시스템 건 목록에서 ~2015년 생산분의 비전자문서 건, 전자문서의 비전자 붙임(별송) 후보, 분리등록 첨부물을 찾는다. 붙임 후보는 지침의 '붙임(별첨) 표기 예시'(별송·따로보냄·우편·택배·인편송부 등 강한 단서 / 그림·필름·CD 등 약한 단서)로 고른다.
- '건 색인목록(서식)'·'붙임(별송) 색인목록(서식)'을 철별로 만든다. 비고(실물확인 여부)는 실물을 대조한 사람이 채운다.
- 국가기록원 '전자기록물 오류조치 공문'의 비전자 붙임(별송) 목록·육안검수결과와 대조한다. 비밀 표시 행은 넣지 않는다.

## 전자기록물 오프라인 이관(폐쇄망 기관)
- 지침: 이관매체(WORM 또는 변경 여부를 입증할 수 있는 재사용 매체)에 디스크 이미지로 담고 암호화, 인계·인수 양쪽에서 해시값을 비교, '전자기록물 인수·인계 확인서'(표준시간·도구명·해시 알고리즘·이관작업 시스템정보·이미지정보·인계자·인수자)를 작성.
- 해시 기준선과 대조는 fixity-check(baseline/verify)로 한다. 매체 수록·암호화·서명은 사람이 기관 승인 도구로 한다.

## 장기보존패키지(NEO2) 구조 점검 (NAK 31-1:2022)
`python "${CLAUDE_PLUGIN_ROOT}"/skills/transfer-reconcile/scripts/neo_check.py NEO폴더/ -o 산출/NEO점검.csv`
- 루트 neo:NAKEncapsulatedObject, 필수 요소(Version·SignedObject·ObjectContent·Record·RecordMetadata 등), 부모가 있을 때 필수인 자식(TitleCon→MainTitle, DateTimeCon→DateTimeCreated·DateTimeClosed, RetentionPeriodCon→RetentionPeriod·PreservationPlace 등), ComponentData Base64 해독과 SourceFileExtent 크기, SignatureBlock.
- 압축 패키지는 최상위에 Metadata.xml과 PackagingInformation.xml이 함께 있어야 한다.
- 서명·인증서 유효성(장기검증, NAK 32-1)은 보지 않는다 — 시스템·전용 검증기로 한다.

## 장기보존패키지(NEO3, 디렉토리 방식) 점검 (NAK 31-2:2022 부속서 A)
`python "${CLAUDE_PLUGIN_ROOT}"/skills/transfer-reconcile/scripts/neo_check.py 패키지.zip(또는 NEO3 폴더) -o 산출/NEO3점검.csv`
- 최상위에 Metadata.xml·PackagingInformation.xml·Metadata.xsd·PackagingInformation.xsd(이름 고정).
- ZIP이면 첫 항목 `mimetype`, 비압축, 내용 `application/neo3+zip`.
- PackagingInformation.xml Version 3.0, PackageManifest/ManifestItem의 Path 파일이 있고 HashValue가 HashAlgorithm(SHA2-256 등)으로 맞는지. MD5·SHA1은 '검토'로 표시, 목록에 없는 파일도 '검토'.
- UniqueID만 있는 외부(원격) 항목은 패키지 안에서 확인할 수 없어 건너뛴다.
