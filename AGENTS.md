# 매주 쓰는 정산 엑셀을 자동으로 정리하는 스크립트

Instructions for coding agents (Codex and others) working in this repository.

> HYPOTHESIS DRAFT: the evidence is below the qualification bar. Confirm with real users before building beyond a prototype.

## Project key
`prd:655e5030-872f-417e-87c7-6914c5d87a79`

## Rules
- Ship one feature at a time and keep each change small.
- For a Seedup project, at session start or after compaction or resuming, call resume_project {"project":"prd:655e5030-872f-417e-87c7-6914c5d87a79"} before working, then get_next_task {"slug":"prd:655e5030-872f-417e-87c7-6914c5d87a79"} to pick a feature. Before stopping, leave a handoff with save_handoff for the same project.
- For a Seedup project, call update_feature_status when each feature starts (in_progress) and ends (done with checks or blocked).
- Name the PRD feature in the commit message, for example Implements F-03 (R-2 ← E-5, E-9).
- If the user has enabled commit uploads for this project, call link_commit after each commit with the project key, commit SHA, message, and changed files (`git diff --name-only HEAD~1`).
- Ask the user in the terminal before changing the locked goal or scope. Record their answer with record_decision.
- Every feature must trace to a requirement and every requirement to user evidence below. Do not add work that has no source. Ask first.
- When a readiness slot says "open question", ask instead of guessing.
- Treat these unverified ideas as out of scope until validated: H-01 사용자는 카드 매출 파일과 계좌 거래 파일을 받아 하나의 정리된 엑셀을 만들 수 있어야 한다, H-02 사용자는 카드 매출과 계좌 입금에서 승인, 취소, 환불, 수수료, 입금을 구분해 볼 수 있어야 한다, H-03 사용자는 주 단위로 카드 매출과 계좌 입금의 차이를 볼 수 있어야 한다, H-04 사용자는 원본 파일을 건드리지 않고 결과 파일을 새로 받을 수 있어야 한다, H-05 사용자는 처음 쓰는 파일 형식의 열 이름을 직접 지정할 수 있어야 한다, H-06 사용자는 정산 기간 전체의 합계를 한 화면에서 볼 수 있어야 한다, H-07 사용자는 카드 단말기가 여러 대여도 파일을 한 번에 넣을 수 있어야 한다.

## Locked goal (v1)
카드 내역과 계좌 내역을 손으로 옮기고 눈으로 대조하던 일을, 파일 두 개를 넣고 한 번 실행하는 일로 줄인다.

## Problem
카드·계좌 내역을 받아 매출을 맞추는 1인 사장님은 매주 같은 작업을 반복한다. 카드사에서 받은 매출 내역과 은행에서 받은 거래 내역을 각각 열어, 금액을 한 줄씩 눈으로 대조하면서 하나의 정산 엑셀에 옮긴다. 옮기는 중에는 승인과 취소, 부분 취소, 수수료, 며칠 뒤에 들어오는 입금이 한 파일에 섞여 있어서, 주 합계가 맞는지 보려면 행을 다시 세어야 한다. 금액이 맞지 않는 주가 보이면 어느 거래가 빠졌는지 찾는 데 시간이 더 든다. 이 제품은 파일 두 개를 받아 분류와 대조까지 계산하고, 사람이 볼 것은 맞지 않는 항목 목록만 남긴다.

## Technical design

### Architecture
로컬에서 도는 파이썬 스크립트 하나가 다섯 단계를 순서대로 처리한다. 설정을 읽고, 카드 파일과 계좌 파일을 각각 정규화하고, 행마다 거래 종류를 붙이고, 주 단위로 묶어 차이를 계산하고, 엑셀을 쓴다. 각 단계는 파일이나 목록을 주고받는 함수로 나누어, 실패하면 어느 파일의 몇 번째 행에서 멈췄는지 출력한다. 서버와 데이터베이스는 두지 않는다.

### Data model
- MappingConfig: cardDateColumn, cardAmountColumn, cardMemoColumn, bankDateColumn, bankAmountColumn, bankMemoColumn, weekStartsOn.
- Transaction: sourceFile, rowNumber, date, amount, kind(매출/취소/환불/수수료/입금/기타), memo.
- WeeklyMatch: weekStart, cardSales, cancels, fees, expectedDeposit, actualDeposit, diff, status(일치/불일치).
- IssueRow: sourceFile, rowNumber, reason, memo.
- Summary: totalSales, totalCancels, totalFees, totalDeposit, totalDiff.

### Interfaces
- settle init: 현재 폴더에 빈 매핑 설정 파일을 만든다. 결과는 만든 파일 경로 한 줄.
- settle run --card <카드매출파일> --bank <계좌파일>: 콘솔에 요약 5개 값, 확인 필요 건수, 결과 파일 경로를 출력하고 결과 엑셀을 저장한다. 종료 코드는 정상 0, 파일이나 설정을 못 읽음 2.
- settle run --card <파일1> --card <파일2> --bank <계좌파일>: 카드 파일을 여러 개 받아 합산한다. R7과 F5에서 지원한다.

### Stack and rationale
- Python 3.11: 카드사와 은행 엑셀은 머리글 행과 병합 셀이 섞여 있어 셀 단위로 열어 봐야 하고, 이를 다루는 라이브러리가 가장 많다.
- openpyxl: 입력 엑셀을 셀 단위로 읽고, 결과 엑셀에 시트 4장을 서식까지 넣어 쓴다.
- pandas: 주 단위 묶기와 합계를 짧은 코드로 처리하고, 금액 합산에서 부동소수 오차가 생기지 않게 정수(원)으로 바꿔 계산한다.
- PyYAML: 매핑 설정을 사람이 열어 고칠 수 있는 텍스트로 둔다. 사장님이 열 이름만 바꿀 수 있어야 한다.
- SQLite나 서버는 쓰지 않는다. 처리할 파일이 한 번에 1~2개이고, 사장님 컴퓨터에서만 돌아간다.

### Implementation order
1. 카드 엑셀과 계좌 CSV를 읽어 날짜·금액·메모를 뽑고, 못 읽은 행은 번호와 사유로 모은다 (F-01)
2. 메모와 금액 부호 규칙으로 거래 종류를 붙이고 기타 행을 따로 모은다 (F-02)
3. 주 단위로 합산해 예상 입금액과 실제 입금액의 차이를 계산하고 불일치 주 목록을 만든다 (F-03)
4. 시트 4장을 가진 엑셀을 원본과 다른 이름으로 저장하고 콘솔에 요약을 출력한다 (F-03, F-04)
5. 매핑 설정 파일 읽기와 빈 설정 파일 생성 명령을 붙인다 (F-01)


## Features
- [ ] F-01 파일 읽기와 열 매핑
  - (P0) 카드 매출 파일과 계좌 거래 파일, 설정 파일을 받아 날짜·금액·메모·출처·원본행번호를 가진 거래 목록 2개를 돌려준다. 읽지 못한 행은 버리지 않고 사유와 행 번호로 따로 돌려준다.
  - inputs: 카드 매출 파일 1개, 계좌 거래 파일 1개, 열 매핑 설정 파일 1개를 실행 인자(또는 CLI 옵션)로 받는다. 각 데이터 파일은 .xlsx 또는 .csv로 읽는다. 확인 필요: CSV 인코딩(UTF-8/CP949)과 구분자 자동 추론 지원 범위. / 설정 파일에는 출처별로 시트명, 헤더 행 번호, 날짜·금액·메모·거래종류에 해당하는 열 이름 또는 열 번호를 지정한다. 열 이름 지정이 우선하고, 없으면 열 번호로 대체한다. / 상태 조건: 세 파일이 모두 존재하고 읽기 가능해야 하며, 처리 대상 출처(카드/계좌)에 대한 매핑이 설정에 정의되어 있어야 한다. 매핑이 없으면 추측하지 않고 오류로 중단한다.
  - rules: 각 데이터 행에서 날짜, 금액, 거래 종류, 메모를 추출하고 출처(카드/계좌)와 원본 행 번호를 붙여, 카드 거래 목록과 계좌 거래 목록 2개를 반환한다. 두 목록은 동일한 6개 필드 구조를 쓴다. / 금액은 통화 기호·쉼표·공백을 제거하고 괄호 음수 표기를 반영해 정수(원)로 변환한다. 카드 매출은 양수, 계좌 출금은 음수로 저장한다. 확인 필요: 카드 매출 취소·환불의 부호 규칙. / 날짜는 설정에 지정된 형식(예: yyyyMMdd, yyyy.MM.dd, yyyy/MM/dd, Excel 일련번호)으로 파싱해 YYYY-MM-DD 문자열로 정규화한다. 시간이 포함된 값은 날짜 부분만 사용한다. / 빈 행은 결과에서도 실패 목록에서도 제외한다. 값이 있는 행 중 하나라도 파싱에 실패하면 해당 행은 정상 목록에서 제외하고 실패 목록에 사유와 함께 넣으며, 원본 행 번호는 헤더를 포함한 엑셀 기준 1-based 번호로 기록한다.
  - errors: 파일 없음, 읽기 권한 없음, 손상된 xlsx, 지정한 시트 없음, 헤더 행을 찾을 수 없음, 설정에 지정한 열을 찾을 수 없음은 각각 구분되는 오류 코드로 즉시 중단하고, 메시지에 파일명·시트명·지정한 열 이름을 포함한다. / 행 단위 실패(날짜 형식 불일치, 금액이 숫자가 아님, 필수값 누락)는 중단하지 않고 출처·파일명·시트명·원본행번호·원본값·실패사유를 수집해 실패 목록으로 반환한다. / 설정 파일이 없거나 JSON/YAML 파싱에 실패하거나 필수 키가 누락된 경우, 설정 파일 경로와 누락된 키 경로를 알려주고 중단한다. / 실패 행이 있어도 정상 행 결과는 그대로 반환하고 종료 코드는 0, 경고 요약(실패 건수와 행 번호 목록)을 표준 출력에 남긴다. 확인 필요: 실패 행 비율이 일정 수를 넘으면 전체 중단할지 여부.
  - access: 인증·인가 없이 로컬 단독 실행을 전제로 하며, 사용자가 지정한 3개 파일 경로만 읽고 그 외 경로에는 접근하지 않는다. 원본 파일은 수정하지 않는다. / 카드·계좌 내역은 민감 금융정보이므로 네트워크 전송, 외부 API 호출, 원격 로깅을 하지 않는다. / 로그와 오류 메시지에서 계좌번호·카드번호는 마스킹하고 메모 원문은 출력하지 않는다. 확인 필요: 로그 파일 저장 위치와 보관 기간.
  - data: DB를 사용하지 않고 파일 기반으로 동작한다. 거래 레코드는 거래일(YYYY-MM-DD 문자열), 금액(정수, 원), 거래종류(문자열), 메모(문자열), 출처(카드|계좌), 원본행번호(정수)의 6개 필드를 갖는다. / 설정 파일은 sources 아래 출처별로 file, sheet, header_row, columns(date, amount, memo, type)를 갖는 구조로 저장한다. 확인 필요: 설정 파일 형식 확정(YAML 또는 JSON 중 택1). / 실패 레코드는 거래 레코드 필드에 원본행값(raw)과 실패사유(reason)를 추가한 구조로 반환하며, 원본행번호는 두 목록에서 유일하게 행을 식별하는 키로 사용한다.
  - done: 카드사 1종, 은행 1종의 실제 파일과 CSV 1종으로 실행했을 때 두 거래 목록의 건수, 금액 합계, 최초·최종 거래일이 원본 파일과 일치한다. / 열 이름이 서로 다른 파일 3종 이상에 대해 코드 수정 없이 설정 파일만 교체해 재실행했을 때 동일하게 동작한다. / 실패 행이 섞인 파일에서 실패 목록의 행 번호가 실제 엑셀 행 번호와 정확히 일치하고, 정상 행은 누락·중복 없이 반환된다. / 단위 테스트로 날짜 형식 3종, 금액 표기 5종(쉼표, 통화 기호, 괄호 음수, 공백, 빈값), 빈 행, 병합 셀, 헤더 없는 파일, 없는 시트, 없는 열 케이스를 검증한다.
- [ ] F-02 거래 분류
  - (P0) 거래 목록을 받아 각 행에 매출·취소·환불·수수료·입금·기타 중 하나를 붙이고, 규칙에 걸리지 않은 행 목록을 함께 돌려준다.
  - rules: 메모에 '취소'가 있으면 취소, '환불'이면 환불, '수수료'면 수수료로 분류한다. / 출처가 카드이고 메모에 '승인'이 있으면 매출, 출처가 계좌이고 금액이 양수면 입금으로 분류한다. / 금액 부호는 매출·입금은 양수(+), 환불·수수료는 음수(-)로 본다. / 분류 우선순위는 취소 > 환불 > 수수료 > 매출 > 입금이며, 어디에도 해당하지 않으면 기타로 분류한다.
  - done: 매출·취소·환불·수수료·입금·기타 각 분류에 대한 테스트 케이스가 통과한다. / 규칙에 걸리지 않은 행이 미분류 목록에, 오류 행이 사유와 함께 정확히 분리되는 테스트가 통과한다. / 입력 엑셀을 넣으면 분류 컬럼이 추가된 출력 엑셀과 미분류 목록이 생성된다.
- [ ] F-03 주 단위 대조
  - (P0) 분류된 거래 목록 2개를 받아 주 시작일별로 매출·취소·수수료·예상 입금액·실제 입금액·차이를 계산하고, 차이가 1원 이상인 주와 그 주에 속한 행 번호를 돌려준다.
  - rules: 예상 입금액은 카드 매출 합계에서 취소액과 수수료를 차감한 값이고, 실제 입금액은 해당 주 계좌 입금 합계이며, 차이는 예상 입금액에서 실제 입금액을 뺀 값이다. / 확인 필요: 주의 시작 요일, 매출과 입금 사이 정산 시차 보정 규칙.
- [ ] F-04 엑셀 출력
  - (P0) 정리 시트, 대조 시트, 요약 시트, 확인 필요 시트를 받아 원본과 다른 이름의 엑셀 파일을 저장하고, 콘솔에 요약 5개 값과 확인 필요 건수, 결과 파일 경로를 출력한다.
- [ ] F-05 여러 파일 입력
  - 카드 매출 파일 여러 개와 계좌 파일 1개를 받아 출처 파일 열을 붙여 하나의 거래 목록으로 합친 뒤 F2~F4에 넘긴다.

(F-02~F-05의 전체 inputs/errors/access/data/done 항목은 Seedup의 전체 PRD를 보세요.)

## Screens, in order (a draft, confirm with users)
S-01. 입력 및 읽기: 카드 매출 파일 여러 개와 계좌 파일, 설정 파일을 읽어 거래 목록과 읽기 실패 행을 확인한다. (F-01, F-05)
S-02. 거래 분류: 거래 분류 규칙을 적용해 매출·취소·환불·수수료·입금·기타로 분류된 거래와 규칙 미적용 행을 확인한다. (F-02)
S-03. 주 단위 대조: 주 시작일별 매출·취소·수수료·예상 입금액·실제 입금액·차이를 확인하고 1원 이상 차이 주와 행 번호를 검토한다. (F-03)
S-04. 엑셀 저장: 정리·대조·요약·확인 필요 시트를 담은 엑셀을 원본과 다른 이름으로 저장하고 콘솔 요약과 결과 경로를 확인한다. (F-04)

## Do not build
- 세금 신고서와 부가세계산서를 만들지 않는다.
- 카드사와 은행 사이트에서 파일을 직접 받아오지 않는다. 자동 로그인이나 스크래핑을 하지 않는다.
- 매장과 직원 단위 권한 관리, 클라우드 저장, 여러 사람이 함께 쓰는 기능을 넣지 않는다.
- 현금 매출을 자동으로 추정하지 않는다.

## Success criteria
- G-1 대조 정확도 (measure: 최근 4주 정산 파일에서 스크립트가 계산한 예상 입금액이 사장님이 손으로 맞춘 값과 4주 모두 일치한다. 측정: 사장님 3명의 기존 정산 엑셀과 결과 대조 시트를 주 단위로 비교해 일치한 주 수를 센다.)
- G-2 사람이 확인할 행 비율 5% 이하. 측정 (measure: 사장님 3명의 최근 4주 정산 파일에서 결과 엑셀의 확인 필요 시트 행 수를 전체 거래 행 수로 나눈 비율을 계산한다.)
- G-3 처음 실행부터 요약 확인까지 5분 이내. 측정 (measure: 사장님 3명이 설치부터 결과 파일의 요약 시트를 확인할 때까지 걸린 시간을 각각 기록한다.)
- G-4 매핑 설정 없이 실행되는 비율. 측정 (measure: 지원 형식 3종 파일 각 1개에 settle init 없이 settle run을 실행해 성공 건수를 센다.)

## Ask the user before deciding
- 차별점으로 내세울 한 가지를 무엇으로 잡을지.
- 실행 방식을 터미널 명령으로 갈지, 더블클릭으로 도는 실행 파일로 포장할지, 간단한 창을 붙일지.
- 첫 버전에서 지원할 카드사·PG·은행 파일 형식 3개를 무엇으로 정할지.
- 현금 매출을 입력으로 받을지, 결과 엑셀에 사장님이 직접 적는 칸만 넣을지.
- 부가세와 세금 계산까지 결과에 넣을지.

This AGENTS.md is abbreviated to fit the rule file limit. Read the full PRD in Seedup.
