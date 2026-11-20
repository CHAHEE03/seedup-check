# seedup-check

카드 매출 파일과 계좌 거래 파일을 읽어 주 단위 정산을 맞추는 로컬 스크립트(`settle`). 지금은 F-01(파일 읽기와 열 매핑)까지 있습니다.

```
pip install -e ".[test]"
settle init                                   # settle.yaml(빈 매핑 설정) 생성 → 열 이름을 채운다
settle run --card 카드매출.xlsx --bank 계좌.csv  # 건수·합계·기간과 읽기 실패 행 번호 출력
pytest -q
```

종료 코드: 정상 0(읽기 실패 행이 있어도 0, 경고로 출력), 파일·설정 오류 2.
