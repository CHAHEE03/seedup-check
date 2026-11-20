import datetime as dt

import openpyxl
import pytest
import yaml

from settle import cli
from settle.read import SettleError, load_config, parse_amount, parse_date, read_inputs, read_source


def make_xlsx(path, rows, merges=(), sheet="내역"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    for m in merges:
        ws.merge_cells(m)
    wb.save(path)
    return path


def src(columns, header_row=1, sheet=None):
    return {"sheet": sheet, "header_row": header_row, "columns": columns}


# ---------- 값 변환 ----------

@pytest.mark.parametrize("value", ["20260105", "2026.01.05", "2026/01/05", "2026-01-05 13:45:00", "2026. 1. 5.",
                                   46027, 20260105, dt.datetime(2026, 1, 5, 9, 30), dt.date(2026, 1, 5)])
def test_parse_date_formats(value):
    assert parse_date(value) == "2026-01-05"


@pytest.mark.parametrize("value", ["2026.13.01", "어제", "", None, 0])
def test_parse_date_rejects(value):
    assert parse_date(value) is None


@pytest.mark.parametrize("value,expected", [
    ("1,234,000", 1234000),      # 쉼표
    ("₩12,000", 12000),          # 통화 기호
    ("12,000원", 12000),
    ("(3,500)", -3500),          # 괄호 음수
    (" 1 500 ", 1500),           # 공백
    ("-2,000", -2000),
    (5000, 5000),
    (5000.0, 5000),
])
def test_parse_amount_formats(value, expected):
    assert parse_amount(value) == expected


@pytest.mark.parametrize("value", ["", "  ", None, "abc", "12.5", 12.5])
def test_parse_amount_rejects(value):
    assert parse_amount(value) is None


# ---------- 파일 읽기 ----------

CARD_ROWS = [
    ["가맹점 매출 내역"],                                   # 1행: 제목
    ["승인일자", "승인금액", "구분", "가맹점"],             # 2행: 머리글
    ["2026.01.05", "12,000", "승인", "아메리카노"],         # 3
    [None, "(12,000)", "취소", "아메리카노"],               # 4: 날짜는 3~4행 병합
    [None, None, None, None],                               # 5: 빈 행
    ["2026/01/07", "₩8,500", "승인", "라떼"],              # 6
    ["날짜아님", "9,000", "승인", "케이크"],                # 7: 날짜 실패
    ["20260109", "구천원", "승인", "쿠키"],                 # 8: 금액 실패
    ["2026-01-11 18:00", "1 500", "승인", "물"],            # 9
]
CARD_COLS = {"date": "승인일자", "amount": "승인금액", "type": "구분", "memo": "가맹점"}


def test_xlsx_rows_failures_and_merged_cells(tmp_path):
    path = make_xlsx(tmp_path / "card.xlsx", CARD_ROWS, merges=["A3:A4"])
    txs, failures = read_source(path, src(CARD_COLS, header_row=2), "card")

    assert [(t.row, t.date, t.amount) for t in txs] == [
        (3, "2026-01-05", 12000), (4, "2026-01-05", -12000), (6, "2026-01-07", 8500), (9, "2026-01-11", 1500)]
    assert txs[0].kind == "승인" and txs[0].memo == "아메리카노" and txs[0].source == "카드"
    assert [(f.row, f.reason) for f in failures] == [(7, "날짜 형식 불일치"), (8, "금액이 숫자가 아님")]
    assert failures[0].raw["date"] == "날짜아님" and failures[0].sheet == "내역" and failures[0].file == "card.xlsx"


def test_missing_required_value_goes_to_failures(tmp_path):
    rows = [["일자", "금액"], ["2026-01-05", None], [None, "1000"]]
    _, failures = read_source(make_xlsx(tmp_path / "b.xlsx", rows), src({"date": "일자", "amount": "금액"}), "bank")
    assert [(f.row, f.reason, f.source) for f in failures] == [(2, "금액 없음", "계좌"), (3, "날짜 없음", "계좌")]


def test_csv_cp949_and_utf8_with_semicolon(tmp_path):
    text = "거래일시;입금액;적요\n2026-01-06;30,000;카드대금\n\n2026-01-08;(1,200);수수료\n"
    for enc in ("cp949", "utf-8-sig"):
        path = tmp_path / f"bank-{enc}.csv"
        path.write_bytes(text.encode(enc))
        txs, failures = read_source(path, src({"date": "거래일시", "amount": "입금액", "memo": "적요"}), "bank")
        assert [(t.row, t.date, t.amount, t.memo) for t in txs] == [
            (2, "2026-01-06", 30000, "카드대금"), (4, "2026-01-08", -1200, "수수료")]
        assert failures == []


def test_three_layouts_same_result_with_config_only(tmp_path):
    data = [("2026-01-05", 10000, "A"), ("2026-01-06", -3000, "B")]
    a = make_xlsx(tmp_path / "a.xlsx", [["일자", "금액", "내용"], *data])
    b = make_xlsx(tmp_path / "b.xlsx", [["무시"], ["금액(원)", "비고", "거래일"], *[(m, n, d) for d, m, n in data]],
                  sheet="Sheet2")
    c = tmp_path / "c.csv"
    c.write_text("\n".join(f"{d.replace('-', '')},{m},{n}" for d, m, n in data), encoding="utf-8")
    results = [
        read_source(a, src({"date": "일자", "amount": "금액", "memo": "내용"}), "card"),
        read_source(b, src({"date": "거래일", "amount": "금액(원)", "memo": "비고"}, header_row=2, sheet="Sheet2"),
                    "card"),
        read_source(c, src({"date": 1, "amount": 2, "memo": 3}, header_row=0), "card"),  # 머리글 없는 파일
    ]
    simplified = [[(t.date, t.amount, t.memo) for t in txs] for txs, _ in results]
    assert simplified[0] == simplified[1] == simplified[2] == [("2026-01-05", 10000, "A"), ("2026-01-06", -3000, "B")]


# ---------- 중단 오류 ----------

def test_errors_stop_with_codes(tmp_path):
    path = make_xlsx(tmp_path / "card.xlsx", CARD_ROWS)
    cases = [
        (tmp_path / "없음.xlsx", src(CARD_COLS, 2), "file_not_found"),
        (path, src(CARD_COLS, 2, sheet="없는시트"), "sheet_not_found"),
        (path, src({**CARD_COLS, "date": "거래일"}, 2), "column_not_found"),
        (path, src(CARD_COLS, 5), "header_not_found"),   # 빈 행을 머리글로 지정
        (path, src({"date": 9, "amount": 2}, 0), "column_not_found"),
    ]
    for p, cfg, code in cases:
        with pytest.raises(SettleError) as e:
            read_source(p, cfg, "card")
        assert e.value.code == code
    with pytest.raises(SettleError) as e:
        read_source(path, src({**CARD_COLS, "date": "거래일"}, 2), "card")
    assert "card.xlsx" in str(e.value) and "내역" in str(e.value) and "거래일" in str(e.value)


def test_corrupt_xlsx(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip")
    with pytest.raises(SettleError) as e:
        read_source(bad, src(CARD_COLS, 2), "card")
    assert e.value.code == "corrupt_file"


def test_config_errors(tmp_path):
    with pytest.raises(SettleError) as e:
        load_config(tmp_path / "settle.yaml")
    assert e.value.code == "config_not_found"

    broken = tmp_path / "broken.yaml"
    broken.write_text("sources: [", encoding="utf-8")
    with pytest.raises(SettleError) as e:
        load_config(broken)
    assert e.value.code == "config_invalid"

    partial = tmp_path / "partial.yaml"
    partial.write_text(yaml.safe_dump({"sources": {"card": src(CARD_COLS, 2)}}, allow_unicode=True), encoding="utf-8")
    with pytest.raises(SettleError) as e:
        load_config(partial)
    assert e.value.code == "config_missing_key" and "sources.bank.header_row" in str(e.value)


# ---------- CLI ----------

def test_cli_init_then_run(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    assert cli.main(["init"]) == 0
    assert capsys.readouterr().out.strip().endswith("settle.yaml")
    assert cli.main(["run", "--card", "c.xlsx", "--bank", "b.csv"]) == 2  # 빈 설정은 추측하지 않고 중단
    assert "config_missing_key" in capsys.readouterr().err

    make_xlsx(tmp_path / "c.xlsx", CARD_ROWS, merges=["A3:A4"])
    (tmp_path / "b.csv").write_text("일자,금액\n2026-01-06,30000\n", encoding="utf-8")
    (tmp_path / "settle.yaml").write_text(yaml.safe_dump({"sources": {
        "card": src(CARD_COLS, 2), "bank": src({"date": "일자", "amount": "금액"})}}, allow_unicode=True),
        encoding="utf-8")
    assert cli.main(["run", "--card", "c.xlsx", "--bank", "b.csv"]) == 0
    out = capsys.readouterr().out
    assert "카드: 4건, 합계 10,000원, 2026-01-05 ~ 2026-01-11" in out
    assert "경고: 카드 읽기 실패 2건 (행 7, 8)" in out
    assert "계좌: 1건, 합계 30,000원, 2026-01-06 ~ 2026-01-06" in out
    assert "아메리카노" not in out  # 메모 원문은 콘솔에 내지 않는다


def test_read_inputs_does_not_modify_sources(tmp_path):
    card = make_xlsx(tmp_path / "c.xlsx", CARD_ROWS, merges=["A3:A4"])
    bank = tmp_path / "b.csv"
    bank.write_text("일자,금액\n2026-01-06,30000\n", encoding="utf-8")
    before = card.read_bytes(), bank.read_bytes()
    cfg = {"sources": {"card": src(CARD_COLS, 2), "bank": src({"date": "일자", "amount": "금액"})}}
    result = read_inputs(card, bank, cfg)
    assert len(result["card"][0]) == 4 and len(result["bank"][0]) == 1
    assert (card.read_bytes(), bank.read_bytes()) == before
