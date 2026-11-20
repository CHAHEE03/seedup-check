"""settle 명령: init(빈 매핑 설정 만들기), run(카드·계좌 파일 읽기)."""
import argparse
import sys
from pathlib import Path

from settle.read import SOURCE_LABELS, SettleError, load_config, read_inputs

CONFIG_NAME = "settle.yaml"

_SOURCE_TEMPLATE = """  {source}:            # {label} 파일
    sheet:            # 엑셀 시트 이름. 비우면 첫 시트 (CSV는 무시)
    header_row: 1     # 열 이름이 있는 행 번호. 머리글이 없으면 0
    columns:          # 머리글 이름("거래일자") 또는 왼쪽부터 센 열 번호(1)
      date:           # 거래일 (필수)
      amount:         # 금액 (필수)
      memo:           # 메모 (선택)
      type:           # 거래종류 (선택)
"""
TEMPLATE = "# settle 열 매핑 설정\nsources:\n" + "".join(
    _SOURCE_TEMPLATE.format(source=s, label=label) for s, label in SOURCE_LABELS.items()
)


def _init():
    path = Path.cwd() / CONFIG_NAME
    if path.exists():
        print(f"이미 있습니다: {path}", file=sys.stderr)
        return 1
    path.write_text(TEMPLATE, encoding="utf-8")
    print(path)
    return 0


def _run(args):
    result = read_inputs(args.card, args.bank, load_config(args.config))
    for source, (txs, failures) in result.items():
        label = SOURCE_LABELS[source]
        if txs:
            dates = sorted(t.date for t in txs)
            total = sum(t.amount for t in txs)
            print(f"{label}: {len(txs)}건, 합계 {total:,}원, {dates[0]} ~ {dates[-1]}")
        else:
            print(f"{label}: 0건")
        if failures:
            rows = ", ".join(str(f.row) for f in failures)
            print(f"경고: {label} 읽기 실패 {len(failures)}건 (행 {rows})")
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(prog="settle")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("init", help=f"현재 폴더에 빈 매핑 설정 파일({CONFIG_NAME})을 만든다")
    run = sub.add_parser("run", help="카드·계좌 파일을 읽는다")
    run.add_argument("--card", required=True, help="카드 매출 파일 (.xlsx/.csv)")
    run.add_argument("--bank", required=True, help="계좌 거래 파일 (.xlsx/.csv)")
    run.add_argument("--config", default=CONFIG_NAME, help=f"매핑 설정 파일 (기본 {CONFIG_NAME})")
    args = parser.parse_args(argv)
    try:
        return _init() if args.command == "init" else _run(args)
    except SettleError as e:
        print(f"오류[{e.code}] {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
