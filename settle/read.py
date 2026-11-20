"""F-01: 카드·계좌 파일(.xlsx/.csv)을 설정의 열 매핑대로 읽어 공통 거래 목록으로 바꾼다."""
import csv
import datetime as dt
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import yaml
from openpyxl.utils.exceptions import InvalidFileException

SOURCE_LABELS = {"card": "카드", "bank": "계좌"}


class SettleError(Exception):
    """즉시 중단해야 하는 오류. code로 종류를 구분한다."""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


@dataclass
class Transaction:
    date: str  # YYYY-MM-DD
    amount: int  # 원, 부호는 원본 그대로
    kind: str  # 원본의 거래종류 열 값(없으면 "")
    memo: str
    source: str  # 카드 | 계좌
    row: int  # 원본 파일 기준 1-based 행 번호


@dataclass
class Failure:
    source: str
    file: str
    sheet: str
    row: int
    raw: dict
    reason: str


# ---------- 설정 ----------

def load_config(path):
    path = Path(path)
    if not path.is_file():
        raise SettleError("config_not_found", f"설정 파일이 없습니다: {path}")
    try:
        cfg = yaml.safe_load(path.read_text(encoding="utf-8"))
    except (yaml.YAMLError, UnicodeDecodeError) as e:
        raise SettleError("config_invalid", f"설정 파일을 해석하지 못했습니다: {path} ({e.__class__.__name__})")
    for source in SOURCE_LABELS:
        for key in ("header_row", "columns.date", "columns.amount"):
            key_path = f"sources.{source}.{key}"
            node = cfg
            for part in key_path.split("."):
                node = node.get(part) if isinstance(node, dict) else None
            if node is None or node == "":
                raise SettleError("config_missing_key", f"설정 파일 {path}에 {key_path} 값이 없습니다.")
    return cfg


# ---------- 값 변환 ----------

_DATE_SEP = re.compile(r"(\d{4})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})")
_DATE_PACKED = re.compile(r"(\d{4})(\d{2})(\d{2})(?!\d)")
_EXCEL_EPOCH = dt.date(1899, 12, 30)


def parse_date(value):
    """YYYY-MM-DD 문자열, 못 읽으면 None. 시간 부분은 버린다."""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer() and 19000101 <= value <= 29991231:
            value = str(int(value))  # yyyyMMdd가 숫자 셀로 들어온 경우
        elif 0 < value < 100000:
            return (_EXCEL_EPOCH + dt.timedelta(days=int(value))).isoformat()  # 엑셀 일련번호
        else:
            return None
    if not isinstance(value, str):
        return None
    s = value.strip()
    m = _DATE_SEP.match(s) or _DATE_PACKED.match(s)
    if not m:
        return None
    try:
        return dt.date(*map(int, m.groups())).isoformat()
    except ValueError:
        return None


_AMOUNT_NOISE = re.compile(r"[\s,₩￦원$]|KRW", re.I)
_AMOUNT_NUM = re.compile(r"[+-]?\d+(?:\.0+)?")


def parse_amount(value):
    """정수(원), 못 읽으면 None."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    if not isinstance(value, str):
        return None
    s = value.strip()
    negative = s.startswith("(") and s.endswith(")")
    s = _AMOUNT_NOISE.sub("", s.strip("()"))
    if not _AMOUNT_NUM.fullmatch(s):
        return None
    n = int(float(s))
    return -abs(n) if negative else n


def _blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


# ---------- 파일 읽기 ----------

def _read_rows(path, sheet):
    """(시트 이름, 행 목록)을 돌려준다. 행 목록의 i번째가 원본 i+1행."""
    path = Path(path)
    if not path.exists():
        raise SettleError("file_not_found", f"파일이 없습니다: {path}")
    try:
        if path.suffix.lower() == ".csv":
            return "", _read_csv(path)
        wb = openpyxl.load_workbook(path, data_only=True)
    except PermissionError:
        raise SettleError("permission_denied", f"파일을 읽을 권한이 없습니다: {path}")
    except (zipfile.BadZipFile, InvalidFileException, KeyError, OSError) as e:
        raise SettleError("corrupt_file", f"엑셀 파일을 열지 못했습니다: {path} ({e.__class__.__name__})")
    if sheet and sheet not in wb.sheetnames:
        raise SettleError("sheet_not_found", f"{path.name}에 '{sheet}' 시트가 없습니다. 있는 시트: {', '.join(wb.sheetnames)}")
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    for rng in ws.merged_cells.ranges:  # 병합 셀은 왼쪽 위 값으로 채운다
        top_left = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                rows[r - 1][c - 1] = top_left
    return ws.title, rows


def _read_csv(path):
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp949"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SettleError("corrupt_file", f"CSV 인코딩을 알 수 없습니다(UTF-8, CP949만 지원): {path}")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [list(r) for r in csv.reader(text.splitlines(), dialect)]


def read_source(path, source_cfg, source):
    """한 출처 파일을 읽어 (거래 목록, 실패 목록)을 돌려준다."""
    label = SOURCE_LABELS[source]
    name = Path(path).name
    sheet, rows = _read_rows(path, source_cfg.get("sheet"))
    where = f"{name}" + (f" '{sheet}' 시트" if sheet else "")
    header_row = int(source_cfg["header_row"])
    header = {}
    if header_row > 0:
        if header_row > len(rows) or all(_blank(v) for v in rows[header_row - 1]):
            raise SettleError("header_not_found", f"{where}의 {header_row}행에서 열 이름을 찾지 못했습니다.")
        header = {str(v).strip(): i for i, v in enumerate(rows[header_row - 1]) if not _blank(v)}
    width = max((len(r) for r in rows), default=0)

    cols = {}
    for field, spec in source_cfg["columns"].items():
        if spec is None or spec == "":
            continue
        if isinstance(spec, int) and 1 <= spec <= width:
            cols[field] = spec - 1
        elif isinstance(spec, str) and spec.strip() in header:
            cols[field] = header[spec.strip()]
        else:
            raise SettleError("column_not_found", f"{where}에서 {field} 열 '{spec}'을(를) 찾지 못했습니다.")

    txs, failures = [], []
    for row_no in range(header_row + 1, len(rows) + 1):
        row = rows[row_no - 1]
        if all(_blank(v) for v in row):
            continue
        cell = {f: (row[i] if i < len(row) else None) for f, i in cols.items()}
        date, amount = parse_date(cell["date"]), parse_amount(cell["amount"])
        reasons = []
        if _blank(cell["date"]):
            reasons.append("날짜 없음")
        elif date is None:
            reasons.append("날짜 형식 불일치")
        if _blank(cell["amount"]):
            reasons.append("금액 없음")
        elif amount is None:
            reasons.append("금액이 숫자가 아님")
        if reasons:
            raw = {f: ("" if v is None else str(v)) for f, v in cell.items()}
            failures.append(Failure(label, name, sheet, row_no, raw, ", ".join(reasons)))
            continue
        kind, memo = ("" if _blank(cell.get(f)) else str(cell[f]).strip() for f in ("type", "memo"))
        txs.append(Transaction(date, amount, kind, memo, label, row_no))
    return txs, failures


def read_inputs(card_path, bank_path, config):
    """카드·계좌 파일을 읽어 {'card': (거래, 실패), 'bank': (거래, 실패)}를 돌려준다."""
    return {
        "card": read_source(card_path, config["sources"]["card"], "card"),
        "bank": read_source(bank_path, config["sources"]["bank"], "bank"),
    }
