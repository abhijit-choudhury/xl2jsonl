import pytest

from xl2jsonl.chunker import (
    _detect_header_row,
    _is_empty_row,
    _normalize_cell,
    _normalize_headers,
    sheet_to_records,
)
from xl2jsonl.exceptions import NoHeaderError
from xl2jsonl.models import SheetData


def test_detect_header_row_simple():
    rows = [["Name", "Age", "City"], ["Alice", 30, "London"]]
    assert _detect_header_row(rows) == 0


def test_detect_header_row_with_empty_leading():
    rows = [[None, None], [None, None], ["Name", "Age"], ["Alice", 30]]
    assert _detect_header_row(rows) == 2


def test_detect_header_row_none():
    rows = [[1, 2, 3], [4, 5, 6]]
    assert _detect_header_row(rows) is None


def test_normalize_headers_basic():
    assert _normalize_headers(["Name", "Age"]) == ["Name", "Age"]


def test_normalize_headers_duplicates():
    result = _normalize_headers(["Rev", "Rev", "Rev"])
    assert result == ["Rev", "Rev_2", "Rev_3"]


def test_normalize_headers_blanks():
    result = _normalize_headers(["Name", None, ""])
    assert result == ["Name", "column_2", "column_3"]


def test_normalize_headers_whitespace():
    result = _normalize_headers(["  Name  ", "Multi\nLine\nHeader"])
    assert result == ["Name", "Multi Line Header"]


def test_is_empty_row():
    assert _is_empty_row([None, None, ""])
    assert _is_empty_row([])
    assert not _is_empty_row(["data"])
    assert not _is_empty_row([0])


def test_normalize_cell_types():
    import datetime

    assert _normalize_cell("  hello  ") == "hello"
    assert _normalize_cell(42) == 42
    assert _normalize_cell(3.14) == 3.14
    assert _normalize_cell(True) is True
    assert _normalize_cell(None) is None
    assert _normalize_cell(datetime.date(2024, 1, 15)) == "2024-01-15"
    assert _normalize_cell(datetime.datetime(2024, 6, 1, 14, 30)) == "2024-06-01T14:30:00"


def test_sheet_to_records_simple():
    sheet = SheetData(
        name="Test",
        index=0,
        rows=[
            ["Name", "Age"],
            ["Alice", 30],
            ["Bob", 25],
        ],
    )
    records = list(sheet_to_records(sheet, "test.xlsx"))
    assert len(records) == 2
    assert records[0].data == {"Name": "Alice", "Age": 30}
    assert records[0].metadata.filename == "test.xlsx"
    assert records[0].metadata.sheet_name == "Test"
    assert records[0].metadata.sheet_number == 1
    assert records[0].metadata.row_number == 2  # header=row1, first data=row2
    assert records[1].data == {"Name": "Bob", "Age": 25}
    assert records[1].metadata.row_number == 3


def test_sheet_to_records_skip_empty():
    sheet = SheetData(
        name="Test",
        index=0,
        rows=[
            ["Name", "Age"],
            ["Alice", 30],
            [None, None],
            ["Bob", 25],
        ],
    )
    records = list(sheet_to_records(sheet, "test.xlsx"))
    assert len(records) == 2


def test_sheet_to_records_keep_empty():
    sheet = SheetData(
        name="Test",
        index=0,
        rows=[
            ["Name", "Age"],
            ["Alice", 30],
            [None, None],
        ],
    )
    records = list(sheet_to_records(sheet, "test.xlsx", skip_empty_rows=False))
    assert len(records) == 2
    assert records[1].data == {"Name": None, "Age": None}


def test_sheet_to_records_explicit_header():
    sheet = SheetData(
        name="Test",
        index=0,
        rows=[
            ["garbage", "row"],
            ["Name", "Age"],
            ["Alice", 30],
        ],
    )
    records = list(sheet_to_records(sheet, "test.xlsx", header_row=1))
    assert len(records) == 1
    assert records[0].data == {"Name": "Alice", "Age": 30}


def test_sheet_to_records_no_header_raises():
    sheet = SheetData(name="Test", index=0, rows=[[1, 2], [3, 4]])
    with pytest.raises(NoHeaderError):
        list(sheet_to_records(sheet, "test.xlsx"))


def test_sheet_to_records_short_row_padded():
    sheet = SheetData(
        name="Test",
        index=0,
        rows=[
            ["A", "B", "C"],
            ["val"],
        ],
    )
    records = list(sheet_to_records(sheet, "test.xlsx"))
    assert records[0].data == {"A": "val", "B": None, "C": None}
