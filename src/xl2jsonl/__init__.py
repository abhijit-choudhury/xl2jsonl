"""xl2jsonl - Convert Excel files to JSONL with robust merged-cell handling."""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

from xl2jsonl.chunker import sheet_to_records
from xl2jsonl.loader import load_workbook
from xl2jsonl.models import RowRecord
from xl2jsonl.writer import record_to_dict, write_jsonl


def convert(
    input_path: str | Path,
    output_path: str | Path | None = None,
    *,
    sheets: list[str | int] | None = None,
    engine: str = "auto",
    header_row: int | None = None,
    skip_empty_rows: bool = True,
) -> list[dict] | int:
    """Convert an Excel file to JSONL.

    If output_path is given, writes JSONL file and returns row count.
    If output_path is None, returns list of dicts.
    """
    input_path = Path(input_path)
    records = iter_records(
        input_path,
        sheets=sheets,
        engine=engine,
        header_row=header_row,
        skip_empty_rows=skip_empty_rows,
    )

    if output_path is not None:
        return write_jsonl(records, Path(output_path))

    return [record_to_dict(r) for r in records]


def iter_records(
    input_path: str | Path,
    *,
    sheets: list[str | int] | None = None,
    engine: str = "auto",
    header_row: int | None = None,
    skip_empty_rows: bool = True,
) -> Iterator[RowRecord]:
    """Lazy iterator over all records from an Excel file."""
    input_path = Path(input_path)
    filename = input_path.name
    sheet_data_list = load_workbook(input_path, sheets=sheets, engine=engine)

    for sheet in sheet_data_list:
        yield from sheet_to_records(
            sheet,
            filename=filename,
            header_row=header_row,
            skip_empty_rows=skip_empty_rows,
        )
