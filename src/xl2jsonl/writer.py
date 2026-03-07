from __future__ import annotations

import io
from pathlib import Path
from typing import Iterator

import orjson

from xl2jsonl.models import RowRecord


def record_to_dict(record: RowRecord) -> dict:
    return {
        "content": record.data,
        "metadata": {
            "filename": record.metadata.filename,
            "sheet_name": record.metadata.sheet_name,
            "sheet_number": record.metadata.sheet_number,
            "row_number": record.metadata.row_number,
        },
    }


def write_jsonl(
    records: Iterator[RowRecord],
    output: Path | io.RawIOBase | io.BufferedIOBase,
) -> int:
    count = 0
    should_close = isinstance(output, Path)
    fh = open(output, "wb") if should_close else output

    try:
        for record in records:
            line = orjson.dumps(record_to_dict(record))
            fh.write(line)
            fh.write(b"\n")
            count += 1
    finally:
        if should_close:
            fh.close()

    return count
