from __future__ import annotations

from dataclasses import dataclass
from typing import Any

CellValue = str | int | float | bool | None


@dataclass(frozen=True)
class Metadata:
    filename: str
    sheet_name: str
    sheet_number: int  # 1-based
    row_number: int  # 1-based, original Excel row


@dataclass
class SheetData:
    name: str
    index: int  # 0-based sheet index
    rows: list[list[CellValue]]
    had_merged_cells: bool = False


@dataclass
class RowRecord:
    data: dict[str, CellValue]
    metadata: Metadata
