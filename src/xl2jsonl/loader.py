from __future__ import annotations

import logging
from pathlib import Path

import openpyxl
from python_calamine import CalamineWorkbook

from xl2jsonl.exceptions import LoaderError
from xl2jsonl.models import CellValue, SheetData

logger = logging.getLogger(__name__)


def load_workbook(
    path: Path,
    sheets: list[str | int] | None = None,
    engine: str = "auto",
) -> list[SheetData]:
    path = Path(path)
    if not path.exists():
        raise LoaderError(f"File not found: {path}")

    calamine_wb = CalamineWorkbook.from_path(str(path))
    all_sheet_names = calamine_wb.sheet_names

    target_sheets = _resolve_sheet_selection(all_sheet_names, sheets)

    openpyxl_wb = None
    if engine in ("auto", "openpyxl"):
        openpyxl_wb = openpyxl.load_workbook(str(path), data_only=True)

    results: list[SheetData] = []
    try:
        for idx, name in target_sheets:
            if engine == "calamine":
                results.append(_load_sheet_calamine(calamine_wb, name, idx))
            elif engine == "openpyxl":
                assert openpyxl_wb is not None
                results.append(_load_sheet_openpyxl(openpyxl_wb, name, idx))
            else:  # auto
                assert openpyxl_wb is not None
                ws = openpyxl_wb[name]
                merged_ranges = list(ws.merged_cells.ranges)
                if merged_ranges:
                    logger.info(
                        "Sheet '%s': %d merged region(s), using openpyxl",
                        name,
                        len(merged_ranges),
                    )
                    results.append(_load_sheet_openpyxl(openpyxl_wb, name, idx))
                else:
                    logger.debug("Sheet '%s': no merges, using calamine", name)
                    results.append(_load_sheet_calamine(calamine_wb, name, idx))
    finally:
        if openpyxl_wb:
            openpyxl_wb.close()

    return results


def _resolve_sheet_selection(
    all_names: list[str],
    sheets: list[str | int] | None,
) -> list[tuple[int, str]]:
    if sheets is None:
        return list(enumerate(all_names))

    result: list[tuple[int, str]] = []
    for s in sheets:
        if isinstance(s, int):
            if s < 0 or s >= len(all_names):
                raise LoaderError(
                    f"Sheet index {s} out of range (0-{len(all_names) - 1})"
                )
            result.append((s, all_names[s]))
        else:
            if s not in all_names:
                raise LoaderError(
                    f"Sheet '{s}' not found. Available: {all_names}"
                )
            result.append((all_names.index(s), s))
    return result


def _load_sheet_calamine(
    wb: CalamineWorkbook,
    sheet_name: str,
    sheet_idx: int,
) -> SheetData:
    data = wb.get_sheet_by_name(sheet_name).to_python()
    rows: list[list[CellValue]] = [list(row) for row in data]
    return SheetData(name=sheet_name, index=sheet_idx, rows=rows)


def _load_sheet_openpyxl(
    wb: openpyxl.Workbook,
    sheet_name: str,
    sheet_idx: int,
) -> SheetData:
    ws = wb[sheet_name]
    rows, had_merges = _resolve_merges(ws)
    return SheetData(
        name=sheet_name, index=sheet_idx, rows=rows, had_merged_cells=had_merges
    )


def _resolve_merges(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> tuple[list[list[CellValue]], bool]:
    merged_ranges = list(ws.merged_cells.ranges)
    had_merges = len(merged_ranges) > 0

    # Collect merge info before unmerging
    merge_fills: list[tuple[object, CellValue]] = []
    for mr in merged_ranges:
        value = ws.cell(mr.min_row, mr.min_col).value
        merge_fills.append((mr, value))

    # Unmerge all ranges
    for mr, _ in merge_fills:
        ws.unmerge_cells(str(mr))

    # Fill every cell in each former merged range
    for mr, value in merge_fills:
        for row in range(mr.min_row, mr.max_row + 1):
            for col in range(mr.min_col, mr.max_col + 1):
                ws.cell(row, col).value = value

    # Convert to 2D list
    rows: list[list[CellValue]] = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))

    return rows, had_merges
