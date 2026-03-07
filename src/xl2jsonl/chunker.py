from __future__ import annotations

import datetime
import logging
from dataclasses import dataclass
from typing import Iterator

from xl2jsonl.exceptions import NoHeaderError
from xl2jsonl.models import CellValue, Metadata, RowRecord, SheetData

logger = logging.getLogger(__name__)


@dataclass
class _TableRegion:
    """A rectangular sub-grid identified as a potential table."""

    rows: list[list[CellValue]]
    row_offset: int  # 0-based row offset in original sheet
    col_offset: int  # 0-based column offset in original sheet


def sheet_to_records(
    sheet: SheetData,
    filename: str,
    header_row: int | None = None,
    skip_empty_rows: bool = True,
) -> Iterator[RowRecord]:
    if not sheet.rows:
        return

    if header_row is not None:
        if header_row >= len(sheet.rows):
            raise NoHeaderError(
                f"Header row {header_row} out of range in sheet '{sheet.name}'"
            )
        yield from _table_to_records(
            sheet.rows, sheet, filename, header_row, skip_empty_rows, row_offset=0
        )
        return

    # Auto-detect table regions
    regions = _detect_table_regions(sheet.rows)
    found_any = False

    for region in regions:
        header_idx = _detect_header_row(region.rows)
        if header_idx is None:
            logger.debug(
                "Sheet '%s': no header in region (row_off=%d, col_off=%d) — skipping",
                sheet.name,
                region.row_offset + 1,
                region.col_offset + 1,
            )
            continue
        found_any = True
        yield from _table_to_records(
            region.rows,
            sheet,
            filename,
            header_idx,
            skip_empty_rows,
            row_offset=region.row_offset,
        )

    if not found_any:
        raise NoHeaderError(f"No header row found in sheet '{sheet.name}'")


def _table_to_records(
    rows: list[list[CellValue]],
    sheet: SheetData,
    filename: str,
    header_idx: int,
    skip_empty_rows: bool,
    row_offset: int = 0,
) -> Iterator[RowRecord]:
    raw_headers = rows[header_idx]
    headers = _normalize_headers(raw_headers)
    num_cols = len(headers)

    logger.debug(
        "Sheet '%s': header at row %d, %d columns: %s",
        sheet.name,
        row_offset + header_idx + 1,
        num_cols,
        headers,
    )

    for i, row in enumerate(rows[header_idx + 1 :]):
        if skip_empty_rows and _is_empty_row(row):
            continue

        padded = list(row) + [None] * max(0, num_cols - len(row))
        data = {h: _normalize_cell(padded[j]) for j, h in enumerate(headers)}
        metadata = Metadata(
            filename=filename,
            sheet_name=sheet.name,
            sheet_number=sheet.index + 1,
            row_number=row_offset + header_idx + i + 2,
        )
        yield RowRecord(data=data, metadata=metadata)


# ---------------------------------------------------------------------------
# Table region detection
# ---------------------------------------------------------------------------


def _detect_table_regions(rows: list[list[CellValue]]) -> list[_TableRegion]:
    """Detect rectangular table regions in a sheet.

    For sheets with a single column group, returns the whole sheet as one region
    (preserving existing single-table behavior).

    For sheets with multiple column groups (side-by-side tables separated by
    empty columns), splits into column groups, then further splits each group
    into row blocks separated by empty row gaps.
    """
    if not rows:
        return []

    col_groups = _find_column_groups(rows)

    if len(col_groups) <= 1:
        return [_TableRegion(rows=rows, row_offset=0, col_offset=0)]

    regions: list[_TableRegion] = []
    for col_start, col_end in col_groups:
        sub_rows = _extract_columns(rows, col_start, col_end)
        blocks = _find_row_blocks(sub_rows)
        for block_start, block_end in blocks:
            block_rows = sub_rows[block_start:block_end]
            if block_rows:
                regions.append(
                    _TableRegion(
                        rows=block_rows,
                        row_offset=block_start,
                        col_offset=col_start,
                    )
                )

    return regions


def _find_column_groups(
    rows: list[list[CellValue]],
    min_gap: int = 1,
) -> list[tuple[int, int]]:
    """Find groups of contiguous non-empty columns.

    Groups are separated by at least ``min_gap`` consecutive empty columns.
    A column is considered "occupied" only if it has data in a meaningful
    number of rows (at least 10% of the most-used column, minimum 2).
    This filters out noise from title/note rows that span across table
    boundaries.

    Returns list of (col_start, col_end) tuples (inclusive).
    """
    max_cols = max((len(r) for r in rows), default=0)
    if not max_cols:
        return []

    col_counts = [0] * max_cols
    for row in rows:
        for c in range(min(len(row), max_cols)):
            if row[c] is not None and (not isinstance(row[c], str) or row[c].strip()):
                col_counts[c] += 1

    max_count = max(col_counts) if col_counts else 0
    threshold = max(2, int(max_count * 0.1))
    col_occupied = [count >= threshold for count in col_counts]

    occupied = [c for c in range(max_cols) if col_occupied[c]]
    if not occupied:
        return []

    groups: list[tuple[int, int]] = []
    group_start = occupied[0]
    prev = occupied[0]

    for c in occupied[1:]:
        if c - prev > min_gap:  # gap of min_gap+ empty columns
            groups.append((group_start, prev))
            group_start = c
        prev = c

    groups.append((group_start, prev))
    return groups


def _extract_columns(
    rows: list[list[CellValue]],
    col_start: int,
    col_end: int,
) -> list[list[CellValue]]:
    """Extract a sub-grid of columns [col_start, col_end] from each row."""
    return [
        [row[c] if c < len(row) else None for c in range(col_start, col_end + 1)]
        for row in rows
    ]


def _find_row_blocks(
    rows: list[list[CellValue]],
    min_gap: int = 1,
) -> list[tuple[int, int]]:
    """Split rows into blocks separated by ``min_gap``+ consecutive empty rows.

    Returns list of (start, end) index pairs (end is exclusive).
    """
    blocks: list[tuple[int, int]] = []
    block_start: int | None = None
    consecutive_empty = 0

    for i, row in enumerate(rows):
        if _is_empty_row(row):
            consecutive_empty += 1
        else:
            if block_start is None:
                block_start = i
            elif consecutive_empty >= min_gap:
                blocks.append((block_start, i - consecutive_empty))
                block_start = i
            consecutive_empty = 0

    if block_start is not None:
        end = len(rows) - consecutive_empty
        if end > block_start:
            blocks.append((block_start, end))

    return blocks


# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------


def _detect_header_row(rows: list[list[CellValue]]) -> int | None:
    """Find the best row that looks like a table header.

    Handles multi-row headers (e.g. a category row with merged cells above
    the actual column headers) by checking if a subsequent row has better
    column coverage when the first candidate has duplicate values.
    """
    for idx, row in enumerate(rows):
        if _is_empty_row(row):
            continue
        if not _is_header_candidate(row, row_num=idx + 1):
            continue

        # Check for multi-row header: if this candidate has duplicate values
        # (from merged category cells), a subsequent row may be the real header.
        non_empty = [c for c in row if c is not None and c != ""]
        distinct = len(set(str(c).strip() for c in non_empty))
        if distinct < len(non_empty):
            better = _find_best_header_in_run(rows, idx)
            if better != idx:
                logger.debug(
                    "Row %d has duplicates (%d/%d distinct), "
                    "preferring row %d with better coverage",
                    idx + 1,
                    distinct,
                    len(non_empty),
                    better + 1,
                )
                return better

        return idx
    return None


def _find_best_header_in_run(
    rows: list[list[CellValue]], start: int
) -> int:
    """Among consecutive header candidates starting at *start*, return the
    index of the one with the most non-empty cells (best column coverage)."""
    best_idx = start
    best_count = _non_empty_count(rows[start])

    for j in range(start + 1, min(start + 5, len(rows))):
        r = rows[j]
        if _is_empty_row(r):
            continue
        if not _is_header_candidate(r, row_num=j + 1):
            break
        count = _non_empty_count(r)
        if count > best_count:
            best_count = count
            best_idx = j

    return best_idx


def _non_empty_count(row: list[CellValue]) -> int:
    return sum(1 for c in row if c is not None and c != "")


def _is_header_candidate(
    row: list[CellValue],
    row_num: int | None = None,
) -> bool:
    """Check whether a row looks like a table header.

    A header candidate must have:
    - Majority (50%+) of non-empty cells are strings
    - More than one distinct value (to skip resolved merged title rows)
    - Reasonable diversity of values (to skip partially merged rows)
    """
    non_empty = [c for c in row if c is not None and c != ""]
    if not non_empty:
        return False

    string_count = sum(1 for c in non_empty if isinstance(c, str))
    if string_count / len(non_empty) < 0.5:
        return False

    distinct_values = len(set(str(c).strip() for c in non_empty))
    prefix = f"Row {row_num}: " if row_num else ""

    if distinct_values == 1 and len(non_empty) > 1:
        logger.debug(
            "%sskipping as merged title (all %d cells have same value)",
            prefix,
            len(non_empty),
        )
        return False

    if len(non_empty) >= 4 and distinct_values / len(non_empty) < 0.5:
        logger.debug(
            "%sskipping as likely merged title (%d distinct / %d cells)",
            prefix,
            distinct_values,
            len(non_empty),
        )
        return False

    return True


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _normalize_headers(raw: list[CellValue]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}

    for i, cell in enumerate(raw):
        if cell is None or (isinstance(cell, str) and cell.strip() == ""):
            name = f"column_{i + 1}"
        else:
            name = str(cell).strip()
            # Collapse internal newlines/whitespace
            name = " ".join(name.split())

        # Deduplicate
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1

        headers.append(name)

    return headers


def _is_empty_row(row: list[CellValue]) -> bool:
    return all(c is None or (isinstance(c, str) and c.strip() == "") for c in row)


def _normalize_cell(value: CellValue) -> CellValue:
    if isinstance(value, datetime.datetime):
        return value.isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    return value
