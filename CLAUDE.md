# xl2jsonl

Convert Excel (.xlsx, .xls, .xlsb) and CSV (.csv, .tsv) files to JSONL.

## Quick Reference

- **Language**: Python 3.10+
- **Package manager**: uv
- **Entry point**: `src/xl2jsonl/cli.py` (Click CLI), `src/xl2jsonl/__init__.py` (Python API)
- **Test runner**: `uv run python -m pytest` (pytest not directly on PATH, use `python -m`)
- **Lint**: `uv run ruff check src/ tests/`

## Project Structure

```
src/xl2jsonl/
  __init__.py    - Public API: convert(), iter_records()
  cli.py         - Click CLI entry point
  loader.py      - Multi-format loader: Excel (calamine + openpyxl), CSV/TSV (stdlib csv)
  chunker.py     - Header detection, row-to-dict conversion, cell normalization
  writer.py      - JSONL streaming output via orjson
  models.py      - Data classes: SheetData, RowRecord, Metadata, CellValue
  exceptions.py  - Xl2JsonlError, LoaderError, EmptySheetError, NoHeaderError
tests/
  conftest.py    - Programmatic xlsx fixture generators (openpyxl)
  test_*.py      - Unit and integration tests
```

## Architecture

1. **Loader** (`loader.py`): Routes by file extension. Excel (.xlsx/.xls/.xlsb) read via calamine; .xlsx gets openpyxl merge detection in `auto` mode. CSV/TSV read via stdlib `csv` module with basic type inference (int, float, bool).
2. **Chunker** (`chunker.py`): Detects table regions, auto-detects headers, converts rows to key-value dicts.
   - **Table region detection**: For sheets with multiple column groups (side-by-side tables separated by empty columns), splits into independent regions. Each column group is further split into row blocks by empty row gaps. Single-column-group sheets use the whole sheet as one region (backward compatible).
   - **Header detection**: First row where 50%+ non-empty cells are strings AND values are sufficiently distinct (skips merged title rows).
3. **Writer** (`writer.py`): Streams `{content, metadata}` JSONL records via orjson.

## Key Design Decisions

- **Multi-table sheets**: Column groups are detected using count-based occupancy (a column needs data in 10%+ of rows to count as occupied, filtering out noise from titles/notes that span across table boundaries). Groups are split by 1+ empty columns. Row blocks within each group are split by 1+ empty rows.
- **Merged title detection**: Header auto-detection skips rows where most non-empty cells share the same value (resolved merged title/description rows).
- **Multi-row headers**: When a header candidate has duplicate values (merged category row), scans subsequent rows for one with better column coverage.
- **Format support**: Merge resolution only works for .xlsx (openpyxl limitation). For .xls/.xlsb, calamine-only (no merge detection). CSV/TSV loaded via stdlib csv with type inference.
- `--header-row` (0-based) CLI flag overrides auto-detection when heuristics fail.
- Duplicate headers from merged columns get `_2`, `_3` suffixes.
- Empty leading rows are skipped during header detection.
- Regions with no detected header are silently skipped (logged at DEBUG level).

## Running

```bash
uv run xl2jsonl input.xlsx -o output.jsonl    # CLI
uv run python -m pytest -v                     # Tests (all 39)
uv run python -m pytest tests/test_chunker.py  # Single file
uv run ruff check src/ tests/                  # Lint
```

## Testing Conventions

- Fixtures in `conftest.py` generate xlsx files programmatically using openpyxl
- No test data files checked in — everything is generated at test time
- Tests cover: loader engines, merge resolution, header detection, cell normalization, CLI flags, end-to-end xlsx→jsonl
