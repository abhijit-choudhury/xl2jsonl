# xl2jsonl

Convert Excel files (`.xlsx`) to JSONL with robust handling of merged cells, multiple sheets, and mixed data types. Each output line is a JSON object containing row data as key-value pairs plus metadata for traceability.

## Why xl2jsonl?

Excel spreadsheets in the wild are messy — merged headers spanning multiple columns, merged row labels, empty leading rows, inconsistent types. Most conversion tools break on these. xl2jsonl handles them correctly by using a dual-engine approach: **calamine** (Rust-backed) for speed on clean sheets, with automatic fallback to **openpyxl** when merged cells are detected.

## Output Format

Each line in the output JSONL file is a self-contained JSON object:

```json
{"content": {"Name": "Alice", "Age": 30, "Department": "Engineering"}, "metadata": {"filename": "staff.xlsx", "sheet_name": "Employees", "sheet_number": 1, "row_number": 2}}
{"content": {"Name": "Bob", "Age": 25, "Department": "Design"}, "metadata": {"filename": "staff.xlsx", "sheet_name": "Employees", "sheet_number": 1, "row_number": 3}}
```

| Field | Description |
|---|---|
| `content` | Key-value pairs mapping column headers to cell values for that row |
| `metadata.filename` | Original Excel filename |
| `metadata.sheet_name` | Name of the worksheet |
| `metadata.sheet_number` | 1-based sheet index |
| `metadata.row_number` | 1-based row number in the original Excel file |

## Features

- **Merged cell resolution** — Merged columns, rows, and blocks are filled so every cell in a merged range gets the top-left cell's value
- **Duplicate header handling** — Merged headers that produce duplicates get `_2`, `_3` suffixes (e.g., `Revenue`, `Revenue_2`, `Revenue_3`)
- **Auto header detection** — Skips empty leading rows, merged title rows, and finds the first row where the majority of non-empty cells are strings with sufficient diversity
- **Multi-table sheet support** — Automatically detects side-by-side tables (separated by empty columns) and stacked tables (separated by empty rows) within a single sheet, processing each independently
- **Multi-sheet support** — Processes all sheets by default, or filter by name/index
- **Mixed type handling** — Dates and datetimes convert to ISO 8601 strings, strings are stripped, numbers and booleans pass through, `None` stays `None`
- **Empty row skipping** — Empty rows are skipped by default (configurable)
- **Short row padding** — Rows shorter than the header are padded with `None`
- **Streaming output** — JSONL is written one record at a time via `orjson` for low memory overhead
- **Dual engine** — calamine for speed on clean sheets, openpyxl fallback for merged cells (configurable)

## Dependencies

| Package | Purpose |
|---|---|
| [python-calamine](https://pypi.org/project/python-calamine/) | Fast Rust-backed Excel reader (default engine) |
| [openpyxl](https://pypi.org/project/openpyxl/) | Merged cell detection and resolution |
| [click](https://pypi.org/project/click/) | CLI framework |
| [orjson](https://pypi.org/project/orjson/) | Fast JSON serialization with native date handling |

Dev dependencies: `pytest`, `pytest-cov`, `ruff`

## Setup

Requires **Python 3.10+** and **[uv](https://docs.astral.sh/uv/)**.

```bash
# Clone and install
git clone <repo-url>
cd excel-chunker
uv sync
```

This creates a virtual environment in `.venv/` and installs all dependencies including dev tools.

## Usage

### CLI

```bash
# Convert all sheets (output defaults to <input>.jsonl)
uv run xl2jsonl data.xlsx

# Specify output file
uv run xl2jsonl data.xlsx -o output.jsonl

# Select specific sheets (by name or 0-based index, repeatable)
uv run xl2jsonl data.xlsx -s "Revenue" -s "Costs"
uv run xl2jsonl data.xlsx -s 0 -s 2

# Force a specific engine
uv run xl2jsonl data.xlsx --engine openpyxl
uv run xl2jsonl data.xlsx --engine calamine

# Set header row explicitly (0-based index)
uv run xl2jsonl data.xlsx --header-row 2

# Keep empty rows in output
uv run xl2jsonl data.xlsx --keep-empty

# Verbose logging (-v for INFO, -vv for DEBUG)
uv run xl2jsonl data.xlsx -vv
```

Full CLI help:

```
Usage: xl2jsonl [OPTIONS] INPUT_FILE

  Convert an Excel file to JSONL.

Options:
  -o, --output PATH              Output JSONL file. Defaults to <input>.jsonl
  -s, --sheet TEXT               Sheet name or 0-based index. Repeatable.
  --engine [auto|calamine|openpyxl]
                                 Reading engine (default: auto)
  --header-row INTEGER           0-based header row index. Default: auto-detect.
  --skip-empty / --keep-empty    Skip empty rows (default: skip)
  -v, --verbose                  -v for INFO, -vv for DEBUG
  --help                         Show this message and exit.
```

### Python API

```python
import xl2jsonl

# Convert to JSONL file — returns row count
count = xl2jsonl.convert("data.xlsx", "output.jsonl")
print(f"Wrote {count} records")

# Convert and get results in memory — returns list of dicts
records = xl2jsonl.convert("data.xlsx")
for r in records:
    print(r["content"], r["metadata"])

# Lazy iterator for large files
for record in xl2jsonl.iter_records("data.xlsx", sheets=["Sheet1"]):
    print(record.data)        # dict of header -> value
    print(record.metadata)    # Metadata(filename, sheet_name, sheet_number, row_number)

# All options
count = xl2jsonl.convert(
    "data.xlsx",
    "output.jsonl",
    sheets=["Revenue", 2],       # filter by name or index
    engine="auto",               # "auto" | "calamine" | "openpyxl"
    header_row=None,             # None for auto-detect, or 0-based int
    skip_empty_rows=True,        # skip rows where all cells are empty
)
```

## How Merged Cells Are Handled

### Merged column headers

If a header cell spans columns B through D with value "Revenue":

| | A | B | C | D |
|---|---|---|---|---|
| **Header** | ID | Revenue (merged B:D) | | Profit |
| **Row 1** | 1 | 100 | 200 | 50 |

After resolution, the headers become `["ID", "Revenue", "Revenue_2", "Revenue_3", "Profit"]`, and each data cell maps to its respective deduplicated header.

### Merged row labels

If a cell in column A spans rows 4-5 with value "Total":

| | A | B |
|---|---|---|
| **Row 4** | Total (merged A4:A5) | 250 |
| **Row 5** | | 300 |

After resolution, both rows get `"Total"` as their value for that column.

### Merged blocks

A 2x3 merged region is filled entirely with the top-left cell's value.

### Multi-table sheets

Sheets with multiple tables laid out side-by-side or stacked vertically are automatically detected and split into independent regions:

| | A | B | C | | E | F | G |
|---|---|---|---|---|---|---|---|
| **Row 1** | Account: | | ACME Corp | | *Note: ...* | | |
| | | | | | | | |
| **Row 3** | **Name** | **Score** | **Grade** | | **Region** | **Revenue** | **Growth** |
| **Row 4** | Alice | 95 | A | | EMEA | 1.2M | 5% |
| **Row 5** | Bob | 88 | B | | APAC | 800K | 12% |

This produces two separate tables: one with headers `Name, Score, Grade` and another with `Region, Revenue, Growth`. The metadata row and note are detected as separate regions and skipped (no valid header found).

**How it works:**
1. Columns are grouped by occupancy — a column needs data in a meaningful number of rows (10%+ of the busiest column) to count, filtering out noise from titles/notes that span across table boundaries
2. Groups are split where 1+ empty columns separate them
3. Within each column group, row blocks are split by empty row gaps
4. Each block gets independent header detection and processing
5. Blocks with no detected header are silently skipped

### Engine selection (`auto` mode)

1. All sheets are read with calamine (fast)
2. Each sheet is probed for merged cell ranges via openpyxl
3. Sheets with merges are re-read through openpyxl with full merge resolution
4. Sheets without merges use the calamine data as-is

This gives you the speed of calamine on clean sheets while correctly handling merges where they exist.

## Project Structure

```
excel-chunker/
├── pyproject.toml              # Package config, dependencies, CLI entry point
├── src/xl2jsonl/
│   ├── __init__.py             # Public API: convert(), iter_records()
│   ├── models.py               # Data classes: SheetData, RowRecord, Metadata
│   ├── exceptions.py           # Xl2JsonlError, LoaderError, EmptySheetError, NoHeaderError
│   ├── loader.py               # Dual-engine Excel reader + merge resolution
│   ├── chunker.py              # Header detection, row-to-dict, cell normalization
│   ├── writer.py               # JSONL streaming output via orjson
│   └── cli.py                  # Click CLI entry point
└── tests/
    ├── conftest.py             # Programmatic xlsx fixture generators
    ├── test_loader.py          # Loader unit tests (engines, merges, sheet selection)
    ├── test_chunker.py         # Chunker unit tests (headers, normalization, edge cases)
    ├── test_writer.py          # Writer unit tests (JSONL format, streaming)
    ├── test_cli.py             # CLI tests via Click CliRunner
    └── test_integration.py     # End-to-end tests (xlsx -> jsonl -> verify)
```

## Running Tests

```bash
# Run all tests
uv run pytest

# Verbose output
uv run pytest -v

# With coverage
uv run pytest --cov=xl2jsonl

# Run a specific test file
uv run pytest tests/test_loader.py
```

## Limitations

- Only `.xlsx` files are supported (not `.xls` or `.xlsb`)
- The entire sheet grid is materialized in memory during loading (Excel files are not row-streamable); the chunker and writer stages stream lazily
- Header auto-detection uses a heuristic (first row where 50%+ of non-empty cells are strings) — use `--header-row` to override if it guesses wrong
- Formulas are evaluated to their cached values (`data_only=True` in openpyxl) — if a file was never opened in Excel after formula changes, values may be stale

## License

MIT
