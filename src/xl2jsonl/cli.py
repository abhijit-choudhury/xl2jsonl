from __future__ import annotations

import logging
from pathlib import Path

import click

from xl2jsonl import convert


def _parse_sheets(raw: tuple[str, ...]) -> list[str | int]:
    result: list[str | int] = []
    for s in raw:
        try:
            result.append(int(s))
        except ValueError:
            result.append(s)
    return result


def _setup_logging(verbosity: int) -> None:
    level = logging.WARNING
    if verbosity == 1:
        level = logging.INFO
    elif verbosity >= 2:
        level = logging.DEBUG
    logging.basicConfig(
        level=level,
        format="%(levelname)s %(name)s: %(message)s",
    )


@click.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "-o",
    "--output",
    type=click.Path(path_type=Path),
    default=None,
    help="Output JSONL file. Defaults to <input>.jsonl",
)
@click.option(
    "-s",
    "--sheet",
    "sheets",
    multiple=True,
    help="Sheet name or 0-based index. Repeatable. Default: all sheets.",
)
@click.option(
    "--engine",
    type=click.Choice(["auto", "calamine", "openpyxl"]),
    default="auto",
    help="Reading engine (default: auto)",
)
@click.option(
    "--header-row",
    type=int,
    default=None,
    help="0-based header row index. Default: auto-detect.",
)
@click.option(
    "--skip-empty/--keep-empty",
    default=True,
    help="Skip empty rows (default: skip)",
)
@click.option(
    "-v",
    "--verbose",
    count=True,
    help="-v for INFO, -vv for DEBUG",
)
def main(
    input_file: Path,
    output: Path | None,
    sheets: tuple[str, ...],
    engine: str,
    header_row: int | None,
    skip_empty: bool,
    verbose: int,
) -> None:
    """Convert an Excel file to JSONL."""
    _setup_logging(verbose)

    if output is None:
        output = input_file.with_suffix(".jsonl")

    parsed_sheets = _parse_sheets(sheets) if sheets else None

    count = convert(
        input_file,
        output,
        sheets=parsed_sheets,
        engine=engine,
        header_row=header_row,
        skip_empty_rows=skip_empty,
    )
    click.echo(f"Wrote {count} records to {output}")
