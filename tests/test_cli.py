import json

from click.testing import CliRunner

from xl2jsonl.cli import main


def test_cli_basic(simple_xlsx, tmp_path):
    output = tmp_path / "output.jsonl"
    runner = CliRunner()
    result = runner.invoke(main, [str(simple_xlsx), "-o", str(output)])
    assert result.exit_code == 0
    assert "Wrote" in result.output

    lines = output.read_text().strip().split("\n")
    assert len(lines) == 5  # 3 People + 2 Scores
    obj = json.loads(lines[0])
    assert "content" in obj
    assert "metadata" in obj


def test_cli_default_output(simple_xlsx):
    runner = CliRunner()
    result = runner.invoke(main, [str(simple_xlsx)])
    assert result.exit_code == 0
    default_output = simple_xlsx.with_suffix(".jsonl")
    assert default_output.exists()
    default_output.unlink()


def test_cli_sheet_filter(simple_xlsx, tmp_path):
    output = tmp_path / "output.jsonl"
    runner = CliRunner()
    result = runner.invoke(main, [str(simple_xlsx), "-o", str(output), "-s", "People"])
    assert result.exit_code == 0
    lines = output.read_text().strip().split("\n")
    assert len(lines) == 3  # Only People sheet


def test_cli_engine_flag(simple_xlsx, tmp_path):
    output = tmp_path / "output.jsonl"
    runner = CliRunner()
    result = runner.invoke(
        main, [str(simple_xlsx), "-o", str(output), "--engine", "openpyxl"]
    )
    assert result.exit_code == 0


def test_cli_verbose(simple_xlsx, tmp_path):
    output = tmp_path / "output.jsonl"
    runner = CliRunner()
    result = runner.invoke(main, [str(simple_xlsx), "-o", str(output), "-vv"])
    assert result.exit_code == 0
