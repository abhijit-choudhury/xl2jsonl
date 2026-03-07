import json

import xl2jsonl


def test_end_to_end_simple(simple_xlsx, tmp_path):
    output = tmp_path / "out.jsonl"
    count = xl2jsonl.convert(simple_xlsx, output)
    assert count == 5  # 3 People + 2 Scores

    lines = output.read_text().strip().split("\n")
    records = [json.loads(line) for line in lines]

    # Check first record
    assert records[0]["content"]["Name"] == "Alice"
    assert records[0]["metadata"]["filename"] == "simple.xlsx"
    assert records[0]["metadata"]["sheet_name"] == "People"
    assert records[0]["metadata"]["sheet_number"] == 1

    # Check last record (from Scores sheet)
    assert records[4]["content"]["Subject"] == "English"
    assert records[4]["metadata"]["sheet_name"] == "Scores"
    assert records[4]["metadata"]["sheet_number"] == 2


def test_end_to_end_return_dicts(simple_xlsx):
    result = xl2jsonl.convert(simple_xlsx)
    assert isinstance(result, list)
    assert len(result) == 5
    assert result[0]["content"]["Name"] == "Alice"
    assert result[0]["metadata"]["sheet_name"] == "People"


def test_end_to_end_merged(merged_cells_xlsx, tmp_path):
    output = tmp_path / "out.jsonl"
    count = xl2jsonl.convert(merged_cells_xlsx, output)
    assert count == 4

    lines = output.read_text().strip().split("\n")
    records = [json.loads(line) for line in lines]

    # Merged header "Revenue" spans B and C -> "Revenue" and "Revenue_2"
    assert "Revenue" in records[0]["content"]
    assert "Revenue_2" in records[0]["content"]

    # Merged rows: last two records should have "Total" in ID
    assert records[2]["content"]["ID"] == "Total"
    assert records[3]["content"]["ID"] == "Total"


def test_end_to_end_empty_rows(empty_rows_xlsx, tmp_path):
    output = tmp_path / "out.jsonl"
    count = xl2jsonl.convert(empty_rows_xlsx, output)
    # Should auto-detect header at row 3 and skip empty row 5
    assert count == 2

    lines = output.read_text().strip().split("\n")
    records = [json.loads(line) for line in lines]
    assert records[0]["content"]["Product"] == "Widget"
    assert records[1]["content"]["Product"] == "Gadget"


def test_end_to_end_mixed_types(mixed_types_xlsx, tmp_path):
    output = tmp_path / "out.jsonl"
    count = xl2jsonl.convert(mixed_types_xlsx, output)
    assert count == 2

    lines = output.read_text().strip().split("\n")
    records = [json.loads(line) for line in lines]

    alice = records[0]["content"]
    assert alice["Active"] is True
    assert alice["Score"] == 95.5
    assert alice["Notes"] == "Line1\nLine2"

    bob = records[1]["content"]
    assert bob["Active"] is False
    assert bob["Notes"] == "padded"  # stripped


def test_end_to_end_merged_block(merged_block_xlsx, tmp_path):
    output = tmp_path / "out.jsonl"
    count = xl2jsonl.convert(merged_block_xlsx, output)
    assert count == 2

    lines = output.read_text().strip().split("\n")
    records = [json.loads(line) for line in lines]

    # Both rows should have "Block" in Header1 and Header2
    assert records[0]["content"]["Header1"] == "Block"
    assert records[0]["content"]["Header2"] == "Block"
    assert records[1]["content"]["Header1"] == "Block"
    assert records[1]["content"]["Header2"] == "Block"


def test_iter_records(simple_xlsx):
    records = list(xl2jsonl.iter_records(simple_xlsx, sheets=["People"]))
    assert len(records) == 3
    assert records[0].data["Name"] == "Alice"
    assert records[0].metadata.row_number >= 2
