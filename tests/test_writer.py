import io
import json

from xl2jsonl.models import Metadata, RowRecord
from xl2jsonl.writer import record_to_dict, write_jsonl


def test_record_to_dict():
    record = RowRecord(
        data={"Name": "Alice", "Age": 30},
        metadata=Metadata(
            filename="test.xlsx",
            sheet_name="Sheet1",
            sheet_number=1,
            row_number=2,
        ),
    )
    result = record_to_dict(record)
    assert result == {
        "content": {"Name": "Alice", "Age": 30},
        "metadata": {
            "filename": "test.xlsx",
            "sheet_name": "Sheet1",
            "sheet_number": 1,
            "row_number": 2,
        },
    }


def test_write_jsonl_to_file(tmp_path):
    records = iter(
        [
            RowRecord(
                data={"X": 1},
                metadata=Metadata("f.xlsx", "S1", 1, 2),
            ),
            RowRecord(
                data={"X": 2},
                metadata=Metadata("f.xlsx", "S1", 1, 3),
            ),
        ]
    )

    output = tmp_path / "out.jsonl"
    count = write_jsonl(records, output)
    assert count == 2

    lines = output.read_text().strip().split("\n")
    assert len(lines) == 2
    obj = json.loads(lines[0])
    assert obj["content"]["X"] == 1
    assert obj["metadata"]["sheet_name"] == "S1"


def test_write_jsonl_to_stream():
    records = iter(
        [
            RowRecord(
                data={"Val": "hello"},
                metadata=Metadata("f.xlsx", "Sheet1", 1, 2),
            ),
        ]
    )

    buf = io.BytesIO()
    count = write_jsonl(records, buf)
    assert count == 1

    buf.seek(0)
    obj = json.loads(buf.read())
    assert obj["content"]["Val"] == "hello"


def test_write_jsonl_none_values(tmp_path):
    records = iter(
        [
            RowRecord(
                data={"A": None, "B": True, "C": 3.14},
                metadata=Metadata("f.xlsx", "S1", 1, 2),
            ),
        ]
    )
    output = tmp_path / "out.jsonl"
    write_jsonl(records, output)
    obj = json.loads(output.read_text().strip())
    assert obj["content"]["A"] is None
    assert obj["content"]["B"] is True
    assert obj["content"]["C"] == 3.14
