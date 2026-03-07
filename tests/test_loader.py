from xl2jsonl.loader import load_workbook


def test_load_simple_calamine(simple_xlsx):
    sheets = load_workbook(simple_xlsx, engine="calamine")
    assert len(sheets) == 2
    assert sheets[0].name == "People"
    assert sheets[0].rows[0] == ["Name", "Age", "City"]
    assert sheets[0].rows[1] == ["Alice", 30.0, "London"]
    assert len(sheets[0].rows) == 4  # header + 3 data rows


def test_load_simple_openpyxl(simple_xlsx):
    sheets = load_workbook(simple_xlsx, engine="openpyxl")
    assert len(sheets) == 2
    assert sheets[0].rows[0] == ["Name", "Age", "City"]


def test_load_auto_no_merges(simple_xlsx):
    sheets = load_workbook(simple_xlsx, engine="auto")
    assert len(sheets) == 2
    assert not sheets[0].had_merged_cells


def test_load_merged_cells_openpyxl(merged_cells_xlsx):
    sheets = load_workbook(merged_cells_xlsx, engine="openpyxl")
    assert len(sheets) == 1
    ws = sheets[0]
    assert ws.had_merged_cells

    # Merged header B1:C1 should be filled
    assert ws.rows[0][1] == "Revenue"
    assert ws.rows[0][2] == "Revenue"

    # Merged rows A4:A5 should both have "Total"
    assert ws.rows[3][0] == "Total"
    assert ws.rows[4][0] == "Total"


def test_load_auto_detects_merges(merged_cells_xlsx):
    sheets = load_workbook(merged_cells_xlsx, engine="auto")
    assert sheets[0].had_merged_cells
    # Should have used openpyxl and filled merges
    assert sheets[0].rows[0][2] == "Revenue"


def test_sheet_selection_by_name(simple_xlsx):
    sheets = load_workbook(simple_xlsx, sheets=["Scores"])
    assert len(sheets) == 1
    assert sheets[0].name == "Scores"


def test_sheet_selection_by_index(simple_xlsx):
    sheets = load_workbook(simple_xlsx, sheets=[1])
    assert len(sheets) == 1
    assert sheets[0].name == "Scores"


def test_merged_block(merged_block_xlsx):
    sheets = load_workbook(merged_block_xlsx, engine="openpyxl")
    ws = sheets[0]
    # 2x2 block A2:B3 should all have "Block"
    assert ws.rows[1][0] == "Block"
    assert ws.rows[1][1] == "Block"
    assert ws.rows[2][0] == "Block"
    assert ws.rows[2][1] == "Block"
