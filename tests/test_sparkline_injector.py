"""Verify the post-save sparkline XML injector lands correctly in .xlsx."""
import re
import zipfile
from pathlib import Path

from openpyxl import Workbook

from xl import sparkline_injector


def _make_simple_workbook(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Plan"
    # Write 10 numbers into A1:A10 to act as the sparkline data range
    for i, v in enumerate([100, 102, 99, 105, 110, 108, 115, 120, 118, 125], start=1):
        ws.cell(row=i, column=1, value=v)
    path = tmp_path / "wb.xlsx"
    wb.save(path)
    return path


def test_inject_adds_sparkline_to_correct_sheet(tmp_path):
    path = _make_simple_workbook(tmp_path)
    sparkline_injector.inject_line_sparklines(path, {
        "Daily Plan": [
            {"data_range": "'Daily Plan'!A1:A10", "target_cell": "B1"},
        ],
    })

    with zipfile.ZipFile(path) as z:
        # Find the worksheet xml that should now contain the sparkline
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
    assert "sparkline" in sheet_xml.lower()
    assert "'Daily Plan'!A1:A10" in sheet_xml
    assert ">B1<" in sheet_xml  # target cell


def test_inject_no_op_when_specs_empty(tmp_path):
    path = _make_simple_workbook(tmp_path)
    before = path.read_bytes()
    sparkline_injector.inject_line_sparklines(path, {})
    after = path.read_bytes()
    assert before == after


def test_inject_handles_multiple_sparklines(tmp_path):
    path = _make_simple_workbook(tmp_path)
    sparkline_injector.inject_line_sparklines(path, {
        "Daily Plan": [
            {"data_range": "'Daily Plan'!A1:A10", "target_cell": "B1"},
            {"data_range": "'Daily Plan'!A1:A10", "target_cell": "B2"},
            {"data_range": "'Daily Plan'!A1:A10", "target_cell": "B3"},
        ],
    })

    with zipfile.ZipFile(path) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
    # Three <x14:sparkline> elements (each contains an <xm:sqref> child)
    assert sheet_xml.count("</x14:sparkline>") == 3


def test_inject_preserves_other_worksheet_content(tmp_path):
    path = _make_simple_workbook(tmp_path)
    sparkline_injector.inject_line_sparklines(path, {
        "Daily Plan": [{"data_range": "'Daily Plan'!A1:A10", "target_cell": "B1"}],
    })

    with zipfile.ZipFile(path) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
    # Original cell values should still be present
    assert "<c " in sheet_xml or "<row" in sheet_xml  # cells/rows preserved


def test_resolve_sheet_files_handles_xl_prefix(tmp_path):
    path = _make_simple_workbook(tmp_path)
    mapping = sparkline_injector._resolve_sheet_files(path)
    # No doubled "xl/xl/"
    for f in mapping.values():
        assert not f.startswith("xl/xl/")
        assert f.startswith("xl/")
