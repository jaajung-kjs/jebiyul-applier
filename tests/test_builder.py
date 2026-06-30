"""Tests for src/builder.py — Output builder (Task 8).

TDD: test_build_writes_rates and test_cell_map_matches_template_layout run first.
"""
import os
import openpyxl
import pytest

from src.builder import build_output, CELL_MAP, ITEMS

TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "assets",
    "template_적용근거.xlsx",
)

SAMPLE_RATES = {
    "간접노무비": 0.191,
    "공구손료": 0.03,
    "산재보험료": 0.03656,
    "고용보험료": 0.0101,
    "건강보험료": 0.03545,
    "연금보험료": 0.045,
    "퇴직공제부금비": 0.023,
    "노인장기요양보험료": 0.1295,
    "산업안전보건관리비": 0.0185,
    "기타경비": 0.055,
    "일반관리비": 0.08,
    "이윤": 0.15,
}


def test_build_writes_rates(tmp_path):
    """build_output copies the template and writes rate values into the I column."""
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"))
    wb = openpyxl.load_workbook(out)
    ws = wb["적용근거"]
    vals = [
        ws.cell(r, c).value
        for r in range(1, ws.max_row + 1)
        for c in range(1, ws.max_column + 1)
    ]
    assert 0.191 in vals
    assert 0.03656 in vals


def test_build_returns_path(tmp_path):
    """build_output must return the output file path."""
    out_path = str(tmp_path / "out.xlsx")
    result = build_output(SAMPLE_RATES, out_path)
    assert result == out_path
    assert os.path.isfile(result)


def test_all_rates_written(tmp_path):
    """Every item in SAMPLE_RATES must be written to the expected cell."""
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"))
    wb = openpyxl.load_workbook(out)
    ws = wb["적용근거"]
    for item, rate in SAMPLE_RATES.items():
        cell_addr = CELL_MAP[item]
        assert ws[cell_addr].value == rate, (
            f"Item '{item}' expected rate {rate} at {cell_addr}, "
            f"got {ws[cell_addr].value}"
        )


def test_cell_map_matches_template_layout():
    """Contract-pinning test: CELL_MAP rows must match column-A label rows in the template.

    If build_template.py ITEMS order ever changes, this test catches the mismatch
    before rates get silently written to the wrong cells.
    """
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["적용근거"]

    # Build a lookup: label → row number from actual template column A
    template_label_row: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val in ITEMS:
            template_label_row[val] = r

    for item in ITEMS:
        assert item in template_label_row, (
            f"Item '{item}' not found in template column A"
        )
        actual_template_row = template_label_row[item]
        builder_cell = CELL_MAP[item]          # e.g. "I3"
        builder_row = int(builder_cell[1:])    # strip "I" prefix → int

        assert builder_row == actual_template_row, (
            f"CELL_MAP mismatch for '{item}': builder targets row {builder_row} "
            f"(cell {builder_cell}) but template has the label at row {actual_template_row}. "
            "Update ITEMS order in build_template.py or src/builder.py to fix."
        )
