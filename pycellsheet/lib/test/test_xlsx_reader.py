# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License

# --------------------------------------------------------------------
# pycellsheet is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# pycellsheet is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with pycellsheet.  If not, see <http://www.gnu.org/licenses/>.
# --------------------------------------------------------------------

"""
test_xlsx_reader
================

Unit tests for xlsx import reader behavior.
"""

from io import BytesIO

import pytest

openpyxl = pytest.importorskip("openpyxl")

try:
    from pycellsheet.interfaces.xlsx import XlsxReader
except ImportError:
    from interfaces.xlsx import XlsxReader


class _DummyCodeArray:
    """Minimal target interface consumed by XlsxReader."""

    def __init__(self):
        self.shape = (1, 1, 1)
        self.sheet_scripts = [""]
        self.row_heights = {}
        self.col_widths = {}
        self.dict_grid = {}
        self.cell_attributes = []


def _workbook_to_stream(wb) -> BytesIO:
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def test_xlsx_reader_sets_shape_codes_and_sheet_scripts():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws["A1"] = 42
    ws["B1"] = "text"
    ws["C1"] = True
    ws["D1"] = "=A1+1"
    ws.row_dimensions[1].height = 24
    ws.column_dimensions["A"].width = 20
    ws.merge_cells("A2:B2")
    ws["A2"] = "merged"

    ws2 = wb.create_sheet("Second")
    ws2["B3"] = "sheet2"

    code_array = _DummyCodeArray()
    reader = XlsxReader(_workbook_to_stream(wb), code_array)

    keys = list(reader)

    assert code_array.shape == (3, 4, 2)
    assert "_sheetnames = ['Main', 'Second']" in code_array.sheet_scripts[0]
    assert code_array.dict_grid[(0, 0, 0)] == "42"
    assert code_array.dict_grid[(0, 1, 0)] == "'text'"
    assert code_array.dict_grid[(0, 2, 0)] == "True"
    formula_code = code_array.dict_grid[(0, 3, 0)]
    assert formula_code in {"'=A1+1'", '_C_("A1") + 1'}
    assert code_array.dict_grid[(2, 1, 1)] == "'sheet2'"
    assert (0, 0, 0) in keys
    assert (2, 1, 1) in keys
    assert code_array.row_heights[(0, 0)] > 0
    assert code_array.col_widths[(0, 0)] > 0
    assert code_array.cell_attributes  # includes merge/style attrs


def test_xlsx_rgba2rgb_conversion_and_transparent_guard():
    assert XlsxReader.xlsx_rgba2rgb("FF112233") == (17, 34, 51)

    with pytest.raises(ValueError):
        XlsxReader.xlsx_rgba2rgb("00112233")


def test_xlsx2code_handles_error_and_unknown_cell_types():
    code_array = _DummyCodeArray()
    wb = openpyxl.Workbook()
    reader = XlsxReader(_workbook_to_stream(wb), code_array)

    class _Cell:
        def __init__(self, data_type, value):
            self.data_type = data_type
            self.value = value

    reader._xlsx2code((0, 0, 0), _Cell("e", "#DIV/0!"))
    assert isinstance(code_array.dict_grid[(0, 0, 0)], Exception)

    reader._xlsx2code((0, 1, 0), _Cell("z", "mystery"))
    assert code_array.dict_grid[(0, 1, 0)] == "'mystery'"


def test_xlsx_reader_appends_cell_attributes_with_expected_table():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws["A1"].font = openpyxl.styles.Font(bold=True, italic=True, underline="single")
    ws["A1"].alignment = openpyxl.styles.Alignment(horizontal="right", vertical="top")
    ws["A1"].fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="FF00FF00")

    code_array = _DummyCodeArray()
    reader = XlsxReader(_workbook_to_stream(wb), code_array)
    list(reader)

    assert code_array.cell_attributes
    assert any(attr[1] == 0 for attr in code_array.cell_attributes)
