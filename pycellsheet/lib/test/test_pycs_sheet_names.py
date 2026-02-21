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

from io import BytesIO
from os.path import abspath, dirname, join
import sys

import pytest

project_path = abspath(join(dirname(__file__) + "/../.."))
sys.path.insert(0, project_path)

from interfaces.pycs import PycsReader, PycsWriter, wxcolor2rgb, qt52qt6_fontweights, qt62qt5_fontweights

sys.path.pop(0)


class _DummyDictGrid:
    def __init__(self):
        self.sheet_names = []
        self.sheet_scripts = []
        self.row_heights = {}
        self.col_widths = {}
        self._grid = {}

    def __setitem__(self, key, value):
        self._grid[key] = value

    def __iter__(self):
        return iter(self._grid)

    def __len__(self):
        return len(self._grid)


class _DummyCodeArray:
    def __init__(self, tables: int):
        self.shape = (1, 1, tables)
        self.dict_grid = _DummyDictGrid()
        self.sheet_scripts = ["" for _ in range(tables)]
        self.dict_grid.sheet_scripts = self.sheet_scripts
        self.row_heights = {}
        self.col_widths = {}
        self.cell_attributes = []
        self.exp_parser_code = "return cell"


class _DummyWriterCodeArray:
    def __init__(self, sheet_names, sheet_scripts):
        self.shape = (1, 1, len(sheet_scripts))
        self.dict_grid = _DummyDictGrid()
        self.dict_grid.sheet_names = sheet_names
        self.dict_grid.sheet_scripts = sheet_scripts
        self.cell_attributes = []
        self.dict_grid.row_heights = {}
        self.dict_grid.col_widths = {}
        self._code = {}
        self.exp_parser_code = "return cell"

    def __iter__(self):
        return iter(self._code)

    def __call__(self, key):
        return self._code[key]


def test_pycs2sheet_names_sanitizes_and_uniquifies():
    code_array = _DummyCodeArray(3)
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2sheet_names("Sheet 0\n")
    reader._pycs2sheet_names("Sheet 0\n")
    reader._pycs2sheet_names("Bad\tName\n")

    assert code_array.dict_grid.sheet_names == ["Sheet 0", "Sheet 0_1", "BadName"]


def test_finalize_sheet_names_fills_missing_defaults():
    code_array = _DummyCodeArray(3)
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2sheet_names("Only One\n")
    reader._finalize_sheet_names()

    assert code_array.dict_grid.sheet_names == ["Only One", "Sheet 1", "Sheet 2"]


def test_writer_normalizes_sheet_names_for_output():
    code_array = _DummyWriterCodeArray(
        ["Good", " ", "Bad\tName", "Good"],
        ["a=1", "b=2", "c=3", "d=4"],
    )
    writer = PycsWriter(code_array)

    assert writer._normalized_sheet_names() == ["Good", "Sheet 1", "BadName", "Good_1"]


def test_reader_iter_finalizes_sheet_names():
    code_array = _DummyCodeArray(1)
    payload = (
        b"[PyCellSheet save file version]\n"
        b"0.0\n"
        b"[shape]\n"
        b"1\t1\t2\n"
        b"[sheet_names]\n"
        b"Only One\n"
    )
    reader = PycsReader(BytesIO(payload), code_array)

    list(reader)

    assert code_array.shape == (1, 1, 2)
    assert code_array.dict_grid.sheet_names == ["Only One", "Sheet 1"]


def test_pycs2sheet_names_ignores_extra_names_after_table_count():
    code_array = _DummyCodeArray(2)
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2sheet_names("A\n")
    reader._pycs2sheet_names("B\n")
    reader._pycs2sheet_names("C\n")

    assert code_array.dict_grid.sheet_names == ["A", "B"]


def test_finalize_sheet_names_truncates_extra_entries():
    code_array = _DummyCodeArray(2)
    code_array.dict_grid.sheet_names = ["A", "B", "C"]
    reader = PycsReader(BytesIO(b""), code_array)

    reader._finalize_sheet_names()

    assert code_array.dict_grid.sheet_names == ["A", "B"]


def test_row_heights_and_col_widths_ignore_out_of_bounds():
    code_array = _DummyCodeArray(1)
    code_array.shape = (2, 2, 1)
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2row_heights("1\t0\t9.5\n")
    reader._pycs2row_heights("2\t0\t3.0\n")  # row out of bounds
    reader._pycs2col_widths("1\t0\t8.0\n")
    reader._pycs2col_widths("2\t0\t4.0\n")  # col out of bounds

    assert code_array.row_heights == {(1, 0): 9.5}
    assert code_array.col_widths == {(1, 0): 8.0}


def test_pycs2sheet_scripts_parses_named_headers():
    code_array = _DummyCodeArray(2)
    code_array.dict_grid.sheet_names = ["Alpha", "Beta"]
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2sheet_scripts("(sheet_script:'Alpha') 2\n")
    reader._pycs2sheet_scripts("a = 1\n")
    reader._pycs2sheet_scripts("b = 2\n")
    reader._pycs2sheet_scripts("(sheet_script:'Beta') 1\n")
    reader._pycs2sheet_scripts("x = 3\n")

    assert code_array.sheet_scripts[0] == "a = 1\nb = 2"
    assert code_array.sheet_scripts[1] == "x = 3"


def test_pycs2sheet_scripts_unknown_sheet_name_raises():
    code_array = _DummyCodeArray(2)
    code_array.dict_grid.sheet_names = ["Alpha", "Beta"]
    reader = PycsReader(BytesIO(b""), code_array)

    with pytest.raises(ValueError, match="Unknown sheet name"):
        reader._pycs2sheet_scripts("(sheet_script:'Unknown') 1\n")


def test_pycs2sheet_scripts_raises_for_invalid_header():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)

    with pytest.raises(ValueError):
        reader._pycs2sheet_scripts("macro:bad header\n")


def test_writer_sheet_scripts_section_uses_normalized_sheet_names():
    code_array = _DummyWriterCodeArray(
        ["Main", " ", "Main"],
        ["a = 1", "b = 2", "c = 3"],
    )
    writer = PycsWriter(code_array)

    sheet_scripts_lines = list(writer._sheet_scripts2pycs())

    assert sheet_scripts_lines[0].startswith("(sheet_script:'Main') 1\n")
    assert sheet_scripts_lines[1].startswith("(sheet_script:'Sheet 1') 1\n")
    assert sheet_scripts_lines[2].startswith("(sheet_script:'Main_1') 1\n")


def test_writer_reader_round_trip_preserves_sheet_names_and_sheet_scripts():
    source = _DummyWriterCodeArray(
        ["Revenue", "Revenue", " "],
        ["a = 1", "b = 2", "c = 3"],
    )
    source._code[(0, 0, 0)] = "'text'"

    serialized = "".join(list(PycsWriter(source))).encode("utf-8")

    target = _DummyCodeArray(3)
    reader = PycsReader(BytesIO(serialized), target)
    list(reader)

    assert target.dict_grid.sheet_names == ["Revenue", "Revenue_1", "Sheet 2"]
    assert target.sheet_scripts == ["a = 1", "b = 2", "c = 3"]


def test_reader_sheet_scripts_round_trip():
    source = _DummyWriterCodeArray(
        ["Main"],
        ["x = 9"],
    )

    serialized = "".join(list(PycsWriter(source))).encode("utf-8")
    target = _DummyCodeArray(1)
    list(PycsReader(BytesIO(serialized), target))

    assert target.sheet_scripts == ["x = 9"]
    assert target.sheet_scripts == ["x = 9"]


def test_writer_reader_round_trip_preserves_parser_settings_section():
    source = _DummyWriterCodeArray(["Main"], ["x = 9"])
    source.exp_parser_code = "if cell.startswith('>'):\n    return cell[1:]"

    serialized = "".join(list(PycsWriter(source))).encode("utf-8")
    target = _DummyCodeArray(1)
    list(PycsReader(BytesIO(serialized), target))

    assert target.exp_parser_code == source.exp_parser_code


def test_color_and_weight_conversion_helpers():
    assert wxcolor2rgb(0x112233) == (0x11, 0x22, 0x33)
    assert qt52qt6_fontweights(50) == 405
    assert qt62qt5_fontweights(405) == 50


def test_pycs_version_rejects_future_versions():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)
    with pytest.raises(ValueError):
        reader._pycs_version("1.0\n")


def test_pycs2shape_rejects_non_positive_dimension():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)
    with pytest.raises(ValueError):
        reader._pycs2shape("0\t1\t1\n")


def test_pycs2code_ignores_out_of_bounds_keys():
    code_array = _DummyCodeArray(1)
    code_array.shape = (1, 1, 1)
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2code("0\t0\t0\tok\n")
    reader._pycs2code("1\t0\t0\tnope\n")

    assert code_array.dict_grid._grid == {(0, 0, 0): "ok"}


def test_attr_convert_1to2_handles_known_transforms():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)

    assert reader._attr_convert_1to2("bordercolor_bottom", 0x010203) == ("bordercolor_bottom", (1, 2, 3))
    assert reader._attr_convert_1to2("fontweight", 90) == ("fontweight", 50)
    assert reader._attr_convert_1to2("fontstyle", 93) == ("fontstyle", 1)
    assert reader._attr_convert_1to2("markup", True) == ("renderer", "markup")
    assert reader._attr_convert_1to2("angle", -10) == ("angle", 350)
    assert reader._attr_convert_1to2("merge_area", None) == (None, None)
    assert reader._attr_convert_1to2("justification", "center") == ("justification", "justify_center")
    assert reader._attr_convert_1to2("vertical_align", "middle") == ("vertical_align", "align_center")
    assert reader._attr_convert_1to2("other", "x") == ("other", "x")


def test_pycs2attributes_parses_cell_attribute_rows():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)
    line = "[]\t[]\t[]\t[]\t[(0, 0)]\t0\t'angle'\t90\n"

    reader._pycs2attributes(line)

    assert len(code_array.cell_attributes) == 1
    selection, table, attr = code_array.cell_attributes[0]
    assert table == 0
    assert (0, 0) in selection.cells
    assert attr["angle"] == 90


def test_pycs2sheet_scripts_nonsequential_numeric_header_rejected():
    code_array = _DummyCodeArray(2)
    reader = PycsReader(BytesIO(b""), code_array)

    with pytest.raises(ValueError, match="Numeric sheet_script headers"):
        reader._pycs2sheet_scripts("(sheet_script:2) 1\n")


def test_pycs2parser_settings_rejects_unknown_key():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)

    with pytest.raises(ValueError, match="Unknown parser_settings key"):
        reader._pycs2parser_settings("pycel_formula_opt_in\tTrue\n")
