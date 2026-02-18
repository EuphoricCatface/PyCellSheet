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

from interfaces.pycs import PycsReader, PycsWriter

sys.path.pop(0)


class _DummyDictGrid:
    def __init__(self):
        self.sheet_names = []
        self.macros = []
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
        self.macros = ["" for _ in range(tables)]
        self.dict_grid.macros = self.macros
        self.row_heights = {}
        self.col_widths = {}
        self.cell_attributes = []

    @property
    def sheet_scripts(self):
        return self.macros

    @sheet_scripts.setter
    def sheet_scripts(self, value):
        self.macros = value


class _DummyWriterCodeArray:
    def __init__(self, sheet_names, macros):
        self.shape = (1, 1, len(macros))
        self.dict_grid = _DummyDictGrid()
        self.dict_grid.sheet_names = sheet_names
        self.dict_grid.macros = macros
        self.cell_attributes = []
        self.dict_grid.row_heights = {}
        self.dict_grid.col_widths = {}
        self._code = {}

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


def test_pycs2macros_parses_named_and_numeric_headers():
    code_array = _DummyCodeArray(2)
    code_array.dict_grid.sheet_names = ["Alpha", "Beta"]
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2macros("(macro:'Alpha') 2\n")
    reader._pycs2macros("a = 1\n")
    reader._pycs2macros("b = 2\n")
    reader._pycs2macros("(macro:1) 1\n")
    reader._pycs2macros("x = 3\n")

    assert code_array.macros[0] == "a = 1\nb = 2"
    assert code_array.macros[1] == "x = 3"


def test_pycs2macros_unknown_sheet_name_falls_back_sequential_index():
    code_array = _DummyCodeArray(2)
    code_array.dict_grid.sheet_names = ["Alpha", "Beta"]
    reader = PycsReader(BytesIO(b""), code_array)

    reader._pycs2macros("(macro:'Unknown') 1\n")
    reader._pycs2macros("u = 1\n")

    assert code_array.macros[0] == "u = 1"


def test_pycs2macros_raises_for_invalid_header():
    code_array = _DummyCodeArray(1)
    reader = PycsReader(BytesIO(b""), code_array)

    with pytest.raises(ValueError):
        reader._pycs2macros("macro:bad header\n")


def test_writer_macros_section_uses_normalized_sheet_names():
    code_array = _DummyWriterCodeArray(
        ["Main", " ", "Main"],
        ["a = 1", "b = 2", "c = 3"],
    )
    writer = PycsWriter(code_array)

    macros_lines = list(writer._macros2pycs())

    assert macros_lines[0].startswith("(macro:'Main') 1\n")
    assert macros_lines[1].startswith("(macro:'Sheet 1') 1\n")
    assert macros_lines[2].startswith("(macro:'Main_1') 1\n")


def test_writer_reader_round_trip_preserves_sheet_names_and_macros():
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
    assert target.macros == ["a = 1", "b = 2", "c = 3"]


def test_reader_macros_are_visible_via_sheet_scripts_alias():
    source = _DummyWriterCodeArray(
        ["Main"],
        ["x = 9"],
    )

    serialized = "".join(list(PycsWriter(source))).encode("utf-8")
    target = _DummyCodeArray(1)
    list(PycsReader(BytesIO(serialized), target))

    assert target.macros == ["x = 9"]
    assert target.sheet_scripts == ["x = 9"]
