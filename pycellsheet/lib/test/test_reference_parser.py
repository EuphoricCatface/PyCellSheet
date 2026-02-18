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

from ..pycellsheet import PythonCode, ReferenceParser


class _DummyDictGrid:
    sheet_names = ["Sheet 0", "Other"]


class _DummyCodeArray:
    def __init__(self):
        self.dict_grid = _DummyDictGrid()
        self.sheet_globals_copyable = [{}, {}]
        self.sheet_globals_uncopyable = [{}, {}]

    def __getitem__(self, key):
        return key


def test_parser_wraps_remaining_single_cell_after_sheet_ref():
    parser = ReferenceParser(None)

    assert parser.parser(PythonCode('"0"!A1 + B2')) == 'Sh("0").C("A1") + C("B2")'


def test_parser_wraps_single_cells_around_sheet_ref():
    parser = ReferenceParser(None)

    assert parser.parser(PythonCode('A1 + "0"!B2 + C3')) == 'C("A1") + Sh("0").C("B2") + C("C3")'


def test_cr_accepts_quoted_sheet_name():
    code_array = _DummyCodeArray()
    parser = ReferenceParser(code_array)
    current_sheet = parser.Sheet("0", code_array)

    assert parser.CR('"Other"!A1', current_sheet) == (0, 0, 1)
