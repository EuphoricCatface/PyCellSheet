# -*- coding: utf-8 -*-

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
