# -*- coding: utf-8 -*-

from io import BytesIO
from os.path import abspath, dirname, join
import sys

pyspread_path = abspath(join(dirname(__file__) + "/../.."))
sys.path.insert(0, pyspread_path)

from interfaces.pycs import PycsReader

sys.path.pop(0)


class _DummyDictGrid:
    def __init__(self):
        self.sheet_names = []


class _DummyCodeArray:
    def __init__(self, tables: int):
        self.shape = (1, 1, tables)
        self.dict_grid = _DummyDictGrid()
        self.cell_attributes = []


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
