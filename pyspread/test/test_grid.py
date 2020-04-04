# -*- coding: utf-8 -*-

# Copyright Martin Manns
# Distributed under the terms of the GNU General Public License

# --------------------------------------------------------------------
# pyspread is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# pyspread is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with pyspread.  If not, see <http://www.gnu.org/licenses/>.
# --------------------------------------------------------------------


"""
test_grid
=========

Unit tests for grid.py

"""

from contextlib import contextmanager
from os.path import abspath, dirname, join
import sys

import pytest

from PyQt5.QtWidgets import QApplication

PYSPREADPATH = abspath(join(dirname(__file__) + "/.."))
LIBPATH = abspath(PYSPREADPATH + "/lib")


@contextmanager
def insert_path(path):
    sys.path.insert(0, path)
    yield
    sys.path.pop(0)


with insert_path(PYSPREADPATH):
    from ..pyspread import MainWindow

app = QApplication([])
main_window = MainWindow()


class TestGrid:
    """Unit tests for Grid in grid.py"""

    grid = main_window.grid

    param_test_row = [(0, 0), (1, 1), (100, 100), (1000, 0), (10000, 0),
                      (-1, 0)]

    @pytest.mark.parametrize("row, res", param_test_row)
    def test_row(self, row, res):
        old_row = self.grid.row
        self.grid.row = row
        assert self.grid.row == res
        self.grid.row = old_row

    param_test_column = [(0, 0), (1, 1), (100, 0), (1000, 0), (10000, 0),
                         (-1, 0)]

    @pytest.mark.parametrize("column, res", param_test_column)
    def test_column(self, column, res):
        old_column = self.grid.column
        self.grid.column = column
        assert self.grid.column == res
        self.grid.column = old_column

    param_test_table = [(0, 0), (1, 1), (3, 0), (-1, 0)]

    @pytest.mark.parametrize("table, res", param_test_table)
    def test_table(self, table, res):
        old_table = self.grid.table
        self.grid.table = table
        assert self.grid.table == res
        self.grid.table = old_table


class TestGridHeaderView:
    """Unit tests for GridHeaderView in grid.py"""


class TestGridTableModel:
    """Unit tests for GridTableModel in grid.py"""

    model = main_window.grid.model

    param_test_shape = [
        ((1, 1, 1), (1, 1, 1), None),
        ((0, 0, 0), (1000, 100, 3), ValueError),
        ((9999999999, 0, 0), (1000, 100, 3), ValueError),
        ((1000000, 10000, 10), (1000000, 10000, 10), None),
        ((1000, 100, 3), (1000, 100, 3), None),
    ]

    @pytest.mark.parametrize("shape, res, error", param_test_shape)
    def test_shape(self, shape, res, error):
        """Unit test for File -> New"""

        try:
            self.model.shape = shape
        except ValueError:
            assert error == ValueError
        else:
            assert error is None
        assert self.model.shape == res

        self.model.shape = (1000, 100, 3)


class TestGridCellNavigator:
    """Unit tests for GridCellNavigator in grid.py"""


class TestGridCellDelegate:
    """Unit tests for GridCellDelegate in grid.py"""


class TestTableChoice:
    """Unit tests for TableChoice in grid.py"""
