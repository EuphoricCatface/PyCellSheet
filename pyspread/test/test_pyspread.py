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

from PyQt6.QtWidgets import QApplication, QAbstractItemView


PYSPREADPATH = abspath(join(dirname(__file__) + "/.."))
LIBPATH = abspath(PYSPREADPATH + "/lib")


@contextmanager
def insert_path(path):
    sys.path.insert(0, path)
    yield
    sys.path.pop(0)


@contextmanager
def multi_selection_mode(grid):
    grid.clearSelection()
    old_selection_mode = grid.selectionMode()
    grid.setSelectionMode(QAbstractItemView.MultiSelection)
    yield
    grid.setSelectionMode(old_selection_mode)


with insert_path(PYSPREADPATH):
    from ..pyspread import MainWindow
    from ..commands import MakeButtonCell, RemoveButtonCell
    from ..lib.selection import Selection


app = QApplication.instance()
if app is None:
    app = QApplication([])
main_window = MainWindow()
zoom_levels = main_window.settings.zoom_levels


class TestMainWindow:
    """Unit tests for MainWindow in pyspread.py"""

    grid = main_window.grid
    cell_attributes = grid.model.code_array.cell_attributes

    def test_safe_mode(self):
        """Unit test for safe_mode"""

        main_window.safe_mode = True
        main_window.main_window_actions.approve.isEnabled()
        assert main_window.safe_mode

        main_window.safe_mode = False
        assert not main_window.main_window_actions.approve.isEnabled()
        assert not main_window.safe_mode

    def test_on_clear_globals(self):
        """Unit test for on_clear_globals"""

        main_window.on_clear_globals()

    def test_recalculate_actions_are_undo_safe(self):
        """Recalculate actions must not push undo commands."""

        code_array = main_window.grid.model.code_array
        old_mode = main_window.settings.recalc_mode
        main_window.settings.recalc_mode = "manual"

        try:
            code_array[0, 0, 0] = "1"
            code_array[0, 1, 0] = "C('A1') + 1"
            _ = code_array[0, 1, 0]
            code_array[0, 0, 0] = "2"
            main_window.grid.current = (0, 1, 0)

            before_count = main_window.undo_stack.count()
            before_index = main_window.undo_stack.index()

            main_window.on_recalculate()
            main_window.on_recalculate_cell_only()
            main_window.on_recalculate_ancestors()
            main_window.on_recalculate_children()
            main_window.on_recalculate_all()

            assert main_window.undo_stack.count() == before_count
            assert main_window.undo_stack.index() == before_index
        finally:
            for key in ((0, 0, 0), (0, 1, 0)):
                try:
                    code_array.pop(key)
                except KeyError:
                    pass
            code_array.dep_graph.remove_cell((0, 0, 0))
            code_array.dep_graph.remove_cell((0, 1, 0))
            code_array.dep_graph.dirty.clear()
            main_window.settings.recalc_mode = old_mode
