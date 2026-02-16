# -*- coding: utf-8 -*-

# Copyright Martin Manns
# Modified by Seongyong Park (EuphCat)
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
test_workflows
==============

Unit tests for workflows.py

"""

from contextlib import contextmanager
import os
from os.path import abspath, dirname, join
from pathlib import Path
import sys

import pytest

from PyQt6.QtCore import Qt, QItemSelectionModel
from PyQt6.QtWidgets import QApplication

try:
    from pyspread.dialogs import GridShapeDialog
except ImportError:
    from dialogs import GridShapeDialog


PYSPREADPATH = abspath(join(dirname(__file__) + "/.."))
LIBPATH = abspath(PYSPREADPATH + "/lib")


@contextmanager
def insert_path(path):
    sys.path.insert(0, path)
    yield
    sys.path.pop(0)


with insert_path(PYSPREADPATH):
    from ..pyspread import MainWindow


app = QApplication.instance()
if app is None:
    app = QApplication([])
main_window = MainWindow()


class TestWorkflows:
    """Unit tests for Workflows in workflows.py"""

    workflows = main_window.workflows

    def test_busy_cursor(self):
        """Unit test for busy_cursor"""

        assert QApplication.overrideCursor() != Qt.CursorShape.WaitCursor

        with self.workflows.busy_cursor():
            assert QApplication.overrideCursor() == Qt.CursorShape.WaitCursor

        assert QApplication.overrideCursor() != Qt.CursorShape.WaitCursor

    def test_prevent_updates(self):
        """Unit test for prevent_updates"""

        assert not main_window.prevent_updates

        with self.workflows.prevent_updates():
            assert main_window.prevent_updates

        assert not main_window.prevent_updates

    def test_reset_changed_since_save(self):
        """Unit test for reset_changed_since_save"""

        main_window.settings.changed_since_save = True
        self.workflows.reset_changed_since_save()
        assert not main_window.settings.changed_since_save

    param_update_main_window_title = [
        (Path.home(), "pyspread"),
        (Path("/test.pys"), "test.pys - pyspread"),
    ]

    @pytest.mark.parametrize("path, title", param_update_main_window_title)
    def test_update_main_window_title(self, path, title):
        """Unit test for update_main_window_title"""

        main_window.settings.last_file_input_path = path
        self.workflows.update_main_window_title()
        assert main_window.windowTitle() == title

    param_file_new = [
        ((1000, 100, 3), (1000, 100, 3), None, True),
        ((100, 100, 3), (100, 100, 3), None, True),
        ((10000000, 100, 3), (1000, 100, 3),
         "Error: Grid shape (10000000, 100, 3) exceeds (1000000, 100000, 100).",
         False),
        (None, (1000, 100, 3), None, False),
    ]

    @pytest.mark.parametrize("shape, res, msg, should_apply", param_file_new)
    def test_file_new(self, shape, res, msg, should_apply, monkeypatch):
        """Unit test for file_new"""

        called = {"count": 0}

        def fake_apply_all_sheet_scripts():
            called["count"] += 1
            return 0, 0

        monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts",
                            fake_apply_all_sheet_scripts)
        monkeypatch.setattr(GridShapeDialog, "shape", shape)
        self.workflows.file_new()

        assert main_window.grid.model.shape == res
        assert main_window.grid.current == (0, 0, 0)
        assert main_window.settings.last_file_input_path == Path.home()
        assert main_window.settings.changed_since_save is False
        assert main_window.safe_mode is False
        expected_calls = 1 if should_apply else 0
        assert called["count"] == expected_calls
        if msg:
            assert main_window.statusBar().currentMessage() == msg

        monkeypatch.setattr(GridShapeDialog, "shape",
                            main_window.settings.shape)
        self.workflows.file_new()

    def test_apply_all_sheet_scripts_executes_each_table(self, monkeypatch):
        """apply_all_sheet_scripts should run one script per table."""

        code_array = main_window.grid.model.code_array
        old_shape = code_array.shape
        old_safe_mode = main_window.safe_mode

        try:
            main_window.safe_mode = False
            main_window.grid.model.shape = (old_shape[0], old_shape[1], 3)
            called_tables = []

            def fake_execute_sheet_script(table):
                called_tables.append(table)
                return "", ""

            monkeypatch.setattr(code_array, "execute_sheet_script",
                                fake_execute_sheet_script)

            executed, errors = self.workflows.apply_all_sheet_scripts()

            assert called_tables == [0, 1, 2]
            assert executed == 3
            assert errors == 0
        finally:
            main_window.grid.model.shape = old_shape
            main_window.safe_mode = old_safe_mode

    def test_apply_all_sheet_scripts_skips_in_safe_mode(self, monkeypatch):
        """apply_all_sheet_scripts should not execute anything in safe mode."""

        code_array = main_window.grid.model.code_array
        old_safe_mode = main_window.safe_mode

        try:
            main_window.safe_mode = True
            called = {"count": 0}

            def fake_execute_sheet_script(table):
                called["count"] += 1
                return "", ""

            monkeypatch.setattr(code_array, "execute_sheet_script",
                                fake_execute_sheet_script)

            executed, errors = self.workflows.apply_all_sheet_scripts()

            assert called["count"] == 0
            assert executed == 0
            assert errors == 0
        finally:
            main_window.safe_mode = old_safe_mode

    def test_apply_all_sheet_scripts_counts_errors_and_updates_current_view(self, monkeypatch):
        """apply_all_sheet_scripts should count errors and update current-table output."""

        code_array = main_window.grid.model.code_array
        old_shape = code_array.shape
        old_safe_mode = main_window.safe_mode
        old_current_table = main_window.sheet_script_panel.current_table

        try:
            main_window.safe_mode = False
            main_window.grid.model.shape = (old_shape[0], old_shape[1], 3)
            main_window.sheet_script_panel.current_table = 1

            calls = {"updated": [], "gui": 0, "data_changed": 0}

            def fake_execute_sheet_script(table):
                if table == 1:
                    return "ok-table-1", ""
                if table == 2:
                    return "", "boom"
                return "ok-table-0", ""

            def fake_update_result_viewer(result, err):
                calls["updated"].append((result, err))

            def fake_gui_update():
                calls["gui"] += 1

            def fake_emit_data_changed_all():
                calls["data_changed"] += 1

            monkeypatch.setattr(code_array, "execute_sheet_script",
                                fake_execute_sheet_script)
            monkeypatch.setattr(main_window.sheet_script_panel, "update_result_viewer",
                                fake_update_result_viewer)
            monkeypatch.setattr(main_window.grid, "gui_update", fake_gui_update)
            monkeypatch.setattr(main_window.grid.model, "emit_data_changed_all",
                                fake_emit_data_changed_all)

            executed, errors = self.workflows.apply_all_sheet_scripts()

            assert executed == 3
            assert errors == 1
            assert calls["updated"] == [("ok-table-1", "")]
            assert calls["gui"] == 1
            assert calls["data_changed"] == 1
        finally:
            main_window.grid.model.shape = old_shape
            main_window.safe_mode = old_safe_mode
            main_window.sheet_script_panel.current_table = old_current_table

    def test_filepath_open_untrusted_defers_scripts_until_approve(self, tmp_path, monkeypatch):
        """Untrusted loads must stay in safe mode and defer script execution."""

        payload = (
            "[PyCellSheet save file version]\n"
            "0.0\n"
            "[shape]\n"
            "1\t1\t1\n"
            "[sheet_names]\n"
            "Sheet 0\n"
            "[macros]\n"
            "(macro:'Sheet 0') 1\n"
            "VALUE = 7\n"
            "[grid]\n"
            "0\t0\t0\t'abc'\n"
            "[attributes]\n"
            "[row_heights]\n"
            "[col_widths]\n"
        )
        filepath = tmp_path / "untrusted.pycsu"
        filepath.write_text(payload, encoding="utf-8")

        code_array = main_window.grid.model.code_array
        apply_calls = {"count": 0}
        exec_calls = {"count": 0}
        old_safe_mode = main_window.safe_mode
        old_cwd = os.getcwd()

        def fake_file_progress_gen(_main_window, iterable, _title, _label, _lines):
            for i, line in enumerate(iterable, start=1):
                yield i, line

        def fake_apply_all_sheet_scripts():
            apply_calls["count"] += 1
            return 1, 0

        def fake_execute_sheet_script(table):
            exec_calls["count"] += 1
            return "", ""

        class _ApproveDialog:
            def __init__(self, _parent):
                self.choice = True

        try:
            monkeypatch.setitem(
                self.workflows.filepath_open.__globals__,
                "file_progress_gen",
                fake_file_progress_gen,
            )
            monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts", fake_apply_all_sheet_scripts)
            monkeypatch.setattr(code_array, "execute_sheet_script", fake_execute_sheet_script)
            monkeypatch.setitem(
                main_window.on_approve.__globals__,
                "ApproveWarningDialog",
                _ApproveDialog,
            )

            self.workflows.filepath_open(filepath)

            assert main_window.safe_mode is True
            assert exec_calls["count"] == 0
            assert apply_calls["count"] == 0

            main_window.on_approve()

            assert main_window.safe_mode is False
            assert apply_calls["count"] == 1
        finally:
            os.chdir(old_cwd)
            main_window.safe_mode = old_safe_mode

    param_count_file_lines = [
        ("", 0, "counttest.txt", None),
        ("\n"*100, 100, "counttest.txt", None),
        ("Test"*100, 0, "counttest.txt", None),
        ("Test\n"*10, 10, "counttest.txt", None),
        ("Test\n"*10, None, "false_filename.txt", "Error"),
    ]

    @pytest.mark.parametrize("txt, res, filename, msg", param_count_file_lines)
    def test_count_file_lines(self, txt, res, filename, msg, tmpdir):
        """Unit test for count_file_lines"""

        tmpfile = tmpdir / "counttest.txt"
        tmpfile.write_text(txt, "utf-8")
        testfile = tmpdir / filename
        assert self.workflows.count_file_lines(testfile) == res
        if msg:
            assert str(testfile) in main_window.statusBar().currentMessage()
        tmpfile.remove()

    def test_edit_sort_ascending(self):
        """Unit test for test_edit_sort_ascending"""

        main_window.grid.model.shape = (1000, 100, 3)
        main_window.grid.model.reset()

        main_window.grid.model.code_array[0, 0, 0] = "1"
        main_window.grid.model.code_array[1, 0, 0] = "3"
        main_window.grid.model.code_array[2, 0, 0] = "2"
        main_window.grid.model.code_array[0, 1, 0] = "12"
        main_window.grid.model.code_array[1, 1, 0] = "33"
        main_window.grid.model.code_array[2, 1, 0] = "24"

        for row in range(3):
            for column in range(2):
                index = main_window.grid.model.index(row, column)
                main_window.grid.selectionModel().select(
                    index, QItemSelectionModel.SelectionFlag.Select)

        self.workflows.edit_sort_ascending()
        assert main_window.grid.model.code_array((0, 0, 0)) == "1"
        assert main_window.grid.model.code_array((1, 0, 0)) == "2"
        assert main_window.grid.model.code_array((2, 1, 0)) == "33"

    def test_edit_sort_descending(self):
        """Unit test for test_edit_sort_descending"""

        main_window.grid.model.shape = (1000, 100, 3)
        main_window.grid.model.reset()

        main_window.grid.model.code_array[0, 0, 0] = "1"
        main_window.grid.model.code_array[1, 0, 0] = "3"
        main_window.grid.model.code_array[2, 0, 0] = "2"
        main_window.grid.model.code_array[0, 1, 0] = "12"
        main_window.grid.model.code_array[1, 1, 0] = "33"
        main_window.grid.model.code_array[2, 1, 0] = "24"

        for row in range(3):
            for column in range(2):
                index = main_window.grid.model.index(row, column)
                main_window.grid.selectionModel().select(
                    index, QItemSelectionModel.SelectionFlag.Select)

        self.workflows.edit_sort_descending()
        assert main_window.grid.model.code_array((0, 0, 0)) == "3"
        assert main_window.grid.model.code_array((1, 0, 0)) == "2"
        assert main_window.grid.model.code_array((2, 1, 0)) == "12"
