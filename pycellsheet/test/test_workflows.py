# -*- coding: utf-8 -*-

# Copyright Martin Manns
# Modified by Seongyong Park (EuphCat)
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
    from pycellsheet.dialogs import GridShapeDialog
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
    from ..pycellsheet import MainWindow
    from ..model.model import INITSCRIPT_DEFAULT


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
        (Path.home(), "PyCellSheet"),
        (Path("/test.pys"), "test.pys - PyCellSheet"),
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
        if should_apply:
            assert main_window.grid.model.code_array.sheet_scripts == [
                INITSCRIPT_DEFAULT for _ in range(res[2])
            ]
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

    def test_resolve_sheet_script_drafts_cancel_blocks_transition(self, monkeypatch):
        """Cancelling draft resolution should block transition workflow actions."""

        code_array = main_window.grid.model.code_array
        panel = main_window.sheet_script_panel
        old_shape = code_array.shape
        old_drafts = list(code_array.sheet_scripts_draft)
        old_safe_mode = main_window.safe_mode
        old_changed = main_window.settings.changed_since_save

        class _DraftDialog:
            def __init__(self, _parent):
                self.choice = None

        called = {"count": 0}

        def fake_apply_all_sheet_scripts():
            called["count"] += 1
            return 0, 0

        try:
            main_window.safe_mode = False
            main_window.settings.changed_since_save = False
            code_array.shape = (old_shape[0], old_shape[1], 1)
            code_array.sheet_scripts_draft = [None]
            code_array.sheet_scripts_draft[0] = "x = 1"
            panel.current_table = 0
            panel.update_()

            monkeypatch.setitem(
                self.workflows.file_new.__globals__,
                "SheetScriptDraftDialog",
                _DraftDialog,
            )
            monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts",
                                fake_apply_all_sheet_scripts)
            monkeypatch.setattr(GridShapeDialog, "shape", (5, 5, 1))

            self.workflows.file_new()

            assert called["count"] == 0
            assert code_array.shape == (old_shape[0], old_shape[1], 1)
            assert code_array.sheet_scripts_draft[0] == "x = 1"
        finally:
            code_array.shape = old_shape
            code_array.sheet_scripts_draft = old_drafts
            main_window.safe_mode = old_safe_mode
            main_window.settings.changed_since_save = old_changed
            panel.current_table = 0
            panel.update_()

    def test_resolve_sheet_script_drafts_apply_promotes_and_executes(self, monkeypatch):
        """Applying draft resolution should promote drafts and execute scripts."""

        code_array = main_window.grid.model.code_array
        panel = main_window.sheet_script_panel
        old_shape = code_array.shape
        old_macros = list(code_array.sheet_scripts)
        old_drafts = list(code_array.sheet_scripts_draft)
        old_safe_mode = main_window.safe_mode
        old_changed = main_window.settings.changed_since_save

        class _DraftDialog:
            def __init__(self, _parent):
                self.choice = "apply"

        called = {"count": 0}

        def fake_apply_all_sheet_scripts():
            called["count"] += 1
            return 1, 0

        try:
            main_window.safe_mode = False
            main_window.settings.changed_since_save = False
            code_array.shape = (old_shape[0], old_shape[1], 1)
            code_array.sheet_scripts = [""]
            code_array.sheet_scripts_draft = ["x = 11"]
            panel.current_table = 0
            panel.update_()

            monkeypatch.setitem(
                self.workflows._resolve_unapplied_sheet_script_drafts.__globals__,
                "SheetScriptDraftDialog",
                _DraftDialog,
            )
            monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts",
                                fake_apply_all_sheet_scripts)

            assert self.workflows._resolve_unapplied_sheet_script_drafts()
            assert called["count"] == 1
            assert code_array.sheet_scripts[0] == "x = 11"
            assert code_array.sheet_scripts_draft[0] is None
            assert main_window.settings.changed_since_save is True
        finally:
            code_array.shape = old_shape
            code_array.sheet_scripts = old_macros
            code_array.sheet_scripts_draft = old_drafts
            main_window.safe_mode = old_safe_mode
            main_window.settings.changed_since_save = old_changed
            panel.current_table = 0
            panel.update_()

    def test_file_save_cancelled_by_sheet_script_draft_dialog(self, monkeypatch):
        """file_save should abort if draft-resolution dialog is cancelled."""

        class _DraftDialog:
            def __init__(self, _parent):
                self.choice = None

        code_array = main_window.grid.model.code_array
        panel = main_window.sheet_script_panel
        old_drafts = list(code_array.sheet_scripts_draft)
        old_changed = main_window.settings.changed_since_save
        called = {"save": 0}

        def fake_save(_filepath):
            called["save"] += 1
            return True

        try:
            code_array.sheet_scripts_draft[0] = "x = 3"
            panel.current_table = 0
            panel.update_()
            main_window.settings.changed_since_save = False

            monkeypatch.setitem(
                self.workflows.file_save.__globals__,
                "SheetScriptDraftDialog",
                _DraftDialog,
            )
            monkeypatch.setattr(self.workflows, "_save", fake_save)

            assert self.workflows.file_save() is False
            assert called["save"] == 0
        finally:
            code_array.sheet_scripts_draft = old_drafts
            main_window.settings.changed_since_save = old_changed
            panel.current_table = 0
            panel.update_()

    def test_default_sheet_script_template_does_not_trigger_draft_warning(self, monkeypatch):
        """Untouched default template draft should not prompt warning."""

        code_array = main_window.grid.model.code_array
        panel = main_window.sheet_script_panel
        old_drafts = list(code_array.sheet_scripts_draft)
        old_macros = list(code_array.sheet_scripts)

        class _FailIfCalledDialog:
            def __init__(self, _parent):
                raise AssertionError("SheetScriptDraftDialog should not be shown")

        called = {"count": 0}

        def fake_apply_all_sheet_scripts():
            called["count"] += 1
            return 0, 0

        try:
            code_array.sheet_scripts = ["" for _ in range(code_array.shape[2])]
            panel.current_table = 0
            panel.update_()

            monkeypatch.setitem(
                self.workflows._resolve_unapplied_sheet_script_drafts.__globals__,
                "SheetScriptDraftDialog",
                _FailIfCalledDialog,
            )
            monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts",
                                fake_apply_all_sheet_scripts)

            assert self.workflows._resolve_unapplied_sheet_script_drafts()
            assert called["count"] == 0
        finally:
            code_array.sheet_scripts = old_macros
            code_array.sheet_scripts_draft = old_drafts
            panel.current_table = 0
            panel.update_()

    def test_filepath_open_untrusted_defers_scripts_until_approve(self, tmp_path, monkeypatch):
        """Untrusted loads must stay in safe mode and defer script execution."""

        payload = (
            "[PyCellSheet save file version]\n"
            "0.0\n"
            "[shape]\n"
            "1\t1\t1\n"
            "[sheet_names]\n"
            "Sheet 0\n"
            "[sheet_scripts]\n"
            "(sheet_script:'Sheet 0') 1\n"
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

    def test_file_open_recent_delegates_to_filepath_open(self, monkeypatch):
        """file_open_recent should forward to filepath_open with Path argument."""

        called = {"path": None}
        old_changed = main_window.settings.changed_since_save

        def fake_filepath_open(path):
            called["path"] = path

        try:
            main_window.settings.changed_since_save = False
            monkeypatch.setattr(self.workflows, "filepath_open", fake_filepath_open)

            self.workflows.file_open_recent("relative/test.pycsu")

            assert isinstance(called["path"], Path)
            assert called["path"] == Path("relative/test.pycsu")
        finally:
            main_window.settings.changed_since_save = old_changed

    def test_sign_file_skips_signing_in_safe_mode(self, monkeypatch, tmp_path):
        """sign_file should not sign files while safe mode is enabled."""

        filepath = tmp_path / "data.pycsu"
        filepath.write_text("payload", encoding="utf-8")
        old_safe_mode = main_window.safe_mode

        called = {"sign": 0}

        def fake_sign(_data, _key):
            called["sign"] += 1
            return b"sig"

        try:
            main_window.safe_mode = True
            monkeypatch.setitem(self.workflows.sign_file.__globals__, "sign", fake_sign)

            self.workflows.sign_file(filepath)

            assert called["sign"] == 0
            assert main_window.statusBar().currentMessage() == \
                "File saved but not signed because it is unapproved."
            assert not filepath.with_suffix(".pycsu.sig").exists()
        finally:
            main_window.safe_mode = old_safe_mode

    def test_sign_file_reports_error_when_signature_generation_fails(self, monkeypatch, tmp_path):
        """sign_file should report signing error when signer returns no signature."""

        filepath = tmp_path / "data.pycsu"
        filepath.write_text("payload", encoding="utf-8")
        old_safe_mode = main_window.safe_mode

        def fake_sign(_data, _key):
            return None

        try:
            main_window.safe_mode = False
            monkeypatch.setitem(self.workflows.sign_file.__globals__, "sign", fake_sign)

            self.workflows.sign_file(filepath)

            assert main_window.statusBar().currentMessage() == "Error signing file."
            assert not filepath.with_suffix(".pycsu.sig").exists()
        finally:
            main_window.safe_mode = old_safe_mode

    def test_sign_file_writes_signature_and_reports_success(self, monkeypatch, tmp_path):
        """sign_file should write .sig file and report success when signing works."""

        filepath = tmp_path / "data.pycsu"
        filepath.write_bytes(b"payload")
        old_safe_mode = main_window.safe_mode

        def fake_sign(data, _key):
            assert data == b"payload"
            return b"signed-data"

        try:
            main_window.safe_mode = False
            monkeypatch.setitem(self.workflows.sign_file.__globals__, "sign", fake_sign)

            self.workflows.sign_file(filepath)

            sig_path = filepath.with_suffix(".pycsu.sig")
            assert sig_path.exists()
            assert sig_path.read_bytes() == b"signed-data"
            assert main_window.statusBar().currentMessage() == "File saved and signed."
        finally:
            main_window.safe_mode = old_safe_mode

    def test_save_returns_false_when_progress_is_canceled(self, monkeypatch, tmp_path):
        """_save should abort and return False on save-progress cancellation."""

        filepath = tmp_path / "out.pycsu"

        def fake_file_progress_gen(*_args, **_kwargs):
            raise self.workflows._save.__globals__["ProgressDialogCanceled"]
            yield  # pragma: no cover

        monkeypatch.setitem(self.workflows._save.__globals__, "file_progress_gen",
                            fake_file_progress_gen)

        result = self.workflows._save(filepath)

        assert result is False
        assert main_window.statusBar().currentMessage() == "File save stopped by user."

    def test_save_returns_false_when_writer_init_fails(self, monkeypatch, tmp_path):
        """_save should return False when PycsWriter raises during setup."""

        filepath = tmp_path / "out.pycsu"
        calls = {"count": 0}

        class _FailWriter:
            def __init__(self, _code_array):
                raise ValueError("writer failed")

        def fake_critical(*_args, **_kwargs):
            calls["count"] += 1

        monkeypatch.setitem(self.workflows._save.__globals__, "PycsWriter", _FailWriter)
        monkeypatch.setattr(self.workflows._save.__globals__["QMessageBox"],
                            "critical", fake_critical)

        result = self.workflows._save(filepath)

        assert result is False
        assert calls["count"] == 1

    def test_save_returns_false_when_move_fails(self, monkeypatch, tmp_path):
        """_save should return False when moving temp file to destination fails."""

        filepath = tmp_path / "out.pycsu"
        calls = {"critical": 0}

        class _DummyWriter:
            def __init__(self, _code_array):
                self.lines = ["x = 1\n"]

            def __len__(self):
                return 1

            def __iter__(self):
                return iter(self.lines)

        def fake_file_progress_gen(_main_window, iterable, _title, _label, _length):
            for i, line in enumerate(iterable, start=1):
                yield i, line

        def fake_move(_src, _dst):
            raise OSError("move failed")

        def fake_critical(*_args, **_kwargs):
            calls["critical"] += 1

        monkeypatch.setitem(self.workflows._save.__globals__, "PycsWriter", _DummyWriter)
        monkeypatch.setitem(self.workflows._save.__globals__, "file_progress_gen",
                            fake_file_progress_gen)
        monkeypatch.setitem(self.workflows._save.__globals__, "move", fake_move)
        monkeypatch.setattr(self.workflows._save.__globals__["QMessageBox"],
                            "critical", fake_critical)

        result = self.workflows._save(filepath)

        assert result is False
        assert calls["critical"] == 1

    def test_save_success_updates_state_and_signs(self, monkeypatch, tmp_path):
        """_save should update state/history and call sign_file on success."""

        filepath = tmp_path / "ok.pycsu"
        old_changed = main_window.settings.changed_since_save
        old_output = main_window.settings.last_file_output_path
        old_history = list(main_window.settings.file_history)
        calls = {"sign": 0, "menu_update": 0}

        class _DummyWriter:
            def __init__(self, _code_array):
                self.lines = ["x = 1\n"]

            def __len__(self):
                return 1

            def __iter__(self):
                return iter(self.lines)

        def fake_file_progress_gen(_main_window, iterable, _title, _label, _length):
            for i, line in enumerate(iterable, start=1):
                yield i, line

        def fake_sign_file(_filepath):
            calls["sign"] += 1

        def fake_menu_update():
            calls["menu_update"] += 1

        try:
            main_window.settings.changed_since_save = True
            monkeypatch.setitem(self.workflows._save.__globals__, "PycsWriter", _DummyWriter)
            monkeypatch.setitem(self.workflows._save.__globals__, "file_progress_gen",
                                fake_file_progress_gen)
            monkeypatch.setattr(self.workflows, "sign_file", fake_sign_file)
            monkeypatch.setattr(
                main_window.menuBar().file_menu.history_submenu,
                "update_",
                fake_menu_update,
            )

            result = self.workflows._save(filepath)

            assert result is None
            assert filepath.exists()
            assert main_window.settings.changed_since_save is False
            assert main_window.settings.last_file_output_path == filepath
            assert main_window.windowTitle() == "ok.pycsu - PyCellSheet"
            assert calls["sign"] == 1
            assert calls["menu_update"] == 1
            assert main_window.settings.file_history[0] == filepath.as_posix()
        finally:
            main_window.settings.changed_since_save = old_changed
            main_window.settings.last_file_output_path = old_output
            main_window.settings.file_history = old_history

    def test_file_save_uses_save_as_when_no_suffix(self, monkeypatch):
        """file_save should delegate to file_save_as when last output path has no suffix."""

        old_path = main_window.settings.last_file_output_path
        main_window.settings.last_file_output_path = Path("untitled")

        calls = {"save_as": 0, "save": 0}

        try:
            monkeypatch.setattr(self.workflows, "_resolve_unapplied_sheet_script_drafts",
                                lambda: True)
            monkeypatch.setattr(self.workflows, "_save",
                                lambda _filepath: calls.__setitem__("save", calls["save"] + 1))
            monkeypatch.setattr(self.workflows, "file_save_as",
                                lambda: calls.__setitem__("save_as", calls["save_as"] + 1) or "fallback")

            result = self.workflows.file_save()

            assert result == "fallback"
            assert calls["save"] == 0
            assert calls["save_as"] == 1
        finally:
            main_window.settings.last_file_output_path = old_path

    def test_file_save_routes_to_save_as_when_save_fails(self, monkeypatch):
        """file_save should fall back to file_save_as when _save returns False."""

        old_path = main_window.settings.last_file_output_path
        main_window.settings.last_file_output_path = Path("data.pycsu")
        calls = {"save_as": 0}

        try:
            monkeypatch.setattr(self.workflows, "_resolve_unapplied_sheet_script_drafts",
                                lambda: True)
            monkeypatch.setattr(self.workflows, "_save", lambda _filepath: False)
            monkeypatch.setattr(self.workflows, "file_save_as",
                                lambda: calls.__setitem__("save_as", calls["save_as"] + 1) or "fallback")

            result = self.workflows.file_save()

            assert result == "fallback"
            assert calls["save_as"] == 1
        finally:
            main_window.settings.last_file_output_path = old_path

    def test_file_save_as_returns_false_when_dialog_is_canceled(self, monkeypatch):
        """file_save_as should return False when save dialog is cancelled."""

        class _SaveDialog:
            def __init__(self, _parent):
                self.file_path = ""
                self.suffix = ".pycsu"

        monkeypatch.setattr(self.workflows, "_resolve_unapplied_sheet_script_drafts",
                            lambda: True)
        monkeypatch.setitem(self.workflows.file_save_as.__globals__,
                            "FileSaveDialog", _SaveDialog)

        assert self.workflows.file_save_as() is False

    def test_file_save_as_normalizes_suffix_before_save(self, monkeypatch):
        """file_save_as should enforce chosen suffix before calling _save."""

        class _SaveDialog:
            def __init__(self, _parent):
                self.file_path = "report"
                self.suffix = ".pycsu"

        called = {"path": None}

        def fake_save(path):
            called["path"] = path
            return True

        monkeypatch.setattr(self.workflows, "_resolve_unapplied_sheet_script_drafts",
                            lambda: True)
        monkeypatch.setitem(self.workflows.file_save_as.__globals__,
                            "FileSaveDialog", _SaveDialog)
        monkeypatch.setattr(self.workflows, "_save", fake_save)

        assert self.workflows.file_save_as() is True
        assert called["path"] == Path("report.pycsu")

    def test_filepath_open_trusted_applies_scripts_and_reports_errors(self, tmp_path, monkeypatch):
        """Trusted file loads should apply all sheet scripts and report script errors."""

        payload = (
            "[PyCellSheet save file version]\n"
            "0.0\n"
            "[shape]\n"
            "1\t1\t1\n"
            "[sheet_names]\n"
            "Sheet 0\n"
            "[sheet_scripts]\n"
            "(sheet_script:'Sheet 0') 1\n"
            "VALUE = 7\n"
            "[grid]\n"
            "0\t0\t0\t'abc'\n"
            "[attributes]\n"
            "[row_heights]\n"
            "[col_widths]\n"
        )
        filepath = tmp_path / "trusted.pycsu"
        filepath.write_text(payload, encoding="utf-8")
        filepath.with_suffix(".pycsu.sig").write_bytes(b"sig")

        old_safe_mode = main_window.safe_mode
        old_cwd = os.getcwd()
        apply_calls = {"count": 0}

        def fake_file_progress_gen(_main_window, iterable, _title, _label, _lines):
            for i, line in enumerate(iterable, start=1):
                yield i, line

        def fake_verify(_data, _sig, _key):
            return True

        def fake_apply_all_sheet_scripts():
            apply_calls["count"] += 1
            return 2, 1

        try:
            monkeypatch.setitem(self.workflows.filepath_open.__globals__,
                                "file_progress_gen", fake_file_progress_gen)
            monkeypatch.setitem(self.workflows.filepath_open.__globals__,
                                "verify", fake_verify)
            monkeypatch.setattr(self.workflows, "apply_all_sheet_scripts",
                                fake_apply_all_sheet_scripts)

            self.workflows.filepath_open(filepath)

            assert main_window.safe_mode is False
            assert apply_calls["count"] == 1
            assert main_window.statusBar().currentMessage() == \
                "Applied 2 sheet scripts (1 with errors)."
        finally:
            os.chdir(old_cwd)
            main_window.safe_mode = old_safe_mode
