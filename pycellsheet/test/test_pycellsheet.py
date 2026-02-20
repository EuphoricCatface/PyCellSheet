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
test_grid
=========

Unit tests for grid.py

"""

from contextlib import contextmanager
from copy import deepcopy
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
    from ..pycellsheet import MainWindow
    from ..commands import MakeButtonCell, RemoveButtonCell
    from ..lib.selection import Selection
    from ..model.model import INITSCRIPT_DEFAULT


app = QApplication.instance()
if app is None:
    app = QApplication([])
main_window = MainWindow(prompt_parser_dialog_on_startup=False)
zoom_levels = main_window.settings.zoom_levels


class TestMainWindow:
    """Unit tests for MainWindow in pycellsheet.py"""

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

    def test_startup_in_no_document_state(self):
        """Fresh startup should not initialize a workbook until New/Open."""

        code_array = main_window.grid.model.code_array
        tables = code_array.shape[2]

        assert not main_window.has_document
        assert code_array.sheet_scripts == ["" for _ in range(tables)]
        assert code_array.sheet_scripts_draft == [
            INITSCRIPT_DEFAULT for _ in range(tables)
        ]

    def test_sheet_script_globals_are_isolated_per_sheet(self):
        """Sheet Script globals should not leak across sheets."""

        code_array = main_window.grid.model.code_array
        old_macro0 = code_array.sheet_scripts[0]
        old_macro1 = code_array.sheet_scripts[1]
        old_draft0 = code_array.sheet_scripts_draft[0]
        old_draft1 = code_array.sheet_scripts_draft[1]
        old_copyable0 = deepcopy(code_array.sheet_globals_copyable[0])
        old_copyable1 = deepcopy(code_array.sheet_globals_copyable[1])
        # Uncopyable globals may include module objects (e.g. random),
        # so keep shallow snapshots of dict mappings for restore.
        old_uncopyable0 = dict(code_array.sheet_globals_uncopyable[0])
        old_uncopyable1 = dict(code_array.sheet_globals_uncopyable[1])

        key_sheet0 = (0, 0, 0)
        key_sheet1 = (0, 0, 1)
        old_parser_code = code_array.exp_parser_code

        try:
            code_array.set_exp_parser_mode("mixed")
            code_array.sheet_scripts[0] = "RATE = 7"
            code_array.sheet_scripts[1] = ""
            code_array.execute_sheet_script(0)
            code_array.execute_sheet_script(1)

            code_array[key_sheet0] = "RATE"
            code_array[key_sheet1] = "RATE"

            assert code_array[key_sheet0] == 7
            assert isinstance(code_array[key_sheet1], NameError)
        finally:
            for key in (key_sheet0, key_sheet1):
                try:
                    code_array.pop(key)
                except KeyError:
                    pass
                code_array.dep_graph.remove_cell(key)

            code_array.sheet_scripts[0] = old_macro0
            code_array.sheet_scripts[1] = old_macro1
            code_array.sheet_scripts_draft[0] = old_draft0
            code_array.sheet_scripts_draft[1] = old_draft1
            code_array.sheet_globals_copyable[0] = old_copyable0
            code_array.sheet_globals_copyable[1] = old_copyable1
            code_array.sheet_globals_uncopyable[0] = old_uncopyable0
            code_array.sheet_globals_uncopyable[1] = old_uncopyable1
            code_array.exp_parser_code = old_parser_code
            code_array.smart_cache.clear()
            code_array.dep_graph.dirty.clear()

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

    def test_recalculate_actions_show_scope_and_count_status(self, monkeypatch):
        """Recalculate actions should report consistent scope/count status text."""

        code_array = main_window.grid.model.code_array
        old_mode = main_window.settings.recalc_mode
        main_window.settings.recalc_mode = "manual"

        refresh_calls = {"count": 0}

        try:
            monkeypatch.setattr(code_array, "recalculate_dirty", lambda: 2)
            monkeypatch.setattr(code_array, "recalculate_cell_only", lambda _key: 1)
            monkeypatch.setattr(code_array, "recalculate_ancestors", lambda _key: 0)
            monkeypatch.setattr(code_array, "recalculate_children", lambda _key: 3)
            monkeypatch.setattr(code_array, "recalculate_all", lambda: 4)
            monkeypatch.setattr(main_window, "_refresh_grid",
                                lambda: refresh_calls.__setitem__("count", refresh_calls["count"] + 1))

            main_window.on_recalculate()
            assert main_window.statusBar().currentMessage() == \
                "Recalculated 2 cells (scope: Dirty cells, mode: Manual)"

            main_window.on_recalculate_cell_only()
            assert main_window.statusBar().currentMessage() == \
                "Recalculated 1 cell (scope: Current cell, mode: Manual)"

            main_window.on_recalculate_ancestors()
            assert main_window.statusBar().currentMessage() == \
                "No cells needed recalculation (scope: Current cell + ancestors, mode: Manual)"

            main_window.on_recalculate_children()
            assert main_window.statusBar().currentMessage() == \
                "Recalculated 3 cells (scope: Current cell + children, mode: Manual)"

            main_window.on_recalculate_all()
            assert main_window.statusBar().currentMessage() == \
                "Recalculated 4 cells (scope: Entire workspace, mode: Manual)"

            # on_recalculate_ancestors returns 0 and should not trigger refresh
            assert refresh_calls["count"] == 4
        finally:
            main_window.settings.recalc_mode = old_mode

    def test_help_and_preferences_action_branding(self):
        """User-facing action labels/tips should use PyCellSheet branding."""

        actions = main_window.main_window_actions

        assert actions.about.text() == "About PyCellSheet..."
        assert actions.about.statusTip() == "About PyCellSheet"
        assert actions.manual.statusTip() == "Display the PyCellSheet manual"
        assert actions.tutorial.statusTip() == "Display a PyCellSheet tutorial"
        assert actions.preferences.statusTip() == "PyCellSheet setup parameters"

    def test_expression_parser_actions_present(self):
        """Tools menu should expose parser settings and migration actions."""

        actions = main_window.main_window_actions
        assert actions.expression_parser_settings.text() == "Expression Parser Settings..."
        assert actions.migrate_expression_parser.text() == "Migrate Parser Mode..."

    def test_set_expression_parser_mode_updates_state_and_status(self):
        """Setting parser mode should update model mode and status text."""

        code_array = main_window.grid.model.code_array
        old_code = code_array.exp_parser_code

        try:
            main_window.on_set_expression_parser_mode("mixed", True)
            assert code_array.exp_parser_mode_id == "mixed"
            assert "expression parser mode set to mixed" in \
                main_window.statusBar().currentMessage().lower()

            main_window.on_set_expression_parser_mode("pure_spreadsheet", True)
            assert code_array.exp_parser_mode_id == "pure_spreadsheet"
            assert "pure_spreadsheet" in main_window.statusBar().currentMessage()
        finally:
            code_array.exp_parser_code = old_code
            main_window.update_action_toggles()

    def test_entryline_parser_indicator_shows_parser_mode(self):
        """Entry line parser badge should reflect current parser mode."""

        code_array = main_window.grid.model.code_array
        old_code = code_array.exp_parser_code

        try:
            code_array.exp_parser_code = "return cell.strip()"
            main_window.update_action_toggles()
            assert "Parser: Custom" in main_window.entry_line.parser_indicator.text()

            main_window.on_set_expression_parser_mode("mixed", True)
            assert "Parser: Mixed" in main_window.entry_line.parser_indicator.text()
        finally:
            code_array.exp_parser_code = old_code
            main_window.update_action_toggles()

    def test_open_expression_parser_settings_dialog_accept(self, monkeypatch):
        """Parser settings dialog should apply selected mode."""

        code_array = main_window.grid.model.code_array
        old_code = code_array.exp_parser_code

        class _Dialog:
            def __init__(self, *_args, **_kwargs):
                self.selection = "mixed"

            def exec(self):
                return True

        try:
            monkeypatch.setitem(
                main_window.on_open_expression_parser_settings_dialog.__globals__,
                "ExpressionParserSelectionDialog",
                _Dialog,
            )
            accepted = main_window.on_open_expression_parser_settings_dialog()
            assert accepted is True
            assert code_array.exp_parser_mode_id == "mixed"
        finally:
            code_array.exp_parser_code = old_code

    def test_open_expression_parser_settings_dialog_cancel_exits_when_requested(
            self, monkeypatch):
        """Parser settings cancel should request app quit when configured."""

        calls = {"quit": 0}

        class _Dialog:
            def __init__(self, *_args, **_kwargs):
                self.selection = "mixed"

            def exec(self):
                return False

        class _App:
            def quit(self):
                calls["quit"] += 1

        def fake_single_shot(_ms, callback):
            callback()

        monkeypatch.setitem(
            main_window.on_open_expression_parser_settings_dialog.__globals__,
            "ExpressionParserSelectionDialog",
            _Dialog,
        )
        monkeypatch.setitem(
            main_window.on_open_expression_parser_settings_dialog.__globals__,
            "QTimer",
            type("_Timer", (), {"singleShot": staticmethod(fake_single_shot)}),
        )
        monkeypatch.setitem(
            main_window.on_open_expression_parser_settings_dialog.__globals__,
            "QApplication",
            type("_QApplication", (), {"instance": staticmethod(lambda: _App())}),
        )

        accepted = main_window.on_open_expression_parser_settings_dialog(
            exit_on_cancel=True
        )
        assert accepted is False
        assert calls["quit"] == 1

    def test_open_expression_parser_migration_dialog_applies_report(self, monkeypatch):
        """Migration dialog acceptance should refresh and report migration summary."""

        calls = {"refresh": 0}

        class _Report:
            summary = {"safe_changed": 2, "risky_skipped": 1, "unchanged": 3}

        class _Dialog:
            def __init__(self, _parent, _code_array, current_mode_id, current_table):
                self.current_mode_id = current_mode_id
                self.current_table = current_table
                self.applied_report = _Report()

            def exec(self):
                return True

        monkeypatch.setitem(
            main_window.on_open_expression_parser_migration_dialog.__globals__,
            "ExpressionParserMigrationDialog",
            _Dialog,
        )
        monkeypatch.setattr(
            main_window,
            "_refresh_grid",
            lambda: calls.__setitem__("refresh", calls["refresh"] + 1),
        )

        main_window.on_open_expression_parser_migration_dialog()

        assert calls["refresh"] == 1
        assert "Parser migration applied" in main_window.statusBar().currentMessage()
        assert "safe: 2" in main_window.statusBar().currentMessage()
