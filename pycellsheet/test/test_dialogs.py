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

"""
test_dialogs
============

Non-interactive dialog branch tests.
"""

import pytest

from PyQt6.QtWidgets import QApplication, QDialog, QMessageBox

from ..dialogs import (
    ApproveWarningDialog,
    DataEntryDialog,
    DiscardChangesDialog,
    ExpressionParserSelectionDialog,
    ExpressionParserMigrationDialog,
    SheetScriptDraftDialog,
)


app = QApplication.instance()
if app is None:
    app = QApplication([])


@pytest.mark.parametrize(
    "button, expected",
    [
        (QMessageBox.StandardButton.Discard, True),
        (QMessageBox.StandardButton.Save, False),
        (QMessageBox.StandardButton.Cancel, None),
    ],
)
def test_discard_changes_dialog_choice_mapping(monkeypatch, button, expected):
    monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: button)
    assert DiscardChangesDialog(None).choice is expected


@pytest.mark.parametrize(
    "button, expected",
    [
        (QMessageBox.StandardButton.Apply, "apply"),
        (QMessageBox.StandardButton.Discard, "discard"),
        (QMessageBox.StandardButton.Cancel, None),
    ],
)
def test_sheet_script_draft_dialog_choice_mapping(monkeypatch, button, expected):
    monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: button)
    assert SheetScriptDraftDialog(None).choice == expected


@pytest.mark.parametrize(
    "button, expected",
    [
        (QMessageBox.StandardButton.Yes, True),
        (QMessageBox.StandardButton.No, False),
        (QMessageBox.StandardButton.Cancel, None),
    ],
)
def test_approve_warning_dialog_choice_mapping(monkeypatch, button, expected):
    monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: button)
    assert ApproveWarningDialog(None).choice is expected


def test_data_entry_dialog_rejects_mismatched_initial_data():
    with pytest.raises(ValueError, match="Length of labels and initial_data not equal"):
        DataEntryDialog(None, "t", ["a", "b"], initial_data=["x"])


def test_data_entry_dialog_rejects_mismatched_validators():
    with pytest.raises(ValueError, match="Length of labels and validators not equal"):
        DataEntryDialog(None, "t", ["a", "b"], validators=[None])


def test_data_entry_dialog_data_returns_tuple_when_accepted(monkeypatch):
    dialog = DataEntryDialog(None, "t", ["A", "B"], initial_data=["x", "y"])
    monkeypatch.setattr(dialog, "exec", lambda: QDialog.DialogCode.Accepted)
    assert dialog.data == ("x", "y")


def test_data_entry_dialog_data_returns_none_when_rejected(monkeypatch):
    dialog = DataEntryDialog(None, "t", ["A"], initial_data=["x"])
    monkeypatch.setattr(dialog, "exec", lambda: QDialog.DialogCode.Rejected)
    assert dialog.data is None


def test_expression_parser_selection_dialog_exposes_selected_values():
    dialog = ExpressionParserSelectionDialog(
        None, current_mode_id="mixed"
    )

    target_idx = dialog.mode_combo.findData("pure_spreadsheet")
    dialog.mode_combo.setCurrentIndex(target_idx)

    mode_id = dialog.selection
    assert mode_id == "pure_spreadsheet"


def test_expression_parser_migration_dialog_preview_and_apply():
    class _Report:
        def __init__(self, source, target):
            self.source_mode_id = source
            self.target_mode_id = target
            self.summary = {
                "safe_changed": 1,
                "risky_skipped": 2,
                "unchanged": 3,
                "invalid_source_assumption": 0,
                "total": 6,
            }

    class _CodeArray:
        def __init__(self):
            self.preview_calls = []
            self.apply_calls = []

        def preview_expression_parser_migration(self, **kwargs):
            self.preview_calls.append(kwargs)
            return _Report(kwargs["source_mode_id"], kwargs["target_mode_id"])

        def apply_expression_parser_migration(self, **kwargs):
            self.apply_calls.append(kwargs)
            return _Report(kwargs["source_mode_id"], kwargs["target_mode_id"])

    code_array = _CodeArray()
    dialog = ExpressionParserMigrationDialog(
        None, code_array, current_mode_id="mixed", current_table=1
    )

    target_index = dialog.target_mode.findData("pure_spreadsheet")
    dialog.target_mode.setCurrentIndex(target_index)
    scope_index = dialog.scope.findText("Current sheet only")
    dialog.scope.setCurrentIndex(scope_index)
    dialog.include_risky.setChecked(True)

    dialog.on_preview()
    assert code_array.preview_calls
    assert "safe" in dialog.summary.toPlainText().lower()
    assert dialog.apply_button.isEnabled()

    dialog.on_apply()
    assert code_array.apply_calls
    assert dialog.applied_report is not None


def test_expression_parser_migration_dialog_requires_different_modes():
    class _CodeArray:
        def preview_expression_parser_migration(self, **kwargs):
            raise AssertionError("preview should not run when source == target")

        def apply_expression_parser_migration(self, **kwargs):
            raise AssertionError("apply should not run when source == target")

    dialog = ExpressionParserMigrationDialog(
        None, _CodeArray(), current_mode_id="mixed", current_table=0
    )
    index = dialog.source_mode.findData("mixed")
    dialog.source_mode.setCurrentIndex(index)
    dialog.target_mode.setCurrentIndex(index)

    dialog.on_preview()
    assert "must be different" in dialog.summary.toPlainText().lower()
    assert not dialog.apply_button.isEnabled()

