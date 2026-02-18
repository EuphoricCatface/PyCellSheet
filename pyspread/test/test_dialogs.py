# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
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
