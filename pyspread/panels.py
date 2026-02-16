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

**Provides**

 * :class:`SheetScriptPanel`

"""

import ast
from io import StringIO
from sys import exc_info
from traceback import print_exception

from PyQt6.QtCore import Qt, pyqtProperty, QModelIndex
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QDialogButtonBox, QSplitter
from PyQt6.QtWidgets import QTextEdit, QLabel, QWidget

try:
    from pyspread.lib.spelltextedit import SpellTextEdit
    from pyspread.lib.exception_handling import get_user_codeframe
except ImportError:
    from lib.spelltextedit import SpellTextEdit
    from lib.exception_handling import get_user_codeframe


class SheetScriptPanel(QDialog):
    """The sheet script panel"""

    class AppliedIndicator(QLabel):

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setAlignment(Qt.AlignmentFlag.AlignHCenter)
            self._applied = True
            self._update_text_and_style()
            if parent:
                parent.installEventFilter(self)

        def _update_text_and_style(self):
            self.setText("Applied" if self._applied else "Draft")
            self.setToolTip("Applied" if self._applied
                            else "Draft changes pending")
            self.stylesheet_update()

        def stylesheet_update(self):
            # Qt does not guarantee the applying of new style
            #  when a property is updated, unless you re-apply the stylesheet.
            self.setStyleSheet("")
            self.setStyleSheet(
                "QLabel {background-color: #1f000000; font-weight: bold; padding: 1px} \n"
                "QLabel[applied=true] { color: blue } \n"
                "QLabel[applied=false] { color: red } \n"
            )

        def _reposition(self):
            if self.parent():
                self.move(
                    self.parent().width() - self.width(), 0
                )

        def eventFilter(self, obj, event):
            if event.type() == event.Type.Resize:
                self._reposition()
            return False

        def showEvent(self, event):
            super().showEvent(event)
            self._reposition()

        @pyqtProperty(bool)
        def applied(self):
            return self._applied

        @applied.setter
        def applied(self, applied):
            self._applied = bool(applied)
            self._update_text_and_style()

    def __init__(self, parent, code_array):
        super().__init__()

        self.parent = parent
        self.code_array = code_array
        self.current_table = 0
        self._implicit_default_draft = self.code_array.macros_draft[0] \
            if self.code_array.macros_draft else None

        self._init_widgets()
        self._layout()

        self.update_()

        self.default_text_color = self.result_viewer.textColor()
        self.error_text_color = QColor("red")

        self.button_box.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.on_apply)
        self.button_box.button(QDialogButtonBox.StandardButton.Reset).clicked.connect(self.on_reset)

    def _init_widgets(self):
        """Inititialize widgets"""

        font_family = self.parent.settings.macro_editor_font_family
        self.macro_editor = SpellTextEdit(self, font_family=font_family)
        self.macro_editor.setPlaceholderText(
            "Sheet Script draft. Press Apply to execute and store for this sheet."
        )
        self.applied_indicator = SheetScriptPanel.AppliedIndicator(self.macro_editor)
        self.macro_editor.textChanged.connect(
            lambda: setattr(self.applied_indicator, "applied", False)
        )

        self.result_viewer = QTextEdit(self)
        self.result_viewer.setReadOnly(True)
        self.result_viewer.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.result_viewer.setPlaceholderText(
            "Applied output and errors appear here."
        )
        self.result_viewer.setFontFamily(font_family)

        self.splitter = QSplitter(Qt.Orientation.Vertical, self)

        self.editor_label = QLabel("Script Editor")
        self.editor_label.setStyleSheet("font-weight: bold;")
        self.editor_section = QWidget(self)
        self.editor_layout = QVBoxLayout(self.editor_section)
        self.editor_layout.setContentsMargins(0, 0, 0, 0)
        self.editor_layout.setSpacing(4)
        self.editor_layout.addWidget(self.editor_label)
        self.editor_layout.addWidget(self.macro_editor)

        self.output_label = QLabel("Apply Output")
        self.output_label.setStyleSheet("font-weight: bold;")
        self.output_section = QWidget(self)
        self.output_layout = QVBoxLayout(self.output_section)
        self.output_layout.setContentsMargins(0, 0, 0, 0)
        self.output_layout.setSpacing(4)
        self.output_layout.addWidget(self.output_label)
        self.output_layout.addWidget(self.result_viewer)

        self.splitter.addWidget(self.editor_section)
        self.splitter.addWidget(self.output_section)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(8)
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 2)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Apply |
            QDialogButtonBox.StandardButton.Reset
        )
        self.button_box.button(
            QDialogButtonBox.StandardButton.Apply
        ).setText("Apply Script")
        self.button_box.button(
            QDialogButtonBox.StandardButton.Reset
        ).setText("Revert Draft")

    def _layout(self):
        """Layout dialog widgets"""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)
        layout.addWidget(self.splitter)
        layout.addWidget(self.button_box)

        self.setLayout(layout)
        self.splitter.setSizes([380, 220])

    def _is_invalid_code(self, current_table) -> str:
        """Preliminary code check

        Returns a string with the error message if code is not valid Python.
        If the code runs without errors, an empty string is returned.

        """

        try:
            ast.parse(self.code_array.macros[current_table])

        except Exception:
            # Grab the traceback and return it
            stringio = StringIO()
            excinfo = exc_info()
            # usr_tb will more than likely be none because ast throws
            # SytnaxErrors as occurring outside of the current execution frame
            usr_tb = get_user_codeframe(excinfo[2]) or None
            print_exception(excinfo[0], excinfo[1], usr_tb, None, stringio)
            return stringio.getvalue()
        else:
            return ''

    def on_apply(self):
        """Event handler for Apply button"""

        self.code_array.macros[self.current_table] = self.macro_editor.toPlainText()

        err = self._is_invalid_code(self.current_table)
        if err:
            self.update_result_viewer(err=err)
        else:
            self.update_result_viewer(*self.code_array.execute_macros(self.current_table))
            self.code_array.macros_draft[self.current_table] = None
            self.applied_indicator.applied = True
            self.parent.grid.model.emit_data_changed_all()

        self.parent.grid.gui_update()

    def update_(self):
        """Update macro content"""
        # NYI: store&load evaluation results

        if self.code_array.macros_draft[self.current_table] is not None:
            self.macro_editor.setPlainText(self.code_array.macros_draft[self.current_table])
            self.applied_indicator.applied = False
        else:
            self.macro_editor.setPlainText(self.code_array.macros[self.current_table])
            self.applied_indicator.applied = True

    def update_current_table(self, current):
        self.persist_current_draft()
        self.current_table = current
        self.update_()

    def persist_current_draft(self):
        """Persist current editor content as draft if not applied."""

        if not self.applied_indicator.applied:
            self.code_array.macros_draft[self.current_table] = self.macro_editor.toPlainText()

    def has_unapplied_drafts(self) -> bool:
        """Returns whether any sheet has unapplied draft content."""

        self.persist_current_draft()
        for idx, draft in enumerate(self.code_array.macros_draft):
            if draft is None:
                continue
            if self._is_implicit_default_draft(idx, draft):
                continue
            if draft != self.code_array.macros[idx]:
                return True
        return False

    def _is_implicit_default_draft(self, table: int, draft: str) -> bool:
        """True for untouched default template drafts."""

        return self.code_array.macros[table] == "" and draft == self._implicit_default_draft

    def discard_all_drafts(self):
        """Discard all unapplied drafts across all tables."""

        for idx in range(len(self.code_array.macros_draft)):
            self.code_array.macros_draft[idx] = None
        self.update_()

    def apply_all_drafts_to_scripts(self) -> int:
        """Apply all draft scripts to applied scripts.

        Returns the number of tables that were updated from draft content.
        """

        self.persist_current_draft()
        updated = 0
        for idx, draft in enumerate(self.code_array.macros_draft):
            if draft is None:
                continue
            self.code_array.macros[idx] = draft
            self.code_array.macros_draft[idx] = None
            updated += 1
        self.update_()
        return updated

    def update_result_viewer(self, result: str = "", err: str = ""):
        """Update event result following execution by main window

        :param result: Text to be shown in the result viewer in default color
        :param err: Text to be shown in the result viewer in error text color

        """

        self.result_viewer.clear()

        if result:
            self.result_viewer.append(result)
        if err:
            self.result_viewer.setTextColor(self.error_text_color)
            self.result_viewer.append(err)
            self.result_viewer.setTextColor(self.default_text_color)

    def on_reset(self):
        self.code_array.macros_draft[self.current_table] = None
        self.macro_editor.setPlainText(self.code_array.macros[self.current_table])
        self.applied_indicator.applied = True
