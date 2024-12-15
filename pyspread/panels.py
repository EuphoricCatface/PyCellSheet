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

**Provides**

 * :class:`MacroPanel`

"""

import ast
from io import StringIO
from sys import exc_info
from traceback import print_exception

from PyQt6.QtCore import Qt, pyqtProperty
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QDialogButtonBox, QSplitter
from PyQt6.QtWidgets import QTextEdit, QLabel

try:
    from pyspread.lib.spelltextedit import SpellTextEdit
    from pyspread.lib.exception_handling import get_user_codeframe
except ImportError:
    from lib.spelltextedit import SpellTextEdit
    from lib.exception_handling import get_user_codeframe


class MacroPanel(QDialog):
    """The macro panel"""

    class AppliedIndicator(QLabel):

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setText("Draft")
            self.setAlignment(Qt.AlignmentFlag.AlignHCenter)
            self._applied = True
            self.stylesheet_update()

        def stylesheet_update(self):
            self.setStyleSheet("")
            self.setStyleSheet(
                "QLabel {background-color: #1f000000; font-weight: bold; padding: 1px} \n"
                "QLabel[applied=true] { color: blue } \n"
                "QLabel[applied=false] { color: red } \n"
            )

        def paintEvent(self, a0):
            super().paintEvent(a0)
            if self.parent():
                self.move(
                    self.parent().width() - self.width(), 0
                )

        @pyqtProperty(bool)
        def applied(self):
            return self._applied

        @applied.setter
        def applied(self, applied):
            if self._applied == applied:
                return
            self._applied = applied
            self.setText("Applied" if applied else "Draft")
            self.stylesheet_update()

    def __init__(self, parent, code_array):
        super().__init__()

        self.parent = parent
        self.code_array = code_array
        self.current_table = 0

        self._init_widgets()
        self._layout()

        self.update()

        self.default_text_color = self.result_viewer.textColor()
        self.error_text_color = QColor("red")

        self.button_box.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.on_apply)

    def _init_widgets(self):
        """Inititialize widgets"""

        font_family = self.parent.settings.macro_editor_font_family
        self.macro_editor = SpellTextEdit(self, font_family=font_family)
        self.applied_indicator = MacroPanel.AppliedIndicator(self.macro_editor)
        self.macro_editor.textChanged.connect(
            lambda: self.applied_indicator.setProperty("applied", False)
        )

        self.result_viewer = QTextEdit(self)
        self.result_viewer.setReadOnly(True)

        self.splitter = QSplitter(Qt.Orientation.Vertical, self)

        self.splitter.addWidget(self.macro_editor)
        self.splitter.addWidget(self.result_viewer)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Apply |
            QDialogButtonBox.StandardButton.Reset
        )

    def _layout(self):
        """Layout dialog widgets"""

        layout = QVBoxLayout(self)
        layout.addWidget(self.splitter)
        layout.addWidget(self.button_box)

        self.setLayout(layout)

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
            if self.current_table in self.code_array.dict_grid.macros_draft:
                del self.code_array.dict_grid.macros_draft[self.current_table]
            self.applied_indicator.setProperty("applied", True)

        self.parent.grid.gui_update()

    def update(self):
        """Update macro content"""

        if self.current_table in self.code_array.dict_grid.macros_draft:
            self.macro_editor.setPlainText(self.code_array.dict_grid.macros_draft[self.current_table])
            self.applied_indicator.setProperty("applied", False)
        else:
            self.macro_editor.setPlainText(self.code_array.dict_grid.macros[self.current_table])
            self.applied_indicator.setProperty("applied", True)

    def update_current_table(self, current):
        if not self.applied_indicator.property("applied"):
            self.code_array.dict_grid.macros_draft[self.current_table] = self.macro_editor.toPlainText()
        self.current_table = current
        self.update()

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
