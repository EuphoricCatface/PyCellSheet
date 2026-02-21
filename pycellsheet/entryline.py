# -*- coding: utf-8 -*-

# Copyright Martin Manns
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

**Provides**

 * :class:`Entryline`

"""

from contextlib import contextmanager

from PyQt6.QtCore import Qt, QEvent
from PyQt6.QtGui import QFontMetrics, QTextOption, QKeyEvent
from PyQt6.QtWidgets import QWidget, QMainWindow, QLabel

try:
    import pycellsheet.commands as commands
    from pycellsheet.lib.spelltextedit import SpellTextEdit
    from pycellsheet.lib.pycellsheet import ExpressionParser
    from pycellsheet.lib.string_helpers import quote
except ImportError:
    import commands
    from lib.spelltextedit import SpellTextEdit
    from lib.pycellsheet import ExpressionParser
    from lib.string_helpers import quote


class Entryline(SpellTextEdit):
    """The entry line for PyCellSheet"""

    class ParserIndicator(QLabel):
        """Shows current parser mode."""

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
            self.setStyleSheet("")
            if parent:
                parent.installEventFilter(self)
                try:
                    scrollbar = parent.verticalScrollBar()
                    scrollbar.rangeChanged.connect(lambda *_: self._reposition())
                    scrollbar.valueChanged.connect(lambda *_: self._reposition())
                except Exception:
                    pass

        def _reposition(self):
            if self.parent():
                self.adjustSize()
                scrollbar = self.parent().verticalScrollBar()
                scrollbar_width = scrollbar.width() if scrollbar.isVisible() else 0
                self.move(
                    max(0, self.parent().width() - self.width() - scrollbar_width - 6),
                    2
                )

        def eventFilter(self, obj, event):
            if event.type() == event.Type.Resize:
                self._reposition()
            return False

        def showEvent(self, event):
            super().showEvent(event)
            self._reposition()

        def update_state(self, parser_label: str):
            self.setText(parser_label)
            self.setToolTip(f"Expression Parser: {parser_label}")
            self.setStyleSheet(
                "QLabel {background-color: #1f000000; font-weight: bold; padding: 1px;}\n"
                "QLabel { color: blue; }"
            )
            self._reposition()

    def __init__(self, main_window: QMainWindow):
        """

        :param main_window: Application main window

        """

        font_family = main_window.settings.entry_line_font_family
        super().__init__(line_numbers=False, font_family=font_family)

        self.main_window = main_window

        min_height = QFontMetrics(self.font()).height() + 20
        self.setMinimumHeight(min_height)

        self.setWordWrapMode(QTextOption.WrapMode.WrapAnywhere)

        self.installEventFilter(self)

        self.last_key = None
        self.parser_indicator = Entryline.ParserIndicator(self)

        # self.highlighter.setDocument(self.document())

    def update_parser_indicator(self):
        """Update parser indicator text and state badge."""

        if not hasattr(self.main_window, "grid"):
            return
        code_array = self.main_window.grid.model.code_array
        mode_id = code_array.exp_parser_mode_id
        if mode_id is None:
            parser_label = "Parser: Custom"
        else:
            parser_label = "Parser: " + ExpressionParser.MODE_ID_TO_LABEL.get(
                mode_id, mode_id
            )
        self.parser_indicator.update_state(parser_label=parser_label)

    @contextmanager
    def disable_updates(self):
        """Disables updates and highlighter"""

        doc = self.highlighter.document()
        self.highlighter.setDocument(None)
        self.main_window.entry_line.setUpdatesEnabled(False)
        yield
        self.main_window.entry_line.setUpdatesEnabled(True)
        self.highlighter.setDocument(doc)

    def keyPressEvent(self, event: QKeyEvent):
        """Key press event filter

        :param event: Key event

        """

        grid = self.main_window.focused_grid

        self.last_key = event.key()

        if self.last_key in (Qt.Key.Key_Enter, Qt.Key.Key_Return):
            if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                self.insertPlainText('\n')
            else:
                if grid.selection_mode:
                    grid.set_selection_mode(False)
                else:
                    self.store_data()
                    grid.row += 1
            return
        elif self.last_key == Qt.Key.Key_Tab:
            self.store_data()
            grid.column += 1
            return
        elif self.last_key == Qt.Key.Key_Escape:
            grid.on_current_changed()
            grid.setFocus()
            return
        elif self.last_key == Qt.Key.Key_Insert:
            grid.set_selection_mode(not grid.selection_mode)
            return

        return super().keyPressEvent(event)

    def store_data(self):
        """Stores current entry line data in grid model"""

        grid = self.main_window.focused_grid

        index = grid.currentIndex()
        model = grid.model

        description = "Set code for cell {}".format(index)
        command = commands.SetCellCode(self.toPlainText(), model, index,
                                       description)
        self.main_window.undo_stack.push(command)

    def on_toggle_spell_check(self, signal: bool):
        """Spell check toggle event handler

        :param signal: Spell check is enabled if True

        """

        self.highlighter.enable_enchant = bool(signal)

    def setPlainText(self, text: str):
        """Overides setPlainText

        Additionally shows busy cursor and disables highlighter on long texts,
        and omits identical replace.

        :param text: Text to be set

        """

        is_long = (text is not None
                   and len(text) > self.main_window.settings.highlighter_limit)

        if text == self.toPlainText():
            return

        if is_long:
            with self.main_window.workflows.busy_cursor():
                self.highlighter.setDocument(None)
                super().setPlainText(text)
        else:
            super().setPlainText(text)
