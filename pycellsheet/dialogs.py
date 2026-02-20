#!/usr/bin/env python3
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


**Modal dialogs**

 * :class:`DiscardChangesDialog`
 * :class:`DiscardDataDialog`
 * :class:`ApproveWarningDialog`
 * :class:`DataEntryDialog`
 * :class:`GridShapeDialog`
 * :class:`SinglePageArea`
 * :class:`MultiPageArea`
 * :class:`CsvExportAreaDialog`
 * :class:`SvgExportAreaDialog`
 * :class:`PrintAreaDialog`
 * (:class:`FileDialogBase`)
 * :class:`FileOpenDialog`
 * :class:`FileSaveDialog`
 * :class:`ImageFileOpenDialog`
 * :class:`CsvFileImportDialog`
 * :class:`FileExportDialog`
 * :class:`FindDialog`
 * :class:`ChartDialog`
 * :class:`CsvImportDialog`
 * :class:`CsvExportDialog`
 * :class:`TutorialDialog`
 * :class:`ManualDialog`
 * :class:`PrintPreviewDialog`
 * :class:`StartupGreeterDialog`
 * :class:`ExpressionParserSelectionDialog`
 * :class:`ExpressionParserMigrationDialog`
"""

from contextlib import redirect_stdout
import csv
from dataclasses import dataclass
from functools import partial
import io
from pathlib import Path
from typing import List, Sequence, Tuple, Union

from PyQt6.QtCore import Qt, QPoint, QSize, QEvent, QAbstractTableModel, QModelIndex
from PyQt6.QtWidgets \
    import (QApplication, QMessageBox, QFileDialog, QDialog, QLineEdit, QLabel,
            QFormLayout, QVBoxLayout, QGroupBox, QDialogButtonBox, QSplitter,
            QTextBrowser, QCheckBox, QGridLayout, QLayout, QHBoxLayout,
            QPushButton, QWidget, QComboBox, QTableView, QAbstractItemView,
            QPlainTextEdit, QToolBar, QMainWindow, QTabWidget, QInputDialog,
            QSpinBox,
            )
from PyQt6.QtGui \
    import (QIntValidator, QImageWriter, QStandardItemModel, QStandardItem,
            QValidator, QWheelEvent)
from PyQt6.QtSvgWidgets import QSvgWidget
from PyQt6.QtPrintSupport import (QPrintPreviewDialog, QPrintPreviewWidget,
                                  QPrinter)

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
except ImportError:
    Figure = None

try:
    import openpyxl
    import pycel
except ImportError:
    openpyxl = None

try:
    import moneyed
except ImportError:
    moneyed = None

try:
    from pycellsheet.actions import ChartDialogActions
    from pycellsheet.toolbar import ChartTemplatesToolBar, RChartTemplatesToolBar
    from pycellsheet.widgets import HelpBrowser, TypeMenuComboBox
    from pycellsheet.lib.csv import sniff, csv_reader, get_header, convert
    from pycellsheet.lib.pycellsheet import ExpressionParser, Formatter
    from pycellsheet.lib.spelltextedit import SpellTextEdit
    from pycellsheet.model.model import INITSCRIPT_DEFAULT, CodeArray
    from pycellsheet.settings import (TUTORIAL_PATH, MANUAL_PATH,
                                      MPL_TEMPLATE_PATH, RPY2_TEMPLATE_PATH,
                                      PLOT9_TEMPLATE_PATH)
except ImportError:
    from actions import ChartDialogActions
    from toolbar import ChartTemplatesToolBar, RChartTemplatesToolBar
    from widgets import HelpBrowser, TypeMenuComboBox
    from lib.csv import sniff, csv_reader, get_header, convert
    from lib.pycellsheet import ExpressionParser, Formatter
    from lib.spelltextedit import SpellTextEdit
    from model.model import INITSCRIPT_DEFAULT, CodeArray
    from settings import (TUTORIAL_PATH, MANUAL_PATH, MPL_TEMPLATE_PATH,
                          RPY2_TEMPLATE_PATH, PLOT9_TEMPLATE_PATH)


class DiscardChangesDialog:
    """Modal dialog that asks if the user wants to discard or save unsaved data

    The modal dialog is shown on accessing the property choice.

    """

    title = "Unsaved changes"
    text = "There are unsaved changes.\nDo you want to save?"
    buttons = QMessageBox.StandardButton
    choices = buttons.Discard | buttons.Cancel | buttons.Save
    default_choice = buttons.Save

    def __init__(self, main_window: QMainWindow):
        """
        :param main_window: Application main window

        """
        self.main_window = main_window

    @property
    def choice(self) -> bool:
        """User choice

        Returns True if the user confirms in a user dialog that unsaved
        changes will be discarded if conformed.
        Returns False if the user chooses to save the unsaved data
        Returns None if the user chooses to abort the operation

        """

        button_approval = QMessageBox.warning(self.main_window, self.title,
                                              self.text, self.choices,
                                              self.default_choice)
        if button_approval == QMessageBox.StandardButton.Discard:
            return True
        if button_approval == QMessageBox.StandardButton.Save:
            return False


class DiscardDataDialog(DiscardChangesDialog):
    """Modal dialog that asks if the user wants to discard data"""

    title = "Data to be discarded"
    buttons = QMessageBox.StandardButton
    choices = buttons.Discard | buttons.Cancel
    default_choice = buttons.Cancel

    def __init__(self, main_window: QMainWindow, text: str):
        """
        :param main_window: Application main window
        :param text: Message text

        """

        super().__init__(main_window)
        self.text = text


class SheetScriptDraftDialog(DiscardChangesDialog):
    """Modal dialog that asks how unapplied Sheet Script drafts are handled."""

    title = "Unapplied Sheet Script drafts"
    text = ("There are unapplied Sheet Script drafts.\n"
            "Do you want to apply them before continuing?")
    buttons = QMessageBox.StandardButton
    choices = buttons.Apply | buttons.Discard | buttons.Cancel
    default_choice = buttons.Apply

    @property
    def choice(self) -> str:
        """User choice

        Returns `"apply"` if the user chooses to apply drafts.
        Returns `"discard"` if the user chooses to discard drafts.
        Returns `None` if the user chooses to abort the operation.

        """

        button_approval = QMessageBox.warning(self.main_window, self.title,
                                              self.text, self.choices,
                                              self.default_choice)
        if button_approval == QMessageBox.StandardButton.Apply:
            return "apply"
        if button_approval == QMessageBox.StandardButton.Discard:
            return "discard"


class ApproveWarningDialog:
    """Modal warning dialog for approving files to be evaled

    The modal dialog is shown on accessing the property choice.

    """

    title = "Security warning"
    text = ("You are going to approve and trust a file that you have not "
            "created yourself. After proceeding, the file is executed.\n \n"
            "It may harm your system as any program can. Please check all "
            "cells thoroughly before proceeding.\n \n"
            "Proceed and sign this file as trusted?")
    buttons = QMessageBox.StandardButton
    choices = buttons.No | buttons.Yes
    default_choice = buttons.No

    def __init__(self, parent: QWidget):
        """
        :param parent: Parent widget, e.g. main window

        """

        self.parent = parent

    @property
    def choice(self) -> bool:
        """User choice

        Returns True iif the user approves leaving safe_mode.
        Returns False iif the user chooses to stay in safe_mode
        Returns None if the user chooses to abort the operation

        """

        button_approval = QMessageBox.warning(self.parent, self.title,
                                              self.text, self.choices,
                                              self.default_choice)
        if button_approval == self.buttons.Yes:
            return True
        if button_approval == self.buttons.No:
            return False


class DataEntryDialog(QDialog):
    """Modal dialog for entering multiple values"""

    def __init__(self, parent: QWidget, title: str, labels: Sequence[str],
                 initial_data: Sequence[str] = None,
                 groupbox_title: str = None,
                 validators: Sequence[Union[QValidator, bool, None]] = None):
        """
        :param parent: Parent widget, e.g. main window
        :param title: Dialog title
        :param labels: Labels for the values in the dialog
        :param initial_data: Initial values to be displayed in the dialog
        :param validators: Validators for the editors of the dialog

        len(initial_data), len(validators) and len(labels) must be equal

        """

        super().__init__(parent)

        self.labels = labels
        self.groupbox_title = groupbox_title

        if initial_data is None:
            self.initial_data = [""] * len(labels)
        elif len(initial_data) != len(labels):
            raise ValueError("Length of labels and initial_data not equal")
        else:
            self.initial_data = initial_data

        if validators is None:
            self.validators = [None] * len(labels)
        elif len(validators) != len(labels):
            raise ValueError("Length of labels and validators not equal")
        else:
            self.validators = validators

        self.editors = []

        layout = QVBoxLayout(self)
        layout.addWidget(self.create_form())
        layout.addStretch(1)
        layout.addWidget(self.create_buttonbox())
        self.setLayout(layout)

        self.setWindowTitle(title)
        self.setMinimumWidth(300)
        self.setMinimumHeight(150)

    @property
    def data(self) -> Tuple[str]:
        """Executes the dialog and returns input as a tuple of strings

        Returns None if the dialog is canceled.

        """

        result = self.exec()

        if result == QDialog.DialogCode.Accepted:
            data = []
            for editor in self.editors:
                if isinstance(editor, QLineEdit):
                    data.append(editor.text())
                elif isinstance(editor, QComboBox):
                    data.append(editor.currentText())
                else:
                    data.append(editor.isChecked())
            return tuple(data)

    def create_form(self) -> QGroupBox:
        """Returns form inside a QGroupBox"""

        form_group_box = QGroupBox()
        if self.groupbox_title:
            form_group_box.setTitle(self.groupbox_title)
        form_layout = QFormLayout()

        for label, initial_value, validator in zip(self.labels,
                                                   self.initial_data,
                                                   self.validators):
            if validator is bool:
                editor = QCheckBox("")
                editor.setChecked(initial_value)
            elif hasattr(validator, "acceptable_values"):
                editor = QComboBox()
                editor.addItems(validator.acceptable_values)
                editor.setCurrentText(initial_value)
            else:
                editor = QLineEdit(str(initial_value))
                editor.setAlignment(Qt.AlignmentFlag.AlignRight)
            if validator and validator is not bool:
                editor.setValidator(validator)
            form_layout.addRow(QLabel(label + " :"), editor)
            self.editors.append(editor)

        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form_group_box.setLayout(form_layout)

        return form_group_box

    def create_buttonbox(self) -> QDialogButtonBox:
        """Returns a QDialogButtonBox with Ok and Cancel"""

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                      | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        return button_box


class GridShapeDialog(DataEntryDialog):
    """Modal dialog for entering the number of rows, columns and tables"""

    def __init__(self, parent: QWidget, shape: Tuple[int, int, int],
                 title: str = "Create a new Grid"):
        """
        :param parent: Parent widget, e.g. main window
        :param shape: Initial shape to be displayed in the dialog

        """

        groupbox_title = "Grid Shape"
        labels = ["Number of Rows", "Number of Columns", "Number of Tables"]
        validator = QIntValidator()
        validator.setBottom(1)  # Do not allow negative values
        validators = [validator] * len(labels)
        super().__init__(parent, title, labels, shape, groupbox_title,
                         validators)

    @property
    def shape(self) -> Tuple[int, int, int]:
        """Executes the dialog and returns an rows, columns, tables

        Returns None if the dialog is canceled.

        """

        try:
            return tuple(map(int, self.data))
        except (TypeError, ValueError):
            pass


@dataclass
class SinglePageArea:
    """Holds single page area boundaries e.g. for export"""

    top: int
    left: int
    bottom: int
    right: int


@dataclass
class MultiPageArea(SinglePageArea):
    """Holds multi page area boundaries e.g. for printing"""

    first: int
    last: int


class CsvExportAreaDialog(DataEntryDialog):
    """Modal dialog for entering csv export area

    Initially, this dialog is filled with the selection bounding box
    if present or with the visible area of <= 1 cell is selected.

    """

    groupbox_title = "Page area"
    labels = ["Top", "Left", "Bottom", "Right"]
    area_cls = SinglePageArea

    def __init__(self, parent: QWidget, grid: QTableView, title: str):
        """
        :param parent: Parent widget, e.g. main window
        :param grid: The main grid widget
        :param title: Dialog title

        """

        self.grid = grid
        self.shape = grid.model.shape

        super().__init__(parent, title, self.labels, self._initial_values,
                         self.groupbox_title, self.validator_list)

    @property
    def _validator(self):
        """Returns int validator for positive numbers"""

        validator = QIntValidator()
        validator.setBottom(0)
        return validator

    @property
    def _row_validator(self) -> QIntValidator:
        """Returns row validator"""

        row_validator = self._validator
        row_validator.setTop(self.shape[0] - 1)
        return row_validator

    @property
    def _column_validator(self) -> QIntValidator:
        """Returns column validator"""

        column_validator = self._validator
        column_validator.setTop(self.shape[1] - 1)
        return column_validator

    @property
    def validator_list(self) -> List[QIntValidator]:
        """Returns list of validators for dialog"""

        return [self._row_validator, self._column_validator] * 2

    @property
    def _initial_values(self) -> Tuple[int, int, int, int]:
        """Returns tuple of initial values"""

        grid = self.grid
        shape = grid.model.shape

        if grid.selection and len(grid.selected_idx) > 1:
            (bb_top, bb_left), (bb_bottom, bb_right) = \
                grid.selection.get_grid_bbox(shape)
        else:
            bb_top, bb_bottom = grid.rowAt(0), grid.rowAt(grid.height())
            bb_left, bb_right = grid.columnAt(0), grid.columnAt(grid.width())

        return bb_top, bb_left, bb_bottom, bb_right

    @property
    def area(self) -> Union[SinglePageArea, MultiPageArea]:
        """Executes the dialog and returns top, left, bottom, right

        Returns None if the dialog is canceled.

        """

        try:
            int_data = map(int, self.data)
            data = (min(self.shape[i % 2], d) for i, d in enumerate(int_data))
        except (TypeError, ValueError):
            return

        if data is not None:
            try:
                return self.area_cls(*data)
            except ValueError:
                return


class SvgExportAreaDialog(CsvExportAreaDialog):
    """Modal dialog for entering svg export area

    Initially, this dialog is filled with the selection bounding box
    if present or with the visible area of <= 1 cell is selected.

    """

    groupbox_title = "SVG export area"


class PrintAreaDialog(CsvExportAreaDialog):
    """Modal dialog for entering print area

    Initially, this dialog is filled with the selection bounding box
    if present or with the visible area of <= 1 cell is selected.
    Initially, the current table is selected.

    """

    labels = ["Top", "Left", "Bottom", "Right", "First table", "Last table"]
    area_cls = MultiPageArea

    @property
    def _table_validator(self) -> QIntValidator:
        """Returns column validator"""

        table_validator = self._validator
        table_validator.setTop(self.shape[1] - 1)
        return table_validator

    @property
    def validator_list(self) -> List[QIntValidator]:
        """Returns list of validators for dialog"""

        validators = super().validator_list
        validators += [self._table_validator] * 2
        return validators

    @property
    def _initial_values(self) -> Tuple[int, int, int, int, int, int]:
        """Returns tuple of initial values"""

        bb_top, bb_left, bb_bottom, bb_right = super()._initial_values
        table = self.grid.table

        return bb_top, bb_left, bb_bottom, bb_right, table, table


class TupleValidator(QValidator):
    """Validator for a tuple of values, normally strings"""

    def __init__(self, *acceptable_values):
        super().__init__()
        self.acceptable_values = acceptable_values

    def validate(self, text, pos):
        """Validator core function. Must be overloaded."""

        if text in self.acceptable_values:
            return self.State.Acceptable
        return self.State.Invalid


class PreferencesDialog(DataEntryDialog):
    """Modal dialog for entering PyCellSheet preferences"""

    def __init__(self, parent: QWidget):
        """
        :param parent: Parent widget, e.g. main window

        """

        title = "Preferences"
        groupbox_title = "Global settings"
        labels = ["Signature key for files", "Number of recent files", "Show sum in statusbar"]
        self.keys = ["signature_key", "max_file_history", "show_statusbar_sum"]
        self.mappers = [str, int, bool]

        int_validator = QIntValidator()
        int_validator.setBottom(0)  # Do not allow negative values

        validators = [None, int_validator, bool]
        if moneyed is not None:
            currencies = map(str, moneyed.list_all_currencies())
            tuple_validator = TupleValidator(*currencies)
            labels.append("Money default currency")
            self.keys.append("currency_iso_code")
            self.mappers.append(str)
            validators.append(tuple_validator)

        data = [getattr(parent.settings, key) for key in self.keys]
        super().__init__(parent, title, labels, data, groupbox_title,
                         validators)

    @property
    def data(self) -> dict:
        """Executes the dialog and returns a dict containing preferences data

        Returns None if the dialog is canceled.

        """

        data = super().data
        if data is not None:
            data_dict = {}
            for key, mapper, data in zip(self.keys, self.mappers, data):
                data_dict[key] = mapper(data)
            return data_dict


class StartupGreeterDialog(QDialog):
    """Startup greeter for parser and initscript selections."""

    @staticmethod
    def simple_initscript_template() -> str:
        lines = []
        for line in INITSCRIPT_DEFAULT.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.startswith("#"):
                continue
            lines.append(line)
        return "\n".join(lines).strip() + "\n"

    def __init__(self, parent: QWidget, settings, current_parser_code: str):
        super().__init__(parent)

        self.settings = settings
        self.current_parser_code = current_parser_code
        self.action = "close"

        self.setWindowTitle("Welcome to PyCellSheet")
        self.setMinimumSize(860, 620)
        self._build_ui()
        self._load_defaults()

    def _build_ui(self):
        self.tabs = QTabWidget(self)
        self._build_parser_tab()
        self._build_initscript_tab()

        self.open_button = QPushButton("Open")
        self.new_button = QPushButton("New")
        self.close_button = QPushButton("Close")
        self.open_button.clicked.connect(lambda: self._done("open"))
        self.new_button.clicked.connect(lambda: self._done("new"))
        self.close_button.clicked.connect(lambda: self._done("close"))

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(self.open_button)
        button_row.addWidget(self.new_button)
        button_row.addWidget(self.close_button)

        layout = QVBoxLayout(self)
        layout.addWidget(self.tabs)
        layout.addLayout(button_row)
        self.setLayout(layout)

    def _build_parser_tab(self):
        parser_tab = QWidget(self)
        layout = QVBoxLayout(parser_tab)

        self.parser_mode_combo = QComboBox()
        for mode in ExpressionParser.list_modes():
            self.parser_mode_combo.addItem(
                f"{mode['label']} ({mode['id']})",
                ("official", mode["id"], mode["code"]),
            )
        self.parser_mode_combo.addItem("Custom (manual code)", ("custom", None, None))

        self.custom_preset_combo = QComboBox()
        self.save_preset_button = QPushButton("Save as custom preset")
        self.delete_preset_button = QPushButton("Delete preset")
        self.save_preset_button.clicked.connect(self._save_custom_preset)
        self.delete_preset_button.clicked.connect(self._delete_custom_preset)
        self.custom_preset_combo.currentIndexChanged.connect(
            self._on_custom_preset_selected
        )

        self.parser_code_editor = QPlainTextEdit()
        self.parser_code_editor.setPlaceholderText(
            "Parser code (function body). Official modes are read-only."
        )

        self.example_plaintext = QPlainTextEdit()
        self.example_plaintext.setPlaceholderText("Plaintext examples")
        self.example_python = QPlainTextEdit()
        self.example_python.setPlaceholderText("Python code examples")
        self.example_formula = QPlainTextEdit()
        self.example_formula.setPlaceholderText("Spreadsheet formula examples")

        form = QFormLayout()
        form.addRow("Parser mode:", self.parser_mode_combo)
        form.addRow("Custom presets:", self.custom_preset_combo)
        row = QHBoxLayout()
        row.addWidget(self.save_preset_button)
        row.addWidget(self.delete_preset_button)
        form.addRow("", row)
        form.addRow("Parser code:", self.parser_code_editor)

        examples = QHBoxLayout()
        examples.addWidget(self.example_plaintext)
        examples.addWidget(self.example_python)
        examples.addWidget(self.example_formula)

        layout.addLayout(form)
        layout.addWidget(QLabel("Try sample content snippets:"))
        layout.addLayout(examples)
        self.tabs.addTab(parser_tab, "Parser")

        self.parser_mode_combo.currentIndexChanged.connect(self._on_parser_mode_changed)

    def _build_initscript_tab(self):
        tab = QWidget(self)
        layout = QVBoxLayout(tab)

        self.initscript_preset_combo = QComboBox()
        self.initscript_preset_combo.addItem("Verbose", "verbose")
        self.initscript_preset_combo.addItem("Simple", "simple")
        self.initscript_preset_combo.addItem("Custom", "custom")
        self.initscript_preset_combo.currentIndexChanged.connect(
            self._on_initscript_preset_changed
        )

        self.initscript_editor = QPlainTextEdit()
        self.initscript_editor.setPlaceholderText("Sheet Script preset text")

        form = QFormLayout()
        form.addRow("Preset:", self.initscript_preset_combo)
        form.addRow("Template:", self.initscript_editor)

        layout.addLayout(form)
        self.tabs.addTab(tab, "Sheet Script Preset")

    def _load_defaults(self):
        self._reload_custom_presets()

        preferred_mode = self.settings.startup_parser_mode_id
        idx = -1
        for i in range(self.parser_mode_combo.count()):
            kind, mode_id, _code = self.parser_mode_combo.itemData(i)
            if kind == "official" and mode_id == preferred_mode:
                idx = i
                break
        if idx == -1:
            idx = 0
        self.parser_mode_combo.setCurrentIndex(idx)
        self._on_parser_mode_changed()

        self.example_plaintext.setPlainText("hello\n42\nTOTAL")
        self.example_python.setPlainText(">1 + 2\n>'x' * 2")
        self.example_formula.setPlainText("=SUM(A1:A3)\n=A1*10")

        preset_choice = self.settings.initscript_preset_choice
        preset_idx = self.initscript_preset_combo.findData(preset_choice)
        self.initscript_preset_combo.setCurrentIndex(max(0, preset_idx))
        self._on_initscript_preset_changed()

    def _reload_custom_presets(self):
        self.custom_preset_combo.clear()
        presets = self.settings.parser_custom_presets or []
        for preset in presets:
            if not isinstance(preset, dict):
                continue
            name = preset.get("name", "Unnamed")
            code = preset.get("code", "")
            self.custom_preset_combo.addItem(name, code)

    def _on_parser_mode_changed(self):
        kind, _mode_id, mode_code = self.parser_mode_combo.currentData()
        if kind == "official":
            self.parser_code_editor.setPlainText(mode_code)
            self.parser_code_editor.setReadOnly(True)
        else:
            self.parser_code_editor.setReadOnly(False)

    def _on_custom_preset_selected(self):
        if self.parser_mode_combo.currentData()[0] != "custom":
            return
        code = self.custom_preset_combo.currentData()
        if code:
            self.parser_code_editor.setPlainText(code)

    def _save_custom_preset(self):
        name, ok = QInputDialog.getText(self, "Preset name", "Name:")
        if not ok or not name.strip():
            return
        presets = list(self.settings.parser_custom_presets or [])
        presets.append({"name": name.strip(), "code": self.parser_code_editor.toPlainText()})
        self.settings.parser_custom_presets = presets
        self._reload_custom_presets()

    def _delete_custom_preset(self):
        idx = self.custom_preset_combo.currentIndex()
        if idx < 0:
            return
        target_name = self.custom_preset_combo.currentText()
        presets = [p for p in (self.settings.parser_custom_presets or [])
                   if not (isinstance(p, dict) and p.get("name") == target_name)]
        self.settings.parser_custom_presets = presets
        self._reload_custom_presets()

    def _on_initscript_preset_changed(self):
        choice = self.initscript_preset_combo.currentData()
        if choice == "verbose":
            self.initscript_editor.setPlainText(INITSCRIPT_DEFAULT)
            self.initscript_editor.setReadOnly(True)
        elif choice == "simple":
            self.initscript_editor.setPlainText(self.simple_initscript_template())
            self.initscript_editor.setReadOnly(True)
        else:
            self.initscript_editor.setReadOnly(False)
            text = self.settings.initscript_preset_custom or ""
            if text and self.initscript_editor.toPlainText() != text:
                self.initscript_editor.setPlainText(text)

    @property
    def parser_code(self) -> str:
        return self.parser_code_editor.toPlainText()

    @property
    def parser_mode_id(self) -> str:
        kind, mode_id, _code = self.parser_mode_combo.currentData()
        return mode_id if kind == "official" else "custom"

    @property
    def initscript_choice(self) -> str:
        return self.initscript_preset_combo.currentData()

    @property
    def initscript_template(self) -> str:
        return self.initscript_editor.toPlainText()

    def _done(self, action: str):
        self.settings.startup_parser_mode_id = (
            self.parser_mode_id if self.parser_mode_id != "custom" else "pure_spreadsheet"
        )
        self.settings.initscript_preset_choice = self.initscript_choice
        if self.initscript_choice == "custom":
            self.settings.initscript_preset_custom = self.initscript_template
        self.action = action
        if action == "close":
            self.reject()
        else:
            self.accept()


class ExpParserAdvancedDialog(QDialog):
    """Advanced Expression Parser configuration with 9-cell playground."""

    OFFICIAL_EXAMPLES = {
        "pure_pythonic": [
            ('"Python"', '"This is a\\n" + "Python string"'),
            ('"Spreadsheet (dud!)"', '=CONCAT("spread", "sheet")'),
            ('"String"', '"text"'),
        ],
        "mixed": [
            ('"Python"', '"This is a\\n" + "Python string"'),
            ('"Spreadsheet"', '=CONCAT("spread", "sheet")'),
            ('"String"', "'text"),
        ],
        "pure_spreadsheet": [
            ('Python', '> "This is a\\n" + "Python string"'),
            ('Spreadsheet', '=CONCAT("spread", "sheet")'),
            ('String', "'text"),
        ],
    }

    class PlaygroundModel(QAbstractTableModel):
        """Small 3x3 worksheet model backed by CodeArray."""

        def __init__(self, settings, parent=None):
            super().__init__(parent)
            self.code_array = CodeArray((3, 3, 1), settings)
            self.code_array.sheet_scripts[0] = INITSCRIPT_DEFAULT
            self.code_array.execute_sheet_script(0)

        def rowCount(self, parent=QModelIndex()):
            if parent.isValid():
                return 0
            return 3

        def columnCount(self, parent=QModelIndex()):
            if parent.isValid():
                return 0
            return 3

        def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
            if role != Qt.ItemDataRole.DisplayRole:
                return None
            if orientation == Qt.Orientation.Horizontal:
                return chr(ord("A") + section)
            return str(section + 1)

        def flags(self, index):
            if not index.isValid():
                return Qt.ItemFlag.NoItemFlags
            return (
                Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsEditable
            )

        def data(self, index, role=Qt.ItemDataRole.DisplayRole):
            if not index.isValid():
                return None
            key = (index.row(), index.column(), 0)

            if role == Qt.ItemDataRole.EditRole:
                code = self.code_array(key)
                return "" if code is None else str(code)

            if role == Qt.ItemDataRole.DisplayRole:
                try:
                    return str(Formatter.display_formatter(self.code_array[key]))
                except Exception as err:
                    return f"{type(err).__name__}: {err}"

            if role == Qt.ItemDataRole.ToolTipRole:
                code = self.code_array(key)
                try:
                    result = self.code_array[key]
                    rendered = Formatter.tooltip_formatter(result)
                except Exception as err:
                    rendered = f"{type(err).__name__}: {err}"
                return f"Code: {code or ''}\nResult: {rendered}"
            return None

        def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
            if not index.isValid() or role != Qt.ItemDataRole.EditRole:
                return False
            self.code_array[index.row(), index.column(), 0] = str(value)
            self._emit_all_changed()
            return True

        def set_parser_code(self, parser_code: str):
            self.code_array.exp_parser_code = parser_code
            self.code_array.smart_cache.clear()
            self.code_array.dep_graph.dirty.clear()
            self._emit_all_changed()

        def load_examples(self, examples):
            for row in range(3):
                for col in range(3):
                    self.code_array[row, col, 0] = ""
            for col, (label, code) in enumerate(examples):
                self.code_array[0, col, 0] = label
                self.code_array[1, col, 0] = code
            self._emit_all_changed()

        def _emit_all_changed(self):
            top_left = self.index(0, 0)
            bottom_right = self.index(2, 2)
            self.dataChanged.emit(top_left, bottom_right)

    def __init__(self, parent: QWidget, settings, mode_id: str, parser_code: str):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Expression Parser Advanced")
        self.setMinimumSize(760, 520)
        self._mode_change_in_progress = False

        self.mode_combo = QComboBox()
        for mode in ExpressionParser.list_modes():
            self.mode_combo.addItem(
                f"[Official] {mode['label']}",
                ("official", mode["id"], mode["code"]),
            )
        for preset in (self.settings.parser_custom_presets or []):
            if isinstance(preset, dict):
                self.mode_combo.addItem(
                    preset.get("name", "Unnamed"),
                    ("custom", preset.get("name", "Unnamed"), preset.get("code", "")),
                )

        idx = self.mode_combo.findData(
            ("official", "pure_spreadsheet",
             ExpressionParser.get_mode_code("pure_spreadsheet"))
        )
        if idx >= 0:
            self.mode_combo.setCurrentIndex(idx)
        else:
            self.mode_combo.setCurrentIndex(0)

        self.parser_code_editor = QPlainTextEdit()
        self.parser_code_editor.setPlainText(parser_code or "")

        self.playground_model = self.PlaygroundModel(settings, self)
        self.playground = QTableView(self)
        self.playground.setModel(self.playground_model)
        self.playground.setAlternatingRowColors(True)
        self.playground.horizontalHeader().setStretchLastSection(True)
        self.playground.verticalHeader().setDefaultSectionSize(28)
        self.playground.horizontalHeader().setDefaultSectionSize(220)
        self.playground_status = QLabel("")
        self.playground_status.setWordWrap(True)

        self.save_preset_button = QPushButton("Save preset")
        self.delete_preset_button = QPushButton("Delete preset")

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        form = QFormLayout()
        form.addRow("Mode:", self.mode_combo)
        form.addRow("Parser code:", self.parser_code_editor)

        preset_row = QHBoxLayout()
        preset_row.addWidget(self.save_preset_button)
        preset_row.addWidget(self.delete_preset_button)
        form.addRow("Presets:", preset_row)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(QLabel("Playground (real worksheet, 9 cells):"))
        layout.addWidget(self.playground)
        layout.addWidget(self.playground_status)
        layout.addWidget(buttons)

        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self.save_preset_button.clicked.connect(self._save_preset)
        self.delete_preset_button.clicked.connect(self._delete_preset)
        self.parser_code_editor.textChanged.connect(self._on_parser_code_edited)

        self._on_mode_changed()

    def _set_top6_examples(self, mode_id: str):
        examples = self.OFFICIAL_EXAMPLES.get(mode_id, [])
        self.playground_model.load_examples(examples)

    def _apply_parser_to_playground(self):
        parser_code = self.parser_code_editor.toPlainText()
        try:
            self.playground_model.set_parser_code(parser_code)
        except Exception as err:
            self.playground_status.setText(f"Parser code error: {err}")
            return False
        self.playground_status.setText("")
        return True

    def _on_parser_code_edited(self):
        if self._mode_change_in_progress:
            return
        self._apply_parser_to_playground()
        self._update_preset_buttons()

    def _on_mode_changed(self):
        self._mode_change_in_progress = True
        kind, mode_id, mode_code = self.mode_combo.currentData()
        if kind == "official" and mode_id in self.OFFICIAL_EXAMPLES:
            self.parser_code_editor.setPlainText(mode_code)
            self.parser_code_editor.setReadOnly(True)
            self._set_top6_examples(mode_id)
        else:
            self.parser_code_editor.setPlainText(mode_code)
            self.parser_code_editor.setReadOnly(False)
        self._mode_change_in_progress = False
        self._apply_parser_to_playground()
        self._update_preset_buttons()

    def _save_preset(self):
        if not self.save_preset_button.isEnabled():
            return
        name, ok = QInputDialog.getText(self, "Save parser preset", "Preset name:")
        if not ok or not name.strip():
            return
        trimmed = name.strip()
        presets = list(self.settings.parser_custom_presets or [])
        code = self.parser_code_editor.toPlainText()
        replaced = False
        for preset in presets:
            if isinstance(preset, dict) and preset.get("name") == trimmed:
                preset["code"] = code
                replaced = True
                break
        if not replaced:
            presets.append({"name": trimmed, "code": code})
        self.settings.parser_custom_presets = presets
        self._reload_mode_combo(select_custom_name=trimmed)

    def _delete_preset(self):
        kind, preset_name, _code = self.mode_combo.currentData()
        if kind != "custom":
            return
        presets = [p for p in (self.settings.parser_custom_presets or [])
                   if not (isinstance(p, dict) and p.get("name") == preset_name)]
        self.settings.parser_custom_presets = presets
        self._reload_mode_combo(select_official_id="pure_spreadsheet")

    def _reload_mode_combo(self, select_custom_name: str = None, select_official_id: str = None):
        current_data = self.mode_combo.currentData()
        self.mode_combo.blockSignals(True)
        self.mode_combo.clear()
        for mode in ExpressionParser.list_modes():
            self.mode_combo.addItem(
                f"[Official] {mode['label']}",
                ("official", mode["id"], mode["code"]),
            )
        for preset in (self.settings.parser_custom_presets or []):
            if isinstance(preset, dict):
                self.mode_combo.addItem(
                    preset.get("name", "Unnamed"),
                    ("custom", preset.get("name", "Unnamed"), preset.get("code", "")),
                )
        self.mode_combo.blockSignals(False)

        if select_custom_name is not None:
            idx = self.mode_combo.findData(("custom", select_custom_name, self._find_custom_code(select_custom_name)))
            if idx >= 0:
                self.mode_combo.setCurrentIndex(idx)
                return
        if select_official_id is not None:
            target = ("official", select_official_id, ExpressionParser.get_mode_code(select_official_id))
            idx = self.mode_combo.findData(target)
            if idx >= 0:
                self.mode_combo.setCurrentIndex(idx)
                return
        if current_data is not None:
            idx = self.mode_combo.findData(current_data)
            if idx >= 0:
                self.mode_combo.setCurrentIndex(idx)
                return
        self.mode_combo.setCurrentIndex(0)

    def _find_custom_code(self, name: str) -> str:
        for preset in (self.settings.parser_custom_presets or []):
            if isinstance(preset, dict) and preset.get("name") == name:
                return preset.get("code", "")
        return ""

    def _update_preset_buttons(self):
        current = self.mode_combo.currentData()
        if not current:
            self.save_preset_button.setEnabled(False)
            self.delete_preset_button.setEnabled(False)
            return
        kind, _id_or_name, code = current
        modified = self.parser_code_editor.toPlainText() != (code or "")
        self.save_preset_button.setEnabled(modified)
        self.delete_preset_button.setEnabled(kind == "custom")

    @property
    def selected_mode_id(self) -> str:
        kind, mode_id, _code = self.mode_combo.currentData()
        if kind == "official":
            return mode_id
        return "custom"

    @property
    def parser_code(self) -> str:
        return self.parser_code_editor.toPlainText()


class SheetScriptPresetAdvancedDialog(QDialog):
    """Advanced Sheet Script preset manager with save/delete."""

    def __init__(self, parent: QWidget, settings, choice: str, template: str):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Sheet Script Preset Advanced")
        self.setMinimumSize(760, 460)

        self.preset_combo = QComboBox()
        self.editor = QPlainTextEdit()
        self.save_button = QPushButton("Save preset")
        self.delete_button = QPushButton("Delete preset")

        self._load_presets()
        idx = self.preset_combo.findData(choice)
        if idx >= 0:
            self.preset_combo.setCurrentIndex(idx)
        else:
            self.preset_combo.setCurrentIndex(0)
        self.editor.setPlainText(template)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        row = QHBoxLayout()
        row.addWidget(self.preset_combo)
        row.addWidget(self.save_button)
        row.addWidget(self.delete_button)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Preset menu"))
        layout.addLayout(row)
        layout.addWidget(self.editor)
        layout.addWidget(buttons)

        self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
        self.save_button.clicked.connect(self._save_preset)
        self.delete_button.clicked.connect(self._delete_preset)
        self.editor.textChanged.connect(self._update_buttons)
        self._on_preset_changed()

    def _load_presets(self):
        self.preset_combo.clear()
        self.preset_combo.addItem("[Official] Verbose", "verbose")
        self.preset_combo.addItem("[Official] Simple", "simple")
        presets = self.settings.initscript_custom_presets or []
        for preset in presets:
            if isinstance(preset, dict):
                preset_id = f"preset:{preset.get('name', 'Unnamed')}"
                self.preset_combo.addItem(preset.get("name", "Unnamed"), preset_id)

    def _on_preset_changed(self):
        choice = self.preset_combo.currentData()
        if choice == "verbose":
            self.editor.setPlainText(INITSCRIPT_DEFAULT)
        elif choice == "simple":
            self.editor.setPlainText(StartupGreeterDialog.simple_initscript_template())
        elif str(choice).startswith("preset:"):
            target = str(choice).split(":", 1)[1]
            for preset in (self.settings.initscript_custom_presets or []):
                if isinstance(preset, dict) and preset.get("name") == target:
                    self.editor.setPlainText(preset.get("code", ""))
                    break
        self._update_buttons()

    def _save_preset(self):
        if not self.save_button.isEnabled():
            return
        name, ok = QInputDialog.getText(self, "Save Sheet Script preset", "Preset name:")
        if not ok or not name.strip():
            return
        trimmed = name.strip()
        presets = list(self.settings.initscript_custom_presets or [])
        code = self.editor.toPlainText()
        replaced = False
        for preset in presets:
            if isinstance(preset, dict) and preset.get("name") == trimmed:
                preset["code"] = code
                replaced = True
                break
        if not replaced:
            presets.append({"name": trimmed, "code": code})
        self.settings.initscript_custom_presets = presets
        self._load_presets()
        data = f"preset:{trimmed}"
        idx = self.preset_combo.findData(data)
        if idx >= 0:
            self.preset_combo.setCurrentIndex(idx)
        self._update_buttons()

    def _delete_preset(self):
        choice = self.preset_combo.currentData()
        if not str(choice).startswith("preset:"):
            return
        target = str(choice).split(":", 1)[1]
        presets = [p for p in (self.settings.initscript_custom_presets or [])
                   if not (isinstance(p, dict) and p.get("name") == target)]
        self.settings.initscript_custom_presets = presets
        self._load_presets()
        self.preset_combo.setCurrentIndex(0)
        self._update_buttons()

    def _current_template_code(self) -> str:
        choice = self.preset_combo.currentData()
        if choice == "verbose":
            return INITSCRIPT_DEFAULT
        if choice == "simple":
            return StartupGreeterDialog.simple_initscript_template()
        if str(choice).startswith("preset:"):
            target = str(choice).split(":", 1)[1]
            for preset in (self.settings.initscript_custom_presets or []):
                if isinstance(preset, dict) and preset.get("name") == target:
                    return preset.get("code", "")
        return ""

    def _update_buttons(self):
        baseline = self._current_template_code()
        modified = self.editor.toPlainText() != baseline
        self.save_button.setEnabled(modified)
        self.delete_button.setEnabled(str(self.preset_combo.currentData()).startswith("preset:"))

    @property
    def choice(self) -> str:
        data = self.preset_combo.currentData()
        if str(data).startswith("preset:"):
            return str(data)
        return data

    @property
    def template(self) -> str:
        return self.editor.toPlainText()


class NewDocumentDialog(QDialog):
    """New document setup dialog with parser/initscript and shape."""

    def __init__(self, parent: QWidget, settings, current_shape: Tuple[int, int, int]):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("New Document")
        self.setMinimumWidth(520)

        self.rows_spin = QSpinBox()
        self.cols_spin = QSpinBox()
        self.tabs_spin = QSpinBox()
        self.rows_spin.setRange(1, parent.settings.maxshape[0])
        self.cols_spin.setRange(1, parent.settings.maxshape[1])
        self.tabs_spin.setRange(1, parent.settings.maxshape[2])
        self.rows_spin.setValue(current_shape[0])
        self.cols_spin.setValue(current_shape[1])
        self.tabs_spin.setValue(current_shape[2])

        self.parser_mode_combo = QComboBox()
        for mode in ExpressionParser.list_modes():
            self.parser_mode_combo.addItem(
                f"[Official] {mode['label']}",
                ("official", mode["id"], mode["code"]),
            )
        for preset in (self.settings.parser_custom_presets or []):
            if isinstance(preset, dict):
                self.parser_mode_combo.addItem(
                    preset.get("name", "Unnamed"),
                    ("custom", preset.get("name", "Unnamed"), preset.get("code", "")),
                )
        self.parser_advanced_button = QPushButton("Advanced...")

        self.initscript_preset_combo = QComboBox()
        self.initscript_preset_combo.addItem("[Official] Verbose", "verbose")
        self.initscript_preset_combo.addItem("[Official] Simple", "simple")
        for preset in (self.settings.initscript_custom_presets or []):
            if isinstance(preset, dict):
                self.initscript_preset_combo.addItem(
                    preset.get("name", "Unnamed"),
                    f"preset:{preset.get('name', 'Unnamed')}",
                )
        self.initscript_advanced_button = QPushButton("Advanced...")

        self._parser_mode_id = self.settings.startup_parser_mode_id
        if self._parser_mode_id in ExpressionParser.MODE_ID_TO_LABEL:
            self._parser_code = ExpressionParser.get_mode_code(self._parser_mode_id)
        else:
            self._parser_mode_id = "pure_spreadsheet"
            self._parser_code = ExpressionParser.get_mode_code("pure_spreadsheet")
        self._initscript_choice = self.settings.initscript_preset_choice
        self._initscript_template = INITSCRIPT_DEFAULT

        pidx = self.parser_mode_combo.findData(
            ("official", self._parser_mode_id, self._parser_code)
        )
        if pidx >= 0:
            self.parser_mode_combo.setCurrentIndex(pidx)
        iidx = self.initscript_preset_combo.findData(self._initscript_choice)
        if iidx >= 0:
            self.initscript_preset_combo.setCurrentIndex(iidx)
        self._refresh_initscript_template()

        form = QFormLayout()
        form.addRow("Rows:", self.rows_spin)
        form.addRow("Columns:", self.cols_spin)
        form.addRow("Tables:", self.tabs_spin)
        parser_row = QHBoxLayout()
        parser_row.addWidget(self.parser_mode_combo)
        parser_row.addWidget(self.parser_advanced_button)
        form.addRow("Expression Parser:", parser_row)
        script_row = QHBoxLayout()
        script_row.addWidget(self.initscript_preset_combo)
        script_row.addWidget(self.initscript_advanced_button)
        form.addRow("Sheet Script Preset:", script_row)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

        self.parser_mode_combo.currentIndexChanged.connect(self._on_parser_mode_changed)
        self.initscript_preset_combo.currentIndexChanged.connect(self._on_initscript_choice_changed)
        self.parser_advanced_button.clicked.connect(self._open_parser_advanced)
        self.initscript_advanced_button.clicked.connect(self._open_initscript_advanced)

    def _reload_parser_preset_combo(self):
        self.parser_mode_combo.blockSignals(True)
        self.parser_mode_combo.clear()
        for mode in ExpressionParser.list_modes():
            self.parser_mode_combo.addItem(
                f"[Official] {mode['label']}",
                ("official", mode["id"], mode["code"]),
            )
        for preset in (self.settings.parser_custom_presets or []):
            if isinstance(preset, dict):
                self.parser_mode_combo.addItem(
                    preset.get("name", "Unnamed"),
                    ("custom", preset.get("name", "Unnamed"), preset.get("code", "")),
                )
        self.parser_mode_combo.blockSignals(False)

    def _reload_initscript_preset_combo(self):
        self.initscript_preset_combo.blockSignals(True)
        self.initscript_preset_combo.clear()
        self.initscript_preset_combo.addItem("[Official] Verbose", "verbose")
        self.initscript_preset_combo.addItem("[Official] Simple", "simple")
        for preset in (self.settings.initscript_custom_presets or []):
            if isinstance(preset, dict):
                self.initscript_preset_combo.addItem(
                    preset.get("name", "Unnamed"),
                    f"preset:{preset.get('name', 'Unnamed')}",
                )
        self.initscript_preset_combo.blockSignals(False)

    def _refresh_initscript_template(self):
        choice = self._initscript_choice
        if choice == "verbose":
            self._initscript_template = INITSCRIPT_DEFAULT
        elif choice == "simple":
            self._initscript_template = StartupGreeterDialog.simple_initscript_template()
        elif str(choice).startswith("preset:"):
            target = str(choice).split(":", 1)[1]
            for preset in (self.settings.initscript_custom_presets or []):
                if isinstance(preset, dict) and preset.get("name") == target:
                    self._initscript_template = preset.get("code", "")
                    break

    def _on_parser_mode_changed(self):
        kind, mode_or_name, code = self.parser_mode_combo.currentData()
        if kind == "official":
            self._parser_mode_id = mode_or_name
            self._parser_code = code
        else:
            self._parser_mode_id = "custom"
            self._parser_code = code

    def _on_initscript_choice_changed(self):
        self._initscript_choice = self.initscript_preset_combo.currentData()
        self._refresh_initscript_template()

    def _open_parser_advanced(self):
        dialog = ExpParserAdvancedDialog(
            self, self.settings, self._parser_mode_id, self._parser_code
        )
        if dialog.exec():
            self._parser_mode_id = dialog.selected_mode_id
            self._parser_code = dialog.parser_code
            self._reload_parser_preset_combo()
            if self._parser_mode_id in ExpressionParser.MODE_ID_TO_LABEL:
                idx = self.parser_mode_combo.findData(
                    ("official", self._parser_mode_id, self._parser_code)
                )
            else:
                idx = self.parser_mode_combo.findData(
                    ("custom", dialog.mode_combo.currentData()[1], self._parser_code)
                )
            if idx >= 0:
                self.parser_mode_combo.setCurrentIndex(idx)

    def _open_initscript_advanced(self):
        dialog = SheetScriptPresetAdvancedDialog(
            self, self.settings, self._initscript_choice, self._initscript_template
        )
        if dialog.exec():
            self._initscript_choice = dialog.choice
            self._initscript_template = dialog.template
            self._reload_initscript_preset_combo()
            idx = self.initscript_preset_combo.findData(self._initscript_choice)
            if idx >= 0:
                self.initscript_preset_combo.setCurrentIndex(idx)

    @property
    def shape(self) -> Tuple[int, int, int]:
        return self.rows_spin.value(), self.cols_spin.value(), self.tabs_spin.value()

    @property
    def parser_mode_id(self) -> str:
        return self._parser_mode_id

    @property
    def parser_code(self) -> str:
        return self._parser_code

    @property
    def initscript_choice(self) -> str:
        return self._initscript_choice

    @property
    def initscript_template(self) -> str:
        return self._initscript_template


class ExpressionParserSelectionDialog(QDialog):
    """Dialog for selecting parser mode."""

    def __init__(self, parent: QWidget, current_mode_id: str):
        super().__init__(parent)

        self._mode_labels = dict(ExpressionParser.MODE_ID_TO_LABEL)
        self.current_mode_id = current_mode_id

        self.setWindowTitle("Expression Parser Settings")
        self.setMinimumWidth(480)

        self.mode_combo = QComboBox()
        for mode_id, label in self._mode_labels.items():
            self.mode_combo.addItem(f"{label} ({mode_id})", mode_id)
        if current_mode_id in self._mode_labels:
            self.mode_combo.setCurrentIndex(self.mode_combo.findData(current_mode_id))

        self.custom_mode_notice = QLabel(
            "Current workbook uses a custom parser code. "
            "Selecting a built-in mode will replace it."
        )
        self.custom_mode_notice.setWordWrap(True)
        self.custom_mode_notice.setVisible(current_mode_id is None)

        form = QFormLayout()
        form.addRow("Expression Parser mode:", self.mode_combo)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.custom_mode_notice)
        layout.addLayout(form)
        layout.addWidget(buttons)
        self.setLayout(layout)

    @property
    def selection(self) -> str:
        return self.mode_combo.currentData()


class ExpressionParserMigrationDialog(QDialog):
    """Preview and apply conservative Expression Parser mode migration."""

    def __init__(self, parent: QWidget, code_array, current_mode_id: str,
                 current_table: int):
        super().__init__(parent)

        self.code_array = code_array
        self.current_mode_id = current_mode_id
        self.current_table = current_table
        self.applied_report = None
        self._last_preview = None

        self._mode_labels = dict(ExpressionParser.MODE_ID_TO_LABEL)
        self._mode_labels.update(ExpressionParser.LEGACY_MODE_ID_TO_LABEL)
        self._known_modes = list(self._mode_labels.items())

        self._build_ui()
        self._load_initial_values()

    def _build_ui(self):
        self.setWindowTitle("Migrate Expression Parser Mode")
        self.setMinimumWidth(560)

        if self.current_mode_id is None:
            current_parser_text = "Current workspace parser: Custom (non-official code)"
        else:
            current_label = ExpressionParser.MODE_ID_TO_LABEL.get(
                self.current_mode_id,
                ExpressionParser.LEGACY_MODE_ID_TO_LABEL.get(self.current_mode_id, self.current_mode_id),
            )
            current_parser_text = (
                f"Current workspace parser: {current_label} ({self.current_mode_id})"
            )
        self.current_parser_label = QLabel(current_parser_text)
        self.current_parser_label.setWordWrap(True)

        self.source_mode = QComboBox()
        self.target_mode = QComboBox()
        for mode_id, label in self._known_modes:
            display = f"{label} ({mode_id})"
            self.source_mode.addItem(display, mode_id)
            self.target_mode.addItem(display, mode_id)

        self.scope = QComboBox()
        self.scope.addItem("All sheets", None)
        self.scope.addItem("Current sheet only", [self.current_table])

        self.include_risky = QCheckBox("Include risky rewrites")
        self.include_risky.setChecked(False)

        self.summary = QPlainTextEdit()
        self.summary.setReadOnly(True)
        self.summary.setPlaceholderText(
            "Preview migration to see safe/risky/unchanged counts."
        )

        self.preview_button = QPushButton("Preview")
        self.apply_button = QPushButton("Apply")
        self.cancel_button = QPushButton("Cancel")
        self.apply_button.setEnabled(False)

        form = QFormLayout()
        form.addRow(QLabel("Source mode:"), self.source_mode)
        form.addRow(QLabel("Target mode:"), self.target_mode)
        form.addRow(QLabel("Scope:"), self.scope)
        form.addRow(QLabel(""), self.include_risky)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(self.preview_button)
        buttons.addWidget(self.apply_button)
        buttons.addWidget(self.cancel_button)

        layout = QVBoxLayout(self)
        layout.addWidget(self.current_parser_label)
        layout.addLayout(form)
        layout.addWidget(self.summary)
        layout.addLayout(buttons)
        self.setLayout(layout)

        self.preview_button.clicked.connect(self.on_preview)
        self.apply_button.clicked.connect(self.on_apply)
        self.cancel_button.clicked.connect(self.reject)

        self.source_mode.currentIndexChanged.connect(self._on_options_changed)
        self.target_mode.currentIndexChanged.connect(self._on_options_changed)
        self.scope.currentIndexChanged.connect(self._on_options_changed)
        self.include_risky.toggled.connect(self._on_options_changed)

    def _load_initial_values(self):
        if self.current_mode_id is not None:
            source_index = self.source_mode.findData(self.current_mode_id)
            if source_index != -1:
                self.source_mode.setCurrentIndex(source_index)
        else:
            self.summary.setPlainText(
                "Current workspace parser is custom. Select the source mode that best matches it."
            )

        target_mode_id = "pure_spreadsheet"
        if self.current_mode_id == target_mode_id:
            target_mode_id = "mixed"
        target_index = self.target_mode.findData(target_mode_id)
        if target_index != -1:
            self.target_mode.setCurrentIndex(target_index)

        self._validate_mode_pair()

    def _on_options_changed(self):
        self._last_preview = None
        self.applied_report = None
        self.apply_button.setEnabled(False)
        self._validate_mode_pair()

    def _migration_args(self):
        source_mode_id = self.source_mode.currentData()
        target_mode_id = self.target_mode.currentData()
        tables = self.scope.currentData()
        include_risky = self.include_risky.isChecked()
        return source_mode_id, target_mode_id, tables, include_risky

    def _validate_mode_pair(self):
        source_mode_id, target_mode_id, _, _ = self._migration_args()
        if source_mode_id == target_mode_id:
            self.summary.setPlainText(
                "Source and target modes must be different."
            )
            return False
        return True

    @staticmethod
    def _format_summary(report):
        summary = report.summary
        lines = [
            "Migration preview",
            f"- Source: {report.source_mode_id}",
            f"- Target: {report.target_mode_id}",
            f"- Safe changes: {summary['safe_changed']}",
            f"- Risky skipped: {summary['risky_skipped']}",
            f"- Unchanged: {summary['unchanged']}",
            f"- Invalid source assumptions: {summary['invalid_source_assumption']}",
            f"- Total examined: {summary['total']}",
        ]
        return "\n".join(lines)

    def on_preview(self):
        if not self._validate_mode_pair():
            self.apply_button.setEnabled(False)
            return

        source_mode_id, target_mode_id, tables, include_risky = self._migration_args()
        report = self.code_array.preview_expression_parser_migration(
            source_mode_id=source_mode_id,
            target_mode_id=target_mode_id,
            tables=tables,
            include_risky=include_risky,
        )
        self._last_preview = (source_mode_id, target_mode_id, tables, include_risky)
        self.summary.setPlainText(self._format_summary(report))
        self.apply_button.setEnabled(True)

    def on_apply(self):
        if not self._validate_mode_pair():
            return

        source_mode_id, target_mode_id, tables, include_risky = self._migration_args()
        args = (source_mode_id, target_mode_id, tables, include_risky)
        if self._last_preview != args:
            self.on_preview()
            if self._last_preview != args:
                return

        report = self.code_array.apply_expression_parser_migration(
            source_mode_id=source_mode_id,
            target_mode_id=target_mode_id,
            tables=tables,
            include_risky=include_risky,
        )
        self.applied_report = report
        self.summary.setPlainText(self._format_summary(report))
        self.accept()


class CellKeyDialog(DataEntryDialog):
    """Modal dialog for entering a cell key, i.e. row, column, table"""

    def __init__(self, parent: QWidget, shape: Tuple[int, int, int]):
        """
        :param parent: Parent widget, e.g. main window
        :param shape: Grid shape

        """

        title = "Go to cell"
        groupbox_title = "Cell index"
        labels = ["Row", "Column", "Table"]

        row_validator = QIntValidator()
        row_validator.setRange(0, shape[0] - 1)
        column_validator = QIntValidator()
        column_validator.setRange(0, shape[1] - 1)
        table_validator = QIntValidator()
        table_validator.setRange(0, shape[2] - 1)
        validators = [row_validator, column_validator, table_validator]

        super().__init__(parent, title, labels, None, groupbox_title,
                         validators)

    @property
    def key(self) -> Tuple[int, int, int]:
        """Executes the dialog and returns rows, columns, tables

        Returns None if the dialog is canceled.

        """

        data = self.data

        if data is None:
            return

        try:
            return tuple(map(int, data))
        except ValueError:
            pass


class FileDialogBase:
    """Base class for modal file dialogs

    The chosen filename is stored in the file_path attribute
    The chosen name filter is stored in the chosen_filter attribute
    If the dialog is aborted then both filepath and chosen_filter are None

    _get_filepath must be overloaded

    """

    file_path = None

    title = "Choose file"

    filters_list = [
        "PyCellSheet un-compressed (*.pycsu)",
        "PyCellSheet compressed (*.pycs)",
    ]
    if openpyxl is not None:
        filters_list.append("Office Open XML - Tabellendokument (*.xlsx)")

    selected_filter = None

    @property
    def filters(self) -> str:
        """Formatted filters for qt"""

        return ";;".join(self.filters_list)

    @property
    def suffix(self) -> str:
        """Suffix for filepath"""

        return self.selected_filter.split("(")[-1].strip()[1:-1]

    def __init__(self, main_window: QMainWindow):
        """
        :param main_window: Application main window

        """

        self.main_window = main_window
        self.selected_filter = self.filters_list[0]

        self.show_dialog()

    def show_dialog(self):
        """Sublasses must overload this method"""

        raise NotImplementedError


class FileOpenDialog(FileDialogBase):
    """Modal dialog for opening a PyCellSheet file"""

    title = "Open"

    def show_dialog(self):
        """Present dialog and update values"""

        path = self.main_window.settings.last_file_input_path
        self.file_path, self.selected_filter = \
            QFileDialog.getOpenFileName(self.main_window, self.title,
                                        str(path), self.filters,
                                        self.selected_filter)


class FileSaveDialog(FileDialogBase):
    """Modal dialog for saving a PyCellSheet file"""

    title = "Save"

    def show_dialog(self):
        """Present dialog and update values"""

        path = self.main_window.settings.last_file_output_path
        self.file_path, self.selected_filter = \
            QFileDialog.getSaveFileName(self.main_window, self.title,
                                        str(path), self.filters,
                                        self.selected_filter)


class ImageFileOpenDialog(FileDialogBase):
    """Modal dialog for inserting an image"""

    title = "Insert image"

    img_formats = QImageWriter.supportedImageFormats()
    img_format_strings = (f"*.{fmt.data().decode()}" for fmt in img_formats)
    img_format_string = " ".join(img_format_strings)
    name_filter = f"Images ({img_format_string})" + ";;" \
                  "Scalable Vector Graphics (*.svg *.svgz)"

    def show_dialog(self):
        """Present dialog and update values"""

        path = self.main_window.settings.last_file_input_path
        self.file_path, self.selected_filter = \
            QFileDialog.getOpenFileName(self.main_window,
                                        self.title,
                                        str(path),
                                        self.name_filter)


class CsvFileImportDialog(FileDialogBase):
    """Modal dialog for importing csv files"""

    title = "Import data"
    filters_list = [
        "CSV file (*.*)",
    ]

    @property
    def suffix(self):
        """Do not offer suffix for filepath"""

        return

    def show_dialog(self):
        """Present dialog and update values"""

        path = self.main_window.settings.last_file_import_path
        self.file_path, self.selected_filter = \
            QFileDialog.getOpenFileName(self.main_window, self.title,
                                        str(path), self.filters,
                                        self.filters_list[0])


class FileExportDialog(FileDialogBase):
    """Modal dialog for exporting csv files"""

    title = "Export data"

    def __init__(self, main_window: QMainWindow, filters_list: List[str]):
        """
        :param main_window: Application main window
        :param filters_list: List of filter strings

        """

        self.filters_list = filters_list
        super().__init__(main_window)

    @property
    def suffix(self) -> str:
        """Suffix for filepath"""

        return f".{self.selected_filter.split()[0].lower()}"

    def show_dialog(self):
        """Present dialog and update values"""

        path = self.main_window.settings.last_file_export_path
        self.file_path, self.selected_filter = \
            QFileDialog.getSaveFileName(self.main_window, self.title,
                                        str(path), self.filters,
                                        self.filters_list[0])


@dataclass
class FindDialogState:
    """Dataclass for FindDialog state storage"""

    pos: QPoint
    case: bool
    results: bool
    more: bool
    backward: bool
    word: bool
    regex: bool
    start: bool


class FindDialog(QDialog):
    """Find dialog that is launched from the main menu"""

    def __init__(self, main_window: QMainWindow):
        """
        :param main_window: Application main window

        """

        super().__init__(main_window)

        self.main_window = main_window
        workflows = main_window.workflows

        self._create_widgets()
        self._layout()
        self._order()

        self.setWindowTitle("Find")

        self.extension.hide()

        self.more_button.toggled.connect(self.extension.setVisible)

        self.find_button.clicked.connect(partial(workflows.find_dialog_on_find,
                                                 self))

        # Restore state
        state = self.main_window.settings.find_dialog_state
        if state is not None:
            self.restore(state)

    def _create_widgets(self):
        """Create find dialog widgets

        :param results_checkbox: Show find results checkbox

        """

        self.search_text_label = QLabel("Search for:")
        self.search_text_editor = QLineEdit()
        self.search_text_label.setBuddy(self.search_text_editor)

        self.case_checkbox = QCheckBox("Match &case")
        self.results_checkbox = QCheckBox("Code and &results")

        self.find_button = QPushButton("&Find")
        self.find_button.setDefault(True)

        self.more_button = QPushButton("&More")
        self.more_button.setCheckable(True)
        self.more_button.setAutoDefault(False)

        self.extension = QWidget()

        self.backward_checkbox = QCheckBox("&Backward")
        self.word_checkbox = QCheckBox("&Whole words")
        self.regex_checkbox = QCheckBox("Regular e&xpression")
        self.from_start_checkbox = QCheckBox("From &start")

        self.button_box = QDialogButtonBox(Qt.Orientation.Vertical)
        self.button_box.addButton(self.find_button,
                                  QDialogButtonBox.ButtonRole.ActionRole)
        self.button_box.addButton(self.more_button,
                                  QDialogButtonBox.ButtonRole.ActionRole)

    def _layout(self):
        """Find dialog layout"""

        self.extension_layout = QVBoxLayout()
        self.extension_layout.setContentsMargins(0, 0, 0, 0)
        self.extension_layout.addWidget(self.backward_checkbox)
        self.extension_layout.addWidget(self.word_checkbox)
        self.extension_layout.addWidget(self.regex_checkbox)
        self.extension_layout.addWidget(self.from_start_checkbox)
        self.extension.setLayout(self.extension_layout)

        self.text_layout = QGridLayout()
        self.text_layout.addWidget(self.search_text_label, 0, 0)
        self.text_layout.addWidget(self.search_text_editor, 0, 1)
        self.text_layout.setColumnStretch(0, 1)

        self.search_layout = QVBoxLayout()
        self.search_layout.addLayout(self.text_layout)
        self.search_layout.addWidget(self.case_checkbox)
        self.search_layout.addWidget(self.results_checkbox)

        self.main_layout = QGridLayout()
        self.main_layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)
        self.main_layout.addLayout(self.search_layout, 0, 0)
        self.main_layout.addWidget(self.button_box, 0, 1)
        self.main_layout.addWidget(self.extension, 1, 0, 1, 2)
        self.main_layout.setRowStretch(2, 1)

        self.setLayout(self.main_layout)

    def _order(self):
        """Find dialog tabOrder"""

        self.setTabOrder(self.results_checkbox, self.backward_checkbox)
        self.setTabOrder(self.backward_checkbox, self.word_checkbox)
        self.setTabOrder(self.word_checkbox, self.regex_checkbox)
        self.setTabOrder(self.regex_checkbox, self.from_start_checkbox)

    def restore(self, state):
        """Restores state from FindDialogState"""

        self.move(state.pos)
        self.case_checkbox.setChecked(state.case)
        self.results_checkbox.setChecked(state.results)
        self.more_button.setChecked(state.more)
        self.backward_checkbox.setChecked(state.backward)
        self.word_checkbox.setChecked(state.word)
        self.regex_checkbox.setChecked(state.regex)
        self.from_start_checkbox.setChecked(state.start)

    # Overrides

    def closeEvent(self, event: QEvent):
        """Store state for next invocation and close

        :param event: Close event

        """

        state = FindDialogState(pos=self.pos(),
                                case=self.case_checkbox.isChecked(),
                                results=self.results_checkbox.isChecked(),
                                more=self.more_button.isChecked(),
                                backward=self.backward_checkbox.isChecked(),
                                word=self.word_checkbox.isChecked(),
                                regex=self.regex_checkbox.isChecked(),
                                start=self.from_start_checkbox.isChecked())
        self.main_window.settings.find_dialog_state = state

        super().closeEvent(event)


class ReplaceDialog(FindDialog):
    """Replace dialog that is launched from the main menu"""

    def __init__(self, main_window: QMainWindow):
        """
        :param main_window: Application main window

        """

        super().__init__(main_window)

        workflows = main_window.workflows

        self.setWindowTitle("Replace")

        self.results_checkbox.setDisabled(True)

        self.replace_text_label = QLabel("Replace with:")
        self.replace_text_editor = QLineEdit()
        self.replace_text_label.setBuddy(self.replace_text_editor)

        self.text_layout.addWidget(self.replace_text_label, 1, 0)
        self.text_layout.addWidget(self.replace_text_editor, 1, 1)

        self.replace_button = QPushButton("&Replace")
        self.replace_all_button = QPushButton("Replace &all")

        self.button_box.addButton(self.replace_button,
                                  QDialogButtonBox.ButtonRole.ActionRole)
        self.button_box.addButton(self.replace_all_button,
                                  QDialogButtonBox.ButtonRole.ActionRole)

        self.setTabOrder(self.search_text_editor, self.replace_text_editor)
        self.setTabOrder(self.more_button, self.replace_button)

        p_onreplace = partial(workflows.replace_dialog_on_replace, self)
        self.replace_button.clicked.connect(p_onreplace)

        p_onreplaceall = partial(workflows.replace_dialog_on_replace_all, self)
        self.replace_all_button.clicked.connect(p_onreplaceall)


class ChartDialog(QDialog):
    """The chart dialog"""

    def __init__(self, parent: QWidget, key: Tuple[int, int, int],
                 size: Tuple[int, int] = (1000, 700)):
        """
        :param parent: Parent window
        :param key: Target cell for chart
        :param size: Initial dialog size

        """

        self.key = key

        if Figure is None:
            raise ImportError

        super().__init__(parent)

        self.actions = ChartDialogActions(self)

        self.chart_templates_toolbar = ChartTemplatesToolBar(self)
        self.rchart_templates_toolbar = RChartTemplatesToolBar(self)

        self.setWindowTitle(f"Chart dialog for cell {key}")

        self.resize(*size)
        self.parent = parent

        self.actions = ChartDialogActions(self)

        self.dialog_ui()

    def on_template(self):
        """Event handler for pressing a template toolbar button"""

        chart_template_name = self.sender().data()
        chart_template_code = None

        tpl_paths = MPL_TEMPLATE_PATH, RPY2_TEMPLATE_PATH, PLOT9_TEMPLATE_PATH
        for tpl_path in tpl_paths:
            full_tpl_path = tpl_path / chart_template_name
            try:
                with open(full_tpl_path, encoding='utf8') as template_file:
                    chart_template_code = template_file.read()
            except OSError:
                pass

        if chart_template_code is None:
            return

        self.editor.insertPlainText(chart_template_code)

    def dialog_ui(self):
        """Sets up dialog UI"""

        msg = "Enter Python code into the editor to the left. Globals " + \
              "such as X, Y, Z, S are available as they are in the grid. " + \
              "The last line must result in a matplotlib figure.\n \n" + \
              "Pressing Apply displays the figure or an error message in " + \
              "the right area."

        self.message = QTextBrowser(self)
        self.message.setText(msg)
        self.editor = SpellTextEdit(self)
        self.splitter = QSplitter(self)

        buttonbox = self.create_buttonbox()

        self.splitter.addWidget(self.editor)
        self.splitter.addWidget(self.message)
        self.splitter.setOpaqueResize(False)
        self.splitter.setSizes([9999, 9999])

        # Layout
        layout = QVBoxLayout(self)

        toolbar_layout = QHBoxLayout()
        toolbar_layout.addWidget(self.chart_templates_toolbar)
        toolbar_layout.addWidget(self.rchart_templates_toolbar)
        layout.addLayout(toolbar_layout)

        layout.addWidget(self.splitter)
        layout.addWidget(buttonbox)

        self.setLayout(layout)

    def apply(self):
        """Executes the code in the dialog and updates the canvas"""

        # Get current cell
        key = self.parent.grid.current
        code = self.editor.toPlainText()

        filelike = io.StringIO()
        with redirect_stdout(filelike):
            figure = self.parent.grid.model.code_array._eval_cell(key, code)
        stdout_str = filelike.getvalue()
        if stdout_str:
            stdout_str += "\n \n"

        if isinstance(figure, Figure):
            canvas = FigureCanvasQTAgg(figure)
            self.splitter.replaceWidget(1, canvas)
            try:
                canvas.draw()
            except Exception:
                pass
        elif isinstance(figure, bytes) or isinstance(figure, str):
            with redirect_stdout(filelike):
                if isinstance(figure, str):
                    figure = bytearray(figure, encoding='utf-8')
                svg_widget = QSvgWidget()
                self.splitter.replaceWidget(1, svg_widget)
                svg_widget.renderer().load(figure)
            stdout_str = filelike.getvalue()
            if stdout_str:
                stdout_str += "\n \n"
                msg = stdout_str + f"Error:\n{figure}"
                self.message.setText(msg)
        else:
            if isinstance(figure, Exception):
                msg = stdout_str + f"Error:\n{figure}"
                self.message.setText(msg)
            else:
                msg = stdout_str
                msg_text = "Error:\n{} has type '{}', " + \
                           "which is no instance of {}."
                msg += msg_text.format(figure, type(figure).__name__,
                                       Figure)
                self.message.setText(msg)

            if self.splitter.widget(1) != self.message:
                self.splitter.replaceWidget(1, self.message)

    def create_buttonbox(self):
        """Returns a QDialogButtonBox with Ok and Cancel"""

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                      | QDialogButtonBox.StandardButton.Apply
                                      | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(
            QDialogButtonBox.StandardButton.Apply).clicked.connect(self.apply)
        return button_box


class CsvParameterGroupBox(QGroupBox):
    """QGroupBox that holds parameter widgets for the csv import dialog"""

    title = "Parameters"

    quotings = "QUOTE_ALL", "QUOTE_MINIMAL", "QUOTE_NONNUMERIC", "QUOTE_NONE"

    # Tooltips
    encoding_widget_tooltip = "CSV file encoding"
    quoting_widget_tooltip = \
        "Controls when quotes should be generated by the writer and " \
        "recognised by the reader."
    quotechar_tooltip = \
        "A one-character string used to quote fields containing special " \
        "characters, such as the delimiter or quotechar, or which contain " \
        "new-line characters."
    delimiter_tooltip = "A one-character string used to separate fields."
    escapechar_tooltip = "A one-character string used by the writer to " \
        "escape the delimiter if quoting is set to QUOTE_NONE and the " \
        "quotechar if doublequote is False. On reading, the escapechar " \
        "removes any special meaning from the following character."
    hasheader_tooltip = \
        "Analyze the CSV file and treat the first row as strings if it " \
        "appears to be a series of column headers."
    keepheader_tooltip = "Import header labels as str in the first row"
    doublequote_tooltip = \
        "Controls how instances of quotechar appearing inside a field " \
        "should be themselves be quoted. When True, the character is " \
        "doubled. When False, the escapechar is used as a prefix to the " \
        "quotechar."
    skipinitialspace_tooltip = "When True, whitespace immediately following " \
        "the delimiter is ignored."

    # Default values that are displayed if the sniffer fails
    default_quoting = "QUOTE_MINIMAL"
    default_quotechar = '"'
    default_delimiter = ','

    def __init__(self, parent: QWidget):
        """
        :param parent: Parent window

        """

        super().__init__(parent)
        self.parent = parent
        self.default_encoding = parent.parent.settings.default_encoding
        self.encodings = parent.parent.settings.encodings

        self.setTitle(self.title)
        self._create_widgets()
        self._layout()

        self.hasheader_widget.toggled.connect(self.on_hasheader_toggled)

    def _create_widgets(self):
        """Create widgets for all parameters"""

        # Encoding
        self.encoding_label = QLabel("Encoding")
        self.encoding_widget = QComboBox(self.parent)
        self.encoding_widget.addItems(self.encodings)
        if self.default_encoding in self.encodings:
            default_encoding_idx = self.encodings.index(self.default_encoding)
            self.encoding_widget.setCurrentIndex(default_encoding_idx)
        self.encoding_widget.setEditable(False)
        self.encoding_widget.setToolTip(self.encoding_widget_tooltip)

        # Quote character
        self.quotechar_label = QLabel("Quote character")
        self.quotechar_widget = QLineEdit(self.default_quotechar, self.parent)
        self.quotechar_widget.setMaxLength(1)
        self.quotechar_widget.setToolTip(self.quotechar_tooltip)

        # Delimiter
        self.delimiter_label = QLabel("Delimiter")
        self.delimiter_widget = QLineEdit(self.default_delimiter, self.parent)
        self.delimiter_widget.setMaxLength(1)
        self.delimiter_widget.setToolTip(self.delimiter_tooltip)

        # Escape character
        self.escapechar_label = QLabel("Escape character")
        self.escapechar_widget = QLineEdit(self.parent)
        self.escapechar_widget.setMaxLength(1)
        self.escapechar_widget.setToolTip(self.escapechar_tooltip)

        # Quote style
        self.quoting_label = QLabel("Quote style")
        self.quoting_widget = QComboBox(self.parent)
        self.quoting_widget.addItems(self.quotings)
        if self.default_quoting in self.quotings:
            default_quoting_idx = self.quotings.index(self.default_quoting)
            self.quoting_widget.setCurrentIndex(default_quoting_idx)
        self.quoting_widget.setEditable(False)
        self.quoting_widget.setToolTip(self.quoting_widget_tooltip)

        # Header present
        self.hasheader_label = QLabel("Header present")
        self.hasheader_widget = QCheckBox(self.parent)
        self.hasheader_widget.setToolTip(self.hasheader_tooltip)

        # Keep header
        self.keepheader_label = QLabel("Keep header")
        self.keepheader_widget = QCheckBox(self.parent)
        self.keepheader_widget.setToolTip(self.keepheader_tooltip)

        # Double quote
        self.doublequote_label = QLabel("Doublequote")
        self.doublequote_widget = QCheckBox(self.parent)
        self.doublequote_widget.setToolTip(self.doublequote_tooltip)

        # Skip initial space
        self.skipinitialspace_label = QLabel("Skip initial space")
        self.skipinitialspace_widget = QCheckBox(self.parent)
        self.skipinitialspace_widget.setToolTip(self.skipinitialspace_tooltip)

        # Mapping to csv dialect

        self.csv_parameter2widget = {
            "encoding": self.encoding_widget,  # Extra dialect attribute
            "quotechar": self.quotechar_widget,
            "delimiter": self.delimiter_widget,
            "escapechar": self.escapechar_widget,
            "quoting": self.quoting_widget,
            "hasheader": self.hasheader_widget,  # Extra dialect attribute
            "keepheader": self.keepheader_widget,  # Extra dialect attribute
            "doublequote": self.doublequote_widget,
            "skipinitialspace": self.skipinitialspace_widget,
        }

    def _layout(self):
        """Layout widgets"""

        hbox_layout = QHBoxLayout()
        left_form_layout = QFormLayout()
        right_form_layout = QFormLayout()

        hbox_layout.addLayout(left_form_layout)
        hbox_layout.addSpacing(20)
        hbox_layout.addLayout(right_form_layout)

        left_form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        right_form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        left_form_layout.addRow(self.encoding_label, self.encoding_widget)
        left_form_layout.addRow(self.quotechar_label, self.quotechar_widget)
        left_form_layout.addRow(self.delimiter_label, self.delimiter_widget)
        left_form_layout.addRow(self.escapechar_label, self.escapechar_widget)

        right_form_layout.addRow(self.quoting_label, self.quoting_widget)
        right_form_layout.addRow(self.hasheader_label, self.hasheader_widget)
        right_form_layout.addRow(self.keepheader_label, self.keepheader_widget)
        right_form_layout.addRow(self.doublequote_label,
                                 self.doublequote_widget)
        right_form_layout.addRow(self.skipinitialspace_label,
                                 self.skipinitialspace_widget)

        self.setLayout(hbox_layout)

    # Event handlers
    def on_hasheader_toggled(self, toggled: bool):
        """Disables keep_header if hasheader is not toggled"""

        self.keepheader_widget.setChecked(False)
        self.keepheader_widget.setEnabled(toggled)

    def adjust_csvdialect(self, dialect: csv.Dialect) -> csv.Dialect:
        """Adjusts csv dialect from widget settings

        Note that the dialect has two extra attributes encoding and hasheader

        :param dialect: Attributes class for csv reading and writing

        """

        for parameter, widget in self.csv_parameter2widget.items():
            if hasattr(widget, "currentText"):
                value = widget.currentText()
            elif hasattr(widget, "isChecked"):
                value = widget.isChecked()
            elif hasattr(widget, "text"):
                value = widget.text()
            else:
                raise AttributeError(f"{widget} unsupported")

            # Convert strings to csv constants
            if parameter == "quoting" and isinstance(value, str):
                value = getattr(csv, value)

            setattr(dialect, parameter, value)

        return dialect

    def set_csvdialect(self, dialect: csv.Dialect):
        """Update widgets from given csv dialect

        :param dialect: Attributes class for csv reading and writing

        """

        for parameter in self.csv_parameter2widget:
            try:
                value = getattr(dialect, parameter)
            except AttributeError:
                value = None

            if value is not None:
                widget = self.csv_parameter2widget[parameter]
                if hasattr(widget, "setCurrentText"):
                    try:
                        widget.setCurrentText(value)
                    except TypeError:
                        try:
                            widget.setCurrentIndex(value)
                        except TypeError:
                            pass
                elif hasattr(widget, "setChecked"):
                    widget.setChecked(bool(value))
                elif hasattr(widget, "setText"):
                    widget.setText(value)
                else:
                    raise AttributeError(f"{widget} unsupported")
        if not self.hasheader_widget.isChecked():
            self.keepheader_widget.setEnabled(False)


class CsvTable(QTableView):
    """Table for previewing csv file content"""

    no_rows = 9

    def __init__(self, parent: QWidget):
        """
        :param parent: Parent window

        """

        super().__init__(parent)

        self.parent = parent

        self.comboboxes = []

        self.model = QStandardItemModel(self)
        self.setModel(self.model)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.verticalHeader().hide()

    def add_choice_row(self, length: int):
        """Adds row with comboboxes for digest choice

        :param length: Number of columns in row

        """

        item_row = map(QStandardItem, [''] * length)
        self.comboboxes = [TypeMenuComboBox() for _ in range(length)]
        self.model.appendRow(item_row)
        for i, combobox in enumerate(self.comboboxes):
            self.setIndexWidget(self.model.index(0, i), combobox)

    def fill(self, filepath: Path, dialect: csv.Dialect,
             digest_types: List[str] = None):
        """Fills the csv table with values from the csv file

        :param filepath: Path to csv file
        :param dialect: Attributes class for csv reading and writing
        :param digest_types: Names of preprocessing functions for csv values

        """

        self.model.clear()

        self.verticalHeader().hide()

        try:
            if hasattr(dialect, "encoding"):
                encoding = dialect.encoding
            else:
                encoding = self.parent.csv_encoding
            try:
                with open(filepath, newline='', encoding=encoding) as csvfile:
                    if hasattr(dialect, 'hasheader') and dialect.hasheader:
                        header = get_header(csvfile, dialect)
                        self.model.setHorizontalHeaderLabels(header)
                        self.horizontalHeader().show()
                    else:
                        self.horizontalHeader().hide()

                    for i, row in enumerate(csv_reader(csvfile, dialect)):
                        if i >= self.no_rows:
                            break
                        if i == 0:
                            self.add_choice_row(len(row))
                        if digest_types is None:
                            item_row = map(QStandardItem, map(str, row))
                        else:
                            codes = (convert(ele, t)
                                     for ele, t in zip(row, digest_types))
                            item_row = map(QStandardItem, codes)
                        self.model.appendRow(item_row)
            except UnicodeDecodeError:
                QMessageBox.warning(self.parent, "Encoding error",
                                    f"File is not encoded in {encoding}.")
                self.model.clear()

        except (OSError, csv.Error) as error:
            title = "CSV Import Error"
            text_tpl = "Error importing csv file {path}.\n \n" +\
                       "{errtype}: {error}"
            text = text_tpl.format(path=filepath, errtype=type(error).__name__,
                                   error=error)
            QMessageBox.warning(self.parent, title, text)

    def get_digest_types(self) -> List[str]:
        """Returns list of digest types from comboboxes"""

        try:
            return [cbox.currentText() for cbox in self.comboboxes]
        except RuntimeError:
            return []

    def update_comboboxes(self, digest_types: List[str]):
        """Updates the cono boxes to show digest_types

        :param digest_types: Names of preprocessing functions for csv values

        """

        for combobox, digest_type in zip(self.comboboxes, digest_types):
            combobox.setCurrentText(digest_type)


class CsvImportDialog(QDialog):
    """Modal dialog for importing csv files"""

    title = "CSV import"

    def __init__(self, parent: QWidget, filepath: Path,
                 digest_types: List[str] = None):
        """
        :param parent: Parent window
        :param filepath: Path to csv file
        :param digest_types: Names of preprocessing functions for csv values

        """

        super().__init__(parent)

        self.parent = parent
        self.filepath = filepath
        self.digest_types = digest_types

        self.sniff_size = parent.settings.sniff_size

        self.csv_encoding = 'utf-8'
        self.dialect = None

        self.setWindowTitle(self.title)

        self.parameter_groupbox = CsvParameterGroupBox(self)
        self.csv_table = CsvTable(self)

        layout = QVBoxLayout(self)
        layout.addWidget(self.parameter_groupbox)
        layout.addWidget(self.csv_table)
        layout.addWidget(self.create_buttonbox())

        self.setLayout(layout)

        self.reset()

    def create_buttonbox(self):
        """Returns a QDialogButtonBox"""

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Reset
                                      | QDialogButtonBox.StandardButton.Apply
                                      | QDialogButtonBox.StandardButton.Ok
                                      | QDialogButtonBox.StandardButton.Cancel)

        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(
            QDialogButtonBox.StandardButton.Reset).clicked.connect(self.reset)
        button_box.button(
            QDialogButtonBox.StandardButton.Apply).clicked.connect(self.apply)

        return button_box

    def _sniff_dialect(self):
        """Sniff the dialect of self.filepath`"""

        try:
            return sniff(self.filepath, self.sniff_size, self.csv_encoding)
        except UnicodeError:
            self.csv_encoding, ok = QInputDialog().getItem(
                self, f"{self.filepath} not encoded in utf-8",
                f"Encoding of {self.filepath}",
                self.parent.settings.encodings)
            if ok:
                return self._sniff_dialect()
        except (OSError, csv.Error) as error:
            title = "CSV Import Error"
            text = f"Error sniffing csv file {self.filepath}.\n \n" + \
                   f"{type(error).__name__}: {error}"
            QMessageBox.warning(self.parent, title, text)

    # Button event handlers

    def reset(self):
        """Button event handler, resets parameter_groupbox and csv_table"""

        dialect = self._sniff_dialect()
        if dialect is None:
            return

        self.parameter_groupbox.set_csvdialect(dialect)
        self.csv_table.fill(self.filepath, dialect)
        if self.digest_types is not None:
            self.csv_table.update_comboboxes(self.digest_types)

    def apply(self):
        """Button event handler, applies parameters to csv_table"""

        sniff_dialect = self._sniff_dialect()
        if sniff_dialect is None:
            return

        try:
            dialect = self.parameter_groupbox.adjust_csvdialect(sniff_dialect)
        except AttributeError as error:
            title = "CSV Import Error"
            text_tpl = "Error setting dialect for csv file {path}.\n \n" +\
                       "{errtype}: {error}"
            text = text_tpl.format(path=self.filepath,
                                   errtype=type(error).__name__, error=error)
            QMessageBox.warning(self.parent, title, text)
            return

        digest_types = self.csv_table.get_digest_types()
        self.csv_table.fill(self.filepath, dialect, digest_types)
        self.csv_table.update_comboboxes(digest_types)

    def accept(self):
        """Button event handler, starts csv import"""

        sniff_dialect = self._sniff_dialect()
        if sniff_dialect is None:
            return

        try:
            dialect = self.parameter_groupbox.adjust_csvdialect(sniff_dialect)
        except AttributeError as error:
            title = "CSV Import Error"
            text_tpl = "Error setting dialect for csv file {path}.\n \n" +\
                       "{errtype}: {error}"
            text = text_tpl.format(path=self.filepath,
                                   errtype=type(error).__name__, error=error)
            QMessageBox.warning(self.parent, title, text)
            self.reject()
            return

        self.digest_types = self.csv_table.get_digest_types()
        self.dialect = dialect

        super().accept()


class CsvExportDialog(QDialog):
    """Modal dialog for exporting csv files"""

    title = "CSV export"
    maxrows = 10

    def __init__(self, parent: QWidget, csv_area: SinglePageArea):
        """
        :param parent: Parent window
        :param csv_area: Grid area to be exported

        """

        super().__init__(parent)

        self.parent = parent

        self.csv_area = csv_area

        self.dialect = self.default_dialect

        self.setWindowTitle(self.title)

        self.parameter_groupbox = CsvParameterGroupBox(self)
        self.csv_preview = QPlainTextEdit(self)
        self.csv_preview.setReadOnly(True)

        layout = QVBoxLayout(self)
        layout.addWidget(self.parameter_groupbox)
        layout.addWidget(self.csv_preview)
        layout.addWidget(self.create_buttonbox())

        self.setLayout(layout)

        self.reset()

    @property
    def default_dialect(self) -> csv.Dialect:
        """Default dialect for export based on excel-tab"""

        dialect = csv.excel
        dialect.encoding = "utf-8"
        dialect.hasheader = False

        return dialect

    def reset(self):
        """Button event handler, resets parameter_groupbox and csv_preview"""

        self.parameter_groupbox.set_csvdialect(self.default_dialect)
        self.csv_preview.clear()

    def apply(self):
        """Button event handler, applies parameters to csv_preview"""

        top = self.csv_area.top
        left = self.csv_area.left
        bottom = self.csv_area.bottom
        right = self.csv_area.right
        table = self.parent.grid.table

        bottom = min(bottom-top, self.maxrows-1) + top

        code_array = self.parent.grid.model.code_array
        csv_data = code_array[top: bottom + 1, left: right + 1, table]

        adjust_csvdialect = self.parameter_groupbox.adjust_csvdialect
        dialect = adjust_csvdialect(self.default_dialect)

        str_io = io.StringIO()
        writer = csv.writer(str_io, dialect=dialect)
        writer.writerows(csv_data)

        self.csv_preview.setPlainText(str_io.getvalue())

    def accept(self):
        """Button event handler, starts csv import"""

        adjust_csvdialect = self.parameter_groupbox.adjust_csvdialect
        self.dialect = adjust_csvdialect(self.default_dialect)
        super().accept()

    def create_buttonbox(self) -> QDialogButtonBox:
        """Returns button box with Reset, Apply, Ok, Cancel"""

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Reset
                                      | QDialogButtonBox.StandardButton.Apply
                                      | QDialogButtonBox.StandardButton.Ok
                                      | QDialogButtonBox.StandardButton.Cancel)

        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(
            QDialogButtonBox.StandardButton.Reset).clicked.connect(self.reset)
        button_box.button(
            QDialogButtonBox.StandardButton.Apply).clicked.connect(self.apply)

        return button_box


class TutorialDialog(QDialog):
    """Dialog for browsing the PyCellSheet tutorial"""

    window_title = "PyCellSheet tutorial"
    size_hint = 1000, 800
    path: Path = TUTORIAL_PATH / 'tutorial.md'

    def __init__(self, parent: QWidget):
        """
        :param parent: Parent window

        """

        super().__init__(parent)

        self._create_widgets()
        self._layout()

    def _create_widgets(self):
        """Creates dialog widgets, e.g. the browser"""

        self.browser = HelpBrowser(self, self.path)

    def _layout(self):
        """Dialog layout management"""

        self.setWindowTitle(self.window_title)

        layout = QHBoxLayout()
        layout.addWidget(self.browser)
        self.setLayout(layout)

    # Overrides

    def sizeHint(self) -> QSize:
        """QDialog.sizeHint override"""

        return QSize(*self.size_hint)


class ManualDialog(TutorialDialog):
    """Dialog for browsing the PyCellSheet manual"""

    window_title = "PyCellSheet manual"
    size_hint = 1000, 800
    title2path = {
        "Overview": MANUAL_PATH / 'overview.md',
        "Concepts": MANUAL_PATH / 'basic_concepts.md',
        "Workspace": MANUAL_PATH / 'workspace.md',
        "File": MANUAL_PATH / 'file_menu.md',
        "Edit": MANUAL_PATH / 'edit_menu.md',
        "View": MANUAL_PATH / 'view_menu.md',
        "Format": MANUAL_PATH / 'format_menu.md',
        "Macro": MANUAL_PATH / 'macro_menu.md',
        "Advanced topics": MANUAL_PATH / 'advanced_topics.md',
    }

    def _create_widgets(self):
        """Creates tabbar and dialog browser"""

        self.tabbar = QTabWidget(self)
        for title, path in self.title2path.items():
            self.tabbar.addTab(HelpBrowser(self, path), title)

    def _layout(self):
        """Dialog layout management"""

        self.setWindowTitle(self.window_title)

        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)

        layout.addWidget(self.tabbar)


class PrintPreviewDialog(QPrintPreviewDialog):
    """Adds Mouse wheel functionality"""

    def __init__(self, printer: QPrinter):
        """
        :param printer: Target printer

        """

        super().__init__(printer)
        self.toolbar = self.findChildren(QToolBar)[0]
        self.actions = self.toolbar.actions()
        self.widget = self.findChildren(QPrintPreviewWidget)[0]
        self.combo_zoom = self.toolbar.widgetForAction(self.actions[3])

    def wheelEvent(self, event: QWheelEvent):
        """Overrides mouse wheel event handler

        :param event: Mouse wheel event

        """

        modifiers = QApplication.keyboardModifiers()
        if modifiers == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0:
                zoom_factor = self.widget.zoomFactor() / 1.1
            else:
                zoom_factor = self.widget.zoomFactor() * 1.1
            self.widget.setZoomFactor(zoom_factor)
            self.combo_zoom.setCurrentText(str(round(zoom_factor*100, 1))+"%")
        else:
            super().wheelEvent(event)
