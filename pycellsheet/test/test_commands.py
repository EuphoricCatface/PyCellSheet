# -*- coding: utf-8 -*-

from contextlib import contextmanager
from os.path import abspath, dirname, join
import sys
from types import SimpleNamespace

import pytest

from PyQt6.QtCore import Qt

PYSPREADPATH = abspath(join(dirname(__file__) + "/.."))


@contextmanager
def insert_path(path):
    sys.path.insert(0, path)
    yield
    sys.path.pop(0)


with insert_path(PYSPREADPATH):
    from ..commands import (
        MakeButtonCell,
        RemoveButtonCell,
        RenameSheet,
        SetCellCode,
        SetCellMerge,
        SetColumnsWidth,
        SetGridSize,
        SetRowsHeight,
    )
    from ..lib.attrdict import AttrDict
    from ..lib.selection import Selection
    from ..model.model import CellAttribute


class DummyIndex:
    def __init__(self, row, column):
        self._row = row
        self._column = column

    def row(self):
        return self._row

    def column(self):
        return self._column


class DummySignal:
    def __init__(self):
        self.calls = []

    def emit(self, *args):
        self.calls.append(args)


class DummyCellAttributes:
    def __init__(self):
        self._stack = []
        self._by_key = {}

    def append(self, attr):
        self._stack.append(attr)
        selection, table, attr_dict = attr
        for row, col in selection.cells:
            key = (row, col, table)
            current = dict(self._by_key.get(key, {}))
            current.update(attr_dict)
            self._by_key[key] = AttrDict(current)

    def pop(self):
        if not self._stack:
            raise IndexError("pop from empty list")
        return self._stack.pop()

    def __getitem__(self, key):
        if key not in self._by_key:
            self._by_key[key] = AttrDict({"locked": False, "button_cell": False})
        return self._by_key[key]

    def clear(self):
        self._stack.clear()
        self._by_key.clear()


class DummyCodeArray:
    def __init__(self):
        self._data = {}
        self.row_heights = {}
        self.col_widths = {}
        self.cell_attributes = DummyCellAttributes()
        self.dict_grid = SimpleNamespace(
            row_heights=self.row_heights,
            col_widths=self.col_widths,
            cell_attributes=self.cell_attributes,
            sheet_names=["Sheet 0", "Sheet 1"],
        )

    def __iter__(self):
        return iter(self._data)

    def __contains__(self, key):
        return key in self._data

    def __getitem__(self, key):
        return self._data[key]

    def __setitem__(self, key, value):
        self._data[key] = value

    def __call__(self, key):
        return self._data.get(key)

    def keys(self):
        return list(self._data.keys())

    def pop(self, key):
        return self._data.pop(key)


class DummyHighlighter:
    def __init__(self):
        self._document = object()

    def document(self):
        return self._document

    def setDocument(self, document):
        self._document = document


class DummyEntryLine:
    def __init__(self):
        self.highlighter = DummyHighlighter()

    @contextmanager
    def disable_updates(self):
        yield


class DummyModel:
    def __init__(self):
        self.shape = (5, 5, 2)
        self.code_array = DummyCodeArray()
        self.emit_data_changed_all_calls = 0
        self.dataChanged = DummySignal()
        self.main_window = SimpleNamespace(entry_line=DummyEntryLine(), grids=[])

    def index(self, row, col, parent=None):
        return DummyIndex(row, col)

    def code(self, index):
        return self.code_array((index.row(), index.column(), 0))

    def setData(self, index_or_indices, value, role, raw=False, table=None):
        if isinstance(index_or_indices, list):
            indices = index_or_indices
        else:
            indices = [index_or_indices]

        if raw:
            index = indices[0]
            key = (index.row(), index.column(), table if table is not None else 0)
            self.code_array[key] = value
            return True

        if role == Qt.ItemDataRole.EditRole:
            for index in indices:
                self.code_array[(index.row(), index.column(), 0)] = value
            return True

        if role == Qt.ItemDataRole.DecorationRole:
            for index in indices:
                selection, table_idx, attr_dict = value
                key = (index.row(), index.column(), table_idx)
                current = dict(self.code_array.cell_attributes[key])
                current.update(attr_dict)
                self.code_array.cell_attributes._by_key[key] = AttrDict(current)
                self.code_array.cell_attributes.append(value)
            return True

        return True

    def emit_data_changed_all(self):
        self.emit_data_changed_all_calls += 1


class DummyHeader:
    def __init__(self, size):
        self._size = size

    def defaultSectionSize(self):
        return self._size


class DummyGrid:
    def __init__(self, model, table=0):
        self.model = model
        self.table = table
        self.zoom = 1.0
        self.row_height_calls = []
        self.column_width_calls = []
        self.update_cell_spans_calls = 0
        self.update_index_widgets_calls = 0
        self.table_choice = SimpleNamespace(
            on_table_changed=lambda _table: None,
            setTabText=lambda _idx, _name: None,
        )
        self.main_window = SimpleNamespace(grids=[self], table_choice=self.table_choice)

    def verticalHeader(self):
        return DummyHeader(20)

    def horizontalHeader(self):
        return DummyHeader(20)

    @contextmanager
    def undo_resizing_row(self):
        yield

    @contextmanager
    def undo_resizing_column(self):
        yield

    def setRowHeight(self, row, height):
        self.row_height_calls.append((row, height))

    def setColumnWidth(self, col, width):
        self.column_width_calls.append((col, width))

    def update_cell_spans(self):
        self.update_cell_spans_calls += 1

    def update_index_widgets(self):
        self.update_index_widgets_calls += 1


def build_grid_pair(table=0):
    model = DummyModel()
    grid_a = DummyGrid(model, table=table)
    grid_b = DummyGrid(model, table=table)
    grid_a.main_window.grids = [grid_a, grid_b]
    grid_b.main_window.grids = [grid_a, grid_b]
    model.main_window.grids = [grid_a, grid_b]
    return model, grid_a, grid_b


def test_set_grid_size_redo_and_undo_restore_cells():
    model, grid, _ = build_grid_pair()
    model.code_array[(0, 0, 0)] = "keep"
    model.code_array[(4, 4, 1)] = "drop"

    cmd = SetGridSize(grid, old_shape=(5, 5, 2), new_shape=(2, 2, 1), description="Resize")
    cmd.redo()
    assert model.shape == (2, 2, 1)
    assert (4, 4, 1) not in model.code_array

    cmd.undo()
    assert model.shape == (5, 5, 2)
    assert model.code_array[(4, 4, 1)] == "drop"


def test_set_rows_height_redo_and_undo_update_storage():
    model, grid_a, grid_b = build_grid_pair()
    cmd = SetRowsHeight(grid_a, rows=[1, 2], table=0, old_height=5.0, new_height=7.0, description="Rows")

    cmd.redo()
    assert model.code_array.row_heights[(1, 0)] == 7.0
    assert model.code_array.row_heights[(2, 0)] == 7.0
    assert len(grid_a.row_height_calls) == 2
    assert len(grid_b.row_height_calls) == 2

    cmd.undo()
    assert model.code_array.row_heights[(1, 0)] == 5.0
    assert model.code_array.row_heights[(2, 0)] == 5.0


def test_set_columns_width_default_size_removes_explicit_width():
    model, grid, _ = build_grid_pair()
    model.code_array.col_widths[(1, 0)] = 9.0
    cmd = SetColumnsWidth(grid, columns=[1], table=0, old_width=9.0, new_width=20.0, description="Cols")

    cmd.redo()
    assert (1, 0) not in model.code_array.col_widths

    cmd.undo()
    assert model.code_array.col_widths[(1, 0)] == 9.0


def test_set_cell_code_merge_and_redo_undo():
    model, _, _ = build_grid_pair()
    index = model.index(0, 0)
    model.code_array[(0, 0, 0)] = "old"

    cmd1 = SetCellCode("new1", model, index, "Edit")
    cmd2 = SetCellCode("new2", model, index, "Edit")
    assert cmd1.mergeWith(cmd2) is True

    cmd1.redo()
    assert model.code_array[(0, 0, 0)] == "new2"
    cmd1.undo()
    assert model.code_array[(0, 0, 0)] == "old"


def test_set_cell_merge_undo_raises_warning_on_empty_attributes():
    model, _, _ = build_grid_pair()
    index = model.index(0, 0)
    attr = CellAttribute(Selection([], [], [], [], [(0, 0)]), 0, AttrDict({"merge_area": (0, 0, 1, 1)}))
    cmd = SetCellMerge(attr, model, index, [index], "Merge")

    with pytest.raises(Warning):
        cmd.undo()


def test_rename_sheet_redo_and_undo_updates_names_and_tabs():
    model, grid, _ = build_grid_pair()
    tab_changes = []
    grid.main_window.table_choice.setTabText = lambda idx, name: tab_changes.append((idx, name))

    cmd = RenameSheet(grid, sheet_index=1, old_name="Sheet 1", new_name="Renamed", description="Rename")
    cmd.redo()
    assert model.code_array.dict_grid.sheet_names[1] == "Renamed"
    cmd.undo()
    assert model.code_array.dict_grid.sheet_names[1] == "Sheet 1"
    assert tab_changes == [(1, "Renamed"), (1, "Sheet 1")]


def test_make_and_remove_button_cell_updates_widgets_for_active_table():
    model, grid_a, grid_b = build_grid_pair(table=0)
    index = model.index(0, 0)

    make_cmd = MakeButtonCell(grid_a, text="Run", index=index, description="MakeButton")
    make_cmd.redo()
    assert model.code_array.cell_attributes[(0, 0, 0)].button_cell == "Run"
    assert grid_a.update_index_widgets_calls == 1
    assert grid_b.update_index_widgets_calls == 1

    remove_cmd = RemoveButtonCell(grid_a, index=index, description="RemoveButton")
    remove_cmd.redo()
    assert model.code_array.cell_attributes[(0, 0, 0)].button_cell is False
    remove_cmd.undo()
    assert model.code_array.cell_attributes[(0, 0, 0)].button_cell == "Run"


def test_make_button_cell_skips_widget_update_if_table_changed_after_init():
    model, grid_a, grid_b = build_grid_pair(table=0)
    index = model.index(0, 0)
    cmd = MakeButtonCell(grid_a, text="Run", index=index, description="MakeButton")

    grid_a.table = 1
    cmd.redo()
    assert grid_a.update_index_widgets_calls == 0
    assert grid_b.update_index_widgets_calls == 0
