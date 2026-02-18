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

The model contains the core data structures of PyCellSheet and is divided
into the following layers.

- Layer 3: :class:`CodeArray`
- Layer 2: :class:`DataArray`
- Layer 1: :class:`DictGrid`
- Layer 0: :class:`KeyValueStore`


**Provides**

 * :class:`DefaultCellAttributeDict`
 * :class:`CellAttribute`
 * :class:`CellAttributes`
 * :class:`KeyValueStore`
 * :class:`DictGrid`
 * :class:`DataArray`
 * :class:`CodeArray`

"""

from __future__ import absolute_import

import typing, types
from builtins import filter
from builtins import str
from builtins import zip
from builtins import range

import ast
from collections import defaultdict
from copy import copy, deepcopy
from inspect import isgenerator
import io
from itertools import product
from pydoc import plaintext, render_doc
import re
import sys
from traceback import print_exception
from typing import (
        Any, Dict, Iterable, List, NamedTuple, Sequence, Tuple, Union)

import numpy

from PyQt6.QtGui import QImage, QPixmap  # Needed

try:
    import dateutil
except ImportError:
    dateutil = None

try:
    from matplotlib.figure import Figure
except ImportError:
    Figure = None

try:
    from moneyed import Money
except ImportError:
    Money = None

try:
    from pycellsheet.settings import Settings
    from pycellsheet.lib.attrdict import AttrDict
    import pycellsheet.lib.charts as charts
    from pycellsheet.lib.exception_handling import get_user_codeframe
    from pycellsheet.lib.typechecks import is_stringlike
    from pycellsheet.lib.selection import Selection
    from pycellsheet.lib.string_helpers import ZEN
    from pycellsheet.lib.sheet_name import generate_unique_sheet_name
    from pycellsheet.lib.pycellsheet import EmptyCell, PythonCode, Range, HelpText, ExpressionParser, \
        ReferenceParser, RangeOutput, PythonEvaluator, CELL_META_GENERATOR, DependencyTracker

    # Dependency tracking and smart caching
    from pycellsheet.lib.dependency_graph import DependencyGraph
    from pycellsheet.lib.smart_cache import SmartCache
    from pycellsheet.lib.exceptions import CircularRefError

except ImportError:
    from settings import Settings
    from lib.attrdict import AttrDict
    import lib.charts as charts  # Needed
    from lib.exception_handling import get_user_codeframe
    from lib.typechecks import is_stringlike
    from lib.selection import Selection
    from lib.string_helpers import ZEN
    from lib.sheet_name import generate_unique_sheet_name
    from lib.pycellsheet import EmptyCell, PythonCode, Range, HelpText, ExpressionParser, \
        ReferenceParser, RangeOutput, PythonEvaluator, CELL_META_GENERATOR, DependencyTracker

    # Dependency tracking and smart caching
    from lib.dependency_graph import DependencyGraph
    from lib.smart_cache import SmartCache
    from lib.exceptions import CircularRefError


INITSCRIPT_DEFAULT = \
"""import random as random_
RANDOM_SEED = 0
random = random_.Random(RANDOM_SEED)
# For non-deterministic runs, uncomment and customize:
# import datetime as datetime_
# random.seed(datetime_.datetime.now().timestamp())

try:
    from pycellsheet.lib.spreadsheet import *
except ImportError:
    from lib.spreadsheet import *
## The above is equivalent to:
# try:
#     from pycellsheet.lib.spreadsheet.array import *
#     from pycellsheet.lib.spreadsheet.database import *
#     ...
# except ImportError:
#     from lib.spreadsheet.array import *
#     from lib.spreadsheet.database import *
#     ...
## If you want more pythonic namespaced imports, you can also go:
# try:
#     from pycellsheet.lib import spreadsheet
# except:
#     from lib import spreadsheet
"""


class_format_functions = {}


def _get_isolated_builtins() -> dict[str, Any]:
    """Returns a per-evaluation builtins mapping that is detached from runtime globals."""

    if isinstance(__builtins__, dict):
        return dict(__builtins__)
    return dict(__builtins__.__dict__)


class DefaultCellAttributeDict(AttrDict):
    """Holds default values for all cell attributes"""

    def __init__(self):
        super().__init__(self)

        self.borderwidth_bottom = 1
        self.borderwidth_right = 1
        self.bordercolor_bottom = 220, 220, 220
        self.bordercolor_right = 220, 220, 220
        self.bgcolor = 255, 255, 255  # Do not use theme
        self.textfont = None
        self.pointsize = 10
        self.fontweight = None
        self.fontstyle = None
        self.textcolor = 0, 0, 0  # Do not use theme
        self.underline = False
        self.strikethrough = False
        self.locked = False
        self.angle = 0.0
        self.vertical_align = "align_top"
        self.justification = "justify_left"
        self.merge_area = None
        self.renderer = "text"
        self.button_cell = False
        self.panel_cell = False


class CellAttribute(NamedTuple):
    """Single cell attribute"""

    selection: Selection
    table: int
    attr: AttrDict


class CellAttributes(list):
    """Stores cell formatting attributes in a list of CellAttribute instances

    The class stores cell attributes as a list of layers.
    Each layer describes attributes for one selection in one table.
    Ultimately, a cell's attributes are determined by going through all
    elements of an `CellAttributes` instance. A default `AttrDict` is updated
    with the one in the list element if it is relevant for the respective cell.
    Therefore, attributes are efficiently stored for large sets of cells.

    The class provides attribute read access to single cells via
    :meth:`__getitem__`.
    Otherwise it behaves similar to a `list`.

    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.__add__ = None
        self.__delattr__ = None
        self.__delitem__ = None
        self.__delslice__ = None
        self.__iadd__ = None
        self.__imul__ = None
        self.__rmul__ = None
        self.__setattr__ = None
        self.__setslice__ = None
        self.insert = None
        self.remove = None
        self.reverse = None
        self.sort = None

    # Cache for __getattr__ maps key to tuple of len and attr_dict

    _attr_cache = AttrDict()
    _table_cache = {}

    def append(self, cell_attribute: CellAttribute):
        """append that clears caches

        :param cell_attribute: Cell attribute to be appended

        """

        if not isinstance(cell_attribute, CellAttribute):
            msg = "{} not instance of CellAttribute".format(cell_attribute)
            raise UserWarning(msg)
            return

        # We need to clean up merge areas
        selection, table, attr = cell_attribute
        if "merge_area" in attr:
            for i, ele in enumerate(reversed(self)):
                if ele[0] == selection and ele[1] == table \
                   and "merge_area" in ele[2]:
                    try:
                        self.pop(-1 - i)
                    except IndexError:
                        pass
            if attr["merge_area"] is not None:
                super().append(cell_attribute)
        else:
            super().append(cell_attribute)

        self._attr_cache.clear()
        self._table_cache.clear()

    def __getitem__(self, key: Tuple[int, int, int]) -> AttrDict:
        """Returns attribute dict for a single key

        :param key: Key of cell for cell_attribute retrieval

        """

#        if any(isinstance(key_ele, slice) for key_ele in key):
#            raise Warning("slice in key {}".format(key))
#            return

        try:
            cache_len, cache_dict = self._attr_cache[key]

            # Use cache result only if no new attrs have been defined
            if cache_len == len(self):
                return cache_dict
        except KeyError:
            pass

        # Update table cache if it is outdated (e.g. when creating a new grid)
        if len(self) != self._len_table_cache():
            self._update_table_cache()

        row, col, tab = key

        result_dict = DefaultCellAttributeDict()

        try:
            for selection, attr_dict in self._table_cache[tab]:
                if (row, col) in selection:
                    result_dict.update(attr_dict)
        except KeyError:
            pass

        # Upddate cache with current length and dict
        self._attr_cache[key] = (len(self), result_dict)

        return result_dict

    def __setitem__(self, index: int, cell_attribute: CellAttribute):
        """__setitem__ that clears caches

        :param index: Index of item in self
        :param cell_attribute: Cell attribute to be set

        """

        if not isinstance(cell_attribute, CellAttribute):
            msg = "{} not instance of CellAttribute".format(cell_attribute)
            raise Warning(msg)
            return

        super().__setitem__(index, cell_attribute)

        self._attr_cache.clear()
        self._table_cache.clear()

    def _len_table_cache(self) -> int:
        """Returns the length of the table cache"""

        length = 0

        for table in self._table_cache:
            length += len(self._table_cache[table])

        return length

    def _update_table_cache(self):
        """Clears and updates the table cache to be in sync with self"""

        self._table_cache.clear()
        for sel, tab, val in self:
            try:
                self._table_cache[tab].append((sel, val))
            except KeyError:
                self._table_cache[tab] = [(sel, val)]

        if len(self) != self._len_table_cache():
            raise Warning("Length of _table_cache does not match")

    def get_merging_cell(self,
                         key: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """Returns key of cell that merges the cell key

        Retuns None if cell key not merged.

        :param key: Key of the cell that is merged

        """

        row, col, tab = key

        # Is cell merged
        for selection, table, attr in self:
            if tab == table and "merge_area" in attr:
                top, left, bottom, right = attr["merge_area"]
                if top <= row <= bottom and left <= col <= right:
                    return top, left, tab

    def for_table(self, table: int) -> list:
        """Return cell attributes for a given table

        Return type should be `CellAttributes`. The list return type is
        provided because PEP 563 is unavailable in Python 3.6.

        Note that the table's presence in the grid is not checked.

        :param table: Table for which cell attributes are returned

        """

        table_cell_attributes = CellAttributes()

        for selection, __table, attr in self:
            if __table == table:
                cell_attribute = CellAttribute(selection, __table, attr)
                table_cell_attributes.append(cell_attribute)

        return table_cell_attributes

# End of class CellAttributes


class KeyValueStore(dict):
    """Key-Value store in memory. Currently a dict with default value None.

    This class represents layer 0 of the model.

    """

    def __init__(self, default_value=None):
        """
        :param default_value: Value that is provided for missing keys

        """

        super().__init__()
        self.default_value = default_value

    def __missing__(self, value: Any) -> Any:
        """Returns the default value None"""

        return self.default_value

# End of class KeyValueStore

# -----------------------------------------------------------------------------


class DictGrid(KeyValueStore):
    """Core data class with all information that is stored in a `.pycs` file.

    Besides grid code access via standard `dict` operations, it provides
    the following attributes:

    * :attr:`~DictGrid.cell_attributes` -  Stores cell formatting attributes
    * :attr:`~DictGrid.sheet_scripts` - Per-sheet initialization scripts

    This class represents layer 1 of the model.

    """

    def __init__(self, shape: Tuple[int, int, int]):
        """
        :param shape: Shape of the grid

        """

        super().__init__()

        self.shape = shape

        # Instance of :class:`CellAttributes`
        self.cell_attributes = CellAttributes()

        # Canonical internal storage: per-sheet init scripts
        self.sheet_scripts: list[str] = [u"" for _ in range(shape[2])]
        # Sheet names corresponding to table indices
        self.sheet_names: list[str] = [f"Sheet {i}" for i in range(shape[2])]
        self.exp_parser_code = u""

        self.row_heights = defaultdict(float)  # Keys have format (row, table)
        self.col_widths = defaultdict(float)  # Keys have format (col, table)

    def __getitem__(self, key: Tuple[int, int, int]) -> Any:
        """
        :param key: Cell key

        """
        shape = self.shape

        for axis, key_ele in enumerate(key):
            if shape[axis] <= key_ele or key_ele < -shape[axis]:
                msg = "Grid index {key} outside grid shape {shape}."
                msg = msg.format(key=key, shape=shape)
                raise IndexError(msg)

        return super().__getitem__(key)

    def __missing__(self, key):
        """Default value is None"""

        return

    @property
    def macros(self) -> list[str]:
        """Compatibility alias for legacy save/load naming."""

        return self.sheet_scripts

    @macros.setter
    def macros(self, macros: list[str]):
        """Compatibility alias setter for legacy save/load naming."""

        self.sheet_scripts = macros

# End of class DictGrid

# -----------------------------------------------------------------------------


class DataArray:
    """DataArray provides enhanced grid read/write access.

    Enhancements comprise:
     * Slicing
     * Multi-dimensional operations such as insertion and deletion along one
       axis

    This class represents layer 2 of the model.

    """

    def __init__(self, shape: Tuple[int, int, int], settings: Settings):
        """
        :param shape: Shape of the grid
        :param settings: Pyspread settings

        """

        self.dict_grid = DictGrid(shape)
        self.settings = settings

        self.sheet_scripts_draft: list[typing.Optional[str]] = [INITSCRIPT_DEFAULT for _ in range(shape[2])]
        self.sheet_globals_copyable: list[dict[str, typing.Any]] = [dict() for _ in range(shape[2])]
        self.sheet_globals_uncopyable: list[dict[str, typing.Any]] = [dict() for i in range(shape[2])]

        self.exp_parser = ExpressionParser()
        self.exp_parser_code = ExpressionParser.DEFAULT_PARSERS["Mixed"]  # Workaround until we make a UI

    def __eq__(self, other) -> bool:
        if not hasattr(other, "dict_grid") or \
           not hasattr(other, "cell_attributes"):
            return False

        return self.dict_grid == other.dict_grid and \
            self.cell_attributes == other.cell_attributes

    def __ne__(self, other) -> bool:
        return not self.__eq__(other)

    @property
    def data(self) -> dict:
        """Returns `dict` of data content.


        - Data is the central content interface for loading / saving data.
        - It shall be used for loading and saving from and to `.pycs` and other
          files.
        - It shall be used for loading and saving sheet scripts.
        - However, it is not used for importing and exporting data because
          these operations are partial to the grid.

        **Content of returned dict**

        :param shape: Grid shape
        :type shape: Tuple[int, int, int]
        :param grid: Cell content
        :type grid: Dict[Tuple[int, int, int], str]
        :param attributes: Cell attributes
        :type attributes: CellAttribute
        :param row_heights: Row heights
        :type row_heights: defaultdict[Tuple[int, int], float]
        :param col_widths: Column widths
        :type col_widths: defaultdict[Tuple[int, int], float]
        :param macros: Legacy sheet script key for compatibility
        :type macros: list[str]

        """

        data = {}

        data["shape"] = self.shape
        data["grid"] = {}.update(self.dict_grid)
        data["attributes"] = self.cell_attributes[:]
        data["row_heights"] = self.row_heights
        data["col_widths"] = self.col_widths
        data["macros"] = self.sheet_scripts
        data["sheet_scripts"] = self.sheet_scripts

        return data

    @data.setter
    def data(self, **kwargs):
        """Sets data from given parameters

        Old values are deleted.
        If a paremeter is not given, nothing is changed.

        **Content of kwargs dict**

        :param shape: Grid shape
        :type shape: Tuple[int, int, int]
        :param grid: Cell content
        :type grid: Dict[Tuple[int, int, int], str]
        :param attributes: Cell attributes
        :type attributes: CellAttribute
        :param row_heights: Row heights
        :type row_heights: defaultdict[Tuple[int, int], float]
        :param col_widths: Column widths
        :type col_widths: defaultdict[Tuple[int, int], float]
        :param macros: Legacy sheet script key for compatibility
        :type macros: list[str]

        """

        if "shape" in kwargs:
            self.shape = kwargs["shape"]

        if "grid" in kwargs:
            self.dict_grid.clear()
            self.dict_grid.update(kwargs["grid"])

        if "attributes" in kwargs:
            self.attributes[:] = kwargs["attributes"]

        if "row_heights" in kwargs:
            self.row_heights = kwargs["row_heights"]

        if "col_widths" in kwargs:
            self.col_widths = kwargs["col_widths"]

        if "sheet_scripts" in kwargs:
            self.sheet_scripts = kwargs["sheet_scripts"]
        elif "macros" in kwargs:
            self.sheet_scripts = kwargs["macros"]

        if "exp_parser_code" in kwargs:
            pass

    @property
    def row_heights(self) -> defaultdict:
        """row_heights interface to dict_grid"""

        return self.dict_grid.row_heights

    @row_heights.setter
    def row_heights(self, row_heights: defaultdict):
        """row_heights interface to dict_grid"""

        self.dict_grid.row_heights = row_heights

    @property
    def col_widths(self) -> defaultdict:
        """col_widths interface to dict_grid"""

        return self.dict_grid.col_widths

    @col_widths.setter
    def col_widths(self, col_widths: defaultdict):
        """col_widths interface to dict_grid"""

        self.dict_grid.col_widths = col_widths

    @property
    def cell_attributes(self) -> CellAttributes:
        """cell_attributes interface to dict_grid"""

        return self.dict_grid.cell_attributes

    @cell_attributes.setter
    def cell_attributes(self, value: CellAttributes):
        """cell_attributes interface to dict_grid"""

        # First empty cell_attributes
        self.cell_attributes[:] = []
        self.cell_attributes.extend(value)

    @property
    def sheet_scripts(self) -> list[str]:
        """Canonical interface to per-sheet init scripts in dict_grid."""

        return self.dict_grid.sheet_scripts

    @sheet_scripts.setter
    def sheet_scripts(self, sheet_scripts: list[str]):
        """Sets per-sheet init scripts."""

        self.dict_grid.sheet_scripts = sheet_scripts

    @property
    def macros(self) -> list[str]:
        """Legacy alias of `sheet_scripts` for compatibility."""

        return self.sheet_scripts

    @macros.setter
    def macros(self, macros: list[str]):
        """Legacy alias setter of `sheet_scripts` for compatibility."""

        self.sheet_scripts = macros

    @property
    def macros_draft(self) -> list[typing.Optional[str]]:
        """Legacy alias of `sheet_scripts_draft` for compatibility."""

        return self.sheet_scripts_draft

    @macros_draft.setter
    def macros_draft(self, macros_draft: list[typing.Optional[str]]):
        """Legacy alias setter of `sheet_scripts_draft` for compatibility."""

        self.sheet_scripts_draft = macros_draft

    @property
    def exp_parser_code(self) -> str:
        """macros interface to dict_grid"""

        return self.dict_grid.exp_parser_code

    @exp_parser_code.setter
    def exp_parser_code(self, exp_parser_code: str):
        """Sets ExpressionParser string"""

        self.dict_grid.exp_parser_code = exp_parser_code
        self.exp_parser.set_parser(exp_parser_code)

    @property
    def shape(self) -> Tuple[int, int, int]:
        """Returns dict_grid shape"""

        return self.dict_grid.shape

    @shape.setter
    def shape(self, shape: Tuple[int, int, int]):
        """Deletes all cells beyond new shape and sets dict_grid shape

        Returns a dict of the deleted cells' contents

        :param shape: Target shape for grid

        """

        # Delete each cell that is beyond new borders

        old_shape = self.shape
        deleted_cells = {}

        if any(new_axis < old_axis
               for new_axis, old_axis in zip(shape, old_shape)):
            for key in list(self.dict_grid.keys()):
                if any(key_ele >= new_axis
                       for key_ele, new_axis in zip(key, shape)):
                    deleted_cells[key] = self.pop(key)

        # Set dict_grid shape attribute
        self.dict_grid.shape = shape

        # Adjust sheet-specific lists to match new table count
        old_tables = old_shape[2]
        new_tables = shape[2]
        if new_tables < old_tables:
            # Trim lists
            self.sheet_scripts = self.sheet_scripts[:new_tables]
            self.sheet_scripts_draft = self.sheet_scripts_draft[:new_tables]
            self.sheet_globals_copyable = self.sheet_globals_copyable[:new_tables]
            self.sheet_globals_uncopyable = self.sheet_globals_uncopyable[:new_tables]
            self.dict_grid.sheet_names = self.dict_grid.sheet_names[:new_tables]
        elif new_tables > old_tables:
            # Extend lists
            for i in range(old_tables, new_tables):
                self.sheet_scripts.append(u"")
                self.sheet_scripts_draft.append(None)
                self.sheet_globals_copyable.append(dict())
                self.sheet_globals_uncopyable.append(dict())
                new_name = generate_unique_sheet_name(
                    f"Sheet {i}",
                    self.dict_grid.sheet_names,
                    fallback_index=i,
                )
                self.dict_grid.sheet_names.append(new_name)

        self._adjust_rowcol(0, 0, 0)
        self._adjust_cell_attributes(0, 0, 0)

        return deleted_cells

    def __iter__(self) -> Iterable:
        """Returns iterator over self.dict_grid"""

        return iter(self.dict_grid)

    def __contains__(self, key: Tuple[int, int, int]) -> bool:
        """True if key is contained in grid

        Handles single keys only.

        :param key: Key of cell to be checked

        """

        if any(not isinstance(ele, int) for ele in key):
            return NotImplemented

        row, column, table = key
        rows, columns, tables = self.shape

        return (0 <= row <= rows
                and 0 <= column <= columns
                and 0 <= table <= tables)

    def __getitem__(self, key: Tuple[int, int, int]) -> str:
        """Cell code retrieval for a single key

        :param key: Cell key(s) that shall be set
        :return: Cell code

        """

        if any(isinstance(key_ele, slice) for key_ele in key):
            raise NotImplementedError("Slice-based cell access is no longer supported.")

        if any(is_stringlike(key_ele) for key_ele in key):
            raise NotImplementedError("Cell string based access not implemented")

        return self.dict_grid[key]

    def __setitem__(self, key: Tuple[int, int, int], value: str):
        """Accepts index keys

        :param key: Cell key that shall be set
        :param value: Code for the key

        """
        if any(isinstance(key_ele, slice) for key_ele in key):
            raise NotImplementedError("Slice-based cell assignment is no longer supported.")
        if any(is_stringlike(key_ele) for key_ele in key):
            raise NotImplementedError("Cell string based assignment not implemented")

        if value:
            merging_cell = self.cell_attributes.get_merging_cell(key)
            if ((merging_cell is None or merging_cell == key)
                    and isinstance(value, str)):
                self.dict_grid[key] = value
            return

        try:
            self.pop(key)
        except (KeyError, TypeError):
            pass

    # Pickle support

    def __getstate__(self) -> Dict[str, DictGrid]:
        """Returns dict_grid for pickling

        Note that all persistent data is contained in the DictGrid class

        """

        return {"dict_grid": self.dict_grid}

    def get_row_height(self, row: int, tab: int) -> float:
        """Returns row height

        :param row: Row for which height is retrieved
        :param tab: Table for which for which row height is retrieved

        """

        try:
            return self.row_heights[(row, tab)]

        except KeyError:
            return

    def get_col_width(self, col: int, tab: int) -> float:
        """Returns column width

        :param col: Column for which width is retrieved
        :param tab: Table for which for which column width is retrieved

        """

        try:
            return self.col_widths[(col, tab)]

        except KeyError:
            return

    def keys(self) -> List[Tuple[int, int, int]]:
        """Returns keys in self.dict_grid"""

        return list(self.dict_grid.keys())

    def pop(self, key: Tuple[int, int, int]) -> Any:
        """dict_grid pop wrapper

        :param key: Cell key

        """

        return self.dict_grid.pop(key)

    def get_last_filled_cell(self, table: int = None) -> Tuple[int, int, int]:
        """Returns key for the bottommost rightmost cell with content

        :param table: Limit search to this table

        """

        maxrow = 0
        maxcol = 0

        for row, col, tab in self.dict_grid:
            if table is None or tab == table:
                maxrow = max(row, maxrow)
                maxcol = max(col, maxcol)

        return maxrow, maxcol, table

    def _shift_rowcol(self, insertion_point: int, no_to_insert: int):
        """Shifts row and column sizes when a table is inserted or deleted

        :param insertion_point: Table at which a new table is inserted
        :param no_to_insert: Number of tables that are inserted

        """

        # Shift row heights

        new_row_heights = {}
        del_row_heights = []

        for row, tab in self.row_heights:
            if tab >= insertion_point:
                new_row_heights[(row, tab + no_to_insert)] = \
                    self.row_heights[(row, tab)]
                del_row_heights.append((row, tab))

        for row, tab in new_row_heights:
            self.set_row_height(row, tab, new_row_heights[(row, tab)])

        for row, tab in del_row_heights:
            if (row, tab) not in new_row_heights:
                self.set_row_height(row, tab, None)

        # Shift column widths

        new_col_widths = {}
        del_col_widths = []

        for col, tab in self.col_widths:
            if tab >= insertion_point:
                new_col_widths[(col, tab + no_to_insert)] = \
                    self.col_widths[(col, tab)]
                del_col_widths.append((col, tab))

        for col, tab in new_col_widths:
            self.set_col_width(col, tab, new_col_widths[(col, tab)])

        for col, tab in del_col_widths:
            if (col, tab) not in new_col_widths:
                self.set_col_width(col, tab, None)

    def _adjust_rowcol(self, insertion_point: int, no_to_insert: int,
                       axis: int, tab: int = None):
        """Adjusts row and column sizes on insertion/deletion

        :param insertion_point: Point on axis at which insertion takes place
        :param no_to_insert: Number of rows or columns that are inserted
        :param axis: Row insertion if 0, column insertion if 1, must be in 0, 1
        :param tab: Table at which insertion takes place, None means all tables

        """

        if axis == 2:
            self._shift_rowcol(insertion_point, no_to_insert)
            return

        if axis not in (0, 1):
            raise Warning("Axis {} not in (0, 1)".format(axis))
            return

        cell_sizes = self.col_widths if axis else self.row_heights
        set_cell_size = self.set_col_width if axis else self.set_row_height

        new_sizes = {}
        del_sizes = []

        for pos, table in cell_sizes:
            if pos >= insertion_point and (tab is None or tab == table):
                if 0 <= pos + no_to_insert < self.shape[axis]:
                    new_sizes[(pos + no_to_insert, table)] = \
                        cell_sizes[(pos, table)]
                if pos < insertion_point + no_to_insert:
                    new_sizes[(pos, table)] = cell_sizes[(pos, table)]
                del_sizes.append((pos, table))

        for pos, table in new_sizes:
            set_cell_size(pos, table, new_sizes[(pos, table)])

        for pos, table in del_sizes:
            if (pos, table) not in new_sizes:
                set_cell_size(pos, table, None)

    def _adjust_merge_area(
            self, attrs: AttrDict, insertion_point: int, no_to_insert: int,
            axis: int) -> Tuple[int, int, int, int]:
        """Returns an updated merge area

        :param attrs: Cell attribute dictionary that shall be adjusted
        :param insertion_point: Point on axis at which insertion takes place
        :param no_to_insert: Number of rows/cols/tabs to be inserted (>=0)
        :param axis: Row insertion if 0, column insertion if 1, must be in 0, 1

        """

        if axis not in (0, 1):
            raise Warning("Axis {} not in (0, 1)".format(axis))
            return

        if "merge_area" not in attrs or attrs["merge_area"] is None:
            return

        top, left, bottom, right = attrs["merge_area"]
        selection = Selection([(top, left)], [(bottom, right)], [], [], [])

        selection.insert(insertion_point, no_to_insert, axis)

        __top, __left = selection.block_tl[0]
        __bottom, __right = selection.block_br[0]

        # Adjust merge area if it is beyond the grid shape
        rows, cols, tabs = self.shape

        if __top < 0 and __bottom < 0:
            return
        if __top >= rows and __bottom >= rows:
            return
        if __left < 0 and __right < 0:
            return
        if __left >= cols and __right >= cols:
            return

        if __top < 0:
            __top = 0

        if __top >= rows:
            __top = rows - 1

        if __bottom < 0:
            __bottom = 0

        if __bottom >= rows:
            __bottom = rows - 1

        if __left < 0:
            __left = 0

        if __left >= cols:
            __left = cols - 1

        if __right < 0:
            __right = 0

        if __right >= cols:
            __right = cols - 1

        return __top, __left, __bottom, __right

    def _adjust_cell_attributes(
            self, insertion_point: int, no_to_insert: int,  axis: int,
            tab: int = None, cell_attrs: AttrDict = None):
        """Adjusts cell attributes on insertion/deletion

        :param insertion_point: Point on axis at which insertion takes place
        :param no_to_insert: Number of rows/cols/tabs to be inserted (>=0)
        :param axis: Row insertion if 0, column insertion if 1, must be in 0, 1
        :param tab: Table at which insertion takes place, None means all tables
        :param cell_attrs: If given replaces the existing cell attributes

        """

        def replace_cell_attributes_table(index: int, new_table: int):
            """Replaces table in cell_attributes item

            :param index: Cell attribute index for table replacement
            :param new_table: New table value for cell attribute

            """

            cell_attr = list(list.__getitem__(self.cell_attributes, index))
            cell_attr[1] = new_table
            self.cell_attributes[index] = CellAttribute(*cell_attr)

        def get_ca_with_updated_ma(
                attrs: AttrDict,
                merge_area: Tuple[int, int, int, int]) -> AttrDict:
            """Returns cell attributes with updated merge area

            :param attrs: Cell attributes to be updated
            :param merge_area: New merge area (top, left, bottom, right)

            """

            new_attrs = copy(attrs)

            if merge_area is None:
                try:
                    new_attrs.pop("merge_area")
                except KeyError:
                    pass
            else:
                new_attrs["merge_area"] = merge_area

            return new_attrs

        if axis not in list(range(3)):
            raise ValueError("Axis must be in [0, 1, 2]")

        if tab is not None and tab < 0:
            raise Warning("tab is negative")
            return

        if cell_attrs is None:
            cell_attrs = []

        if cell_attrs:
            self.cell_attributes[:] = cell_attrs

        elif axis < 2:
            # Adjust selections on given table

            ca_updates = {}
            for i, (selection, table, attrs) \
                    in enumerate(self.cell_attributes):
                selection = copy(selection)
                if tab is None or tab == table:
                    selection.insert(insertion_point, no_to_insert, axis)
                    # Update merge area if present
                    merge_area = self._adjust_merge_area(attrs,
                                                         insertion_point,
                                                         no_to_insert, axis)
                    new_attrs = get_ca_with_updated_ma(attrs, merge_area)

                    ca_updates[i] = CellAttribute(selection, table, new_attrs)

            for idx in ca_updates:
                self.cell_attributes[idx] = ca_updates[idx]

        elif axis == 2:
            # Adjust tabs

            pop_indices = []

            for i, cell_attribute in enumerate(self.cell_attributes):
                selection, table, value = cell_attribute

                if no_to_insert < 0 and insertion_point <= table:
                    if insertion_point > table + no_to_insert:
                        # Delete later
                        pop_indices.append(i)
                    else:
                        replace_cell_attributes_table(i, table + no_to_insert)

                elif insertion_point <= table:
                    # Insert
                    replace_cell_attributes_table(i, table + no_to_insert)

            for i in pop_indices[::-1]:
                self.cell_attributes.pop(i)

        self.cell_attributes._attr_cache.clear()
        self.cell_attributes._update_table_cache()

    def insert(self, insertion_point: int, no_to_insert: int, axis: int,
               tab: int = None):
        """Inserts no_to_insert rows/cols/tabs/... before insertion_point

        :param insertion_point: Point on axis at which insertion takes place
        :param no_to_insert: Number of rows/cols/tabs to be inserted (>=0)
        :param axis: Row/Column/Table insertion if 0/1/2 must be in 0, 1, 2
        :param tab: Table at which insertion takes place, None means all tables

        """

        if not 0 <= axis <= len(self.shape):
            raise ValueError("Axis not in grid dimensions")

        if insertion_point > self.shape[axis] or \
           insertion_point < -self.shape[axis]:
            raise IndexError("Insertion point not in grid")

        new_keys = {}
        del_keys = []

        for key in list(self.dict_grid.keys()):
            if key[axis] >= insertion_point and (tab is None or tab == key[2]):
                new_key = list(key)
                new_key[axis] += no_to_insert
                if 0 <= new_key[axis] < self.shape[axis]:
                    new_keys[tuple(new_key)] = self(key)
                del_keys.append(key)

        # Now re-insert moved keys

        for key in del_keys:
            if key not in new_keys and self(key) is not None:
                self.pop(key)

        self._adjust_rowcol(insertion_point, no_to_insert, axis, tab=tab)
        self._adjust_cell_attributes(insertion_point, no_to_insert, axis, tab)

        if axis == 2:
            for i in range(no_to_insert):
                self.sheet_scripts.insert(insertion_point, u"")
                self.sheet_scripts_draft.insert(insertion_point, None)
                self.sheet_globals_copyable.insert(insertion_point, dict())
                self.sheet_globals_uncopyable.insert(insertion_point, dict())
                new_name = generate_unique_sheet_name(
                    f"Sheet {insertion_point + i}",
                    self.dict_grid.sheet_names,
                    fallback_index=insertion_point + i,
                )
                self.dict_grid.sheet_names.insert(insertion_point, new_name)

        for key in new_keys:
            self.__setitem__(key, new_keys[key])

    def delete(self, deletion_point: int, no_to_delete: int, axis: int,
               tab: int = None):
        """Deletes no_to_delete rows/cols/... starting with deletion_point

        :param deletion_point: Point on axis at which deletion takes place
        :param no_to_delete: Number of rows/cols/tabs to be deleted (>=0)
        :param axis: Row/Column/Table deletion if 0/1/2, must be in 0, 1, 2
        :param tab: Table at which insertion takes place, None means all tables

        """

        if not 0 <= axis < len(self.shape):
            raise ValueError("Axis not in grid dimensions")

        if no_to_delete < 0:
            raise ValueError("Cannot delete negative number of rows/cols/...")

        if deletion_point > self.shape[axis] or \
           deletion_point <= -self.shape[axis]:
            raise IndexError("Deletion point not in grid")

        new_keys = {}
        del_keys = []

        # Note that the loop goes over a list that copies all dict keys
        for key in list(self.dict_grid.keys()):
            if tab is None or tab == key[2]:
                if deletion_point <= key[axis] < deletion_point + no_to_delete:
                    del_keys.append(key)

                elif key[axis] >= deletion_point + no_to_delete:
                    new_key = list(key)
                    new_key[axis] -= no_to_delete

                    new_keys[tuple(new_key)] = self(key)
                    del_keys.append(key)

        self._adjust_rowcol(deletion_point, -no_to_delete, axis, tab=tab)
        self._adjust_cell_attributes(deletion_point, -no_to_delete, axis, tab)

        if axis == 2:
            for _ in range(no_to_delete):
                if deletion_point < len(self.sheet_scripts):
                    self.sheet_scripts.pop(deletion_point)
                if deletion_point < len(self.sheet_scripts_draft):
                    self.sheet_scripts_draft.pop(deletion_point)
                if deletion_point < len(self.sheet_globals_copyable):
                    self.sheet_globals_copyable.pop(deletion_point)
                if deletion_point < len(self.sheet_globals_uncopyable):
                    self.sheet_globals_uncopyable.pop(deletion_point)
                if deletion_point < len(self.dict_grid.sheet_names):
                    self.dict_grid.sheet_names.pop(deletion_point)

        # Now re-insert moved keys

        for key in new_keys:
            self.__setitem__(key, new_keys[key])

        for key in del_keys:
            if key not in new_keys and self(key) is not None:
                self.pop(key)

    def set_row_height(self, row: int, tab: int, height: float):
        """Sets row height

        :param row: Row for height setting
        :param tab: Table, in which row height is set
        :param height: Row height to be set

        """

        try:
            self.row_heights.pop((row, tab))
        except KeyError:
            pass

        if height is not None:
            self.row_heights[(row, tab)] = float(height)

    def set_col_width(self, col: int, tab: int, width: float):
        """Sets column width

        :param col: Column for width setting
        :param tab: Table, in which column width is set
        :param width: Column width to be set

        """

        try:
            self.col_widths.pop((col, tab))
        except KeyError:
            pass

        if width is not None:
            self.col_widths[(col, tab)] = float(width)

    # Element access via call

    __call__ = __getitem__

# End of class DataArray

# -----------------------------------------------------------------------------


class CodeArray(DataArray):
    """CodeArray provides objects when accessing cells via `__getitem__`

    Cell code can be accessed via function call

    This class represents layer 3 of the model.

    """

    # Safe mode: If True then whether PyCellSheet is operating in safe_mode
    # In safe_mode, cells are not evaluated but its code is returned instead.
    safe_mode = False

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ref_parser = ReferenceParser(self)
        self.cell_meta_gen = CELL_META_GENERATOR(self)

        self.dep_graph = DependencyGraph()
        self.smart_cache = SmartCache(self.dep_graph)
        self._eval_warnings: dict[Tuple[int, int, int], list[str]] = {}
        self._format_warnings: dict[Tuple[int, int, int], str] = {}

    @staticmethod
    def _looks_like_cell_ref_name(name: str) -> bool:
        """Return True if `name` resembles a spreadsheet cell reference."""

        return bool(re.fullmatch(r"[A-Za-z]{1,4}[1-9][0-9]*", name))

    @staticmethod
    def _script_duplicate_import_warnings(sheet_script: str) -> list[str]:
        """Collect duplicate import-binding warnings from a sheet script."""

        warnings = []
        try:
            tree = ast.parse(sheet_script, mode="exec")
        except Exception:
            return warnings

        bindings: dict[str, int] = defaultdict(int)
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    bound = alias.asname or alias.name.split(".")[0]
                    bindings[bound] += 1
            elif isinstance(node, ast.ImportFrom):
                for alias in node.names:
                    if alias.name == "*":
                        continue
                    bound = alias.asname or alias.name
                    bindings[bound] += 1

        for name in sorted(bindings):
            if bindings[name] > 1:
                warnings.append(
                    f"WARNING: Duplicate import binding '{name}' appears {bindings[name]} times "
                    f"in this Sheet Script."
                )
        return warnings

    def _sheet_global_name_warnings(self, sheet_globals: dict[str, typing.Any]) -> list[str]:
        """Collect warnings for risky global names."""

        warnings = []
        for name in sorted(sheet_globals):
            if name == "__builtins__":
                continue
            if self._looks_like_cell_ref_name(name):
                warnings.append(
                    f"WARNING: Sheet global '{name}' looks like a cell reference and may be confusing."
                )
        return warnings

    def _clear_cell_warnings(self, key: Tuple[int, int, int]):
        self._eval_warnings.pop(key, None)
        self._format_warnings.pop(key, None)

    def _set_eval_warnings(self, key: Tuple[int, int, int], warnings: list[str]):
        if warnings:
            self._eval_warnings[key] = list(warnings)
            return
        self._eval_warnings.pop(key, None)

    def set_format_warning(self, key: Tuple[int, int, int], warning: typing.Optional[str]):
        if warning:
            self._format_warnings[key] = warning
            return
        self._format_warnings.pop(key, None)

    def get_cell_warnings(self, key: Tuple[int, int, int]) -> list[str]:
        warnings = []
        warnings.extend(self._eval_warnings.get(key, []))
        fmt = self._format_warnings.get(key)
        if fmt:
            warnings.append(fmt)
        return warnings

    def has_cell_warnings(self, key: Tuple[int, int, int]) -> bool:
        return bool(self._eval_warnings.get(key) or self._format_warnings.get(key))

    def __setitem__(self, key: Tuple[int, int, int], value: str):
        """Sets cell code and resets result cache

        :param key: Cell key(s) that shall be set
        :param value: Code for cell(s) to be set

        """

        if any(isinstance(key_ele, slice) for key_ele in key):
            raise NotImplementedError("Slice-based cell assignment is no longer supported.")

        # Change numpy array repr function for grid cell results
        try:
            numpy.set_printoptions(formatter = {'all': lambda s: repr(s.tolist() if hasattr(s, "tolist") else s)})
        except AttributeError:
            numpy.set_string_function(lambda s: repr(s.tolist() if hasattr(s, "tolist") else s))

        # Prevent unchanged cells from being recalculated on cursor movement.
        # Compare against stored code, not cache state.
        old_value = self(key)
        old_empty = old_value in (None, "")
        new_empty = value in (None, "")
        unchanged = (old_value == value) or (old_empty and new_empty)

        super().__setitem__(key, value)
        self._clear_cell_warnings(key)

        if not unchanged:
            # Invalidate cache for this cell and its dependents
            # IMPORTANT: Invalidate BEFORE removing from dep graph,
            # so invalidate() can propagate to dependents
            preserve_dependents_cache = (
                self.settings.recalc_mode == "manual"
            )
            self.smart_cache.invalidate(
                key,
                preserve_dependents_cache=preserve_dependents_cache,
            )
            # Remove dependencies for this cell (will be re-tracked on next eval)
            # Reverse edges are preserved by default for cycle detection
            self.dep_graph.remove_cell(key)

    def __getitem__(self, key: Tuple[int, int, int]) -> Any:
        """Returns _eval_cell

        :param key: Cell key for result retrieval (code if in safe mode)

        """

        if any(isinstance(key_ele, slice) for key_ele in key):
            raise NotImplementedError("Slice-based cell access is no longer supported.")

        code = self(key)

        if self.settings.recalc_mode == "manual" \
           and self.dep_graph.is_dirty(key):
            cached = self.smart_cache.get_raw(key)
            if cached is not SmartCache.INVALID:
                return deepcopy(cached)

        # Smart cache handling (check even for empty cells)
        cached = self.smart_cache.get(key)
        if cached is not SmartCache.INVALID:
            # Cache hit! Return deepcopied value (cache stores original)
            return deepcopy(cached)

        # Empty cells still need to be cached and have dirty flag cleared
        if code is None:
            # Store EmptyCell in cache
            self.smart_cache.set(key, EmptyCell)
            # Clear dirty flag
            self.dep_graph.clear_dirty(key)
            self._set_eval_warnings(key, [])
            return EmptyCell

        # Button cell handling
        if self.cell_attributes[key].button_cell is not False:
            self._set_eval_warnings(key, [])
            return

        # Normal cell handling - evaluate the cell
        result, eval_warnings = self._eval_cell(key, code, return_warnings=True)

        # Store in smart cache
        self.smart_cache.set(key, result)
        self._set_eval_warnings(key, eval_warnings)

        # Clear dirty flag after successful evaluation
        self.dep_graph.clear_dirty(key)

        return result

    def _make_nested_list(self, gen: Union[Iterable, Iterable[Iterable],
                                           Iterable[Iterable[Iterable]]]
                          ) -> Union[Sequence, Sequence[Sequence],
                                     Sequence[Sequence[Sequence]]]:
        """Makes nested list from generator for creating numpy.array"""

        res = []

        for ele in gen:
            if ele is None:
                res.append(None)

            elif not is_stringlike(ele) and isgenerator(ele):
                # Nested generator
                res.append(self._make_nested_list(ele))

            else:
                res.append(ele)

        return res

    def _eval_cell(self, key: Tuple[int, int, int], cell_contents: str,
                   return_warnings: bool = False) -> Any:
        """Evaluates one cell and returns its result

        :param key: Key of cell to be evaled
        :param cell_contents: Whatever the user typed in

        """
        eval_warnings: list[str] = []

        # Help helper function that fixes help being displayed in stdout
        def help(*args) -> HelpText:
            """Returns help string for object arguments"""

            if not args:
                return HelpText(args, ZEN)

            return HelpText(args, render_doc(*args, renderer=plaintext))

        if self.safe_mode:
            # Safe mode is active
            if return_warnings:
                return cell_contents, eval_warnings
            return cell_contents

        #  --- ExpParser START ---  #
        if self.exp_parser.handle_empty(cell_contents):
            if return_warnings:
                return EmptyCell, eval_warnings
            return EmptyCell
        exp_parsed = self.exp_parser.parse(cell_contents)
        if exp_parsed is EmptyCell and cell_contents.strip():
            eval_warnings.append(
                "Expression parser returned EmptyCell for non-empty cell contents."
            )
        if not isinstance(exp_parsed, PythonCode):
            if return_warnings:
                return exp_parsed, eval_warnings
            return exp_parsed
        #  --- ExpParser END ---  #

        #  --- RefParser ---  #
        try:
            ref_parsed = self.ref_parser.parser(exp_parsed)
        except Exception as err:
            if return_warnings:
                return err, eval_warnings
            return err

        # Rebuild this cell's forward dependency set on each evaluation.
        # Keep reverse edges so dependents of this cell remain known.
        self.dep_graph.remove_cell(key)

        #  --- Dependency Tracking & Cycle Detection START ---  #
        # Check for circular references before evaluating
        try:
            self.dep_graph.check_for_cycles(key)
        except CircularRefError as err:
            if return_warnings:
                return err, eval_warnings
            return err

        #  --- PythonEval START ---  #
        env = deepcopy(self.sheet_globals_copyable[key[2]])
        env.update(self.sheet_globals_uncopyable[key[2]])
        # Keep eval globals isolated from module/runtime globals.
        env["__builtins__"] = _get_isolated_builtins()
        cur_sheet = self.ref_parser.Sheet(str(key[2]), self)
        self.cell_meta_gen.set_context(key)
        local = {
            "help": help,
            "cell_single_ref": cur_sheet.cell_single_ref,   "C": cur_sheet.C,
            "cell_range_ref": cur_sheet.cell_range_ref,     "R": cur_sheet.R,
            "global_var": cur_sheet.global_var,             "G": cur_sheet.G,
            "sheet_ref": self.ref_parser.sheet_ref,         "Sh": self.ref_parser.Sh,
            "cell_ref": lambda exp: self.ref_parser.cell_ref(exp, cur_sheet),
                                                            "CR": lambda exp: self.ref_parser.CR(exp, cur_sheet),
            "cell_meta": self.cell_meta_gen.cell_meta,      "CM":  self.cell_meta_gen.CM,
            "RangeOutput": RangeOutput  # Needed for RangeOutput.OFFSET evaluation
        }
        try:
            # Track dependencies during evaluation
            with DependencyTracker.track(key):
                # lstrip() here prevents IndentationError, in case the user puts a space after a "code marker"
                result = PythonEvaluator.exec_then_eval(ref_parsed.lstrip(), env, local)
                if isinstance(result, RangeOutput):
                    PythonEvaluator.range_output_handler(self, result, key)
                if isinstance(result, RangeOutput.OFFSET):
                    result = PythonEvaluator.range_offset_handler(self, result, key)

        except CircularRefError as err:
            # Circular reference detected during evaluation
            result = err

        except AttributeError as err:
            # Attribute Error includes RunTimeError
            result = AttributeError(err)

        except RuntimeError as err:
            result = RuntimeError(err)

        except Exception as err:
            result = err
        #  --- PythonEval END ---  #

        # Change back cell value for evaluation from other cells
        # self.dict_grid[key] = _old_code

        if return_warnings:
            return result, eval_warnings
        return result

    def pop(self, key: Tuple[int, int, int]):
        """pop with cache support

        :param key: Cell key that shall be popped

        """

        # Invalidate cache and remove dependencies
        # Reverse edges are preserved by default to allow proper invalidation
        # when the cell is re-added
        self.dep_graph.remove_cell(key)
        self.smart_cache.invalidate(key)
        self._clear_cell_warnings(key)

        return super().pop(key)

    def recalculate_dirty(self) -> int:
        """Force recalculation of all dirty cells

        Returns the number of cells recalculated

        """

        dirty_cells = list(self.dep_graph.get_all_dirty())

        if not dirty_cells:
            return 0

        # Force recalculation by accessing each dirty cell
        # This will evaluate and cache, clearing dirty flags
        for key in dirty_cells:
            try:
                self.smart_cache.drop(key)
                # Access the cell - this triggers evaluation
                _ = self[key]
            except Exception:
                # Ignore errors during recalculation
                pass

        return len(dirty_cells)

    def _filter_recalc_keys(self, keys) -> list[Tuple[int, int, int]]:
        """Return a unique list of valid cell keys"""

        rows, columns, tables = self.shape
        seen = set()
        filtered = []
        for key in keys:
            if not isinstance(key, tuple) or len(key) != 3:
                continue
            row, column, table = key
            if any(not isinstance(ele, int) for ele in key):
                continue
            if not (0 <= row < rows and 0 <= column < columns
                    and 0 <= table < tables):
                continue
            if key in seen:
                continue
            seen.add(key)
            filtered.append(key)
        return filtered

    def _recalculate_keys(self, keys) -> int:
        """Recalculate a set of cells, ignoring evaluation errors"""

        recalculated = 0
        for key in keys:
            try:
                self.smart_cache.drop(key)
                _ = self[key]
                recalculated += 1
            except Exception:
                pass
        return recalculated

    def recalculate_cell_only(self, key: Tuple[int, int, int]) -> int:
        """Recalculate a single cell and mark dependents dirty if it changed"""

        keys = self._filter_recalc_keys([key])
        if not keys:
            return 0
        key = keys[0]

        old = self.smart_cache.get_raw(key)
        changed = old is SmartCache.INVALID

        try:
            self.smart_cache.drop(key)
            new = self[key]
            if old is not SmartCache.INVALID:
                try:
                    changed = (new != old)
                except Exception:
                    changed = True
        except Exception:
            changed = True

        if changed:
            direct_dependents = self.dep_graph.dependents.get(key, set())
            if self.settings.recalc_mode == "auto":
                for dependent in direct_dependents:
                    self.smart_cache.invalidate(dependent)
            else:
                for dependent in direct_dependents:
                    self.dep_graph.mark_dirty(dependent)

        return 1

    def execute_button_cell(self, key: Tuple[int, int, int]) -> Any:
        """Execute button-cell code with cache and dependency lifecycle handling."""

        if self.cell_attributes[key].button_cell is False:
            return self[key]

        code = self(key)
        if code is None:
            self._set_eval_warnings(key, [])
            return None

        old = self.smart_cache.get_raw(key)
        changed = old is SmartCache.INVALID

        result, eval_warnings = self._eval_cell(key, code, return_warnings=True)
        self.smart_cache.set(key, result)
        self._set_eval_warnings(key, eval_warnings)
        self.dep_graph.clear_dirty(key)

        if old is not SmartCache.INVALID:
            try:
                changed = (result != old)
            except Exception:
                changed = True

        if changed:
            direct_dependents = self.dep_graph.dependents.get(key, set())
            if self.settings.recalc_mode == "auto":
                for dependent in direct_dependents:
                    self.smart_cache.invalidate(dependent)
            else:
                for dependent in direct_dependents:
                    self.dep_graph.mark_dirty(dependent)

        return result

    def recalculate_ancestors(self, key: Tuple[int, int, int]) -> int:
        """Recalculate a cell and all transitive dependencies"""

        deps = self.dep_graph.get_all_dependencies(key)
        keys = self._filter_recalc_keys(list(deps) + [key])
        return self._recalculate_keys(keys)

    def recalculate_children(self, key: Tuple[int, int, int]) -> int:
        """Recalculate a cell and all transitive dependents"""

        dependents = self.dep_graph.get_all_dependents(key)
        keys = self._filter_recalc_keys([key] + list(dependents))
        return self._recalculate_keys(keys)

    def recalculate_all(self) -> int:
        """Recalculate all cells in the workspace"""

        keys = self._filter_recalc_keys(list(self.dict_grid.keys()))
        return self._recalculate_keys(keys)

    def execute_sheet_script(self, current_table) -> Tuple[str, str]:
        """Executes a sheet script and returns result string and error string."""

        if self.safe_mode:
            return '', "Safe mode activated. Code not executed."

        # We need to execute each cell so that assigned globals are updated
        for key in self:
            self[key]

        # Windows exec does not like Windows newline
        self.sheet_scripts[current_table] = self.sheet_scripts[current_table].replace('\r\n', '\n')

        # Create file-like string to capture output
        code_out = io.StringIO()
        code_err = io.StringIO()
        err_msg = io.StringIO()

        # Capture output and errors
        sys.stdout = code_out
        sys.stderr = code_err

        sheet_script = self.sheet_scripts[current_table]
        sheet_globals = {"__builtins__": _get_isolated_builtins()}
        try:
            exec(sheet_script, sheet_globals)

        except Exception:
            exc_info = sys.exc_info()
            user_tb = get_user_codeframe(exc_info[2]) or exc_info[2]
            print_exception(exc_info[0], exc_info[1], user_tb, None, err_msg)
        # Restore stdout and stderr
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

        results = code_out.getvalue()
        errs = code_err.getvalue() + err_msg.getvalue()
        diagnostics = self._script_duplicate_import_warnings(sheet_script)
        diagnostics.extend(self._sheet_global_name_warnings(sheet_globals))
        if diagnostics:
            errs += "".join(msg + "\n" for msg in diagnostics)

        code_out.close()
        code_err.close()

        # Reset cache - clear all since init scripts affect globals
        self.smart_cache.clear()
        self.sheet_globals_copyable[current_table] = dict()
        self.sheet_globals_uncopyable[current_table] = dict()

        for k, v in sheet_globals.items():
            try:
                self.sheet_globals_copyable[current_table][k] = deepcopy(v)
            except TypeError:
                if type(v) not in [types.ModuleType]:
                    errs += (
                        f"WARNING: Sheet global '{k}' ({type(v).__name__}) is not deepcopyable; "
                        f"falling back to by-reference use.\n"
                    )
                self.sheet_globals_uncopyable[current_table][k] = v
        return results, errs

    def execute_macros(self, current_table) -> Tuple[str, str]:
        """Compatibility alias for execute_sheet_script()."""

        return self.execute_sheet_script(current_table)

    def _sorted_keys(self, keys: Iterable[Tuple[int, int, int]],
                     startkey: Tuple[int, int, int],
                     reverse: bool = False) -> Iterable[Tuple[int, int, int]]:
        """Generator that yields sorted keys starting with startkey

        :param keys: Key sequence that is sorted
        :param startkey: First key to be yielded
        :param reverse: Sort direction reversed if True

        """

        def tuple_key(tpl):
            return tpl[::-1]

        if reverse:
            def tuple_cmp(tpl):
                return tpl[::-1] > startkey[::-1]
        else:
            def tuple_cmp(tpl):
                return tpl[::-1] < startkey[::-1]

        searchkeys = sorted(keys, key=tuple_key, reverse=reverse)
        searchpos = sum(1 for _ in filter(tuple_cmp, searchkeys))

        searchkeys = searchkeys[searchpos:] + searchkeys[:searchpos]

        for key in searchkeys:
            yield key

    def string_match(self, datastring: str, findstring: str, word: bool,
                     case: bool, regexp: bool) -> int:
        """Returns position of findstring in datastring or None if not found

        :param datastring: String to be searched
        :param findstring: Search string
        :param word: Search full words only if True
        :param case: Search case sensitively if True
        :param regexp: Regular expression search if True

        """

        if not isinstance(datastring, str):  # Empty cell
            return

        if regexp:
            match = re.search(findstring, datastring)
            if match is None:
                pos = -1
            else:
                pos = match.start()
        else:
            if not case:
                datastring = datastring.lower()
                findstring = findstring.lower()

            if word:
                pos = -1
                matchstring = r'\b' + findstring + r'+\b'
                for match in re.finditer(matchstring, datastring):
                    pos = match.start()
                    break  # find 1st occurrance
            else:
                pos = datastring.find(findstring)

        if pos == -1:
            return None

        return pos

    def findnextmatch(self, startkey: Tuple[int, int, int], find_string: str,
                      up: bool = False, word: bool = False, case: bool = False,
                      regexp: bool = False, results: bool = True
                      ) -> Tuple[int, int, int]:
        """Returns tuple with position of the next match of find_string or None

        :param startkey: Start position of search
        :param find_string: String to be searched for
        :param up: Search up instead of down if True
        :param word: Search full words only if True
        :param case: Search case sensitively if True
        :param regexp: Regular expression search if True
        :param results: Search includes result string if True (slower)

        """

        def is_matching(key, find_string, word, case, regexp):
            code = self(key)
            pos = self.string_match(code, find_string, word, case, regexp)
            if results:
                if pos is not None:
                    return True
                r_str = str(self[key])
                pos = self.string_match(r_str, find_string, word, case, regexp)
            return pos is not None

        # List of keys in sgrid in search order

        table = startkey[2]
        keys = [key for key in self.keys() if key[2] == table]

        for key in self._sorted_keys(keys, startkey, reverse=up):
            try:
                if is_matching(key, find_string, word, case, regexp):
                    return key

            except Exception:
                # re errors are cryptical: sre_constants,...
                pass

# End of class CodeArray
