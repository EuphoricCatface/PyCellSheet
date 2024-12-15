#!/usr/bin/env python
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

This file contains interfaces to the native pycs file format.

PycsReader and PycsWriter classed are structured into the following sections:
 * shape
 * code
 * attributes
 * row_heights
 * col_widths
 * macros


**Provides**

 * :func:`wxcolor2rgb`
 * :func:`qt52qt6_fontweights`
 * :func:`qt62qt5_fontweights`
 * :dict:`wx2qt_fontweights`
 * :dict:`wx2qt_fontstyles`
 * :class:`PycsReader`
 * :class:`PycsWriter`

"""

from builtins import str, map, object

import ast
from base64 import b64decode, b85encode
from collections import OrderedDict
from typing import Any, BinaryIO, Callable, Iterable, Tuple

try:
    from pyspread.lib.attrdict import AttrDict
    from pyspread.lib.selection import Selection
    from pyspread.model.model import CellAttribute, CodeArray
except ImportError:
    from lib.attrdict import AttrDict
    from lib.selection import Selection
    from model.model import CellAttribute, CodeArray


def wxcolor2rgb(wxcolor: int) -> Tuple[int, int, int]:
    """Returns red, green, blue for given wxPython binary color value

    :param wxcolor: Color value from wx.Color

    """

    red = wxcolor >> 16
    green = wxcolor - (red << 16) >> 8
    blue = wxcolor - (red << 16) - (green << 8)

    return red, green, blue


def qt52qt6_fontweights(qt5_weight):
    """Approximates the mapping from Qt5 to Qt6 font weight"""

    return int((qt5_weight - 20) * 13.5)


def qt62qt5_fontweights(qt6_weight):
    """Approximates the mapping from Qt6 to Qt5 font weight"""

    return int(qt6_weight / 13.5 + 20)


wx2qt_fontweights = {
    90: 50,  # wx.FONTWEIGHT_NORMAL
    91: 40,  # wx.FONTWEIGHT_LIGHT
    92: 75,  # wx.FONTWEIGHT_BOLD
    93: 87,  # wx.FONTWEIGHT_MAX
    }

wx2qt_fontstyles = {
    90: 0,  # wx.FONTSTYLE_NORMAL
    93: 1,  # wx.FONTSTYLE_ITALIC
    94: 1,  # wx.FONTSTYLE_SLANT
    95: 2,  # wx.FONTSTYLE_MAX
    }


class PycsReader:
    """Reads pycs file into a code_array"""

    def __init__(self, pycs_file: BinaryIO, code_array: CodeArray):
        """
        :param pycs_file: The pycs or pycsu file to be read
        :param code_array: Target code_array

        """

        self.pycs_file = pycs_file
        self.code_array = code_array

        self._section2reader = {
            "[PyCellSheet save file version]\n": self._pycs_version,
            "[shape]\n": self._pycs2shape,
            "[macros]\n": self._pycs2macros,
            "[grid]\n": self._pycs2code,
            "[attributes]\n": self._pycs2attributes,
            "[row_heights]\n": self._pycs2row_heights,
            "[col_widths]\n": self._pycs2col_widths,
        }

        # When converting old versions, cell attributes are required that
        # take place after the cell attribute readout
        self.cell_attributes_postfixes = []

        # Now there are many pages of macros
        self.current_macro = -1
        self.current_macro_remaining = 0

    def __iter__(self):
        """Iterates over self.pycs_file, replacing everything in code_array"""

        state = None

        # Reset pycs_file to start to enable multiple calls of this method
        self.pycs_file.seek(0)

        for line in self.pycs_file:
            line = line.decode("utf8")
            if line in self._section2reader:
                state = line
            elif state is not None:
                self._section2reader[state](line)
            yield line

        # Apply cell attributes post fixes
        for cell_attribute in self.cell_attributes_postfixes:
            self.code_array.cell_attributes.append(cell_attribute)

    # Decorators

    def version_handler(method: Callable) -> Callable:
        """When we need to handle changes for each version, then this will handle it.

        :param method: Method to be replaced in case of old pycs file version

        """

        return method

    # Helpers

    def _split_tidy(self, string: str, maxsplit: int = None) -> str:
        """Rstrips string for \n and splits string for \t

        :param string: String to be rstripped and  split
        :param maxsplit: Maximum number of splits

        """

        if maxsplit is None:
            return string.rstrip("\n").split("\t")
        else:
            return string.rstrip("\n").split("\t", maxsplit)

    def _get_key(self, *keystrings: str) -> Tuple[int, ...]:
        """Returns int key tuple from key string list

        :param keystrings: Strings that contain integers that are key elements

        """

        return tuple(map(int, keystrings))

    # Sections

    def _pycs_version(self, line: str):
        """pycs file version including assertion

        :param line: Pycs file line to be parsed

        """

        self.version = float(line.strip())

        if self.version > 0.0:
            # Abort if file version not supported
            msg = "File version {version} unsupported (> 0.0)."
            raise ValueError(msg.format(version=line.strip()))

    def _pycs2shape(self, line: str):
        """Updates shape in code_array

        :param line: Pycs file line to be parsed

        """

        shape = self._get_key(*self._split_tidy(line))
        if any(dim <= 0 for dim in shape):
            # Abort if any axis is 0 or less
            msg = "Code array has invalid shape {shape}."
            raise ValueError(msg.format(shape=shape))
        self.code_array.shape = shape

    def _pycs2code(self, line: str):
        """Updates code in pycs code_array

        :param line: Pycs file line to be parsed

        """

        row, col, tab, code = self._split_tidy(line, maxsplit=3)
        key = self._get_key(row, col, tab)

        if all(0 <= key[i] < self.code_array.shape[i] for i in range(3)):
            self.code_array.dict_grid[key] = code

    def _attr_convert_1to2(self, key: str, value: Any) -> Tuple[str, Any]:
        """Converts key, value attribute pair from v1.0 to v2.0

        :param key: AttrDict key
        :param value: AttrDict value for key

        """

        color_attrs = ["bordercolor_bottom", "bordercolor_right", "bgcolor",
                       "textcolor"]
        if key in color_attrs:
            return key, wxcolor2rgb(value)

        elif key == "fontweight":
            return key, wx2qt_fontweights[value]

        elif key == "fontstyle":
            return key, wx2qt_fontstyles[value]

        elif key == "markup" and value:
            return "renderer", "markup"

        elif key == "angle" and value < 0:
            return "angle", 360 + value

        elif key == "merge_area":
            # Value in v1.0 None if the cell was merged
            # In v 2.0 this is no longer necessary
            return None, value

        # Update justifiaction and alignment values
        elif key in ["vertical_align", "justification"]:
            just_align_value_tansitions = {
                    "left": "justify_left",
                    "center": "justify_center",
                    "right": "justify_right",
                    "top": "align_top",
                    "middle": "align_center",
                    "bottom": "align_bottom",
            }
            return key, just_align_value_tansitions[value]

        return key, value

    def _pycs2attributes(self, line: str):
        """Updates attributes in code_array

        :param line: Pycs file line to be parsed

        """

        splitline = self._split_tidy(line)

        selection_data = list(map(ast.literal_eval, splitline[:5]))
        selection = Selection(*selection_data)

        tab = int(splitline[5])

        attr_dict = AttrDict()

        for col, ele in enumerate(splitline[6:]):
            if not (col % 2):
                # Odd entries are keys
                key = ast.literal_eval(ele)

            else:
                # Even cols are values
                value = ast.literal_eval(ele)
                attr_dict[key] = value

        if attr_dict:  # Ignore empty attribute settings
            attr = CellAttribute(selection, tab, attr_dict)
            self.code_array.cell_attributes.append(attr)

    def _pycs2row_heights(self, line: str):
        """Updates row_heights in code_array

        :param line: Pycs file line to be parsed

        """

        # Split with maxsplit 3
        split_line = self._split_tidy(line)
        key = row, tab = self._get_key(*split_line[:2])
        height = float(split_line[2])

        shape = self.code_array.shape

        try:
            if row < shape[0] and tab < shape[2]:
                self.code_array.row_heights[key] = height

        except ValueError:
            pass

    def _pycs2col_widths(self, line: str):
        """Updates col_widths in code_array

        :param line: Pycs file line to be parsed

        """

        # Split with maxsplit 3
        split_line = self._split_tidy(line)
        key = col, tab = self._get_key(*split_line[:2])
        width = float(split_line[2])

        shape = self.code_array.shape

        try:
            if col < shape[1] and tab < shape[2]:
                self.code_array.col_widths[key] = width

        except ValueError:
            pass

    def _pycs2macros(self, line: str):
        """Updates macros in code_array

        :param line: Pycs file line to be parsed

        """
        # self.code_array.macros += line

        if self.current_macro_remaining:
            if self.current_macro_remaining == 1:
                # For the last line, we need to exclude the line break,
                #  otherwise the scripts will gain a line break every save&load.
                line = line.rstrip("\r\n")
            self.code_array.macros[self.current_macro] += line
            self.current_macro_remaining -= 1
            return

        sheet_number_str = ""  # later we will use string names
        if not line.startswith("(macro:"):
            raise ValueError("The save file does not follow the macro name conventions")
        for i in iter(line[7:]):
            if i == ")":
                break
            sheet_number_str += i
        new_sheet_number = int(sheet_number_str)
        if self.current_macro + 1 != new_sheet_number:
            raise ValueError("The save file does not follow the macro name conventions")
        self.current_macro = new_sheet_number

        line_count_str = ""
        for i in iter(line[9 + len(sheet_number_str)]):
            line_count_str += i
        self.current_macro_remaining = int(line_count_str)

class PycsWriter(object):
    """Interface between code_array and pycs file data

    Iterating over it yields pycs file lines

    """

    def __init__(self, code_array: CodeArray):
        """
        :param code_array: The code_array object data structure

        """

        self.code_array = code_array

        self.version = 0.0  # NOT STABILIZED YET!

        self._section2writer = OrderedDict([
            ("[PyCellSheet save file version]\n", self._version2pycs),
            ("[shape]\n", self._shape2pycs),
            ("[macros]\n", self._macros2pycs),
            ("[grid]\n", self._code2pycs),
            ("[attributes]\n", self._attributes2pycs),
            ("[row_heights]\n", self._row_heights2pycs),
            ("[col_widths]\n", self._col_widths2pycs),
        ])

    def __iter__(self) -> Iterable[str]:
        """Yields a pycs_file line wise from code_array"""

        for key in self._section2writer:
            yield key
            for line in self._section2writer[key]():
                yield line

    def __len__(self) -> int:
        """Returns how many lines will be written when saving the code_array"""

        lines = 9  # Headers + 1 line version + 1 line shape
        lines += len(self.code_array.dict_grid)
        lines += len(self.code_array.cell_attributes)
        lines += len(self.code_array.dict_grid.row_heights)
        lines += len(self.code_array.dict_grid.col_widths)
        lines += self.code_array.dict_grid.macros.count('\n')

        return lines

    def _version2pycs(self) -> Iterable[str]:
        """Returns pycs file version information in pycs format

        Format: <version>\n

        """

        yield repr(self.version) + "\n"

    def _shape2pycs(self) -> Iterable[str]:
        """Returns shape information in pycs format

        Format: <rows>\t<cols>\t<tabs>\n

        """

        yield u"\t".join(map(str, self.code_array.shape)) + u"\n"

    def _code2pycs(self) -> Iterable[str]:
        """Returns cell code information in pycs format

        Format: <row>\t<col>\t<tab>\t<code>\n

        """

        for key in self.code_array:
            key_str = u"\t".join(repr(ele) for ele in key)
            if self.version <= 1.0:
                code_str = self.code_array(key)
            else:
                code_str = repr(self.code_array(key))
            out_str = key_str + u"\t" + code_str + u"\n"

            yield out_str

    def _attributes2pycs(self) -> Iterable[str]:
        """Returns cell attributes information in pycs format

        Format:
        <selection[0]>\t[...]\t<tab>\t<key>\t<value>\t[...]\n

        """

        # Remove doublettes
        purged_cell_attributes = []
        purged_cell_attributes_keys = []
        for selection, tab, attr_dict in self.code_array.cell_attributes:
            if purged_cell_attributes_keys and \
               (selection, tab) == purged_cell_attributes_keys[-1]:
                purged_cell_attributes[-1][2].update(attr_dict)
            else:
                purged_cell_attributes_keys.append((selection, tab))
                purged_cell_attributes.append([selection, tab, attr_dict])

        for selection, tab, attr_dict in purged_cell_attributes:
            if not attr_dict:
                continue

            sel_list = [selection.block_tl, selection.block_br,
                        selection.rows, selection.columns, selection.cells]

            tab_list = [tab]

            attr_dict_list = []
            for key in attr_dict:
                if key is not None:
                    attr_dict_list.append(key)
                    attr_dict_list.append(attr_dict[key])

            line_list = list(map(repr, sel_list + tab_list + attr_dict_list))

            yield u"\t".join(line_list) + u"\n"

    def _row_heights2pycs(self) -> Iterable[str]:
        """Returns row height information in pycs format

        Format: <row>\t<tab>\t<value>\n

        """

        for row, tab in self.code_array.dict_grid.row_heights:
            if row < self.code_array.shape[0] and \
               tab < self.code_array.shape[2]:
                height = self.code_array.dict_grid.row_heights[(row, tab)]
                height_strings = list(map(repr, [row, tab, height]))
                yield u"\t".join(height_strings) + u"\n"

    def _col_widths2pycs(self) -> Iterable[str]:
        """Returns column width information in pycs format

        Format: <col>\t<tab>\t<value>\n

        """

        for col, tab in self.code_array.dict_grid.col_widths:
            if col < self.code_array.shape[1] and \
               tab < self.code_array.shape[2]:
                width = self.code_array.dict_grid.col_widths[(col, tab)]
                width_strings = list(map(repr, [col, tab, width]))
                yield u"\t".join(width_strings) + u"\n"

    def _macros2pycs(self) -> Iterable[str]:
        """Returns macros information in pycs format

        Format: <macro code line>\n

        """

        macros = self.code_array.dict_grid.macros
        for i, macro in enumerate(macros):
            macro_list = [
                f"(macro:{i}) {macro.count('\n') + 1}",
                macro,
                ""  # To append a linebreak at the end
            ]
            yield str.join("\n", macro_list)
