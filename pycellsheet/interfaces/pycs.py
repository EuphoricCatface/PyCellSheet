#!/usr/bin/env python
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

This file contains interfaces to the native pycs file format.

PycsReader and PycsWriter classed are structured into the following sections:
 * shape
 * code
 * attributes
 * row_heights
 * col_widths
 * sheet_scripts


**Provides**

 * :func:`wxcolor2rgb`
 * :func:`qt52qt6_fontweights`
 * :func:`qt62qt5_fontweights`
 * ``dict``: ``wx2qt_fontweights``
 * ``dict``: ``wx2qt_fontstyles``
 * :class:`PycsReader`
 * :class:`PycsWriter`

"""

from builtins import str, map, object

import ast
from base64 import b64decode, b85encode
from collections import OrderedDict
import re
from typing import Any, BinaryIO, Callable, Iterable, Tuple

try:
    from pycellsheet.lib.attrdict import AttrDict
    from pycellsheet.lib.sheet_name import sanitize_loaded_sheet_name, generate_unique_sheet_name
    from pycellsheet.lib.selection import Selection
    from pycellsheet.model.model import CellAttribute, CodeArray
except ImportError:
    from lib.attrdict import AttrDict
    from lib.sheet_name import sanitize_loaded_sheet_name, generate_unique_sheet_name
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
            "[sheet_names]\n": self._pycs2sheet_names,
            "[parser_specs]\n": self._pycs2parser_specs,
            "[parser_settings]\n": self._pycs2parser_settings,
            "[sheet_scripts]\n": self._pycs2sheet_scripts,
            "[grid]\n": self._pycs2code,
            "[attributes]\n": self._pycs2attributes,
            "[row_heights]\n": self._pycs2row_heights,
            "[col_widths]\n": self._pycs2col_widths,
        }

        # When converting old versions, cell attributes are required that
        # take place after the cell attribute readout
        self.cell_attributes_postfixes = []

        # [sheet_scripts] section stores per-sheet script blocks.
        self.current_sheet_script = -1
        self.current_sheet_script_remaining = 0
        self._sheet_script_header_re = re.compile(r"^\(sheet_script:(.+)\)\s+([0-9]+)$")
        self.current_parser_spec_id = None
        self.current_parser_spec_remaining = 0
        self.current_parser_spec_lines = []
        self._parser_spec_header_re = re.compile(r"^\(parser_spec:(.+)\)\s+([0-9]+)$")
        self._loaded_parser_specs = []
        self._pending_active_parser_id = None

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

        # Ensure sheet_names length always matches table count after load.
        self._finalize_sheet_names()
        if self._loaded_parser_specs:
            self.code_array.parser_specs = self._loaded_parser_specs
        if self._pending_active_parser_id is not None:
            self.code_array.active_parser_id = self._pending_active_parser_id

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

    def _pycs2sheet_names(self, line: str):
        """Updates sheet names in code_array

        :param line: Pycs file line to be parsed (one sheet name per line)

        """

        raw_sheet_name = line.rstrip('\n')
        # Initialize sheet_names list if first call in this section
        if not hasattr(self, '_sheet_names_initialized'):
            if hasattr(self.code_array.dict_grid, 'sheet_names'):
                self.code_array.dict_grid.sheet_names = []
            self._sheet_names_initialized = True
        # Append to sheet_names list (order matters)
        if hasattr(self.code_array.dict_grid, 'sheet_names'):
            sheet_names = self.code_array.dict_grid.sheet_names
            # Ignore extra names beyond table count.
            if len(sheet_names) >= self.code_array.shape[2]:
                return
            sheet_name = sanitize_loaded_sheet_name(
                raw_sheet_name,
                sheet_names,
                fallback_index=len(sheet_names),
            )
            sheet_names.append(sheet_name)

    def _finalize_sheet_names(self):
        """Normalize sheet names count and fill missing names deterministically."""

        sheet_names = getattr(self.code_array.dict_grid, 'sheet_names', None)
        if sheet_names is None:
            return

        expected_count = self.code_array.shape[2]
        normalized = []

        for i in range(expected_count):
            raw = sheet_names[i] if i < len(sheet_names) else ""
            normalized.append(
                sanitize_loaded_sheet_name(raw, normalized, fallback_index=i)
            )

        self.code_array.dict_grid.sheet_names = normalized

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

    def _pycs2sheet_scripts(self, line: str):
        """Updates sheet scripts in code_array from [sheet_scripts] section.

        :param line: Pycs file line to be parsed

        """
        if self.current_sheet_script_remaining:
            if self.current_sheet_script_remaining == 1:
                # For the last line, we need to exclude the line break,
#  otherwise the scripts will gain a line break every save&load.
                line = line.rstrip("\r\n")
            self.code_array.sheet_scripts[self.current_sheet_script] += line
            self.current_sheet_script_remaining -= 1
            return

        header_line = line.rstrip("\r\n")
        header_match = self._sheet_script_header_re.fullmatch(header_line)
        if header_match is None:
            raise ValueError("The save file does not follow sheet script header conventions")
        raw_sheet_identifier, line_count_str = header_match.groups()

        try:
            parsed_identifier = ast.literal_eval(raw_sheet_identifier)
        except Exception:
            parsed_identifier = raw_sheet_identifier

        # v0.5+ contract: only named sheet_script headers are accepted.
        if isinstance(parsed_identifier, int):
            raise ValueError(
                "Numeric sheet_script headers are no longer supported in v0.5+. "
                "Use named headers, e.g. (sheet_script:'Sheet 0') N."
            )

        sheet_names = getattr(self.code_array.dict_grid, 'sheet_names', None)
        sheet_identifier = str(parsed_identifier)
        if not sheet_names or sheet_identifier not in sheet_names:
            raise ValueError(
                f"Unknown sheet name in sheet_script header: {sheet_identifier!r}. "
                "Ensure [sheet_names] includes the referenced sheet before [sheet_scripts]."
            )
        new_sheet_number = sheet_names.index(sheet_identifier)

        self.current_sheet_script = new_sheet_number

        self.current_sheet_script_remaining = int(line_count_str)

    def _pycs2parser_settings(self, line: str):
        """Updates parser settings from [parser_settings] section."""

        key, value_repr = self._split_tidy(line, maxsplit=1)
        value = ast.literal_eval(value_repr)
        if key == "exp_parser_code":
            self.code_array.exp_parser_code = value
        elif key == "active_parser_id":
            self._pending_active_parser_id = str(value)
        else:
            raise ValueError(
                f"Unknown parser_settings key: {key}. "
                "Supported keys in v0.6+: exp_parser_code, active_parser_id."
            )

    def _pycs2parser_specs(self, line: str):
        """Updates parser specs from [parser_specs] section."""

        if self.current_parser_spec_remaining:
            if self.current_parser_spec_remaining == 1:
                line = line.rstrip("\r\n")
            self.current_parser_spec_lines.append(line)
            self.current_parser_spec_remaining -= 1
            if self.current_parser_spec_remaining == 0:
                payload = "\n".join(self.current_parser_spec_lines)
                parsed = ast.literal_eval(payload)
                if not isinstance(parsed, dict):
                    raise ValueError("parser_spec payload must be a dict literal.")
                parsed["id"] = self.current_parser_spec_id
                self._loaded_parser_specs.append(parsed)
                self.current_parser_spec_lines = []
            return

        header_line = line.rstrip("\r\n")
        header_match = self._parser_spec_header_re.fullmatch(header_line)
        if header_match is None:
            raise ValueError("The save file does not follow parser_spec header conventions")
        raw_parser_identifier, line_count_str = header_match.groups()
        try:
            parsed_identifier = ast.literal_eval(raw_parser_identifier)
        except Exception:
            parsed_identifier = raw_parser_identifier

        self.current_parser_spec_id = str(parsed_identifier)
        self.current_parser_spec_remaining = int(line_count_str)

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
            ("[sheet_names]\n", self._sheet_names2pycs),
            ("[parser_specs]\n", self._parser_specs2pycs),
            ("[parser_settings]\n", self._parser_settings2pycs),
            ("[sheet_scripts]\n", self._sheet_scripts2pycs),
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

    def _normalized_sheet_names(self) -> list[str]:
        """Return valid, unique sheet names aligned to current table count."""

        table_count = self.code_array.shape[2]
        raw_names = getattr(self.code_array.dict_grid, 'sheet_names', []) or []

        normalized = []
        for i in range(table_count):
            raw = raw_names[i] if i < len(raw_names) else ""
            normalized.append(
                generate_unique_sheet_name(raw, normalized, fallback_index=i)
            )
        return normalized

    def __len__(self) -> int:
        """Returns how many lines will be written when saving the code_array"""

        # Keep progress lengths exact by counting the same stream __iter__ emits.
        return sum(1 for _ in self)

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

    def _sheet_names2pycs(self) -> Iterable[str]:
        """Returns sheet names in pycs format

        Format: one sheet name per line

        """

        for name in self._normalized_sheet_names():
            yield name + u"\n"

    def _parser_settings2pycs(self) -> Iterable[str]:
        """Returns parser settings information in pycs format."""

        active_parser_id = getattr(self.code_array, "active_parser_id", None)
        if active_parser_id is not None:
            yield f"active_parser_id\t{active_parser_id!r}\n"
        yield f"exp_parser_code\t{self.code_array.exp_parser_code!r}\n"

    def _parser_specs2pycs(self) -> Iterable[str]:
        """Returns parser spec blocks in pycs format."""

        parser_specs = getattr(self.code_array, "parser_specs", None) or []
        for parser_spec in parser_specs:
            parser_id = str(parser_spec.get("id"))
            payload = dict(parser_spec)
            payload.pop("id", None)
            payload_str = repr(payload)
            payload_count = payload_str.count("\n") + 1
            block = [
                f"(parser_spec:{parser_id!r}) {payload_count}",
                payload_str,
                "",
            ]
            yield "\n".join(block)

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

    def _sheet_scripts2pycs(self) -> Iterable[str]:
        """Returns sheet script information in pycs format

        Format: <sheet script code line>\n

        """

        sheet_scripts = self.code_array.dict_grid.sheet_scripts
        sheet_names = self._normalized_sheet_names()
        for i, sheet_script in enumerate(sheet_scripts):
            # Use sheet name if available, otherwise fall back to index
            sheet_identifier = sheet_names[i] if i < len(sheet_names) else str(i)
            sheet_script_count = sheet_script.count('\n') + 1
            macro_list = [
                f"(sheet_script:{sheet_identifier!r}) {sheet_script_count}",
                sheet_script,
                ""  # To append a linebreak at the end
            ]
            yield str.join("\n", macro_list)
