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
test_model
==========

Unit tests for model.py

"""
from builtins import zip
from builtins import map
from builtins import str
from builtins import range
from builtins import object

import fractions  # Yes, it is required
import math  # Yes, it is required
from os.path import abspath, dirname, join
import sys

import pytest
import numpy

project_path = abspath(join(dirname(__file__) + "/../.."))
sys.path.insert(0, project_path)

from model.model import (KeyValueStore, CellAttributes, DictGrid, DataArray,
                         CodeArray, CellAttribute, DefaultCellAttributeDict,
                         INITSCRIPT_DEFAULT, _get_isolated_builtins,
                         PYCEL_FORMULA_PROMOTION_ENABLED)

from lib.attrdict import AttrDict
from lib.exceptions import SpillRefError
from lib.selection import Selection
from lib.pycellsheet import EmptyCell, ExpressionParser, PythonCode, SpreadSheetCode
sys.path.pop(0)


class Settings:
    """Simulates settings class"""

    timeout = 1000
    recalc_mode = "auto"


def test_get_isolated_builtins_returns_detached_mapping():
    builtins_map = _get_isolated_builtins()
    sentinel_key = "__pycellsheet_test_sentinel__"

    assert isinstance(builtins_map, dict)
    assert sentinel_key not in builtins_map

    builtins_map[sentinel_key] = 1
    fresh_map = _get_isolated_builtins()
    assert sentinel_key not in fresh_map


def test_pycel_formula_promotion_policy_default_off():
    assert PYCEL_FORMULA_PROMOTION_ENABLED is False


class TestCellAttributes(object):
    """Unit tests for CellAttributes"""

    def setup_method(self, method):
        """Creates empty CellAttributes"""

        self.cell_attr = CellAttributes()

        selection_1 = Selection([(2, 2)], [(4, 5)], [55], [55, 66], [(34, 56)])
        selection_2 = Selection([], [], [], [], [(32, 53), (34, 56)])

        ca1 = CellAttribute(selection_1, 0, AttrDict([("testattr", 3)]))
        ca2 = CellAttribute(selection_2, 0, AttrDict([("testattr", 2)]))

        self.cell_attr.append(ca1)
        self.cell_attr.append(ca2)

    def test_append(self):
        """Test append"""

        selection = Selection([], [], [], [], [(23, 12)])
        table = 0
        attr = AttrDict([("angle", 0.2)])

        self.cell_attr.append(CellAttribute(selection, table, attr))

        # Check if 1 item - the actual action has been added
        assert len(self.cell_attr) == 3

    def test_getitem(self):
        """Test __getitem__"""

        assert self.cell_attr[32, 53, 0].testattr == 2
        assert self.cell_attr[2, 2, 0].testattr == 3

    def test_setitem(self):
        """Test __setitem__"""

        selection_3 = Selection([], [], [], [], [(2, 53), (34, 56)])
        ca3 = CellAttribute(selection_3, 0, AttrDict([("testattr", 5)]))
        self.cell_attr[1] = ca3

        assert not self.cell_attr._attr_cache
        assert not self.cell_attr._table_cache

        assert self.cell_attr[2, 53, 0].testattr == 5

    def test_len_table_cache(self):
        """Test _len_table_cache"""

        self.cell_attr[32, 53, 0]

        assert self.cell_attr._len_table_cache() == 2

    def test_update_table_cache(self):
        """Test _update_table_cache"""

        assert self.cell_attr._len_table_cache() == 0
        self.cell_attr._update_table_cache()
        assert self.cell_attr._len_table_cache() == 2

    def test_get_merging_cell(self):
        """Test get_merging_cell"""

        selection_1 = Selection([(2, 2)], [(5, 5)], [], [], [])
        selection_2 = Selection([(3, 2)], [(9, 9)], [], [], [])
        selection_3 = Selection([(2, 2)], [(9, 9)], [], [], [])

        attr_dict_1 = AttrDict([("merge_area", (2, 2, 5, 5))])
        attr_dict_2 = AttrDict([("merge_area", (3, 2, 9, 9))])
        attr_dict_3 = AttrDict([("merge_area", (2, 2, 9, 9))])

        cell_attribute_1 = CellAttribute(selection_1, 0, attr_dict_1)
        cell_attribute_2 = CellAttribute(selection_2, 0, attr_dict_2)
        cell_attribute_3 = CellAttribute(selection_3, 1, attr_dict_3)

        self.cell_attr.append(cell_attribute_1)
        self.cell_attr.append(cell_attribute_2)
        self.cell_attr.append(cell_attribute_3)

        # Cell 1. 1, 0 is not merged
        assert self.cell_attr.get_merging_cell((1, 1, 0)) is None

        # Cell 3. 3, 0 is merged to cell 3, 2, 0
        assert self.cell_attr.get_merging_cell((3, 3, 0)) == (2, 2, 0)

        # Cell 2. 2, 0 is merged to cell 2, 2, 0
        assert self.cell_attr.get_merging_cell((2, 2, 0)) == (2, 2, 0)

    def test_for_table(self):
        """Test for_table"""

        selection_3 = Selection([], [], [], [], [(2, 53), (34, 56)])
        ca3 = CellAttribute(selection_3, 2, AttrDict([("testattr", 5)]))
        self.cell_attr.append(ca3)

        assert len(self.cell_attr.for_table(0)) == 2

        result_cas = CellAttributes()
        result_cas.append(ca3)
        assert self.cell_attr.for_table(2) == result_cas


def test_sheet_script_alias_property():
    """sheet_scripts remain aligned by table resize."""

    code_array = CodeArray((2, 2, 2), Settings())
    code_array.sheet_scripts = ["a = 1", "b = 2"]

    assert code_array.sheet_scripts == ["a = 1", "b = 2"]
    assert code_array.sheet_scripts == ["a = 1", "b = 2"]


def test_execute_sheet_script_alias():
    """execute_sheet_script delegates to execute_macros."""

    code_array = CodeArray((2, 2, 1), Settings())
    code_array.set_exp_parser_mode("mixed")
    code_array.sheet_scripts = ["x = 7"]
    code_array.execute_sheet_script(0)

    assert code_array._eval_cell((0, 0, 0), "x") == 7


def test_sheet_scripts_alias_tracks_shape_resize():
    """sheet_scripts should stay aligned when table count changes."""

    code_array = CodeArray((2, 2, 2), Settings())
    code_array.sheet_scripts = ["a = 1", "b = 2"]

    code_array.shape = (2, 2, 3)
    assert len(code_array.sheet_scripts) == 3
    assert len(code_array.sheet_scripts) == 3

    code_array.shape = (2, 2, 1)
    assert code_array.sheet_scripts == ["a = 1"]
    assert code_array.sheet_scripts == ["a = 1"]


class TestKeyValueStore(object):
    """Unit tests for KeyValueStore"""

    def setup_method(self, method):
        """Creates empty KeyValueStore"""

        self.k_v_store = KeyValueStore()

    def test_missing(self):
        """Test if missing value returns None"""

        key = (1, 2, 3)
        assert self.k_v_store[key] is None

        self.k_v_store[key] = 7

        assert self.k_v_store[key] == 7


class TestDictGrid(object):
    """Unit tests for DictGrid"""

    def setup_method(self, method):
        """Creates empty DictGrid"""

        self.dict_grid = DictGrid((100, 100, 100))

    def test_getitem(self):
        """Unit test for __getitem__"""

        with pytest.raises(IndexError):
            self.dict_grid[100, 0, 0]

        self.dict_grid[(2, 4, 5)] = "Test"
        assert self.dict_grid[(2, 4, 5)] == "Test"

    def test_missing(self):
        """Test if missing value returns None"""

        key = (1, 2, 3)
        assert self.dict_grid[key] is None

        self.dict_grid[key] = 7

        assert self.dict_grid[key] == 7


class TestDataArray(object):
    """Unit tests for DataArray"""

    def setup_method(self, method):
        """Creates empty DataArray"""

        self.data_array = DataArray((100, 100, 100), Settings())

    def test_iter(self):
        """Unit test for __iter__"""

        assert list(iter(self.data_array)) == []

        self.data_array[(1, 2, 3)] = "12"
        self.data_array[(1, 2, 4)] = "13"

        assert sorted(list(iter(self.data_array))) == [(1, 2, 3), (1, 2, 4)]

    def test_keys(self):
        """Unit test for keys"""

        assert list(self.data_array.keys()) == []

        self.data_array[(1, 2, 3)] = "12"
        self.data_array[(1, 2, 4)] = "13"

        assert sorted(self.data_array.keys()) == [(1, 2, 3), (1, 2, 4)]

    def test_pop(self):
        """Unit test for pop"""

        self.data_array[(1, 2, 3)] = "12"
        self.data_array[(1, 2, 4)] = "13"

        assert self.data_array.pop((1, 2, 3)) == "12"

        assert sorted(self.data_array.keys()) == [(1, 2, 4)]

    def test_get_shape(self):
        """Unit test for _get_shape"""

        assert self.data_array.shape == (100, 100, 100)

    def test_set_shape(self):
        """Unit test for _set_shape"""

        self.data_array.shape = (10000, 100, 100)
        assert self.data_array.shape == (10000, 100, 100)

    def test_shape_resize_updates_sheet_scoped_lists_and_names(self):
        data_array = DataArray((2, 2, 2), Settings())
        data_array.sheet_scripts = ["a=1", "b=2"]
        data_array.sheet_scripts_draft = [None, "draft"]
        data_array.sheet_globals_copyable = [{"x": 1}, {"y": 2}]
        data_array.sheet_globals_uncopyable = [{"u": 1}, {"v": 2}]
        data_array.dict_grid.sheet_names = ["Main", "Main"]

        data_array.shape = (2, 2, 4)
        assert len(data_array.sheet_scripts) == 4
        assert len(data_array.sheet_scripts_draft) == 4
        assert len(data_array.sheet_globals_copyable) == 4
        assert len(data_array.sheet_globals_uncopyable) == 4
        assert data_array.dict_grid.sheet_names == ["Main", "Main", "Sheet 2", "Sheet 3"]

        data_array.shape = (2, 2, 1)
        assert data_array.sheet_scripts == ["a=1"]
        assert data_array.sheet_scripts_draft == [None]
        assert data_array.dict_grid.sheet_names == ["Main"]

    def test_data_setter_uses_sheet_scripts(self):
        data_array = DataArray((2, 2, 1), Settings())
        DataArray.data.fset(
            data_array,
            shape=(1, 1, 1),
            grid={(0, 0, 0): "x"},
            row_heights={(0, 0): 10.0},
            col_widths={(0, 0): 12.0},
            sheet_scripts=["script=1"],
            exp_parser_code="return PythonCode(cell)",
        )
        assert data_array.shape == (1, 1, 1)
        assert data_array.sheet_scripts == ["script=1"]
        assert data_array.row_heights[(0, 0)] == 10.0
        assert data_array.col_widths[(0, 0)] == 12.0
        assert data_array.exp_parser_code == "return PythonCode(cell)"

        DataArray.data.fset(data_array, sheet_scripts=["legacy_only=1"])
        assert data_array.sheet_scripts == ["legacy_only=1"]

        snapshot = data_array.data
        assert snapshot["grid"] == {(0, 0, 0): "x"}
        assert snapshot["exp_parser_code"] == "return PythonCode(cell)"

    def test_exp_parser_code_setter_updates_parser_behavior(self):
        data_array = DataArray((2, 2, 1), Settings())
        data_array.exp_parser_code = ExpressionParser.DEFAULT_PARSERS["Pure Spreadsheet"]

        assert data_array.exp_parser_code == ExpressionParser.DEFAULT_PARSERS["Pure Spreadsheet"]
        assert data_array.exp_parser.parse("42") == 42
        assert data_array.exp_parser.parse("=1+1") == SpreadSheetCode("1+1")

    def test_exp_parser_mode_id_contract(self):
        data_array = DataArray((2, 2, 1), Settings())
        assert data_array.exp_parser_mode_id == "pure_spreadsheet"

        data_array.set_exp_parser_mode("mixed")
        assert data_array.exp_parser_mode_id == "mixed"
        assert data_array.exp_parser.parse("=A1+1") == SpreadSheetCode("A1+1")

        data_array.exp_parser_code = "return cell.strip()"
        assert data_array.exp_parser_mode_id is None

    def test_codearray_expression_parser_migration_wrappers(self):
        code_array = CodeArray((2, 2, 1), Settings())
        code_array[0, 0, 0] = "1 + 2"
        code_array[0, 1, 0] = "'hello"

        preview = code_array.preview_expression_parser_migration(
            "mixed", "pure_spreadsheet"
        )
        assert preview.summary["safe_changed"] == 1

        applied = code_array.apply_expression_parser_migration(
            "mixed", "pure_spreadsheet"
        )
        assert applied.summary["safe_changed"] == 1
        assert code_array((0, 0, 0)) == ">1 + 2"
        assert code_array((0, 1, 0)) == "'hello"

    param_get_last_filled_cell = [
        ({(0, 0, 0): "2"}, 0, (0, 0)),
        ({(2, 0, 2): "2"}, 0, (0, 0)),
        ({(2, 0, 2): "2"}, None, (2, 0)),
        ({(2, 0, 2): "2"}, 2, (2, 0)),
        ({(32, 30, 0): "432"}, 0, (32, 30)),
    ]

    @pytest.mark.parametrize("content,table,res", param_get_last_filled_cell)
    def test_get_last_filled_cell(self, content, table, res):
        """Unit test for get_last_filled_cellet_end"""

        for key in content:
            self.data_array[key] = content[key]

        assert self.data_array.get_last_filled_cell(table)[:2] == res

    def test_getstate(self):
        """Unit test for __getstate__ (pickle support)"""

        assert "dict_grid" in self.data_array.__getstate__()

    def test_slicing(self):
        """Unit test for __getitem__ and __setitem__"""

        self.data_array[0, 0, 0] = "'Test'"
        self.data_array[0, 0, 0] = "'Tes'"

        assert self.data_array[0, 0, 0] == "'Tes'"

    def test_slice_access_unsupported(self):
        """Slice-style access is no longer supported."""

        with pytest.raises(NotImplementedError):
            self.data_array[:5, 0, 0]

        with pytest.raises(NotImplementedError):
            self.data_array[:5, :5, 0]

        with pytest.raises(NotImplementedError):
            self.data_array[:5, :5, :5]

    param_adjust_rowcol = [
        ({(0, 0): 3.0}, 0, 2, 0, 0, (0, 0), 3.0),
        ({(0, 0): 3.0}, 0, 2, 0, 0, (2, 0), 3.0),
        ({(0, 0): 3.0}, 0, 1, 1, 0, (1, 0), 3.0),
        ({(0, 0): 3.0}, 0, 1, 1, 0, (0, 1), 0.0),
    ]

    @pytest.mark.parametrize("vals, ins_point, no2ins, axis, tab, target, res",
                             param_adjust_rowcol)
    def test_adjust_rowcol(self, vals, ins_point, no2ins, axis, tab, target,
                           res):
        """Unit test for _adjust_rowcol"""

        if axis == 0:
            __vals = self.data_array.row_heights
        elif axis == 1:
            __vals = self.data_array.col_widths
        else:
            raise ValueError("{} out of 0, 1".format(axis))

        __vals.update(vals)

        self.data_array._adjust_rowcol(ins_point, no2ins, axis, tab)
        assert __vals[target] == res

    def test_set_cell_attributes(self):
        """Unit test for _set_cell_attributes"""

        cell_attributes = self.data_array.cell_attributes

        attr = CellAttribute(Selection([], [], [], [], []), 0,
                             AttrDict([("Test", None)]))
        cell_attributes.clear()
        cell_attributes.append(attr)

        assert self.data_array.cell_attributes == cell_attributes

    def test_default_expression_parser_mode_contract(self):
        """DataArray defaults to Pure Spreadsheet mode."""

        assert self.data_array.exp_parser_code == ExpressionParser.DEFAULT_PARSERS["Pure Spreadsheet"]
        assert self.data_array.exp_parser_mode_id == "pure_spreadsheet"
        assert self.data_array.exp_parser.parse(">1 + 2") == PythonCode("1 + 2")
        assert self.data_array.exp_parser.parse("=SUM(A1:A2)") == SpreadSheetCode("SUM(A1:A2)")

    param_adjust_cell_attributes = [
        (0, 5, 0, (4, 3, 0), (9, 3, 0)),
        (34, 5, 0, (4, 3, 0), (4, 3, 0)),
        (0, 0, 0, (4, 3, 0), (4, 3, 0)),
        (1, 5, 1, (4, 3, 0), (4, 8, 0)),
        (1, 5, 1, (4, 3, 1), (4, 8, 1)),
        (0, -1, 2, (4, 3, 1), None),
        (0, -1, 2, (4, 3, 2), (4, 3, 1)),
    ]

    @pytest.mark.parametrize("inspoint, noins, axis, src, target",
                             param_adjust_cell_attributes)
    def test_adjust_cell_attributes(self, inspoint, noins, axis, src, target):
        """Unit test for _adjust_cell_attributes"""

        row, col, tab = src

        cell_attributes = self.data_array.cell_attributes

        attr_dict = AttrDict([("angle", 0.2)])
        attr = CellAttribute(Selection([], [], [], [], [(row, col)]), tab,
                             attr_dict)
        cell_attributes.clear()
        cell_attributes.append(attr)

        self.data_array._adjust_cell_attributes(inspoint, noins, axis)

        if target is None:
            for key in attr_dict:
                # Should be at default value
                default_ca = DefaultCellAttributeDict()[key]
                assert cell_attributes[src][key] == default_ca
        else:
            for key in attr_dict:
                assert cell_attributes[target][key] == attr_dict[key]

    param_test_insert = [
        ({(2, 3, 0): "42"}, 1, 1, 0, None,
         {(2, 3, 0): None, (3, 3, 0): "42"}),
        ({(0, 0, 0): "0", (0, 0, 2): "2"}, 1, 1, 2, None,
         {(0, 0, 3): "2", (0, 0, 4): None}),
    ]

    @pytest.mark.parametrize("data, inspoint, notoins, axis, tab, res",
                             param_test_insert)
    def test_insert(self, data, inspoint, notoins, axis, tab, res):
        """Unit test for insert operation"""

        self.data_array.dict_grid.update(data)
        self.data_array.insert(inspoint, notoins, axis, tab)

        for key in res:
            assert self.data_array[key] == res[key]

    param_test_delete = [
        ({(2, 3, 4): "42"}, 1, 1, 0, None, {(1, 3, 4): "42"}),
        ({(0, 0, 0): "1"}, 0, 1, 0, 0, {(0, 0, 0): None}),
        ({(0, 0, 1): "1"}, 0, 1, 2, None, {(0, 0, 0): "1"}),
        ({(3, 3, 2): "3"}, 0, 2, 2, None, {(3, 3, 0): "3"}),
        ({(4, 2, 1): "3"}, 2, 1, 1, 1, {(4, 2, 1): None}),
        ({(10, 0, 0): "1"}, 0, 10, 0, 0, {(0, 0, 0): "1"}),
    ]

    @pytest.mark.parametrize("data, delpoint, notodel, axis, tab, res",
                             param_test_delete)
    def test_delete(self, data, delpoint, notodel, axis, tab, res):
        """Tests delete operation"""

        self.data_array.dict_grid.update(data)
        self.data_array.delete(delpoint, notodel, axis, tab)

        for key in res:
            assert self.data_array[key] == res[key]

    def test_delete_error(self):
        """Tests delete operation error"""

        self.data_array[2, 3, 4] = "42"

        try:
            self.data_array.delete(1, 1, 20)
            assert False
        except ValueError:
            pass

    def test_set_row_height(self):
        """Unit test for set_row_height"""

        self.data_array.set_row_height(7, 1, 22.345)
        assert self.data_array.row_heights[7, 1] == 22.345

    def test_set_col_width(self):
        """Unit test for set_col_width"""

        self.data_array.set_col_width(7, 1, 22.345)
        assert self.data_array.col_widths[7, 1] == 22.345


class TestCodeArray(object):
    """Unit tests for CodeArray"""

    def setup_method(self, method):
        """Creates empty DataArray"""

        self.code_array = CodeArray((100, 10, 3), Settings())
        self.code_array.set_exp_parser_mode("mixed")

    param_test_setitem = [
        ({(2, 3, 2): "42"}, {(1, 3, 2): "42"},
         {(1, 3, 2): "42", (2, 3, 2): "42"}),
    ]

    @pytest.mark.parametrize("data, items, res_data", param_test_setitem)
    def test_setitem(self, data, items, res_data):
        """Unit test for __setitem__"""

        self.code_array.dict_grid.update(data)
        for key in items:
            self.code_array[key] = items[key]
        for key in res_data:
            assert res_data[key] == self.code_array(key)

    def test_slicing(self):
        """Slice-style access is no longer supported."""

        with pytest.raises(NotImplementedError):
            self.code_array[:1, :1, :1]

        with pytest.raises(NotImplementedError):
            self.code_array[:1, :1, :1] = "1"

        # Basic evaluation correctness remains intact for non-slice access.
        shape = self.code_array.shape
        x_list = [0, shape[0]-1]
        y_list = [0, shape[1]-1]
        z_list = [0, shape[2]-1]
        for x, y, z in zip(x_list, y_list, z_list):
            assert self.code_array[x, y, z] is EmptyCell

        gridsize = 100
        filled_grid = CodeArray((gridsize, 10, 1), Settings())
        filled_grid.set_exp_parser_mode("mixed")
        for i in [-2**99, 2**99, 0]:
            for j in range(gridsize):
                filled_grid[j, 0, 0] = str(i)
                filled_grid[j, 1, 0] = str(i) + '+' + str(j)
                filled_grid[j, 2, 0] = str(i) + '*' + str(j)

            for j in range(gridsize):
                assert filled_grid[j, 0, 0] == i
                assert filled_grid[j, 1, 0] == i + j
                assert filled_grid[j, 2, 0] == i * j

            for j, funcname in enumerate(['int']):
                filled_grid[j, 3, 0] = funcname + ' (' + str(i) + ')'
                assert filled_grid[j, 3, 0] == eval(funcname + "(" + "i" + ")")
        # Test X, Y, Z
        for i in range(10):
            self.code_array[i, 0, 0] = str(i)
        assert [self.code_array((i, 0, 0)) for i in range(10)] == \
            list(map(str, range(10)))

        assert [self.code_array[i, 0, 0] for i in range(10)] == list(range(10))

        # Test dependency-aware evaluation

        filled_grid[0, 0, 0] = "[1, 2, 3, 4]"
        filled_grid[1, 0, 0] = "sum(C('A1'))"

        assert filled_grid[1, 0, 0] == 10

    def test_make_nested_list(self):
        """Unit test for _make_nested_list"""

        def gen():
            """Nested generator"""

            yield (("Test" for _ in range(2)) for _ in range(2))

        res = self.code_array._make_nested_list(gen())

        assert res == [[["Test" for _ in range(2)] for _ in range(2)]]

    data_eval_cell = [
        ((0, 0, 0), "2 + 4", 6),
        ((1, 0, 0), "S[0, 0, 0]", NameError),
        ((43, 2, 1), "X, Y, Z", NameError),
    ]

    @pytest.mark.parametrize("key, code, res", data_eval_cell)
    def test_eval_cell(self, key, code, res):
        """Unit test for _eval_cell"""

        self.code_array[key] = code
        result = self.code_array._eval_cell(key, code)
        if isinstance(res, type) and issubclass(res, Exception):
            assert isinstance(result, res)
        else:
            assert result == res

    def test_spreadsheet_formula_requires_opt_in(self):
        self.code_array[0, 0, 0] = "=1+2"
        result = self.code_array[0, 0, 0]

        assert isinstance(result, RuntimeError)
        assert "disabled" in str(result).lower()

    def test_spreadsheet_formula_evaluates_with_pycel_when_opted_in(self):
        excel_formula = self.code_array._eval_spreadsheet_code.__globals__["ExcelFormula"]
        if excel_formula is None:
            pytest.skip("pycel is not installed in this environment")

        self.code_array.set_pycel_formula_opt_in(True)
        self.code_array[0, 0, 0] = "=1+2"
        result = self.code_array[0, 0, 0]

        assert not isinstance(result, RuntimeError)
        assert not isinstance(result, ImportError)
        if isinstance(result, Exception):
            assert "compile expression" in str(result).lower()

    def test_spreadsheet_formula_reports_missing_pycel(self, monkeypatch):
        self.code_array.set_pycel_formula_opt_in(True)
        monkeypatch.setitem(self.code_array._eval_spreadsheet_code.__globals__,
                            "ExcelFormula", None)
        self.code_array[0, 0, 0] = "=1+2"
        result = self.code_array[0, 0, 0]

        assert isinstance(result, ImportError)
        assert "pycel" in str(result).lower()

    def test_pure_spreadsheet_python_marker_allows_space_after_token(self):
        self.code_array.set_exp_parser_mode("pure_spreadsheet")
        self.code_array[0, 0, 0] = "> 123"

        assert self.code_array[0, 0, 0] == 123

    def test_globals_result_with_modules_survives_cache_hit(self):
        """globals() result should not crash on cached deepcopy fallback."""

        self.code_array.sheet_scripts[0] = INITSCRIPT_DEFAULT
        self.code_array.execute_sheet_script(0)
        self.code_array[0, 0, 0] = "globals()"

        first = self.code_array[0, 0, 0]
        second = self.code_array[0, 0, 0]

        assert isinstance(first, dict)
        assert isinstance(second, dict)
        assert "random_" in second

    def test_locals_result_with_modules_survives_cache_hit(self):
        """locals() result should not crash on cached deepcopy fallback."""

        self.code_array.sheet_scripts[0] = INITSCRIPT_DEFAULT
        self.code_array.execute_sheet_script(0)
        self.code_array[0, 0, 0] = "locals()"

        first = self.code_array[0, 0, 0]
        second = self.code_array[0, 0, 0]

        assert isinstance(first, dict)
        assert isinstance(second, dict)
        assert "C" in second

    def test_legacy_slice_replacement_helpers(self):
        """Reference helpers replace legacy slice-style references."""

        self.code_array[0, 0, 0] = "2"
        self.code_array[1, 0, 0] = "C('A1')"
        self.code_array[2, 0, 0] = "sum(R('A1', 'A2').flatten())"
        self.code_array[3, 0, 0] = "CR('A1')"
        self.code_array[4, 0, 0] = "S[0, 0, 0]"

        assert self.code_array._eval_cell((1, 0, 0), self.code_array((1, 0, 0))) == 2
        assert self.code_array._eval_cell((2, 0, 0), self.code_array((2, 0, 0))) == 4
        assert self.code_array._eval_cell((3, 0, 0), self.code_array((3, 0, 0))) == 2

        legacy = self.code_array._eval_cell((4, 0, 0), self.code_array((4, 0, 0)))
        assert isinstance(legacy, NameError)

    def test_parser_emptycell_warning_contract(self):
        """Parser returning EmptyCell for non-empty code should warn."""

        self.code_array.exp_parser_code = "return EmptyCell"
        self.code_array[0, 0, 0] = "nonempty"

        result = self.code_array[0, 0, 0]
        warnings = self.code_array.get_cell_warnings((0, 0, 0))

        assert result is EmptyCell
        assert any("Expression parser returned EmptyCell" in warning for warning in warnings)

    def test_execute_button_cell_updates_cache_and_dependents(self):
        """Button cell execution should not bypass cache/dirty lifecycle."""

        key = (0, 0, 0)
        dependent = (1, 0, 0)
        self.code_array[key] = "2"

        selection = Selection([(0, 0)], [(0, 0)], [], [], [])
        attr = AttrDict([("button_cell", "Run")])
        self.code_array.cell_attributes.append(CellAttribute(selection, 0, attr))

        self.code_array.smart_cache.set(key, 1)
        self.code_array.smart_cache.set(dependent, 99)
        self.code_array.dep_graph.dependents[key] = {dependent}

        result = self.code_array.execute_button_cell(key)

        assert result == 2
        assert self.code_array.smart_cache.get_raw(key) == 2
        assert self.code_array.dep_graph.is_dirty(dependent)

    def test_execute_sheet_script(self):
        """Unit test for execute_sheet_script"""

        self.code_array.sheet_scripts = ["a = 5\ndef f(x): return x ** 2"]
        self.code_array.execute_sheet_script(0)
        assert self.code_array._eval_cell((0, 0, 0), "a") == 5
        assert self.code_array._eval_cell((0, 0, 0), "f(2)") == 4

    def test_init_script_default_has_deterministic_random_policy(self):
        assert "from pycellsheet.lib.spreadsheet import *" in INITSCRIPT_DEFAULT
        assert "RANDOM_SEED = 0" in INITSCRIPT_DEFAULT
        assert "random = random_.Random(RANDOM_SEED)" in INITSCRIPT_DEFAULT

    def test_sheet_script_random_seed_is_deterministic_after_reapply(self):
        self.code_array.sheet_scripts = [
            "import random as random_\n"
            "RANDOM_SEED = 123\n"
            "random = random_.Random(RANDOM_SEED)\n"
        ]
        self.code_array[0, 0, 0] = "random.random()"

        self.code_array.execute_sheet_script(0)
        first = self.code_array[0, 0, 0]
        self.code_array.execute_sheet_script(0)
        second = self.code_array[0, 0, 0]

        assert first == second

    def test_execute_sheet_script_warns_on_cell_like_globals(self):
        self.code_array.sheet_scripts = ["A1 = 10\nvalue = 3"]
        _, errs = self.code_array.execute_sheet_script(0)
        assert "looks like a cell reference" in errs

    def test_execute_sheet_script_warns_on_duplicate_import_bindings(self):
        self.code_array.sheet_scripts = ["import math\nimport random as math"]
        _, errs = self.code_array.execute_sheet_script(0)
        assert "Duplicate import binding 'math'" in errs

    def test_range_output_spill_conflict_returns_spill_ref_error(self):
        self.code_array.set_exp_parser_mode("pure_spreadsheet")
        self.code_array[0, 1, 0] = ">99"
        self.code_array[0, 0, 0] = ">RangeOutput(2, [1, 2])"

        result = self.code_array[0, 0, 0]
        assert isinstance(result, SpillRefError)

    def test_execute_sheet_script_restores_streams_on_base_exception(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        self.code_array.sheet_scripts = ["raise KeyboardInterrupt('stop')"]

        with pytest.raises(KeyboardInterrupt):
            self.code_array.execute_sheet_script(0)

        assert sys.stdout is old_stdout
        assert sys.stderr is old_stderr

    def test_sorted_keys(self):
        """Unit test for _sorted_keys"""

        code_array = self.code_array

        keys = [(1, 0, 0), (2, 0, 0), (0, 1, 0), (0, 99, 0), (0, 0, 0),
                (0, 0, 99), (1, 2, 3)]
        sorted_keys = [(0, 1, 0), (0, 99, 0), (1, 2, 3), (0, 0, 99), (0, 0, 0),
                       (1, 0, 0), (2, 0, 0)]
        rev_sorted_keys = [(0, 1, 0), (2, 0, 0), (1, 0, 0), (0, 0, 0),
                           (0, 0, 99), (1, 2, 3), (0, 99, 0)]

        sort_gen = code_array._sorted_keys(keys, (0, 1, 0))
        for result, expected_result in zip(sort_gen, sorted_keys):
            assert result == expected_result

        rev_sort_gen = code_array._sorted_keys(keys, (0, 3, 0), reverse=True)
        for result, expected_result in zip(rev_sort_gen, rev_sorted_keys):
            assert result == expected_result

    def test_string_match(self):
        """Tests creation of string_match"""

        code_array = self.code_array

        test_strings = [
            "", "Hello", " Hello", "Hello ", " Hello ", "Hello\n",
            "THelloT", " HelloT", "THello ", "hello", "HELLO", "sd"
        ]

        search_string = "Hello"

        # Normal search
        flags = False, False, False
        results = [None, 0, 1, 0, 1, 0, 1, 1, 1, 0, 0, None]
        for test_string, result in zip(test_strings, results):
            res = code_array.string_match(test_string, search_string, *flags)
            assert res == result

        # Case sensitive
        flags = False, True, False
        results = [None, 0, 1, 0, 1, 0, 1, 1, 1, None, None, None]
        for test_string, result in zip(test_strings, results):
            res = code_array.string_match(test_string, search_string, *flags)
            assert res == result

        # Word search
        flags = True, False, False
        results = [None, 0, 1, 0, 1, 0, None, None, None, 0, 0, None]
        for test_string, result in zip(test_strings, results):
            res = code_array.string_match(test_string, search_string, *flags)
            assert res == result

    def test_findnextmatch(self):
        """Find method test"""

        code_array = self.code_array

        for i in range(100):
            code_array[i, 0, 0] = str(i)

        assert code_array[3, 0, 0] == 3
        assert code_array.findnextmatch((0, 0, 0), "3", False) == (3, 0, 0)
        assert code_array.findnextmatch((0, 0, 0), "99", True) == (99, 0, 0)
