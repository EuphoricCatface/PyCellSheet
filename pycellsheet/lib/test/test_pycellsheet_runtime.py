# Copyright Seongyong Park (EuphCat)
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
test_pycellsheet_runtime
========================

Focused contract tests for runtime helpers in pycellsheet.py.
"""

import pytest
import random

try:
    from pycellsheet.lib.exceptions import SpillRefError
except ImportError:
    from ..exceptions import SpillRefError

from ..pycellsheet import (
    CELL_META_GENERATOR,
    EmptyCell,
    Formatter,
    HelpText,
    Range,
    RangeOutput,
    safe_deepcopy,
)


class _FakeCellAttributes:
    def __call__(self, key):
        return {"key": key, "marker": "attrs"}


class _FakeCodeArray:
    def __init__(self):
        class _DictGrid:
            sheet_names = ["Sheet1", "Sheet2"]

        self.dict_grid = _DictGrid()
        self.cell_attributes = _FakeCellAttributes()
        self._values = {}

    def __call__(self, key):
        return self._values.get(key, "")


@pytest.fixture(autouse=True)
def reset_cell_meta_singleton():
    CELL_META_GENERATOR._CELL_META_GENERATOR__INSTANCE = None
    yield
    CELL_META_GENERATOR._CELL_META_GENERATOR__INSTANCE = None


def test_helptext_query_formatting():
    assert HelpText((), "x").query == "help()"
    assert HelpText((str,), "x").query == "help(str)"
    assert HelpText((123,), "x").query == "help(123)"
    assert HelpText(("a", "b"), "x").query == "help('a', 'b')"


def test_formatter_display_and_tooltip_contracts():
    assert Formatter.display_formatter(ValueError("bad")) == "ValueError"
    assert Formatter.tooltip_formatter(ValueError("bad")) == "bad"
    spill_err = SpillRefError((0, 0, 0), (0, 1, 0))
    assert Formatter.display_formatter(spill_err) == "#REF!"
    assert "Spill conflict" in Formatter.tooltip_formatter(spill_err)

    help_text = HelpText((str,), "string help")
    assert Formatter.display_formatter(help_text) == "help(str)"
    assert Formatter.tooltip_formatter(help_text) == "string help"

    assert Formatter.display_formatter(RangeOutput(1, [7])) == 7
    assert Formatter.display_formatter(RangeOutput(1, [])) == "EMPTY_RANGEOUTPUT"
    assert Formatter.tooltip_formatter(5) == "int"


def test_range_output_from_range_copies_shape_and_data():
    source = Range("A1", 2, [1, EmptyCell, 3, 4])
    out = RangeOutput.from_range(source)

    assert out.width == 2
    assert out.lst == [1, EmptyCell, 3, 4]


def test_cell_meta_generator_requires_explicit_init():
    with pytest.raises(AssertionError):
        CELL_META_GENERATOR.get_instance()


def test_cell_meta_generator_current_and_explicit_refs():
    code_array = _FakeCodeArray()
    key = (1, 1, 0)
    code_array._values[key] = "B2_code"
    code_array._values[(0, 0, 0)] = "A1_sheet1"
    code_array._values[(1, 1, 1)] = "B2_sheet2"

    generator = CELL_META_GENERATOR.get_instance(code_array)
    generator.set_context(key)

    current = generator.cell_meta()
    assert current.code == "B2_code"
    assert current.attributes["key"] == key

    sheet1_a1 = generator.cell_meta("A1")
    assert sheet1_a1.code == "A1_sheet1"

    sheet2_b2 = generator.cell_meta('"Sheet2"!B2')
    assert sheet2_b2.code == "B2_sheet2"


def test_safe_deepcopy_handles_nested_uncopyables_in_dict():
    value = {"module": random, "nested": {"n": 1}}

    copied = safe_deepcopy(value)

    assert copied is not value
    assert copied["nested"] is not value["nested"]
    assert copied["nested"]["n"] == 1
    assert copied["module"] is random
