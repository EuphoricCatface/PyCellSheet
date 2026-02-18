# -*- coding: utf-8 -*-

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

"""Contract tests for ExpressionParser mode behavior."""

import pytest

from ..pycellsheet import ExpressionParser, PythonCode


@pytest.mark.parametrize(
    "cell, expected",
    [
        ("1 + 2", PythonCode("1 + 2")),
        ("'hello'", PythonCode("'hello'")),
        (">A1", PythonCode(">A1")),
    ],
)
def test_pure_pythonic_mode_contract(cell, expected):
    parser = ExpressionParser()
    parser.set_parser(ExpressionParser.DEFAULT_PARSERS["Pure Pythonic"])

    assert parser.parse(cell) == expected


@pytest.mark.parametrize(
    "cell, expected",
    [
        ("'hello", "hello"),
        ("1 + 2", PythonCode("1 + 2")),
        ("=A1", PythonCode("=A1")),
    ],
)
def test_mixed_mode_contract(cell, expected):
    parser = ExpressionParser()
    parser.set_parser(ExpressionParser.DEFAULT_PARSERS["Mixed"])

    assert parser.parse(cell) == expected


@pytest.mark.parametrize(
    "cell, expected",
    [
        (">1 + 2", PythonCode("1 + 2")),
        ("plain text", "plain text"),
        ("'>A1", ">A1"),
    ],
)
def test_reverse_mixed_mode_contract(cell, expected):
    parser = ExpressionParser()
    parser.set_parser(ExpressionParser.DEFAULT_PARSERS["Reverse Mixed"])

    assert parser.parse(cell) == expected


@pytest.mark.parametrize(
    "cell, expected",
    [
        ("=1 + 2", PythonCode("1 + 2")),
        ("42", 42),
        ("3.5", 3.5),
        ("'42", "42"),
        ("plain text", "plain text"),
    ],
)
def test_pure_spreadsheet_mode_contract(cell, expected):
    parser = ExpressionParser()
    parser.set_parser(ExpressionParser.DEFAULT_PARSERS["Pure Spreadsheet"])

    assert parser.parse(cell) == expected


@pytest.mark.parametrize(
    "cell, expected",
    [
        (None, True),
        ("", True),
        (" ", False),
        ("0", False),
    ],
)
def test_handle_empty_contract(cell, expected):
    parser = ExpressionParser()

    assert parser.handle_empty(cell) is expected
