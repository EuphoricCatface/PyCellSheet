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

"""Contract tests for PythonEvaluator execution semantics."""

import pytest

from ..pycellsheet import PythonCode, PythonEvaluator
from ..compile_cache import CompileCache


def test_exec_then_eval_uses_terminal_expression():
    code = PythonCode("a = 2\na + 3")
    assert PythonEvaluator.exec_then_eval(code, {}, {}) == 5


def test_exec_then_eval_rejects_terminal_return():
    code = PythonCode("x = 7\nreturn x * 2")
    with pytest.raises(SyntaxError, match="Top-level 'return' is not allowed"):
        PythonEvaluator.exec_then_eval(code, {}, {})


def test_exec_then_eval_accepts_terminal_non_expression_statement():
    code = PythonCode("x = 1\nif x:\n    y = 3")
    assert PythonEvaluator.exec_then_eval(code, {}, {}) is None


def test_exec_then_eval_allows_terminal_try_block_without_crash():
    code = PythonCode("x = 1\ntry:\n    x += 2\nexcept Exception:\n    x = -1")
    assert PythonEvaluator.exec_then_eval(code, {}, {}) is None


def test_exec_then_eval_rejects_non_terminal_top_level_return():
    code = PythonCode("return 1\n2 + 3")
    with pytest.raises(SyntaxError, match="Top-level 'return' is not allowed"):
        PythonEvaluator.exec_then_eval(code, {}, {})


def test_exec_then_eval_rejects_top_level_return_inside_try_finally():
    code = PythonCode(
        "try:\n"
        "    return 5\n"
        "finally:\n"
        "    x = 99\n"
    )
    with pytest.raises(SyntaxError, match="Top-level 'return' is not allowed"):
        PythonEvaluator.exec_then_eval(code, {}, {})


def test_exec_then_eval_allows_return_inside_nested_function():
    code = PythonCode("def f(x):\n    return x + 1\nf(4)")
    assert PythonEvaluator.exec_then_eval(code, {}, {}) == 5


def test_exec_then_eval_reuses_compile_cache_artifact():
    cache = CompileCache()
    key = ("table0", "parser_signature", "a = 2\na + 3")
    code = PythonCode("a = 2\na + 3")

    first = PythonEvaluator.exec_then_eval(
        code, {}, {}, compile_cache=cache, cache_key=key
    )
    second = PythonEvaluator.exec_then_eval(
        code, {}, {}, compile_cache=cache, cache_key=key
    )

    assert first == 5
    assert second == 5
    assert len(cache) == 1
