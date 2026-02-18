#!/usr/bin/env python
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


"""

**Integration tests for dependency graph + smart cache**

Tests the full pipeline from cell editing to dependency tracking
to cache invalidation.

"""

from os.path import abspath, dirname, join
import sys

import pytest

project_path = abspath(join(dirname(__file__) + "/../.."))
sys.path.insert(0, project_path)

from model.model import CodeArray
from lib.exceptions import CircularRefError
from lib.smart_cache import SmartCache

sys.path.pop(0)


class Settings:
    """Simulates settings class"""
    timeout = 1000
    recalc_mode = "auto"


@pytest.fixture
def code_array():
    """Create a CodeArray for testing"""
    return CodeArray((100, 100, 3), Settings())


# --- Basic Dependency Tracking Tests ---

def test_simple_dependency_tracked(code_array):
    """Test that C() calls record dependencies"""

    # A1 = 10
    code_array[0, 0, 0] = "10"
    assert code_array[0, 0, 0] == 10

    # A2 = C("A1") + 5
    code_array[1, 0, 0] = 'C("A1") + 5'
    assert code_array[1, 0, 0] == 15

    # Check dependency was recorded
    deps = code_array.dep_graph.dependencies[(1, 0, 0)]
    assert (0, 0, 0) in deps


def test_multiple_dependencies_tracked(code_array):
    """Test that multiple C() calls record all dependencies"""

    # A1 = 10, B1 = 20
    code_array[0, 0, 0] = "10"
    code_array[0, 1, 0] = "20"

    # A2 = C("A1") + C("B1")
    code_array[1, 0, 0] = 'C("A1") + C("B1")'
    assert code_array[1, 0, 0] == 30

    # Check both dependencies were recorded
    deps = code_array.dep_graph.dependencies[(1, 0, 0)]
    assert (0, 0, 0) in deps
    assert (0, 1, 0) in deps


def test_range_dependencies_tracked(code_array):
    """Test that R() calls record dependencies for all cells in range"""

    # A1=1, A2=2, A3=3
    code_array[0, 0, 0] = "1"
    code_array[1, 0, 0] = "2"
    code_array[2, 0, 0] = "3"

    # A4 = SUM(R("A1", "A3")) -> manual flattening, to avoid pulling in an unnecessary dependency
    code_array[3, 0, 0] = 'sum(R("A1", "A3").flatten())'
    result = code_array[3, 0, 0]
    assert result == 6

    # Check all three cells are dependencies
    deps = code_array.dep_graph.dependencies[(3, 0, 0)]
    assert (0, 0, 0) in deps
    assert (1, 0, 0) in deps
    assert (2, 0, 0) in deps


def test_cross_sheet_dependency_tracked(code_array):
    """Test that Sh() calls record cross-sheet dependencies"""

    # Sheet 0, A1 = 10
    code_array[0, 0, 0] = "10"

    # Sheet 1, A1 = Sh("0").C("A1") * 2
    code_array[0, 0, 1] = 'Sh("0").C("A1") * 2'
    assert code_array[0, 0, 1] == 20

    # Check cross-sheet dependency was recorded
    deps = code_array.dep_graph.dependencies[(0, 0, 1)]
    assert (0, 0, 0) in deps


# --- Cache Invalidation Tests ---

def test_smart_cache_invalidation_simple(code_array):
    """Test that editing a cell invalidates dependent caches"""

    # A1 = 10
    code_array[0, 0, 0] = "10"

    # A2 = C("A1") + 5
    code_array[1, 0, 0] = 'C("A1") + 5'
    assert code_array[1, 0, 0] == 15

    # Edit A1
    code_array[0, 0, 0] = "20"

    # A2 should be recalculated (not cached old value)
    assert code_array[1, 0, 0] == 25


def test_smart_cache_invalidation_chain(code_array):
    """Test that invalidation propagates through dependency chain"""

    # A1 = 1
    code_array[0, 0, 0] = "1"

    # A2 = C("A1") + 1
    code_array[1, 0, 0] = 'C("A1") + 1'
    assert code_array[1, 0, 0] == 2

    # A3 = C("A2") + 1
    code_array[2, 0, 0] = 'C("A2") + 1'
    assert code_array[2, 0, 0] == 3

    # A4 = C("A3") + 1
    code_array[3, 0, 0] = 'C("A3") + 1'
    assert code_array[3, 0, 0] == 4

    # Edit A1 - should invalidate entire chain
    code_array[0, 0, 0] = "10"

    # Check that all dependent cells are recalculated
    assert code_array[0, 0, 0] == 10
    assert code_array[1, 0, 0] == 11
    assert code_array[2, 0, 0] == 12
    assert code_array[3, 0, 0] == 13


def test_smart_cache_unrelated_cells_not_invalidated(code_array):
    """Test that unrelated cells keep their cache"""

    # A1 = 100
    code_array[0, 0, 0] = "100"

    # B1 = 200
    code_array[0, 1, 0] = "200"

    # A2 = C("A1") * 2
    code_array[1, 0, 0] = 'C("A1") * 2'
    assert code_array[1, 0, 0] == 200

    # Cache A2's result
    code_array.smart_cache.set((1, 0, 0), 200)

    # Edit B1 (unrelated to A2)
    code_array[0, 1, 0] = "999"

    # A2's cache should still be valid
    assert code_array.smart_cache.is_valid((1, 0, 0))
    cached = code_array.smart_cache.get((1, 0, 0))
    assert cached == 200


def test_smart_cache_diamond_pattern(code_array):
    """Test cache invalidation with diamond dependency pattern"""

    # A1 = 1
    code_array[0, 0, 0] = "1"

    # A2 = C("A1") + 1
    code_array[1, 0, 0] = 'C("A1") + 1'
    assert code_array[1, 0, 0] == 2

    # A3 = C("A1") + 2
    code_array[2, 0, 0] = 'C("A1") + 2'
    assert code_array[2, 0, 0] == 3

    # A4 = C("A2") + C("A3")  (depends on both A2 and A3)
    code_array[3, 0, 0] = 'C("A2") + C("A3")'
    assert code_array[3, 0, 0] == 5

    # Edit A1 - should invalidate A2, A3, and A4
    code_array[0, 0, 0] = "10"

    assert code_array[0, 0, 0] == 10
    assert code_array[1, 0, 0] == 11  # A2 recalculated
    assert code_array[2, 0, 0] == 12  # A3 recalculated
    assert code_array[3, 0, 0] == 23  # A4 recalculated


# --- Circular Reference Detection Tests ---

def test_circular_reference_self(code_array):
    """Test detecting self-reference circular dependency"""

    # A1 = C("A1") + 1
    code_array[0, 0, 0] = 'C("A1") + 1'
    result = code_array[0, 0, 0]

    # Should return CircularRefError
    assert isinstance(result, CircularRefError)
    assert "Circular reference" in str(result)


def test_circular_reference_two_cells(code_array):
    """Test detecting two-cell circular dependency"""

    # A1 = C("A2") + 1
    code_array[0, 0, 0] = 'C("A2") + 1'

    # A2 = C("A1") + 1  (creates cycle)
    code_array[1, 0, 0] = 'C("A1") + 1'

    # Both cells should detect the cycle
    result1 = code_array[0, 0, 0]
    result2 = code_array[1, 0, 0]

    assert isinstance(result1, CircularRefError) or isinstance(result2, CircularRefError)


def test_circular_reference_complex(code_array):
    """Test detecting complex multi-cell circular dependency"""

    # A1 = 1
    code_array[0, 0, 0] = "1"

    # A2 = C("A1") + 1
    code_array[1, 0, 0] = 'C("A1") + 1'
    assert code_array[1, 0, 0] == 2

    # A3 = C("A2") + 1
    code_array[2, 0, 0] = 'C("A2") + 1'
    assert code_array[2, 0, 0] == 3

    # A1 = C("A3") + 1  (creates cycle: A1 -> A2 -> A3 -> A1)
    code_array[0, 0, 0] = 'C("A3") + 1'

    # Should detect cycle
    result = code_array[0, 0, 0]
    assert isinstance(result, CircularRefError)


# --- Performance Tests ---

def test_cache_hit_avoids_recomputation(code_array):
    """Test that cache hits avoid re-evaluation"""

    # A1 = 10
    code_array[0, 0, 0] = "10"

    # A2 = C("A1") * 2
    code_array[1, 0, 0] = 'C("A1") * 2'
    first_result = code_array[1, 0, 0]
    assert first_result == 20

    # Access A2 again - should be cached
    second_result = code_array[1, 0, 0]
    assert second_result == 20

    # Verify cache was used (both results should be same value)
    assert first_result == second_result


def test_dependency_removal_on_edit(code_array):
    """Test that editing a cell removes its old dependencies"""

    # A1 = 10, B1 = 20
    code_array[0, 0, 0] = "10"
    code_array[0, 1, 0] = "20"

    # A2 = C("A1") + 5
    code_array[1, 0, 0] = 'C("A1") + 5'
    assert code_array[1, 0, 0] == 15

    # Check A2 depends on A1
    assert (0, 0, 0) in code_array.dep_graph.dependencies[(1, 0, 0)]

    # Edit A2 to depend on B1 instead
    code_array[1, 0, 0] = 'C("B1") + 5'
    assert code_array[1, 0, 0] == 25

    # Check A2 now depends on B1, not A1
    deps = code_array.dep_graph.dependencies[(1, 0, 0)]
    assert (0, 1, 0) in deps
    # Old dependency on A1 should be gone
    assert (0, 0, 0) not in deps


# --- Edge Cases ---

def test_empty_cell_dependency(code_array):
    """Test that referencing an empty cell works"""

    # A1 is empty
    # A2 = C("A1") + 1
    code_array[1, 0, 0] = 'C("A1") + 1'

    # Should handle EmptyCell (which acts as 0)
    result = code_array[1, 0, 0]
    assert result == 1


def test_none_vs_empty_caching(code_array):
    """Test that None results are cached correctly (not confused with INVALID)"""

    # A1 = None
    code_array[0, 0, 0] = "None"
    result = code_array[0, 0, 0]
    assert result is None

    # Should be cached
    cached = code_array.smart_cache.get((0, 0, 0))
    assert cached is None
    assert cached is not SmartCache.INVALID


def test_branching_formula_rebuilds_dependencies(code_array):
    """Dependencies should track only the currently executed branch."""

    # C1 controls branch
    code_array[0, 2, 0] = "True"
    code_array[0, 0, 0] = "10"   # A1
    code_array[0, 3, 0] = "99"   # D1

    # B1 = A1 if C1 else D1
    code_array[0, 1, 0] = 'C("A1") if C("C1") else C("D1")'
    assert code_array[0, 1, 0] == 10
    deps = code_array.dep_graph.dependencies[(0, 1, 0)]
    assert (0, 0, 0) in deps   # A1
    assert (0, 2, 0) in deps   # C1
    assert (0, 3, 0) not in deps  # D1 not taken

    # Flip condition and reevaluate B1
    code_array[0, 2, 0] = "False"
    assert code_array[0, 1, 0] == 99
    deps = code_array.dep_graph.dependencies[(0, 1, 0)]
    assert (0, 0, 0) not in deps
    assert (0, 2, 0) in deps
    assert (0, 3, 0) in deps


def test_cr_dynamic_target_rebuilds_dependencies(code_array):
    """CR(...) target changes should update dependency edges."""

    # A1 stores target address as string
    code_array[0, 0, 0] = '"C1"'
    code_array[0, 2, 0] = "11"  # C1
    code_array[0, 3, 0] = "22"  # D1

    # B1 reads target from A1
    code_array[0, 1, 0] = 'CR(C("A1"))'
    assert code_array[0, 1, 0] == 11
    deps = code_array.dep_graph.dependencies[(0, 1, 0)]
    assert (0, 0, 0) in deps  # A1
    assert (0, 2, 0) in deps  # C1
    assert (0, 3, 0) not in deps

    # Change target to D1 and reevaluate
    code_array[0, 0, 0] = '"D1"'
    assert code_array[0, 1, 0] == 22
    deps = code_array.dep_graph.dependencies[(0, 1, 0)]
    assert (0, 0, 0) in deps
    assert (0, 2, 0) not in deps
    assert (0, 3, 0) in deps
