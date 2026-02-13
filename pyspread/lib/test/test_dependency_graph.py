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

**Unit tests for dependency_graph.py**

"""

import pytest
from ..dependency_graph import DependencyGraph
from ..exceptions import CircularRefError


@pytest.fixture
def graph():
    """Create a fresh DependencyGraph for each test"""
    return DependencyGraph()


# --- Basic Add/Remove Tests ---

def test_add_dependency_simple(graph):
    """Test adding a simple dependency"""

    # A2 depends on A1
    graph.add_dependency(dependent=(0, 1, 0), dependency=(0, 0, 0))

    assert (0, 0, 0) in graph.dependencies[(0, 1, 0)]
    assert (0, 1, 0) in graph.dependents[(0, 0, 0)]


def test_add_multiple_dependencies(graph):
    """Test adding multiple dependencies for one cell"""

    # A3 depends on A1 and A2
    graph.add_dependency((0, 2, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    assert (0, 0, 0) in graph.dependencies[(0, 2, 0)]
    assert (0, 1, 0) in graph.dependencies[(0, 2, 0)]
    assert (0, 2, 0) in graph.dependents[(0, 0, 0)]
    assert (0, 2, 0) in graph.dependents[(0, 1, 0)]


def test_add_dependency_idempotent(graph):
    """Test that adding same dependency twice is idempotent"""

    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 1, 0), (0, 0, 0))

    assert len(graph.dependencies[(0, 1, 0)]) == 1
    assert len(graph.dependents[(0, 0, 0)]) == 1


def test_remove_cell(graph):
    """Test removing a cell removes all its dependencies"""

    # Setup: A3 depends on A1 and A2
    graph.add_dependency((0, 2, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    # Remove A3
    graph.remove_cell((0, 2, 0))

    # A3 should have no dependencies
    assert len(graph.dependencies.get((0, 2, 0), set())) == 0

    # A1 and A2 should not list A3 as dependent
    assert (0, 2, 0) not in graph.dependents.get((0, 0, 0), set())
    assert (0, 2, 0) not in graph.dependents.get((0, 1, 0), set())


def test_remove_cell_with_dependents(graph):
    """Test removing a cell that other cells depend on"""

    # Setup: A2 depends on A1, A3 depends on A1
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 0, 0))

    # Remove A1 (which A2 and A3 depend on)
    graph.remove_cell((0, 0, 0))

    # A2 should still have A1 in dependencies (orphaned reference)
    # But A1 should not have A2 in dependents
    assert len(graph.dependents.get((0, 0, 0), set())) == 0


def test_remove_nonexistent_cell(graph):
    """Test removing a cell that doesn't exist"""

    # Should not crash
    graph.remove_cell((9, 9, 9))


# --- Cycle Detection Tests ---

def test_no_cycle_simple(graph):
    """Test that simple chain has no cycle"""

    # A1 -> A2 -> A3
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    # Should not raise
    graph.check_for_cycles((0, 0, 0))
    graph.check_for_cycles((0, 1, 0))
    graph.check_for_cycles((0, 2, 0))


def test_cycle_self_reference(graph):
    """Test detecting self-reference cycle"""

    # A1 -> A1
    graph.add_dependency((0, 0, 0), (0, 0, 0))

    with pytest.raises(CircularRefError):
        graph.check_for_cycles((0, 0, 0))


def test_cycle_simple_two_cells(graph):
    """Test detecting simple two-cell cycle"""

    # A1 -> A2 -> A1
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 0, 0), (0, 1, 0))

    with pytest.raises(CircularRefError):
        graph.check_for_cycles((0, 0, 0))


def test_cycle_complex(graph):
    """Test detecting complex multi-cell cycle"""

    # A1 -> A2 -> A3 -> A4 -> A2 (cycle at A2)
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))
    graph.add_dependency((0, 3, 0), (0, 2, 0))
    graph.add_dependency((0, 1, 0), (0, 3, 0))

    # Check from A2, which is part of the cycle
    with pytest.raises(CircularRefError):
        graph.check_for_cycles((0, 1, 0))


def test_cycle_diamond_no_cycle(graph):
    """Test that diamond dependency pattern has no cycle"""

    # A1 -> A2 -> A4
    #    -> A3 -> A4
    # (A4 depends on both A2 and A3)
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 0, 0))
    graph.add_dependency((0, 3, 0), (0, 1, 0))
    graph.add_dependency((0, 3, 0), (0, 2, 0))

    # Should not raise
    graph.check_for_cycles((0, 0, 0))


def test_cycle_cross_sheet(graph):
    """Test detecting cycle across sheets"""

    # Sheet0!A1 -> Sheet1!A1 -> Sheet0!A1
    graph.add_dependency((0, 0, 1), (0, 0, 0))
    graph.add_dependency((0, 0, 0), (0, 0, 1))

    with pytest.raises(CircularRefError):
        graph.check_for_cycles((0, 0, 0))


# --- Dirty Flag Tests ---

def test_mark_dirty_simple(graph):
    """Test marking a cell dirty"""

    graph.mark_dirty((0, 0, 0))

    assert (0, 0, 0) in graph.dirty


def test_mark_dirty_propagates(graph):
    """Test that marking a cell dirty propagates to dependents"""

    # A2 depends on A1, A3 depends on A2
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    # Mark A1 dirty
    graph.mark_dirty((0, 0, 0))

    # A1, A2, A3 should all be dirty
    assert (0, 0, 0) in graph.dirty
    assert (0, 1, 0) in graph.dirty
    assert (0, 2, 0) in graph.dirty


def test_mark_dirty_diamond(graph):
    """Test dirty propagation through diamond pattern"""

    # A1 -> A2 -> A4
    #    -> A3 -> A4
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 0, 0))
    graph.add_dependency((0, 3, 0), (0, 1, 0))
    graph.add_dependency((0, 3, 0), (0, 2, 0))

    # Mark A1 dirty
    graph.mark_dirty((0, 0, 0))

    # All cells should be dirty
    assert (0, 0, 0) in graph.dirty
    assert (0, 1, 0) in graph.dirty
    assert (0, 2, 0) in graph.dirty
    assert (0, 3, 0) in graph.dirty


def test_is_dirty_simple(graph):
    """Test checking if a cell is dirty"""

    assert not graph.is_dirty((0, 0, 0))

    graph.mark_dirty((0, 0, 0))

    assert graph.is_dirty((0, 0, 0))


def test_clear_dirty_simple(graph):
    """Test clearing dirty flag"""

    graph.mark_dirty((0, 0, 0))
    assert graph.is_dirty((0, 0, 0))

    graph.clear_dirty((0, 0, 0))
    assert not graph.is_dirty((0, 0, 0))


def test_clear_dirty_multiple(graph):
    """Test clearing dirty flags for multiple cells"""

    graph.mark_dirty((0, 0, 0))
    graph.mark_dirty((0, 1, 0))
    graph.mark_dirty((0, 2, 0))

    graph.clear_dirty((0, 0, 0))
    graph.clear_dirty((0, 1, 0))

    assert not graph.is_dirty((0, 0, 0))
    assert not graph.is_dirty((0, 1, 0))
    assert graph.is_dirty((0, 2, 0))


# --- Transitive Dependency Tests ---

def test_get_all_dependencies_simple(graph):
    """Test getting all dependencies (transitive closure)"""

    # A3 depends on A2, A2 depends on A1
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    deps = graph.get_all_dependencies((0, 2, 0))

    assert (0, 0, 0) in deps
    assert (0, 1, 0) in deps


def test_get_all_dependencies_diamond(graph):
    """Test getting all dependencies for diamond pattern"""

    # A1 -> A2 -> A4
    #    -> A3 -> A4
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 0, 0))
    graph.add_dependency((0, 3, 0), (0, 1, 0))
    graph.add_dependency((0, 3, 0), (0, 2, 0))

    deps = graph.get_all_dependencies((0, 3, 0))

    assert (0, 0, 0) in deps
    assert (0, 1, 0) in deps
    assert (0, 2, 0) in deps


def test_get_all_dependencies_no_deps(graph):
    """Test getting dependencies for cell with no dependencies"""

    deps = graph.get_all_dependencies((0, 0, 0))

    assert len(deps) == 0


def test_get_all_dependents_simple(graph):
    """Test getting all dependents (transitive closure)"""

    # A3 depends on A2, A2 depends on A1
    graph.add_dependency((0, 1, 0), (0, 0, 0))
    graph.add_dependency((0, 2, 0), (0, 1, 0))

    deps = graph.get_all_dependents((0, 0, 0))

    assert (0, 1, 0) in deps
    assert (0, 2, 0) in deps


def test_get_all_dependents_no_deps(graph):
    """Test getting dependents for cell with no dependents"""

    deps = graph.get_all_dependents((0, 0, 0))

    assert len(deps) == 0
