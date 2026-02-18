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

**Unit tests for smart_cache.py**

"""

import pytest
from ..smart_cache import SmartCache
from ..dependency_graph import DependencyGraph


@pytest.fixture
def dep_graph():
    """Create a fresh DependencyGraph for each test"""
    return DependencyGraph()


@pytest.fixture
def cache(dep_graph):
    """Create a fresh SmartCache for each test"""
    return SmartCache(dep_graph)


# --- INVALID Sentinel Tests ---

def test_invalid_sentinel_is_unique(cache):
    """Test that INVALID sentinel is a unique object"""

    assert SmartCache.INVALID is not None
    assert SmartCache.INVALID is not False
    assert SmartCache.INVALID is not 0
    assert SmartCache.INVALID is not ""


def test_invalid_sentinel_identity(cache):
    """Test that INVALID sentinel maintains identity"""

    sentinel1 = SmartCache.INVALID
    sentinel2 = SmartCache.INVALID

    assert sentinel1 is sentinel2


def test_cache_miss_returns_invalid(cache):
    """Test that cache miss returns INVALID sentinel"""

    result = cache.get((0, 0, 0))

    assert result is SmartCache.INVALID


def test_cached_none_vs_invalid(cache):
    """Test that cached None is different from INVALID"""

    # Store None in cache
    cache.set((0, 0, 0), None)

    result = cache.get((0, 0, 0))

    # Should get None, not INVALID
    assert result is None
    assert result is not SmartCache.INVALID


# --- Basic Cache Operations ---

def test_set_and_get_simple(cache):
    """Test basic set and get operations"""

    cache.set((0, 0, 0), 42)

    result = cache.get((0, 0, 0))

    assert result == 42


def test_set_and_get_various_types(cache):
    """Test caching different types"""

    test_values = [
        42,
        "hello",
        [1, 2, 3],
        {"key": "value"},
        None,
        0,
        False,
        "",
    ]

    for i, value in enumerate(test_values):
        key = (0, i, 0)
        cache.set(key, value)
        result = cache.get(key)

        # Use 'is' for None/False, '==' for others
        if value is None or value is False:
            assert result is value
        else:
            assert result == value


def test_overwrite_cached_value(cache):
    """Test overwriting an existing cached value"""

    key = (0, 0, 0)

    cache.set(key, 42)
    assert cache.get(key) == 42

    cache.set(key, 99)
    assert cache.get(key) == 99


def test_multiple_cells(cache):
    """Test caching multiple cells"""

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)
    cache.set((0, 2, 0), 30)

    assert cache.get((0, 0, 0)) == 10
    assert cache.get((0, 1, 0)) == 20
    assert cache.get((0, 2, 0)) == 30


# --- Dependency-Based Invalidation Tests ---

def test_invalidate_single_cell(cache, dep_graph):
    """Test invalidating a single cell"""

    cache.set((0, 0, 0), 42)
    assert cache.get((0, 0, 0)) == 42

    cache.invalidate((0, 0, 0))

    result = cache.get((0, 0, 0))
    assert result is SmartCache.INVALID


def test_invalidate_marks_dirty(cache, dep_graph):
    """Test that invalidate marks the cell and its dependents as dirty"""

    # Setup: A2 depends on A1
    dep_graph.add_dependency((0, 1, 0), (0, 0, 0))

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)

    # Invalidate A1
    cache.invalidate((0, 0, 0))

    # Both A1 and A2 should be marked dirty
    assert dep_graph.is_dirty((0, 0, 0))
    assert dep_graph.is_dirty((0, 1, 0))


def test_is_valid_checks_dirty_flag(cache, dep_graph):
    """Test that is_valid checks the dirty flag"""

    key = (0, 0, 0)

    cache.set(key, 42)
    assert cache.is_valid(key)

    # Mark dirty without invalidating cache
    dep_graph.mark_dirty(key)

    # Cache entry exists but is not valid due to dirty flag
    assert not cache.is_valid(key)


def test_is_valid_checks_dependencies(cache, dep_graph):
    """Test that is_valid checks if dependencies are dirty"""

    # Setup: A2 depends on A1
    dep_graph.add_dependency((0, 1, 0), (0, 0, 0))

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)

    assert cache.is_valid((0, 1, 0))

    # Mark A1 dirty
    dep_graph.mark_dirty((0, 0, 0))

    # A2's cache is invalid because its dependency (A1) is dirty
    assert not cache.is_valid((0, 1, 0))


def test_is_valid_transitive_dependencies(cache, dep_graph):
    """Test that is_valid checks transitive dependencies"""

    # Setup: A3 depends on A2, A2 depends on A1
    dep_graph.add_dependency((0, 1, 0), (0, 0, 0))
    dep_graph.add_dependency((0, 2, 0), (0, 1, 0))

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)
    cache.set((0, 2, 0), 30)

    assert cache.is_valid((0, 2, 0))

    # Mark A1 dirty
    dep_graph.mark_dirty((0, 0, 0))

    # A3's cache is invalid because transitive dependency (A1) is dirty
    assert not cache.is_valid((0, 2, 0))


def test_get_respects_dirty_flag(cache, dep_graph):
    """Test that get returns INVALID when cell is dirty"""

    key = (0, 0, 0)

    cache.set(key, 42)
    dep_graph.mark_dirty(key)

    result = cache.get(key)

    # Even though cache has value, should return INVALID due to dirty flag
    assert result is SmartCache.INVALID


def test_get_respects_dirty_dependencies(cache, dep_graph):
    """Test that get returns INVALID when dependencies are dirty"""

    # Setup: A2 depends on A1
    dep_graph.add_dependency((0, 1, 0), (0, 0, 0))

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)

    # Mark A1 dirty
    dep_graph.mark_dirty((0, 0, 0))

    result = cache.get((0, 1, 0))

    # A2's cache is invalid because A1 is dirty
    assert result is SmartCache.INVALID


# --- Clear Tests ---

def test_clear_all(cache):
    """Test clearing all cache entries"""

    cache.set((0, 0, 0), 10)
    cache.set((0, 1, 0), 20)
    cache.set((0, 2, 0), 30)

    cache.clear()

    assert cache.get((0, 0, 0)) is SmartCache.INVALID
    assert cache.get((0, 1, 0)) is SmartCache.INVALID
    assert cache.get((0, 2, 0)) is SmartCache.INVALID


def test_clear_after_invalidate(cache, dep_graph):
    """Test that clear removes all entries"""

    cache.set((0, 0, 0), 42)
    cache.invalidate((0, 0, 0))
    cache.clear()

    result = cache.get((0, 0, 0))

    assert result is SmartCache.INVALID


# --- Edge Cases ---

def test_invalidate_nonexistent_cell(cache, dep_graph):
    """Test invalidating a cell that was never cached"""

    # Should not crash
    cache.invalidate((9, 9, 9))

    # Cell should be marked dirty
    assert dep_graph.is_dirty((9, 9, 9))


def test_is_valid_nonexistent_cell(cache):
    """Test is_valid on a cell that was never cached"""

    # Should return False (not valid because not cached)
    assert not cache.is_valid((9, 9, 9))


def test_cache_after_clear_dirty(cache, dep_graph):
    """Test that caching after clearing dirty flag works"""

    key = (0, 0, 0)

    cache.set(key, 42)
    dep_graph.mark_dirty(key)

    # Clear dirty flag
    dep_graph.clear_dirty(key)

    # Now cache should be valid again
    assert cache.is_valid(key)
    assert cache.get(key) == 42
