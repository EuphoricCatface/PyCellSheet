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

**Smart Cache for Cell Results**

Provides:
 * SmartCache - Dependency-aware cache with INVALID sentinel

"""


class SmartCache:
    """Dependency-aware cache for cell evaluation results

    Uses an INVALID sentinel to distinguish "not cached" from "cached None".
    Checks dependency graph dirty flags before returning cached values.

    Key features:
    - INVALID sentinel: cache miss vs cached None
    - Dependency-aware: checks if dependencies are dirty
    - Lazy deepcopy: stores original, deepcopy happens at retrieval (in CodeArray)

    """

    # Sentinel object to distinguish "not cached" from "cached None"
    INVALID = object()

    def __init__(self, dep_graph):
        """Initialize SmartCache

        Parameters
        ----------
        dep_graph: DependencyGraph
            The dependency graph to check dirty flags against

        """

        self.dep_graph = dep_graph
        self._cache = {}

    def set(self, key, value):
        """Store a value in the cache

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table)
        value: any
            Value to cache (can be None)

        """

        self._cache[key] = value

    def get(self, key):
        """Get a cached value if valid, otherwise return INVALID

        Checks:
        1. Is the value in cache?
        2. Is the cell marked dirty?
        3. Are any of its dependencies dirty?

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table)

        Returns
        -------
        any or INVALID
            Cached value if valid, INVALID sentinel otherwise

        """

        if not self.is_valid(key):
            return self.INVALID

        return self._cache.get(key, self.INVALID)

    def is_valid(self, key):
        """Check if a cached value is valid

        A cached value is valid if:
        1. It exists in the cache
        2. The cell is not marked dirty
        3. None of its dependencies are dirty

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table)

        Returns
        -------
        bool
            True if cached value is valid, False otherwise

        """

        # Not in cache at all
        if key not in self._cache:
            return False

        # Cell itself is dirty
        if self.dep_graph.is_dirty(key):
            return False

        # Check if any dependency is dirty
        all_deps = self.dep_graph.get_all_dependencies(key)
        for dep in all_deps:
            if self.dep_graph.is_dirty(dep):
                return False

        return True

    def invalidate(self, key):
        """Invalidate a cache entry

        Removes the cache entry and marks the cell and all its dependents dirty.

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table)

        """

        # Remove from cache
        self._cache.pop(key, None)

        # Mark dirty (propagates to dependents)
        self.dep_graph.mark_dirty(key)

    def clear(self):
        """Clear all cache entries"""

        self._cache.clear()
