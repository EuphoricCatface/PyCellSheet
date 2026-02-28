# -*- coding: utf-8 -*-

# Created by Seongyong Park (EuphCat)
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

**Dependency Graph for Cell References**

Provides:
 * DependencyGraph - Tracks dependencies between cells and manages dirty flags

"""

from collections import defaultdict
import logging
from .exceptions import CircularRefError


logger = logging.getLogger(__name__)


class DependencyGraph:
    """Tracks dependencies between cells and manages dirty flags

    Maintains two-way edges:
    - Forward edges: dependencies[A2] = {A1} (A2 depends on A1)
    - Reverse edges: dependents[A1] = {A2} (A1 is depended on by A2)

    Also tracks dirty flags for cells that need recalculation.

    """

    def __init__(self):
        """Initialize an empty dependency graph"""

        # Forward edges: {dependent: {dependencies}}
        # e.g., dependencies[(0, 1, 0)] = {(0, 0, 0)} means A2 depends on A1
        self.dependencies = defaultdict(set)

        # Reverse edges: {dependency: {dependents}}
        # e.g., dependents[(0, 0, 0)] = {(0, 1, 0)} means A1 is depended on by A2
        self.dependents = defaultdict(set)

        # Dirty flags: set of cells that need recalculation
        self.dirty = set()

        # Transitive-closure caches to avoid repeated DFS on hot paths.
        self._deps_closure_cache = {}
        self._dependents_closure_cache = {}

    def _invalidate_closure_cache(self):
        """Invalidate cached transitive-closure lookups."""

        self._deps_closure_cache.clear()
        self._dependents_closure_cache.clear()

    def add_dependency(self, dependent, dependency):
        """Add a dependency relationship

        Parameters
        ----------
        dependent: tuple
            Cell key (row, col, table) that depends on another cell
        dependency: tuple
            Cell key (row, col, table) that is depended upon

        Example
        -------
        add_dependency((0, 1, 0), (0, 0, 0))  # A2 depends on A1

        """

        self.dependencies[dependent].add(dependency)
        self.dependents[dependency].add(dependent)
        self._invalidate_closure_cache()
        logger.debug("Added dependency %s -> %s", dependent, dependency)

    def remove_cell(self, key, remove_reverse_edges=False, clear_dirty=True):
        """Remove dependency relationships for a cell

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to remove
        remove_reverse_edges: bool
            If True, removes both forward and reverse edges (full removal).
            If False (default), only removes forward edges (what this cell depends on),
            keeping reverse edges (what depends on this cell) intact.

            Default is False to preserve dependencies from cells that reference
            this cell, allowing proper invalidation when the cell is re-added.
        clear_dirty: bool
            If True (default), clear the cell dirty flag. Set False when
            rebuilding dependencies during evaluation so dirty status can be
            resolved after evaluation logic completes.

        """

        logger.debug("Removing cell %s (remove_reverse_edges=%s, clear_dirty=%s)",
                     key, remove_reverse_edges, clear_dirty)

        # Remove forward edges: this cell no longer depends on anything
        if key in self.dependencies:
            for dependency in self.dependencies[key]:
                self.dependents[dependency].discard(key)
            del self.dependencies[key]
            logger.debug("Removed forward edges for cell %s", key)

        # Remove reverse edges: nothing depends on this cell anymore
        if remove_reverse_edges and key in self.dependents:
            for dependent in self.dependents[key]:
                self.dependencies[dependent].discard(key)
            del self.dependents[key]
            logger.debug("Removed reverse edges for cell %s", key)

        # Clear dirty flag for removed cell when requested.
        if clear_dirty:
            self.dirty.discard(key)
        self._invalidate_closure_cache()
        if clear_dirty:
            logger.debug("Cleared dirty flag for removed cell %s", key)

    def check_for_cycles(self, start_key):
        """Check if there are any cycles starting from the given cell

        Uses depth-first search with a recursion stack to detect cycles.

        Parameters
        ----------
        start_key: tuple
            Cell key (row, col, table) to start checking from

        Raises
        ------
        CircularRefError
            If a cycle is detected

        """

        logger.debug("Checking for cycles starting from %s", start_key)
        visited = set()
        rec_stack = []
        rec_stack_set = set()

        stack = [(start_key, iter(self.dependencies.get(start_key, set())))]
        while stack:
            key, iterator = stack[-1]

            if key not in visited:
                visited.add(key)
                rec_stack.append(key)
                rec_stack_set.add(key)

            try:
                dependency = next(iterator)
            except StopIteration:
                stack.pop()
                rec_stack.pop()
                rec_stack_set.remove(key)
                continue

            if dependency in rec_stack_set:
                cycle_start_idx = rec_stack.index(dependency)
                cycle = rec_stack[cycle_start_idx:] + [dependency]
                logger.debug("Cycle detected: %s", cycle)
                raise CircularRefError(cycle)

            if dependency not in visited:
                stack.append((dependency, iter(self.dependencies.get(dependency, set()))))

        logger.debug("No cycles detected from %s", start_key)

    def mark_dirty(self, key):
        """Mark a cell and all its dependents as dirty

        Propagates the dirty flag through all transitive dependents.

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to mark dirty

        """

        stack = [key]
        while stack:
            current = stack.pop()
            if current in self.dirty:
                logger.debug("Cell %s already dirty; skipping", current)
                continue

            self.dirty.add(current)
            logger.debug("Marked cell %s dirty", current)
            stack.extend(self.dependents.get(current, set()))

    def is_dirty(self, key):
        """Check if a cell is marked dirty

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to check

        Returns
        -------
        bool
            True if cell is dirty, False otherwise

        """

        dirty = key in self.dirty
        logger.debug("Dirty check for %s: %s", key, dirty)
        return dirty

    def clear_dirty(self, key):
        """Clear the dirty flag for a cell

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to clear

        """

        self.dirty.discard(key)
        logger.debug("Cleared dirty flag for %s", key)

    def get_all_dependencies(self, key):
        """Get all transitive dependencies for a cell

        Returns the transitive closure of all cells this cell depends on.

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to get dependencies for

        Returns
        -------
        set
            Set of all cell keys (transitive dependencies)

        Example
        -------
        If A3 depends on A2, and A2 depends on A1:
        get_all_dependencies((0, 2, 0)) returns {(0, 0, 0), (0, 1, 0)}

        """

        cached = self._deps_closure_cache.get(key)
        if cached is not None:
            logger.debug("Transitive dependencies cache hit for %s", key)
            return set(cached)

        result = set()
        visited = {key}
        stack = [key]
        while stack:
            current_key = stack.pop()
            for dependency in self.dependencies.get(current_key, set()):
                if dependency in visited:
                    continue
                visited.add(dependency)
                result.add(dependency)
                stack.append(dependency)

        self._deps_closure_cache[key] = frozenset(result)
        logger.debug("Transitive dependencies for %s: %s", key, result)
        return result

    def get_all_dependents(self, key):
        """Get all transitive dependents for a cell

        Returns the transitive closure of all cells that depend on this cell.

        Parameters
        ----------
        key: tuple
            Cell key (row, col, table) to get dependents for

        Returns
        -------
        set
            Set of all cell keys (transitive dependents)

        Example
        -------
        If A3 depends on A2, and A2 depends on A1:
        get_all_dependents((0, 0, 0)) returns {(0, 1, 0), (0, 2, 0)}

        """

        cached = self._dependents_closure_cache.get(key)
        if cached is not None:
            logger.debug("Transitive dependents cache hit for %s", key)
            return set(cached)

        result = set()
        visited = {key}
        stack = [key]
        while stack:
            current_key = stack.pop()
            for dependent in self.dependents.get(current_key, set()):
                if dependent in visited:
                    continue
                visited.add(dependent)
                result.add(dependent)
                stack.append(dependent)

        self._dependents_closure_cache[key] = frozenset(result)
        logger.debug("Transitive dependents for %s: %s", key, result)
        return result

    def get_all_dirty(self):
        """Get all dirty cells

        Returns
        -------
        set
            Set of all dirty cell keys

        """

        all_dirty = set(self.dirty)
        logger.debug("All dirty cells requested: %s", all_dirty)
        return all_dirty
