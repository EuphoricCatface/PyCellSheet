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

**PyCellSheet Exception Hierarchy**

Provides:
 * PyCellSheetError - Base exception for all PyCellSheet-specific errors
 * CircularRefError - Raised when circular reference is detected
 * SpillRefError - Raised when RangeOutput spill collides with existing content

"""


class PyCellSheetError(RuntimeError):
    """Base exception for all PyCellSheet-specific errors"""
    pass


class CircularRefError(PyCellSheetError):
    """Raised when a circular reference is detected in cell dependencies"""

    def __init__(self, cycle):
        """Initialize CircularRefError with cycle path

        Parameters
        ----------
        cycle: list of tuples or str
            List of cell keys (row, col, table) that form the cycle,
            or a string message (when unpickling).

        """

        # Handle both list (normal creation) and str (unpickling)
        if isinstance(cycle, str):
            # Called during unpickling with the formatted message
            # We can't recover the cycle list, so store empty list
            self.cycle = []
            super().__init__(cycle)
        else:
            # Normal creation with cycle list
            self.cycle = cycle if isinstance(cycle, list) else list(cycle)
            super().__init__(self._format_message())

    def _format_message(self):
        """Format the error message showing the cycle path"""

        if not self.cycle:
            return "Circular reference detected (empty cycle)"

        if len(self.cycle) == 1:
            return f"Circular reference: {self.cycle[0]} references itself"

        if len(self.cycle) == 2 and self.cycle[0] == self.cycle[1]:
            return f"Circular reference: {self.cycle[0]} references itself"

        # Format cycle as: A1 → A2 → A3 → A1
        cycle_str = " → ".join(str(cell) for cell in self.cycle)
        return f"Circular reference: {cycle_str}"

    def __reduce__(self):
        """Custom pickle support to preserve cycle list"""
        return (self.__class__, (self.cycle,))


class SpillRefError(PyCellSheetError):
    """Raised when RangeOutput spill expansion hits a conflicting cell."""

    def __init__(self, anchor_key, conflict_key):
        self.anchor_key = anchor_key
        if isinstance(conflict_key, (list, tuple, set)):
            self.conflict_keys = list(conflict_key)
        else:
            self.conflict_keys = [conflict_key]
        # Compatibility alias for older call sites/tests expecting a singular key.
        self.conflict_key = self.conflict_keys[0] if self.conflict_keys else None
        if len(self.conflict_keys) == 1:
            msg = f"Spill conflict from {anchor_key} blocked by occupied cell {self.conflict_key}"
        else:
            msg = f"Spill conflict from {anchor_key} blocked by occupied cell(s) {self.conflict_keys}"
        super().__init__(msg)
