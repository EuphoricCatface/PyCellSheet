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

**PyCellSheet Exception Hierarchy**

Provides:
 * PyCellSheetError - Base exception for all PyCellSheet-specific errors
 * CircularRefError - Raised when circular reference is detected

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
        cycle: list of tuples
            List of cell keys (row, col, table) that form the cycle.
            The first and last elements should be the same cell.

        """

        self.cycle = cycle
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
