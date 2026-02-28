# -*- coding: utf-8 -*-
#
# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License
#
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

"""Storage backend abstraction for cell code persistence."""

from __future__ import annotations

import os
from typing import Any, Iterable, Protocol, Tuple


CellKey = Tuple[int, int, int]


class CellStorageBackend(Protocol):
    """Protocol for pluggable cell-code storage backends."""

    def get_code(self, key: CellKey) -> Any:
        ...

    def set_code(self, key: CellKey, value: Any):
        ...

    def pop_code(self, key: CellKey) -> Any:
        ...

    def iter_keys(self) -> Iterable[CellKey]:
        ...

    def resize(self, shape: tuple[int, int, int]):
        ...

    def as_dict(self) -> dict[CellKey, Any]:
        ...

    def replace_from_dict(self, values: dict[CellKey, Any]):
        ...


class Dict3DBackend:
    """Adapter around the legacy DictGrid mapping behavior."""

    def __init__(self, dict_grid):
        self.dict_grid = dict_grid

    def get_code(self, key: CellKey) -> Any:
        return self.dict_grid[key]

    def set_code(self, key: CellKey, value: Any):
        self.dict_grid[key] = value

    def pop_code(self, key: CellKey) -> Any:
        return self.dict_grid.pop(key)

    def iter_keys(self) -> Iterable[CellKey]:
        return iter(self.dict_grid.keys())

    def resize(self, shape: tuple[int, int, int]):
        self.dict_grid.shape = shape

    def as_dict(self) -> dict[CellKey, Any]:
        return dict(self.dict_grid)

    def replace_from_dict(self, values: dict[CellKey, Any]):
        self.dict_grid.clear()
        self.dict_grid.update(values)


class MatrixSheetsBackend:
    """Array-of-2D-sheets backend with sparse per-sheet storage."""

    def __init__(self, dict_grid):
        self.dict_grid = dict_grid
        rows, cols, tabs = dict_grid.shape
        self._shape = (rows, cols, tabs)
        self._sheets: list[dict[tuple[int, int], Any]] = [dict() for _ in range(tabs)]

    def _validate_key(self, key: CellKey):
        row, col, tab = key
        rows, cols, tabs = self._shape
        if not (0 <= row < rows and 0 <= col < cols and 0 <= tab < tabs):
            raise IndexError(f"Grid index {key} outside grid shape {self._shape}.")

    def get_code(self, key: CellKey) -> Any:
        self._validate_key(key)
        row, col, tab = key
        return self._sheets[tab].get((row, col))

    def set_code(self, key: CellKey, value: Any):
        self._validate_key(key)
        row, col, tab = key
        self._sheets[tab][(row, col)] = value

    def pop_code(self, key: CellKey) -> Any:
        self._validate_key(key)
        row, col, tab = key
        return self._sheets[tab].pop((row, col))

    def iter_keys(self) -> Iterable[CellKey]:
        for tab, sheet in enumerate(self._sheets):
            for row, col in sheet:
                yield (row, col, tab)

    def resize(self, shape: tuple[int, int, int]):
        old_rows, old_cols, old_tabs = self._shape
        new_rows, new_cols, new_tabs = shape
        self._shape = shape
        self.dict_grid.shape = shape

        if new_tabs < old_tabs:
            self._sheets = self._sheets[:new_tabs]
        elif new_tabs > old_tabs:
            self._sheets.extend(dict() for _ in range(new_tabs - old_tabs))

        if new_rows < old_rows or new_cols < old_cols:
            for tab in range(new_tabs):
                sheet = self._sheets[tab]
                stale = [(r, c) for (r, c) in sheet if r >= new_rows or c >= new_cols]
                for rc in stale:
                    sheet.pop(rc, None)

    def as_dict(self) -> dict[CellKey, Any]:
        out: dict[CellKey, Any] = {}
        for tab, sheet in enumerate(self._sheets):
            for (row, col), value in sheet.items():
                out[(row, col, tab)] = value
        return out

    def replace_from_dict(self, values: dict[CellKey, Any]):
        rows, cols, tabs = self._shape
        self._sheets = [dict() for _ in range(tabs)]
        for key, value in values.items():
            row, col, tab = key
            if 0 <= row < rows and 0 <= col < cols and 0 <= tab < tabs:
                self._sheets[tab][(row, col)] = value
            else:
                raise IndexError(f"Grid index {key} outside grid shape {self._shape}.")


def create_storage_backend(dict_grid):
    """Factory for storage backend selection.

    Supported values:
      - ``matrix2d`` (default): array of sparse 2D sheet maps
      - ``dict3d``: legacy DictGrid adapter
    """

    backend = os.environ.get("PYCELLSHEET_STORAGE_BACKEND", "matrix2d").strip().lower()
    if backend == "dict3d":
        return Dict3DBackend(dict_grid)
    if backend == "matrix2d":
        return MatrixSheetsBackend(dict_grid)
    raise ValueError(
        f"Unknown storage backend {backend!r}. "
        "Supported backends: 'matrix2d', 'dict3d'."
    )
