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


def create_storage_backend(dict_grid):
    """Factory for storage backend selection.

    NOTE: matrix backend is intentionally not selectable yet, because several
    code paths still write directly via ``dict_grid`` and must be migrated first.
    """

    _ = os.environ.get("PYCELLSHEET_STORAGE_BACKEND", "dict3d").strip().lower()
    return Dict3DBackend(dict_grid)

