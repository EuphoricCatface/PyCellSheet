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

from os.path import abspath, dirname, join
import sys

import pytest

project_path = abspath(join(dirname(__file__) + "/../.."))
sys.path.insert(0, project_path)

from model.model import DictGrid
from model.storage_backend import (
    Dict3DBackend,
    MatrixSheetsBackend,
    create_storage_backend,
)

sys.path.pop(0)


def test_create_storage_backend_defaults_to_matrix2d(monkeypatch):
    monkeypatch.delenv("PYCELLSHEET_STORAGE_BACKEND", raising=False)
    dict_grid = DictGrid((4, 4, 2))

    backend = create_storage_backend(dict_grid)

    assert isinstance(backend, MatrixSheetsBackend)


def test_create_storage_backend_honors_dict3d_override(monkeypatch):
    monkeypatch.setenv("PYCELLSHEET_STORAGE_BACKEND", "dict3d")
    dict_grid = DictGrid((4, 4, 2))

    backend = create_storage_backend(dict_grid)

    assert isinstance(backend, Dict3DBackend)


def test_matrix_backend_basic_mutation_and_roundtrip():
    dict_grid = DictGrid((3, 3, 2))
    backend = MatrixSheetsBackend(dict_grid)

    backend.set_code((1, 2, 1), "x")
    backend.set_code((0, 0, 0), "y")
    assert backend.get_code((1, 2, 1)) == "x"
    assert backend.get_code((2, 2, 1)) is None

    keys = sorted(backend.iter_keys())
    assert keys == [(0, 0, 0), (1, 2, 1)]
    assert backend.as_dict() == {(0, 0, 0): "y", (1, 2, 1): "x"}

    assert backend.pop_code((0, 0, 0)) == "y"
    assert backend.as_dict() == {(1, 2, 1): "x"}


def test_matrix_backend_resize_prunes_out_of_bounds_cells():
    dict_grid = DictGrid((5, 5, 2))
    backend = MatrixSheetsBackend(dict_grid)
    backend.replace_from_dict({
        (0, 0, 0): "a",
        (4, 4, 0): "b",
        (3, 3, 1): "c",
    })

    backend.resize((4, 4, 1))

    assert dict_grid.shape == (4, 4, 1)
    assert backend.as_dict() == {(0, 0, 0): "a"}


def test_matrix_backend_replace_from_dict_rejects_out_of_bounds():
    dict_grid = DictGrid((2, 2, 1))
    backend = MatrixSheetsBackend(dict_grid)

    with pytest.raises(IndexError):
        backend.replace_from_dict({(2, 0, 0): "bad"})
