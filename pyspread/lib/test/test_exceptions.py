# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
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

**Unit tests for exceptions.py**

"""

import pytest
from ..exceptions import PyCellSheetError, CircularRefError


def test_pycellsheet_error_hierarchy():
    """Test that PyCellSheetError inherits from RuntimeError"""

    assert issubclass(PyCellSheetError, RuntimeError)


def test_circular_ref_error_hierarchy():
    """Test that CircularRefError inherits from PyCellSheetError"""

    assert issubclass(CircularRefError, PyCellSheetError)
    assert issubclass(CircularRefError, RuntimeError)


def test_circular_ref_error_simple():
    """Test CircularRefError with simple cycle"""

    cycle = [(0, 0, 0), (1, 0, 0), (0, 0, 0)]  # A1 -> B1 -> A1
    err = CircularRefError(cycle)

    assert "Circular reference" in str(err)
    assert "(0, 0, 0)" in str(err)
    assert "(1, 0, 0)" in str(err)


def test_circular_ref_error_complex():
    """Test CircularRefError with complex cycle"""

    cycle = [(0, 0, 0), (1, 0, 0), (2, 0, 0), (3, 0, 0), (0, 0, 0)]
    err = CircularRefError(cycle)

    assert "Circular reference" in str(err)
    assert len(str(err)) > 0


def test_circular_ref_error_self_reference():
    """Test CircularRefError with self-reference"""

    cycle = [(0, 0, 0), (0, 0, 0)]  # A1 -> A1
    err = CircularRefError(cycle)

    assert "Circular reference" in str(err)
    assert "(0, 0, 0)" in str(err)


def test_circular_ref_error_empty_cycle():
    """Test CircularRefError with empty cycle list"""

    cycle = []
    err = CircularRefError(cycle)

    # Should not crash, should have reasonable message
    assert "Circular reference" in str(err)


def test_circular_ref_error_can_be_raised():
    """Test that CircularRefError can be raised and caught"""

    with pytest.raises(CircularRefError) as exc_info:
        raise CircularRefError([(0, 0, 0), (1, 0, 0), (0, 0, 0)])

    assert "Circular reference" in str(exc_info.value)


def test_circular_ref_error_can_be_caught_as_runtime_error():
    """Test that CircularRefError can be caught as RuntimeError"""

    with pytest.raises(RuntimeError):
        raise CircularRefError([(0, 0, 0)])
