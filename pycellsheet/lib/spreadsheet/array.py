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

try:
    from pycellsheet.lib.pycellsheet import EmptyCell, Range, flatten_args, RangeOutput
except ImportError:
    from lib.pycellsheet import EmptyCell, Range, flatten_args, RangeOutput

_ARRAY_FUNCTIONS = [
    'ARRAY_CONSTRAIN', 'BYCOL', 'BYROW', 'CHOOSECOLS', 'CHOOSEROWS', 'FLATTEN', 'FREQUENCY',
    'GROWTH', 'HSTACK', 'LINEST', 'LOGEST', 'MAKEARRAY', 'MAP', 'MDETERM', 'MINVERSE', 'MMULT',
    'REDUCE', 'SCAN', 'SUMPRODUCT', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'TOCOL', 'TOROW', 'TRANSPOSE',
    'TREND', 'VSTACK', 'WRAPCOLS', 'WRAPROWS'
]
__all__ = _ARRAY_FUNCTIONS + ["_ARRAY_FUNCTIONS"]


def ARRAY_CONSTRAIN(a, b):
    raise NotImplementedError("ARRAY_CONSTRAIN() not implemented yet")


def BYCOL(array, func):
    """Apply a Python function to each column of an array."""
    if not callable(func):
        raise ValueError("Second argument must be a callable function")

    if isinstance(array, Range):
        rows = len(array)
        cols = array.width
        result = []
        for c in range(cols):
            col = [array[r][c] for r in range(rows)]
            result.append(func(col))
        return result
    else:
        # 1D array - treat as single column
        lst = flatten_args(array)
        return func(lst)

def BYROW(array, func):
    """Apply a Python function to each row of an array."""
    if not callable(func):
        raise ValueError("Second argument must be a callable function")

    if isinstance(array, Range):
        rows = len(array)
        cols = array.width
        result = []
        for r in range(rows):
            row = [array[r][c] for c in range(cols)]
            result.append(func(row))
        return result
    else:
        # 1D array - treat as single row
        lst = flatten_args(array)
        return func(lst)

def CHOOSECOLS(array: Range, *col_nums):
    """Select specific columns from an array by index (1-based)."""
    if not isinstance(array, Range):
        raise ValueError("First argument must be a Range")

    rows = len(array)
    cols = array.width
    result = []

    for r in range(rows):
        for col_num in col_nums:
            col_idx = int(col_num) - 1  # Convert to 0-based
            if 0 <= col_idx < cols:
                result.append(array[r][col_idx])
            else:
                raise ValueError(f"Column index {col_num} out of range")

    return RangeOutput(rows, result)

def CHOOSEROWS(array: Range, *row_nums):
    """Select specific rows from an array by index (1-based)."""
    if not isinstance(array, Range):
        raise ValueError("First argument must be a Range")

    cols = array.width
    result = []

    for row_num in row_nums:
        row_idx = int(row_num) - 1  # Convert to 0-based
        if 0 <= row_idx < len(array):
            for c in range(cols):
                result.append(array[row_idx][c])
        else:
            raise ValueError(f"Row index {row_num} out of range")

    return RangeOutput(len(row_nums), result)

def FLATTEN(a, b):
    raise NotImplementedError("FLATTEN() not implemented yet")


def FREQUENCY(data, bins):
    """Count values into bins and return len(bins)+1 bucket counts."""
    data_vals = flatten_args(data)
    bin_vals = flatten_args(bins)
    # If bins is empty, you might decide how to handle. Let's just do normal logic.
    bin_vals.sort()

    # Initialize counts with len(bins)+1
    result = [0] * (len(bin_vals) + 1)
    for d in data_vals:
        placed = False
        for i, b in enumerate(bin_vals):
            if d <= b:
                result[i] += 1
                placed = True
                break
        if not placed:
            result[-1] += 1
    return result


def GROWTH(a, b):
    raise NotImplementedError("GROWTH() not implemented yet")


def HSTACK(*arrays):
    """Stack arrays horizontally (side by side)."""
    if not arrays:
        raise ValueError("At least one array is required")

    # Convert all arrays to 2D structure
    processed = []
    max_rows = 0

    for arr in arrays:
        if isinstance(arr, Range):
            rows = len(arr)
            cols = arr.width
            max_rows = max(max_rows, rows)
            processed.append((rows, cols, arr))
        else:
            # Flatten to 1D and treat as single column
            lst = flatten_args(arr) if hasattr(arr, '__iter__') else [arr]
            max_rows = max(max_rows, len(lst))
            processed.append((len(lst), 1, lst))

    # Build result by concatenating rows
    result = []
    for r in range(max_rows):
        for rows, cols, arr in processed:
            if isinstance(arr, Range):
                for c in range(cols):
                    if r < rows:
                        result.append(arr[r][c])
                    else:
                        result.append(0)  # Pad with zeros if needed
            else:
                if r < len(arr):
                    result.append(arr[r])
                else:
                    result.append(0)

    return RangeOutput(max_rows, result)

def LINEST(a, b):
    raise NotImplementedError("LINEST() not implemented yet")


def LOGEST(a, b):
    raise NotImplementedError("LOGEST() not implemented yet")


def MAKEARRAY(rows, cols, func):
    """Create an array by calling a Python function for each position."""
    if not callable(func):
        raise ValueError("Third argument must be a callable function")

    rows = int(rows)
    cols = int(cols)
    result = []

    for r in range(rows):
        for c in range(cols):
            # Call func with 0-based row, col indices
            result.append(func(r, c))

    return RangeOutput(rows, result)

def MAP(array, func):
    """Apply a Python function to each element of an array."""
    if not callable(func):
        raise ValueError("Second argument must be a callable function")

    if isinstance(array, Range):
        rows = len(array)
        cols = array.width
        result = []
        for r in range(rows):
            for c in range(cols):
                result.append(func(array[r][c]))
        return RangeOutput(rows, result)
    else:
        lst = flatten_args(array)
        return [func(x) for x in lst]

def MDETERM(a, b):
    raise NotImplementedError("MDETERM() not implemented yet")


def MINVERSE(a, b):
    raise NotImplementedError("MINVERSE() not implemented yet")


def MMULT(a, b):
    raise NotImplementedError("MMULT() not implemented yet")


def REDUCE(initial_value, array, func):
    """Reduce an array to a single value using a Python function."""
    if not callable(func):
        raise ValueError("Third argument must be a callable function")

    lst = flatten_args(array)
    accumulator = initial_value

    for value in lst:
        accumulator = func(accumulator, value)

    return accumulator

def SCAN(initial_value, array, func):
    """Running accumulation of an array using a Python function."""
    if not callable(func):
        raise ValueError("Third argument must be a callable function")

    lst = flatten_args(array)
    result = [initial_value]
    accumulator = initial_value

    for value in lst:
        accumulator = func(accumulator, value)
        result.append(accumulator)

    return result

def SUMPRODUCT(a, b):
    raise NotImplementedError("SUMPRODUCT() not implemented yet")


def SUMX2MY2(a, b):
    raise NotImplementedError("SUMX2MY2() not implemented yet")


def SUMX2PY2(a, b):
    raise NotImplementedError("SUMX2PY2() not implemented yet")


def SUMXMY2(a, b):
    raise NotImplementedError("SUMXMY2() not implemented yet")


def TOCOL(array, ignore=0, scan_by_column=False):
    """Convert array to a single column."""
    if isinstance(array, Range):
        rows = len(array)
        cols = array.width
        result = []

        if scan_by_column:
            # Scan column by column
            for c in range(cols):
                for r in range(rows):
                    val = array[r][c]
                    if ignore == 0 or (ignore == 1 and val != 0) or (ignore == 2 and val != "" and val is not None):
                        result.append(val)
        else:
            # Scan row by row (default)
            for r in range(rows):
                for c in range(cols):
                    val = array[r][c]
                    if ignore == 0 or (ignore == 1 and val != 0) or (ignore == 2 and val != "" and val is not None):
                        result.append(val)

        return RangeOutput(len(result), result)
    else:
        lst = flatten_args(array)
        return RangeOutput(len(lst), lst)

def TOROW(array, ignore=0, scan_by_column=False):
    """Convert array to a single row."""
    # Same as TOCOL but returns as single-row range
    col_result = TOCOL(array, ignore, scan_by_column)
    # Return as a single row (height=1)
    return col_result.lst if hasattr(col_result, 'lst') else col_result

def TRANSPOSE(a, b):
    raise NotImplementedError("TRANSPOSE() not implemented yet")


def TREND(a, b):
    raise NotImplementedError("TREND() not implemented yet")


def VSTACK(*arrays):
    """Stack arrays vertically (one on top of another)."""
    if not arrays:
        raise ValueError("At least one array is required")

    # Convert all arrays and track max width
    processed = []
    max_cols = 0

    for arr in arrays:
        if isinstance(arr, Range):
            rows = len(arr)
            cols = arr.width
            max_cols = max(max_cols, cols)
            processed.append((rows, cols, arr))
        else:
            # Flatten to 1D and treat as single row
            lst = flatten_args(arr) if hasattr(arr, '__iter__') else [arr]
            max_cols = max(max_cols, len(lst))
            processed.append((1, len(lst), lst))

    # Build result by stacking rows
    result = []
    total_rows = 0

    for rows, cols, arr in processed:
        if isinstance(arr, Range):
            for r in range(rows):
                for c in range(max_cols):
                    if c < cols:
                        result.append(arr[r][c])
                    else:
                        result.append(0)  # Pad with zeros
            total_rows += rows
        else:
            for c in range(max_cols):
                if c < len(arr):
                    result.append(arr[c])
                else:
                    result.append(0)
            total_rows += 1

    return RangeOutput(total_rows, result)

def WRAPCOLS(vector, wrap_count, pad_with=None):
    """Wrap a 1D array into columns."""
    lst = flatten_args(vector)
    wrap_count = int(wrap_count)

    if wrap_count <= 0:
        raise ValueError("wrap_count must be positive")

    # Calculate number of rows needed
    num_rows = (len(lst) + wrap_count - 1) // wrap_count
    result = []

    for r in range(num_rows):
        for c in range(wrap_count):
            idx = r + c * num_rows
            if idx < len(lst):
                result.append(lst[idx])
            else:
                result.append(pad_with if pad_with is not None else 0)

    return RangeOutput(num_rows, result)

def WRAPROWS(vector, wrap_count, pad_with=None):
    """Wrap a 1D array into rows."""
    lst = flatten_args(vector)
    wrap_count = int(wrap_count)

    if wrap_count <= 0:
        raise ValueError("wrap_count must be positive")

    # Calculate number of rows needed
    num_rows = (len(lst) + wrap_count - 1) // wrap_count
    result = []

    for r in range(num_rows):
        for c in range(wrap_count):
            idx = r * wrap_count + c
            if idx < len(lst):
                result.append(lst[idx])
            else:
                result.append(pad_with if pad_with is not None else 0)

    return RangeOutput(num_rows, result)
