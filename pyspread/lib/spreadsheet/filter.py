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

try:
    from pyspread.lib.pycellsheet import EmptyCell, Range, RangeOutput, flatten_args
except ImportError:
    from lib.pycellsheet import EmptyCell, Range, RangeOutput, flatten_args

_FILTER_FUNCTIONS = [
    'FILTER', 'SORT', 'SORTN', 'UNIQUE'
]
__all__ = _FILTER_FUNCTIONS + ["_FILTER_FUNCTIONS"]


def FILTER(data_range: Range, *test_pairs) -> RangeOutput:
    """
    FILTER(data_range, (test_range1, crit_func1), (test_range2, crit_func2), ...)

    - data_range: the range whose rows we'll output.
    - Each (range_to_test, criterion_func):
       - range_to_test must have the same 'height' as data_range, and 1 width.
       - criterion_func(row_value, row_index, row_data) -> bool
         If None, default = bool(row_value).

    We'll only keep rows where *all* criteria return True.
    """
    if not test_pairs:
        # No pairs => just output the entire data_range
        return RangeOutput(width=data_range.width, lst=data_range.lst)

    # Prepare to store final row data
    new_list = []

    for y in range(data_range.height):
        row_data = data_range[y]  # the row from the main data
        keep = True

        # Check each (range_to_test, criterion_func)
        for (rng_test, crit_func) in test_pairs:
            # Validate shapes
            if rng_test.width != 1:
                raise ValueError("FILTER: test_range doesn't have 1 as its width.")
            if rng_test.height != data_range.height:
                raise ValueError("FILTER: mismatch in row count between data_range and test_range.")

            test_value = rng_test[y][0]

            # If crit_func is None, default to bool
            if crit_func is None:
                pass_check = bool(test_value)
            else:
                # pass test_value, row_index, row_data
                pass_check = crit_func(test_value)

            if not pass_check:
                keep = False
                break

        if keep:
            new_list.extend(row_data)

    return RangeOutput(width=data_range.width, lst=new_list)


def SORT(data_range: Range, *sort_pairs) -> RangeOutput:
    """
    SORT(data_range, (range_to_sort1, is_asc1), (range_to_sort2, is_asc2), ...)

    Multi-level sort:
      - We'll do stable sorts from rightmost pair to leftmost pair, so the leftmost pair is highest priority.
    """
    # Convert each row to a tuple (row_data, row_index) for stable sorting
    row_count = data_range.height
    rows = [(data_range[y], y) for y in range(row_count)]

    # We apply each (range_to_sort, is_ascending) in REVERSE order
    # because Python's sort is stable => final leftmost key is last sort
    for (rng_to_sort, is_asc) in reversed(sort_pairs):
        # Check shape
        if rng_to_sort.width != 1:
            raise ValueError("SORT: range_to_sort doesn't have 1 as its width.")
        if rng_to_sort.height != row_count:
            raise ValueError("SORT: range_to_sort row count mismatch with data_range.")

        def sort_key(item):
            (row_data, row_idx) = item
            val = rng_to_sort[row_idx]
            return val[0]

        rows.sort(key=sort_key, reverse=(not is_asc))

    # Now reassemble final ordering
    new_list = []
    for (row_data, _) in rows:
        new_list.extend(row_data)

    return RangeOutput(width=data_range.width, lst=new_list)


def SORTN(data_range: Range, n=1, display_ties_mode='default', *sort_pairs) -> RangeOutput:
    """
    SORTN(data_range, n=1, display_ties_mode='default', (range_to_sort1, is_asc1), ...)

    1. Perform multi-level sort like SORT.
    2. Truncate to top n rows (or handle ties if display_ties_mode='include_ties').
    """
    # Step 1: do a multi-level sort, reusing logic from SORT.
    sorted_ro = SORT(data_range, *sort_pairs)
    sorted_list = sorted_ro.lst  # a flat list
    row_width = sorted_ro.width
    total_rows = len(sorted_list) // row_width

    if n <= 0:
        # Return empty
        return RangeOutput(width=row_width, lst=[])

    # Step 2: handle 'display_ties_mode'
    if display_ties_mode == 'default':
        # Just take first n rows
        sliced = sorted_list[: n * row_width]

    elif display_ties_mode == 'include_ties':
        # We need to see if the n-th row ties with the (n+1)-th row
        # We'll compare the final sort key. That means we need the last sort pair's key, or single-col assumption.
        # For simplicity, let's assume single-col last sort or just compare the entire row?
        # We'll approximate.

        # Start by slicing the top n rows
        sliced = sorted_list[: n * row_width]
        if n < total_rows:
            # check tie with row n and row n+1
            nth_row = sliced[-row_width:]
            next_row = sorted_list[n * row_width : (n + 1) * row_width]
            if nth_row == next_row:
                # The row n is identical to row n+1 => there's a tie => let's expand
                # In practice, we might keep expanding until the row changes
                idx = n
                while idx < total_rows:
                    rowstart = idx * row_width
                    rowend = rowstart + row_width
                    candidate = sorted_list[rowstart:rowend]
                    if candidate == nth_row:
                        sliced.extend(candidate)
                        idx += 1
                    else:
                        break
            # else if no tie => do nothing

    else:
        # unrecognized or advanced modes => default to 'default'
        sliced = sorted_list[: n * row_width]

    return RangeOutput(width=row_width, lst=sliced)


def UNIQUE(data_range: Range, by_col=False, exactly_once=False) -> RangeOutput:
    """
    UNIQUE(data_range, by_col=False, exactly_once=False)

    Returns unique rows (or columns if by_col=True) from data_range.

    - data_range: the range to filter for unique values
    - by_col: if True, transpose logic to find unique columns instead of rows
    - exactly_once: if True, only return rows/cols that appear exactly once

    Preserves the order of first occurrence.
    """
    if by_col:
        # Transpose: treat columns as the unit of uniqueness
        width = data_range.width
        height = data_range.height

        # Extract each column as a tuple for hashability
        seen = {}  # column_tuple -> count
        order = []  # list of column indices in order of first appearance

        for col_idx in range(width):
            col_data = []
            for row_idx in range(height):
                row = data_range[row_idx]
                col_data.append(row[col_idx] if col_idx < len(row) else EmptyCell())

            col_tuple = tuple(col_data)
            if col_tuple not in seen:
                seen[col_tuple] = 1
                order.append(col_tuple)
            else:
                seen[col_tuple] += 1

        # Filter by exactly_once if needed
        if exactly_once:
            unique_cols = [col for col in order if seen[col] == 1]
        else:
            unique_cols = order

        # Transpose back: convert columns to rows
        new_list = []
        for row_idx in range(height):
            for col_tuple in unique_cols:
                new_list.append(col_tuple[row_idx])

        return RangeOutput(width=len(unique_cols), lst=new_list)

    else:
        # Default: find unique rows
        width = data_range.width
        height = data_range.height

        seen = {}  # row_tuple -> count
        order = []  # list of row tuples in order of first appearance

        for row_idx in range(height):
            row = data_range[row_idx]
            row_tuple = tuple(row)

            if row_tuple not in seen:
                seen[row_tuple] = 1
                order.append(row_tuple)
            else:
                seen[row_tuple] += 1

        # Filter by exactly_once if needed
        if exactly_once:
            unique_rows = [row for row in order if seen[row] == 1]
        else:
            unique_rows = order

        # Flatten back to a list
        new_list = []
        for row_tuple in unique_rows:
            new_list.extend(row_tuple)

        return RangeOutput(width=width, lst=new_list)
