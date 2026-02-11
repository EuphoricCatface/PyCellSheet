try:
    from pyspread.lib.pycellsheet import Range, flatten_args
except ImportError:
    from lib.pycellsheet import Range, flatten_args

_FILTER_FUNCTIONS = [
    'FILTER', 'SORT', 'SORTN', 'UNIQUE'
]
__all__ = _FILTER_FUNCTIONS + ["_FILTER_FUNCTIONS"]


def FILTER(range_, condition):
    lst = flatten_args(range_)
    cond = flatten_args(condition) if isinstance(condition, Range) else condition
    return [v for v, c in zip(lst, cond) if c]


def SORT(range_, sort_column=1, is_ascending=True, *args):
    """Sort the contents of a range or array."""
    if isinstance(range_, Range):
        # Handle 2D range sorting
        rows = len(range_)
        cols = range_.width

        # Extract rows as lists
        data = []
        for r in range(rows):
            row = [range_[r][c] for c in range(cols)]
            data.append(row)

        # Sort by the specified column (1-indexed)
        sort_col_idx = int(sort_column) - 1
        if sort_col_idx >= cols or sort_col_idx < 0:
            sort_col_idx = 0

        # Sort the data
        sorted_data = sorted(data, key=lambda row: row[sort_col_idx], reverse=not is_ascending)

        # Flatten back to list for Range output
        from lib.pycellsheet import RangeOutput
        result = []
        for row in sorted_data:
            result.extend(row)

        return RangeOutput(len(sorted_data), result)
    else:
        # Handle 1D array sorting
        lst = flatten_args(range_)
        return sorted(lst, reverse=not is_ascending)


def SORTN(range_, n=1, display_ties_mode=0, sort_column=1, is_ascending=True):
    raise NotImplementedError("SORTN() not implemented yet")


def UNIQUE(range_):
    lst = flatten_args(range_)
    seen = set()
    result = []
    for v in lst:
        if v not in seen:
            seen.add(v)
            result.append(v)
    return result
