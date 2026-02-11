try:
    from pyspread.lib.pycellsheet import Range, flatten_args
except ImportError:
    from lib.pycellsheet import Range, flatten_args

_LOOKUP_FUNCTIONS = [
    'ADDRESS', 'CHOOSE', 'COLUMN', 'COLUMNS', 'FORMULATEXT', 'GETPIVOTDATA', 'HLOOKUP',
    'INDEX', 'INDIRECT', 'LOOKUP', 'MATCH', 'OFFSET', 'ROW', 'ROWS', 'VLOOKUP', 'XLOOKUP'
]
__all__ = _LOOKUP_FUNCTIONS + ["_LOOKUP_FUNCTIONS"]


def ADDRESS(row, column, abs_num=1, a1=True, sheet_text=""):
    row = int(row)
    col = int(column)
    col_str = ""
    c = col
    while c > 0:
        c, remainder = divmod(c - 1, 26)
        col_str = chr(65 + remainder) + col_str
    if a1:
        match abs_num:
            case 1:
                ref = f"${col_str}${row}"
            case 2:
                ref = f"{col_str}${row}"
            case 3:
                ref = f"${col_str}{row}"
            case 4:
                ref = f"{col_str}{row}"
            case _:
                raise ValueError("abs_num must be 1, 2, 3, or 4")
    else:
        match abs_num:
            case 1:
                ref = f"R{row}C{col}"
            case 2:
                ref = f"R{row}C[{col}]"
            case 3:
                ref = f"R[{row}]C{col}"
            case 4:
                ref = f"R[{row}]C[{col}]"
            case _:
                raise ValueError("abs_num must be 1, 2, 3, or 4")
    if sheet_text:
        return f"'{sheet_text}'!{ref}"
    return ref


def CHOOSE(index, *values):
    idx = int(index)
    if idx < 1 or idx > len(values):
        raise ValueError(f"Index {idx} is out of range (1-{len(values)})")
    return values[idx - 1]


def COLUMN(reference=None):
    # Without arguments, requires evaluation engine context
    if reference is None:
        raise NotImplementedError("COLUMN() without arguments requires evaluation engine support")
    if isinstance(reference, Range):
        return reference.width
    raise NotImplementedError("COLUMN() requires evaluation engine support")


def COLUMNS(range_: Range):
    return range_.width


def FORMULATEXT(reference):
    # Requires evaluation engine context
    raise NotImplementedError("FORMULATEXT() requires evaluation engine support")


def GETPIVOTDATA(*args):
    raise NotImplementedError("GETPIVOTDATA() not implemented yet")


def HLOOKUP(search_key, range_: Range, index, is_sorted=True):
    first_row = range_[0]
    for col_idx in range(range_.width):
        val = first_row[col_idx]
        if is_sorted:
            if val == search_key:
                return range_[int(index) - 1][col_idx]
        else:
            if val == search_key:
                return range_[int(index) - 1][col_idx]
    if is_sorted:
        # For sorted data, find the largest value <= search_key
        best_col = None
        for col_idx in range(range_.width):
            val = first_row[col_idx]
            if val <= search_key:
                best_col = col_idx
        if best_col is not None:
            return range_[int(index) - 1][best_col]
    raise ValueError(f"HLOOKUP: '{search_key}' not found")


def INDEX(reference, row, column=None):
    if isinstance(reference, Range):
        r = int(row)
        if column is not None:
            c = int(column)
            return reference[r - 1][c - 1]
        return reference[r - 1]
    if isinstance(reference, (list, tuple)):
        return reference[int(row) - 1]
    raise ValueError("INDEX: reference must be a Range or list")


def INDIRECT(reference_string):
    # Requires evaluation engine context
    raise NotImplementedError("INDIRECT() requires evaluation engine support")


def LOOKUP(search_key, search_range, result_range=None):
    if result_range is None:
        result_range = search_range
    search_list = flatten_args(search_range)
    result_list = flatten_args(result_range)
    best_idx = None
    for i, v in enumerate(search_list):
        if v <= search_key:
            best_idx = i
    if best_idx is not None:
        return result_list[best_idx]
    raise ValueError(f"LOOKUP: '{search_key}' not found")


def MATCH(search_key, search_range, search_type=1):
    lst = flatten_args(search_range)
    if search_type == 0:
        for i, v in enumerate(lst):
            if v == search_key:
                return i + 1
        raise ValueError(f"MATCH: '{search_key}' not found")
    elif search_type == 1:
        best_idx = None
        for i, v in enumerate(lst):
            if v <= search_key:
                best_idx = i
        if best_idx is not None:
            return best_idx + 1
        raise ValueError(f"MATCH: '{search_key}' not found")
    else:  # search_type == -1
        best_idx = None
        for i, v in enumerate(lst):
            if v >= search_key:
                best_idx = i
        if best_idx is not None:
            return best_idx + 1
        raise ValueError(f"MATCH: '{search_key}' not found")


def OFFSET(reference, rows, cols, height=None, width=None):
    # Requires evaluation engine context
    raise NotImplementedError("OFFSET() requires evaluation engine support")


def ROW(reference=None):
    # Without arguments, requires evaluation engine context
    if reference is None:
        raise NotImplementedError("ROW() without arguments requires evaluation engine support")
    raise NotImplementedError("ROW() requires evaluation engine support")


def ROWS(range_: Range):
    return len(range_)


def VLOOKUP(search_key, range_: Range, index, is_sorted=True):
    for row_idx in range(len(range_)):
        val = range_[row_idx][0]
        if not is_sorted and val == search_key:
            return range_[row_idx][int(index) - 1]
        if is_sorted and val == search_key:
            return range_[row_idx][int(index) - 1]
    if is_sorted:
        best_row = None
        for row_idx in range(len(range_)):
            val = range_[row_idx][0]
            if val <= search_key:
                best_row = row_idx
        if best_row is not None:
            return range_[best_row][int(index) - 1]
    raise ValueError(f"VLOOKUP: '{search_key}' not found")


def XLOOKUP(search_key, lookup_range, return_range, if_not_found=None, match_mode=0, search_mode=1):
    lookup_list = flatten_args(lookup_range)
    return_list = flatten_args(return_range)
    if match_mode == 0:
        for i, v in enumerate(lookup_list):
            if v == search_key:
                return return_list[i]
    elif match_mode == -1:
        best_idx = None
        for i, v in enumerate(lookup_list):
            if v <= search_key:
                if best_idx is None or v > lookup_list[best_idx]:
                    best_idx = i
        if best_idx is not None:
            return return_list[best_idx]
    elif match_mode == 1:
        best_idx = None
        for i, v in enumerate(lookup_list):
            if v >= search_key:
                if best_idx is None or v < lookup_list[best_idx]:
                    best_idx = i
        if best_idx is not None:
            return return_list[best_idx]
    if if_not_found is not None:
        return if_not_found
    raise ValueError(f"XLOOKUP: '{search_key}' not found")
