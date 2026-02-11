try:
    from pyspread.lib.pycellsheet import Range, RangeOutput, flatten_args
except ImportError:
    from lib.pycellsheet import Range, RangeOutput, flatten_args

_ARRAY_FUNCTIONS = [
    'ARRAY_CONSTRAIN', 'BYCOL', 'BYROW', 'CHOOSECOLS', 'CHOOSEROWS', 'FLATTEN', 'FREQUENCY',
    'GROWTH', 'HSTACK', 'LINEST', 'LOGEST', 'MAKEARRAY', 'MAP', 'MDETERM', 'MINVERSE', 'MMULT',
    'REDUCE', 'SCAN', 'SUMPRODUCT', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'TOCOL', 'TOROW', 'TRANSPOSE',
    'TREND', 'VSTACK', 'WRAPCOLS', 'WRAPROWS'
]
__all__ = _ARRAY_FUNCTIONS + ["_ARRAY_FUNCTIONS"]


def ARRAY_CONSTRAIN(input_range, num_rows, num_cols):
    raise NotImplementedError("ARRAY_CONSTRAIN() not implemented yet")


def BYCOL(array, func):
    raise NotImplementedError("BYCOL() not implemented yet")


def BYROW(array, func):
    raise NotImplementedError("BYROW() not implemented yet")


def CHOOSECOLS(array: Range, *col_nums):
    raise NotImplementedError("CHOOSECOLS() not implemented yet")


def CHOOSEROWS(array: Range, *row_nums):
    raise NotImplementedError("CHOOSEROWS() not implemented yet")


def FLATTEN(*args):
    return flatten_args(*args)


def FREQUENCY(data, bins):
    data_list = sorted(flatten_args(data))
    bin_list = sorted(flatten_args(bins))
    counts = [0] * (len(bin_list) + 1)
    for val in data_list:
        placed = False
        for i, b in enumerate(bin_list):
            if val <= b:
                counts[i] += 1
                placed = True
                break
        if not placed:
            counts[-1] += 1
    return counts


def GROWTH(known_ys, known_xs=None, new_xs=None, const=True):
    raise NotImplementedError("GROWTH() not implemented yet")


def HSTACK(*arrays):
    raise NotImplementedError("HSTACK() not implemented yet")


def LINEST(known_ys, known_xs=None, const=True, stats=False):
    raise NotImplementedError("LINEST() not implemented yet")


def LOGEST(known_ys, known_xs=None, const=True, stats=False):
    raise NotImplementedError("LOGEST() not implemented yet")


def MAKEARRAY(rows, cols, func):
    raise NotImplementedError("MAKEARRAY() not implemented yet")


def MAP(array, func):
    raise NotImplementedError("MAP() not implemented yet")


def MDETERM(matrix: Range):
    raise NotImplementedError("MDETERM() not implemented yet")


def MINVERSE(matrix: Range):
    raise NotImplementedError("MINVERSE() not implemented yet")


def MMULT(matrix1: Range, matrix2: Range):
    raise NotImplementedError("MMULT() not implemented yet")


def REDUCE(initial_value, array, func):
    raise NotImplementedError("REDUCE() not implemented yet")


def SCAN(initial_value, array, func):
    raise NotImplementedError("SCAN() not implemented yet")


def SUMPRODUCT(*arrays):
    lists = [flatten_args(a) for a in arrays]
    length = len(lists[0])
    for lst in lists[1:]:
        if len(lst) != length:
            raise ValueError("All arrays must have the same dimensions")
    total = 0
    for i in range(length):
        product = 1
        for lst in lists:
            product *= lst[i]
        total += product
    return total


def SUMX2MY2(array_x, array_y):
    xs = flatten_args(array_x)
    ys = flatten_args(array_y)
    return sum(x**2 - y**2 for x, y in zip(xs, ys))


def SUMX2PY2(array_x, array_y):
    xs = flatten_args(array_x)
    ys = flatten_args(array_y)
    return sum(x**2 + y**2 for x, y in zip(xs, ys))


def SUMXMY2(array_x, array_y):
    xs = flatten_args(array_x)
    ys = flatten_args(array_y)
    return sum((x - y)**2 for x, y in zip(xs, ys))


def TOCOL(array, ignore=0, scan_by_column=False):
    raise NotImplementedError("TOCOL() not implemented yet")


def TOROW(array, ignore=0, scan_by_column=False):
    raise NotImplementedError("TOROW() not implemented yet")


def TRANSPOSE(array: Range):
    rows = len(array)
    cols = array.width
    result = []
    for c in range(cols):
        for r in range(rows):
            result.append(array[r][c])
    return RangeOutput(rows, result)


def TREND(known_ys, known_xs=None, new_xs=None, const=True):
    raise NotImplementedError("TREND() not implemented yet")


def VSTACK(*arrays):
    raise NotImplementedError("VSTACK() not implemented yet")


def WRAPCOLS(vector, wrap_count, pad_with=None):
    raise NotImplementedError("WRAPCOLS() not implemented yet")


def WRAPROWS(vector, wrap_count, pad_with=None):
    raise NotImplementedError("WRAPROWS() not implemented yet")
