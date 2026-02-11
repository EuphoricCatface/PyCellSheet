try:
    import numpy as np
except ImportError:
    np = None

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
    """Calculate values along an exponential trend."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use GROWTH")

    y = np.array(flatten_args(known_ys))
    log_y = np.log(y)

    if known_xs is None:
        x = np.arange(1, len(y) + 1)
    else:
        x = np.array(flatten_args(known_xs))

    if const:
        X = np.column_stack([x, np.ones(len(x))])
    else:
        X = x.reshape(-1, 1)

    # Calculate regression coefficients on log scale
    coeffs, residuals, rank, s = np.linalg.lstsq(X, log_y, rcond=None)

    # Predict for new_xs or original xs
    if new_xs is None:
        new_x = x
    else:
        new_x = np.array(flatten_args(new_xs))

    if const:
        X_new = np.column_stack([new_x, np.ones(len(new_x))])
    else:
        X_new = new_x.reshape(-1, 1)

    log_predictions = X_new @ coeffs
    predictions = np.exp(log_predictions)
    return list(predictions)


def HSTACK(*arrays):
    raise NotImplementedError("HSTACK() not implemented yet")


def LINEST(known_ys, known_xs=None, const=True, stats=False):
    """Calculate linear regression parameters."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use LINEST")

    y = np.array(flatten_args(known_ys))

    if known_xs is None:
        x = np.arange(1, len(y) + 1)
    else:
        x = np.array(flatten_args(known_xs))

    if const:
        # Add intercept term
        X = np.column_stack([x, np.ones(len(x))])
    else:
        X = x.reshape(-1, 1)

    # Solve using least squares
    coeffs, residuals, rank, s = np.linalg.lstsq(X, y, rcond=None)

    if stats:
        # Return extended statistics (simplified)
        n = len(y)
        k = len(coeffs)
        y_pred = X @ coeffs
        sse = np.sum((y - y_pred) ** 2)
        sst = np.sum((y - np.mean(y)) ** 2)
        r_squared = 1 - sse / sst if sst != 0 else 0
        se = np.sqrt(sse / (n - k)) if n > k else 0

        # Return as flat list: [coeffs..., intercept, se, r_squared, ...]
        result = list(coeffs) + [se, r_squared]
        return result
    else:
        return list(coeffs)


def LOGEST(known_ys, known_xs=None, const=True, stats=False):
    """Calculate exponential regression parameters."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use LOGEST")

    y = np.array(flatten_args(known_ys))

    # Take log of y values for exponential fit
    log_y = np.log(y)

    if known_xs is None:
        x = np.arange(1, len(y) + 1)
    else:
        x = np.array(flatten_args(known_xs))

    if const:
        X = np.column_stack([x, np.ones(len(x))])
    else:
        X = x.reshape(-1, 1)

    # Solve using least squares
    coeffs, residuals, rank, s = np.linalg.lstsq(X, log_y, rcond=None)

    # Convert coefficients back to exponential form
    exp_coeffs = [np.exp(c) for c in coeffs]

    if stats:
        # Return extended statistics (simplified)
        n = len(y)
        k = len(coeffs)
        log_y_pred = X @ coeffs
        sse = np.sum((log_y - log_y_pred) ** 2)
        sst = np.sum((log_y - np.mean(log_y)) ** 2)
        r_squared = 1 - sse / sst if sst != 0 else 0
        se = np.sqrt(sse / (n - k)) if n > k else 0

        result = exp_coeffs + [se, r_squared]
        return result
    else:
        return exp_coeffs


def MAKEARRAY(rows, cols, func):
    raise NotImplementedError("MAKEARRAY() not implemented yet")


def MAP(array, func):
    raise NotImplementedError("MAP() not implemented yet")


def MDETERM(matrix: Range):
    """Calculate the determinant of a matrix."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use MDETERM")

    # Convert Range to numpy array
    rows = len(matrix)
    cols = matrix.width
    arr = np.array([[matrix[r][c] for c in range(cols)] for r in range(rows)])

    return float(np.linalg.det(arr))


def MINVERSE(matrix: Range):
    """Calculate the inverse of a matrix."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use MINVERSE")

    # Convert Range to numpy array
    rows = len(matrix)
    cols = matrix.width
    arr = np.array([[matrix[r][c] for c in range(cols)] for r in range(rows)])

    inv = np.linalg.inv(arr)

    # Convert back to flat list for RangeOutput
    result = []
    for r in range(inv.shape[0]):
        for c in range(inv.shape[1]):
            result.append(float(inv[r, c]))

    return RangeOutput(inv.shape[0], result)


def MMULT(matrix1: Range, matrix2: Range):
    """Multiply two matrices."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use MMULT")

    # Convert ranges to numpy arrays
    rows1 = len(matrix1)
    cols1 = matrix1.width
    arr1 = np.array([[matrix1[r][c] for c in range(cols1)] for r in range(rows1)])

    rows2 = len(matrix2)
    cols2 = matrix2.width
    arr2 = np.array([[matrix2[r][c] for c in range(cols2)] for r in range(rows2)])

    result_arr = np.matmul(arr1, arr2)

    # Convert back to flat list for RangeOutput
    result = []
    for r in range(result_arr.shape[0]):
        for c in range(result_arr.shape[1]):
            result.append(float(result_arr[r, c]))

    return RangeOutput(result_arr.shape[0], result)


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
    """Calculate values along a linear trend."""
    if np is None:
        raise NotImplementedError("Install `numpy` python package to use TREND")

    y = np.array(flatten_args(known_ys))

    if known_xs is None:
        x = np.arange(1, len(y) + 1)
    else:
        x = np.array(flatten_args(known_xs))

    if const:
        X = np.column_stack([x, np.ones(len(x))])
    else:
        X = x.reshape(-1, 1)

    # Calculate regression coefficients
    coeffs, residuals, rank, s = np.linalg.lstsq(X, y, rcond=None)

    # Predict for new_xs or original xs
    if new_xs is None:
        new_x = x
    else:
        new_x = np.array(flatten_args(new_xs))

    if const:
        X_new = np.column_stack([new_x, np.ones(len(new_x))])
    else:
        X_new = new_x.reshape(-1, 1)

    predictions = X_new @ coeffs
    return list(predictions)


def VSTACK(*arrays):
    raise NotImplementedError("VSTACK() not implemented yet")


def WRAPCOLS(vector, wrap_count, pad_with=None):
    raise NotImplementedError("WRAPCOLS() not implemented yet")


def WRAPROWS(vector, wrap_count, pad_with=None):
    raise NotImplementedError("WRAPROWS() not implemented yet")
