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
