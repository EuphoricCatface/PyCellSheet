import math
import typing
import random

try:
    from pyspread.lib.pycellsheet import EmptyCell, Range, flatten_args, RangeOutput
except ImportError:
    from lib.pycellsheet import EmptyCell, Range, flatten_args, RangeOutput

_MATH_FUNCTIONS = [
    'ABS', 'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'BASE',
    'CEILING', 'COMBIN', 'COMBINA', 'COS', 'COSH', 'COT', 'COTH', 'COUNTBLANK', 'COUNTIF',
    'COUNTIFS', 'COUNTUNIQUE', 'CSC', 'CSCH', 'DECIMAL', 'DEGREES', 'ERFC', 'EVEN', 'EXP', 'FACT',
    'FACTDOUBLE', 'FLOOR', 'GAMMALN', 'GCD', 'IMLN', 'IMPOWER', 'IMSQRT', 'INT', 'ISEVEN', 'ISO',
    'ISODD', 'LCM', 'LN', 'LOG', 'LOG10', 'MOD', 'MROUND', 'MULTINOMIAL', 'MUNIT', 'ODD', 'PI',
    'POWER', 'PRODUCT', 'QUOTIENT', 'RADIANS', 'RAND', 'RANDARRAY', 'RANDBETWEEN', 'ROUND',
    'ROUNDDOWN', 'ROUNDUP', 'SEC', 'SECH', 'SEQUENCE', 'SERIESSUM', 'SIGN', 'SIN', 'SINH', 'SQRT',
    'SQRTPI', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMSQ', 'TAN', 'TANH', 'TRUNC'
]
__all__ = _MATH_FUNCTIONS + ["_MATH_FUNCTIONS"]


def ABS(x):
    return abs(x)


def ACOS(x):
    return math.acos(x)


def ACOSH(x):
    return math.acosh(x)


def ACOT(x):
    return math.atan(1 / x) if x != 0 else math.pi / 2


def ACOTH(x):
    if abs(x) <= 1:
        raise ValueError("ACOTH is undefined for |x| ≤ 1")
    return 0.5 * math.log((x + 1) / (x - 1))


def ASIN(x):
    return math.asin(x)


def ASINH(x):
    return math.asinh(x)


def ATAN(x):
    return math.atan(x)


def ATAN2(y, x):
    return math.atan2(y, x)


def ATANH(x):
    return math.atanh(x)


def BASE(value, base, minimum_length=0):
    match base:
        case 2:
            result = bin(value)[2:]  # Strip "0b" prefix
        case 8:
            result = oct(value)[2:]  # Strip "0o" prefix
        case 10:
            result = str(value)  # Decimal base
        case 16:
            result = hex(value)[2:].upper()  # Strip "0x" prefix and capitalize
        case _:
            raise NotImplementedError("Only base value of 2, 8, 10, 16 are supported")

    # Add padding for positive numbers only
    if value >= 0 and len(result) < minimum_length:
        result = f"{result:0>{minimum_length}}"  # Right-align with '0' padding
    return result


class CEILING:
    def __new__(cls, value, factor=1):
        # Proper implementation is NYI
        return CEILING.PRECISE(value, factor)

    @staticmethod
    def MATH(value, significance=1, mode=0):
        if significance == 0:
            raise ValueError("Significance cannot be zero")

        significance = abs(significance)

        if value > 0 or mode == 0:
            # Use math.ceil for positive values
            return math.ceil(value / significance) * significance
        return math.floor(value / significance) * significance

    @staticmethod
    def PRECISE(number, significance=1):
        if significance == 0:
            raise ValueError("Significance cannot be zero")
        return math.ceil(number / significance) * significance


def COMBIN(n, k):
    return math.comb(n, k)


def COMBINA(n, k):
    return math.factorial(n+k-1)/(math.factorial(k)*math.factorial(n-1))


def COS(x):
    return math.cos(x)


def COSH(x):
    return math.cosh(x)


def COT(x):
    return 1/math.tan(x)


def COTH(x):
    return 1/math.tanh(x)


def COUNTBLANK(r: Range):
    return len(r.lst) - len(r.flatten())


def COUNTIF(r: Range, criterion: typing.Callable[[typing.Any], bool]):
    return len(list(filter(criterion, r.flatten())))


def COUNTIFS(*args):
    if len(args) % 2:
        raise ValueError("Number of arguments has to be even")
    current_range: Range
    filtered = []
    for i, arg in enumerate(args):
        if i % 2 == 0:
            current_range = arg
            continue
        filtered.extend(list(filter(arg, current_range.flatten())))
    return len(filtered)


def COUNTUNIQUE(r: Range):
    return len(set(r.flatten()) - {EmptyCell})


def CSC(x):
    return 1/math.sin(x)


def CSCH(x):
    return 1/math.sinh(x)


def DECIMAL(value, base):
    return int(value, base=base)


def DEGREES(x):
    return x / math.pi * 180


class ERFC:
    def __new__(cls, x):
        return math.erfc(x)

    @staticmethod
    def PRECISE(x):
        return math.erfc(x)


def EVEN(x):
    value = math.ceil(abs(x))
    if value % 2 == 1:
        value += 1
    return value if x >= 0 else -value


def EXP(x):
    return math.exp(x)


def FACT(x):
    return math.factorial(x)


def FACTDOUBLE(x):
    if x < 0 or not isinstance(x, int):
        raise ValueError("x must be non-negative integer")
    rtn = 1
    while x > 1:
        rtn *= x
        x -= 2
    return rtn


class FLOOR:
    def __new__(cls, value, factor=1):
        # Proper implementation is NYI
        return FLOOR.PRECISE(value, factor)

    @staticmethod
    def MATH(value, significance=1, mode=0):
        if significance == 0:
            raise ValueError("Significance cannot be zero")

        significance = abs(significance)

        if value > 0 or mode == 0:
            # Use math.ceil for positive values
            return math.floor(value / significance) * significance
        return math.ceil(value / significance) * significance

    @staticmethod
    def PRECISE(number, significance=1):
        if significance == 0:
            raise ValueError("Significance cannot be zero")
        return math.floor(number / significance) * significance


class GAMMALN:
    def __new__(cls, value):
        return math.lgamma(value)

    @staticmethod
    def PRECISE(value):
        return math.lgamma(value)


def GCD(*integers):
    return math.gcd(*integers)


def IMLN(value):
    import cmath
    return complex(cmath.log(complex(value)))


def IMPOWER(complex_base, exponent):
    return complex(complex(complex_base) ** exponent)


def IMSQRT(complex_number):
    import cmath
    return complex(cmath.sqrt(complex(complex_number)))


def INT(x):
    return int(x)


def ISEVEN(x):
    return (x % 2) == 0


class ISO:
    @staticmethod
    def CEILING(number, significance=1):
        return CEILING.PRECISE(number, significance)


def ISODD(x):
    return (x % 2) == 1


def LCM(*integers):
    return math.lcm(*integers)


def LN(x):
    return math.log(x)


def LOG(x, base):
    return math.log(x, base)


def LOG10(x):
    return math.log10(x)


def MOD(x, y):
    return x % y


def MROUND(value, multiple):
    if multiple == 0:
        return 0
    return round(value / multiple) * multiple


def MULTINOMIAL(*args):
    total = sum(args)
    result = math.factorial(total)
    for a in args:
        result //= math.factorial(a)
    return result


def MUNIT(dimension):
    result = []
    for i in range(dimension):
        for j in range(dimension):
            result.append(1 if i == j else 0)
    return RangeOutput(dimension, result)


def ODD(x):
    value = math.ceil(abs(x))
    if value % 2 == 0:
        value += 1
    return value if x >= 0 else -value


def PI():
    return math.pi


def POWER(x, y):
    return x ** y


def PRODUCT(*args):
    rtn = 1
    lst = flatten_args(*args)
    for v in lst:
        rtn *= v
    return rtn


def QUOTIENT(dividend, divisor):
    return int(dividend / divisor)


def RADIANS(x):
    return x / 180 * math.pi


def RAND():
    return random.random()


def RANDARRAY(x, y):
    rangeoutput_lst = []
    for i in range(x):
        for j in range(y):
            rangeoutput_lst.append(random.random())
    rtn = RangeOutput(x, rangeoutput_lst)
    return rtn


def RANDBETWEEN(x, y):
    return x + (y-x) * random.random()


def ROUND(value, places=0):
    return round(value, places)


def ROUNDDOWN(value, places=0):
    factor = 10 ** places
    return math.trunc(value * factor) / factor


def ROUNDUP(value, places=0):
    factor = 10 ** places
    if value >= 0:
        return math.ceil(value * factor) / factor
    else:
        return math.floor(value * factor) / factor


def SEC(x):
    return 1/math.cos(x)


def SECH(x):
    return 1/math.cosh(x)


def SEQUENCE(rows, columns=1, start=1, step=1):
    result = []
    val = start
    for i in range(rows):
        for j in range(columns):
            result.append(val)
            val += step
    return RangeOutput(columns, result)


def SERIESSUM(x, n, m, coefficients):
    lst = flatten_args(coefficients) if isinstance(coefficients, Range) else coefficients
    result = 0
    for i, a in enumerate(lst):
        result += a * (x ** (n + i * m))
    return result


def SIGN(x):
    if x == 0:
        return 0
    return 1 if x > 0 else -1


def SIN(x):
    return math.sin(x)


def SINH(x):
    return math.sinh(x)


def SQRT(x):
    return math.sqrt(x)


def SQRTPI(x):
    return math.sqrt(x * math.pi)


def SUBTOTAL(function_num, *ranges):
    lst = flatten_args(*ranges)
    match function_num:
        case 1 | 101:
            return sum(lst) / len(lst)
        case 2 | 102:
            return len([v for v in lst if isinstance(v, (int, float))])
        case 3 | 103:
            return len([v for v in lst if v != EmptyCell])
        case 4 | 104:
            return max(lst)
        case 5 | 105:
            return min(lst)
        case 6 | 106:
            result = 1
            for v in lst:
                result *= v
            return result
        case 7 | 107:
            import statistics
            return statistics.stdev(lst)
        case 8 | 108:
            import statistics
            return statistics.pstdev(lst)
        case 9 | 109:
            return sum(lst)
        case 10 | 110:
            import statistics
            return statistics.variance(lst)
        case 11 | 111:
            import statistics
            return statistics.pvariance(lst)
        case _:
            raise ValueError(f"Invalid function_num: {function_num}")


def SUM(*args):
    lst = flatten_args(*args)
    return sum(lst)


def SUMIF(r: Range, criterion, sum_range: Range | None = None):
    if sum_range is not None and (len(r) != len(sum_range) or r.width != sum_range.width):
        raise ValueError("The dimensions of r and sum_range don't match")
    if sum_range is None:
        sum_range = r
    sum_ = 0
    for i in range(len(r)):
        for j in range(r.width):
            if not criterion(r[i][j]):
                continue
            if sum_range[i][j] == EmptyCell:
                continue
            sum_ += sum_range[i][j]
    return sum_


def SUMIFS(sum_range: Range, *criteria_pairs):
    if len(criteria_pairs) % 2:
        raise ValueError("Number of criteria arguments has to be even")
    sum_ = 0
    flat_sum = sum_range.lst
    for idx in range(len(flat_sum)):
        include = True
        for i in range(0, len(criteria_pairs), 2):
            criteria_range = criteria_pairs[i]
            criterion = criteria_pairs[i + 1]
            if not criterion(criteria_range.lst[idx]):
                include = False
                break
        if include and flat_sum[idx] != EmptyCell:
            sum_ += flat_sum[idx]
    return sum_


def SUMSQ(*args):
    sum_ = 0
    for arg in args:
        if arg == EmptyCell:
            continue
        sum_ += arg * arg
    return sum_


def TAN(x):
    return math.tan(x)


def TANH(x):
    return math.tanh(x)


def TRUNC(value, places=0):
    return int(value  / 10 ** -places) * 10 ** -places