import math
import typing
import random

try:
    from pyspread.lib.pycellsheet import EmptyCell, Range
except ImportError:
    from lib.pycellsheet import EmptyCell, Range

__all__ = [
    'ABS', 'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'BASE',
    'CEILING', 'COMBIN', 'COMBINA', 'COS', 'COSH', 'COT', 'COTH', 'COUNTBLANK', 'COUNTIF',
    'COUNTIFS', 'COUNTUNIQUE', 'CSC', 'CSCH', 'DECIMAL', 'DEGREES', 'ERFC', 'EVEN', 'EXP', 'FACT',
    'FACTDOUBLE', 'FLOOR', 'GAMMALN', 'GCD', 'IMLN', 'IMPOWER', 'IMSQRT', 'INT', 'ISEVEN', 'ISO',
    'ISODD', 'LCM', 'LN', 'LOG', 'LOG10', 'MOD', 'MROUND', 'MULTINOMIAL', 'MUNIT', 'ODD', 'PI',
    'POWER', 'PRODUCT', 'QUOTIENT', 'RADIANS', 'RAND', 'RANDARRAY', 'RANDBETWEEN', 'ROUND',
    'ROUNDDOWN', 'ROUNDUP', 'SEC', 'SECH', 'SEQUENCE', 'SERIESSUM', 'SIGN', 'SIN', 'SINH', 'SQRT',
    'SQRTPI', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMSQ', 'TAN', 'TANH', 'TRUNC'
]


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
        CEILING.PRECISE(value, factor)

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
    r.flatten().count(EmptyCell)


def COUNTIF(r: Range, criterion: typing.Callable[[typing.Any], bool]):
    len(list(filter(criterion, r.flatten())))


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
        math.erfc(x)

    @staticmethod
    def PRECISE(x):
        return ERFC(x)


def EVEN(x):
    raise NotImplemented("EVEN() not implemented yet")


def EXP(x, y):
    return x ** y


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
        FLOOR.PRECISE(value, factor)

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
    @staticmethod
    def __new__(cls, value):
        raise NotImplemented("GAMMALN() not implemented yet")

    @staticmethod
    def PRECISE(value):
        raise NotImplemented("GAMMALN.PRECISE() not implemented yet")


def GCD(*integers):
    return math.gcd(*integers)


def IMLN(value):
    raise NotImplemented("IMLN() not implemented yet")


def IMPOWER(complex_base, exponent):
    raise NotImplemented("IMPOWER() not implemented yet")


def IMSQRT(complex_number):
    raise NotImplemented("IMSQRT() not implemented yet")


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


def MROUND(x, y):
    raise NotImplemented("MROUND() not implemented yet")


def MULTINOMIAL(x, y):
    raise NotImplemented("MULTINOMIAL() not implemented yet")


def MUNIT(x, y):
    raise NotImplemented("MUNIT() not implemented yet")


def ODD(x, y):
    raise NotImplemented("ODD() not implemented yet")


def PI():
    return math.pi


def POWER(x, y):
    return x ** y


def PRODUCT(*args):
    rtn = 1
    for arg in args:
        if arg == EmptyCell:
            continue
        rtn *= arg
    return rtn


def QUOTIENT(x, y):
    raise NotImplemented("QUOTIENT() not implemented yet")


def RADIANS(x):
    return x / 180 * math.pi


def RAND():
    return random.random()


def RANDARRAY(x, y):
    raise NotImplemented("RANDARRAY() not implemented yet")


def RANDBETWEEN(x, y):
    return x + (y-x) * random.random()


def ROUND(x, y):
    raise NotImplemented("ROUND() not implemented yet")


def ROUNDDOWN(x, y):
    raise NotImplemented("ROUNDDOWN() not implemented yet")


def ROUNDUP(x, y):
    raise NotImplemented("ROUNDUP() not implemented yet")


def SEC(x):
    return 1/math.cos(x)


def SECH(x):
    return 1/math.cosh(x)


def SEQUENCE(x, y):
    raise NotImplemented("SEQUENCE() not implemented yet")


def SERIESSUM(x, y):
    raise NotImplemented("SERIESSUM() not implemented yet")


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


def SUBTOTAL(x, y):
    raise NotImplemented("SUBTOTAL() not implemented yet")


def SUM(*args):
    lst = []
    for arg in args:
        if isinstance(arg, Range):
            lst.extend(arg.flatten())
            continue
        if isinstance(arg, list):
            lst.extend(arg)
            continue
        lst.append(arg)
    return sum(filter(lambda a: a != EmptyCell, lst))


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


def SUMIFS(x, y):
    raise NotImplemented("SUMIFS() not implemented yet")


def SUMSQ(*args):
    sum_ = 0
    for arg in args:
        if arg == EmptyCell:
            continue
        sum_ += arg * arg
    return sum_


def TAN(x):
    return math.tan(x)


def TANH(x, y):
    return math.tanh(x)


def TRUNC(value, places=0):
    return int(value  / 10 ** -places) * 10 ** -places