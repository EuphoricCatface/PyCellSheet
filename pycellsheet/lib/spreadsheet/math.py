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

import math
import typing
import random

try:
    from pycellsheet.lib.pycellsheet import EmptyCell, Range, flatten_args, RangeOutput
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


def COUNTIF(range_: Range, criterion_func):
    """Count items in a range that satisfy ``criterion_func``."""
    vals = range_.flatten()
    return sum(1 for v in vals if criterion_func(v))


def COUNTIFS(*range_crit_pairs):
    """
    COUNTIFS((range1, crit1), (range2, crit2), ...)
      Returns the count of items that pass all criteria in their respective ranges.
      We assume all ranges have the same length.
    """
    if not range_crit_pairs:
        raise ValueError("COUNTIFS: no criteria provided")
    # Start with None => uninitialized set of valid indices
    valid_indices = None
    lengths_checked = None

    # Process each (range_, crit)
    for (rng, crit) in range_crit_pairs:
        r_vals = rng.lst
        if lengths_checked is None:
            lengths_checked = len(r_vals)
            valid_indices = set(range(lengths_checked))
        else:
            if len(r_vals) != lengths_checked:
                raise ValueError("COUNTIFS: ranges differ in length")

        local_indices = {i for i, v in enumerate(r_vals) if crit(v)}
        valid_indices &= local_indices
        if not valid_indices:
            return 0
    return len(valid_indices) if valid_indices is not None else 0


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
    raise NotImplementedError("EVEN() not implemented yet")


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
        raise NotImplementedError("GAMMALN() not implemented yet")

    @staticmethod
    def PRECISE(value):
        raise NotImplementedError("GAMMALN.PRECISE() not implemented yet")


def GCD(*integers):
    return math.gcd(*integers)


def IMLN(value):
    raise NotImplementedError("IMLN() not implemented yet")


def IMPOWER(complex_base, exponent):
    raise NotImplementedError("IMPOWER() not implemented yet")


def IMSQRT(complex_number):
    raise NotImplementedError("IMSQRT() not implemented yet")


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
    raise NotImplementedError("MROUND() not implemented yet")


def MULTINOMIAL(x, y):
    raise NotImplementedError("MULTINOMIAL() not implemented yet")


def MUNIT(x, y):
    raise NotImplementedError("MUNIT() not implemented yet")


def ODD(x, y):
    raise NotImplementedError("ODD() not implemented yet")


def PI():
    return math.pi


def POWER(x, y):
    """
    Raises 'base' to the given 'exponent'.
    """
    return x ** y


def PRODUCT(*args):
    lst = flatten_args(*args)
    if not lst:
        return 0
    return math.prod(lst)


def QUOTIENT(numerator, denominator):
    """
    Returns the integer portion of the division numerator / denominator.
    Typically equivalent to trunc(numerator / denominator).
    Some spreadsheets do floor for positive and negative differently,
    so watch for sign behavior. We'll do 'trunc' here.
    """
    return math.trunc(numerator / denominator)


def RADIANS(x):
    return x / 180 * math.pi


def RAND():
    return random.random()


def RANDARRAY(row, column):
    rangeoutput_lst = []
    for i in range(row):
        for j in range(column):
            rangeoutput_lst.append(random.random())
    rtn = RangeOutput(column, rangeoutput_lst)
    return rtn


def RANDBETWEEN(x, y):
    return x + (y-x) * random.random()


def ROUND(value, decimals):
    """
    ROUND(value, decimals) rounds 'value' to 'decimals' decimal places.
    In Python, we can use built-in round().
    Note that Python's round uses "Banker's Rounding" for .5 cases,
    while Excel/LibreOffice typically use "round half away from zero."
    If you want to emulate the spreadsheet's rounding more precisely,
    you'd implement a custom approach.
    """
    # Simple approach using built-in round():
    return round(value, decimals)


def ROUNDDOWN(value, decimals):
    """
    ROUNDDOWN(value, decimals) truncates (towards zero) the number at 'decimals' places.
    In many spreadsheets, "Round Down" means floor for positive, but for negative
    values, it effectively moves toward zero. So we emulate that logic here.
    """
    multiplier = 10 ** decimals
    if value >= 0:
        return math.floor(value * multiplier) / multiplier
    else:
        return math.ceil(value * multiplier) / multiplier


def ROUNDUP(value, decimals):
    """
    ROUNDUP(value, decimals) rounds away from zero at 'decimals' places.
    For positive values, it's math.ceil; for negative, math.floor.
    """
    multiplier = 10 ** decimals
    if value >= 0:
        return math.ceil(value * multiplier) / multiplier
    else:
        return math.floor(value * multiplier) / multiplier


def SEC(x):
    return 1/math.cos(x)


def SECH(x):
    return 1/math.cosh(x)


def SEQUENCE(x, y):
    raise NotImplementedError("SEQUENCE() not implemented yet")


def SERIESSUM(x, y):
    raise NotImplementedError("SERIESSUM() not implemented yet")


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
    raise NotImplementedError("SUBTOTAL() not implemented yet")


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


def SUMIFS(x, y):
    raise NotImplementedError("SUMIFS() not implemented yet")


def SUMSQ(*args):
    """
    SUMSQ(...) = sum of squares of all numeric values in args.
    Example:
      =SUMSQ(A1:A3, B1:B3)
    If flatten_args returns [1, 2, 3], result = 1^2 + 2^2 + 3^2 = 14.
    """
    vals = flatten_args(*args)
    return sum(x * x for x in vals)


def TAN(x):
    return math.tan(x)


def TANH(x):
    return math.tanh(x)


def TRUNC(value, places=0):
    return int(value  / 10 ** -places) * 10 ** -places
