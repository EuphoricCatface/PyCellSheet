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

import cmath
import math

try:
    from pycellsheet.lib.pycellsheet import flatten_args
except ImportError:
    from lib.pycellsheet import flatten_args

_ENGINEERING_FUNCTIONS = [
    'BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR',
    'COMPLEX', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'DELTA', 'ERF', 'GESTEP', 'HEX2BIN',
    'HEX2DEC', 'HEX2OCT', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH',
    'IMCOT', 'IMCOTH', 'IMCSC', 'IMCSCH', 'IMDIV', 'IMEXP', 'IMLOG', 'IMLOG10', 'IMLOG2',
    'IMPRODUCT', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSUB', 'IMSUM', 'IMTAN',
    'IMTANH', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX'
]
__all__ = _ENGINEERING_FUNCTIONS + ["_ENGINEERING_FUNCTIONS"]


def BIN2DEC(x):
    return int(x, 2)


def BIN2HEX(x):
    return hex(int(x, 2))


def BIN2OCT(x):
    return oct(int(x, 2))


def BITAND(a, b):
    return a & b


def BITLSHIFT(a, b):
    return a << b


def BITOR(a, b):
    return a | b


def BITRSHIFT(a, b):
    return a >> b


def BITXOR(a, b):
    return a ^ b


def COMPLEX(a, b):
    return complex(a, b)


def DEC2BIN(x):
    return bin(x)


def DEC2HEX(x):
    return hex(x)


def DEC2OCT(x):
    return oct(x)


def DELTA(a, b=0):
    return int(a==b)


class ERF:
    def __new__(cls, lower_limit, upper_limit=None):
        return cls.PRECISE(lower_limit, upper_limit)

    @staticmethod
    def PRECISE(lower_limit, upper_limit=None):
        if upper_limit is None:
            return math.erf(lower_limit)
        return math.erf(upper_limit) - math.erf(lower_limit)


def GESTEP(x, step=0):
    return 1 if x >= step else 0


def HEX2BIN(x):
    return bin(int(x, 16))


def HEX2DEC(x):
    return int(x, 16)


def HEX2OCT(x):
    return oct(int(x, 16))


def IMABS(z):
    return abs(z)


def IMAGINARY(z: complex):
    return z.imag


def IMARGUMENT(z: complex):
    return cmath.phase(z)


def IMCONJUGATE(z: complex):
    return complex(z.real, -z.imag)


def IMCOS(z):
    return cmath.cos(z)


def IMCOSH(z):
    return cmath.cosh(z)


def IMCOT(z):
    return 1/cmath.tan(z)


def IMCOTH(z):
    return 1/cmath.tanh(z)


def IMCSC(z):
    return 1/cmath.sin(z)


def IMCSCH(z):
    return 1/cmath.sinh(z)


def IMDIV(a, b):
    return a / b


def IMEXP(z):
    return math.e ** z


def IMLOG(value, base):
    return cmath.log(value, base)


def IMLOG10(value):
    return cmath.log(value, 10)


def IMLOG2(value):
    return cmath.log(value, 2)


def IMPRODUCT(a, b):
    return a * b


def IMREAL(z: complex):
    return z.real


def IMSEC(z):
    return 1/cmath.cos(z)


def IMSECH(z):
    return 1/cmath.cosh(z)


def IMSIN(z):
    return cmath.sin(z)


def IMSINH(z):
    return cmath.sinh(z)


def IMSUB(a, b):
    return a - b


def IMSUM(*args):
    return sum(flatten_args(*args))


def IMTAN(z):
    return cmath.tan(z)


def IMTANH(z):
    return cmath.tanh(z)


def OCT2BIN(x):
    return bin(int(x, 8))


def OCT2DEC(x):
    return int(x, 8)


def OCT2HEX(x):
    return hex(int(x, 8))