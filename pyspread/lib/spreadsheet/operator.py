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

_OPERATOR_FUNCTIONS = [
    'ADD', 'CONCAT', 'DIVIDE', 'EQ', 'GT', 'GTE', 'ISBETWEEN', 'LT', 'LTE', 'MINUS', 'MULTIPLY',
    'NE', 'POW', 'UMINUS', 'UNARY_PERCENT', 'UPLUS'
]
__all__ = _OPERATOR_FUNCTIONS + ["_OPERATOR_FUNCTIONS"]


def ADD(a, b):
    return a + b


def CONCAT(a, b):
    return str(a) + str(b)


def DIVIDE(a, b):
    return a / b


def EQ(a, b):
    return a == b


def GT(a, b):
    return a > b


def GTE(a, b):
    return a >= b


def ISBETWEEN(value, lower, upper, lower_inclusive, upper_inclusive):
    if value > lower or (lower_inclusive and value == lower):
        pass  # okay
    else:
        return False
    if value < upper or (upper_inclusive and value == upper):
        pass  # okay
    else:
        return False
    return True


def LT(a, b):
    return a < b


def LTE(a, b):
    return a <= b


def MINUS(a):
    return -a


def MULTIPLY(a, b):
    return a * b


def NE(a, b):
    return a != b


def POW(a, b):
    return a ** b


def UMINUS(a):
    return -a


def UNARY_PERCENT(a):
    return a / 100


def UPLUS(a):
    return +a
