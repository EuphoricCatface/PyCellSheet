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

try:
    from pycellsheet.lib.pycellsheet import EmptyCell, flatten_args
    from pycellsheet.lib.spreadsheet.errors import SpreadsheetErrorNa
except ImportError:
    from lib.pycellsheet import EmptyCell, flatten_args
    from lib.spreadsheet.errors import SpreadsheetErrorNa

_LOGICAL_FUNCTIONS = [
    'AND', 'FALSE', 'IF', 'IFERROR', 'IFNA', 'IFS', 'LAMBDA', 'LET', 'NOT', 'OR', 'SWITCH',
    'TRUE', 'XOR'
]
__all__ = _LOGICAL_FUNCTIONS + ["_LOGICAL_FUNCTIONS"]


def AND(*args):
    return all(flatten_args(*args))


def FALSE():
    return False


def IF(expr, true, false=False):
    return true if expr else false


def IFERROR(value, value_if_error=EmptyCell):
    return value_if_error if isinstance(value, BaseException) else value


def IFNA(value, value_if_na):
    if isinstance(value, SpreadsheetErrorNa):
        return value_if_na
    return value


def IFS(*args):
    matched = False
    for i, arg in enumerate(args):
        if (i % 2) == 0:
            # arg is condition
            matched = bool(arg)
            continue
        if matched:
            return arg
    return SpreadsheetErrorNa()


def LAMBDA(name, formula_expression):
    raise NotImplementedError("LAMBDA() not implemented yet. Use Python's native lambda.")


def LET(a, b):
    raise NotImplementedError("LET() not implemented yet. Do you really need this function, when you have python?")


def NOT(expr):
    return not expr


def OR(*args):
    return any(flatten_args(*args))


def SWITCH(expr, *args):
    default = SpreadsheetErrorNa()
    if len(args) % 2:
        default = args[-1]
        args = args[:-1]

    matched = False
    for i, arg in enumerate(args):
        if (i % 2) == 0:
            # arg is condition
            if expr == arg:
                matched = True
            continue
        if matched:
            return arg
    return default



def TRUE():
    return True


def XOR(args):
    return len(list(filter(bool, flatten_args(args)))) % 2
