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

try:
    from pyspread.lib.spreadsheet.array import *
    from pyspread.lib.spreadsheet.database import *
    from pyspread.lib.spreadsheet.date import *
    from pyspread.lib.spreadsheet.engineering import *
    from pyspread.lib.spreadsheet.filter import *
    from pyspread.lib.spreadsheet.financial import *
    from pyspread.lib.spreadsheet.info import *
    from pyspread.lib.spreadsheet.logical import *
    from pyspread.lib.spreadsheet.lookup import *
    from pyspread.lib.spreadsheet.math import *
    from pyspread.lib.spreadsheet.operator import *
    from pyspread.lib.spreadsheet.parser import *
    from pyspread.lib.spreadsheet.statistical import *
    from pyspread.lib.spreadsheet.text import *
    from pyspread.lib.spreadsheet.web import *
except ImportError:
    from lib.spreadsheet.array import *
    from lib.spreadsheet.database import *
    from lib.spreadsheet.date import *
    from lib.spreadsheet.engineering import *
    from lib.spreadsheet.filter import *
    from lib.spreadsheet.financial import *
    from lib.spreadsheet.info import *
    from lib.spreadsheet.logical import *
    from lib.spreadsheet.lookup import *
    from lib.spreadsheet.math import *
    from lib.spreadsheet.operator import *
    from lib.spreadsheet.parser import *
    from lib.spreadsheet.statistical import *
    from lib.spreadsheet.text import *
    from lib.spreadsheet.web import *

def T(*args, **kwargs):
    raise NotImplementedError(
        "T() is ambiguous: use T_TEXT() for the text function "
        "or T_STAT.DIST() / T_STAT.INV() / T_STAT.TEST() for statistical functions"
    )


__all__ = (
    _ARRAY_FUNCTIONS +
    _DATABASE_FUNCTIONS +
    _DATE_FUNCTIONS +
    _ENGINEERING_FUNCTIONS +
    _FILTER_FUNCTIONS +
    _FINANCIAL_FUNCTIONS +
    _INFO_FUNCTIONS +
    _LOGICAL_FUNCTIONS +
    _LOOKUP_FUNCTIONS +
    _MATH_FUNCTIONS +
    _OPERATOR_FUNCTIONS +
    _PARSER_FUNCTIONS +
    _STATISTICAL_FUNCTIONS +
    _TEXT_FUNCTIONS +
    _WEB_FUNCTIONS +
    ['T']
)
