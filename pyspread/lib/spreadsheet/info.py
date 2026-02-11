import re
from datetime import date, datetime, time

try:
    from pyspread.lib.pycellsheet import EmptyCell
    from pyspread.lib.spreadsheet.errors import SpreadsheetErrorNa
except ImportError:
    from lib.pycellsheet import EmptyCell
    from lib.spreadsheet.errors import SpreadsheetErrorNa

_INFO_FUNCTIONS = [
    'ERROR', 'ISBLANK', 'ISDATE', 'ISEMAIL', 'ISERR', 'ISERROR', 'ISFORMULA', 'ISLOGICAL',
    'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISREF', 'ISTEXT', 'N', 'NA', 'TYPE', 'CELL'
]
__all__ = _INFO_FUNCTIONS + ["_INFO_FUNCTIONS"]


class ERROR:
    @staticmethod
    def TYPE(error_val):
        raise NotImplementedError("ERROR.TYPE() not implemented yet")


def ISBLANK(value):
    return value == EmptyCell or value is None


def ISDATE(value):
    return isinstance(value, (date, datetime, time))


def ISEMAIL(value):
    if not isinstance(value, str):
        return False
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', value))


def ISERR(value):
    return isinstance(value, Exception) and not isinstance(value, SpreadsheetErrorNa)


def ISERROR(value):
    return isinstance(value, Exception)


def ISFORMULA(cell_reference):
    # Requires evaluation engine context
    raise NotImplementedError("ISFORMULA() requires evaluation engine support")


def ISLOGICAL(value):
    return isinstance(value, bool)


def ISNA(value):
    return isinstance(value, SpreadsheetErrorNa)


def ISNONTEXT(value):
    return not isinstance(value, str)


def ISNUMBER(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value != EmptyCell


def ISREF(value):
    raise NotImplementedError("ISREF() requires evaluation engine support")


def ISTEXT(value):
    return isinstance(value, str)


def N(value):
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    return 0


def NA():
    raise SpreadsheetErrorNa()


def TYPE(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return 1
    if isinstance(value, str):
        return 2
    if isinstance(value, bool):
        return 4
    if isinstance(value, Exception):
        return 16
    if isinstance(value, (list, tuple)):
        return 64
    return 1


def CELL(info_type, reference):
    # Requires evaluation engine context
    raise NotImplementedError("CELL() requires evaluation engine support")
