import datetime
import re

try:
    from pyspread.lib.spreadsheet import errors
    from pyspread.lib.pycellsheet import EmptyCell
except ImportError:
    from lib.spreadsheet import errors
    from lib.pycellsheet import EmptyCell

_INFO_FUNCTIONS = [
    'ERROR', 'ISBLANK', 'ISDATE', 'ISEMAIL', 'ISERR', 'ISERROR', 'ISFORMULA', 'ISLOGICAL',
    'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISREF', 'ISTEXT', 'N', 'NA', 'TYPE', 'CELL'
]
__all__ = _INFO_FUNCTIONS + ["_INFO_FUNCTIONS"]


class ERROR:
    @staticmethod
    def TYPE(value):
        error_code_map = {
            "#NULL!": 1,
            "#DIV/0!": 2,
            "#VALUE!": 3,
            "#REF!": 4,
            "#NAME?": 5,
            "#NUM!": 6,
            "#N/A": 7
        }
        if isinstance(value, errors.SpreadsheetErrorBase):
            # check the code
            code_str = value.cell_output()
            if code_str in error_code_map:
                return error_code_map[code_str]
            else:
                # unknown error => return 8
                return 8
        if isinstance(value, Exception):
            # Python error => return 8
            return 8
        # Not an error => return #N/A
        raise errors.SpreadsheetErrorNa


def ISBLANK(value) -> bool:
    """
    ISBLANK(value) => True if value is an 'empty cell' (e.g. EmptyCell), else False.
    """
    if value == EmptyCell:
        return True
    return False


def ISDATE(value) -> bool:
    """
    ISDATE(value) => True if value is a date/datetime object, else False.
    """
    return isinstance(value, (datetime.date, datetime.datetime))


_EMAIL_REGEX = re.compile(r"^[^@]+@[^@]+\.[^@]+$")
def ISEMAIL(value) -> bool:
    """
    ISEMAIL(value) => True if value is a string that looks like a valid email, else False.
    """
    if not isinstance(value, str):
        return False
    return bool(_EMAIL_REGEX.match(value))


def ISERR(value) -> bool:
    """
    ISERR(value) => True if value is an error but not #N/A
    """
    if isinstance(value, errors.SpreadsheetErrorBase):
        return value.cell_output() != "#N/A"
    if isinstance(value, Exception):
        return True
    return False


def ISERROR(value) -> bool:
    """
    ISERROR(value) => True if value is any error, including Python exceptions
    """
    return isinstance(value, Exception)


def ISFORMULA(a, b):
    raise NotImplementedError("ISFORMULA() not implemented yet")


def ISLOGICAL(value) -> bool:
    """
    ISLOGICAL(value) => True if value is a bool (True/False), else False.
    """
    return isinstance(value, bool)


def ISNA(value) -> bool:
    """
    ISNA(value) => True if value is the #N/A error, else False.
    """
    if isinstance(value, errors.SpreadsheetErrorBase):
        return value.cell_output() == "#N/A"
    return False


def ISNONTEXT(value) -> bool:
    """
    ISNONTEXT(value) => True if value is not a text string.
    """
    return not isinstance(value, str)


def ISNUMBER(value) -> bool:
    """
    ISNUMBER(value) => True if value is a numeric type, else False.
    """
    return isinstance(value, (int, float))


def ISREF(a, b):
    raise NotImplementedError("ISREF() not implemented yet")


def ISTEXT(value) -> bool:
    """
    ISTEXT(value) => True if value is a string, else False.
    """
    return isinstance(value, str)


def N(value):
    """
    N(value):
      - If number => return the number
      - If True => 1, If False => 0
      - If error => return the error
      - Else => 0
    """
    if isinstance(value, (errors.SpreadsheetErrorBase, Exception)):
        return value  # pass error along unchanged
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, bool):
        return 1 if value else 0
    # else => 0
    return 0


def NA():
    """
    NA() => returns the #N/A error value.
    """
    return errors.SpreadsheetErrorNa


def TYPE(value) -> int:
    """
    TYPE(value):
      1 => number
      2 => text
      4 => logical
      16 => error
      64 => range/array reference (depending on your design)
    If none of the above, might mimic Excel's #VALUE! or return 0 or something.
    """
    if isinstance(value, (errors.SpreadsheetErrorBase, Exception)):
        return 16
    if isinstance(value, (int, float)):
        return 1
    if isinstance(value, str):
        return 2
    if isinstance(value, bool):
        return 4
    # If you consider a Range or list => 64
    if isinstance(value, (Range, RangeOutput)):
        return 64
    # If something else => could raise an error or return something.
    # Let's return 0 or #VALUE! error. We'll do 0 for demonstration:
    return 0


def CELL(a, b):
    raise NotImplementedError("CELL() not implemented yet")