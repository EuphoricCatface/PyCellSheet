from datetime import date, datetime

_PARSER_FUNCTIONS = [
    'CONVERT', 'TO_DATE', 'TO_DOLLARS', 'TO_PERCENT', 'TO_PURE_NUMBER', 'TO_TEXT'
]
__all__ = _PARSER_FUNCTIONS + ["_PARSER_FUNCTIONS"]


def CONVERT(value, start_unit, end_unit):
    raise NotImplementedError("CONVERT() not implemented yet")


def TO_DATE(value):
    if isinstance(value, (date, datetime)):
        return value
    if isinstance(value, (int, float)):
        return date.fromordinal(int(value) + 693594)
    if isinstance(value, str):
        try:
            import dateutil.parser
        except ImportError:
            raise NotImplementedError("Install `dateutil` python package to use TO_DATE")
        return dateutil.parser.parse(value).date()
    raise ValueError(f"Cannot convert {type(value)} to date")


def TO_DOLLARS(value):
    return f"${float(value):,.2f}"


def TO_PERCENT(value):
    return f"{float(value) * 100:.2f}%"


def TO_PURE_NUMBER(value):
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace('$', '').replace(',', '').replace('%', '')
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        raise ValueError(f"Cannot convert '{value}' to number")


def TO_TEXT(value):
    return str(value)
