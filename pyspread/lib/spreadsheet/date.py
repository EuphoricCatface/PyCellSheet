from datetime import datetime, date, time, timedelta
from typing import Union, Optional, SupportsIndex

try:
    from pyspread.lib.pycellsheet import Range
except ImportError:
    from lib.pycellsheet import Range

__all__ = [
    'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360', 'EDATE', 'EOMONTH', 'EPOCHTODATE',
    'HOUR', 'ISOWEEKNUM', 'MINUTE', 'MONTH', 'NETWORKDAYS', 'NOW', 'SECOND', 'TIME', 'TIMEVALUE',
    'TODAY', 'WEEKDAY', 'WEEKNUM','WORKDAY', 'YEAR', 'YEARFRAC'
]


def DATE(year: SupportsIndex, month: SupportsIndex, day: SupportsIndex)\
        -> date:
    return date(year, month, day)


def DATEDIF(start: date, end: date, unit: Optional[str] = None)\
        -> timedelta:
    if unit is not None:
        raise NotImplementedError("`unit` operation is not supported yet on DATEDIF")
    return end - start


def DATEVALUE(date_string: str)\
        -> date:
    try:
        import dateutil.parser
    except ImportError:
        raise NotImplementedError("Install `dateutil` python package to use DATEVALUE")
    return dateutil.parser.parse(date_string).date()


def DAY(date_: date)\
        -> int:
    return date_.day


def DAYS(end_date: date, start_date: date):
    return (start_date - end_date).days


def DAYS360():
    raise NotImplementedError("DAYS360 is not implemented yet")


def EDATE(start_date: date, months: int)\
        -> date:
    raise NotImplementedError("EDATE is not implemented yet")


def EOMONTH(start_date: date, months: int)\
        -> int:
    raise NotImplementedError("EOMONTH is not implemented yet")


def HOUR(time_: time)\
        -> int:
    return time_.hour


def ISOWEEKNUM(date_: date)\
        -> int:
    return date_.isocalendar()[1]


def MINUTE(time_: time)\
        -> int:
    return time_.minute


def MONTH(date_: date)\
        -> int:
    return date_.month


class NETWORKDAYS:
    @staticmethod
    def __call__(start_date: date, end_date: date, holidays: Optional[Range] = None)\
            -> int:
        raise NotImplementedError("NETWORKDAYS is not implemented yet")

    @staticmethod
    def INTL(start_date: date, end_date: date, weekend: Union[int, str] = 1,
             holidays: Optional[list[date]] = None)\
            -> int:
        raise NotImplementedError("NETWORKDAYS.INTL is not implemented yet")


def NOW():
    return datetime.now()


def SECOND(time_: time):
    return time_.second


def TIME(hour: SupportsIndex, minute: SupportsIndex, second: SupportsIndex):
    return time(hour, minute, second)


def TODAY():
    return date.today()


def TIMEVALUE(time_string: str)\
        -> time:
    try:
        import dateutil.parser
    except ImportError:
        raise NotImplementedError("Install `dateutil` python package to use TIMEVALUE")
    return dateutil.parser.parse(time_string).time()


def WEEKDAY(date_: date, type_: int = 1)\
        -> int:
    match type_:
        case 1:
            return (date_.weekday() + 1) % 7 + 1
        case 2:
            return date_.isoweekday()
        case 3:
            return date_.weekday()
        case _:
            raise ValueError("type_ value should be 1, 2 or 3")


def WEEKNUM(date_: date, type_: int = 1):
    raise NotImplementedError("WEEKNUM is not implemented yet")


class WORKDAY:
    @staticmethod
    def __call__(start_date: date, num_days: int, holidays: Optional[Range] = None):
        raise NotImplementedError("WORKDAY is not implemented yet")

    @staticmethod
    def INTL(start_date: date, num_days: int, weekend: Union[int, str] = 1, holidays: Optional[Range] = None):
        raise NotImplementedError("WORKDAY.INTL is not implemented yet")


def YEAR(date_: date):
    return date_.year


def YEARFRAC():
    raise NotImplementedError("YEARFRAC is not implemented yet")


def EPOCHTODATE(timestamp: float, unit: int)\
        -> datetime:
    if unit not in [1, 2, 3]:
        raise ValueError("unit value should be 1, 2 or 3")

    return datetime.fromtimestamp(timestamp / (1000 ** (unit - 1)))

