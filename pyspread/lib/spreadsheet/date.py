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

import calendar
from datetime import datetime, date, time, timedelta
from typing import Union, Optional, SupportsIndex

try:
    from pyspread.lib.pycellsheet import Range
except ImportError:
    from lib.pycellsheet import Range

_DATE_FUNCTIONS = [
    'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360', 'EDATE', 'EOMONTH', 'EPOCHTODATE',
    'HOUR', 'ISOWEEKNUM', 'MINUTE', 'MONTH', 'NETWORKDAYS', 'NOW', 'SECOND', 'TIME', 'TIMEVALUE',
    'TODAY', 'WEEKDAY', 'WEEKNUM','WORKDAY', 'YEAR', 'YEARFRAC'
]
__all__ = _DATE_FUNCTIONS + ["_DATE_FUNCTIONS"]


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


def DAYS360(start_date: date, end_date: date, method=False):
    s_y, s_m, s_d = start_date.year, start_date.month, min(start_date.day, 30)
    e_y, e_m, e_d = end_date.year, end_date.month, min(end_date.day, 30)
    if not method:  # US method
        if s_d == 31:
            s_d = 30
        if e_d == 31 and s_d >= 30:
            e_d = 30
    return (e_y - s_y) * 360 + (e_m - s_m) * 30 + (e_d - s_d)


def EDATE(start_date: date, months: int)\
        -> date:
    month = start_date.month + months
    year = start_date.year
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    day = min(start_date.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def EOMONTH(start_date: date, months: int)\
        -> date:
    target = EDATE(start_date, months)
    last_day = calendar.monthrange(target.year, target.month)[1]
    return date(target.year, target.month, last_day)


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
    def __new__(cls, start_date: date, end_date: date, holidays: Optional[Range] = None)\
            -> int:
        holiday_set = set()
        if holidays is not None:
            holiday_set = set(holidays.flatten())
        count = 0
        step = 1 if end_date >= start_date else -1
        current = start_date
        while (step > 0 and current <= end_date) or (step < 0 and current >= end_date):
            if current.weekday() < 5 and current not in holiday_set:
                count += 1
            current += timedelta(days=step)
        return count if step > 0 else -count

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
    if type_ == 1:
        jan1 = date(date_.year, 1, 1)
        jan1_weekday = (jan1.weekday() + 1) % 7
        day_of_year = (date_ - jan1).days
        return (day_of_year + jan1_weekday) // 7 + 1
    elif type_ == 2:
        return date_.isocalendar()[1]
    else:
        raise ValueError("type_ value should be 1 or 2")


class WORKDAY:
    def __new__(cls, start_date: date, num_days: int, holidays: Optional[Range] = None):
        holiday_set = set()
        if holidays is not None:
            holiday_set = set(holidays.flatten())
        current = start_date
        step = 1 if num_days > 0 else -1
        remaining = abs(num_days)
        while remaining > 0:
            current += timedelta(days=step)
            if current.weekday() < 5 and current not in holiday_set:
                remaining -= 1
        return current

    @staticmethod
    def INTL(start_date: date, num_days: int, weekend: Union[int, str] = 1, holidays: Optional[Range] = None):
        raise NotImplementedError("WORKDAY.INTL is not implemented yet")


def YEAR(date_: date):
    return date_.year


def YEARFRAC(start_date: date, end_date: date, basis: int = 0):
    days = abs((end_date - start_date).days)
    if basis == 0:
        return days / 360
    elif basis == 1:
        year = start_date.year
        year_days = 366 if calendar.isleap(year) else 365
        return days / year_days
    elif basis == 2:
        return days / 360
    elif basis == 3:
        return days / 365
    elif basis == 4:
        return days / 360
    else:
        raise ValueError("basis must be 0, 1, 2, 3, or 4")


def EPOCHTODATE(timestamp: float, unit: int)\
        -> datetime:
    if unit not in [1, 2, 3]:
        raise ValueError("unit value should be 1, 2 or 3")

    return datetime.fromtimestamp(timestamp / (1000 ** (unit - 1)))
