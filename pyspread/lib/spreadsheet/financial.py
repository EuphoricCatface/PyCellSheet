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

import math
from datetime import datetime, timedelta

try:
    from pyspread.lib.pycellsheet import flatten_args
except ImportError:
    from lib.pycellsheet import flatten_args

_FINANCIAL_FUNCTIONS = [
    'ACCRINT', 'ACCRINTM', 'AMORLINC', 'COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM',
    'COUPPCD', 'CUMIPMT', 'CUMPRINC', 'DB', 'DDB', 'DISC', 'DOLLARDE', 'DOLLARFR', 'DURATION',
    'EFFECT', 'FV', 'FVSCHEDULE', 'INTRATE', 'IPMT', 'IRR', 'ISPMT', 'MDURATION', 'MIRR', 'NOMINAL',
    'NPER', 'NPV', 'PDURATION', 'PMT', 'PPMT', 'PRICE', 'PRICEDISC', 'PRICEMAT', 'PV', 'RATE',
    'RECEIVED', 'RRI', 'SLN', 'SYD', 'TBILLEQ', 'TBILLPRICE', 'TBILLYIELD', 'VDB', 'XIRR', 'XNPV',
    'YIELD', 'YIELDDISC', 'YIELDMAT'
]
__all__ = _FINANCIAL_FUNCTIONS + ["_FINANCIAL_FUNCTIONS"]


def ACCRINT(issue, first_interest, settlement, rate, par=1000, frequency=2, basis=0):
    """Calculate accrued interest for a security that pays periodic interest."""
    # Simplified implementation
    if isinstance(issue, datetime):
        days = (settlement - issue).days
    else:
        days = settlement - issue
    return par * rate * days / 365.0


def ACCRINTM(issue, settlement, rate, par=1000, basis=0):
    """Calculate accrued interest for a security that pays interest at maturity."""
    if isinstance(issue, datetime):
        days = (settlement - issue).days
    else:
        days = settlement - issue
    return par * rate * days / 365.0


def AMORLINC(cost, date_purchased, first_period, salvage, period, rate, basis=0):
    """Calculate depreciation for each accounting period (linear)."""
    depreciation = cost * rate
    if period == 0:
        # First period
        return depreciation
    elif (cost - depreciation * period) < salvage:
        # Last period
        return max(0, cost - salvage - depreciation * (period - 1))
    return depreciation


def COUPDAYBS(settlement, maturity, frequency, basis=0):
    """Number of days from beginning of coupon period to settlement."""
    # Simplified: assume monthly periods
    days_in_period = 365 / frequency
    return days_in_period / 2


def COUPDAYS(settlement, maturity, frequency, basis=0):
    """Number of days in the coupon period containing settlement."""
    return 365 / frequency


def COUPDAYSNC(settlement, maturity, frequency, basis=0):
    """Number of days from settlement to next coupon date."""
    days_in_period = 365 / frequency
    return days_in_period / 2


def COUPNCD(settlement, maturity, frequency, basis=0):
    """Next coupon date after settlement."""
    # Simplified: add one period
    if isinstance(settlement, datetime):
        days = int(365 / frequency)
        return settlement + timedelta(days=days)
    return settlement + (365 / frequency)


def COUPNUM(settlement, maturity, frequency, basis=0):
    """Number of coupons between settlement and maturity."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        years = (maturity - settlement).days / 365.0
    else:
        years = (maturity - settlement) / 365.0
    return int(years * frequency)


def COUPPCD(settlement, maturity, frequency, basis=0):
    """Previous coupon date before settlement."""
    # Simplified: subtract one period
    if isinstance(settlement, datetime):
        days = int(365 / frequency)
        return settlement - timedelta(days=days)
    return settlement - (365 / frequency)


def CUMIPMT(rate, nper, pv, start_period, end_period, type_):
    total = 0
    for period in range(int(start_period), int(end_period) + 1):
        total += IPMT(rate, period, nper, pv, 0, type_)
    return total


def CUMPRINC(rate, nper, pv, start_period, end_period, type_):
    total = 0
    for period in range(int(start_period), int(end_period) + 1):
        total += PPMT(rate, period, nper, pv, 0, type_)
    return total


def DB(cost, salvage, life, period, month=12):
    rate = 1 - (salvage / cost) ** (1 / life)
    rate = round(rate, 3)
    depreciation = cost * rate * month / 12
    if period == 1:
        return depreciation
    total_dep = depreciation
    for p in range(2, int(period) + 1):
        depreciation = (cost - total_dep) * rate
        total_dep += depreciation
    return depreciation


def DDB(cost, salvage, life, period, factor=2):
    rate = factor / life
    value = cost
    for p in range(1, int(period) + 1):
        depreciation = value * rate
        if value - depreciation < salvage:
            depreciation = value - salvage
        value -= depreciation
    return depreciation


def DISC(settlement, maturity, pr, redemption, basis=0):
    """Discount rate for a security."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    b = 365 if basis == 0 else 360
    return (redemption - pr) / redemption * (b / dsm)


def DOLLARDE(fractional_dollar, fraction):
    fraction = int(fraction)
    integer_part = int(fractional_dollar)
    frac_part = fractional_dollar - integer_part
    return integer_part + frac_part * 10 ** math.ceil(math.log10(fraction)) / fraction


def DOLLARFR(decimal_dollar, fraction):
    fraction = int(fraction)
    integer_part = int(decimal_dollar)
    frac_part = decimal_dollar - integer_part
    return integer_part + frac_part * fraction / 10 ** math.ceil(math.log10(fraction))


def DURATION(settlement, maturity, coupon, yld, frequency, basis=0):
    """Macaulay duration for a security with periodic interest payments."""
    # Simplified Macaulay duration calculation
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        years = (maturity - settlement).days / 365.0
    else:
        years = (maturity - settlement) / 365.0

    n = int(years * frequency)
    if n == 0:
        return 0

    coupon_pmt = coupon / frequency
    discount_rate = yld / frequency

    pv_weighted = 0
    pv_total = 0

    for t in range(1, n + 1):
        cf = coupon_pmt
        if t == n:
            cf += 1  # Add principal at maturity
        discount = (1 + discount_rate) ** t
        pv_weighted += t * cf / discount
        pv_total += cf / discount

    return (pv_weighted / pv_total) / frequency


def EFFECT(nominal_rate, npery):
    return (1 + nominal_rate / npery) ** npery - 1


def FV(rate, nper, pmt, pv=0, type_=0):
    if rate == 0:
        return -(pv + pmt * nper)
    return -(pv * (1 + rate) ** nper + pmt * ((1 + rate) ** nper - 1) / rate * (1 + rate * type_))


def FVSCHEDULE(principal, schedule):
    lst = flatten_args(schedule) if hasattr(schedule, 'flatten') else schedule
    result = principal
    for rate in lst:
        result *= (1 + rate)
    return result


def INTRATE(settlement, maturity, investment, redemption, basis=0):
    """Interest rate for a fully invested security."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    b = 365 if basis == 0 else 360
    return (redemption - investment) / investment * (b / dsm)


def IPMT(rate, per, nper, pv, fv=0, type_=0):
    payment = PMT(rate, nper, pv, fv, type_)
    if per == 1 and type_ == 1:
        return 0
    if type_ == 1:
        fv_prior = FV(rate, per - 2, payment, pv, 1)
        return fv_prior * rate
    else:
        fv_prior = FV(rate, per - 1, payment, pv, 0)
        return fv_prior * rate


def IRR(values, guess=0.1):
    lst = flatten_args(values) if hasattr(values, 'flatten') else list(values)
    rate = guess
    for _ in range(100):
        npv = sum(v / (1 + rate) ** i for i, v in enumerate(lst))
        dnpv = sum(-i * v / (1 + rate) ** (i + 1) for i, v in enumerate(lst))
        if abs(dnpv) < 1e-10:
            break
        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate
    return rate


def ISPMT(rate, per, nper, pv):
    return -pv * rate * (1 - per / nper)


def MDURATION(settlement, maturity, coupon, yld, frequency, basis=0):
    """Modified duration for a security with periodic interest payments."""
    macaulay_dur = DURATION(settlement, maturity, coupon, yld, frequency, basis)
    return macaulay_dur / (1 + yld / frequency)


def MIRR(values, finance_rate, reinvest_rate):
    lst = flatten_args(values) if hasattr(values, 'flatten') else list(values)
    n = len(lst)
    neg_pv = sum(v / (1 + finance_rate) ** i for i, v in enumerate(lst) if v < 0)
    pos_fv = sum(v * (1 + reinvest_rate) ** (n - 1 - i) for i, v in enumerate(lst) if v > 0)
    return (-pos_fv / neg_pv) ** (1 / (n - 1)) - 1


def NOMINAL(effect_rate, npery):
    return npery * ((1 + effect_rate) ** (1 / npery) - 1)


def NPER(rate, pmt, pv, fv=0, type_=0):
    if rate == 0:
        return -(pv + fv) / pmt
    z = pmt * (1 + rate * type_) / rate
    return math.log((-fv + z) / (pv + z)) / math.log(1 + rate)


def NPV(rate, *values):
    lst = flatten_args(*values)
    return sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(lst))


def PDURATION(rate, pv, fv):
    return (math.log(fv) - math.log(pv)) / math.log(1 + rate)


def PMT(rate, nper, pv, fv=0, type_=0):
    if rate == 0:
        return -(pv + fv) / nper
    factor = (1 + rate) ** nper
    return -(pv * factor + fv) / ((factor - 1) / rate * (1 + rate * type_))


def PPMT(rate, per, nper, pv, fv=0, type_=0):
    return PMT(rate, nper, pv, fv, type_) - IPMT(rate, per, nper, pv, fv, type_)


def PRICE(settlement, maturity, rate, yld, redemption=100, frequency=2, basis=0):
    """Price per $100 face value of a security that pays periodic interest."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        years = (maturity - settlement).days / 365.0
    else:
        years = (maturity - settlement) / 365.0

    n = int(years * frequency)
    if n == 0:
        return redemption

    coupon_pmt = rate * redemption / frequency
    discount_rate = yld / frequency

    pv = 0
    for t in range(1, n + 1):
        pv += coupon_pmt / ((1 + discount_rate) ** t)

    pv += redemption / ((1 + discount_rate) ** n)
    return pv


def PRICEDISC(settlement, maturity, discount, redemption=100, basis=0):
    """Price per $100 face value of a discounted security."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    b = 365 if basis == 0 else 360
    return redemption - discount * redemption * dsm / b


def PRICEMAT(settlement, maturity, issue, rate, yld, basis=0):
    """Price per $100 face value of a security that pays interest at maturity."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
        dim = (maturity - issue).days
        dsi = (settlement - issue).days
    else:
        dsm = maturity - settlement
        dim = maturity - issue
        dsi = settlement - issue

    b = 365 if basis == 0 else 360
    return (100 + rate * 100 * dim / b) / (1 + yld * dsm / b) - rate * 100 * dsi / b


def PV(rate, nper, pmt, fv=0, type_=0):
    if rate == 0:
        return -(fv + pmt * nper)
    factor = (1 + rate) ** nper
    return -(fv + pmt * (factor - 1) / rate * (1 + rate * type_)) / factor


def RATE(nper, pmt, pv, fv=0, type_=0, guess=0.1):
    rate = guess
    for _ in range(100):
        factor = (1 + rate) ** nper
        y = pv * factor + pmt * (factor - 1) / rate * (1 + rate * type_) + fv
        dy = nper * pv * (1 + rate) ** (nper - 1) + pmt * (1 + rate * type_) * (
            nper * (1 + rate) ** (nper - 1) * rate - (factor - 1)
        ) / rate ** 2
        new_rate = rate - y / dy
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate
    return rate


def RECEIVED(settlement, maturity, investment, discount, basis=0):
    """Amount received at maturity for a fully invested security."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    b = 365 if basis == 0 else 360
    return investment / (1 - discount * dsm / b)


def RRI(nper, pv, fv):
    return (fv / pv) ** (1 / nper) - 1


def SLN(cost, salvage, life):
    return (cost - salvage) / life


def SYD(cost, salvage, life, per):
    return (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1))


def TBILLEQ(settlement, maturity, discount):
    """Bond-equivalent yield for a Treasury bill."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    return (365 * discount) / (360 - discount * dsm)


def TBILLPRICE(settlement, maturity, discount):
    """Price per $100 face value for a Treasury bill."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    return 100 * (1 - discount * dsm / 360)


def TBILLYIELD(settlement, maturity, pr):
    """Yield for a Treasury bill."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    return (100 - pr) / pr * (360 / dsm)


def VDB(cost, salvage, life, start_period, end_period, factor=2, no_switch=False):
    """Variable declining balance depreciation."""
    # Simplified VDB using DDB
    rate = factor / life
    value = cost
    total_dep = 0

    for p in range(1, int(end_period) + 1):
        if p > int(start_period):
            depreciation = value * rate
            if not no_switch:
                # Check if switching to straight-line is better
                remaining_life = life - p + 1
                sl_dep = (value - salvage) / remaining_life if remaining_life > 0 else 0
                if sl_dep > depreciation:
                    depreciation = sl_dep

            if value - depreciation < salvage:
                depreciation = value - salvage

            total_dep += depreciation
            value -= depreciation
        else:
            # Still computing depreciation but not accumulating for return
            depreciation = value * rate
            if value - depreciation < salvage:
                depreciation = value - salvage
            value -= depreciation

    return total_dep


def XIRR(values, dates, guess=0.1):
    """Internal rate of return for irregular cash flows."""
    vals = flatten_args(values) if hasattr(values, '__iter__') else [values]
    dts = flatten_args(dates) if hasattr(dates, '__iter__') else [dates]

    # Convert dates to days from first date
    if isinstance(dts[0], datetime):
        days = [(d - dts[0]).days for d in dts]
    else:
        days = [d - dts[0] for d in dts]

    rate = guess
    for _ in range(100):
        npv = sum(v / (1 + rate) ** (d / 365.0) for v, d in zip(vals, days))
        dnpv = sum(-d / 365.0 * v / (1 + rate) ** (d / 365.0 + 1) for v, d in zip(vals, days))

        if abs(dnpv) < 1e-10:
            break

        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate

    return rate


def XNPV(rate, values, dates):
    """Net present value for irregular cash flows."""
    vals = flatten_args(values) if hasattr(values, '__iter__') else [values]
    dts = flatten_args(dates) if hasattr(dates, '__iter__') else [dates]

    # Convert dates to days from first date
    if isinstance(dts[0], datetime):
        days = [(d - dts[0]).days for d in dts]
    else:
        days = [d - dts[0] for d in dts]

    return sum(v / (1 + rate) ** (d / 365.0) for v, d in zip(vals, days))


def YIELD(settlement, maturity, rate, pr, redemption=100, frequency=2, basis=0):
    """Yield for a security that pays periodic interest."""
    # Use Newton-Raphson to solve for yield
    yld = 0.1  # Initial guess

    for _ in range(100):
        price = PRICE(settlement, maturity, rate, yld, redemption, frequency, basis)
        error = price - pr

        if abs(error) < 0.0001:
            return yld

        # Numerical derivative
        delta = 0.00001
        price_plus = PRICE(settlement, maturity, rate, yld + delta, redemption, frequency, basis)
        derivative = (price_plus - price) / delta

        if abs(derivative) < 1e-10:
            break

        yld = yld - error / derivative

        if yld < -1:  # Prevent negative yields from going too far
            yld = -0.5

    return yld


def YIELDDISC(settlement, maturity, pr, redemption=100, basis=0):
    """Annual yield for a discounted security."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
    else:
        dsm = maturity - settlement
    b = 365 if basis == 0 else 360
    return (redemption - pr) / pr * (b / dsm)


def YIELDMAT(settlement, maturity, issue, rate, pr, basis=0):
    """Annual yield of a security that pays interest at maturity."""
    if isinstance(settlement, datetime) and isinstance(maturity, datetime):
        dsm = (maturity - settlement).days
        dim = (maturity - issue).days
        dsi = (settlement - issue).days
    else:
        dsm = maturity - settlement
        dim = maturity - issue
        dsi = settlement - issue

    b = 365 if basis == 0 else 360
    return ((100 + rate * 100 * dim / b) / (pr + rate * 100 * dsi / b) - 1) * (b / dsm)
