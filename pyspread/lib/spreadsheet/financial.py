import math

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


def ACCRINT(*args):
    raise NotImplementedError("ACCRINT() not implemented yet")


def ACCRINTM(*args):
    raise NotImplementedError("ACCRINTM() not implemented yet")


def AMORLINC(*args):
    raise NotImplementedError("AMORLINC() not implemented yet")


def COUPDAYBS(*args):
    raise NotImplementedError("COUPDAYBS() not implemented yet")


def COUPDAYS(*args):
    raise NotImplementedError("COUPDAYS() not implemented yet")


def COUPDAYSNC(*args):
    raise NotImplementedError("COUPDAYSNC() not implemented yet")


def COUPNCD(*args):
    raise NotImplementedError("COUPNCD() not implemented yet")


def COUPNUM(*args):
    raise NotImplementedError("COUPNUM() not implemented yet")


def COUPPCD(*args):
    raise NotImplementedError("COUPPCD() not implemented yet")


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


def DISC(*args):
    raise NotImplementedError("DISC() not implemented yet")


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


def DURATION(*args):
    raise NotImplementedError("DURATION() not implemented yet")


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


def INTRATE(*args):
    raise NotImplementedError("INTRATE() not implemented yet")


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


def MDURATION(*args):
    raise NotImplementedError("MDURATION() not implemented yet")


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


def PRICE(*args):
    raise NotImplementedError("PRICE() not implemented yet")


def PRICEDISC(*args):
    raise NotImplementedError("PRICEDISC() not implemented yet")


def PRICEMAT(*args):
    raise NotImplementedError("PRICEMAT() not implemented yet")


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


def RECEIVED(*args):
    raise NotImplementedError("RECEIVED() not implemented yet")


def RRI(nper, pv, fv):
    return (fv / pv) ** (1 / nper) - 1


def SLN(cost, salvage, life):
    return (cost - salvage) / life


def SYD(cost, salvage, life, per):
    return (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1))


def TBILLEQ(*args):
    raise NotImplementedError("TBILLEQ() not implemented yet")


def TBILLPRICE(*args):
    raise NotImplementedError("TBILLPRICE() not implemented yet")


def TBILLYIELD(*args):
    raise NotImplementedError("TBILLYIELD() not implemented yet")


def VDB(*args):
    raise NotImplementedError("VDB() not implemented yet")


def XIRR(values, dates, guess=0.1):
    raise NotImplementedError("XIRR() not implemented yet")


def XNPV(rate, values, dates):
    raise NotImplementedError("XNPV() not implemented yet")


def YIELD(*args):
    raise NotImplementedError("YIELD() not implemented yet")


def YIELDDISC(*args):
    raise NotImplementedError("YIELDDISC() not implemented yet")


def YIELDMAT(*args):
    raise NotImplementedError("YIELDMAT() not implemented yet")
