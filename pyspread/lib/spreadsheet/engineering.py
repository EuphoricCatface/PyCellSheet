import cmath
import math

_ENGINEERING_FUNCTIONS = [
    'BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR',
    'COMPLEX', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'DELTA', 'ERF', 'GESTEP', 'HEX2BIN',
    'HEX2DEC', 'HEX2OCT', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH',
    'IMCOT', 'IMCOTH', 'IMCSC', 'IMCSCH', 'IMDIV', 'IMEXP', 'IMLOG', 'IMLOG10', 'IMLOG2',
    'IMPRODUCT', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSUB', 'IMSUM', 'IMTAN',
    'IMTANH', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX'
]
__all__ = _ENGINEERING_FUNCTIONS + ["_ENGINEERING_FUNCTIONS"]


def _parse_complex(text):
    s = str(text).replace(' ', '')
    if s.endswith('i'):
        s = s[:-1] + 'j'
    try:
        return complex(s)
    except ValueError:
        return complex(text)


def _format_complex(c):
    r = c.real
    i = c.imag
    if i == 0:
        return str(r) if r != int(r) else str(int(r))
    if r == 0:
        return f"{i}i" if i != int(i) else f"{int(i)}i"
    sign = '+' if i > 0 else ''
    i_str = str(i) if i != int(i) else str(int(i))
    r_str = str(r) if r != int(r) else str(int(r))
    return f"{r_str}{sign}{i_str}i"


def BIN2DEC(number):
    return int(str(number), 2)


def BIN2HEX(number, places=None):
    dec = int(str(number), 2)
    result = hex(dec)[2:].upper()
    if places is not None:
        result = result.zfill(int(places))
    return result


def BIN2OCT(number, places=None):
    dec = int(str(number), 2)
    result = oct(dec)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def BITAND(number1, number2):
    return int(number1) & int(number2)


def BITLSHIFT(number, shift_amount):
    return int(number) << int(shift_amount)


def BITOR(number1, number2):
    return int(number1) | int(number2)


def BITRSHIFT(number, shift_amount):
    return int(number) >> int(shift_amount)


def BITXOR(number1, number2):
    return int(number1) ^ int(number2)


def COMPLEX(real_num, i_num, suffix='i'):
    if suffix not in ('i', 'j'):
        raise ValueError("Suffix must be 'i' or 'j'")
    return _format_complex(complex(real_num, i_num))


def DEC2BIN(number, places=None):
    n = int(number)
    if n < 0:
        result = bin(n & 0x3FF)[2:]
    else:
        result = bin(n)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def DEC2HEX(number, places=None):
    n = int(number)
    if n < 0:
        result = hex(n & 0xFFFFFFFFFF)[2:].upper()
    else:
        result = hex(n)[2:].upper()
    if places is not None:
        result = result.zfill(int(places))
    return result


def DEC2OCT(number, places=None):
    n = int(number)
    if n < 0:
        result = oct(n & 0x3FFFFFFF)[2:]
    else:
        result = oct(n)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def DELTA(number1, number2=0):
    return 1 if number1 == number2 else 0


class ERF:
    def __new__(cls, lower_limit, upper_limit=None):
        if upper_limit is None:
            return math.erf(lower_limit)
        return math.erf(upper_limit) - math.erf(lower_limit)

    @staticmethod
    def PRECISE(x):
        return math.erf(x)


def GESTEP(number, step=0):
    return 1 if number >= step else 0


def HEX2BIN(number, places=None):
    dec = int(str(number), 16)
    result = bin(dec)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def HEX2DEC(number):
    return int(str(number), 16)


def HEX2OCT(number, places=None):
    dec = int(str(number), 16)
    result = oct(dec)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def IMABS(inumber):
    return abs(_parse_complex(inumber))


def IMAGINARY(inumber):
    return _parse_complex(inumber).imag


def IMARGUMENT(inumber):
    return cmath.phase(_parse_complex(inumber))


def IMCONJUGATE(inumber):
    return _format_complex(_parse_complex(inumber).conjugate())


def IMCOS(inumber):
    return _format_complex(cmath.cos(_parse_complex(inumber)))


def IMCOSH(inumber):
    return _format_complex(cmath.cosh(_parse_complex(inumber)))


def IMCOT(inumber):
    c = _parse_complex(inumber)
    return _format_complex(cmath.cos(c) / cmath.sin(c))


def IMCOTH(inumber):
    c = _parse_complex(inumber)
    return _format_complex(cmath.cosh(c) / cmath.sinh(c))


def IMCSC(inumber):
    return _format_complex(1 / cmath.sin(_parse_complex(inumber)))


def IMCSCH(inumber):
    return _format_complex(1 / cmath.sinh(_parse_complex(inumber)))


def IMDIV(inumber1, inumber2):
    return _format_complex(_parse_complex(inumber1) / _parse_complex(inumber2))


def IMEXP(inumber):
    return _format_complex(cmath.exp(_parse_complex(inumber)))


def IMLOG(inumber, base=None):
    c = _parse_complex(inumber)
    if base is None:
        return _format_complex(cmath.log(c))
    return _format_complex(cmath.log(c) / cmath.log(base))


def IMLOG10(inumber):
    return _format_complex(cmath.log10(_parse_complex(inumber)))


def IMLOG2(inumber):
    return _format_complex(cmath.log(_parse_complex(inumber)) / cmath.log(2))


def IMPRODUCT(*args):
    result = complex(1, 0)
    for a in args:
        result *= _parse_complex(a)
    return _format_complex(result)


def IMREAL(inumber):
    return _parse_complex(inumber).real


def IMSEC(inumber):
    return _format_complex(1 / cmath.cos(_parse_complex(inumber)))


def IMSECH(inumber):
    return _format_complex(1 / cmath.cosh(_parse_complex(inumber)))


def IMSIN(inumber):
    return _format_complex(cmath.sin(_parse_complex(inumber)))


def IMSINH(inumber):
    return _format_complex(cmath.sinh(_parse_complex(inumber)))


def IMSUB(inumber1, inumber2):
    return _format_complex(_parse_complex(inumber1) - _parse_complex(inumber2))


def IMSUM(*args):
    result = complex(0, 0)
    for a in args:
        result += _parse_complex(a)
    return _format_complex(result)


def IMTAN(inumber):
    return _format_complex(cmath.tan(_parse_complex(inumber)))


def IMTANH(inumber):
    return _format_complex(cmath.tanh(_parse_complex(inumber)))


def OCT2BIN(number, places=None):
    dec = int(str(number), 8)
    result = bin(dec)[2:]
    if places is not None:
        result = result.zfill(int(places))
    return result


def OCT2DEC(number):
    return int(str(number), 8)


def OCT2HEX(number, places=None):
    dec = int(str(number), 8)
    result = hex(dec)[2:].upper()
    if places is not None:
        result = result.zfill(int(places))
    return result
