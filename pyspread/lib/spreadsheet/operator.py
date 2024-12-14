def ADD(a, b):
    return a + b


def CONCAT(a, b):
    return str(a) + str(b)


def DIVIDE(a, b):
    return a / b


def EQ(a, b):
    return a == b


def GT(a, b):
    return a > b


def GTE(a, b):
    return a >= b


def ISBETWEEN(value, lower, upper, lower_inclusive, upper_inclusive):
    if value > lower or (lower_inclusive and value == lower):
        pass  # okay
    else:
        return False
    if value < upper or (upper_inclusive and value == upper):
        pass  # okay
    else:
        return False
    return True


def LT(a, b):
    return a < b


def LTE(a, b):
    return a <= b


def MINUS(a):
    return -a


def MULTIPLY(a, b):
    return a * b


def NE(a, b):
    return a != b


def POW(a, b):
    return a ** b


def UMINUS(a):
    return -a


def UNARY_PERCENT(a):
    return a / 100


# def UNIQUE():
# Google Sheets help includes UNIQUE as one of operators
#  but I think it's a typo or something.
# It is also listed in the "Filter" section.


def UPLUS(a):
    return +a
