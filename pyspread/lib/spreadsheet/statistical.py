import math
import statistics

try:
    from pyspread.lib.pycellsheet import EmptyCell, Range, flatten_args
except ImportError:
    from lib.pycellsheet import EmptyCell, Range, flatten_args

_STATISTICAL_FUNCTIONS = [
    'AVEDEV', 'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'BETA', 'BETADIST', 'BETAINV',
    'BINOM', 'BINOMDIST', 'CHIDIST', 'CHIINV', 'CHISQ', 'CHITEST', 'CONFIDENCE', 'CORREL', 'COUNT',
    'COUNTA', 'COVAR', 'COVARIANCE', 'CRITBINOM', 'DEVSQ', 'EXPON', 'EXPONDIST', 'F', 'FDIST',
    'FINV', 'FISHER', 'FISHERINV', 'FORECAST', 'FTEST', 'GAMMA', 'GAMMADIST', 'GAMMAINV', 'GAUSS',
    'GEOMEAN', 'HARMEAN', 'HYPGEOM', 'HYPGEOMDIST', 'INTERCEPT', 'KURT', 'LARGE', 'LOGINV',
    'LOGNORM', 'LOGNORMDIST', 'MARGINOFERROR', 'MAX', 'MAXA', 'MAXIFS', 'MEDIAN', 'MIN', 'MINA',
    'MINIFS', 'MODE', 'NEGBINOM', 'NEGBINOMDIST', 'NORM', 'NORMDIST', 'NORMINV', 'NORMSDIST',
    'NORMSINV', 'PEARSON', 'PERCENTILE', 'PERCENTRANK', 'PERMUT', 'PERMUTATIONA', 'PHI', 'POISSON',
    'PROB', 'QUARTILE', 'RANK', 'RSQ', 'SKEW', 'SLOPE', 'SMALL', 'STANDARDIZE', 'STDEV', 'STDEVA',
    'STDEVP', 'STDEVPA', 'STEYX', 'T_STAT', 'TDIST', 'TINV', 'TRIMMEAN', 'TTEST', 'VAR', 'VARA', 'VARP',
    'VARPA', 'WEIBULL', 'Z', 'ZTEST'
]
__all__ = _STATISTICAL_FUNCTIONS + ["_STATISTICAL_FUNCTIONS"]


def _numeric_filter(lst):
    return [v for v in lst if isinstance(v, (int, float)) and v != EmptyCell]


def AVEDEV(*args):
    lst = _numeric_filter(flatten_args(*args))
    mean = sum(lst) / len(lst)
    return sum(abs(x - mean) for x in lst) / len(lst)


class AVERAGE:
    def __new__(cls, *args):
        lst = flatten_args(*args)
        return sum(lst) / len(lst)

    @staticmethod
    def WEIGHTED(values, weights):
        vals = flatten_args(values)
        wts = flatten_args(weights)
        if len(vals) != len(wts):
            raise ValueError("Values and weights must have the same length")
        return sum(v * w for v, w in zip(vals, wts)) / sum(wts)


def AVERAGEA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
        else:
            converted.append(0)
    if not converted:
        raise ValueError("No values")
    return sum(converted) / len(converted)


def AVERAGEIF(criteria_range: Range, criterion, average_range: Range = None):
    if average_range is None:
        average_range = criteria_range
    values = []
    for i, v in enumerate(criteria_range.lst):
        if criterion(v):
            av = average_range.lst[i]
            if isinstance(av, (int, float)) and av != EmptyCell:
                values.append(av)
    if not values:
        raise ValueError("No matching values")
    return sum(values) / len(values)


def AVERAGEIFS(average_range: Range, *criteria_pairs):
    if len(criteria_pairs) % 2:
        raise ValueError("Number of criteria arguments has to be even")
    values = []
    for idx in range(len(average_range.lst)):
        include = True
        for i in range(0, len(criteria_pairs), 2):
            cr = criteria_pairs[i]
            criterion = criteria_pairs[i + 1]
            if not criterion(cr.lst[idx]):
                include = False
                break
        if include:
            av = average_range.lst[idx]
            if isinstance(av, (int, float)) and av != EmptyCell:
                values.append(av)
    if not values:
        raise ValueError("No matching values")
    return sum(values) / len(values)


class BETA:
    @staticmethod
    def DIST(x, alpha, beta, cumulative=True):
        raise NotImplementedError("BETA.DIST() not implemented yet")

    @staticmethod
    def INV(probability, alpha, beta):
        raise NotImplementedError("BETA.INV() not implemented yet")


def BETADIST(x, alpha, beta):
    raise NotImplementedError("BETADIST() not implemented yet")


def BETAINV(probability, alpha, beta):
    raise NotImplementedError("BETAINV() not implemented yet")


class BINOM:
    @staticmethod
    def DIST(number_s, trials, probability_s, cumulative):
        raise NotImplementedError("BINOM.DIST() not implemented yet")

    @staticmethod
    def INV(trials, probability_s, alpha):
        raise NotImplementedError("BINOM.INV() not implemented yet")


def BINOMDIST(number_s, trials, probability_s, cumulative):
    raise NotImplementedError("BINOMDIST() not implemented yet")


def CHIDIST(x, degrees_freedom):
    raise NotImplementedError("CHIDIST() not implemented yet")


def CHIINV(probability, degrees_freedom):
    raise NotImplementedError("CHIINV() not implemented yet")


class CHISQ:
    class DIST:
        def __new__(cls, x, degrees_freedom, cumulative):
            raise NotImplementedError("CHISQ.DIST() not implemented yet")

        @staticmethod
        def RT(x, degrees_freedom):
            raise NotImplementedError("CHISQ.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, probability, degrees_freedom):
            raise NotImplementedError("CHISQ.INV() not implemented yet")

        @staticmethod
        def RT(probability, degrees_freedom):
            raise NotImplementedError("CHISQ.INV.RT() not implemented yet")

    @staticmethod
    def TEST(actual_range, expected_range):
        raise NotImplementedError("CHISQ.TEST() not implemented yet")


def CHITEST(actual_range, expected_range):
    raise NotImplementedError("CHITEST() not implemented yet")


class CONFIDENCE:
    def __new__(cls, alpha, standard_dev, size):
        raise NotImplementedError("CONFIDENCE() not implemented yet")

    @staticmethod
    def NORM(alpha, standard_dev, size):
        raise NotImplementedError("CONFIDENCE.NORM() not implemented yet")

    @staticmethod
    def T(alpha, standard_dev, size):
        raise NotImplementedError("CONFIDENCE.T() not implemented yet")


def CORREL(data_y, data_x):
    y = _numeric_filter(flatten_args(data_y))
    x = _numeric_filter(flatten_args(data_x))
    return statistics.correlation(x, y)


def COUNT(*args):
    lst = flatten_args(*args)
    return len([v for v in lst if isinstance(v, (int, float)) and v != EmptyCell])


def COUNTA(*args):
    lst = flatten_args(*args)
    return len([v for v in lst if v != EmptyCell])


def COVAR(data_y, data_x):
    y = _numeric_filter(flatten_args(data_y))
    x = _numeric_filter(flatten_args(data_x))
    return statistics.covariance(x, y)


class COVARIANCE:
    @staticmethod
    def P(data_y, data_x):
        y = _numeric_filter(flatten_args(data_y))
        x = _numeric_filter(flatten_args(data_x))
        n = len(x)
        mean_x = sum(x) / n
        mean_y = sum(y) / n
        return sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x, y)) / n

    @staticmethod
    def S(data_y, data_x):
        y = _numeric_filter(flatten_args(data_y))
        x = _numeric_filter(flatten_args(data_x))
        return statistics.covariance(x, y)


def CRITBINOM(trials, probability_s, alpha):
    raise NotImplementedError("CRITBINOM() not implemented yet")


def DEVSQ(*args):
    lst = _numeric_filter(flatten_args(*args))
    mean = sum(lst) / len(lst)
    return sum((x - mean) ** 2 for x in lst)


class EXPON:
    @staticmethod
    def DIST(x, _lambda, cumulative):
        raise NotImplementedError("EXPON.DIST() not implemented yet")


def EXPONDIST(x, _lambda, cumulative):
    raise NotImplementedError("EXPONDIST() not implemented yet")


class F:
    class DIST:
        def __new__(cls, x, deg_freedom1, deg_freedom2, cumulative):
            raise NotImplementedError("F.DIST() not implemented yet")

        @staticmethod
        def RT(x, deg_freedom1, deg_freedom2):
            raise NotImplementedError("F.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, probability, deg_freedom1, deg_freedom2):
            raise NotImplementedError("F.INV() not implemented yet")

        @staticmethod
        def RT(probability, deg_freedom1, deg_freedom2):
            raise NotImplementedError("F.INV.RT() not implemented yet")

    @staticmethod
    def TEST(array1, array2):
        raise NotImplementedError("F.TEST() not implemented yet")


def FDIST(x, deg_freedom1, deg_freedom2):
    raise NotImplementedError("FDIST() not implemented yet")


def FINV(probability, deg_freedom1, deg_freedom2):
    raise NotImplementedError("FINV() not implemented yet")


def FISHER(x):
    return 0.5 * math.log((1 + x) / (1 - x))


def FISHERINV(y):
    e2y = math.exp(2 * y)
    return (e2y - 1) / (e2y + 1)


class FORECAST:
    def __new__(cls, x, known_ys, known_xs):
        y_vals = _numeric_filter(flatten_args(known_ys))
        x_vals = _numeric_filter(flatten_args(known_xs))
        n = len(x_vals)
        mean_x = sum(x_vals) / n
        mean_y = sum(y_vals) / n
        num = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x_vals, y_vals))
        den = sum((xi - mean_x) ** 2 for xi in x_vals)
        slope = num / den
        intercept = mean_y - slope * mean_x
        return intercept + slope * x

    @staticmethod
    def LINEAR(x, known_ys, known_xs):
        return FORECAST.__new__(FORECAST, x, known_ys, known_xs)


def FTEST(array1, array2):
    raise NotImplementedError("FTEST() not implemented yet")


class GAMMA:
    def __new__(cls, value):
        return math.gamma(value)

    @staticmethod
    def DIST(x, alpha, beta, cumulative):
        raise NotImplementedError("GAMMA.DIST() not implemented yet")

    @staticmethod
    def INV(probability, alpha, beta):
        raise NotImplementedError("GAMMA.INV() not implemented yet")


def GAMMADIST(x, alpha, beta, cumulative):
    raise NotImplementedError("GAMMADIST() not implemented yet")


def GAMMAINV(probability, alpha, beta):
    raise NotImplementedError("GAMMAINV() not implemented yet")


def GAUSS(z):
    return statistics.NormalDist().cdf(z) - 0.5


def GEOMEAN(*args):
    lst = _numeric_filter(flatten_args(*args))
    return statistics.geometric_mean(lst)


def HARMEAN(*args):
    lst = _numeric_filter(flatten_args(*args))
    return statistics.harmonic_mean(lst)


class HYPGEOM:
    @staticmethod
    def DIST(sample_s, number_sample, population_s, number_pop, cumulative):
        raise NotImplementedError("HYPGEOM.DIST() not implemented yet")


def HYPGEOMDIST(sample_s, number_sample, population_s, number_pop):
    raise NotImplementedError("HYPGEOMDIST() not implemented yet")


def INTERCEPT(known_ys, known_xs):
    y_vals = _numeric_filter(flatten_args(known_ys))
    x_vals = _numeric_filter(flatten_args(known_xs))
    n = len(x_vals)
    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n
    num = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x_vals, y_vals))
    den = sum((xi - mean_x) ** 2 for xi in x_vals)
    slope = num / den
    return mean_y - slope * mean_x


def KURT(*args):
    lst = _numeric_filter(flatten_args(*args))
    n = len(lst)
    if n < 4:
        raise ValueError("Need at least 4 data points")
    mean = sum(lst) / n
    s = (sum((x - mean) ** 2 for x in lst) / (n - 1)) ** 0.5
    m4 = sum(((x - mean) / s) ** 4 for x in lst)
    return (n * (n + 1) * m4 / ((n - 1) * (n - 2) * (n - 3))) - (3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))


def LARGE(data, k):
    lst = sorted(_numeric_filter(flatten_args(data)), reverse=True)
    return lst[int(k) - 1]


def LOGINV(probability, mean, standard_dev):
    raise NotImplementedError("LOGINV() not implemented yet")


class LOGNORM:
    @staticmethod
    def DIST(x, mean, standard_dev, cumulative):
        raise NotImplementedError("LOGNORM.DIST() not implemented yet")

    @staticmethod
    def INV(probability, mean, standard_dev):
        raise NotImplementedError("LOGNORM.INV() not implemented yet")


def LOGNORMDIST(x, mean, standard_dev):
    raise NotImplementedError("LOGNORMDIST() not implemented yet")


def MARGINOFERROR(confidence_level, standard_dev, size):
    raise NotImplementedError("MARGINOFERROR() not implemented yet")


def MAX(*args):
    lst = _numeric_filter(flatten_args(*args))
    return max(lst)


def MAXA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return max(converted)


def MAXIFS(max_range: Range, *criteria_pairs):
    if len(criteria_pairs) % 2:
        raise ValueError("Number of criteria arguments has to be even")
    values = []
    for idx in range(len(max_range.lst)):
        include = True
        for i in range(0, len(criteria_pairs), 2):
            cr = criteria_pairs[i]
            criterion = criteria_pairs[i + 1]
            if not criterion(cr.lst[idx]):
                include = False
                break
        if include and isinstance(max_range.lst[idx], (int, float)):
            values.append(max_range.lst[idx])
    if not values:
        return 0
    return max(values)


def MEDIAN(*args):
    lst = _numeric_filter(flatten_args(*args))
    return statistics.median(lst)


def MIN(*args):
    lst = _numeric_filter(flatten_args(*args))
    return min(lst)


def MINA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return min(converted)


def MINIFS(min_range: Range, *criteria_pairs):
    if len(criteria_pairs) % 2:
        raise ValueError("Number of criteria arguments has to be even")
    values = []
    for idx in range(len(min_range.lst)):
        include = True
        for i in range(0, len(criteria_pairs), 2):
            cr = criteria_pairs[i]
            criterion = criteria_pairs[i + 1]
            if not criterion(cr.lst[idx]):
                include = False
                break
        if include and isinstance(min_range.lst[idx], (int, float)):
            values.append(min_range.lst[idx])
    if not values:
        return 0
    return min(values)


class MODE:
    def __new__(cls, *args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.mode(lst)

    @staticmethod
    def MULT(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.multimode(lst)

    @staticmethod
    def SNGL(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.mode(lst)


class NEGBINOM:
    @staticmethod
    def DIST(number_f, number_s, probability_s, cumulative):
        raise NotImplementedError("NEGBINOM.DIST() not implemented yet")


def NEGBINOMDIST(number_f, number_s, probability_s):
    raise NotImplementedError("NEGBINOMDIST() not implemented yet")


class NORM:
    @staticmethod
    def DIST(x, mean, standard_dev, cumulative):
        nd = statistics.NormalDist(mean, standard_dev)
        if cumulative:
            return nd.cdf(x)
        return nd.pdf(x)

    @staticmethod
    def INV(probability, mean, standard_dev):
        nd = statistics.NormalDist(mean, standard_dev)
        return nd.inv_cdf(probability)

    class S:
        @staticmethod
        def DIST(z, cumulative=True):
            nd = statistics.NormalDist()
            if cumulative:
                return nd.cdf(z)
            return nd.pdf(z)

        @staticmethod
        def INV(probability):
            return statistics.NormalDist().inv_cdf(probability)


def NORMDIST(x, mean, standard_dev, cumulative):
    return NORM.DIST(x, mean, standard_dev, cumulative)


def NORMINV(probability, mean, standard_dev):
    return NORM.INV(probability, mean, standard_dev)


def NORMSDIST(z):
    return NORM.S.DIST(z, True)


def NORMSINV(probability):
    return NORM.S.INV(probability)


def PEARSON(data_y, data_x):
    y = _numeric_filter(flatten_args(data_y))
    x = _numeric_filter(flatten_args(data_x))
    return statistics.correlation(x, y)


class PERCENTILE:
    def __new__(cls, data, k):
        lst = sorted(_numeric_filter(flatten_args(data)))
        n = len(lst)
        idx = k * (n - 1)
        lo = int(idx)
        hi = lo + 1
        if hi >= n:
            return lst[lo]
        frac = idx - lo
        return lst[lo] * (1 - frac) + lst[hi] * frac

    @staticmethod
    def EXC(data, k):
        lst = sorted(_numeric_filter(flatten_args(data)))
        n = len(lst)
        idx = k * (n + 1) - 1
        lo = int(idx)
        hi = lo + 1
        if lo < 0 or hi >= n:
            raise ValueError("k is out of range for PERCENTILE.EXC")
        frac = idx - lo
        return lst[lo] * (1 - frac) + lst[hi] * frac

    @staticmethod
    def INC(data, k):
        return PERCENTILE.__new__(PERCENTILE, data, k)


class PERCENTRANK:
    def __new__(cls, data, x, significance=3):
        lst = sorted(_numeric_filter(flatten_args(data)))
        n = len(lst)
        if x < lst[0] or x > lst[-1]:
            raise ValueError("x is outside the data range")
        for i in range(n - 1):
            if lst[i] <= x <= lst[i + 1]:
                if lst[i] == lst[i + 1]:
                    frac_rank = i / (n - 1)
                else:
                    frac_rank = (i + (x - lst[i]) / (lst[i + 1] - lst[i])) / (n - 1)
                factor = 10 ** significance
                return int(frac_rank * factor) / factor
        return 1.0

    @staticmethod
    def EXC(data, x, significance=3):
        raise NotImplementedError("PERCENTRANK.EXC() not implemented yet")

    @staticmethod
    def INC(data, x, significance=3):
        return PERCENTRANK.__new__(PERCENTRANK, data, x, significance)


def PERMUT(n, k):
    return math.perm(int(n), int(k))


def PERMUTATIONA(n, k):
    return int(n) ** int(k)


def PHI(x):
    return statistics.NormalDist().pdf(x)


class POISSON:
    def __new__(cls, x, mean, cumulative):
        raise NotImplementedError("POISSON() not implemented yet")

    @staticmethod
    def DIST(x, mean, cumulative):
        raise NotImplementedError("POISSON.DIST() not implemented yet")


def PROB(x_range, prob_range, lower_limit, upper_limit=None):
    xs = flatten_args(x_range)
    ps = flatten_args(prob_range)
    if upper_limit is None:
        upper_limit = lower_limit
    total = 0
    for x, p in zip(xs, ps):
        if lower_limit <= x <= upper_limit:
            total += p
    return total


class QUARTILE:
    def __new__(cls, data, quart):
        return PERCENTILE.__new__(PERCENTILE, data, quart * 0.25)

    @staticmethod
    def EXC(data, quart):
        return PERCENTILE.EXC(data, quart * 0.25)

    @staticmethod
    def INC(data, quart):
        return PERCENTILE.__new__(PERCENTILE, data, quart * 0.25)


class RANK:
    def __new__(cls, number, ref, order=0):
        lst = sorted(_numeric_filter(flatten_args(ref)), reverse=(order == 0))
        return lst.index(number) + 1

    @staticmethod
    def AVG(number, ref, order=0):
        lst = sorted(_numeric_filter(flatten_args(ref)), reverse=(order == 0))
        positions = [i + 1 for i, v in enumerate(lst) if v == number]
        return sum(positions) / len(positions)

    @staticmethod
    def EQ(number, ref, order=0):
        return RANK.__new__(RANK, number, ref, order)


def RSQ(known_ys, known_xs):
    r = PEARSON(known_ys, known_xs)
    return r ** 2


class SKEW:
    def __new__(cls, *args):
        lst = _numeric_filter(flatten_args(*args))
        n = len(lst)
        if n < 3:
            raise ValueError("Need at least 3 data points")
        mean = sum(lst) / n
        s = (sum((x - mean) ** 2 for x in lst) / (n - 1)) ** 0.5
        return (n / ((n - 1) * (n - 2))) * sum(((x - mean) / s) ** 3 for x in lst)

    @staticmethod
    def P(*args):
        lst = _numeric_filter(flatten_args(*args))
        n = len(lst)
        if n < 3:
            raise ValueError("Need at least 3 data points")
        mean = sum(lst) / n
        m2 = sum((x - mean) ** 2 for x in lst) / n
        m3 = sum((x - mean) ** 3 for x in lst) / n
        return m3 / (m2 ** 1.5)


def SLOPE(known_ys, known_xs):
    y_vals = _numeric_filter(flatten_args(known_ys))
    x_vals = _numeric_filter(flatten_args(known_xs))
    n = len(x_vals)
    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n
    num = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x_vals, y_vals))
    den = sum((xi - mean_x) ** 2 for xi in x_vals)
    return num / den


def SMALL(data, k):
    lst = sorted(_numeric_filter(flatten_args(data)))
    return lst[int(k) - 1]


def STANDARDIZE(x, mean, standard_dev):
    return (x - mean) / standard_dev


class STDEV:
    def __new__(cls, *args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.stdev(lst)

    @staticmethod
    def P(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.pstdev(lst)

    @staticmethod
    def S(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.stdev(lst)


def STDEVA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return statistics.stdev(converted)


def STDEVP(*args):
    lst = _numeric_filter(flatten_args(*args))
    return statistics.pstdev(lst)


def STDEVPA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return statistics.pstdev(converted)


def STEYX(known_ys, known_xs):
    y_vals = _numeric_filter(flatten_args(known_ys))
    x_vals = _numeric_filter(flatten_args(known_xs))
    n = len(x_vals)
    if n < 3:
        raise ValueError("Need at least 3 data points")
    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n
    ss_xy = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x_vals, y_vals))
    ss_xx = sum((xi - mean_x) ** 2 for xi in x_vals)
    ss_yy = sum((yi - mean_y) ** 2 for yi in y_vals)
    return math.sqrt((ss_yy - ss_xy ** 2 / ss_xx) / (n - 2))


class T_STAT:
    class DIST:
        def __new__(cls, x, degrees_freedom, cumulative):
            raise NotImplementedError("T_STAT.DIST() not implemented yet")

        @staticmethod
        def _2T(x, degrees_freedom):
            raise NotImplementedError("T_STAT.DIST.2T() not implemented yet")

        @staticmethod
        def RT(x, degrees_freedom):
            raise NotImplementedError("T_STAT.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, probability, degrees_freedom):
            raise NotImplementedError("T_STAT.INV() not implemented yet")

        @staticmethod
        def _2T(probability, degrees_freedom):
            raise NotImplementedError("T_STAT.INV.2T() not implemented yet")

    @staticmethod
    def TEST(range1, range2, tails, type_):
        raise NotImplementedError("T_STAT.TEST() not implemented yet")


def TDIST(x, degrees_freedom, tails):
    raise NotImplementedError("TDIST() not implemented yet")


def TINV(probability, degrees_freedom):
    raise NotImplementedError("TINV() not implemented yet")


def TRIMMEAN(data, percent):
    lst = sorted(_numeric_filter(flatten_args(data)))
    n = len(lst)
    trim_count = int(n * percent / 2)
    trimmed = lst[trim_count:n - trim_count]
    return sum(trimmed) / len(trimmed)


def TTEST(range1, range2, tails, type_):
    raise NotImplementedError("TTEST() not implemented yet")


class VAR:
    def __new__(cls, *args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.variance(lst)

    @staticmethod
    def P(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.pvariance(lst)

    @staticmethod
    def S(*args):
        lst = _numeric_filter(flatten_args(*args))
        return statistics.variance(lst)


def VARA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return statistics.variance(converted)


def VARP(*args):
    lst = _numeric_filter(flatten_args(*args))
    return statistics.pvariance(lst)


def VARPA(*args):
    lst = flatten_args(*args)
    converted = []
    for v in lst:
        if v == EmptyCell:
            continue
        if isinstance(v, bool):
            converted.append(int(v))
        elif isinstance(v, (int, float)):
            converted.append(v)
        elif isinstance(v, str):
            converted.append(0)
    return statistics.pvariance(converted)


class WEIBULL:
    def __new__(cls, x, alpha, beta, cumulative):
        raise NotImplementedError("WEIBULL() not implemented yet")

    @staticmethod
    def DIST(x, alpha, beta, cumulative):
        raise NotImplementedError("WEIBULL.DIST() not implemented yet")


class Z:
    @staticmethod
    def TEST(array, x, sigma=None):
        raise NotImplementedError("Z.TEST() not implemented yet")


def ZTEST(array, x, sigma=None):
    raise NotImplementedError("ZTEST() not implemented yet")
