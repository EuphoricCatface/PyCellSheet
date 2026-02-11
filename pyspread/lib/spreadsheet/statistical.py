import math
import statistics

try:
    from scipy import stats as scipy_stats
except ImportError:
    scipy_stats = None

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
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use BETA.DIST")
        if cumulative:
            return scipy_stats.beta.cdf(x, alpha, beta)
        return scipy_stats.beta.pdf(x, alpha, beta)

    @staticmethod
    def INV(probability, alpha, beta):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use BETA.INV")
        return scipy_stats.beta.ppf(probability, alpha, beta)


def BETADIST(x, alpha, beta):
    return BETA.DIST(x, alpha, beta, cumulative=True)


def BETAINV(probability, alpha, beta):
    return BETA.INV(probability, alpha, beta)


class BINOM:
    @staticmethod
    def DIST(number_s, trials, probability_s, cumulative):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use BINOM.DIST")
        k = int(number_s)
        n = int(trials)
        p = float(probability_s)
        if cumulative:
            return scipy_stats.binom.cdf(k, n, p)
        return scipy_stats.binom.pmf(k, n, p)

    @staticmethod
    def INV(trials, probability_s, alpha):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use BINOM.INV")
        n = int(trials)
        p = float(probability_s)
        return scipy_stats.binom.ppf(alpha, n, p)


def BINOMDIST(number_s, trials, probability_s, cumulative):
    return BINOM.DIST(number_s, trials, probability_s, cumulative)


def CHIDIST(x, degrees_freedom):
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use CHIDIST")
    return scipy_stats.chi2.sf(x, degrees_freedom)


def CHIINV(probability, degrees_freedom):
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use CHIINV")
    return scipy_stats.chi2.isf(probability, degrees_freedom)


class CHISQ:
    class DIST:
        def __new__(cls, x, degrees_freedom, cumulative):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use CHISQ.DIST")
            if cumulative:
                return scipy_stats.chi2.cdf(x, degrees_freedom)
            return scipy_stats.chi2.pdf(x, degrees_freedom)

        @staticmethod
        def RT(x, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use CHISQ.DIST.RT")
            return scipy_stats.chi2.sf(x, degrees_freedom)

    class INV:
        def __new__(cls, probability, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use CHISQ.INV")
            return scipy_stats.chi2.ppf(probability, degrees_freedom)

        @staticmethod
        def RT(probability, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use CHISQ.INV.RT")
            return scipy_stats.chi2.isf(probability, degrees_freedom)

    @staticmethod
    def TEST(actual_range, expected_range):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use CHISQ.TEST")
        actual = flatten_args(actual_range)
        expected = flatten_args(expected_range)
        chi2_stat = sum((a - e) ** 2 / e for a, e in zip(actual, expected))
        df = len(actual) - 1
        return scipy_stats.chi2.sf(chi2_stat, df)


def CHITEST(actual_range, expected_range):
    return CHISQ.TEST(actual_range, expected_range)


class CONFIDENCE:
    def __new__(cls, alpha, standard_dev, size):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use CONFIDENCE")
        z = scipy_stats.norm.ppf(1 - alpha / 2)
        return z * standard_dev / math.sqrt(size)

    @staticmethod
    def NORM(alpha, standard_dev, size):
        return CONFIDENCE.__new__(CONFIDENCE, alpha, standard_dev, size)

    @staticmethod
    def T(alpha, standard_dev, size):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use CONFIDENCE.T")
        t_val = scipy_stats.t.ppf(1 - alpha / 2, size - 1)
        return t_val * standard_dev / math.sqrt(size)


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
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use CRITBINOM")
    return scipy_stats.binom.ppf(alpha, int(trials), float(probability_s))


def DEVSQ(*args):
    lst = _numeric_filter(flatten_args(*args))
    mean = sum(lst) / len(lst)
    return sum((x - mean) ** 2 for x in lst)


class EXPON:
    @staticmethod
    def DIST(x, _lambda, cumulative):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use EXPON.DIST")
        if cumulative:
            return scipy_stats.expon.cdf(x, scale=1/_lambda)
        return scipy_stats.expon.pdf(x, scale=1/_lambda)


def EXPONDIST(x, _lambda, cumulative):
    return EXPON.DIST(x, _lambda, cumulative)


class F:
    class DIST:
        def __new__(cls, x, deg_freedom1, deg_freedom2, cumulative):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use F.DIST")
            if cumulative:
                return scipy_stats.f.cdf(x, deg_freedom1, deg_freedom2)
            return scipy_stats.f.pdf(x, deg_freedom1, deg_freedom2)

        @staticmethod
        def RT(x, deg_freedom1, deg_freedom2):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use F.DIST.RT")
            return scipy_stats.f.sf(x, deg_freedom1, deg_freedom2)

    class INV:
        def __new__(cls, probability, deg_freedom1, deg_freedom2):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use F.INV")
            return scipy_stats.f.ppf(probability, deg_freedom1, deg_freedom2)

        @staticmethod
        def RT(probability, deg_freedom1, deg_freedom2):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use F.INV.RT")
            return scipy_stats.f.isf(probability, deg_freedom1, deg_freedom2)

    @staticmethod
    def TEST(array1, array2):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use F.TEST")
        arr1 = _numeric_filter(flatten_args(array1))
        arr2 = _numeric_filter(flatten_args(array2))
        var1 = statistics.variance(arr1)
        var2 = statistics.variance(arr2)
        f_stat = var1 / var2 if var1 > var2 else var2 / var1
        df1 = len(arr1) - 1 if var1 > var2 else len(arr2) - 1
        df2 = len(arr2) - 1 if var1 > var2 else len(arr1) - 1
        return 2 * min(scipy_stats.f.cdf(f_stat, df1, df2), scipy_stats.f.sf(f_stat, df1, df2))


def FDIST(x, deg_freedom1, deg_freedom2):
    return F.DIST.RT(x, deg_freedom1, deg_freedom2)


def FINV(probability, deg_freedom1, deg_freedom2):
    return F.INV.RT(probability, deg_freedom1, deg_freedom2)


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
    return F.TEST(array1, array2)


class GAMMA:
    def __new__(cls, value):
        return math.gamma(value)

    @staticmethod
    def DIST(x, alpha, beta, cumulative):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use GAMMA.DIST")
        if cumulative:
            return scipy_stats.gamma.cdf(x, alpha, scale=beta)
        return scipy_stats.gamma.pdf(x, alpha, scale=beta)

    @staticmethod
    def INV(probability, alpha, beta):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use GAMMA.INV")
        return scipy_stats.gamma.ppf(probability, alpha, scale=beta)


def GAMMADIST(x, alpha, beta, cumulative):
    return GAMMA.DIST(x, alpha, beta, cumulative)


def GAMMAINV(probability, alpha, beta):
    return GAMMA.INV(probability, alpha, beta)


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
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use HYPGEOM.DIST")
        if cumulative:
            return scipy_stats.hypergeom.cdf(sample_s, number_pop, population_s, number_sample)
        return scipy_stats.hypergeom.pmf(sample_s, number_pop, population_s, number_sample)


def HYPGEOMDIST(sample_s, number_sample, population_s, number_pop):
    return HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative=False)


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
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use LOGINV")
    return scipy_stats.lognorm.ppf(probability, standard_dev, scale=math.exp(mean))


class LOGNORM:
    @staticmethod
    def DIST(x, mean, standard_dev, cumulative):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use LOGNORM.DIST")
        if cumulative:
            return scipy_stats.lognorm.cdf(x, standard_dev, scale=math.exp(mean))
        return scipy_stats.lognorm.pdf(x, standard_dev, scale=math.exp(mean))

    @staticmethod
    def INV(probability, mean, standard_dev):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use LOGNORM.INV")
        return scipy_stats.lognorm.ppf(probability, standard_dev, scale=math.exp(mean))


def LOGNORMDIST(x, mean, standard_dev):
    return LOGNORM.DIST(x, mean, standard_dev, cumulative=True)


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
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use NEGBINOM.DIST")
        if cumulative:
            return scipy_stats.nbinom.cdf(number_f, number_s, probability_s)
        return scipy_stats.nbinom.pmf(number_f, number_s, probability_s)


def NEGBINOMDIST(number_f, number_s, probability_s):
    return NEGBINOM.DIST(number_f, number_s, probability_s, cumulative=False)


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
        lst = sorted(_numeric_filter(flatten_args(data)))
        n = len(lst)
        if x < lst[0] or x > lst[-1]:
            raise ValueError("x is outside the data range")
        for i in range(n - 1):
            if lst[i] <= x <= lst[i + 1]:
                if lst[i] == lst[i + 1]:
                    frac_rank = (i + 1) / (n + 1)
                else:
                    frac_rank = ((i + 1) + (x - lst[i]) / (lst[i + 1] - lst[i])) / (n + 1)
                factor = 10 ** significance
                return int(frac_rank * factor) / factor
        return n / (n + 1)

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
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use POISSON")
        k = int(x)
        if cumulative:
            return scipy_stats.poisson.cdf(k, mean)
        return scipy_stats.poisson.pmf(k, mean)

    @staticmethod
    def DIST(x, mean, cumulative):
        return POISSON.__new__(POISSON, x, mean, cumulative)


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
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use T_STAT.DIST")
            if cumulative:
                return scipy_stats.t.cdf(x, degrees_freedom)
            return scipy_stats.t.pdf(x, degrees_freedom)

        @staticmethod
        def _2T(x, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use T_STAT.DIST.2T")
            return 2 * scipy_stats.t.sf(abs(x), degrees_freedom)

        @staticmethod
        def RT(x, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use T_STAT.DIST.RT")
            return scipy_stats.t.sf(x, degrees_freedom)

    class INV:
        def __new__(cls, probability, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use T_STAT.INV")
            return scipy_stats.t.ppf(probability, degrees_freedom)

        @staticmethod
        def _2T(probability, degrees_freedom):
            if scipy_stats is None:
                raise NotImplementedError("Install `scipy` python package to use T_STAT.INV.2T")
            return scipy_stats.t.ppf(1 - probability / 2, degrees_freedom)

    @staticmethod
    def TEST(range1, range2, tails, type_):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use T_STAT.TEST")
        arr1 = _numeric_filter(flatten_args(range1))
        arr2 = _numeric_filter(flatten_args(range2))
        if type_ == 1:
            # Paired
            result = scipy_stats.ttest_rel(arr1, arr2)
        elif type_ == 2:
            # Two-sample equal variance (pooled)
            result = scipy_stats.ttest_ind(arr1, arr2, equal_var=True)
        elif type_ == 3:
            # Two-sample unequal variance (Welch's)
            result = scipy_stats.ttest_ind(arr1, arr2, equal_var=False)
        else:
            raise ValueError("type_ must be 1, 2, or 3")
        p_value = result.pvalue
        if tails == 1:
            return p_value / 2
        return p_value


def TDIST(x, degrees_freedom, tails):
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use TDIST")
    if tails == 1:
        return scipy_stats.t.sf(x, degrees_freedom)
    elif tails == 2:
        return 2 * scipy_stats.t.sf(abs(x), degrees_freedom)
    raise ValueError("tails must be 1 or 2")


def TINV(probability, degrees_freedom):
    if scipy_stats is None:
        raise NotImplementedError("Install `scipy` python package to use TINV")
    return scipy_stats.t.ppf(1 - probability / 2, degrees_freedom)


def TRIMMEAN(data, percent):
    lst = sorted(_numeric_filter(flatten_args(data)))
    n = len(lst)
    trim_count = int(n * percent / 2)
    trimmed = lst[trim_count:n - trim_count]
    return sum(trimmed) / len(trimmed)


def TTEST(range1, range2, tails, type_):
    return T_STAT.TEST(range1, range2, tails, type_)


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
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use WEIBULL")
        if cumulative:
            return scipy_stats.weibull_min.cdf(x, alpha, scale=beta)
        return scipy_stats.weibull_min.pdf(x, alpha, scale=beta)

    @staticmethod
    def DIST(x, alpha, beta, cumulative):
        return WEIBULL.__new__(WEIBULL, x, alpha, beta, cumulative)


class Z:
    @staticmethod
    def TEST(array, x, sigma=None):
        if scipy_stats is None:
            raise NotImplementedError("Install `scipy` python package to use Z.TEST")
        arr = _numeric_filter(flatten_args(array))
        n = len(arr)
        mean = sum(arr) / n
        if sigma is None:
            sigma = statistics.stdev(arr)
        z = (mean - x) / (sigma / math.sqrt(n))
        return scipy_stats.norm.sf(z)


def ZTEST(array, x, sigma=None):
    return Z.TEST(array, x, sigma)
