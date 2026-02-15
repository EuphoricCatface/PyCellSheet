import typing

import statistics
import math
try:
    import scipy.stats
except ImportError:
    scipy = None

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


def AVEDEV(*args):
    """
    Average absolute deviation of the data points from their mean.
    AveDev = sum( |x - mean(x)| ) / n
    """
    values = flatten_args(*args)
    if not values:
        return 0
    mean_val = sum(values) / len(values)
    return sum(abs(x - mean_val) for x in values) / len(values)


class AVERAGE:
    def __new__(cls, *args):
        """
        Computes the mean of all values in args.
        """
        lst = flatten_args(*args)
        return sum(lst) / len(lst)

    @staticmethod
    def WEIGHTED(values: Range, weights: Range):
        """
        Weighted average: sum(value_i * weight_i) / sum(weight_i)
        """
        if values.width != weights.width or values.height != weights.height:
            raise ValueError("Dimensions of values and weights don't match")

        weighted_sum = 0
        for v, w in zip(values.lst, weights.lst):
            if (v == EmptyCell) != (w == EmptyCell):
                raise ValueError("Mismatched empty cell(s) in values and weights")
            if w < 0:
                raise ValueError("At least one of the weights is negative")
            weighted_sum += v * w
        weights_sum = sum(weights.flatten())
        return weighted_sum / weights_sum


def AVERAGEA(*args):
    """
    Like AVERAGE but treats text/booleans as numeric.
    Texts are treated as 0, bool True as 1, bool False as 0.
    """
    raw_vals = flatten_args(*args)
    numeric_vals = []
    for v in raw_vals:
        if isinstance(v, (int, float)):
            numeric_vals.append(v)
        elif isinstance(v, bool):
            numeric_vals.append(int(v))
        else:
            numeric_vals.append(0)
    return sum(numeric_vals) / len(numeric_vals)


def AVERAGEIF(range_: Range, criterion: typing.Callable[[typing.Any], bool], average_range: typing.Optional[Range] = None):
    """
    AVERAGEIF(range_, criterion, [average_range]):
      - range_: the data to test with the criterion_func
      - criterion_func: a callable that returns True/False
      - average_range: the data to average (optional)

    If average_range is None, we average over the same range.
    """
    if average_range is None:
        average_range = range_

    if range_.width != average_range.width or range_.height != average_range.height:
        raise ValueError("Dimension mismatch in r and average_r")

    avg_list = []
    for test, value in zip(range_.lst, average_range.lst):
        if test == EmptyCell or value == EmptyCell:
            continue
        if criterion(test):
            avg_list.append(value)
    return sum(avg_list) / len(avg_list)


def AVERAGEIFS(average_range, *range_crit_pairs):
    """
    AVERAGEIFS(average_range, (range1, crit1), (range2, crit2), ...)
      Example usage:
        =AVERAGEIFS(B1:B10, (A1:A10, lambda x: x>5), (C1:C10, lambda x: x=='Yes'))
      Returns the average of items in average_range that pass all criteria in each range_crit pair.
    """
    avg_vals = average_range.lst
    if not range_crit_pairs:
        raise ValueError("AVERAGEIFS: no criteria provided")

    valid_indices = set(range(len(avg_vals)))
    for (rng, crit) in range_crit_pairs:
        r_vals = rng.lst
        if len(r_vals) != len(avg_vals):
            raise ValueError("AVERAGEIFS: range length differs from average_range length")
        local_indices = {i for i, v in enumerate(r_vals) if crit(v)}
        valid_indices &= local_indices

    matched_vals = [avg_vals[i] for i in valid_indices if avg_vals[i] != EmptyCell]
    return sum(matched_vals) / len(matched_vals)


class BETA:
    @staticmethod
    def DIST(x, alpha, beta, cumulative=True, A=0, B=1):
        """
        BETA.DIST
        Uses scipy.stats.beta if available, otherwise raises RuntimeError.
        Relies on SciPy to raise ValueError/TypeError for invalid params.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute BETA.DIST.")
        dist = scipy.stats.beta(alpha, beta, loc=A, scale=(B - A))
        return dist.cdf(x) if cumulative else dist.pdf(x)

    @staticmethod
    def INV(prob, alpha, beta, A=0, B=1):
        """
        BETA.INV
        Also uses scipy.stats.beta. SciPy will raise if parameters are invalid.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute BETA.INV.")
        dist = scipy.stats.beta(alpha, beta, loc=A, scale=(B - A))
        return dist.ppf(prob)


def BETADIST(a, b):
    raise NotImplementedError("BETADIST() not implemented yet")


def BETAINV(a, b):
    raise NotImplementedError("BETAINV() not implemented yet")


class BINOM:
    def __new__(cls, *args):
        raise NameError("BINOM cannot be called directly.")

    @staticmethod
    def DIST(x, n, p, cumulative):
        """
        If cumulative=False => PMF => Probability of exactly x successes
        If cumulative=True  => CDF => Probability of up to x successes
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute BINOM.DIST.")
        dist = scipy.stats.binom(n, p)
        if cumulative:
            return dist.cdf(x)
        else:
            return dist.pmf(x)

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("BINOM.INV() not implemented yet")


def BINOMDIST(a, b):
    raise NotImplementedError("BINOMDIST() not implemented yet")


def CHIDIST(a, b):
    raise NotImplementedError("CHIDIST() not implemented yet")


def CHIINV(a, b):
    raise NotImplementedError("CHIINV() not implemented yet")


class CHISQ:
    def __new__(cls, *args):
        raise NameError("CHISQ cannot be called directly.")

    class DIST:
        def __new__(cls, x, df, cumulative):
            if scipy is None:
                raise RuntimeError("SciPy is missing. Cannot compute CHISQ.DIST.")
            dist = scipy.stats.chi2(df)
            if cumulative:
                return dist.cdf(x)
            else:
                return dist.pdf(x)

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("CHISQ.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, a, b):
            raise NotImplementedError("CHISQ.INV() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("CHISQ.INV.RT() not implemented yet")

    @staticmethod
    def TEST(observed, expected):
        """
        CHISQ.TEST(observed, expected):
          Chi-square test statistic & p-value for 1D or 2D data.
          For simplicity, we assume 1D data.
          In many spreadsheets, CHISQ.TEST can handle contingency tables.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute CHISQ.TEST.")
        obs_vals = flatten_args(observed)
        exp_vals = flatten_args(expected)
        if len(obs_vals) != len(exp_vals):
            raise ValueError("CHISQ.TEST: observed and expected must have same length")

        # Use chisquare from scipy
        chi_stat, p_val = scipy.stats.chisquare(obs_vals, f_exp=exp_vals)
        # Some spreadsheets just return p_val; others might do more. We'll do p_val here.
        return p_val


def CHITEST(a, b):
    raise NotImplementedError("CHITEST() not implemented yet")


class CONFIDENCE:
    def __new__(cls, a, b):
        raise NotImplementedError("CONFIDENCE() not implemented yet")

    @staticmethod
    def NORM(alpha, std_dev, size):
        """
        CONFIDENCE.NORM(alpha, std_dev, size):
          Returns margin of error => z * (std_dev / sqrt(size))
          where z is the z-value for alpha/2 in a 2-tailed normal distribution.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute CONFIDENCE.NORM.")
        if size <= 0:
            raise ValueError("CONFIDENCE.NORM: size must be > 0")
        if alpha <= 0 or alpha >= 1:
            raise ValueError("CONFIDENCE.NORM: alpha must be in (0,1)")

        # z-value for a 2-tailed test => 1 - alpha/2
        z = scipy.stats.norm.ppf(1 - alpha / 2.0)
        return z * (std_dev / math.sqrt(size))

    @staticmethod
    def T(alpha, std_dev, size):
        """
        CONFIDENCE.T(alpha, std_dev, size):
          Similar to CONFIDENCE.NORM, but uses T distribution with df=(size-1).
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute CONFIDENCE.T.")
        if size <= 1:
            raise ValueError("CONFIDENCE.T: size must be > 1")
        if alpha <= 0 or alpha >= 1:
            raise ValueError("CONFIDENCE.T: alpha must be in (0,1)")

        df = size - 1
        # t-value for 2-tailed => 1 - alpha/2
        t_val = scipy.stats.t.ppf(1 - alpha / 2.0, df)
        return t_val * (std_dev / math.sqrt(size))


def CORREL(x_range, y_range):
    """
    CORREL(x_range, y_range):
      Returns the correlation coefficient (Pearson's r) between x_range and y_range.
      If lengths mismatch or are too small, raise an error.
    """
    return statistics.correlation(x_range, y_range)


def COUNT(*args):
    """
    Counts how many items are numeric (int or float).
    """
    return len(list(filter(
        lambda a: isinstance(a, (int, float)), flatten_args(*args)
    )))


def COUNTA(*args):
    """
    Counts how many items are not EmptyCell (i.e., not "empty").
    """
    return len(flatten_args(*args))


def COVAR(x_range, y_range):
    return COVARIANCE.S(x_range, y_range)

class COVARIANCE:
    def __new__(cls, *args):
        # Disallow direct calls like PERCENTILE(args).
        raise NameError(
            "COVARIANCE cannot be called directly. "
            "Use COVARIANCE.P(...) or COVARIANCE.S(...)."
        )
    @staticmethod
    def P(x_range: Range, y_range: Range):
        """
        Population covariance = sum((x_i - mean_x)*(y_i - mean_y)) / N
        """
        x_vals = x_range.flatten()
        y_vals = y_range.flatten()
        sample_cov = statistics.covariance(x_vals, y_vals)
        n = len(x_vals)
        return sample_cov * (n - 1) / n


    @staticmethod
    def S(x_range, y_range):
        """
        Sample covariance = sum((x_i - mean_x)*(y_i - mean_y)) / (N - 1)
        """
        return statistics.covariance(x_range, y_range)


def CRITBINOM(a, b):
    raise NotImplementedError("CRITBINOM() not implemented yet")


def DEVSQ(*args):
    """
    DEVSQ(...) = sum((x - mean)^2 for each x in the data).
    This is like the 'sum of squares of deviations'.
    """
    values = flatten_args(*args)
    mean_val = statistics.mean(values)
    return sum((v - mean_val) ** 2 for v in values)


class EXPON:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("EXPON.DIST() not implemented yet")


def EXPONDIST(a, b):
    raise NotImplementedError("EXPONDIST() not implemented yet")


class F:
    def __new__(cls, *args):
        raise NameError("F cannot be called directly. Use F.DIST(...)")

    class DIST:
        def __new__(cls, x, df1, df2, cumulative):
            """
            F distribution.
            If cumulative=True => cdf, else pdf.
            """
            if scipy is None:
                raise RuntimeError("SciPy is missing. Cannot compute F.DIST.")
            dist = scipy.stats.f(df1, df2)
            if cumulative:
                return dist.cdf(x)
            else:
                return dist.pdf(x)

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("F.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, a, b):
            raise NotImplementedError("F.INV() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("F.INV.RT() not implemented yet")

    @staticmethod
    def TEST(range1, range2):
        """
        FTEST(range1, range2):
          In Excel, FTEST returns the two-tailed p-value for the F-test,
          checking ratio of variances. We'll do a naive approach via
          scipy.stats.levene or bartlett or f_oneway. But let's do an approach:
          We do an F = var1/var2, df1 = n1-1, df2 = n2-1 => cdf => p-value * 2
          (two-tailed).
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute FTEST.")
        vals1 = flatten_args(range1)
        vals2 = flatten_args(range2)
        if len(vals1) < 2 or len(vals2) < 2:
            raise ValueError("FTEST needs at least 2 data points in each range")

        var1 = statistics.variance(vals1)  # sample var
        var2 = statistics.variance(vals2)
        df1, df2 = len(vals1) - 1, len(vals2) - 1
        if var2 == 0:
            raise ValueError("FTEST: second dataset variance is zero")

        f = var1 / var2
        dist = scipy.stats.f(df1, df2)
        # two-tailed => we check which side f is on.
        if f > 1:
            # upper tail
            p_one_tailed = 1 - dist.cdf(f)
        else:
            # lower tail
            p_one_tailed = dist.cdf(f)
        return 2 * p_one_tailed  # two-tailed p-value


def FDIST(a, b):
    raise NotImplementedError("FDIST() not implemented yet")


def FINV(a, b):
    raise NotImplementedError("FINV() not implemented yet")


def FISHER(a, b):
    raise NotImplementedError("FISHER() not implemented yet")


def FISHERINV(a, b):
    raise NotImplementedError("FISHERINV() not implemented yet")


class FORECAST:
    def __new__(cls, x0, y_range, x_range):
        cls.LINEAR(x0, y_range, x_range)

    @staticmethod
    def LINEAR(x0, y_range, x_range):
        """
        FORECAST(x, known_y, known_x):
          Basic linear regression formula: y = a + b*x
          where b = COV(x,y) / VAR(x), a = mean(y) - b*mean(x).
        This returns the predicted y for a given x.
        """
        x_vals = x_range.flatten()
        y_vals = y_range.flatten()
        result = scipy.stats.linregress(x_vals, y_vals)
        # Predicted y = intercept + slope * x0
        return result.intercept + result.slope * x0


def FTEST(range1, range2):
    return F.TEST(range1, range2)


class GAMMA:
    def __new__(cls, a, b):
        raise NotImplementedError("GAMMA() not implemented yet")

    @staticmethod
    def DIST(x, shape, scale, cumulative=False):
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute GAMMA.DIST.")
        if x < 0:
            raise ValueError("GAMMA.DIST: x must be >= 0")
        dist = scipy.stats.gamma(shape, scale=scale)
        return dist.cdf(x) if cumulative else dist.pdf(x)

    @staticmethod
    def INV(prob, shape, scale):
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute GAMMA.INV.")
        if not (0 < prob < 1):
            raise ValueError("GAMMA.INV: prob must be in (0,1)")
        dist = scipy.stats.gamma(shape, scale=scale)
        return dist.ppf(prob)


def GAMMADIST(a, b):
    raise NotImplementedError("GAMMADIST() not implemented yet")


def GAMMAINV(a, b):
    raise NotImplementedError("GAMMAINV() not implemented yet")


def GAUSS(a, b):
    raise NotImplementedError("GAUSS() not implemented yet")


def GEOMEAN(*args):
    """
    GEOMEAN(...) returns the geometric mean of all numeric values in args.
    Geometric Mean = (Product of all values)^(1/Count)
    - We skip negative/zero checks here. If there's a non-positive number,
      math.prod(...) might yield zero or negative, and that affects the result or raises.
    - Some spreadsheets raise an error if any value <= 0 in GEOMEAN.
    """
    vals = flatten_args(*args)
    if any(v <= 0 for v in vals):
        raise ValueError("GEOMEAN requires all values > 0")
    return statistics.geometric_mean(vals)


def HARMEAN(*args):
    """
    HARMEAN(...) returns the harmonic mean of the numeric values in args.
    - The built-in statistics.harmonic_mean exists in Python 3.8+
    """
    return statistics.harmonic_mean(flatten_args(*args))


class HYPGEOM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("HYPGEOM.DIST() not implemented yet")


def HYPGEOMDIST(a, b):
    raise NotImplementedError("HYPGEOMDIST() not implemented yet")


def INTERCEPT(y_range, x_range):
    """
    INTERCEPT(known_y, known_x):
      Returns the intercept 'a' in a simple linear regression
      y = a + b*x
      a = mean(y) - b*mean(x)
    """
    x_vals = x_range.flatten()
    y_vals = y_range.flatten()
    result = scipy.stats.linregress(x_vals, y_vals)
    return result.intercept


def KURT(a, b):
    if scipy is None:
        raise RuntimeError("SciPy is missing. Cannot compute KURT.")
    return scipy.stats.kurtosis(flatten_args(*args), fisher=True, bias=False)


def LARGE(data, k):
    """
    LARGE(data, k) => the k-th largest value (1-based).
    """
    vals = sorted(flatten_args(data), reverse=True)
    if not vals:
        raise ValueError("LARGE with no data")
    if k < 1 or k > len(vals):
        raise ValueError("LARGE: k is out of range")
    return vals[k - 1]


def LOGINV(a, b):
    raise NotImplementedError("LOGINV() not implemented yet")


class LOGNORM:
    def __new__(cls, *args):
        raise NameError("LOGNORM cannot be called directly.")

    @staticmethod
    def DIST(x, mean, sd, cumulative):
        """
        If x <= 0 => return 0 or error, depending on spreadsheet's approach.
        This calls scipy.stats.lognorm(s=sigma, scale=exp(mu))
        where mu=mean, sigma=sd in log scale.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute LOGNORM.DIST.")
        if x <= 0:
            raise ValueError("LOGNORM.DIST: x must be > 0")

        # According to SciPy: lognorm takes shape=sigma, scale=e^mean
        dist = scipy.stats.lognorm(sd, scale=math.exp(mean))
        return dist.cdf(x) if cumulative else dist.pdf(x)

    @staticmethod
    def INV(prob, mean, sd):
        """
        LOGNORM.INV(prob, mean, sd) => inverse of the lognormal distribution
        i.e., x such that LOGNORM.DIST(x, mean, sd, True)=prob
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute LOGNORM.INV.")
        if not (0 < prob < 1):
            raise ValueError("LOGNORM.INV: prob must be in (0,1)")

        dist = scipy.stats.lognorm(sd, scale=math.exp(mean))
        return dist.ppf(prob)


def LOGNORMDIST(a, b):
    raise NotImplementedError("LOGNORMDIST() not implemented yet")


def MARGINOFERROR(a, b):
    raise NotImplementedError("MARGINOFERROR() not implemented yet")


def MAX(*args):
    """
    Return the maximum numeric value among arguments.
    """
    flattened = flatten_args(*args)
    return max(flattened) if flattened else 0


def MAXA(a, b):
    raise NotImplementedError("MAXA() not implemented yet")


def MAXIFS(a, b):
    raise NotImplementedError("MAXIFS() not implemented yet")


def MEDIAN(*args):
    """
    MEDIAN(...) returns the median of the numeric values in args.
    """
    return statistics.median(flatten_args(*args))


def MIN(*args):
    """
    Return the minimum numeric value among arguments.
    """
    flattened = flatten_args(*args)
    return min(flattened) if flattened else 0


def MINA(a, b):
    raise NotImplementedError("MINA() not implemented yet")


def MINIFS(a, b):
    raise NotImplementedError("MINIFS() not implemented yet")


class MODE:
    @staticmethod
    def __new__(cls, *args):
        """
        MODE(...) returns the most common (frequent) value in the data.
        """
        return statistics.mode(flatten_args(*args))

    @staticmethod
    def MULT(a, b):
        raise NotImplementedError("MODE.MULT() not implemented yet")

    @staticmethod
    def SNGL(a, b):
        raise NotImplementedError("MODE.SNGL() not implemented yet")


class NEGBINOM:
    """
    NEGBINOM.DIST(k, r, p, cumulative=False)
    Interpreted as:
      Probability of k-th success on a certain trial, or the distribution
      of number of failures before r successes, etc.
    SciPy uses n, p => # successes, success prob => distribution.
    """
    def __new__(cls, *args):
        raise NameError("NEGBINOM cannot be called directly. Use NEGBINOM.DIST(...)")

    @staticmethod
    def DIST(k, r, p, cumulative):
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute NEGBINOM.DIST.")
        # SciPy param note: n = r, p = 1-p for failure-based param?
        # Actually, stats.nbinom(n, p) => # of failures before n successes with success prob p.
        # The usage can vary.
        dist = scipy.stats.nbinom(r, p)
        if cumulative:
            return dist.cdf(k)
        else:
            return dist.pmf(k)


def NEGBINOMDIST(k, r, p, cumulative):
    return NEGBINOM.DIST(k, r, p, cumulative)


class NORM:
    def __new__(cls, *args):
        raise NameError("NORM cannot be called directly.")

    @staticmethod
    def DIST(x, mean, std, cumulative):
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute NORM.DIST.")
        dist = scipy.stats.norm(mean, std)
        return dist.cdf(x) if cumulative else dist.pdf(x)

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("NORM.INV() not implemented yet")

    class S:
        @staticmethod
        def DIST(x):
            """
            Standard normal distribution (mean=0, std=1).
            """
            if scipy is None:
                raise RuntimeError("SciPy is missing. Cannot compute NORM.S.DIST.")
            dist = scipy.stats.norm(0, 1)
            return dist.cdf(x)

        @staticmethod
        def INV(a, b):
            raise NotImplementedError("NORM.S.INV() not implemented yet")


def NORMDIST(a, b):
    raise NotImplementedError("NORMDIST() not implemented yet")


def NORMINV(a, b):
    raise NotImplementedError("NORMINV() not implemented yet")


def NORMSDIST(a, b):
    raise NotImplementedError("NORMSDIST() not implemented yet")


def NORMSINV(a, b):
    raise NotImplementedError("NORMSINV() not implemented yet")


def PEARSON(data_x, data_y):
    return statistics.correlation(data_x, data_y)


class PERCENTILE:
    def __new__(cls, data, percentile):
        cls.INC(data, percentile)

    @staticmethod
    def EXC(data, percentile):
        """
        Exclusive percentile: percentile in (0..1), meaning
        0 or 1 might be disallowed or treated differently.
        Different spreadsheets do it differently:
        - Some skip interpolation at the extreme ends
        - Others rely on a known formula
        We'll do a simplistic approach similar to INC but disallow percentile=0 or 1.
        """
        values = flatten_args(data)
        if not values:
            raise ValueError("PERCENTILE.EXC with no data")
        if not (0 < percentile < 1):
            raise ValueError("k must be strictly between 0 and 1 for .EXC")
        vals = sorted(values)
        n = len(vals)
        # position in zero-based index, but skip endpoints
        pos = (n + 1) * percentile - 1
        # This is a common approach from some Excel docs (with rounding).
        lower_index = int(math.floor(pos))
        upper_index = int(math.ceil(pos))
        # If pos is out of range, you might raise or clamp it;
        # but that depends on spreadsheet definitions.
        # We'll clamp here for safety:
        if lower_index < 0:
            return vals[0]
        if upper_index >= n:
            return vals[-1]
        if lower_index == upper_index:
            return vals[lower_index]
        fraction = pos - lower_index
        return vals[lower_index] + fraction * (vals[upper_index] - vals[lower_index])

    @staticmethod
    def INC(data, percentile):
        """
        Inclusive percentile: percentile in [0..1], with endpoints 0 -> min(data), 1 -> max(data).
        We'll do a naive linear interpolation or rely on 'statistics.quantiles'.
        """
        values = flatten_args(data)
        if not values:
            raise ValueError("PERCENTILE.INC with no data")
        if not (0 <= percentile <= 1):
            raise ValueError("k must be between 0 and 1 (inclusive)")
        # Sort the data
        vals = sorted(values)
        n = len(vals)
        if percentile == 0:
            return vals[0]
        if percentile == 1:
            return vals[-1]
        # position in zero-based index
        pos = (n - 1) * percentile
        lower_index = int(math.floor(pos))
        upper_index = int(math.ceil(pos))
        if lower_index == upper_index:
            return vals[lower_index]
        # linear interpolation
        fraction = pos - lower_index
        return vals[lower_index] + fraction * (vals[upper_index] - vals[lower_index])


class PERCENTRANK:
    @staticmethod
    def __new__(cls, a, b):
        raise NotImplementedError("PERCENTRANK() not implemented yet")

    @staticmethod
    def EXC(data, x, significance=3):
        """
        PERCENTRANK.EXC(data, x, significance=3):
          - Returns percentile rank in (0,1), exclusive of endpoints.
          - Similar logic but excludes 0 or 1 for min/max if x exactly matches extremes.
          Real spreadsheets do more precise interpolation;
          this is a simplified version for demonstration.
        """
        vals = sorted(flatten_args(data))
        n = len(vals)
        if n < 2:
            raise ValueError("PERCENTRANK.EXC needs at least 2 data points")

        if x <= vals[0]:
            return 0.0
        if x >= vals[-1]:
            return 1.0

        # Count how many are strictly < x, how many are <= x, etc.
        count_lt = sum(1 for v in vals if v < x)
        frac = count_lt / (n - 1)  # naive approach
        # clamp to (0,1)
        frac = max(0.000001, min(frac, 0.999999))
        return round(frac, significance)

    @staticmethod
    def INC(data, x, significance=3):
        """
        PERCENTRANK.INC(data, x, significance=3):
          - Return the percentile rank of x in data, from 0..1 inclusive.
          - 'significance' is how many decimal places to round to.
          This is a simplified approach:
            rank = (# of values <= x) / (n - 1)
          Then clamp to [0,1].
          In real spreadsheets, there's interpolation for values not exactly in data.
        """
        vals = sorted(flatten_args(data))
        n = len(vals)
        if n < 1:
            raise ValueError("PERCENTRANK.INC with no data")

        count_le = sum(1 for v in vals if v <= x)
        # a simplistic approach. Some spreadsheets do interpolation here.
        frac = (count_le - 1) / (n - 1) if n > 1 else 0
        frac = max(0, min(frac, 1))  # clamp to [0..1]
        return round(frac, significance)


def PERMUTATIONA(a, b):
    raise NotImplementedError("PERMUTATIONA() not implemented yet")


def PERMUT(a, b):
    raise NotImplementedError("PERMUT() not implemented yet")


def PHI(a, b):
    raise NotImplementedError("PHI() not implemented yet")


class POISSON:
    def __new__(cls, x, lam, cumulative=False):
        cls.DIST(x, lam, cumulative)

    @staticmethod
    def DIST(x, lam, cumulative=False):
        """
        POISSON.DIST(x, mean, cumulative=False)
        In Excel, 'POISSON.DIST(x, lambda, cumulative)' can be either the PMF or CDF.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute POISSON.DIST.")
        if x < 0:
            raise ValueError("POISSON.DIST: x must be >= 0")
        # from scipy.stats import poisson
        dist = scipy.stats.poisson(lam)
        if cumulative:
            return dist.cdf(x)
        else:
            return dist.pmf(x)


def PROB(a, b):
    raise NotImplementedError("PROB() not implemented yet")


class QUARTILE:
    def __new__(cls, data, quartile_number):
        return cls.INC(data, quartile_number)

    @staticmethod
    def EXC(data, quartile_number):
        """
        QUARTILE.EXC(data, quart)
        quart in {1, 2, 3} typically, ignoring 0..4 or skipping them as "exclusive".
        Some spreadsheets treat 0/4 differently (like min/max).
        We'll do a partial approach:
            - 1 => 25% (exclusive)
            - 2 => 50% (exclusive)
            - 3 => 75% (exclusive)
        """
        if quartile_number not in [1, 2, 3]:
            raise ValueError("QUARTILE.EXC quart must be 1..3")
        vals = flatten_args(data)

        vals = sorted(vals)
        k = quartile_number * 0.25  # 1 => 0.25, 2 => 0.5, 3 => 0.75
        return PERCENTILE.EXC(vals, k)

    @staticmethod
    def INC(data, quartile_number):
        """
        QUARTILE.INC(data, quart)
        quart in {0, 1, 2, 3, 4}
        - 0 => min
        - 4 => max
        - 1 => 25th percentile
        - 2 => 50th percentile (median)
        - 3 => 75th percentile
        Uses a PERCENTILE.INC approach under the hood.
        """
        if quartile_number not in [0, 1, 2, 3, 4]:
            raise ValueError("QUARTILE.INC quart must be 0..4")
        vals = flatten_args(data)

        vals = sorted(vals)
        if quartile_number == 0:
            return vals[0]
        if quartile_number == 4:
            return vals[-1]
        # For 1..3, we do percentiles. 1->25%, 2->50%, 3->75%
        k = quartile_number * 0.25  # 1 => 0.25, 2 => 0.5, 3 => 0.75
        return PERCENTILE.INC(vals, k)


class RANK:
    def __new__(cls, a, b):
        raise NotImplementedError("RANK() not implemented yet")

    @staticmethod
    def AVG(value, data, order=0):
        """
        RANK.AVG(value, data, order=0):
          - If order=0 (descending): largest => rank 1, then 2, etc.
          - If order<>0 (ascending): smallest => rank 1, etc.
          - RANK.AVG handles ties by averaging the positions.
        Example: data=[10, 10, 8, 7], value=10 => average rank of positions
                 for the tied items.
        """
        vals = flatten_args(data)
        if not vals:
            raise ValueError("RANK.AVG with no data")

        # Sort ascending or descending
        if order == 0:
            # descending
            sorted_vals = sorted(vals, reverse=True)
        else:
            # ascending
            sorted_vals = sorted(vals)

        # Identify positions of 'value' in the sorted list
        matches = [i for i, v in enumerate(sorted_vals) if v == value]
        if not matches:
            # value not in data => some spreadsheets might still do an approximate rank
            # but we’ll raise an error or return None
            raise ValueError("RANK.AVG: value not found in data")

        # The rank is average of (i+1) for each match i
        # e.g. if the item is at indices [0,1], rank => (1+2)/2 = 1.5
        # +1 because ranks are 1-based, not 0-based
        ranks = [(m + 1) for m in matches]
        return sum(ranks) / len(ranks)

    @staticmethod
    def EQ(value, data, order=0):
        """
        RANK.EQ mimics a typical spreadsheet approach:
        - If order=0 => descending rank
        - If order<>0 => ascending rank
        """
        vals = flatten_args(data)
        if not vals:
            raise ValueError("RANK.EQ with no data")
        # Descending rank => bigger numbers get rank 1
        if order == 0:
            sorted_vals = sorted(vals, reverse=True)
        else:
            sorted_vals = sorted(vals)

        # Python's 'list.index()' finds the first occurrence
        # but we might want to handle duplicates.
        # RANK.EQ in many spreadsheets basically returns
        # 1 + (number of items that exceed 'value').
        if order == 0:  # descending
            rank_val = 1 + sum(1 for v in sorted_vals if v > value)
        else:           # ascending
            rank_val = 1 + sum(1 for v in sorted_vals if v < value)
        return rank_val


def RSQ(y_range: Range, x_range: Range):
    """
    RSQ(y_range, x_range):
      Returns the square of the Pearson correlation coefficient
      (coefficient of determination, r^2).
    """
    y_vals = y_range.flatten()
    x_vals = x_range.flatten()
    r = statistics.correlation(x_vals, y_vals)
    return r * r


class SKEW:
    def __new__(cls, *args):
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute SKEW.")
        values = flatten_args(*args)
        if not values:
            raise ValueError("SKEW called with no data")
        # By default, scipy.stats.skew uses Fisher's definition of skewness.
        # bias=False => corrected skew
        return scipy.stats.skew(values, bias=False)

    @staticmethod
    def P(a, b):
        raise NotImplementedError("SKEW.P() not implemented yet")


def SLOPE(x_range, y_range):
    """
    SLOPE(known_y, known_x):
      Returns the slope 'b' in a simple linear regression
      y = a + b*x
      Slope = [Σ(xy) - (Σx)(Σy)/n ] / [Σ(x^2) - (Σx)^2 / n]
      or equivalently => Cov(x,y)/Var(x).
    """
    x_vals = x_range.flatten()
    y_vals = y_range.flatten()
    result = scipy.stats.linregress(x_vals, y_vals)
    return result.slope


def SMALL(data, k):
    """
    SMALL(data, k) => the k-th smallest value in data (1-based).
    """
    vals = sorted(flatten_args(data))
    if not vals:
        raise ValueError("SMALL with no data")
    if k < 1 or k > len(vals):
        raise ValueError("SMALL: k is out of range")
    return vals[k - 1]


def STANDARDIZE(a, b):
    raise NotImplementedError("STANDARDIZE() not implemented yet")


class STDEV:
    def __new__(cls, *args):
        return cls.S(*args)

    @staticmethod
    def P(*args):
        """
        STDEV.P - population standard deviation
        """
        return statistics.pstdev(flatten_args(*args))

    @staticmethod
    def S(*args):
        return statistics.stdev(flatten_args(*args))


def STDEVA(*args):
    """
    STDEVA - STDEV but interpret text/booleans as numeric
    """
    processed = []
    for v in flatten_args(*args):
        if isinstance(v, (int, float)):
            processed.append(v)
        elif isinstance(v, bool):
            processed.append(int(v))
        else:
            # text => 0, etc.
            processed.append(0)
    if len(processed) < 2:
        return 0
    return statistics.stdev(processed)


def STDEVP(a, b):
    raise NotImplementedError("STDEVP() not implemented yet")


def STDEVPA(a, b):
    raise NotImplementedError("STDEVPA() not implemented yet")


def STEYX(a, b):
    raise NotImplementedError("STEYX() not implemented yet")


class T_STAT:
    class DIST:
        def __new__(cls, x, df, cumulative):
            """
            T distribution with 'df' degrees of freedom.
            If cumulative=True => CDF, else PDF.
            """
            if scipy is None:
                raise RuntimeError("SciPy is missing. Cannot compute T.DIST.")
            dist = scipy.stats.t(df)
            if cumulative:
                return dist.cdf(x)
            else:
                return dist.pdf(x)

        @staticmethod
        def _2T(a, b):
            raise NotImplementedError("T.DIST.2T() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("T.DIST.RT() not implemented yet")

    class INV:
        def __new__(cls, prob, df):
            """
            T.INV(prob, df):
              prob in (0,1)
              df => degrees of freedom
            """
            if scipy is None:
                raise RuntimeError("SciPy is missing. Cannot compute T_.INV.")
            if not (0 < prob < 1):
                raise ValueError("T_.INV: prob must be in (0,1)")

            return scipy.stats.t.ppf(prob, df)

        @staticmethod
        def _2T(a, b):
            raise NotImplementedError("T.INV.2T() not implemented yet")

    @staticmethod
    def TEST(range1, range2, tails=2, type_=2):
        """
        TTEST(range1, range2, tails, type):
          - tails=1 or 2 (one-tailed or two-tailed).
          - type=1..3 (paired, two-sample equal var, two-sample unequal var).
          We'll do a minimal approach using scipy.stats:
            - type=1 => paired => ttest_rel
            - type=2 => two-sample equal var => ttest_ind(..., equal_var=True)
            - type=3 => two-sample unequal var => ttest_ind(..., equal_var=False)
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute TTEST.")
        vals1 = flatten_args(range1)
        vals2 = flatten_args(range2)
        if not vals1 or not vals2:
            raise ValueError("TTEST with empty data")

        # pick the function
        if type_ == 1:
            # paired
            if len(vals1) != len(vals2):
                raise ValueError("TTEST type=1 (paired) requires same length data")
            t_stat, p_val = scipy.stats.ttest_rel(vals1, vals2)
        elif type_ == 2:
            t_stat, p_val = scipy.stats.ttest_ind(vals1, vals2, equal_var=True)
        elif type_ == 3:
            t_stat, p_val = scipy.stats.ttest_ind(vals1, vals2, equal_var=False)
        else:
            raise ValueError("TTEST: invalid type (must be 1..3)")

        if tails == 1:
            # one-tailed => p_val is two-tailed from the test => divide by 2
            return p_val / 2
        elif tails == 2:
            return p_val
        else:
            raise ValueError("TTEST: tails must be 1 or 2")


def TDIST(a, b):
    raise NotImplementedError("TDIST() not implemented yet")


def TINV(a, b):
    raise NotImplementedError("TINV() not implemented yet")


def TRIMMEAN(*args, proportiontocut=0.1):
    """
    TRIMMEAN(data, proportiontocut=0.1):
      Returns the mean of the data excluding a fraction of high/low values.
      proportiontocut is how much to trim from each tail. e.g., 0.1 => cut top 10% and bottom 10%.
    """
    vals = flatten_args(*args)
    if not vals:
        raise ValueError("TRIMMEAN called with no data")
    if not (0 <= proportiontocut < 0.5):
        raise ValueError("proportiontocut must be in [0, 0.5)")

    sorted_vals = sorted(vals)
    n = len(sorted_vals)
    cut = int(n * proportiontocut)  # number of elements to cut from each end
    trimmed = sorted_vals[cut : n - cut]
    if not trimmed:
        raise ValueError("TRIMMEAN: proportiontocut is too large, no data remains")
    return statistics.mean(trimmed)


def TTEST(range1, range2, tails=2, type_=2):
    T.TEST(range1, range2, tails, type_)


class VAR:
    def __new__(cls, *args):
        return cls.S(*args)

    @staticmethod
    def P(*args):
        return statistics.pvariance(flatten_args(*args))

    @staticmethod
    def S(*args):
        return statistics.variance(flatten_args(*args))


def VARA(a, b):
    raise NotImplementedError("VARA() not implemented yet")


def VARP(a, b):
    raise NotImplementedError("VARP() not implemented yet")


def VARPA(a, b):
    raise NotImplementedError("VARPA() not implemented yet")


class WEIBULL:
    def __new__(cls, a, b):
        raise NotImplementedError("WEIBULL() not implemented yet")

    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("WEIBULL.DIST() not implemented yet")


class Z:
    @staticmethod
    def TEST(range_, x, sigma=None):
        """
        Z.TEST(range, x, [sigma]):
          Usually returns the one-tailed p-value for a z-test, comparing
          sample mean vs. hypothesized population mean x, with known sigma (optional).
          Some spreadsheets do =1 - Norm.S.Dist((mean - x)/ (stdev/sqrt(n)), True).
          We'll do a simplified approach.
        """
        if scipy is None:
            raise RuntimeError("SciPy is missing. Cannot compute Z.TEST.")
        vals = flatten_args(range_)
        n = len(vals)
        if n < 1:
            raise ValueError("Z.TEST: no data")

        mean_val = statistics.mean(vals)
        if sigma is None:
            # use sample stdev
            if n < 2:
                raise ValueError("Z.TEST with no sigma requires at least 2 data points")
            stdev = statistics.pstdev(vals)  # or sample stdev, depending on spreadsheet
        else:
            stdev = sigma
        se = stdev / math.sqrt(n)  # standard error
        z = (mean_val - x) / se
        # For a one-tailed test, we want the area to the right if mean_val > x, else to the left
        # i.e. 1 - cdf if z>0, else cdf
        cdf_val = scipy.stats.norm.cdf(z)
        if z > 0:
            return 1 - cdf_val
        else:
            return cdf_val


def ZTEST(range_, x, sigma=None):
    Z.TEST(range_, x, sigma)
