try:
    from pyspread.lib.pycellsheet import EmptyCell, Range
except ImportError:
    from lib.pycellsheet import EmptyCell, Range

__all__ = [
    'AVEDEV', 'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'BETA', 'BETADIST', 'BETAINV',
    'BINOM', 'BINOMDIST', 'CHIDIST', 'CHIINV', 'CHISQ', 'CHITEST', 'CONFIDENCE', 'CORREL', 'COUNT',
    'COUNTA', 'COVAR', 'COVARIANCE', 'CRITBINOM', 'DEVSQ', 'EXPON', 'EXPONDIST', 'F', 'FDIST',
    'FINV', 'FISHER', 'FISHERINV', 'FORECAST', 'FTEST', 'GAMMA', 'GAMMADIST', 'GAMMAINV', 'GAUSS',
    'GEOMEAN', 'HARMEAN', 'HYPGEOM', 'HYPGEOMDIST', 'INTERCEPT', 'KURT', 'LARGE', 'LOGINV',
    'LOGNORM', 'LOGNORMDIST', 'MARGINOFERROR', 'MAX', 'MAXA', 'MAXIFS', 'MEDIAN', 'MIN', 'MINA',
    'MINIFS', 'MODE', 'NEGBINOM', 'NEGBINOMDIST', 'NORM', 'NORMDIST', 'NORMINV', 'NORMSDIST',
    'NORMSINV', 'PEARSON', 'PERCENTILE', 'PERCENTRANK', 'PERMUT', 'PERMUTATIONA', 'PHI', 'POISSON',
    'PROB', 'QUARTILE', 'RANK', 'RSQ', 'SKEW', 'SLOPE', 'SMALL', 'STANDARDIZE', 'STDEV', 'STDEVA',
    'STDEVP', 'STDEVPA', 'STEYX', 'T', 'TDIST', 'TINV', 'TRIMMEAN', 'TTEST', 'VAR', 'VARA', 'VARP',
    'VARPA', 'WEIBULL', 'Z', 'ZTEST'
]


def AVEDEV(a, b):
    raise NotImplementedError("AVEDEV() not implemented yet")


class AVERAGE:
    def __new__(cls, *args):
        lst = []
        for arg in args:
            if isinstance(arg, Range):
                lst.extend(arg.flatten())
                continue
            if isinstance(arg, list):
                lst.extend(arg)
                continue
            lst.append(arg)
        return sum(filter(lambda a: a != EmptyCell, lst))

    @staticmethod
    def WEIGHTED(a, b):
        raise NotImplementedError("AVERAGE.WEIGHTED() not implemented yet")


def AVERAGEA(a, b):
    raise NotImplementedError("AVERAGEA() not implemented yet")


def AVERAGEIF(a, b):
    raise NotImplementedError("AVERAGEIF() not implemented yet")


def AVERAGEIFS(a, b):
    raise NotImplementedError("AVERAGEIFS() not implemented yet")


class BETA:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("BETA.DIST() not implemented yet")

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("BETA.INV() not implemented yet")


def BETADIST(a, b):
    raise NotImplementedError("BETADIST() not implemented yet")


def BETAINV(a, b):
    raise NotImplementedError("BETAINV() not implemented yet")


class BINOM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("BINOM.DIST() not implemented yet")

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
    class DIST:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("CHISQ.DIST() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("CHISQ.DIST.RT() not implemented yet")

    class INV:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("CHISQ.INV() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("CHISQ.INV.RT() not implemented yet")

    @staticmethod
    def TEST(a, b):
        raise NotImplementedError("CHISQ.TEST() not implemented yet")


def CHITEST(a, b):
    raise NotImplementedError("CHITEST() not implemented yet")


class CONFIDENCE:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("CONFIDENCE() not implemented yet")

    @staticmethod
    def NORM(a, b):
        raise NotImplementedError("CONFIDENCE.NORM() not implemented yet")

    @staticmethod
    def T(a, b):
        raise NotImplementedError("CONFIDENCE.T() not implemented yet")


def CORREL(a, b):
    raise NotImplementedError("CORREL() not implemented yet")


def COUNT(a, b):
    raise NotImplementedError("COUNT() not implemented yet")


def COUNTA(a, b):
    raise NotImplementedError("COUNTA() not implemented yet")


def COVAR(a, b):
    raise NotImplementedError("COVAR() not implemented yet")

class COVARIANCE:
    @staticmethod
    def P(a, b):
        raise NotImplementedError("COVARIANCE.P() not implemented yet")

    @staticmethod
    def S(a, b):
        raise NotImplementedError("COVARIANCE.S() not implemented yet")


def CRITBINOM(a, b):
    raise NotImplementedError("CRITBINOM() not implemented yet")


def DEVSQ(a, b):
    raise NotImplementedError("DEVSQ() not implemented yet")


class EXPON:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("EXPON.DIST() not implemented yet")


def EXPONDIST(a, b):
    raise NotImplementedError("EXPONDIST() not implemented yet")


class F:
    class DIST:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("F.DIST() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("F.DIST.RT() not implemented yet")

    class INV:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("F.INV() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("F.INV.RT() not implemented yet")

    @staticmethod
    def TEST(a, b):
        raise NotImplementedError("F.TEST() not implemented yet")


def FDIST(a, b):
    raise NotImplementedError("FDIST() not implemented yet")


def FINV(a, b):
    raise NotImplementedError("FINV() not implemented yet")


def FISHER(a, b):
    raise NotImplementedError("FISHER() not implemented yet")


def FISHERINV(a, b):
    raise NotImplementedError("FISHERINV() not implemented yet")


class FORECAST:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("FORECAST() not implemented yet")

    @staticmethod
    def LINEAR(a, b):
        raise NotImplementedError("FORECAST.LINEAR() not implemented yet")


def FTEST(a, b):
    raise NotImplementedError("FTEST() not implemented yet")


class GAMMA:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("GAMMA() not implemented yet")

    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("GAMMA.DIST() not implemented yet")

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("GAMMA.INV() not implemented yet")


def GAMMADIST(a, b):
    raise NotImplementedError("GAMMADIST() not implemented yet")


def GAMMAINV(a, b):
    raise NotImplementedError("GAMMAINV() not implemented yet")


def GAUSS(a, b):
    raise NotImplementedError("GAUSS() not implemented yet")


def GEOMEAN(a, b):
    raise NotImplementedError("GEOMEAN() not implemented yet")


def HARMEAN(a, b):
    raise NotImplementedError("HARMEAN() not implemented yet")


class HYPGEOM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("HYPGEOM.DIST() not implemented yet")


def HYPGEOMDIST(a, b):
    raise NotImplementedError("HYPGEOMDIST() not implemented yet")


def INTERCEPT(a, b):
    raise NotImplementedError("INTERCEPT() not implemented yet")


def KURT(a, b):
    raise NotImplementedError("KURT() not implemented yet")


def LARGE(a, b):
    raise NotImplementedError("LARGE() not implemented yet")


def LOGINV(a, b):
    raise NotImplementedError("LOGINV() not implemented yet")


class LOGNORM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("LOGNORM.DIST() not implemented yet")

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("LOGNORM.INV() not implemented yet")


def LOGNORMDIST(a, b):
    raise NotImplementedError("LOGNORMDIST() not implemented yet")


def MARGINOFERROR(a, b):
    raise NotImplementedError("MARGINOFERROR() not implemented yet")


def MAX(a, b):
    raise NotImplementedError("MAX() not implemented yet")


def MAXA(a, b):
    raise NotImplementedError("MAXA() not implemented yet")


def MAXIFS(a, b):
    raise NotImplementedError("MAXIFS() not implemented yet")


def MEDIAN(a, b):
    raise NotImplementedError("MEDIAN() not implemented yet")


def MIN(a, b):
    raise NotImplementedError("MIN() not implemented yet")


def MINA(a, b):
    raise NotImplementedError("MINA() not implemented yet")


def MINIFS(a, b):
    raise NotImplementedError("MINIFS() not implemented yet")


class MODE:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("MODE() not implemented yet")

    @staticmethod
    def MULT(a, b):
        raise NotImplementedError("MODE.MULT() not implemented yet")

    @staticmethod
    def SNGL(a, b):
        raise NotImplementedError("MODE.SNGL() not implemented yet")


class NEGBINOM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("NEGBINOM.DIST() not implemented yet")


def NEGBINOMDIST(a, b):
    raise NotImplementedError("NEGBINOMDIST() not implemented yet")


class NORM:
    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("NORM.DIST() not implemented yet")

    @staticmethod
    def INV(a, b):
        raise NotImplementedError("NORM.INV() not implemented yet")

    class S:
        @staticmethod
        def DIST(a, b):
            raise NotImplementedError("NORM.S.DIST() not implemented yet")

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


def PEARSON(a, b):
    raise NotImplementedError("PEARSON() not implemented yet")


class PERCENTILE:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("PERCENTILE() not implemented yet")

    @staticmethod
    def EXC(a, b):
        raise NotImplementedError("PERCENTILE.EXC() not implemented yet")

    @staticmethod
    def INC(a, b):
        raise NotImplementedError("PERCENTILE.INC() not implemented yet")


class PERCENTRANK:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("PERCENTRANK() not implemented yet")

    @staticmethod
    def EXC(a, b):
        raise NotImplementedError("PERCENTRANK.EXC() not implemented yet")

    @staticmethod
    def INC(a, b):
        raise NotImplementedError("PERCENTRANK.INC() not implemented yet")


def PERMUTATIONA(a, b):
    raise NotImplementedError("PERMUTATIONA() not implemented yet")


def PERMUT(a, b):
    raise NotImplementedError("PERMUT() not implemented yet")


def PHI(a, b):
    raise NotImplementedError("PHI() not implemented yet")


class POISSON:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("POISSON() not implemented yet")

    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("POISSON.DIST() not implemented yet")


def PROB(a, b):
    raise NotImplementedError("PROB() not implemented yet")


class QUARTILE:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("QUARTILE() not implemented yet")

    @staticmethod
    def EXC(a, b):
        raise NotImplementedError("QUARTILE.EXC() not implemented yet")

    @staticmethod
    def INC(a, b):
        raise NotImplementedError("QUARTILE.INC() not implemented yet")


class RANK:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("RANK() not implemented yet")

    @staticmethod
    def AVG(a, b):
        raise NotImplementedError("RANK.AVG() not implemented yet")

    @staticmethod
    def EQ(a, b):
        raise NotImplementedError("RANK.EQ() not implemented yet")


def RSQ(a, b):
    raise NotImplementedError("RSQ() not implemented yet")


class SKEW:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("SKEW() not implemented yet")

    @staticmethod
    def P(a, b):
        raise NotImplementedError("SKEW.P() not implemented yet")


def SLOPE(a, b):
    raise NotImplementedError("SLOPE() not implemented yet")


def SMALL(a, b):
    raise NotImplementedError("SMALL() not implemented yet")


def STANDARDIZE(a, b):
    raise NotImplementedError("STANDARDIZE() not implemented yet")


class STDEV:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("STDEV() not implemented yet")

    @staticmethod
    def P(a, b):
        raise NotImplementedError("STDEV.P() not implemented yet")

    @staticmethod
    def S(a, b):
        raise NotImplementedError("STDEV.S() not implemented yet")


def STDEVA(a, b):
    raise NotImplementedError("STDEVA() not implemented yet")


def STDEVP(a, b):
    raise NotImplementedError("STDEVP() not implemented yet")


def STDEVPA(a, b):
    raise NotImplementedError("STDEVPA() not implemented yet")


def STEYX(a, b):
    raise NotImplementedError("STEYX() not implemented yet")


class T:
    class DIST:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("T.DIST() not implemented yet")

        @staticmethod
        def _2T(a, b):
            raise NotImplementedError("T.DIST.2T() not implemented yet")

        @staticmethod
        def RT(a, b):
            raise NotImplementedError("T.DIST.RT() not implemented yet")

    class INV:
        @staticmethod
        def __call__(a, b):
            raise NotImplementedError("T.INV() not implemented yet")

        @staticmethod
        def _2T(a, b):
            raise NotImplementedError("T.INV.2T() not implemented yet")

    @staticmethod
    def TEST(a, b):
        raise NotImplementedError("T.TEST() not implemented yet")


def TDIST(a, b):
    raise NotImplementedError("TDIST() not implemented yet")


def TINV(a, b):
    raise NotImplementedError("TINV() not implemented yet")


def TRIMMEAN(a, b):
    raise NotImplementedError("TRIMMEAN() not implemented yet")


def TTEST(a, b):
    raise NotImplementedError("TTEST() not implemented yet")


class VAR:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("VAR() not implemented yet")

    @staticmethod
    def P(a, b):
        raise NotImplementedError("VAR.P() not implemented yet")

    @staticmethod
    def S(a, b):
        raise NotImplementedError("VAR.S() not implemented yet")


def VARA(a, b):
    raise NotImplementedError("VARA() not implemented yet")


def VARP(a, b):
    raise NotImplementedError("VARP() not implemented yet")


def VARPA(a, b):
    raise NotImplementedError("VARPA() not implemented yet")


class WEIBULL:
    @staticmethod
    def __call__(a, b):
        raise NotImplementedError("WEIBULL() not implemented yet")

    @staticmethod
    def DIST(a, b):
        raise NotImplementedError("WEIBULL.DIST() not implemented yet")


class Z:
    @staticmethod
    def TEST(a, b):
        raise NotImplementedError("Z.TEST() not implemented yet")


def ZTEST(a, b):
    raise NotImplementedError("ZTEST() not implemented yet")