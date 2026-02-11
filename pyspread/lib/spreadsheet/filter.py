try:
    from pyspread.lib.pycellsheet import Range, flatten_args
except ImportError:
    from lib.pycellsheet import Range, flatten_args

_FILTER_FUNCTIONS = [
    'FILTER', 'SORT', 'SORTN', 'UNIQUE'
]
__all__ = _FILTER_FUNCTIONS + ["_FILTER_FUNCTIONS"]


def FILTER(range_, condition):
    lst = flatten_args(range_)
    cond = flatten_args(condition) if isinstance(condition, Range) else condition
    return [v for v, c in zip(lst, cond) if c]


def SORT(range_, sort_column=1, is_ascending=True, *args):
    raise NotImplementedError("SORT() not implemented yet")


def SORTN(range_, n=1, display_ties_mode=0, sort_column=1, is_ascending=True):
    raise NotImplementedError("SORTN() not implemented yet")


def UNIQUE(range_):
    lst = flatten_args(range_)
    seen = set()
    result = []
    for v in lst:
        if v not in seen:
            seen.add(v)
            result.append(v)
    return result
