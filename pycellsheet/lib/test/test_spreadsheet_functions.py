# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License

import pytest

from ..spreadsheet import info as info_fn
from ..spreadsheet import math as math_fn
from ..spreadsheet import statistical as stat_fn


def _rng(width, values, topleft="A1"):
    return math_fn.Range(topleft, width, values)


def _stat_rng(width, values, topleft="A1"):
    return stat_fn.Range(topleft, width, values)


def test_base_supported_bases_and_padding():
    assert math_fn.BASE(10, 2) == "1010"
    assert math_fn.BASE(255, 16) == "FF"
    assert math_fn.BASE(7, 10, 3) == "007"


def test_base_unsupported_base_raises():
    with pytest.raises(NotImplementedError, match="Only base value"):
        math_fn.BASE(10, 3)


def test_countif_counts_matching_values():
    values = _rng(2, [1, 2, 3, math_fn.EmptyCell, 4, 5])
    assert math_fn.COUNTIF(values, lambda x: x >= 3) == 3


def test_countifs_applies_all_criteria():
    values = _rng(2, [1, 2, 3, 4])
    categories = _rng(2, ["x", "y", "x", "y"])
    result = math_fn.COUNTIFS((values, lambda x: x >= 2), (categories, lambda c: c == "y"))
    assert result == 2


def test_countifs_rejects_mismatched_lengths():
    a = _rng(2, [1, 2, 3, 4])
    b = _rng(3, [10, 20, 30])
    with pytest.raises(ValueError, match="ranges differ in length"):
        math_fn.COUNTIFS((a, lambda _: True), (b, lambda _: True))


def test_sum_accepts_scalars_lists_and_ranges():
    assert math_fn.SUM(1, [2, 3], _rng(2, [4, math_fn.EmptyCell, 5, 6])) == 21


def test_sumif_uses_criteria_and_skips_empty_sum_cells():
    criteria = _rng(2, [1, 2, 3, 4])
    sums = _rng(2, [10, math_fn.EmptyCell, 30, 40])
    assert math_fn.SUMIF(criteria, lambda x: x >= 2, sums) == 70


def test_sumif_dimension_mismatch_raises():
    criteria = _rng(2, [1, 2, 3, 4])
    sums = _rng(1, [10, 20, 30, 40])
    with pytest.raises(ValueError, match="dimensions"):
        math_fn.SUMIF(criteria, lambda _: True, sums)


def test_rounding_and_truncation_contracts():
    assert math_fn.ROUND(2.675, 2) == round(2.675, 2)
    assert math_fn.ROUNDDOWN(3.19, 1) == 3.1
    assert math_fn.ROUNDDOWN(-3.19, 1) == -3.1
    assert math_fn.ROUNDUP(3.11, 1) == 3.2
    assert math_fn.ROUNDUP(-3.11, 1) == -3.2
    assert math_fn.TRUNC(12.345, 2) == 12.34
    assert math_fn.TRUNC(-12.345, 2) == -12.34


def test_sumsq_squares_flattened_args():
    assert math_fn.SUMSQ(1, [2, 3], _rng(2, [4, math_fn.EmptyCell, 5, 6])) == 91


def test_error_type_maps_python_errors():
    assert info_fn.ERROR.TYPE(ValueError("bad")) == 8


def test_error_type_non_error_raises_na():
    with pytest.raises(info_fn.errors.SpreadsheetErrorNa):
        info_fn.ERROR.TYPE("ok")


def test_info_predicates_and_na():
    na_error = info_fn.NA()

    assert info_fn.ISEMAIL("person@example.com")
    assert not info_fn.ISEMAIL("not-an-email")
    assert info_fn.ISERR(ValueError("x"))
    assert info_fn.ISERROR(na_error)
    assert info_fn.TYPE("text") == 2


def test_n_passthrough_and_coercion():
    err = info_fn.errors.SpreadsheetErrorValue()
    assert info_fn.N(err) is err
    assert info_fn.N(4.5) == 4.5
    assert info_fn.N(True) == 1
    assert info_fn.N("abc") == 0


def test_avedev_and_averagea():
    assert stat_fn.AVEDEV([2, 4, 6]) == 4 / 3
    assert stat_fn.AVERAGEA([1, True, "x", 3.0]) == 1.25


def test_averageif_uses_aligned_criteria_and_average_range():
    criteria = _rng(2, [1, 2, 3, 4])
    average_values = _rng(2, [10, 20, math_fn.EmptyCell, 40])
    assert stat_fn.AVERAGEIF(criteria, lambda x: x % 2 == 0, average_values) == 30


def test_averageifs_intersects_multiple_criteria():
    avg = _rng(2, [10, 20, 30, 40])
    r1 = _rng(2, [1, 2, 3, 4])
    r2 = _rng(2, ["a", "b", "a", "b"])
    assert stat_fn.AVERAGEIFS(avg, (r1, lambda x: x >= 2), (r2, lambda x: x == "b")) == 30


def test_averageifs_rejects_mismatched_range_length():
    avg = _rng(2, [10, 20, 30, 40])
    mismatch = _rng(3, [1, 2, 3])
    with pytest.raises(ValueError, match="range length differs"):
        stat_fn.AVERAGEIFS(avg, (mismatch, lambda _: True))


def test_count_and_counta_behaviors():
    assert stat_fn.COUNT(1, 2.0, "x", True) == 3
    assert stat_fn.COUNTA(1, [2, stat_fn.EmptyCell], _stat_rng(2, [3, stat_fn.EmptyCell, 4, 5])) == 6


def test_geomean_requires_positive_values():
    assert stat_fn.GEOMEAN([1, 4, 16]) == pytest.approx(4.0)
    with pytest.raises(ValueError, match="all values > 0"):
        stat_fn.GEOMEAN([1, 0, 2])


def test_percentrank_exc_and_inc():
    data = [10, 20, 30, 40]
    assert stat_fn.PERCENTRANK.EXC(data, 25) == 0.667
    assert stat_fn.PERCENTRANK.EXC(data, 10) == 0.0
    assert stat_fn.PERCENTRANK.EXC(data, 40) == 1.0
    assert stat_fn.PERCENTRANK.INC(data, 25) == 0.333
    assert stat_fn.PERCENTRANK.INC(data, 5) == 0.0


def test_quartile_inc_and_exc():
    data = [1, 2, 3, 4]
    assert stat_fn.QUARTILE.INC(data, 0) == 1
    assert stat_fn.QUARTILE.INC(data, 2) == 2.5
    assert stat_fn.QUARTILE.INC(data, 4) == 4
    assert stat_fn.QUARTILE.EXC(data, 1) == 1.25
    with pytest.raises(ValueError, match="quart must be 1..3"):
        stat_fn.QUARTILE.EXC(data, 0)
