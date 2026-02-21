# -*- coding: utf-8 -*-
#
# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License
#
# --------------------------------------------------------------------
# pycellsheet is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# pycellsheet is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with pycellsheet.  If not, see <http://www.gnu.org/licenses/>.
# --------------------------------------------------------------------

import pytest

from ..expression_parser_migrator import (
    KNOWN_MODE_IDS,
    SAFE_CHANGED,
    UNCHANGED,
    RISKY_SKIPPED,
    INVALID_SOURCE_ASSUMPTION,
    _migrate_cell,
    preview_migration,
    apply_migration,
)


class _DummyCodeArray:
    def __init__(self, grid):
        self.dict_grid = dict(grid)

    def __call__(self, key):
        return self.dict_grid[key]

    def __setitem__(self, key, value):
        self.dict_grid[key] = value


def test_preview_mixed_to_pure_spreadsheet_marks_safe_rewrites():
    code_array = _DummyCodeArray({
        (0, 0, 0): "1 + 2",
        (0, 1, 0): "'hello",
        (0, 2, 0): "=SUM(A1:A2)",
    })

    report = preview_migration(code_array, "mixed", "pure_spreadsheet")
    classes = [entry.classification for entry in report.entries]
    values = [entry.new_text for entry in report.entries]

    assert classes == [SAFE_CHANGED, "unchanged", "unchanged"]
    assert values == [">1 + 2", "'hello", "=SUM(A1:A2)"]


def test_preview_pure_spreadsheet_to_mixed_skips_ambiguous_literals():
    code_array = _DummyCodeArray({(0, 0, 0): "plain text"})
    report = preview_migration(code_array, "pure_spreadsheet", "mixed")

    assert report.entries[0].classification == RISKY_SKIPPED


def test_apply_migration_updates_only_safe_changes():
    code_array = _DummyCodeArray({
        (0, 0, 0): "1 + 2",
        (0, 1, 0): "plain text",
    })

    report = apply_migration(code_array, "mixed", "pure_spreadsheet")

    assert report.summary[SAFE_CHANGED] == 2
    assert code_array.dict_grid[(0, 0, 0)] == ">1 + 2"
    assert code_array.dict_grid[(0, 1, 0)] == ">plain text"


def test_preview_migration_validates_mode_ids():
    code_array = _DummyCodeArray({})

    with pytest.raises(ValueError, match="Unknown source parser mode id"):
        preview_migration(code_array, "unknown", "mixed")
    with pytest.raises(ValueError, match="Unknown target parser mode id"):
        preview_migration(code_array, "mixed", "unknown")


def test_preview_migration_filters_tables():
    code_array = _DummyCodeArray({
        (0, 0, 0): "1+2",
        (0, 1, 1): "3+4",
    })

    report = preview_migration(code_array, "mixed", "pure_spreadsheet", tables=[1])

    assert [entry.key for entry in report.entries] == [(0, 1, 1)]
    assert report.entries[0].new_text == ">3+4"


def test_apply_migration_does_not_write_unchanged_entries():
    code_array = _DummyCodeArray({(0, 0, 0): "'literal"})
    report = apply_migration(code_array, "mixed", "pure_spreadsheet")

    assert report.summary[UNCHANGED] == 1
    assert code_array.dict_grid[(0, 0, 0)] == "'literal"


def test_migrate_cell_include_risky_for_unsupported_pair():
    _, cls_risky, _ = _migrate_cell("x", "unknown_source", "unknown_target", include_risky=False)
    _, cls_forced, reason = _migrate_cell("x", "unknown_source", "unknown_target", include_risky=True)

    assert cls_risky == RISKY_SKIPPED
    assert cls_forced == SAFE_CHANGED
    assert "include_risky=True" in reason


def test_migration_report_summary_counts_include_all_classifications():
    code_array = _DummyCodeArray({(0, 0, 0): "x"})
    report = preview_migration(code_array, "pure_spreadsheet", "mixed")
    counts = report.summary

    assert set([SAFE_CHANGED, UNCHANGED, RISKY_SKIPPED, INVALID_SOURCE_ASSUMPTION, "total"]).issubset(counts)
    assert counts["total"] == 1
