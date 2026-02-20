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

from ..expression_parser_migrator import (
    SAFE_CHANGED,
    RISKY_SKIPPED,
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
