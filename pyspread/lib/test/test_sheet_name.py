# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License

# --------------------------------------------------------------------
# pyspread is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# pyspread is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with pyspread.  If not, see <http://www.gnu.org/licenses/>.
# --------------------------------------------------------------------

import pytest

from ..sheet_name import (
    generate_unique_sheet_name,
    sanitize_loaded_sheet_name,
    validate_sheet_name,
)


def test_validate_sheet_name_accepts_printable_nonempty():
    ok, reason = validate_sheet_name("Revenue 2026/Q1", existing_names=[])

    assert ok
    assert reason == ""


@pytest.mark.parametrize("name", ["", "   ", "\n", "\t"])
def test_validate_sheet_name_rejects_empty_or_whitespace(name):
    ok, reason = validate_sheet_name(name, existing_names=[])

    assert not ok
    assert "empty" in reason or "control" in reason


def test_validate_sheet_name_rejects_control_characters():
    ok, reason = validate_sheet_name("Bad\nName", existing_names=[])

    assert not ok
    assert "control" in reason


def test_validate_sheet_name_rejects_duplicates():
    ok, reason = validate_sheet_name("Sheet 1", existing_names=["Sheet 1", "Sheet 2"])

    assert not ok
    assert "already exists" in reason


def test_validate_sheet_name_allows_same_name_for_current_sheet():
    ok, reason = validate_sheet_name(
        "Sheet 1",
        existing_names=["Sheet 1", "Sheet 2"],
        current_name="Sheet 1",
    )

    assert ok
    assert reason == ""


def test_generate_unique_sheet_name_appends_suffix():
    name = generate_unique_sheet_name("Sheet 1", ["Sheet 1", "Sheet 1_1"])

    assert name == "Sheet 1_2"


def test_sanitize_loaded_sheet_name_removes_controls_and_uniquifies():
    existing = ["Clean"]
    name = sanitize_loaded_sheet_name("Bad\tName", existing, fallback_index=1)

    assert name == "BadName"
