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

"""Conservative migration helpers for Expression Parser modes."""

from dataclasses import dataclass
from typing import Optional


SAFE_CHANGED = "safe_changed"
UNCHANGED = "unchanged"
RISKY_SKIPPED = "risky_skipped"
INVALID_SOURCE_ASSUMPTION = "invalid_source_assumption"

KNOWN_MODE_IDS = {
    "pure_pythonic",
    "mixed",
    "pure_spreadsheet",
}


@dataclass(frozen=True)
class MigrationEntry:
    key: tuple[int, int, int]
    old_text: str
    new_text: str
    classification: str
    reason: str


@dataclass
class MigrationReport:
    source_mode_id: str
    target_mode_id: str
    entries: list[MigrationEntry]

    @property
    def summary(self) -> dict[str, int]:
        counts = {
            SAFE_CHANGED: 0,
            UNCHANGED: 0,
            RISKY_SKIPPED: 0,
            INVALID_SOURCE_ASSUMPTION: 0,
        }
        for entry in self.entries:
            counts[entry.classification] += 1
        counts["total"] = len(self.entries)
        return counts


def _looks_number_literal(cell: str) -> bool:
    if not cell:
        return False
    try:
        int(cell)
        return True
    except ValueError:
        pass
    try:
        float(cell)
        return True
    except ValueError:
        return False


def _quote_as_python_string(value: str) -> str:
    return repr(value)


def _migrate_cell(cell: str, source_mode_id: str, target_mode_id: str,
                  include_risky: bool) -> tuple[str, str, str]:
    if source_mode_id == target_mode_id:
        return cell, UNCHANGED, "Source and target modes are identical."

    # Mixed -> Pure Spreadsheet
    if source_mode_id == "mixed" and target_mode_id == "pure_spreadsheet":
        if cell.startswith("'"):
            return cell, UNCHANGED, "String-literal marker is compatible."
        if cell.startswith("="):
            return cell, UNCHANGED, "Spreadsheet marker is compatible."
        return f">{cell}", SAFE_CHANGED, "Python code marker moved to '>'."

    # Pure Spreadsheet -> Mixed
    if source_mode_id == "pure_spreadsheet" and target_mode_id == "mixed":
        if cell.startswith(">"):
            payload = cell[1:]
            if payload.startswith("'"):
                return cell, RISKY_SKIPPED, "Removing '>' would trigger mixed string marker."
            return payload, SAFE_CHANGED, "Removed pure-spreadsheet Python marker."
        if cell.startswith("="):
            return cell, UNCHANGED, "Spreadsheet marker is compatible."
        return cell, RISKY_SKIPPED, "Ambiguous literal/code intent; requires manual review."

    # Pure Pythonic -> Mixed
    if source_mode_id == "pure_pythonic" and target_mode_id == "mixed":
        if cell.startswith("'"):
            return _quote_as_python_string(cell), SAFE_CHANGED, "Escaped leading quote for mixed mode."
        if cell.startswith("="):
            return cell, RISKY_SKIPPED, "Leading '=' changes meaning in mixed mode."
        return cell, UNCHANGED, "Python code stays Python code."

    # Mixed -> Pure Pythonic
    if source_mode_id == "mixed" and target_mode_id == "pure_pythonic":
        if cell.startswith("'"):
            return _quote_as_python_string(cell[1:]), SAFE_CHANGED, "Converted mixed literal marker to Python string literal."
        if cell.startswith("="):
            return cell, RISKY_SKIPPED, "Spreadsheet code has no pure-pythonic equivalent."
        return cell, UNCHANGED, "Python code stays Python code."

    # Pure Spreadsheet -> Pure Pythonic
    if source_mode_id == "pure_spreadsheet" and target_mode_id == "pure_pythonic":
        if cell.startswith(">"):
            return cell[1:], SAFE_CHANGED, "Converted pure-spreadsheet Python marker to plain Python code."
        if cell.startswith("="):
            return cell, RISKY_SKIPPED, "Spreadsheet code has no pure-pythonic equivalent."
        if cell.startswith("'"):
            return _quote_as_python_string(cell[1:]), SAFE_CHANGED, "Converted literal marker to Python string literal."
        if _looks_number_literal(cell):
            return cell, UNCHANGED, "Numeric literal remains valid Python."
        return _quote_as_python_string(cell), SAFE_CHANGED, "Converted plain spreadsheet literal to Python string literal."

    # Pure Pythonic -> Pure Spreadsheet
    if source_mode_id == "pure_pythonic" and target_mode_id == "pure_spreadsheet":
        if cell.startswith("="):
            return cell, RISKY_SKIPPED, "Leading '=' has spreadsheet-code meaning in target mode."
        return f">{cell}", SAFE_CHANGED, "Moved Python code behind pure-spreadsheet marker."

    if include_risky:
        return cell, SAFE_CHANGED, "Risky migration forced by include_risky=True."
    return cell, RISKY_SKIPPED, "Unsupported migration pair for conservative mode."


def _iter_candidate_cells(code_array, tables: Optional[list[int]]):
    allowed_tables = None if tables is None else set(tables)
    for key in sorted(code_array.keys()):
        row, col, table = key
        if allowed_tables is not None and table not in allowed_tables:
            continue
        value = code_array(key)
        if isinstance(value, str):
            yield key, value


def preview_migration(code_array, source_mode_id: str, target_mode_id: str,
                      tables: Optional[list[int]] = None,
                      include_risky: bool = False) -> MigrationReport:
    if source_mode_id not in KNOWN_MODE_IDS:
        raise ValueError(f"Unknown source parser mode id: {source_mode_id}")
    if target_mode_id not in KNOWN_MODE_IDS:
        raise ValueError(f"Unknown target parser mode id: {target_mode_id}")

    entries = []
    for key, old_text in _iter_candidate_cells(code_array, tables):
        new_text, classification, reason = _migrate_cell(
            old_text, source_mode_id, target_mode_id, include_risky
        )
        entries.append(MigrationEntry(
            key=key,
            old_text=old_text,
            new_text=new_text,
            classification=classification,
            reason=reason,
        ))
    return MigrationReport(source_mode_id=source_mode_id, target_mode_id=target_mode_id, entries=entries)


def apply_migration(code_array, source_mode_id: str, target_mode_id: str,
                    tables: Optional[list[int]] = None,
                    include_risky: bool = False) -> MigrationReport:
    report = preview_migration(code_array, source_mode_id, target_mode_id, tables, include_risky)
    for entry in report.entries:
        if entry.classification != SAFE_CHANGED:
            continue
        if entry.old_text == entry.new_text:
            continue
        code_array[entry.key] = entry.new_text
    return report
