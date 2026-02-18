# -*- coding: utf-8 -*-

# Copyright Seongyong Park (EuphCat)
# Distributed under the terms of the GNU General Public License

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

"""Sheet-name validation and normalization helpers."""

from typing import Iterable


def _is_printable_name(name: str) -> bool:
    """Return True if every character in `name` is printable."""

    return all(ch.isprintable() for ch in name)


def validate_sheet_name(name: str,
                        existing_names: Iterable[str],
                        current_name: str = None) -> tuple[bool, str]:
    """Validate a sheet name according to PyCellSheet stabilization rules."""

    if not isinstance(name, str):
        return False, "Sheet name must be a string."

    if name.strip() == "":
        return False, "Sheet name cannot be empty or whitespace-only."

    if not _is_printable_name(name):
        return False, "Sheet name cannot contain control characters."

    if name in existing_names and name != current_name:
        return False, f"Sheet '{name}' already exists."

    return True, ""


def generate_unique_sheet_name(preferred: str,
                               existing_names: Iterable[str],
                               fallback_index: int = 0) -> str:
    """Return a unique, valid sheet name derived from `preferred`."""

    if not isinstance(preferred, str):
        preferred = str(preferred)

    if preferred.strip() == "":
        base = f"Sheet {fallback_index}"
    else:
        # Remove control/non-printable chars while preserving user intent.
        base = "".join(ch for ch in preferred if ch.isprintable())
        if base.strip() == "":
            base = f"Sheet {fallback_index}"

    existing = set(existing_names)
    if base not in existing:
        return base

    suffix = 1
    while True:
        candidate = f"{base}_{suffix}"
        if candidate not in existing:
            return candidate
        suffix += 1


def sanitize_loaded_sheet_name(raw_name: str,
                               existing_names: Iterable[str],
                               fallback_index: int) -> str:
    """Sanitize a sheet name loaded from a file into a valid unique name."""

    return generate_unique_sheet_name(raw_name, existing_names, fallback_index=fallback_index)
