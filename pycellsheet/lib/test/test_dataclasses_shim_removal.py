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

"""Regression tests for removing the vendored dataclasses compatibility shim."""
from pathlib import Path


MODULE_FILES = {
    "pycellsheet.dialogs": Path("pycellsheet/dialogs.py"),
    "pycellsheet.installer": Path("pycellsheet/installer.py"),
    "pycellsheet.grid_renderer": Path("pycellsheet/grid_renderer.py"),
}


def test_no_legacy_dataclasses_fallback_imports():
    for path in MODULE_FILES.values():
        source = path.read_text(encoding="utf-8")
        assert "pycellsheet.lib.dataclasses" not in source
        assert "lib.dataclasses" not in source

def test_module_sources_compile():
    for path in MODULE_FILES.values():
        source = path.read_text(encoding="utf-8")
        compile(source, str(path), "exec")
