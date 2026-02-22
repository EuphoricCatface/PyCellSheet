# -*- coding: utf-8 -*-

# Copyright Martin Manns
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


"""
test_cli
========

Unit tests for cli.py

"""

from argparse import Namespace
from contextlib import contextmanager
from os.path import abspath, dirname, join
import sys
from unittest.mock import patch
from pathlib import Path, PosixPath
import types
from importlib.util import spec_from_file_location, module_from_spec

import pytest

PYCELLSHEETPATH = abspath(join(dirname(__file__) + "/.."))
LIBPATH = abspath(PYCELLSHEETPATH + "/lib")


@contextmanager
def insert_path(path):
    sys.path.insert(0, path)
    yield
    sys.path.pop(0)


with insert_path(PYCELLSHEETPATH):
    from .. import cli
    from ..cli import PyCellSheetArgumentParser

param_test_cli = [
    (['pycellsheet'],
     Namespace(file=None, default_settings=False, loglevel=30)),
    (['pycellsheet', 'test.pys'],
     Namespace(file=PosixPath("test.pys"), default_settings=False,
               loglevel=30)),
    (['pycellsheet', '--help'],
     None),
    (['pycellsheet', '--version'],
     None),
    (['pycellsheet', '--default-settings'],
     Namespace(file=None, default_settings=True, loglevel=30)),
    (['pycellsheet', '-d'],
     Namespace(file=None, default_settings=False, loglevel=10)),
    (['pycellsheet', '-v'],
     Namespace(file=None, default_settings=False, loglevel=20)),
]


@pytest.mark.parametrize("argv, res", param_test_cli)
def test_cli(argv, res):
    with patch('argparse._sys.argv', argv):
        parser = PyCellSheetArgumentParser()
        if res is None:
            with pytest.raises(SystemExit) as exc:
                args, unknown = parser.parse_known_args()
            assert exc.value.code == 0
        else:
            args, unknown = parser.parse_known_args()
            assert args == res


def test_check_mandatory_dependencies_warns_for_missing_and_old_modules(monkeypatch):
    class _Dep:
        def __init__(self, name, installed, version, required):
            self.name = name
            self._installed = installed
            self.version = version
            self.required_version = required

        def is_installed(self):
            return self._installed

    monkeypatch.setattr(cli, "pyqtsvg", None)
    monkeypatch.setattr(
        cli,
        "REQUIRED_DEPENDENCIES",
        [
            _Dep("okmod", True, "1.0", "1.0"),
            _Dep("missingmod", False, "0", "1.0"),
            _Dep("oldmod", True, "0.1", "1.0"),
        ],
    )
    monkeypatch.setattr(cli, "sys", types.SimpleNamespace(
        version_info=types.SimpleNamespace(major=3, minor=5, micro=9),
        stdout=sys.stdout,
    ))

    with patch("sys.stdout.write") as write_mock:
        cli.check_mandatory_dependencies()

    written = "".join(call.args[0] for call in write_mock.call_args_list)
    assert "Python has version 3.5.9" in written
    assert "Required module missingmod not found." in written
    assert "Module oldmod has version 0.1but 1.0 is required." in written
    assert "Required module PyQt6.QtSvg not found." in written


def test_main_module_invokes_main(monkeypatch):
    called = {"count": 0}

    fake_entry = types.ModuleType("pycellsheet.pycellsheet")

    def _fake_main():
        called["count"] += 1

    fake_entry.main = _fake_main
    fake_pkg = types.ModuleType("pycellsheet")
    fake_pkg.main = _fake_main
    monkeypatch.setitem(sys.modules, "pycellsheet", fake_pkg)
    monkeypatch.setitem(sys.modules, "pycellsheet.pycellsheet", fake_entry)
    monkeypatch.setitem(sys.modules, "pycellsheet.pycellsheet.pycellsheet", fake_entry)

    main_path = Path(PYCELLSHEETPATH) / "__main__.py"
    spec = spec_from_file_location("test_pycellsheet_main", main_path)
    module = module_from_spec(spec)
    spec.loader.exec_module(module)

    assert called["count"] == 1
