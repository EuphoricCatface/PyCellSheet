# -*- coding: utf-8 -*-

# Copyright Martin Manns
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

"""

**Provides**

* :func:`check_mandatory_dependencies`:
* :class:`PathAction`:
* :class:`CommandLineParser`:

"""

from argparse import Action, ArgumentParser
from pathlib import Path
import sys

try:
    import PyQt5.QtSvg as pyqtsvg
except ImportError:
    pyqtsvg = None

try:
    from pyspread.__init__ import APP_NAME, VERSION
    from pyspread.installer import REQUIRED_DEPENDENCIES
except ImportError:
    from __init__ import APP_NAME, VERSION
    from installer import REQUIRED_DEPENDENCIES


def check_mandatory_dependencies():
    """Checks mandatory dependencies and exits if they are not met"""

    def dependency_warning(message: str):
        """Print warning message to stdout

        :param message: Warning message to be displayed

        """

        sys.stdout.write('Warning: {}\n'.format(message))

    # Check Python version
    major = sys.version_info.major
    minor = sys.version_info.minor
    micro = sys.version_info.micro
    if major < 3 or major == 3 and minor < 6:
        msg_tpl = "Python has version {}.{}.{} but ≥ 3.6 is required."
        msg = msg_tpl.format(major, minor, micro)
        dependency_warning(msg)

    for module in REQUIRED_DEPENDENCIES:
        if module.is_installed() is None or not module.is_installed():
            msg_tpl = "Required module {} not found."
            msg = msg_tpl.format(module.name)
            dependency_warning(msg)
        elif module.version < module.required_version:
            msg_tpl = "Module {} has version {} but {} is required."
            msg = msg_tpl.format(module.name, module.version,
                                 module.required_version)
            dependency_warning(msg)
    if pyqtsvg is None:
        # Import of mandatory module failed
        msg = "Required module PyQt5.QtSvg not found."
        dependency_warning(msg)


class PyspreadArgumentParser(ArgumentParser):
    """Parser for the command line"""

    def __init__(self):
        check_mandatory_dependencies()

        description = "pyspread is a non-traditional spreadsheet that is " \
                      "based on and written in the programming language " \
                      "Python."

        super().__init__(prog=APP_NAME, description=description)

        self.add_argument('--version', action='version', version=VERSION)

        self.add_argument('--default-settings', action='store_true',
                          help='start with default settings and save them on '
                               'exit')

        self.add_argument('file', type=Path, nargs='?', default=None,
                          help='open pyspread file in pys or pysu format')
