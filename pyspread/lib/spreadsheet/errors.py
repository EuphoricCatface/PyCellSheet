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

class SpreadsheetErrorBase(Exception):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "BASE?!"

    def cell_output(self):
        return f"#{self.__string}"


class SpreadsheetErrorNull(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NULL!"


class SpreadsheetErrorDivZero(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "DIV/0!"


class SpreadsheetErrorValue(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "VALUE!"


class SpreadsheetErrorRef(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "REF!"


class SpreadsheetErrorName(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NAME?"


class SpreadsheetErrorNum(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NUM!"


class SpreadsheetErrorNa(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "N/A"
