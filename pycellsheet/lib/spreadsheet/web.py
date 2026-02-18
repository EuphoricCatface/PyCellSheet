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

import urllib.parse

_WEB_FUNCTIONS = [
    'ENCODEURL', 'HYPERLINK', 'IMPORTDATA', 'IMPORTFEED', 'IMPORTHTML', 'IMPORTRANGE', 'IMPORTXML',
    'ISURL'
]
__all__ = _WEB_FUNCTIONS + ["_WEB_FUNCTIONS"]


def ENCODEURL(text):
    return urllib.parse.quote(str(text), safe='')


def HYPERLINK(url, link_label=None):
    if link_label is None:
        return str(url)
    return str(link_label)


def IMPORTDATA(url):
    raise NotImplementedError("IMPORTDATA() not implemented yet")


def IMPORTFEED(url, query=None, headers=False, num_items=None):
    raise NotImplementedError("IMPORTFEED() not implemented yet")


def IMPORTHTML(url, query, index):
    raise NotImplementedError("IMPORTHTML() not implemented yet")


def IMPORTRANGE(spreadsheet_url, range_string):
    raise NotImplementedError("IMPORTRANGE() not implemented yet")


def IMPORTXML(url, xpath_query):
    raise NotImplementedError("IMPORTXML() not implemented yet")


def ISURL(value):
    if not isinstance(value, str):
        return False
    parsed = urllib.parse.urlparse(value)
    return bool(parsed.scheme and parsed.netloc)
