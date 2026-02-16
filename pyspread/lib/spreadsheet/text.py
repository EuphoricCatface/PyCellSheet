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

import re

try:
    from pyspread.lib.pycellsheet import EmptyCell
except ImportError:
    from lib.pycellsheet import EmptyCell

_TEXT_FUNCTIONS = [
    'ARABIC', 'ASC', 'CHAR', 'CLEAN', 'CODE', 'CONCATENATE', 'DOLLAR', 'EXACT', 'FIND', 'FINDB',
    'FIXED', 'JOIN', 'LEFT', 'LEFTB', 'LEN', 'LENB', 'LOWER', 'MID', 'MIDB', 'PROPER',
    'REGEXEXTRACT', 'REGEXMATCH', 'REGEXREPLACE', 'REPLACE', 'REPLACEB', 'REPT', 'RIGHT',
    'RIGHTB', 'ROMAN', 'SEARCH', 'SEARCHB', 'SPLIT', 'SUBSTITUTE', 'T_TEXT', 'TEXT', 'TEXTJOIN',
    'TRIM', 'UNICHAR', 'UNICODE', 'UPPER', 'VALUE'
]
__all__ = _TEXT_FUNCTIONS + ["_TEXT_FUNCTIONS"]


_ROMAN_VALUES = {
    'I': 1, 'V': 5, 'X': 10, 'L': 50,
    'C': 100, 'D': 500, 'M': 1000
}


def ARABIC(roman_numeral):
    text = str(roman_numeral).upper().strip()
    if not text:
        return 0
    result = 0
    prev = 0
    for ch in reversed(text):
        val = _ROMAN_VALUES.get(ch)
        if val is None:
            raise ValueError(f"Invalid Roman numeral character: {ch}")
        if val < prev:
            result -= val
        else:
            result += val
        prev = val
    return result


def ASC(text):
    result = []
    for ch in str(text):
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            result.append(' ')
        else:
            result.append(ch)
    return ''.join(result)


def CHAR(number):
    return chr(int(number))


def CLEAN(text):
    return ''.join(ch for ch in str(text) if ch.isprintable())


def CODE(text):
    s = str(text)
    if not s:
        raise ValueError("Empty string")
    return ord(s[0])


def CONCATENATE(*args):
    return ''.join(str(a) for a in args)


def DOLLAR(number, decimals=2):
    if decimals >= 0:
        formatted = f"{number:,.{decimals}f}"
    else:
        rounded = round(number, decimals)
        formatted = f"{rounded:,.0f}"
    return f"${formatted}"


def EXACT(string1, string2):
    return str(string1) == str(string2)


def FIND(find_text, within_text, start_position=1):
    s = str(within_text)
    f = str(find_text)
    pos = s.find(f, start_position - 1)
    if pos == -1:
        raise ValueError(f"'{find_text}' not found in '{within_text}'")
    return pos + 1


def FINDB(find_text, within_text, start_position=1):
    return FIND(find_text, within_text, start_position)


def FIXED(number, decimals=2, no_commas=False):
    if decimals >= 0:
        formatted = f"{number:.{decimals}f}"
    else:
        rounded = round(number, decimals)
        formatted = f"{rounded:.0f}"
    if not no_commas:
        parts = formatted.split('.')
        int_part = parts[0]
        sign = ''
        if int_part.startswith('-'):
            sign = '-'
            int_part = int_part[1:]
        int_part = f"{int(int_part):,}"
        parts[0] = sign + int_part
        formatted = '.'.join(parts)
    return formatted


def JOIN(delimiter, *args):
    parts = []
    for arg in args:
        if isinstance(arg, (list, tuple)):
            parts.extend(str(a) for a in arg)
        else:
            parts.append(str(arg))
    return str(delimiter).join(parts)


def LEFT(text, num_chars=1):
    return str(text)[:int(num_chars)]


def LEFTB(text, num_bytes=1):
    return LEFT(text, num_bytes)


def LEN(text):
    return len(str(text))


def LENB(text):
    return len(str(text).encode('utf-8'))


def LOWER(text):
    return str(text).lower()


def MID(text, start_position, num_chars):
    s = str(text)
    start = int(start_position) - 1
    length = int(num_chars)
    return s[start:start + length]


def MIDB(text, start_position, num_bytes):
    return MID(text, start_position, num_bytes)


def PROPER(text):
    return str(text).title()


def REGEXEXTRACT(text, regular_expression):
    match = re.search(regular_expression, str(text))
    if match is None:
        raise ValueError(f"No match found for pattern '{regular_expression}'")
    if match.groups():
        return match.group(1)
    return match.group(0)


def REGEXMATCH(text, regular_expression):
    return bool(re.search(regular_expression, str(text)))


def REGEXREPLACE(text, regular_expression, replacement):
    return re.sub(regular_expression, replacement, str(text))


def REPLACE(text, position, length, new_text):
    s = str(text)
    pos = int(position) - 1
    ln = int(length)
    return s[:pos] + str(new_text) + s[pos + ln:]


def REPLACEB(text, position, num_bytes, new_text):
    return REPLACE(text, position, num_bytes, new_text)


def REPT(text, number_times):
    return str(text) * int(number_times)


def RIGHT(text, num_chars=1):
    s = str(text)
    n = int(num_chars)
    if n == 0:
        return ''
    return s[-n:]


def RIGHTB(text, num_bytes=1):
    return RIGHT(text, num_bytes)


def ROMAN(number):
    number = int(number)
    if number <= 0 or number >= 4000:
        raise ValueError("Number must be between 1 and 3999")
    values = [
        (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
        (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
        (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
    ]
    result = []
    for val, numeral in values:
        while number >= val:
            result.append(numeral)
            number -= val
    return ''.join(result)


def SEARCH(find_text, within_text, start_position=1):
    s = str(within_text).lower()
    f = str(find_text).lower()
    pos = s.find(f, start_position - 1)
    if pos == -1:
        raise ValueError(f"'{find_text}' not found in '{within_text}'")
    return pos + 1


def SEARCHB(find_text, within_text, start_position=1):
    return SEARCH(find_text, within_text, start_position)


def SPLIT(text, delimiter, split_by_each=True, remove_empty=True):
    s = str(text)
    if split_by_each:
        pattern = '[' + re.escape(str(delimiter)) + ']'
        parts = re.split(pattern, s)
    else:
        parts = s.split(str(delimiter))
    if remove_empty:
        parts = [p for p in parts if p]
    return parts


def SUBSTITUTE(text, old_text, new_text, instance_num=None):
    s = str(text)
    old = str(old_text)
    new = str(new_text)
    if instance_num is None:
        return s.replace(old, new)
    count = 0
    start = 0
    while True:
        pos = s.find(old, start)
        if pos == -1:
            break
        count += 1
        if count == instance_num:
            return s[:pos] + new + s[pos + len(old):]
        start = pos + 1
    return s


def T_TEXT(value):
    if isinstance(value, str):
        return value
    return ""


def TEXT(number, format_string):
    fmt = str(format_string)
    if fmt == '0':
        return f"{number:.0f}"
    if fmt == '0.00':
        return f"{number:.2f}"
    if fmt == '#,##0':
        return f"{number:,.0f}"
    if fmt == '#,##0.00':
        return f"{number:,.2f}"
    if fmt == '0%':
        return f"{number * 100:.0f}%"
    if fmt == '0.00%':
        return f"{number * 100:.2f}%"
    try:
        return format(number, fmt)
    except (ValueError, TypeError):
        return str(number)


def TEXTJOIN(delimiter, ignore_empty, *args):
    parts = []
    for arg in args:
        if isinstance(arg, (list, tuple)):
            for item in arg:
                if ignore_empty and (item == EmptyCell or item == '' or item is None):
                    continue
                parts.append(str(item))
        else:
            if ignore_empty and (arg == EmptyCell or arg == '' or arg is None):
                continue
            parts.append(str(arg))
    return str(delimiter).join(parts)


def TRIM(text):
    return ' '.join(str(text).split())


def UNICHAR(number):
    return chr(int(number))


def UNICODE(text):
    s = str(text)
    if not s:
        raise ValueError("Empty string")
    return ord(s[0])


def UPPER(text):
    return str(text).upper()


def VALUE(text):
    s = str(text).strip()
    has_percent = '%' in s
    s = s.replace('$', '').replace(',', '').replace('%', '')
    try:
        if '.' in s:
            result = float(s)
        else:
            result = int(s)
    except ValueError:
        raise ValueError(f"Cannot convert '{text}' to a number")
    if has_percent:
        result = result / 100
    return result
