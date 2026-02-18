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

from datetime import date, datetime

_PARSER_FUNCTIONS = [
    'CONVERT', 'TO_DATE', 'TO_DOLLARS', 'TO_PERCENT', 'TO_PURE_NUMBER', 'TO_TEXT'
]
__all__ = _PARSER_FUNCTIONS + ["_PARSER_FUNCTIONS"]


def CONVERT(value, start_unit, end_unit):
    """Convert a value from one unit to another."""
    # Conversion table: unit -> (base_unit, multiplier_to_base)
    conversions = {
        # Weight/Mass
        'g': ('kg', 0.001),
        'kg': ('kg', 1),
        'lbm': ('kg', 0.45359237),
        'ozm': ('kg', 0.028349523125),
        'sg': ('kg', 14.593903),  # slug
        'u': ('kg', 1.6605402e-27),  # atomic mass unit
        'stone': ('kg', 6.35029318),
        'ton': ('kg', 907.18474),
        'uk_ton': ('kg', 1016.0469088),

        # Distance
        'm': ('m', 1),
        'mi': ('m', 1609.344),
        'Nmi': ('m', 1852),
        'in': ('m', 0.0254),
        'ft': ('m', 0.3048),
        'yd': ('m', 0.9144),
        'ang': ('m', 1e-10),
        'Picapt': ('m', 0.000352778),
        'pica': ('m', 0.00423333),
        'ell': ('m', 1.143),
        'parsec': ('m', 3.08567758149137e16),
        'lightyear': ('m', 9.46073047258e15),
        'survey_mi': ('m', 1609.347219),

        # Time
        'yr': ('sec', 31557600),
        'day': ('sec', 86400),
        'hr': ('sec', 3600),
        'mn': ('sec', 60),
        'sec': ('sec', 1),

        # Pressure
        'Pa': ('Pa', 1),
        'atm': ('Pa', 101325),
        'mmHg': ('Pa', 133.322),
        'psi': ('Pa', 6894.757),
        'Torr': ('Pa', 133.322),

        # Force
        'N': ('N', 1),
        'dyn': ('N', 1e-5),
        'lbf': ('N', 4.4482216152605),
        'pond': ('N', 0.00980665),

        # Energy
        'J': ('J', 1),
        'e': ('J', 1e-7),  # erg
        'c': ('J', 4.1868),  # thermochemical calorie
        'cal': ('J', 4.184),  # IT calorie
        'eV': ('J', 1.602176634e-19),
        'HPh': ('J', 2684519.537),
        'Wh': ('J', 3600),
        'flb': ('J', 1.3558179483314004),
        'BTU': ('J', 1055.05585262),

        # Power
        'W': ('W', 1),
        'HP': ('W', 745.699871582),
        'PS': ('W', 735.49875),

        # Magnetism
        'T': ('T', 1),
        'ga': ('T', 0.0001),

        # Temperature (requires special handling)
        'C': ('C', 1),
        'F': ('F', 1),
        'K': ('K', 1),

        # Liquid measures
        'tsp': ('l', 0.00492892),
        'tbs': ('l', 0.0147868),
        'oz': ('l', 0.0295735),
        'cup': ('l', 0.24),
        'pt': ('l', 0.473176),
        'us_pt': ('l', 0.473176),
        'uk_pt': ('l', 0.568261),
        'qt': ('l', 0.946353),
        'uk_qt': ('l', 1.1365225),
        'gal': ('l', 3.785411784),
        'uk_gal': ('l', 4.54609),
        'l': ('l', 1),
        'lt': ('l', 1),

        # Area
        'uk_acre': ('m2', 4046.8564224),
        'us_acre': ('m2', 4046.8564224),
        'ha': ('m2', 10000),
        'Morgen': ('m2', 2500),

        # Information
        'bit': ('bit', 1),
        'byte': ('bit', 8),
    }

    # Handle temperature conversions separately
    if start_unit in ('C', 'F', 'K') or end_unit in ('C', 'F', 'K'):
        # Convert to Celsius first
        if start_unit == 'C':
            celsius = value
        elif start_unit == 'F':
            celsius = (value - 32) * 5/9
        elif start_unit == 'K':
            celsius = value - 273.15
        else:
            raise ValueError(f"Unknown temperature unit: {start_unit}")

        # Convert from Celsius to target
        if end_unit == 'C':
            return celsius
        elif end_unit == 'F':
            return celsius * 9/5 + 32
        elif end_unit == 'K':
            return celsius + 273.15
        else:
            raise ValueError(f"Unknown temperature unit: {end_unit}")

    # Check if units exist
    if start_unit not in conversions:
        raise ValueError(f"Unknown unit: {start_unit}")
    if end_unit not in conversions:
        raise ValueError(f"Unknown unit: {end_unit}")

    # Get base units
    start_base, start_mult = conversions[start_unit]
    end_base, end_mult = conversions[end_unit]

    # Check if units are compatible
    if start_base != end_base:
        raise ValueError(f"Cannot convert {start_unit} to {end_unit} - incompatible unit types")

    # Convert
    return value * start_mult / end_mult


def TO_DATE(value):
    if isinstance(value, (date, datetime)):
        return value
    if isinstance(value, (int, float)):
        return date.fromordinal(int(value) + 693594)
    if isinstance(value, str):
        try:
            import dateutil.parser
        except ImportError:
            raise NotImplementedError("Install `dateutil` python package to use TO_DATE")
        return dateutil.parser.parse(value).date()
    raise ValueError(f"Cannot convert {type(value)} to date")


def TO_DOLLARS(value):
    return f"${float(value):,.2f}"


def TO_PERCENT(value):
    return f"{float(value) * 100:.2f}%"


def TO_PURE_NUMBER(value):
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace('$', '').replace(',', '').replace('%', '')
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        raise ValueError(f"Cannot convert '{value}' to number")


def TO_TEXT(value):
    return str(value)
