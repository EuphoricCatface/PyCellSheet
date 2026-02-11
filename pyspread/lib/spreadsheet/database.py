import statistics

_DATABASE_FUNCTIONS = [
    'DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP',
    'DSUM', 'DVAR', 'DVARP'
]
__all__ = _DATABASE_FUNCTIONS + ["_DATABASE_FUNCTIONS"]


def _extract_matching_values(database, field, criteria):
    """
    Helper function to extract values from database that match criteria.
    Simplified implementation: assumes database and criteria are lists of dicts.
    """
    # Simplified: assume database is a list of rows (dicts)
    # and criteria is a dict of field: value pairs
    # In a real implementation, this would parse Range objects

    # For now, return empty list as a placeholder
    # Real implementation would need to parse the database structure
    return []


def DAVERAGE(database, field, criteria):
    """Average of selected database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    if not values:
        return 0
    return sum(values) / len(values)


def DCOUNT(database, field, criteria):
    """Count of numeric entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    return len([v for v in values if isinstance(v, (int, float))])


def DCOUNTA(database, field, criteria):
    """Count of non-empty entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    return len([v for v in values if v is not None and v != ""])


def DGET(database, field, criteria):
    """Single value from database matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    if len(values) == 0:
        raise ValueError("No records match criteria")
    if len(values) > 1:
        raise ValueError("Multiple records match criteria")
    return values[0]


def DMAX(database, field, criteria):
    """Maximum value from database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return 0
    return max(numeric)


def DMIN(database, field, criteria):
    """Minimum value from database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return 0
    return min(numeric)


def DPRODUCT(database, field, criteria):
    """Product of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return 0
    result = 1
    for v in numeric:
        result *= v
    return result


def DSTDEV(database, field, criteria):
    """Sample standard deviation of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if len(numeric) < 2:
        return 0
    return statistics.stdev(numeric)


def DSTDEVP(database, field, criteria):
    """Population standard deviation of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return 0
    return statistics.pstdev(numeric)


def DSUM(database, field, criteria):
    """Sum of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    return sum(numeric)


def DVAR(database, field, criteria):
    """Sample variance of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if len(numeric) < 2:
        return 0
    return statistics.variance(numeric)


def DVARP(database, field, criteria):
    """Population variance of database entries matching criteria."""
    values = _extract_matching_values(database, field, criteria)
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return 0
    return statistics.pvariance(numeric)
