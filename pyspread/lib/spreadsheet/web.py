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
