# Created by EuphoricCatface
from typing import Any


class Empty:
    def __str__(self):
        return ""

    def __int__(self):
        return 0

    def __float__(self):
        return 0.0

EmptyCell = Empty()


class PythonCode(str):
    pass


class HelpText:
    def __init__(self, query, contents):
        if len(query) == 0:
            self.query = "help()"
        elif len(query) == 1:
            self.query = f"help({query[0]})"
        else:
            self.query = f"help{query}"
        self.contents = contents


class Range(list):
    def __init__(self):
        super().__init__()
        self.top_left = None

    def flatten(self) -> list:
        self: list[list[Any]]
        return [elem for inner_lst in self for elem in inner_lst]
