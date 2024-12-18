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


class Range:
    def __init__(self, topleft: str, lst: list, width: int):
        self.topleft = topleft
        self.lst = lst
        self.width = width

        if len(self.lst) % width:
            raise ValueError("Length of the list is not divisible with the width")

    def flatten(self) -> list:
        return self.lst

    def __getitem__(self, item: int):
        if item >= len(self.lst) / self.width:
            raise IndexError("Index out of range")
        return self.lst[self.width * item:self.width * (item + 1)]

