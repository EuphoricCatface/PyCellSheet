# Created by EuphoricCatface
from typing import Any


EmptyCell = object()


class PythonCode(str):
    pass


class Range(list):
    def __init__(self):
        super().__init__()
        self.top_left = None

    def flatten(self) -> list:
        self: list[list[Any]]
        return [elem for inner_lst in self for elem in inner_lst]
