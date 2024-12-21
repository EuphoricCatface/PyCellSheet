# Created by EuphoricCatface
import typing
import copy
import warnings


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


def coord_to_spreadsheet_ref(coord: tuple[int, int]) -> str:
    """Calculate a coordinate from spreadsheet-like address string"""
    row, col = coord

    if col == 0:
        col_str = "A"
    else:
        col_inter = []
        while col:
            col_inter.append(col % 26)
            col = col // 26
        col_inter = list(map(lambda a: chr(a + ord("A")), col_inter))
        col_inter.reverse()
        col_str = str.join("", col_inter)

    return col_str + str(row + 1)


class Range:
    def __init__(self, topleft: typing.Union[str, tuple[int, int]], width: int, lst: typing.Optional[list] = None):
        self.topleft = topleft
        self.lst = lst if lst else []
        self.width = width

        if len(self.lst) % width:
            warnings.warn("Length of the list is not divisible with the width")

    def flatten(self) -> list:
        return self.lst

    def __getitem__(self, item: int):
        if item >= len(self):
            raise IndexError("Index out of range")
        return copy.deepcopy(self.lst[self.width * item:self.width * (item + 1)])

    def __len__(self):
        dangling = False
        if len(self.lst) % self.width:
            warnings.warn("Length of the list is not divisible with the width")
            dangling = True
        return len(self.lst) // self.width + int(dangling)

    def append(self, item: typing.Any):
        self.lst.append(item)

    def __str__(self):
        topleft_str = coord_to_spreadsheet_ref(self.topleft)
        contents_list = [self.__getitem__(i) for i in range(len(self))]
        return topleft_str + str(contents_list)
