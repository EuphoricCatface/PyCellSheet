# Created by EuphoricCatface
import typing
import copy
import warnings
import re


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


class ExpressionParser:
    DEFAULT_PARSERS = {
        "Pure Pythonic": (
            "return PythonCode(cell)\n"
        ),
        "Mixed": (
            "# Inspired by the spreadsheet string marker\n"
            "if cell.startswith('\\''):\n"
            "    return cell[1:]\n"
            "return PythonCode(cell)\n"
        ),
        "Reverse Mixed": (
            "# Inspired by the python shell prompt `>>>`\n"
            "if cell.startswith(">"):\n"
            "    return PythonCode(cell[1:])\n"
            "if cell.startswith('\\''):\n"
            "    cell = cell[1:]\n"
            "return cell\n"
        ),
        "Pure Spreadsheet": (
            "if cell.startswith('='):\n"
            "    return PythonCode(cell[1:])\n"
            "try:\n"
            "    return int(cell)\n"
            "except ValueError:\n"
            "    pass\n"
            "try:\n"
            "    return float(cell)\n"
            "except ValueError:\n"
            "    pass\n"
            "if cell.startswith('\\''):\n"
            "    cell = cell[1:]\n"
            "return cell\n"
        )
    }

    def __init__(self):
        self.cached_fn = None

    def set_parser(self, code: str):
        local = {}
        code_list = ["def parser(cell):"]
        code_list.extend(map(lambda a: "    " + a, code.splitlines(keepends=False)))
        code = str.join("\n", code_list)
        exec(code, globals(), local)
        self.cached_fn = local["parser"]

    def parse(self, cell):
        return self.cached_fn(cell)

    def handle_empty(self, cell):
        """Returns true if cell is empty"""
        if cell is None or cell == "":
            return True
        return False


class ReferenceParser:
    COMPILED_RANGE_RE = re.compile(r"[A-Z]+[1-9][0-9]*:[A-Z][1-9][0-9]*")
    def parser(self, code: PythonCode):
        # A1:B2 -> R("A1", "B2")
        # A1 -> C("A1")
        # '0'!A1 -> Sh("0").C("A1")
        # '0'!A1:B2 -> Sh("0").R("A1", "B2")
        # '0'!non_cell -> Sh("0").G("non_cell")
        # Any combinations mixing [Sh(), R(), C(), G()] and [!, :] are not supported, and are not likely to work.

        # Matches should not happen inside strings (Comments are okay, they cannot hurt you)
        # We will make use of `ast` to find this.
        # In case of '!', each will be turned into `_` unless it is part of "!=".
        #     Anything enclosed with double quotation marks before the '!'
        #     will have the whitespaces and special characters including the quotation marks turned into underscores.
        # In case of ':', each will be turned into `_` only if it looks like `[A-Z]+[0-9]+:[A-Z]+[0-9]+`
        #  but does not have a colon prepended or appended.
        # Each (line, col) of replacement will be recorded.
        # ast.parse() and then ast.walk() to find every ast.Name().
        # Only the replaces that are found to be in the ranges are valid, to be processed into `Sh()` and `R()`.
        # Any remaining cellname-like Name-s will be processed into `C()`

        # CAUTION! arr[A1:B2] will be interpreted as arr[R("A1", "B2")], but not arr[C("A1"):C("B2")]
        # If you want the latter,
        #  Workaround 1: Put at least one space(s) around the colon.
        #  Workaround 2: Specify the step in the slice, like arr[A1:B2:1]

        def find_all_occurrences(target: str, sub: str, start=0, end=-1) -> list[int]:
            rtn = []
            start_pos = start
            while True:
                start_pos = target.find(sub, start_pos, end)
                if start_pos == -1:
                    break
                rtn.append(start_pos)
                start_pos += 1
            return rtn

        # Step 1: Sheet reference check
        # 1-1: Find all exclamation marks
        replacements_exc = find_all_occurrences(code, '!')
        # 1-2: Check if it's part of "!="
        for i in replacements_exc.copy():
            if code[i+1] == '=':
                replacements_exc.remove(i)
        # 1-3: Search in reverse, and see if it has prepending single quote
        # NOTE: Only plain string quoted with single quote is supported.
        # f-strings, function returns, etc. are not supported and not likely to work.
        replacements_side = []
        for i in replacements_exc:
            if code[i-1] != "'":
                continue
            quoted_start = code.rfind("'", 0, i-1)
            replacements_side.append(quoted_start)
            replacements_side.append(i-1)
            replacements_side.extend(
                find_all_occurrences(code, ' ', quoted_start, i-1)
            )

        # Step 2: Range check
        col_match_dict = {}
        iters = re.finditer(self.COMPILED_RANGE_RE, code)
        for match in iters:
            if len(code) < match.end(0) and code[match.end(0) + 1] == ':':
                continue
            colon_pos = code.find(':', match.start(0), match.end(0))
            col_match_dict[colon_pos] = match

        all_replacements = [
            *replacements_exc, *replacements_side, *col_match_dict.keys()
        ]
        code_inspect = str.join("", list(
            map(lambda enm: "_" if enm[0] in all_replacements else enm[1], enumerate(code))
        ))
        print(code_inspect)

        return code
