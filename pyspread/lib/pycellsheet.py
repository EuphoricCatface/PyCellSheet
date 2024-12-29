# Created by EuphoricCatface
import typing
import copy
import warnings
import re
import ast
import collections


class Empty:
    def __eq__(self, other):
        return isinstance(other, Empty)

    def __repr__(self):
        return "EmptyCell"

    def __str__(self):
        return ""

    def __int__(self):
        return 0

    def __float__(self):
        return 0.0

    def __bool__(self):
        return False

    def __add__(self, other):
        if isinstance(other, (int, float)):
            return other
        if isinstance(other, str):
            return other
        return NotImplemented

    def __radd__(self, other):
        self.__add__(other)

    def __sub__(self, other):
        if isinstance(other, (int, float)):
            return -other
        return NotImplemented

    def __rsub__(self, other):
        if isinstance(other, (int, float)):
            return other

    def __mul__(self, other):
        if isinstance(other, (int, float)):
            return 0
        return NotImplemented

    def __rmul__(self, other):
        return self.__mul__(other)

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


def spreadsheet_ref_to_coord(addr: str) -> tuple[int, int]:
    """Calculate a coordinate from spreadsheet-like address string"""
    col_str = None
    row_str = None
    for i, ch in enumerate(addr):
        if ch.isalpha():
            continue
        col_str = addr[:i]
        row_str = addr[i:]
        break

    if not col_str or not row_str:
        raise ValueError(f"Malformed spreadsheet-type address {addr=}")

    try:
        row_num = int(row_str) - 1
    except ValueError:
        raise ValueError(f"Malformed spreadsheet-type address {addr=}")

    col_str = col_str.upper()
    col_num = 0
    for ch in col_str:
        col_num *= 26
        col_num += ord(ch) - ord('A')

    return row_num, col_num


class RangeBase:
    def __init__(self, width: int, lst: typing.Optional[list] = None):
        self.lst = lst if lst else []
        self.width = width

        if len(self.lst) % width:
            warnings.warn("Length of the list is not divisible with the width")

    def flatten(self) -> list:
        # The usage doesn't care about the dimensions, so we can probably ignore EmptyCell-s
        return list(filter(lambda a: a != EmptyCell, self.lst))

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

    @property
    def height(self):
        return len(self)

    def append(self, item: typing.Any):
        self.lst.append(item)



class Range(RangeBase):
    def __init__(self, topleft: typing.Union[str, tuple[int, int]], width: int, lst: typing.Optional[list] = None):
        super().__init__(width, lst)
        self.topleft = topleft if isinstance(topleft, tuple) else spreadsheet_ref_to_coord(topleft)


class RangeOutput(RangeBase):
    def __init__(self, width: int, lst: typing.Optional[list] = None):
        super().__init__(width, lst)

    @classmethod
    def from_range(cls, r: Range):
        return cls(r.width, r.lst)

    class OFFSET:
        def __init__(self, x, y):
            self.x = x
            self.y = y


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


def flatten_args(*args: list | Range | typing.Any) -> list:
    lst = []
    for arg in args:
        if isinstance(arg, Range):
            lst.extend(arg.flatten())
            continue
        if isinstance(arg, list):
            lst.extend(arg)
            continue
        lst.append(arg)
    return lst


class ReferenceParser:
    COMPILED_RANGE_RE = re.compile(r"[A-Z]+[1-9][0-9]*:[A-Z][1-9][0-9]*")
    COMPILED_CELL_RE = re.compile(r"[A-Z]+[1-9][0-9]*")

    def __init__(self, code_array):
        self.code_array = code_array

    @staticmethod
    def sheet_name_to_idx(sheet_name):
        return int(sheet_name)

    # NYI: cache sheet, and invalidate the cache if the sheet gets modified
    class Sheet:
        def __init__(self, sheet_name: str, code_array):
            self.code_array = code_array

            self.sheet_idx = ReferenceParser.sheet_name_to_idx(sheet_name)
            self.sheet_global_var = copy.deepcopy(self.code_array.sheet_globals_copyable[self.sheet_idx])
            self.sheet_global_var.update(self.code_array.sheet_globals_uncopyable[self.sheet_idx])

        def cell_single_ref(self, addr: str):
            return copy.deepcopy(self.code_array[*spreadsheet_ref_to_coord(addr), self.sheet_idx])

        def cell_range_ref(self, addr1: str, addr2: str) -> Range:
            coord1 = spreadsheet_ref_to_coord(addr1)
            coord2 = spreadsheet_ref_to_coord(addr2)
            topleft = min(coord1[0], coord2[0]), min(coord1[1], coord2[1])
            botright = max(coord1[0], coord2[0]), max(coord1[1], coord2[1])
            width = botright[1] - topleft[1] + 1

            rtn = Range(topleft, width)
            for row in range(topleft[0], botright[0] + 1):
                for col in range(topleft[1], botright[1] + 1):
                    rtn.append(copy.deepcopy(self.code_array[row, col, self.sheet_idx]))

            return rtn

        def global_var(self, name):
            return self.sheet_global_var[name]

        C = cell_single_ref
        R = cell_range_ref
        G = global_var

    def sheet_ref(self, sheet_name):
        return self.Sheet(sheet_name, self.code_array)
    Sh = sheet_ref

    def cell_ref(self, spreadsheet_ref_notation: str, current_sheet: "ReferenceParser.Sheet"):
        target_sheet: "ReferenceParser.Sheet" = current_sheet
        exc_index = -1
        non_sheet = spreadsheet_ref_notation
        if "!" in spreadsheet_ref_notation:
            exc_index = spreadsheet_ref_notation.index("!")
            sheet_name = spreadsheet_ref_notation[:exc_index]
            sheet_name.strip('"')
            target_sheet = self.sheet_ref(spreadsheet_ref_notation[:exc_index])
            non_sheet = spreadsheet_ref_notation[exc_index + 1:]
        if ":" in non_sheet:
            col_index = non_sheet.index(":")
            range_start = exc_index + 1
            return target_sheet.cell_range_ref(
                non_sheet[:col_index],
                non_sheet[col_index + 1:]
            )
        else:
            target_sheet.global_var(non_sheet)
    CR = cell_ref

    def parser(self, code: PythonCode):
        # A1:B2 -> R("A1", "B2")
        # A1 -> C("A1")
        # "0"!A1 -> Sh("0").C("A1")
        # "0"!A1:B2 -> Sh("0").R("A1", "B2")
        # "0"!non_cell -> Sh("0").G("non_cell")
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

        def find_all_occurrences(target: str, sub: str, start=0, end=-1) -> collections.deque[int]:
            rtn = collections.deque()
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
        # 1-3: Search in reverse, and see if it has a prepending double-quoted sheet name
        # NOTE: Only plain string quoted with double quote is supported.
        # f-strings, function returns, etc. are not supported and not likely to work.
        replacements_side = []
        for i in replacements_exc:
            if code[i-1] != '"':
                continue
            quoted_start = code.rfind('"', 0, i-1)
            replacements_side.append(quoted_start)
            replacements_side.append(i-1)
            replacements_side.extend(
                find_all_occurrences(code, ' ', quoted_start, i-1)
            )

        # Step 2: Range operator check
        replacements_col = collections.deque()
        iters = re.finditer(self.COMPILED_RANGE_RE, code)
        for match in iters:
            if len(code) < match.end(0) and code[match.end(0) + 1] == ':':
                continue
            colon_pos = code.find(':', match.start(0), match.end(0))
            replacements_col.append(colon_pos)

        # Step 3: Replace
        all_replacements = {
            *replacements_exc, *replacements_side, *replacements_col }
        code_inspect = str.join("", list(
            map(lambda enm: "_" if enm[0] in all_replacements else enm[1], enumerate(code)) ))
        parsed = ast.parse(code_inspect)

        # Step 4: Get all names
        # Step 5 (combined): Separate out single-cell-like names (without a sheet reference!)
        single_cell_idx_name = dict()

        names_indices: list[tuple[int, int]] = []
        split_lines = code_inspect.splitlines()
        line_lengths = [0]
        for line in split_lines:
            line_lengths.append(line_lengths[-1] + len(line))
        for node in ast.walk(parsed):
            if not isinstance(node, ast.Name):
                continue
            start_index = line_lengths[node.lineno - 1] + node.col_offset
            end_index = line_lengths[node.end_lineno - 1] + node.end_col_offset
            if re.fullmatch(self.COMPILED_CELL_RE, node.id):
                single_cell_idx_name[(start_index, end_index)] = node.id
            else:
                names_indices.append((start_index, end_index))

        # Step 6: Find applicable names
        names_idx_applicable = dict()
        for start, end in names_indices:
            if not (replacements_exc or replacements_col):
                break
            while replacements_exc and replacements_exc[0] < start:
                del replacements_exc[0]
            while replacements_col and replacements_col[0] < start:
                del replacements_col[0]
            exc_idx = -1
            col_idx = -1
            if replacements_exc and start <= replacements_exc[0] < end:
                exc_idx = replacements_exc.popleft()
            if replacements_col and start <= replacements_col[0] < end:
                col_idx = replacements_col.popleft()
            if (exc_idx == -1) and (col_idx == -1):
                continue
            result = (exc_idx, col_idx)
            names_idx_applicable[(start, end)] = result

        # Step 7: Prepare replacements
        names_idx_replacement_str = dict()
        single_cell_indices = collections.deque(sorted(single_cell_idx_name.keys()))
        for (start, end), (exc_idx, col_idx) in names_idx_applicable.items():

            # Step 7-0: Prepare single cell reference
            while True:
                if not single_cell_indices:
                    break
                s_index = single_cell_indices.popleft()
                if s_index[0] > end:
                    break
                s_str = f"C(\"{single_cell_idx_name.pop(s_index)}\")"
                names_idx_replacement_str[s_index] = s_str

            # Step 7-1: Prepare sheet reference
            sheet_parsed = ""
            if exc_idx != -1:
                sheet_name = code[start:exc_idx]
                if sheet_name.startswith('"') and sheet_name.endswith('"'):
                    sheet_parsed = f"Sh({sheet_name})."
                else:
                    sheet_parsed = f"Sh('{sheet_name}')."

            # Step 7-2: Prepare the rest
            range_or_cell_or_global_parsed = ""
            if col_idx != -1:
                range_start = (exc_idx + 1) if exc_idx != -1 else start
                range1 = code[range_start:col_idx]
                range2 = code[col_idx + 1:end]
                range_or_cell_or_global_parsed = f"R(\"{range1}\", \"{range2}\")"
            else:
                # If this code piece doesn't have a range, then it has to have a sheet reference
                assert exc_idx != -1, "Filtered name has neither sheet ref nor range op?"
                var_start = exc_idx + 1
                var_end = end
                var = code[var_start:var_end]
                if re.fullmatch(self.COMPILED_CELL_RE, var):
                    range_or_cell_or_global_parsed = f"C(\"{var}\")"
                else:
                    range_or_cell_or_global_parsed = f"G(\"{var}\")"
            names_idx_replacement_str[(start, end)] = [sheet_parsed, range_or_cell_or_global_parsed]

        # Step 7-0-1: Prepare remaining single cell references
        while single_cell_indices:
            s_idx = single_cell_indices.popleft()
            s_str = f"C(\"{single_cell_idx_name.pop(s_idx)}\")"
            names_idx_replacement_str[s_idx] = s_str

        # Step 8: Finally, assemble the code with the replacements
        last_end = 0
        parsed_code_list = []
        for (start, end), replacements in names_idx_replacement_str.items():
            parsed_code_list.append(code[last_end:start])
            parsed_code_list.extend(replacements)
            last_end = end
        parsed_code_list.append(code[last_end:])

        return str.join("", parsed_code_list)

class PythonEvaluator:
    @staticmethod
    def exec_then_eval(code: PythonCode,
                       _globals: dict = None, _locals: dict = None):
        """execs multiline code and returns eval of last code line

        :param code: Code to be executed / evaled
        :param _globals: Globals dict for code execution and eval
        :param _locals: Locals dict for code execution and eval

        """

        if _globals is None:
            _globals = {}

        if _locals is None:
            _locals = {}

        block = ast.parse(code, mode='exec')

        # assumes last node is an expression
        last_body = block.body.pop()
        last = ast.Expression(last_body.value)

        exec(compile(block, '<string>', mode='exec'), _globals, _locals)
        res = eval(compile(last, '<string>', mode='eval'), _globals, _locals)

        return res

    @staticmethod
    def range_output_handler(code_array, range_output: RangeOutput, current_key):
        x1, y1, current_table = current_key
        for xo in range(range_output.height):
            for yo in range(range_output.width):
                if xo == 0 and yo == 0:
                    continue
                if code_array((x1 + xo, y1 + yo, current_table)) == f"RangeOutput.OFFSET({xo}, {yo})":
                    code_array[x1 + xo, y1 + yo, current_table] = ""
                if code_array[x1 + xo, y1 + yo, current_table] != EmptyCell:
                    print(x1 + xo, y1 + yo, code_array[x1 + xo, y1 + yo, current_table], )
                    raise ValueError("Cannot expand RangeOutput")
        for xo in range(range_output.height):
            for yo in range(range_output.width):
                if xo == 0 and yo == 0:
                    continue
                code_array[x1 + xo, y1 + yo, current_table] = \
                    f"RangeOutput.OFFSET({xo}, {yo})"

    @staticmethod
    def range_offset_handler(code_array, range_offset: RangeOutput.OFFSET, current_key):
        x, y, table = current_key
        xo = range_offset.x
        yo = range_offset.y
        return code_array[x-xo, y-yo, table][xo][yo]

