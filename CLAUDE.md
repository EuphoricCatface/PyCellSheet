# CLAUDE.md - PyCellSheet Developer Guide

## Project Overview

PyCellSheet is a fork of pyspread v2.3.1 that aims to be a comfortable middle ground between a conventional spreadsheet and pyspread's purely Pythonic approach. The key philosophical difference from pyspread is **copy-priority semantics**: cell references return `deepcopy`'d values by default, so cells behave like independent values in a normal spreadsheet, rather than pyspread's reference-priority system where cells share mutable objects.

The project is at an early proof-of-concept stage (v0.0.5). It has been dormant for a couple of years.

## Design Philosophy

See `Application Design Note.txt` for the full design document. Key principles:

- **Copy-priority**: `deepcopy` on every cell reference. Picklable objects are deepcopied; unpicklable ones (e.g. `types.ModuleType`) are passed by reference with a warning. This means `A1 = [1,2,3]; A2 = A1.sort()` leaves A1 unchanged and A2 shows `None`.
- **Per-sheet init scripts** (not global macros): Each sheet has its own initialization script that defines imports and global variables. Init scripts cannot reference cells or cross-reference other sheets.
- **Spreadsheet function library** provided via imports in init scripts, not built-in. Functions are UPPERCASE to avoid shadowing Python builtins.
- **Expression Parser** is configurable per workspace with four modes, allowing users at different comfort levels to use the tool.

## Cell Evaluation Pipeline

```
Cell Contents -> Expression Parser -> Reference Parser -> Python Evaluator -> (Formatter)
    (str)         (str or Result)       (PythonCode)        (Result)          (str for display)
```

1. **Expression Parser** (`ExpressionParser` in `pycellsheet.py`): Determines if cell contents are code or literal values. Returns either a Python object (immediate result) or a `PythonCode(str)` subclass for further evaluation. Four preset modes:
   - **Pure Pythonic**: Everything is code
   - **Mixed**: `'` prefix marks string literals; everything else is code
   - **Reverse Mixed** (default): `>` prefix marks code; everything else is a string. `'` prefix can escape the `>`.
   - **Pure Spreadsheet**: `=` prefix marks code; tries int/float parsing; rest is string

   **Important convention for Mixed/Pure Spreadsheet modes**: Since `'` at the start of a cell is consumed as the string-literal prefix, users must use double quotes (`"`) when writing Python code that begins with a string literal (e.g. `"hello" + str(expr)` instead of `'hello' + str(expr)`).

2. **Reference Parser** (`ReferenceParser` in `pycellsheet.py`): Transforms spreadsheet-style references into Python function calls using AST analysis:
   - `A1` -> `C("A1")` (single cell)
   - `A1:B2` -> `R("A1", "B2")` (range)
   - `"Sheet1"!A1` -> `Sh("Sheet1").C("A1")` (cross-sheet)
   - `CR(expr)` is the catch-all that parses a string at runtime

3. **Python Evaluator** (`PythonEvaluator` in `pycellsheet.py`): Uses `exec()`/`eval()` with:
   - Globals: deepcopy of the current sheet's init script globals
   - Locals: helper functions (`C`, `R`, `Sh`, `CR`, `G`, `help`, `RangeOutput`, `CM`/`cell_meta`)
   - Last expression in the cell becomes the result (REPL-style)

4. **Formatter** (`Formatter` in `pycellsheet.py`): Static methods `display_formatter` and `tooltip_formatter` with type-based dispatch. Exceptions show class name in cell + details in tooltip; HelpText shows query in cell + contents in tooltip; RangeOutput shows first element. Used by `grid.py`'s `DisplayRole`/`ToolTipRole`.

## Key Source Files

### PyCellSheet additions (where most development happens)
- `pyspread/lib/pycellsheet.py` - Core types and parsers: `Empty`/`EmptyCell`, `PythonCode`, `Range`, `RangeOutput`, `ExpressionParser`, `ReferenceParser`, `PythonEvaluator`, `Formatter`, `HelpText`, `CELL_META_GENERATOR`, `flatten_args()`
- `pyspread/lib/spreadsheet/` - Spreadsheet function library (UPPERCASE functions). `math.py` (~60 functions), `engineering.py`, `filter.py`, `info.py`, `logical.py` are largely implemented. `statistical.py` has basic functions (AVERAGE, COUNT, MAX, MIN, MEDIAN, etc.) but ~60 stubs remain. Other modules (financial, text, etc.) are mostly stubs.
- `pyspread/lib/spreadsheet/__init__.py` - Re-exports all function modules via `__all__`

### Modified pyspread files
- `pyspread/model/model.py` - Data model layers (DictGrid -> DataArray -> CodeArray). `_eval_cell()` runs the full pipeline. `execute_macros()` runs init scripts and separates globals into copyable/uncopyable dicts.
- `pyspread/panels.py` - Init script editor UI with draft/applied buffer and AST validation
- `pyspread/grid.py` - Cell display: `DisplayRole` and `ToolTipRole` delegate to `Formatter` class
- `pyspread/interfaces/pycs.py` - File format: `[macros]` section stores per-sheet init scripts as `(macro:X) linecount`
- `pyspread/lib/string_helpers.py` - Modified `wrap_text()` to preserve newlines (for tooltips)

### Inherited pyspread (largely unchanged)
- `pyspread/grid_renderer.py` - Cell painting and border rendering
- `pyspread/workflows.py`, `pyspread/actions.py`, `pyspread/commands.py` - UI workflow logic
- `pyspread/main_window.py` - Main application window
- `pyspread/settings.py` - Application settings

## Data Model Architecture

Four-layer model (inherited from pyspread, modified):

- **Layer 0 - KeyValueStore**: `dict` with default `None` for missing keys
- **Layer 1 - DictGrid**: Stores cell contents as `{(row, col, table): str}`. Also holds `cell_attributes`, `macros` (list of init script strings per sheet), `exp_parser_code`, row/col sizes.
- **Layer 2 - DataArray**: Adds slicing, insertion/deletion. Holds non-persisted state: `macros_draft`, `sheet_globals_copyable`, `sheet_globals_uncopyable`, `exp_parser` instance.
- **Layer 3 - CodeArray**: Evaluates cells on access via `__getitem__` -> `_eval_cell()`. Maintains `result_cache` and `frozen_cache`.

The grid is currently a 3D dict keyed by `(row, col, table)`. A later goal is to replace this with an array of 2D matrices.

## Types

- `EmptyCell` - Singleton `Empty()` instance with `__copy__`/`__deepcopy__` preserving identity. Returned for empty cells. Has `__int__() -> 0`, `__float__() -> 0.0`, `__str__() -> ""`, `__bool__() -> False`, and arithmetic dunders so it behaves as zero/empty in calculations.
- `PythonCode(str)` - Marker subclass indicating the string should proceed through Reference Parser and Python Evaluator.
- `Range(RangeBase)` - 1D list with `width` and `topleft` coordinate. `__getitem__` returns rows (deepcopied). `flatten()` strips EmptyCells.
- `RangeOutput(RangeBase)` - Spill-over output. When a cell evaluates to `RangeOutput`, neighboring cells get filled with `RangeOutput.OFFSET(row, col)` expressions. Invalid offsets self-erase.
- `HelpText` - Wraps `help()` output for display in tooltips. Tries `__name__` for the query display.
- `CELL_META_GENERATOR` - Singleton that provides `CM()`/`cell_meta()` in cells. Returns a `CELL_META` object with `.code` and `.attributes` properties for the current or referenced cell.

## Known Issues and Quirks

- `coord_to_spreadsheet_ref` has a bug with column 0: it hardcodes "A" but the general algorithm would also produce "A" for col=0 via the else branch if it handled it
- Expression parser currently hardcoded to "Mixed" mode as a workaround until UI is built (see `DataArray.__init__`)
- No dependency graph or recalculation ordering yet
- No circular reference detection
- Sheet names are currently numeric indices (string `"0"`, `"1"`, ...) - named sheets are a later goal

## Build and Run

```bash
# Dependencies
pip install -r requirements.txt  # PyQt6, numpy, etc.

# Run
python -m pyspread
# or
python pyspread/__main__.py
```

## Testing

Inherited pyspread tests exist in `pyspread/test/` and `pyspread/lib/test/` but have not been updated for PyCellSheet changes. No PyCellSheet-specific tests exist yet.

## Later Goals (from design note, not yet implemented)

- Replace 3D matrix with array of 2D matrices
- Dependency graph, update chaining, circular reference detection
- `compile()` caching for cell code
- Named sheets (currently numeric)
- Custom function scripts per workbook
- Configurable Formatter (per-workspace and per-cell scope; basic static version now exists)
- `FILTER`/`COUNTIF`-style lambda-based range operations
- Asyncio/generator cells
- Cell handle dragging
