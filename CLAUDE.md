# CLAUDE.md - PyCellSheet Developer Guide

## Project Overview

PyCellSheet is a fork of pyspread v2.3.1 that aims to be a comfortable middle ground between a conventional spreadsheet and pyspread's purely Pythonic approach. The key philosophical difference from pyspread is **copy-priority semantics**: cell references return `deepcopy`'d values by default, so cells behave like independent values in a normal spreadsheet, rather than pyspread's reference-priority system where cells share mutable objects.

v0.1.0, v0.2.0, v0.3.0, and v0.4.0 are released. Current development is on the v0.5.0 parser/spill feature expansion cycle.

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

## Import Pattern

**CRITICAL**: Always use try/except pattern for imports to support both development and installed modes:

```python
try:
    from pycellsheet.lib.module import Class
except ImportError:
    from lib.module import Class
```

**NEVER** import inline in methods. All imports must be in the try/except block at the top of the file.

**Example violation**:

```python
def some_method(self):
    from pycellsheet.icons import Icon  # ❌ WRONG - inline import
    ...
```

**Correct pattern**:

```python
# At top of file
try:
    from pycellsheet.icons import Icon
except ImportError:
    from icons import Icon


def some_method(self):
    # Use Icon here ✅
    ...
```

### PyCellSheet additions (where most development happens)
- `pycellsheet/lib/pycellsheet.py` - Core types and parsers: `Empty`/`EmptyCell`, `PythonCode`, `Range`, `RangeOutput`, `ExpressionParser`, `ReferenceParser`, `PythonEvaluator`, `Formatter`, `HelpText`, `CELL_META_GENERATOR`, `DependencyTracker`, `flatten_args()`
- `pycellsheet/lib/dependency_graph.py` - Dependency tracking with forward/reverse edges, dirty flags, and circular reference detection via DFS
- `pycellsheet/lib/smart_cache.py` - Dependency-aware cache with INVALID sentinel and transitive dirty checking
- `pycellsheet/lib/exceptions.py` - Custom exceptions: `PyCellSheetError`, `CircularRefError`
- `pycellsheet/lib/spreadsheet/` - Spreadsheet function library (UPPERCASE functions). `math.py` (~60 functions), `engineering.py`, `filter.py`, `info.py`, `logical.py` are largely implemented. `statistical.py` has basic functions (AVERAGE, COUNT, MAX, MIN, MEDIAN, etc.) but ~60 stubs remain. Other modules (financial, text, etc.) are mostly stubs.
- `pycellsheet/lib/spreadsheet/__init__.py` - Re-exports all function modules via `__all__`

### Modified core files
- `pycellsheet/model/model.py` - Data model layers (DictGrid -> DataArray -> CodeArray). `_eval_cell()` runs the full pipeline. `execute_sheet_script()` is the canonical sheet-script execution entrypoint.
- `pycellsheet/panels.py` - Sheet Script editor UI with draft/applied buffer and AST validation
- `pycellsheet/grid.py` - Cell display: `DisplayRole` and `ToolTipRole` delegate to `Formatter` class
- `pycellsheet/interfaces/pycs.py` - File format: `[sheet_scripts]` section stores per-sheet init scripts as `(sheet_script:'Name') linecount`; parser settings are persisted under `[parser_settings]`
- `pycellsheet/lib/string_helpers.py` - Modified `wrap_text()` to preserve newlines (for tooltips)

### Inherited pyspread lineage (largely unchanged architecture)
- `pycellsheet/grid_renderer.py` - Cell painting and border rendering
- `pycellsheet/workflows.py`, `pycellsheet/actions.py`, `pycellsheet/commands.py` - UI workflow logic
- `pycellsheet/main_window.py` - Main application window
- `pycellsheet/settings.py` - Application settings

## Data Model Architecture

Four-layer model (inherited from pyspread, modified):

- **Layer 0 - KeyValueStore**: `dict` with default `None` for missing keys
- **Layer 1 - DictGrid**: Stores cell contents as `{(row, col, table): str}`. Also holds `cell_attributes`, `sheet_scripts` (list of init script strings per sheet), `exp_parser_code`, row/col sizes.
- **Layer 2 - DataArray**: Adds slicing, insertion/deletion. Holds non-persisted state: `sheet_scripts_draft`, `sheet_globals_copyable`, `sheet_globals_uncopyable`, `exp_parser` instance, `dep_graph`, `smart_cache`.
- **Layer 3 - CodeArray**: Evaluates cells on access via `__getitem__` -> `_eval_cell()`. Uses `SmartCache` for dependency-aware caching and `DependencyGraph` for tracking cell relationships and circular reference detection.

The grid is currently a 3D dict keyed by `(row, col, table)` where the key format is **(row, col, table)**, not (col, row, table). A later goal is to replace this with an array of 2D matrices.

## Types

- `EmptyCell` - Singleton `Empty()` instance with `__copy__`/`__deepcopy__` preserving identity. Returned for empty cells. Has `__int__() -> 0`, `__float__() -> 0.0`, `__str__() -> ""`, `__bool__() -> False`, and arithmetic dunders so it behaves as zero/empty in calculations.
- `PythonCode(str)` - Marker subclass indicating the string should proceed through Reference Parser and Python Evaluator.
- `Range(RangeBase)` - 1D list with `width` and `topleft` coordinate. `__getitem__` returns rows (deepcopied). `flatten()` strips EmptyCells.
- `RangeOutput(RangeBase)` - Spill-over output. When a cell evaluates to `RangeOutput`, neighboring cells get filled with `RangeOutput.OFFSET(row, col)` expressions. Invalid offsets self-erase.
- `HelpText` - Wraps `help()` output for display in tooltips. Tries `__name__` for the query display.
- `CELL_META_GENERATOR` - Singleton that provides `CM()`/`cell_meta()` in cells. Returns a `CELL_META` object with `.code` and `.attributes` properties for the current or referenced cell.

## Dependency Tracking & Smart Caching

**DependencyGraph** (`lib/dependency_graph.py`):
- **Forward edges**: `dependencies[A2] = {A1}` means A2 depends on A1
- **Reverse edges**: `dependents[A1] = {A2}` means A1 is depended on by A2
- **Dirty flags**: `dirty = {A1, A2}` tracks cells needing recalculation
- **Circular reference detection**: DFS with recursion stack, raises `CircularRefError`
- **remove_cell(remove_reverse_edges=False)**: Preserves reverse edges by default to maintain dependency chains when cells are edited or deleted

**SmartCache** (`lib/smart_cache.py`):
- **INVALID sentinel**: `object()` distinguishes "not cached" from "cached None"
- **Dependency-aware invalidation**: Marks cell and all transitive dependents dirty
- **Validation**: Checks if cell or any transitive dependencies are dirty before returning cached value
- **Storage**: Stores original values, returns `deepcopy()` on retrieval

**DependencyTracker** (context manager in `lib/pycellsheet.py`):
- **Thread-local storage**: Tracks currently evaluating cell
- **Nested evaluation support**: Saves and restores previous context for recursive cell evaluation
- **Automatic recording**: When `C("A1")` is called, records dependency from current cell to A1
- **Circular reference detection**: Calls `check_for_cycles()` after adding each dependency

**Cell Evaluation Pipeline** with dependency tracking:
1. **__getitem__**: Checks cache validity (dirty flags + dependencies)
2. **_eval_cell**: Wraps evaluation in `DependencyTracker.track(key)` context
3. **C()/R()**: Records dependency and checks for circular references before accessing cell
4. **__setitem__**: Invalidates cache (propagates to dependents), removes forward edges only
5. **pop()**: Same as __setitem__ - preserves reverse edges for proper dependency tracking

## Known Issues and Quirks

- `coord_to_spreadsheet_ref` has a bug with column 0: it hardcodes "A" but the general algorithm would also produce "A" for col=0 via the else branch if it handled it
- Expression parser currently hardcoded to "Mixed" mode as a workaround until UI is built (see `DataArray.__init__`)
- ~~No dependency graph or recalculation ordering yet~~ ✅ **DONE** - DependencyGraph implemented with smart caching
- ~~No circular reference detection~~ ✅ **DONE** - Circular references detected via DFS during evaluation
- Dirty visualization/recalc baseline is implemented (dirty indicator + recalc actions), but dependency-inspector-style UI is still a later goal
- Named sheets and sheet renaming are implemented. Current `.pycs` contract uses `[sheet_scripts]` headers with named sheet identifiers.

## Build and Run

```bash
# Dependencies
pip install -r requirements.txt  # PyQt6, numpy, etc.

# Run
python -m pycellsheet
# or
python pycellsheet/__main__.py
```

## Testing

**Dependency Tracking Tests** (67 tests total):
- `pycellsheet/lib/test/test_dependency_graph.py` - 30 tests for DependencyGraph (add/remove, cycles, dirty flags, transitive closure)
- `pycellsheet/lib/test/test_smart_cache.py` - 20 tests for SmartCache (INVALID sentinel, dirty checking, invalidation propagation)
- `pycellsheet/model/test/test_dependency_integration.py` - 17 integration tests (C()/R()/Sh() tracking, cache invalidation chains, circular reference detection, dynamic refs)

As of 2026-02-21, `QT_QPA_PLATFORM=offscreen PYTHONPATH=pycellsheet pytest -q` passes with 896 tests on the active interpreter.

Legacy test suites under `pycellsheet/test/` and `pycellsheet/lib/test/` are part of the active regression baseline and should remain aligned with shipped PyCellSheet behavior.

## Later Goals (from design note, not yet implemented)

- Replace 3D matrix with array of 2D matrices
- ~~Dependency graph, update chaining, circular reference detection~~ ✅ **DONE**
- `compile()` caching for cell code
- UI for dirty cell visualization (refresh icon, recalc mode, F9 shortcuts)
- Custom function scripts per workbook
- Configurable Formatter (per-workspace and per-cell scope; basic static version now exists)
- `FILTER`/`COUNTIF`-style lambda-based range operations
- Asyncio/generator cells
- Cell handle dragging
