---
layout: default
section: manual
parent: ../
title: Basic Concepts
---

# Basic concepts

## Python code as cell language

PyCellSheet executes Python code in each cell. This is similar to typing in the
Python shell. Normal cells are only executed when required, for example for
displaying results. **Execution order between cells is not guaranteed to be
stable** and may differ across Python versions.

Normally, one line of code containing a [Python expression](https://docs.python.org/3.7/reference/expressions.html)
is entered in each cell. However, a cell can contain additional lines of
arbitrary Python code that precede the final expression. The object returned by
the cell is always the result of **the last line's expression**. Note that only
the last line **must** be an expression.

In order to enter a new line in one cell, press `<Shift> + <Enter>`. Only pressing `<Enter>` accepts the entered code and switches to the next cell.

While editing cell code in the entry line (not in a cell editor), pressing
`<Insert>` switches the grid to selection mode, indicated by an icon in the
status bar. In selection mode, selecting cells in the grid generates relative
references in the entry line. Pressing `<Meta>` while clicking generates
absolute references instead. You can exit selection mode by selecting the entry
line, by focusing the entry line and pressing `<Insert>` again, or by pressing
`<Escape>` while inside the grid. Note that you cannot edit cell code in cell
editors while in selection mode.

### Example

Let us enter an expression into the top left cell in table 0:
```python
1 + 1
```
After pressing `<Enter>`, the cell displays
```python
2
```
as expected. List comprehensions are also valid expressions.
```python
[i ** 2 for i in range(100) if i % 3]
```
works.

However, statements such as
```python
import math
```
are not valid in the last line of a cell. In contrast,
```python
import math
math
```
is valid. Note that multi-line cells have been added to make some 3rd party modules such as `rpy2` accessible. Abusing the feature for module imports and complex control flows is discouraged.

## Module import

Modules should be imported via the Sheet Script panel. If the panel is hidden,
open it from `View -> Sheet Script`. Enter code in the editor and press
`Apply`. If errors are raised, they are displayed in the message area below the editor.

While it is possible to import modules from within a cell, there are drawbacks:
 * The module is not imported until the cell is executed, which is not guaranteed in any way.
 * A spreadsheet may quickly become hard to understand when importing from cells.

## Variable assignment

Besides Python expressions, one variable assignment is accepted within the last line of a cell. The assignment consists of one variable name at the start followed by “=” and a Python expression. The variable is considered as global. Therefore, it is accessible from other cells.

For example `a = 5 + 3` assigns `8` to the global variable `a`.

`b = c = 4`, `+=`, `-=` etc. are not valid in the last line of a cell. In
preceding lines, such code is valid. However, variables assigned there stay in
the local scope of the cell, while assignment in the last line enters the
global scope.

Since evaluation order of cells is not guaranteed, assigning a variable twice may result in unpredictable behaviour of the spreadsheet.

## Displaying results in the grid

Result objects from cells are interpreted by the cell renderer. Therefore, two
renderers may display the same object differently. Cell renderers can be
changed in `Format -> Cell renderer`. PyCellSheet provides four renderers:

1. The `Text renderer` is selected by default. It displays the string representation of the object as plain text. The exception is the object `None`, which is displayed as empty cell. This behavior allows empty cells returning `None` without the grid appearing cluttered.

2. The `Image renderer` renders a QImage object as an image. It renders a 2D array as a monochrome bitmap and a 3D array of shape `(m, n, 3)` as a color image. Furthermore, it renders a `str` object with valid svg content as an SVG image.

3. The `Markup renderer` renders the object string representation as markup text. It supports the [limited subset of static HTML 4 / CSS 2.1](https://doc.qt.io/qt-5/richtext-html-subset.html) that is provided by QT5's [QTextDocument](https://doc.qt.io/qt-5/qtextdocument.html) class.

4. The `Matplotlib chart renderer` renders a [matplotlib Figure object](https://matplotlib.org/api/_as_gen/matplotlib.figure.Figure.html#matplotlib.figure.Figure).

Note that the concept of different cell renderers was introduced in upstream
pyspread and retained in PyCellSheet.

## Absolute cell access

The result objects, for which string representations are displayed in the
grid, can be accessed from other cells via the grid getitem method. The grid
object is globally accessible via the name `S`. For example
```python
S[3, 2, 1]
```
returns the result object from the cell in row 3, column 2, table 1. This type of access is
called **absolute** because the targeted cell does not change when the code is copied to another cell similar to a call "$A$1" in a common spreadsheet.

## Relative cell access

In order to access a cell relative to the current cell position, 3 variables X, Y and Z are provided that point to the row, the column and the table of the calling cell. The values stay the same for called cells. Therefore,
```python
S[X-1, Y+1, Z]
```
returns the result object of the cell that is in the same table two rows above and 1 column right of the current cell. This type of access is called relative because the targeted cell changes when the code is copied to another cell similar to a call "A1" in a common spreadsheet.

## Slicing the grid

Cell access can refer to multiple cells by slicing similar to slicing a matrix in numpy. Therefore, a slice object is used in the getitem call. For example
```python
S[:3, 0, 0]
```
returns the first three rows of column 0 in table 0 and
```python
S[1:4:2, :2, -1]
```
returns row 1 and 3 and column 0 and 1 of the last table of the grid.

The returned object is a numpy object array of the result objects. This object allows utilization of the numpy commands such as numpy.sum that address all array dimensions instead of only the outermost. For example
```python
numpy.sum(S[1:10, 2:4, 0])
```
sums up the results of all cells from 1, 2, 0 to 9, 3, 0 instead of summing
each row, which Python's `sum` function does.

One disadvantage of this approach is that slicing results are not sparse as the grid itself and therefore consume memory for each cell. Therefore,
```python
S[:, :, :]
```
may lock up or even crash with a memory error if the grid size is too large.

## How cells are evaluated

PyCellSheet differs from traditional spreadsheets in how it evaluates cells.
The approach is difficult for Python because side-effects must be taken into
account. PyCellSheet employs dependency tracking and smart caching.

PyCellSheet evaluates applied sheet script code and visible cell code.
Whenever a cell that has not been evaluated before becomes visible, it is
evaluated on the fly. Each evaluated cell result is stored in smart cache.
When a cell is evaluated again, the cached result is used if the cell and its
dependencies have not changed. This also applies for cells addressed from other
cells (for example through helper references).

Cell caches are invalidated when:
  * sheet script changes have been applied
  * a cell code has been changed
  * any of the cell's dependencies have been changed

The dependency graph automatically tracks which cells depend on other cells, ensuring that cached values are only used when they are still valid.

## Everything is accessible

All parts of PyCellSheet are written in Python, therefore all objects can be
accessed from within each cell. This is also the case for external modules.

There are five convenient “magical” objects, which are merely syntactic sugar: `S`, `X`, `Y`, `Z` and `nn`.

`S` is the grid data object. It is ultimately based on a `dict`. However, it consists of several layers on top. When indexing or slicing, it behaves similarly to a 3D numpy-array that returns result objects. When calling it (like a function) with a 3 tuple, it returns the cell code.

`X`, `Y` and `Z` represent the current cell coordinates. When copied to another cell, these coordinates change accordingly. This approach allows relative addressing by adding the relative coordinates to X, Y or Z. Therefore, no special relative addressing methods are needed.

`nn` is a function that flattens a NumPy array and removes all `None` objects.
This function makes special-casing `None` for operations such as `sum`
unnecessary.

## Security

Since Python expressions are evaluated in PyCellSheet, a spreadsheet is as
powerful as any program. It could harm the system or even send confidential
data to third parties over the Internet.

The risk is similar to other office applications, which is why precautions are
provided. The model is that you, the user, are trusted and outside files are
not. When starting PyCellSheet for the first time, a secret key is generated
and stored in `~/.config/pyspread/pyspread.conf` (on many Linux systems). You
can edit this key in the Preferences dialog (`File -> Preferences...`).

If you save a file then a signature is saved with it (suffix `.sig` for
`.pycs`/`.pycsu` files). Only if the signature is valid for the stored secret
key, you can re-open the file directly. Otherwise, e.g. if anyone else opens
the file, it is displayed in `Safe mode`, i.e. each cell displays the cell
code and no cell code is evaluated. The user can approve the file by selecting
`Approve file` in the `File` menu. Afterwards, cell code is evaluated. When the
user then saves the file, it is newly signed.

Never approve foreign spreadsheet files unless you have thoroughly checked each
cell. Each cell may delete valuable files. Malicious cells are likely to be
hidden in large sheets. If unsure, inspect the `.pycsu`/`.pycs` file contents
first. It may also be a good idea to run PyCellSheet with a restricted user or
sandbox.

## Current Limitations

* Execution of certain operations cannot be interrupted or terminated if slow. An example is creating very large integers. Such long-running code may block PyCellSheet.
* Maximum recursion depth is limited. Its value is a tradeoff between handling complex cell dependencies and time until stopping when cyclic dependencies are present.
* Legacy Python2 code from older pyspread files is not automatically converted when opening old save files.
