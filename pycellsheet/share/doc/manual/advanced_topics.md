---
layout: default
section: manual
parent: ../
title: Advanced topics
---

# Advanced topics

## Accessing the current cell from a sheet script helper

The variables `X`, `Y`, and `Z` are cell-context values. In sheet scripts they
are not automatically bound to a calling cell. To use caller coordinates in a
script helper or external library function, pass coordinates as parameters.

## Conditional formatting

For conditionally formatting the background color of a cell, enter
```python
def color(value, condition_func, X, Y, Z):
    if condition_func(value):
        color = 255, 0, 0
    else:
        color = None

    S.cell_attributes[X,Y,Z]["bgcolor"] = color

    return value
```
into the Sheet Script panel and
```
color(5, lambda x: x>4, X, Y, Z)
```
into a cell.
If you change the first parameter in the cell's function from 5 into 1 then the background color changes back to white.


## Adjusting the float accuracy that is displayed in a cell

While one can use the `round` function to adjust accuracy, this may be tedious for larger spreadsheets.
`class_format_functions` is available as a dictionary mapping a type to a
display-formatting function.
It should be configured in the Sheet Script panel. The following code adjusts
`float` output:

```python
class_format_functions[float] = lambda x: f"{x:.4f}"
```

Note that depending on your use case, `float` may not be the best choice. Consider using Python's builtin [decimal module](https://docs.python.org/3/library/decimal.html). When creating decimals from given numbers, do not forget to provide them as strings.

```
Decimal(3.2)    # 3.20000000000000017763568394002504646778106689453125
Decimal('3.2')  # 3.2
```

For arbitrary precision, you may want to try out the [mpmath module](https://pypi.org/project/mpmath/), which
provides the pretty attribute for human friendly representation.

If you are working with currencies, you may be interested in the [Python Money Class](https://pypi.org/project/money/).
Putting their currency presets approach, which is stated on their project page,
into the Sheet Script panel works well:

```
class EUR(Money):
    def __init__(self, amount='0'):
        super().__init__(amount=amount, currency='EUR')
```


## Cyclic references

PyCellSheet uses dependency tracking and circular-reference detection. If a
cycle is created, the cycle is detected and reported as an error.

## Result stability

Result stability is not guaranteed when redefining global variables because execution order may be changed. This happens for when in large spreadsheets the result cache is full and cell results that are purged from the cache are re-evaluated.

## Security annoyance when approving files in read only folders

If a `.pycs` or `.pycsu` file is located in a folder without write or file
creation permissions, a signature file cannot be created. Therefore, the file
must be approved each time it is opened.

## Handling large amounts of data

While the main grid may be large, filling many cells can consume considerable
memory. For large datasets, load data into one cell (for example, as a NumPy
array) and work from there.

## Substituting pivot tables

In the examples directory, a Pivot table replacement is shown using list comprehensions.

## Memory consumption for sheets with many matplotlib charts

If there are hundreds of charts in a spreadsheet, PyCellSheet may consume
considerable memory. This is most obvious when printing or creating PDF files.
