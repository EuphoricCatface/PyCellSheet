---
layout: default
section: manual
parent: ../
title: The Workspace
---

# The Workspace

## Starting and exiting PyCellSheet

Type

`$ python -m pycellsheet`

from the command prompt. 

If you want to run PyCellSheet without installation, change into the project
directory and type

`$ python pycellsheet/__main__.py`

You can also launch the file via a file manager (for example Windows Explorer).

You can exit PyCellSheet by closing the main window or by selecting
`File -> Quit` from the menu.

## PyCellSheet main window

The main window comprises the following components (see Figure):

- Titlebar
- Menu
- Toolbars
- Entryline
- Grid
- Sheet Script panel
- Table choice
- Statusbar

![PyCellSheet main window](images/screenshot_main_window.png)

## Title bar

When PyCellSheet is started, or when a new spreadsheet is created, the title
bar displays `PyCellSheet`. When a file is opened or saved, the filename is
shown in front of ` - PyCellSheet`.

Whenever a spreadsheet is changed, an asterisk `*` is displayed in front of the
title text.

## Menu

Since all actions are available via menus, this manual provides a section for
each menu. Menu items do not change position or content while working, but the
state of toggle actions changes, for example when selecting a cell in the grid.

## Toolbars

Toolbars make a subset of actions quickly accessible. Status updates of toggle actions are visualized. Actions with multiple states such as the cell renderer choice are displayed with the icon of the current state. The state is changed by clicking on the button, which is indicated by a changed icon.

Toolbar content can be shown or hidden using the toolbar menu at the right side
of each toolbar. The show/hide state is restored at the next application start.

## Entry line

Code may be entered into the grid cells via the entry line. The entry line can be detached, moved and attached at another location on the screen.

Code is accepted and evaluated when `<Enter>` is pressed. It is discarded when a new cell is selected. Multiple lines of code within one cell can be entered using `<Shift> + <Enter>`.

Code can also be entered by selecting a cell and then typing into the appearing cell editor. However, code highlighting and spell checking is displayed only in the entry line.

When data shall be displayed as text, it has to be quoted so that the code represents a Python string. In order to make such data entry easier, quotation is automatically added if `<Ctrl>+<Enter>` is pressed after editing a cell. If multiple cells are selected then `<Ctrl>+<Enter>` quotes all selected cells.

## Grid

## Grid cells

Grid cells display a representation of the cell's result. The actual output depends on the renderer. 
 * The text renderer displays the return string from the `__str__` method. 
 * The image renderer displays an image, which is derived from a QImage or from an SVG string or from a numpy.ndarray. In the latter case, the matrix values are displayed as image.
 * The markup renderer displays the result from an html string. While it only supports basic html features, it allows to display formatted text.
 * The matplotlib renderer displays a rendered matplotlib figure. Note that animations and interactive figures are not supported.

Regardless of the selected renderer, each cell has a tooltip associated that displays up to 1000 characters of the result of the text renderer. Move the mouse over a cell to display the tooltip.

#### Navigating the grid
The grid can be navigated via the arrow keys as well as using the scrollbars. Note that if a cell that is outside the view is newly selected, the view may jump to the cell.

#### Multiple views
The grid provides up to four views of the grid. These views can be accessed by dragging the splitters that initially reside at the lower and right corners of the grid. All views may be edited. Changes are immediately affecting all views. The relevant selection for editing the grid is always the view that is having focus. Note that it is not possible to have two views of two different tables at the same time.

#### Changing cell content
In order to change cell content, double-click on the cell or select the cell and edit the text in the entry line.

#### Deleting cell content
A cell can be deleted by selecting it and pressing `<Del>`. This also works for selections.

#### Selecting cells
Cells can be selected by the following actions:

- Keeping the left mouse button pressed while over cells selects a block
- Pressing `<Ctrl>` when left-clicking on cells selects these cells individually
- Clicking on row or column labels selects all cells of a row or column
- Pressing `<Ctrl> + A` selects all grid cells of the current table

Only cells of the current table can be selected at any time. Switching tables switches cell selections to the new table, i.e. the same cells in the new table are selected and no cells of the old table are selected.

Be careful when selecting all cells in a large table. Operations may take considerable time.

#### Secondary grid views
You can pull out up to 3 secondary grid views from the right and lower border
of the grid.
These views display the same grid content but can be independently scrolled and zoomed
in order to effectively work in separated sections of large grids.

Note that actions that affect the grid such as formatting cells refer to the last focused grid.

## Sheet Script panel

Sheet scripts can be edited in the Sheet Script panel. The editor allows editing
Python code that is executed when script changes are applied.

![Sheet Script panel](images/screenshot_macros.png)

The Apply button executes the current sheet script. Output (including exceptions)
is shown in the lower part of the panel.

The scope of a script is per sheet. Imported modules, helper functions, and
script-level variables become available to cells in the same sheet.

Sheet scripts are the recommended place for module imports and shared helper
code that should not be repeated in every cell.

## Table choice

Tables can be switched using the table choice. On right click, a context menu for insertion and deletion of tables is shown.

## Statusbar

Status and error messages may appear in the Statusbar. Safe mode is indicated with an attention icon `⚠`.
