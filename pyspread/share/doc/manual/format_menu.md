---
layout: default
section: manual
parent: ../
title: Format actions
---

# Format menu

## Format → Copy format

Copies only the format of the selected cells. Copying formats has been separated from copying content in order to prevent unwanted behavior.

## Format → Paste format

Pastes copied cell formats.

----------

## Format → Font
Assigns a font including style and size to the current cell if no selection is present. If a selection is present, the  font is assigned to each selected cell.

Fonts are not embedded in `.pycs`/`.pycsu` files. Therefore, fonts have to be
available on the target system when opening a file, otherwise the default font
is used.

## Format → Bold
Bold toggles the current selection’s cell font bold attribute. If no cell is selected, then the attribute is toggled for the current cell. The shortcut is `<Ctrl> + B`.

## Format → Italics
Italics toggles the current selection’s cell font italics attribute. If no cell is selected, then the attribute is toggled for the current cell. The shortcut for Italics is `<Ctrl> + I`.

## Format → Underline
Underline toggles the current selection’s cell font underline attribute. If no cell is selected, then the attribute is toggled for the current cell. The shortcut for Underline is `<Ctrl> + U`.

## Format → Strikethrough
Strikethrough toggles the current selection’s cell font strikethrough attribute. If no cell is selected, then the attribute is toggled for the current cell.


----------

## Format → Cell renderer

Opens a sub-menu, in which the cell renderer for the current cell can be chosen.

## Format → Lock cell

Lock toggles the current selection's cell lock attribute. If no cell is
selected, then the current cell is locked. Locking means that the cell cannot
be edited until it is unlocked again.

## Format → Button cell

Creates a button cell from the current cell. Button cells may be employed to
provide an intuitive interface for the user that allows executing functions
from the Sheet Script panel by pressing a button in the grid.

On selecting button cell, a dialog for querying the user for a button text is displayed. Next, a button that is labeled with this text is displayed in the cell instead of cell results. The button cell's code is executed when activating the button (on release).

Note that button cells can be activated even if they are locked.

## Format → Merge cells

Merge cells merge all cells that are in the bounding box of the current selection. If there is no selection the the current cell will not be merged or unmerged if already merged. Merged cells act as one. Output is shown for the top left cell, which stays intact upon a merge.


----------


## Format → Rotation

Opens a sub-menu, in which cell rotation can be chosen from 0, 90, 180 and 270 degree. The chosen rotation is applied to all cells in the current selection. If no cell is selected, then it is applied to the current cell. Besides text output, rotation also applied to bitmap and vector graphics that are displayed in the cell. Matplotlib charts may be dislocated in rotated cells.

## Format → Justification

Opens a sub-menu, in which cell justifications can be chosen from left, center and right. The chosen justification is applied to all cells in the current selection. If no cell is selected, then it is applied to the current cell. Besides text output, justification also applied to bitmap and vector graphics that are displayed in the cell.

## Format → Alignment

Opens a sub-menu, in which cell alignment can be chosen from top, center and bottom. The chosen alignment is applied to all cells in the current selection. If no cell is selected, then it is applied to the current cell. Besides text output, alignment also applied to bitmap and vector graphics that are displayed in the cell.

----------

## Format → Formatted borders

When changing border color or width, the command affects the selection, or if no selection is present. the current cell.

Since a cell has four borders, all borders are affected by default. The border choice box allows changing this
behaviour by providing the following options:

- All borders: All borders are affected
- Left border: Only the left border of the smallest containing bounding box is affected.
- Right border: Only the right border of the smallest containing bounding box is affected.
- Top border: Only the top border of the smallest containing bounding box is affected.
- Bottom border: Only the bottom border of the smallest containing bounding box is affected.
- Outer borders: All outer borders of the smallest containing bounding box are affected.
- Top and bottom borders: Only the top and the bottom border of the smallest containing bounding box are affected.

## Format → Border width

Choice box that changes cell border widths. The section Border choice box explains, which borders are affected. There are eleven different border widths. The first width is `0`, which means that no border is drawn.

----------

## Format → Text color

Opens a dialog, in which a color can be chosen. On o.k., the text color is set to the chosen color for all cells in the current selection. If no cell is selected, then the text color is set for the current cell.

## Format → Line color

Invokes a color choice dialog that changes cell border color. The section Border choice box explains, which borders are affected. The border color is chosen as an RGB value. The color choice dialog may look different depending on the operating system.

## Format → Background color
Opens a dialog, in which a color can be chosen. On o.k., the background color is set to the chosen color for all cells in the current selection. If no cell is selected, then the background color is set for the current cell.
