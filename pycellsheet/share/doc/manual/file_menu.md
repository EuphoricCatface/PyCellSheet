---
layout: default
section: manual
parent: ../
title: File actions
---

# File menu

## File → New

![New dialog](images/screenshot_new_dialog.png)

An empty spreadsheet can be created by **`File → New`**.

A Dialog pops up, in which the size of the new spreadsheet grid can be entered. Note
that grid size has been limited to 1 000 000 rows, 100 000 columns and 100 tables.

## File → Open

Loading a spreadsheet from disk can be initiated with **`File → Open`**. Opening a
file expects a file with extension `.pycsu`, `.pycs` or - if the `pycel`
package is installed - `.xlsx`. The native format is PyCellSheet specific.
`.pycs` is the compressed variant of `.pycsu`. `pycsu` is the default option
and is useful with version control systems such as git.

Using `.xlsx`, Excel files can be opened. Excel formulas are converted into
Python code via the `pycel` package. Note that many common files may not work
as expected because `pycel` does not support all features (for example,
relative references may be treated as absolute). Using the resulting files may
also require `pycel` to be installed to run without errors.

Since spreadsheet files are ultimately Python programs, a file is opened in safe mode if
it has not been previously signed with the key that is shown in the Preference dialog.

Safe mode means that the cell content is loaded and displayed in the grid but not executed, so that 2+2 remains 2+2 and is not computed into 4. You can leave safe mode with **`File → Approve file`**.

----------

## File → Save

A spreadsheet can be stored to disk with **`File → Save`** . If a file is already opened, it is
overwritten. Otherwise, Save prompts for a filename.

When a file is saved, it is signed in an additional file with the suffix `.sig` using the key that is shown in the Preference dialog. Note that the save file is not encrypted.

The `.pycsu` file format is a UTF-8 text file (without BOM) with sectioned data:

```
[PyCellSheet save file version]
0.0
[shape]
1000 100 3
[sheet_names]
0 Sheet0
[grid]
7 22 0 'Testcode1'
8 9 0 'Testcode2'
[attributes]
[] [] [] [] [(0, 0)]
0 'textfont' u'URW Chancery L'
[] [] [] [] [(0, 0)]
0 'pointsize' 20
[row_heights]
0 0 56.0
7 0 25.0
[col_widths]
0 0 80.0
[macros]
(macro:0) 1
import math
```

## File → Save As
**`File → Save As`** saves the spreadsheet as does **`File → Save`**. While
Save overwrites currently opened files, Save As always prompts for a file name.

----------

## File → Import

![CSV import dialog](images/screenshot_csv_import.png)

A csv file can be imported via `File → Import`.

If the selected file is not encoded in UTF-8, an encoding has to be chosen in a dialog. If the file is encoded in UTF-8 or if the chosen encoding can be read, the CSV file import dialog opens. In this dialog, CSV import options can be set. Furthermore, target Python types can be specified, so that import of dates becomes possible. The grid of the import dialog only shows the first few rows of CSV files in order to give an impression how import data will look.

Be careful when importing a file that contains code, because code in the CSV file might prove harmful.

For importing money data, it is recommended to use the decimal or the Money datatype. The latter supports specific currencies and requires the optional dependency [py-moneyed](https://pypi.org/project/py-moneyed/).

## File → Export

PyCellSheet can export spreadsheets to `csv` files and `svg` files.

When exporting a file then a dialog is displayed, in which the area to be exported can be chosen.

When exporting a `.csv` file then an export dialog is shown next, in which the format of the csv file may be specified. The start of the exported file is shown below the options.

![CSV export dialog](images/screenshot_csv_export.png)

----------

## File → Approve file

PyCellSheet cells contain Python code. Instead of a special purpose language,
you enter code in a general purpose language. This code can do everything that
the operating system allows.

Even though the situation differs little to common spreadsheet applications, the approach makes malicious attacks easy. Instead of knowing how to circumvent blocks of the domain specific language to make the computer do what you want, everything is straight forward.

In order to make working with PyCellSheet as safe as possible, save files
(`.pycs` and `.pycsu`) are signed in a signature file. Only a user with a
private key can open the file without approving it. Files without a valid
signature are opened in safe mode, i.e. code is displayed and not executed.
The file can then be approved after inspection.

![Approve dialog](images/screenshot_approve_dialog.png)

Therefore, never approve foreign files unless you have checked each cell
thoroughly. If unsure, inspect the `.pycsu`/`.pycs` content first. It may also
be a good idea to run PyCellSheet with restricted privileges or in a sandbox.

## File → Print preview

When selecting print preview, a dialog box is shown, in which the spreadsheet extents (rows, columns and tables) that should be printed can be selected.

![Print preview dialog 1](images/screenshot_print_preview_1.png)

After pressing o.k., a second dialog window displays the print preview.

![Print preview dialog 2](images/screenshot_print_preview_2.png)

## File → Print

**Print** prints the spreadsheet. First, a dialog similar to **Print preview** is opened, in which the spreadsheet extents (rows, columns and tables) can be selected. After pressing o.k., a operation system specific print dialog is opened. This dialog provided an option to start printing.

----------

## File → Preferences

The preferences dialog allows changing:

- **Signature key for files**: The private key that is used for signing `.pycs` and `.pycsu` files
- **Number of recent files**: The maximum number of files displayed in recent files. Changes come into effect after restart.

![Preferences dialog 2](images/screenshot_preferences_dialog.png)

On *nix, configuration is stored in `~/.config/PyCellSheet/PyCellSheet.conf`
- This file is created when PyCellSheet is started the first time
- Removing it resets configuration.

----------

## File → Quit

**`File → Quit`** exits PyCellSheet. If changes have been made to a new or
loaded file then a dialog pops up and asks if the changes shall be saved.
