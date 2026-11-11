# Tables

**Insert > Table** inserts a table: choose the rows and columns, and whether the first row is a bold header. In an empty document an empty paragraph stays above the table, ready for a title or other text; delete it if the table is to come first.

## A table from a file

**From File...** in the Insert Table dialog makes the table from an Excel workbook (.xlsx, .xlsm), a CSV or TSV file, or a JSON file. Excel doesn't need to be installed, and the workbook may be open in Excel while you do it (its last saved state is read). For a workbook, choose the sheet and, if you like, a range such as `A1:D20`; without one the cells in use are taken, empty rows and columns around them left out. Values come as Excel shows them: number formats, currency, percentages and dates, in your computer's regional format; a formula gives its last result. The columns are made as wide as their text needs. CSV files may use commas, semicolons (as Excel saves them where the decimal mark is a comma) or tabs, in UTF-8 or Windows' own encoding.

The table is an ordinary table afterwards: it isn't linked to the file. A script can fill it again from the file, keeping its formats (see below).

While the caret is in a table, the **Table Tools > Table Layout** tab appears:

- **Insert** rows above or below and columns left or right; **delete** rows, columns or the whole table.
- **Merge Cells** and **Split Cells**, and **Delete Cells** (see below).
- **Properties** (also on the right-click menu): column widths, the table's alignment on the page, cell padding and row height for the whole table; borders, shading and vertical alignment for the selected cells or the whole table.
- **Borders** (none, all, outside, inside ..., with line width and colour) and **Shading** for the selected cells.

**Tab** moves to the next cell and **Shift+Tab** to the previous one; Tab in the last cell adds a row, as in Word. This works in tables from Word and other programs too.

Table formatting is kept in RTF, Word, OpenDocument and HTML files. Markdown tables have no widths, borders or shading (see [Markdown and HTML views](source-views.md)). One limit: borders are single lines (Word's double and dotted borders show as single lines of the same width and colour). See also the [known limitations](../reference/limitations.md).

## Merging, splitting and deleting cells

To work on several cells, drag over them with the mouse: from a cell in one row to a cell in another, AhkWord selects the cells between them (shaded as a selection), not the whole rows. Within one row, select from the first cell's text to the last's. From the keyboard, **Shift+Up** and **Shift+Down** in a cell select the cells above or below it as a block, and then **Shift+Left** and **Shift+Right** add or take away a column.

- **Merge Cells** makes the selected cells one cell. Their text is kept, each cell's as paragraphs of its own, as in Word. Cells merged across columns make one wider cell; cells merged down rows show as one tall cell.
- **Split Cells...** asks how many columns and rows. A cell is split into columns of equal width. A cell merged down splits back into its rows (or into a number of rows that divides them); a cell of one row split into rows gets new rows, the row's other cells merged down over them, so that they stay whole. Several selected cells are merged first.
- **Delete Cells...** asks what becomes of the cells around them: the cells after them in their rows move left, the cells below them move up (their text; empty cells come at the bottom), or their whole rows or columns go. With cells selected, **Delete Row** and **Delete Column** delete all of their rows or columns.
- **Delete** (the key) on selected cells empties them.
- **Copy** (Ctrl+C) on selected cells copies them as a table of their own (paste it into AhkWord or Word), and as tab-separated text (Excel puts it in cells). **Cut** copies them and empties them.

Columns are found by where they are, not by counting cells: deleting or inserting a column that a merged cell spans makes that cell narrower or wider instead of deleting it. Each of these is one step for **Undo**.

!!! tip "From a script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: see the [API reference](../automation/api.md#tables).
    Cells: `Rte.SelectCells(3, 2, 4, 2)` (row 3, cell 2 to row 4, cell 2), then `Rte.MergeCells()`, `Rte.SplitCells(2, 1)`, `Rte.DeleteCells("up")` or `Rte.ClearCells()`.

!!! tip "Data from a script"
    `Rte.InsertTableFromFile("C:\Data\clients.xlsx", 1, "Clients")` and `Rte.InsertTableFromData(csvOrJson)` make a table from data. `Rte.SetTableData(data, "", 1, 2)` fills the document's first table again, keeping its header row and formats: rows are added or deleted to fit, and JSON objects go in the columns their names head. `Rte.TableData(1, "json")` gives a table's cells back as CSV, TSV or JSON, and `Rte.CellText(2, 3)` / `Rte.SetCellText(2, 3, "text")` read and change one cell. See the [API reference](../automation/api.md#tables-from-data-and-data-from-tables).

!!! tip "Back to the file the data came from"
    `Rte.SaveTableData("C:\Data\clients.xlsx", 1, "Clients")` writes the document's first table back to the workbook: only the cells whose value changed (formulas and formats of the others are left alone), the cells the table no longer reaches emptied; the workbook must be closed in Excel. A .csv, .tsv or .json file is written whole, as it was (its delimiter; JSON objects stay objects). The macros **Table to source file** and **Table from source file** (Tools > Macros) do it for the table at the caret: they ask for the file the first time and remember it for that table of that document (hold Shift to choose another).
