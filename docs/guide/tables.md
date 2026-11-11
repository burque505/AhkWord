# Tables

**Insert > Table** inserts a table: choose the rows and columns, and whether the first row is a bold header. In an empty document an empty paragraph stays above the table, ready for a title or other text; delete it if the table is to come first.

## A table from a file

**From File...** in the Insert Table dialog makes the table from an Excel workbook (.xlsx, .xlsm), a CSV or TSV file, or a JSON file. Excel doesn't need to be installed, and the workbook may be open in Excel while you do it (its last saved state is read). For a workbook, choose the sheet and, if you like, a range such as `A1:D20`; without one the cells in use are taken, empty rows and columns around them left out. Values come as Excel shows them: number formats, currency, percentages and dates, in your computer's regional format; a formula gives its last result. The columns are made as wide as their text needs. CSV files may use commas, semicolons (as Excel saves them where the decimal mark is a comma) or tabs, in UTF-8 or Windows' own encoding.

The table is an ordinary table afterwards: it isn't linked to the file. A script can fill it again from the file, keeping its formats (see below).

While the caret is in a table, the **Table Tools** tab appears:

- **Insert** rows above or below and columns left or right; **delete** rows, columns or the whole table.
- **Properties** (also on the right-click menu): column widths, the table's alignment on the page, cell padding and row height for the whole table; borders, shading and vertical alignment for the selected cells or the whole table.
- **Borders** (none, all, outside, inside ..., with line width and colour) and **Shading** for the selected cells.

**Tab** moves to the next cell and **Shift+Tab** to the previous one; Tab in the last cell adds a row, as in Word. This works in tables from Word and other programs too.

Table formatting is kept in RTF, Word, OpenDocument and HTML files. Markdown tables have no widths, borders or shading (see [Markdown and HTML views](source-views.md)). Two limits: borders are single lines (Word's double and dotted borders show as single lines of the same width and colour), and cells can't be merged or split, though tables from Word with merged cells open as they are.

!!! tip "From a script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: see the [API reference](../automation/api.md#tables).

!!! tip "Data from a script"
    `Rte.InsertTableFromFile("C:\Data\clients.xlsx", 1, "Clients")` and `Rte.InsertTableFromData(csvOrJson)` make a table from data. `Rte.SetTableData(data, "", 1, 2)` fills the document's first table again, keeping its header row and formats: rows are added or deleted to fit, and JSON objects go in the columns their names head. `Rte.TableData(1, "json")` gives a table's cells back as CSV, TSV or JSON, and `Rte.CellText(2, 3)` / `Rte.SetCellText(2, 3, "text")` read and change one cell. See the [API reference](../automation/api.md#tables-from-data-and-data-from-tables).
