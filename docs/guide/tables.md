# Tables

**Insert > Table** inserts a table: choose the rows and columns, and whether the first row is a bold header.

While the caret is in a table, the **Table Tools** tab appears:

- **Insert** rows above or below and columns left or right; **delete** rows, columns or the whole table.
- **Properties** (also on the right-click menu): column widths, the table's alignment on the page, cell padding and row height for the whole table; borders, shading and vertical alignment for the selected cells or the whole table.
- **Borders** (none, all, outside, inside ..., with line width and colour) and **Shading** for the selected cells.

**Tab** moves to the next cell and **Shift+Tab** to the previous one; Tab in the last cell adds a row, as in Word. This works in tables from Word and other programs too.

Table formatting is kept in RTF, Word, OpenDocument and HTML files. Markdown tables have no widths, borders or shading (see [Markdown and HTML views](source-views.md)). Two limits: borders are single lines (Word's double and dotted borders show as single lines of the same width and colour), and cells can't be merged or split, though tables from Word with merged cells open as they are.

!!! tip "From a script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: see the [API reference](../automation/api.md#tables).
