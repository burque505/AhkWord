#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Fills the table at the caret again from the file it is linked to (an Excel workbook, CSV, TSV or JSON), keeping
; the table's formats, as Table Layout > Update Table does: rows are added or deleted to fit. For a table not linked
; yet, this asks for the file (and the sheet and range of a workbook) and links the table to it first
; (Rte.LinkTable). Hold Shift while starting it to link the table to another file. Rte.UpdateTables() updates every
; linked table of the document at once.
Title := "Table from source file"
if !Rte.TableNumberAtCaret() {
    MsgBox("Put the caret in a table first.", Title, "Icon!")
    ExitApp
}
if Rte.TableSource() = "" || GetKeyState("Shift") {
    Path := FileSelect("1", , Title ": fill the table from", "Data (*.xlsx; *.xlsm; *.csv; *.tsv; *.txt; *.json)")
    if Path = ""
        ExitApp
    Sheet := "", Range := ""
    if RegExMatch(Path, "i)\.xls[xm]$") {
        Sheets := StrReplace(Rte.WorkbookSheets(Path), "`n", ", ")
        Answer := InputBox("Sheet (" Sheets "), blank for the first:", Title)
        if Answer.Result != "OK"
            ExitApp
        Sheet := Trim(Answer.Value)
        Answer := InputBox("Range, such as A1:D20, or blank for the cells in use:", Title)
        if Answer.Result != "OK"
            ExitApp
        Range := Trim(Answer.Value)
    }
    Rte.LinkTable(Path, 0, Sheet, Range)
}
try
    Count := Rte.UpdateTable()
catch as e {
    MsgBox("The table was not filled:`n" e.Message, Title, "Icon!")
    ExitApp
}
MsgBox("The table ← " StrSplit(Rte.TableSource(), "|")[1] "`n`n" Count " row" (Count = 1 ? "" : "s") " filled.", Title, "Iconi T5")
