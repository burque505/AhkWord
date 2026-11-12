#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Writes the table at the caret back to the file it is linked to (an Excel workbook, CSV, TSV or JSON), as
; Table Layout > Save to File does. A table inserted from a file is linked already; for another table this asks for
; the file (and the sheet and range of a workbook) and links the table to it first (Rte.LinkTable): the link is kept
; in the document, with the table. A workbook gets only the cells whose value changed; it must be closed in Excel.
; Hold Shift while starting it to link the table to another file. See also "Table from source file".
Title := "Table to source file"
if !Rte.TableNumberAtCaret() {
    MsgBox("Put the caret in a table first.", Title, "Icon!")
    ExitApp
}
if Rte.TableSource() = "" || GetKeyState("Shift") {
    Path := FileSelect("S16", , Title ": write the table to", "Data (*.xlsx; *.xlsm; *.csv; *.tsv; *.txt; *.json)")
    if Path = ""
        ExitApp
    if !RegExMatch(Path, "i)\.(xlsx|xlsm|csv|tsv|txt|json)$")
        Path .= ".csv"
    Sheet := "", Range := ""
    if RegExMatch(Path, "i)\.xls[xm]$") {
        try Sheets := StrReplace(Rte.WorkbookSheets(Path), "`n", ", ")
        catch {
            MsgBox("There is no such workbook: save one in Excel first.", Title, "Icon!")
            ExitApp
        }
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
    Warning := Rte.SaveTableCheck()
catch as e {
    MsgBox("The table was not written:`n" e.Message, Title, "Icon!")
    ExitApp
}
; past the link's range, or over a sheet's rows below the data: ask first
if Warning != "" && MsgBox(Warning "`n`nWrite the table to the file anyway?", Title, "YesNo Icon! Default2") != "Yes"
    ExitApp
try
    Count := Rte.SaveTableToSource(0, Warning != "")
catch as e {
    MsgBox("The table was not written:`n" e.Message, Title, "Icon!")
    ExitApp
}
Link := StrSplit(Rte.TableSource(), "|")    ; (after: a range the table went past has grown)
Done := RegExMatch(Link[1], "i)\.xls[xm]$") ? (Count = 0 ? "No cells had changed." : Count " cell" (Count = 1 ? "" : "s") " changed.")
    : Count " row" (Count = 1 ? "" : "s") " written."
MsgBox("The table → " Link[1] (Link[2] != "" ? " (" Link[2] (Link[3] != "" ? ", " Link[3] : "") ")" : "") "`n`n" Done, Title, "Iconi T5")
