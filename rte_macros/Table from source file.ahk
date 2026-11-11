#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Refreshes the table at the caret from its file (an Excel workbook, CSV, TSV or JSON), keeping the table's formats:
; rows are added or deleted to fit. Uses the file "Table to source file" remembered for this table (table-links.ini),
; or asks for one and remembers it; hold Shift while starting it to choose again.
Title := "Table from source file"
TableNo := Rte.TableNumberAtCaret()
if !TableNo {
    MsgBox("Put the caret in a table first.", Title, "Icon!")
    ExitApp
}
Doc := Rte.FileName
Links := (Rte.Portable ? A_ScriptDir "\..\data" : A_AppData "\AhkWord") "\table-links.ini"
Key := "table " TableNo
Link := Doc != "" && !GetKeyState("Shift") ? IniRead(Links, Doc, Key, "") : ""
Parts := StrSplit(Link, "|")
if Link = "" {
    Path := FileSelect("1", , Title ": fill table " TableNo " from", "Data (*.xlsx; *.xlsm; *.csv; *.tsv; *.txt; *.json)")
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
    Parts := [Path, Sheet, Range]
    if Doc != ""
        DirCreate(RegExReplace(Links, "\\[^\\]+$")), IniWrite(Path "|" Sheet "|" Range, Links, Doc, Key)
}
Path := Parts[1], Sheet := Parts.Length > 1 ? Parts[2] : "", Range := Parts.Length > 2 ? Parts[3] : ""
try
    Count := Rte.SetTableDataFromFile(Path, TableNo, 1, Sheet, Range)
catch as e {
    MsgBox("The table was not filled:`n" e.Message, Title, "Icon!")
    ExitApp
}
MsgBox("Table " TableNo " ← " Path "`n`n" Count " row" (Count = 1 ? "" : "s") " filled.", Title, "Iconi T5")
