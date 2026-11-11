#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Writes the table at the caret back to the file it came from (an Excel workbook, CSV, TSV or JSON) after you edit it
; in AhkWord. The first time, it asks for the file, and the sheet and range for a workbook, and remembers them for
; this table of this document (in table-links.ini); hold Shift while starting it to choose again.
; A workbook gets only the cells whose value changed; it must be closed in Excel. See also "Table from source file".
Title := "Table to source file"
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
    Path := FileSelect("S16", , Title ": write table " TableNo " to", "Data (*.xlsx; *.xlsm; *.csv; *.tsv; *.txt; *.json)")
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
    Parts := [Path, Sheet, Range]
    if Doc != ""
        DirCreate(RegExReplace(Links, "\\[^\\]+$")), IniWrite(Path "|" Sheet "|" Range, Links, Doc, Key)
}
Path := Parts[1], Sheet := Parts.Length > 1 ? Parts[2] : "", Range := Parts.Length > 2 ? Parts[3] : ""
try
    Count := Rte.SaveTableData(Path, TableNo, Sheet, Range)
catch as e {
    MsgBox("The table was not written:`n" e.Message, Title, "Icon!")
    ExitApp
}
Done := RegExMatch(Path, "i)\.xls[xm]$") ? (Count = 0 ? "No cells had changed." : Count " cell" (Count = 1 ? "" : "s") " changed.")
    : Count " row" (Count = 1 ? "" : "s") " written."
MsgBox("Table " TableNo " → " Path (Sheet != "" ? " (" Sheet (Range != "" ? ", " Range : "") ")" : "") "`n`n" Done, Title, "Iconi T5")
