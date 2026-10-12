#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Title-cases the selection, keeping its formatting (bold, colors, links...).
if Rte.SelectionLength = 0 {
    Rte.SetStatus("Select some text first.")
    ExitApp
}
Rte.ChangeCase("title")
Rte.SetStatus("Selection converted to Title Case.")
