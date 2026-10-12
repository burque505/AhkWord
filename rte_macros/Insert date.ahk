#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Inserts today's date at the caret.
Rte.InsertText(FormatTime(, "dddd, d MMMM yyyy"))
