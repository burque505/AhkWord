#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Collects text from any program into the document: select something in a browser, a PDF reader, an e-mail ... and press
; Win+Shift+C; it is added at the end of the document, with where it came from and when. The editor doesn't need to be
; in front. Runs until the editor closes, or until you exit it from its tray icon.
#+c:: {
    Saved := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(1) {
        A_Clipboard := Saved
        TrayTip("Nothing was selected to collect.", "Collect from anywhere", "Iconi")
        return
    }
    Text := Trim(StrReplace(StrReplace(A_Clipboard, "`r`n", "`r"), "`n", "`r"), " `t`r")
    Source := WinGetTitle("A")
    A_Clipboard := Saved
    try {
        Rte.GoTo(-1)
        Rte.InsertText((Rte.TextLength > 0 ? "`r`r" : "") Text "`r— " Source ", " FormatTime(, "ShortDate") " " FormatTime(, "Time"))
        TrayTip(SubStr(Text, 1, 80) (StrLen(Text) > 80 ? "…" : ""), "Collected into the document", "Iconi")
    } catch as e
        TrayTip(e.Message, "Collect from anywhere", "Iconx")
}
Rte.On("Closed", (*) => ExitApp())
Rte.SetStatus("Collect from anywhere is on: select text in any program and press Win+Shift+C to add it to the end of this document.")
Persistent()
