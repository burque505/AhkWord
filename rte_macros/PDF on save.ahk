#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Each time the document is saved, a PDF of it is made beside it (report.docx -> report.pdf), laid out as it prints.
; The PDF goes through Microsoft Print to PDF; your default printer stays as it is.
; (Iteration 56) An event hook: AhkWord runs it on each save (Tools > Event Hooks, Saved), so nothing keeps running.
; Run it yourself (Tools > Macros, or a hotkey from Macro Hotkeys) to switch PDF on save on or off.

if A_Args.Length >= 1 && A_Args[1] = "Saved" {
    ; run by AhkWord after a save: A_Args[2] is the file
    SavedFile := A_Args.Length >= 2 ? A_Args[2] : ""
    if SavedFile = ""
        ExitApp
    SplitPath(SavedFile, , &Folder, &Ext, &Name)
    if Ext = "pdf"
        ExitApp
    try {
        Rte.ExportPdf(Folder "\" Name ".pdf")
        Rte.SetStatus("Saved, and made " Name ".pdf beside it.")
    } catch as e
        TrayTip(e.Message, "PDF on save failed", "Iconx")
    ExitApp
}

; run from Tools > Macros: PDF on save on or off
Hooked := Rte.EventHook("Saved")
if Hooked = "" || Hooked = A_ScriptName {
    if Hooked = "" {
        Rte.SetEventHook("Saved", A_ScriptName)
        On := true
    } else
        On := !Rte.EventHookOn("Saved")
    Rte.SetEventHookOn("Saved", On)
    Rte.SetStatus(On ? "PDF on save is on: each save also makes a PDF beside the document (run PDF on save again to switch it off)."
        : "PDF on save is off (run it again, or tick On in Tools > Event Hooks, to switch it on).")
} else if MsgBox("Saving already runs another macro, " Hooked " (Tools > Event Hooks, Saved), and an event runs one "
    . "macro.`n`nRun PDF on save on each save instead?", "PDF on save", "YesNo Icon?") = "Yes" {
    Rte.SetEventHook("Saved", A_ScriptName)
    Rte.SetEventHookOn("Saved", true)
    Rte.SetStatus("PDF on save is on, instead of " Hooked ".")
}
