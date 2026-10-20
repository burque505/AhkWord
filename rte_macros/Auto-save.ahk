#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Saves the document 3 seconds after you stop typing (needs AhkWord.ahk). Only documents that already have a file
; name are saved, so save a new document once yourself. Stops when the editor closes; to stop it earlier, exit it
; from its tray icon.
; The save uses Rte.Post: this script queues it and does not wait; a failure raises the AsyncError event.
Delay := 3000
Rte.On("TextChanged", (*) => SetTimer(SaveNow, -Delay))
Rte.On("AsyncError", (*) => TrayTip(Rte.LastAsyncError, "Auto-save failed", "Iconx"))
Rte.On("Closed", (*) => ExitApp())
Rte.SetStatus("Auto-save is on: the document is saved " Delay // 1000 " s after you stop typing.")

SaveNow() {
    try {
        if Rte.FileName != "" && Rte.Modified
            Rte.Post("Save")
    }
}
