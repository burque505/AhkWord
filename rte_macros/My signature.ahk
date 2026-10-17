#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Inserts your signature at the caret: your name, title, phone and e-mail, one a line (unlike the Signature Block
; template part, which asks each time, this keeps your details). The first time, it asks
; for them and keeps them in "My signature.ini" beside this script; hold Ctrl while starting it to change them.
; Give it a hotkey in Tools > Macro Hotkeys (Ctrl+Alt+S, say) to sign with one key.
Ini := A_ScriptDir "\My signature.ini"
Fields := ["Name", "Title", "Phone", "E-mail"]
if !FileExist(Ini) || GetKeyState("Ctrl") {
    for Field in Fields {
        Answer := InputBox("Your " StrLower(Field) " for the signature block (leave it empty to leave the line out):",
            "My signature", "w420 h130", IniRead(Ini, "Signature", Field, ""))
        if Answer.Result = "Cancel"
            ExitApp
        IniWrite(Answer.Value, Ini, "Signature", Field)
    }
}
Lines := []
for Field in Fields
    if (Value := IniRead(Ini, "Signature", Field, "")) != ""
        Lines.Push(Value)
if Lines.Length = 0 {
    Rte.SetStatus("My signature: no details yet. Hold Ctrl while starting it to enter them.")
    ExitApp
}
Block := ""
for Line in Lines
    Block .= (A_Index > 1 ? "`r" : "") Line
Rte.InsertText(Block)
Rte.SetStatus("Your signature inserted. Hold Ctrl while starting the macro to change it.")
