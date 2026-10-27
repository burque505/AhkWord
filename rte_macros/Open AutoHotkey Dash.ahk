#Requires AutoHotkey v2.0
; Opens AutoHotkey's Dash: its help, Window Spy, and the settings for how .ahk files are opened and edited (Editor
; settings: the program AhkWord's Edit opens a macro in). The Dash comes with AutoHotkey when it is installed; AhkWord's
; own download carries AutoHotkey without it.

; where AutoHotkey is installed (the installer's InstallDir; else its usual place)
Dir := ""
for Root in ["HKLM\SOFTWARE\AutoHotkey", "HKCU\SOFTWARE\AutoHotkey"]
    if Dir = ""
        try Dir := RegRead(Root, "InstallDir")
if Dir = ""
    Dir := A_ProgramFiles "\AutoHotkey"
Dash := Dir "\UX\ui-dash.ahk", UX := Dir "\UX\AutoHotkeyUX.exe"

if FileExist(Dash) {
    ; as the Start menu's AutoHotkey Dash opens it
    if FileExist(UX)
        Run('"' UX '" "' Dash '"')
    else
        Run('"' Dash '"')
    ExitApp
}
if MsgBox("The AutoHotkey Dash comes with AutoHotkey, which isn't installed on this computer (AhkWord carries its own "
    . "copy, without the Dash). You don't need it to use AhkWord or its macros.`n`nOpen autohotkey.com to download "
    . "AutoHotkey?", "Open AutoHotkey Dash", "YesNo Iconi") = "Yes"
    Run("https://www.autohotkey.com/")
