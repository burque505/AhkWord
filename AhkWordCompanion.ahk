#Requires AutoHotkey v2.0 64-bit
#SingleInstance Off    ; (one for each AhkWord.exe window)
#Include "RteClient2.ahk"
; Iteration 88: AhkWord.exe starts this script beside itself, with its process id:
;   AutoHotkey64.exe AhkWordCompanion.ahk <pid>
; It gives that window what AhkWord.ahk gives the editor it runs in-process: the hotkeys, the tray notice on saving and
; the snippets as hotstrings (AhkWordHotkeys.ahk). Its calls go to that window's automation server only, and it ends
; when that window's process does.

if A_Args.Length < 1 || !(A_Args[1] ~= "^\d+$") || !ProcessExist(A_Args[1]) {
    MsgBox("AhkWordCompanion.ahk is started by AhkWord.exe, with its process id.", "AhkWord", "Iconi")
    ExitApp
}
EditorPid := A_Args[1] + 0
AhkWordEditorPid() => EditorPid

; This window's automation server ("AhkWord automation server <pid>"), so that Rte never talks to another AhkWord
; window (Rte takes RTE_WRITER_HWND first). The editor makes it while starting: wait for it.
CompanionServer() {
    Hidden := A_DetectHiddenWindows
    A_DetectHiddenWindows := true
    try return WinWait("AhkWord automation server " EditorPid " ahk_pid " EditorPid, , 60)
    finally A_DetectHiddenWindows := Hidden
}
if !(Server := CompanionServer())
    ExitApp
EnvSet("RTE_WRITER_HWND", Server)
; The server is made while the editor is still being built: ask nothing before its window is shown (a request in the
; middle of that waits until the editor is free, and may run out of time).
if !WinWait("AhkWord ahk_pid " EditorPid, , 120)    ; (hidden windows not counted: the server is one)
    ExitApp

; Gone with the editor (also when it stops on an error).
SetTimer(CompanionWatch, 1000)
CompanionWatch() {
    if !ProcessExist(EditorPid)
        ExitApp
}

if FileExist(A_ScriptDir "\images\icon\ahkword.ico")
    TraySetIcon(A_ScriptDir "\images\icon\ahkword.ico")
A_IconTip := "AhkWord (hotkeys and snippets)"
Persistent()

#Include "AhkWordHotkeys.ahk"
