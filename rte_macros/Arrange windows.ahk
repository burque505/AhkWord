#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Arranges the open AhkWord windows side by side on the monitor of the active one
; (up to 3 in columns, more in a grid), each in Print Layout with its page fitted to its width.

A_TitleMatchMode := "RegEx"
Windows := WinGetList(" - AhkWord$")          ; front to back: the active one first
if Windows.Length = 0
    ExitApp
; the monitor the front window is on, without the taskbar
WinGetPos(&WX, &WY, &WW, &WH, Windows[1])
Mon := MonitorFromPoint(WX + WW // 2, WY + WH // 2)
MonitorGetWorkArea(Mon, &Left, &Top, &Right, &Bottom)
N := Windows.Length
Cols := N <= 3 ? N : Ceil(Sqrt(N)), Rows := Ceil(N / Cols)
W := (Right - Left) // Cols, H := (Bottom - Top) // Rows
for i, Hwnd in Windows {
    ; maximized first (a minimized or off-screen window comes back), then restored, as WinMove needs
    WinMaximize(Hwnd)
    WinRestore(Hwnd)
    Col := Mod(i - 1, Cols), Row := (i - 1) // Cols
    MoveVisible(Hwnd, Left + Col * W, Top + Row * H, W, H)
}
; each window's editor: Print Layout, the page's width fitted to the window (each window keeps its own zoom)
for Hwnd in Windows {
    if !(Server := ServerOf(Hwnd))
        continue
    EnvSet("RTE_WRITER_HWND", Server)
    try {
        Rte.PrintLayout := true
        Rte.ZoomMode := "PageWidth"
    }
}
EnvSet("RTE_WRITER_HWND")
WinActivate(Windows[1])

; WinMove by the visible frame: Windows 10/11 windows have invisible borders around them.
MoveVisible(Hwnd, X, Y, W, H) {
    WinMove(X, Y, W, H, Hwnd)
    WinGetPos(&OX, &OY, &OW, &OH, Hwnd)
    Rect := Buffer(16)
    if DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", Hwnd, "UInt", 9, "Ptr", Rect, "UInt", 16) != 0
        return                                    ; (9: DWMWA_EXTENDED_FRAME_BOUNDS)
    L := NumGet(Rect, 0, "Int"), T := NumGet(Rect, 4, "Int"), R := NumGet(Rect, 8, "Int"), B := NumGet(Rect, 12, "Int")
    WinMove(X - (L - OX), Y - (T - OY), W + (OW - (R - L)), H + (OH - (B - T)), Hwnd)
}

; The hidden automation window of the AhkWord that owns a window (0: none).
ServerOf(Hwnd) {
    Hidden := A_DetectHiddenWindows
    A_DetectHiddenWindows := true
    try return WinExist("^AhkWord automation server " WinGetPID(Hwnd) "$")
    finally A_DetectHiddenWindows := Hidden
}

MonitorFromPoint(X, Y) {
    loop MonitorGetCount() {
        MonitorGet(A_Index, &ML, &MT, &MR, &MB)
        if X >= ML && X < MR && Y >= MT && Y < MB
            return A_Index
    }
    return MonitorGetPrimary()
}
