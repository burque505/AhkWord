#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Shows the editor's events as they happen (needs AhkWord.ahk). Close this window, or the editor, to stop.
Monitor := Gui("+AlwaysOnTop +Resize +MinSize300x150", "AhkWord events")
Monitor.SetFont("s9", "Segoe UI")
List := Monitor.AddListView("w560 r14", ["Time", "Event", "Details"])
List.ModifyCol(1, 70), List.ModifyCol(2, 120), List.ModifyCol(3, 340)
Monitor.OnEvent("Size", (G, MinMax, W, H) => MinMax != -1 && List.Move(, , W - 20, H - 20))
Monitor.OnEvent("Close", (*) => (Rte.Off(), ExitApp()))

for Name in StrSplit(Rte.EventNames, "`n")
    Rte.On(Name, ShowEvent)
Monitor.Show("NoActivate")

ShowEvent(e) {
    switch e.Name {
        case "TextChanged": Details := e.TextLength " characters"
        case "SelectionChanged": Details := "start " e.Start (e.Length ? ", " e.Length " selected" : "")
        case "ModifiedChanged": Details := e.Modified ? "unsaved changes" : "no unsaved changes"
        case "Saved", "Opened": Details := (Path := Rte.FileName) != "" ? Path : "(new document)"
        case "AsyncError": Details := Rte.LastAsyncError
        case "Closed": Details := "the editor has closed; this window closes in 2 s"
        default: Details := ""
    }
    Row := List.Add(, FormatTime(, "HH:mm:ss"), e.Name, Details)
    List.Modify(Row, "Vis")
    if e.Name = "Closed"
        SetTimer(() => ExitApp(), -2000)
}
