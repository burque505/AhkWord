#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; A small always-on-top window showing progress towards a word goal, updated as you type (needs rte_writer2.ahk).
Answer := InputBox("Word goal for this document:", "Word goal", "w240 h110", 500)
if Answer.Result != "OK" || !IsInteger(Answer.Value) || Answer.Value < 1
    ExitApp
Goal := Integer(Answer.Value)

Tracker := Gui("+AlwaysOnTop +ToolWindow", "Word goal")
Tracker.SetFont("s9", "Segoe UI")
Bar := Tracker.AddProgress("w240 h16 Range0-" Goal)
Label := Tracker.AddText("w240 Center")
Tracker.OnEvent("Close", (*) => (Rte.Off(), ExitApp()))

UpdateGoal(*) {
    Words := Rte.WordCount()
    Bar.Value := Min(Words, Goal)
    Label.Text := Words " of " Goal " words" (Words >= Goal ? "  -  goal reached!" : "")
}
Rte.On("TextChanged", UpdateGoal)
Rte.On("Opened", UpdateGoal)
Rte.On("Closed", (*) => ExitApp())
UpdateGoal()
Tracker.Show("NoActivate")
