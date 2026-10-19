#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Shows the page view's state (View > Show Pages) for a bug report: where the printed pages start, each gap between
; the sheets, and any raised text that isn't a gap's. Positions and sizes only: none of the document's text.
Report := Rte.PageViewReport()
G := Gui("+Resize", "Page view report")
G.SetFont("s9", "Consolas")
Box := G.AddEdit("ReadOnly -Wrap HScroll VScroll w900 h420", Report)
G.AddButton("Default w80", "Close").OnEvent("Click", (*) => ExitApp())
G.OnEvent("Close", (*) => ExitApp())
G.OnEvent("Size", (g, minMax, w, h) => minMax = -1 ? 0 : Box.Move(, , w - 20, h - 50))
G.Show()
