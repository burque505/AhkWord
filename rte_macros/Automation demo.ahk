#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Builds a sample document by calling the editor from AHK, and reports how many calls it made and how long they took.
if Rte.Modified && MsgBox("The automation demo replaces the current document.`n`nDiscard your unsaved changes?", "Automation demo", "YesNo Icon!") != "Yes"
    ExitApp

Start := A_TickCount, Calls := 0
Do(Name, Args*) {    ; Rte.Call(Name, Args*) is Rte.<Name>(Args*), counted
    global Calls
    Calls++
    return Rte.Call(Name, Args*)
}
Prop(Name) {
    global Calls
    Calls++
    return Rte.%Name%
}

Do("New")
Rte.Html := "<h1>Automation demo</h1>"
    . "<p>This document was written by an AutoHotkey script through RteClient.ahk. "
    . "Every word below was formatted by a separate call.</p>"
    . "<p>Important: bold, italic, wavy, colored and highlighted.</p>"
    . "<p>Shopping list</p><p>Apples</p><p>Pears</p><p>Green</p><p>Red</p><p>Bread</p>"
Calls++

Format1(Word, Actions*) {
    if Do("Find", Word, 1, 1) < 0
        throw Error("Not found: " Word)
    for Action in Actions
        Do(Action*)
}
Format1("Important", ["ToggleBold"], ["SetFontColor", "#C00000"])
Format1("bold", ["ToggleBold"])
Format1("italic", ["ToggleItalic"])
Format1("wavy", ["ToggleUnderline", "Wavy"])
Format1("colored", ["SetFontColor", "SteelBlue"])
Format1("highlighted", ["SetBackgroundColor", "Yellow"])

; A numbered list with a nested level (Nest = Tab), under a heading
Format1("Shopping list", ["SetBlockStyle", "Heading2"])
Do("Find", "Apples", 1, 1)
Start1 := Prop("SelectionStart")
Do("Find", "Bread", 1, 1)
Do("Select", Start1, Prop("SelectionStart") + Prop("SelectionLength") - Start1)
Do("ToggleList", "Decimal")
Do("Find", "Green", 1, 1)
Start2 := Prop("SelectionStart")
Do("Find", "Red", 1, 1)
Do("Select", Start2, Prop("SelectionStart") + 3 - Start2)
Do("Nest", 1)

Do("GoTo", -1)
Do("SetStatus", Format("Automation demo: {} calls in {} ms.", Calls + 1, A_TickCount - Start))
