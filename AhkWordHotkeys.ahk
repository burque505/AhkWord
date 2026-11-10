; AhkWord's hotkeys, its tray notice on saving, and the snippets as hotstrings: everything of AhkWord that runs in
; AutoHotkey beside the editor. Included by AhkWord.ahk (the editor in the same process) and by AhkWordCompanion.ahk
; (VS build trial: started by AhkWord.exe, the editor in another process). Rte (RteClient2.ahk) finds the editor in both
; cases; AhkWordEditorPid() is the editor's process, defined by the script that includes this file.
;
; Hotkeys defined here keep working while the editor is open (in AhkWord.ahk they call it directly, in-process):
#HotIf WinActive("AhkWord ahk_pid " AhkWordEditorPid())    ; (each window's own: several can be open)
^!d::Rte.InsertText(FormatTime(, "d MMMM yyyy"))                           ; Ctrl+Alt+D: insert today's date
^!w::Rte.SetStatus(Rte.WordCount() " words, " Rte.CharacterCount() " characters")   ; Ctrl+Alt+W: counts in the status bar
#HotIf

; Events work in-process too: a tray notification whenever the document is saved. The subscription is made from a
; timer, which runs as soon as the editor is up (this script's main thread is busy running it). A save made on closing
; (Save changes? Yes) reaches the callback after the window has gone, when Rte can't be asked any more: then the
; notification has no file name.
SetTimer(SubscribeToSaved, -100)
SubscribeToSaved() {
    try Rte.On("Saved", SavedNotice)
    catch TargetError, TimeoutError    ; the editor is not up yet (or still busy starting)
        SetTimer(SubscribeToSaved, -200)
}
SavedNotice(e) {
    try FileName := Rte.FileName
    catch TargetError    ; (closed already)
        FileName := "AhkWord"
    TrayTip(FileName, "AhkWord: saved", "Iconi Mute")
}

; Iteration 55: the snippets (Insert > Snippets) as hotstrings. Typed in this AhkWord's document, an abbreviation (then
; Space, Enter, Tab or punctuation, or at once) becomes its snippet: the editor takes out what was typed and puts the
; snippet's part in (FireSnippet), so the hotstrings send no keys (B0). They fire only while this AhkWord's own window
; is active with the document (or the Markdown or HTML view) in focus: each AhkWord window is a process of its own, and
; only the active one's fire. Set up once the editor is up, and again when the snippets' file changes (from this
; window, another, or a script).
SetTimer(SnippetsStart, -400)
SnippetsStart() {
    global SnippetsFile
    try SnippetsFile := Rte.SnippetsFile
    catch TargetError, TimeoutError {    ; the editor is not up yet (or still busy starting)
        SetTimer(SnippetsStart, -300)
        return
    }
    SetTimer(SnippetsWatch, 1500)
    SnippetsWatch()
}
SnippetsWatch() {
    static Last := "-"
    Now := FileExist(SnippetsFile) ? FileGetTime(SnippetsFile) "|" FileGetSize(SnippetsFile) : ""
    if Now == Last
        return
    Last := Now
    SnippetsLoad()
}
SnippetsLoad() {
    static On := []
    try Lines := Rte.SnippetHotstrings
    catch
        return
    HotIf(SnippetsHere)
    for Key in On
        try Hotstring(Key, , "Off")
    On := []
    Loop Parse Lines, "`n" {
        Tab := InStr(A_LoopField, "`t")
        if !Tab
            continue
        Abbreviation := SubStr(A_LoopField, Tab + 1)
        Key := ":" SubStr(A_LoopField, 1, Tab - 1) ":" Abbreviation
        try {
            Hotstring(Key, SnippetFire.Bind(Abbreviation), "On")
            On.Push(Key)
        }
    }
    HotIf()
}
; (this AhkWord's main window, not one of its dialogs, with the focus in the document or the source view)
SnippetsHere(*) {
    if !WinActive("AhkWord ahk_pid " AhkWordEditorPid())
        return false
    try return WinGetClass("ahk_id " ControlGetFocus("A")) ~= "i)RichEdit50W|Scintilla"
    return false
}
SnippetFire(Abbreviation, *) {
    try Rte.FireSnippet(Abbreviation, A_EndChar)
}
