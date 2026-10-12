#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Shows word, character and paragraph counts, and the current file.
Paragraphs := 0
Loop Parse Rte.Text, "`n"
    if Trim(A_LoopField) != ""
        Paragraphs++
Path := Rte.FileName
MsgBox(Format("Words:`t{}`nCharacters:`t{}`nParagraphs:`t{}`n`nFile:`t{}"
    , Rte.WordCount(), Rte.CharacterCount(), Paragraphs, Path != "" ? Path : "(not saved yet)"), "Document statistics", "Iconi")
