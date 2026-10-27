#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
; Tidies the selected text, as text pasted from a PDF, an e-mail or a web page often needs: lines broken in the middle
; of sentences are joined again (an empty line still separates paragraphs), runs of spaces and tabs become one space,
; "straight" quotes become curly ones, and -- becomes a dash. Select the text first (Ctrl+A for all of it). The
; selection's formatting is not kept, so use it on plain pasted text. Ctrl+Z undoes it.
if Rte.SelectionLength = 0 {
    Rte.SetStatus("Select the pasted text first (or Ctrl+A for all of it), then run Clean up pasted text again.")
    ExitApp
}
Text := StrReplace(StrReplace(Rte.SelectedText, "`r`n", "`r"), "`n", "`r")
; (iteration 57) footnotes, cross-references and a table of contents keep hidden markers in the text ([[fn]], [[xref:...]],
; [[toc=...]]): put back as plain text, they would be no markers, shown as text, and the notes, references and links gone
if RegExMatch(Text, "\[\[(?:/?xref[:=\]]|/?tocs?[:=\]]|pg\]\]|/?toc\]\]|fn[\]:])") {
    Rte.SetStatus("Clean up pasted text: the selection has footnotes, cross-references or a table of contents, which it "
        . "would take apart. Select only the pasted text, then run it again.")
    ExitApp
}
Before := Text
Text := RegExReplace(Text, "[ \t]+", " ")                      ; runs of spaces and tabs
Text := RegExReplace(Text, " ?`r ?", "`r")                      ; spaces at the ends of lines
Text := RegExReplace(Text, "(?<=\S)-`r(?=\S)", "-")             ; a word split at a line's end stays one (hyphenated)
Text := RegExReplace(Text, "(?<=[^`r])`r(?=[^`r])", " ")        ; a single line break inside a paragraph
Text := RegExReplace(Text, "`r{3,}", "`r`r")                    ; at most one empty line
Text := StrReplace(Text, "--", "—")
Text := RegExReplace(Text, "(^|[\s(\[{`r])`"", "$1“")           ; opening double quotes
Text := StrReplace(Text, "`"", "”")                             ; the rest are closing ones
Text := RegExReplace(Text, "(^|[\s(\[{`r])'", "$1‘")            ; opening single quotes
Text := StrReplace(Text, "'", "’")                              ; apostrophes and closing single quotes
if Text == Before {
    Rte.SetStatus("Clean up pasted text: the selection was tidy already.")
    ExitApp
}
Rte.InsertText(Text)
Rte.SetStatus("Clean up pasted text: lines joined, spaces tidied, quotes and dashes made typographic (Ctrl+Z undoes it).")
