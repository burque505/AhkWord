#Requires AutoHotkey v2.0 64-bit
#Include "Csharp-in-AHK.ahk"
#Include "RteClient2.ahk"
; RTE Writer, iteration 15: a word processor built on RichEditorNET (rte_writer15.cs), hosted in this AHK process.
; Adds to iteration 14:
;   - Pictures' alt text: shown when the mouse rests on a picture; right-click a picture > Alt text... to change it;
;     Insert > Image asks for it.
; Iteration 14 added to iteration 13:
;   - The header and footer are shown above and below the text, as they are printed on the first page; click one to
;     change them (View > Header and footer shows or hides them).
;   - Pictures keep their alt text in RTF (and so through Word and LibreOffice).
; Iteration 13 added to iteration 12:
;   - Headers and footers belong to each document (File > Header and Footer; the settings are the default for new
;     documents), and are saved in it: real headers and footers in RTF, Word and OpenDocument files, with page
;     numbers that Word and LibreOffice keep up to date, and the page size and margins of Page Setup.
; Iteration 12 added to iteration 11:
;   - Manual page breaks: Insert > Page break (Ctrl+Enter), shown in the editor as a dashed line. Printing and PDF
;     export start a new page there; they are kept in RTF, Word and OpenDocument files, and in HTML and Markdown as
;     <div style="page-break-after: always; break-after: page"></div>, which browsers obey when printing.
; Iteration 11 added to iteration 10:
;   - Exports through Word and LibreOffice (.docx, .odt) give legal lists real list numbering, which renumbers itself
;     when the document is edited in Word or LibreOffice (rte_office3.cs).
; Iteration 10 added to iteration 9:
;   - Legal numbering: Paragraph > Legal numbering (Ctrl+Shift+G, or the 1.1 button) numbers list items 1., 1.1, 1.1.1 by
;     level; Tab / Shift+Tab, Enter, Backspace and Delete keep the numbers right; Restart / Continue numbering work on
;     them; Paragraph > List levels sets their look. Saved in HTML, Markdown ([type=legal]), RTF and .docx.
; Iteration 9 added tables, RTF, Word and OpenDocument files (Word, LibreOffice, the built-in Open XML converter or
; pandoc), and the table right-click menu and shortcuts.
; Iteration 8 added Markdown formatting (HTML and M-ArkDown), Paragraph > List levels and Restart / Continue numbering.
; Earlier iterations added events, recent files, settings, drag and drop, printing, PDF export, headers and footers,
; spelling suggestions, the outline, macro recording and hotkeys, and real Markdown. Settings are kept in
; %APPDATA%\RTE Writer\settings.json (Rte.SettingsFile); set RTE_WRITER_SETTINGS to a file path to keep them elsewhere.
; Optionally pass a document to open:  rte_writer15.ahk "C:\path\document.md"
;
; AHK automation: other scripts #Include "RteClient2.ahk" and call Rte.<Name>(...), react to events with
; Rte.On(Event, Callback) and queue calls with Rte.Post(...). Help > Automation API lists everything that can be called,
; e.g. Rte.InsertPageBreak(), Rte.LegalNumbering(), Rte.LegalFormat := "All", Rte.InsertTable(3, 4), Rte.ExportDocx("C:\path\copy.docx").
;
; Hotkeys defined here keep working while the editor is open, and call it directly (in-process):
#HotIf WinActive("RTE Writer ahk_pid " ProcessExist())
^!d::Rte.InsertText(FormatTime(, "d MMMM yyyy"))                           ; Ctrl+Alt+D: insert today's date
^!w::Rte.SetStatus(Rte.WordCount() " words, " Rte.CharacterCount() " characters")   ; Ctrl+Alt+W: counts in the status bar
#HotIf

; Events work in-process too: a tray notification whenever the document is saved. The subscription is made from a
; timer, which runs as soon as the editor is up (this script's main thread is busy running it).
SetTimer(SubscribeToSaved, -100)
SubscribeToSaved() {
    try Rte.On("Saved", (e) => TrayTip(Rte.FileName, "RTE Writer: saved", "Iconi Mute"))
    catch TargetError    ; the editor is not up yet
        SetTimer(SubscribeToSaved, -200)
}

StartFile := A_Args.Length ? A_Args[1] : ""
RunCSharpFile(A_ScriptDir "\rte_writer15.cs", A_AhkPath, A_ScriptDir, StartFile)
if CSExitCode    ; an unhandled exception in the editor; its message is in the console output
    MsgBox("RTE Writer stopped because of an error (" CSExitCode "):`n`n" CSOutput, "RTE Writer", "Iconx")
ExitApp
