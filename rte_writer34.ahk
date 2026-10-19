#Requires AutoHotkey v2.0 64-bit
#SingleInstance Off    ; (File > New Window starts this script again: one instance for each window)
#Include "Csharp-in-AHK.ahk"
#Include "RteClient2.ahk"
; RTE Writer, iteration 34: a word processor built on RichEditorNET (rte_writer34.cs), hosted in this AHK process.
; Adds to iteration 31:
;   - Paragraph styles: the Style box and Home > Styles (Modify, New, Update to Match, Delete; the Classic, Plain and
;     Modern style sets; Save Styles as My Default). Headings are what their style says. Scripts: Rte.Style := "Heading 1".
; Iteration 31 added to iteration 30:
;   - Several windows: File > New Window and Open in New Window start this script again (so it allows more than one
;     instance); the windows share the settings without undoing each other's. Scripts: Rte.NewWindow().
; Iteration 30 added to iteration 29:
;   - Templates (File > New from Template: Letter, Memo, Report, Court Paper, Meeting Notes, and your own), with
;     {{fields}} asked for and remembered; Insert from Template (parts: Caption, Signature Block, Certificate of
;     Service) and Fill In Fields; File > Start Page; more macros in rte_macros, and on the Home tab.
;     Scripts: Rte.NewFromTemplate("Letter"), Rte.InsertTemplate("Signature Block"), Rte.FillFields(), Rte.StartPage := 1.
; Iteration 29 added to iteration 28:
;   - Autosave and recovery: a recovery copy every minute while there are unsaved changes, offered at the next start if
;     RTE Writer didn't close normally. Portable mode: a file named "portable" (or portable.txt) beside this script keeps
;     the settings and recovery copies in a "data" folder here. Scripts: Rte.AutosaveSeconds, Rte.Portable.
; Iteration 28 added to iteration 27:
;   - File > Close (Ctrl+W) closes the document and starts an empty one; Export as PDF is on the Layout tab and in the
;     Quick Access Toolbar's list. Scripts: Rte.Close().
; Iteration 27 added to iteration 26:
;   - View > Zoom > Multiple Pages: the printed pages side by side, read-only; double-click one to edit it there.
;     Scripts: Rte.MultiplePages := 1.
;   - Print Preview fits the screen above the taskbar, with page buttons and "Page 2 of 6" instead of a spin box.
; Iteration 26 added to iteration 25:
;   - Table cells keep their paragraphs' indents, tab stops and spacing in HTML files.
; Iteration 25 added to iteration 24:
;   - Space before and after paragraphs (Home > Paragraph > Line Spacing), and Word files open spaced as in Word: their
;     paragraphs' line spacing and space before and after (and, through the built-in converter, indents and tab stops).
;     Scripts: Rte.SpaceBefore := 12, Rte.SpaceAfter := 6.
; Iteration 24 added to iteration 23:
;   - Print Layout, on by default (View > Print Layout): lines break where they will on the printed page, as in Word
;     and LibreOffice, with the page on grey; line numbers on screen are then the printed ones. Its zoom: Page Width
;     (the default: the page fills the window's width), One Page, or manual (Zoom In / Out / 100%, Ctrl+wheel).
; Iteration 23 added to iteration 22:
;   - Fonts at the caret, as in Word: with nothing selected, a font or size chosen in the font box (or the Font dialog)
;     is what is typed next, and an empty line takes it at once and keeps it; the font box shows it.
; Iteration 22 added to iteration 21:
;   - Line spacing: Home > Paragraph > Line Spacing (1.0 to 3.0 lines, or Line Spacing Options for at least / exactly so
;     many points) for the selected paragraphs. Kept in RTF, Word, OpenDocument and HTML files. Scripts:
;     Rte.SetLineSpacing("1.5"), Rte.SetLineSpacing("exactly 14pt"), Rte.LineSpacing.
; Iteration 21 added to iteration 20:
;   - The Markdown and HTML views edit in Scintilla: syntax colours, line numbers (View > Source line numbers) and
;     Ctrl+wheel zoom; everything else in the views works as before.
; Iteration 20 added to iteration 19:
;   - Line numbers are printed: Print, Print Preview and Export as PDF number the printed lines, left of the text.
;     Rte.PrintedLineAt(position) gives "page:line" as printed, e.g. for citing a page and line.
; Iteration 19 added to iteration 18:
;   - Line numbers, as in Word: Layout > Line Numbers numbers the lines in the left margin, continuously or restarting on
;     each page, with Line Numbering Options for the first number, count-by (every 5th ...) and the distance from the
;     text. Lines in tables and page breaks don't count; Suppress for Current Paragraph leaves out others. Saved in RTF,
;     Word and OpenDocument files (Word and LibreOffice show and print them), HTML and Markdown. Scripts:
;     Rte.LineNumbers := "continuous", Rte.SetLineNumbers("page", 1, 5), Rte.SuppressLineNumbers(1), Rte.LineNumberAt(pos).
; Iteration 18 added to iteration 17:
;   - Table properties: Table Tools > Properties (or right-click a table) sets column widths, the table's alignment, cell
;     padding, row height, and the borders, shading and vertical alignment of the selected cells or the whole table;
;     Borders and Shading are on Table Tools too. Kept in RTF, Word and OpenDocument files and in HTML; Markdown tables
;     can't keep them. Word's cell background colors now show. Scripts: Rte.SetColumnWidths("2in, 3in"),
;     Rte.SetCellBorders("outside", 1, "#C00000"), Rte.SetCellShading("LightYellow"), Rte.SetTableAlignment("center") ...
; Iteration 17 added to iteration 16:
;   - Markdown and HTML views (View > Document View): the document as Markdown or HTML in a plain-text editor, alone or
;     side by side with the formatted text, which follows as you type. A .md (or .html) file edited in its source view
;     is saved exactly as typed. Rte.View := "Markdown", Rte.SideBySide := 1 and Rte.SourceText work from scripts too.
;     In the source view, the formatting keys and ribbon buttons insert Markdown or HTML: Ctrl+B **bold** / <strong>,
;     Ctrl+1 # heading / <h1>, the list buttons - / 1. / <ul>, Tab-like indenting of list items (Ctrl+M), alignment as
;     <div align>, the font control as <span style>, and Table, Link, Picture, Page break and Date.
;   - Line breaks (Shift+Enter) stay line breaks in HTML (<br>) and Markdown, in headings, lists and tables too.
;   - Markdown tables are written lined up; in the Markdown view, Tab / Shift+Tab in a table line it up and move
;     between cells (Tab in the last cell adds a row), and the Table button tidies the table with the caret. Tables
;     with long cells are written compactly instead. In the formatted text, Tab / Shift+Tab move between the cells of
;     any table (also from Word), and Tab in the last cell adds a row.
; Iteration 16 added to iteration 15:
;   - A ribbon (the Windows Ribbon, through the WindowsRibbon NuGet package) instead of the menu bar and toolbar: a File
;     menu with recent documents, a quick access toolbar, Home / Insert / Layout / Review / View / Tools tabs, and Table
;     Tools while the caret is in a table. Press Alt for KeyTips; the keyboard shortcuts are unchanged. The ribbon is
;     ribbon\RteRibbon.ribbon, built from ribbon\RteRibbon.xml by ribbon\build_ribbon.ahk.
; Iteration 15 added to iteration 14:
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
; %APPDATA%\RTE Writer\settings.json (Rte.SettingsFile), or data\settings.json here in portable mode; set
; RTE_WRITER_SETTINGS to a file path to keep them elsewhere.
; Optionally pass a document to open:  rte_writer34.ahk "C:\path\document.md"
;
; AHK automation: other scripts #Include "RteClient2.ahk" and call Rte.<Name>(...), react to events with
; Rte.On(Event, Callback) and queue calls with Rte.Post(...). Tools > Automation API lists everything that can be called,
; e.g. Rte.InsertPageBreak(), Rte.LegalNumbering(), Rte.LegalFormat := "All", Rte.InsertTable(3, 4), Rte.ExportDocx("C:\path\copy.docx").
;
; Hotkeys defined here keep working while the editor is open, and call it directly (in-process):
#HotIf WinActive("RTE Writer ahk_pid " ProcessExist())    ; (each window's own: several can be open)
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
RunCSharpFile(A_ScriptDir "\rte_writer34.cs", A_AhkPath, A_ScriptDir, StartFile, A_ScriptFullPath)
if CSExitCode    ; an unhandled exception in the editor; its message is in the console output
    MsgBox("RTE Writer stopped because of an error (" CSExitCode "):`n`n" CSOutput, "RTE Writer", "Iconx")
ExitApp
