#:package EllipticBit.RichEditorNET
#:package WindowsRibbon@2.16.2
#:package Scintilla5.NET@7.0.0
#load "rte_markdown6.cs"
#load "rte_htmllists2.cs"
#load "rte_tables2.cs"
#load "rte_tablecells.cs"    // (iteration 92: Merge Cells, Split Cells, Delete Cells)
#load "rte_office7.cs"
#load "rte_docxodus.cs"
#load "rte_rtfdocx.cs"
#load "rte_legal.cs"
#load "rte_pagebreaks.cs"
#load "rte_headerfooter.cs"
#load "rte_linebreaks.cs"
#load "rte_linenumbers2.cs"
#load "rte_sourcebox.cs"
#load "rte_linespacing2.cs"
#load "rte_printlayout.cs"
#load "rte_pageview.cs"
#load "rte_recovery.cs"
#load "rte_templates.cs"
#load "rte_styles.cs"
#load "rte_paragraph.cs"
#load "rte_docxstyles.cs"
#load "rte_pagegaps.cs"
#load "rte_crossrefs.cs"
#load "rte_footnotes.cs"
#load "rte_toc.cs"
#load "rte_pdf.cs"
#load "rte_snippets.cs"
#load "rte_hooks.cs"
#load "rte_palette.cs"
#load "rte_theme.cs"
#load "rte_sections.cs"
#load "rte_extensions.cs"
#load "rte_headline.cs"
#load "rte_strings.cs"
#load "rte_ruler.cs"
#load "rte_pagesetup.cs"
#load "rte_paraframes.cs"    // (iteration 87: paragraph borders and shading)
#load "rte_tabledata.cs"    // (iteration 91: tables from data: CSV, Excel, JSON)
// AhkWord: a word processor built on EllipticBit.RichEditorNET, hosted in AutoHotkey by Csharp-in-AHK.
// Run it with AhkWord.ahk. Iteration 66 adds to iteration 65:
//   - Print Layout: a document whose first paragraphs are hidden (Word puts floating shapes' anchors there) no longer has
//     its first lines painted over by the top margin: page 1's gap goes on the first paragraph that shows.
// Iteration 65 added to iteration 64:
//   - Paste as Plain Text (Ctrl+Shift+V, the right-click menu, the command palette, PastePlainText for macros): the
//     clipboard's text only, taking the formatting where it lands, so a web page's or a chat's fonts, spacing and
//     rules stay behind.
// Iteration 64 added Change Case > Headline Case (rte_headline.cs). Iteration 63 added to iteration 62:
//   - Sections (rte_sections.cs): Insert > Section Break starts a new section on a new page, with its own header and
//     footer (at first the section before's, "same as previous") and its page numbers starting again (at 1, or a number
//     of its own); File > Header and Footer edits the caret's section. Printing, Print Preview, PDF and the page view
//     give each page its section's header, footer and number; {sectionpages} is the pages in the section.
// Iteration 62 added to iteration 61:
//   - RenderTemplate, a hook for an extension's own template language: given a template's RTF (New from Template,
//     Insert from Template) or the document's (Fill In Fields), it returns it filled in, or null when the template
//     isn't in its language; AhkWord's own {{fields}} left in it are filled in afterwards, as before.
//   - The header's and footer's parts filled in too (New from Template, Fill In Fields): AhkWord's own {{fields}} in
//     them, asked for with the text's; an extension's language through RenderTemplateText.
// Iteration 61 added to iteration 60:
//   - Hooks for what the extension point loads: AddExtensionCommand (Tools > Extensions, shown once there is one, and
//     the command palette), FieldSource (values for template fields before the dialog asks), SetExtensionStatus (a
//     place in the status bar).
// Iteration 60 added to iteration 59:
//   - An extension point (rte_extensions.cs): empty in AhkWord; a build with more features puts its own there, which
//     loads its files and is called once the window is open (AhkWordExtensions.Started).
// Iteration 59 added to iteration 58:
//   - AhkWord's own .NET (its dotnet folder) kept up to date (AhkWordSetup.ahk, before .NET is loaded): Microsoft's
//     list of .NET 10 releases fetched once a week in the background; a newer .NET offered at the next start (Update
//     now, Later, Skip), downloaded and checked as the first one, put in in one move (the old folder renamed away first:
//     refused while another window runs on it). An empty no-update-check file beside AhkWord turns the look off.
// Iteration 58 added to iteration 57:
//   - Export as PDF (and printing) show their progress in the status bar, page by page, and Esc stops them (while
//     AhkWord's window is in front: DocumentPrinter.Progress and CancelRequested); a stopped export leaves no file.
//     After the last page, what is going on: Windows finishing the file, AhkWord adding the bookmarks.
// Iteration 57 added to iteration 56:
//   - The command palette (Ctrl+Shift+P, Tools > Command Palette, the quick access toolbar; rte_palette.cs): type part
//     of a command's name. The ribbon's commands, as the ribbon names and places them (RteRibbon.xml), with their
//     shortcuts, run as their buttons do (registered by Button, Toggle and Check); the macros, styles, snippets and
//     recent files too. Greyed out where the ribbon greys them out; the ones used last first. Scripts: PaletteItems.
// Iteration 56 added to iteration 55:
//   - Event hooks (Tools > Event Hooks; rte_hooks.cs): a macro that AhkWord runs by itself when it starts, a document
//     is opened, made, saved, printed or closed, or a style is changed, told the event and the file or style
//     (A_Args, AHKWORD_EVENT, AHKWORD_EVENT_DETAIL): no script of the user's kept running. BeforeSave and Closing wait
//     for their macro (its Rte calls answered meanwhile, the window taking no input); an event doesn't start its macro
//     again while it runs. Kept in the settings (EventHooks). Scripts: SetEventHook, EventHook, EventHookList,
//     HookEventNames.
// Iteration 55 added to iteration 54:
//   - Snippets (Insert > Snippets; rte_snippets.cs): a part with an abbreviation, typed in the document (then Space,
//     Enter, Tab or punctuation, or at once) to insert the part, formatted, its fields asked for. The abbreviations are
//     AutoHotkey hotstrings, which AhkWord.ahk sets up for its own window only and the document in focus (each AhkWord
//     window by itself; a snippet goes in through the editor: FireSnippet waits until the editor has the abbreviation
//     and the ending character typed, then swaps them for the part, as one undo step). New from Selection makes a
//     part of the selection; Export as Script writes the snippets as plain-text hotstrings for any program. Kept in
//     snippets.json beside the user's own templates. Scripts: Snippets, AddSnippet, RemoveSnippet, InsertSnippet,
//     SaveSelectionAsSnippet, ExportSnippets.
// Iteration 54 added to iteration 53:
//   - A download that carries everything (make_release.ps1): the libraries (lib\packages, as NuGet lays them out, and
//     their list), the C# compiler's two DLLs (lib\roslyn) and the C# host built beforehand (lib\host), named in
//     Csharp-in-AHK.ini, so that nothing is downloaded at the first start and nothing written outside AhkWord's folder
//     (Csharp-in-AHK.ahk: HostDll, PackagesManifest, CacheDir; relative paths). Without .NET 10 on the computer,
//     AhkWord offers to download it from Microsoft for itself only, checked against Microsoft's SHA-512, into its own
//     dotnet folder, or to open Microsoft's page (AhkWordSetup.ahk); the "complete" download has it inside.
//     SkiaSharp's native library is found beside its package there too (OfficeConverters.LoadSkia).
// Iteration 53 added to iteration 52:
//   - PDF bookmarks (rte_pdf.cs): Export as PDF puts in a bookmark for each heading, under the one of the level above
//     it, each opening its page at the heading, and makes each table of contents' entry a link to its heading's page,
//     as Word's PDFs have them; the file opens with its bookmarks shown. (Microsoft Print to PDF writes the pages only:
//     the bookmarks and links go into the file it wrote, with PDFsharp, where DocumentPrinter.PdfPlaces says the
//     headings and entries are printed.) Scripts: ExportPdf(path, bookmarks := 1).
// Iteration 52 added to iteration 51:
//   - A table of contents (Insert > Table of Contents; rte_toc.cs), as Word's: the headings of levels 1 to 3 (1 to 6),
//     each with its page as printed at the right margin after dots (or dashes, a line, nothing; or just after the text;
//     or no page numbers), under a title ("Contents", in TOC Heading), in the styles TOC 1 to TOC 6. Update (F9, or the
//     table's right-click menu) makes it again from the headings; its page numbers alone are brought up to date before
//     saving, printing and exporting. Ctrl+click an entry to go to its heading. Insert > Table of Contents with the
//     caret in a table changes its options. Scripts: InsertTableOfContents, UpdateTableOfContents, TableOfContents.
//     In the editor the table is paragraphs between hidden markers, each entry with a hidden link to a target at its
//     heading; RTF (and so Word and LibreOffice files) gets Word's TOC field, its entries HYPERLINK fields to
//     bookmarks at the headings with PAGEREF fields for the pages, which Word and LibreOffice can update; the built-in
//     .docx converter the same; HTML and Markdown links to anchors at the headings. All are read back as tables,
//     Word's and LibreOffice's own too.
// Iteration 51 added to iteration 50:
//   - Footnotes (Insert > Footnote, Ctrl+Alt+F; rte_footnotes.cs): a superscript number in the text, its note (its
//     paragraphs, with bold, italic and underline) at the foot of its page: printed, in Print Preview and PDFs, and on Show Pages' sheets,
//     under a short line, in the document's Footnote Text style (Footnote dialog > Size of all the notes, Modify
//     Style, Rte.FootnoteSize; a Word or LibreOffice document's notes' size kept: Footnotes.LoadedSize), else the text's
//     font at three quarters of its size (12 points: 9). The pages make room
//     for their notes as Word does (DocumentPrinter.CountPages): a page ends at its last line for which the text, the
//     notes going on from the page before and its own notes up to the last one's first line fit; the notes' lines
//     that don't fit go on at the next page's foot. Show Pages' sheets show the printed pages' lines, page breaks and all. Numbered 1, 2, 3 ... in order; hover over a
//     number for its note, Ctrl+click it (or Insert > Footnote on it) to change it, an empty note takes it out.
//     Scripts: InsertFootnote, FootnoteText, FootnoteList, FootnoteCount, SetFootnote, DeleteFootnote (`n between a
//     note's paragraphs, <b> <i> <u> for its formatting).
//     In the editor a footnote is hidden text around its number, its note in it, so it goes wherever its number goes.
//     RTF (and so Word and LibreOffice files) gets real footnotes (\footnote, with \fet0: else Word made them
//     endnotes); Word's and LibreOffice's come back as footnotes; the built-in .docx converter writes and reads real
//     footnotes too; HTML gets links to a list of the notes at its end, Markdown [^1] and "[^1]: note".
//   - Numbered styles' sub-levels' spacing (Modify Style > Space between sub-levels; ParagraphStyle.SubLevelSpace,
//     StyleEngine.FitSubLevelSpacing, after every renumbering): None at first (no space after a paragraph with points
//     under it, nor between the points; the style's space after the last), AfterParagraph, or Each (as before: a
//     12-point space after every paragraph looked like a blank line between the points). A paragraph's own space after
//     (set by hand) stays.
//   - Cross-references' Word fields: the text's own space after them was taken for a control word's delimiter.
//   - Show Pages: a black streak at the caret's line after Tab on a numbered paragraph, the whole window black a moment
//     after inserting a footnote: RichEdit, not done laying the change out, left those lines out of the bitmap the
//     window is painted from (black when new). The bitmap starts with the window's background now, and while RichEdit's
//     layout is pending (TOM's E_PENDING) the window keeps what it shows and is painted again a moment later.
// Iteration 50 added to iteration 49:
//   - Cross-references to numbered paragraphs (Insert > Cross-reference; rte_crossrefs.cs): "see paragraph 3(b)",
//     to a legal item, a numbered style's paragraph or a numbered list's item, showing its number ("(b)") or its full
//     number ("3(b)": a letter or roman sub-level in parentheses, as citations have it), without a trailing period;
//     also a sub-point typed in front of a paragraph after a numbered one ("e. ", not a list item: "23(e)"). The number follows the paragraph's: as legal items and numbered
//     styles are renumbered (in the same undo step), and, for lists RichEdit numbers itself, before saving, printing
//     and exporting, and with F9 (Insert > Update). Ctrl+click goes to the paragraph. A reference whose paragraph was
//     deleted says "(reference not found)". Scripts: CrossReferenceTargets, InsertCrossReference, UpdateCrossReferences.
//     Both ends are hidden text in the editor; RTF (and so Word and LibreOffice files) gets Word's own bookmarks and
//     REF fields, which Word and LibreOffice keep up to date; HTML and Markdown get anchors and links; the built-in
//     .docx converter gets bookmarks and links to them. All are read back.
//   - Lists, fixed (there before):
//     - An ordinary numbered list right after legal items (or a numbered style's) started at "2.": RichEdit numbered it
//       on from them. The legal items before it get another kind of numbers (hidden, so not shown; LegalLists.BreakChains).
//     - Saved as .rtf and opened again, such a list came back as typed numbers ("2.<tab>"), no longer a list.
//     - A numbered style's sub-level ("a.") came back from .html and .md as legal numbering ("5.1"), its paragraphs in
//       Normal: HTML keeps the style on each item (data-rte-style), Markdown after [type=legal] ([style=...]).
//     - Markdown wrote a paragraph with a typed number and a tab ("2.<tab>Text") as a table's row.
//     - Legal items followed by a numbered style's: the built-in .docx converter gave both one Word list with the style's
//       formats ("a." for "2.1"), as Word's and LibreOffice's RTF did; now a Word list for each run, the second going on
//       from the first. Read back, the style's sub-level was taken for legal numbering; from Word's file, the style
//       wasn't found numbered. And LibreOffice numbered an ordinary list after legal items on from them ("6.").
// Iteration 49 added to iteration 48:
//   - A different first page header and footer (Layout > Header & Footer > First page; DifferentFirstPage,
//     SetFirstHeader, SetFirstFooter for scripts): page 1 gets its own parts, or none if they are empty (a letter's
//     or a pleading's first page). Shown on the sheets (Show Pages) and the strips, printed, and kept in the files:
//     RTF \titlepg with \headerf / \footerf, .docx w:titlePg with first-page header and footer parts, .odt
//     style:header-first / footer-first (as LibreOffice writes them); read back from all three, also as Word or
//     LibreOffice changed them. "Use for new documents too" keeps them as the default.
//   - Show Pages: a blank document has its sheet, header and footer before anything is typed (the gaps were left out
//     for a document of one paragraph mark).
//   - Show Pages: after a recovery copy (every AutosaveSeconds while there are changes; its RTF is read without the
//     gaps), the text sometimes stayed drawn without page 1's top margin, over the header, until scrolled: RichEdit
//     lays the text out again only once the messages waiting are done, so the window is painted once more then.
//   - A header or footer changed (the dialog, scripts) is shown at once, not half a second later.
// Iteration 48 added to iteration 47:
//   - Show Pages is on by default (redesigned in iteration 42, without flicker since 44): new settings have it; saved
//     settings keep their own.
//   - A crash fixed (there since the ribbon came, iteration 16): now and then, after opening a long document, AhkWord
//     ended with an access violation in UIRibbon.dll, in the ribbon's own timer (COfficeSpaceUser::OnUpdate >
//     InvalidateUICommandInternal, as Microsoft's symbols name it). Changing the Styles menu's label, RibbonLib tells
//     the ribbon with a pointer to a key it has just made; the ribbon queues it for its timer, and a garbage collection
//     in between moved or freed the key. The label is now changed with a key pinned for good (SetRibbonLabel).
//   - Show Pages: the window painted whole when the gaps are in, its painting held meanwhile (WriterEditBox.HoldPaint,
//     not LockWindowUpdate any more): after opening, the window sometimes kept what it showed before the gaps went in
//     (a table row drawn over a paragraph, until scrolling painted it again).
//   - The ribbon is sent only what changed (WriterForm.Push): it was sent some 45 values at every caret move, nearly
//     all unchanged (less work for it; not the crash's cause, as it turned out).
// Iteration 47 added to iteration 46:
//   - Saving as .rtf (and the recovery copy) keeps every paragraph's style: the list indents are fitted before the
//     styles go in. Fitting them goes through RichEdit, which writes only its own styles (Normal, the headings): a new
//     document's numbered style paragraphs (Numbered Paragraph) came back as Normal, keeping only their numbers.
// Iteration 46 added to iteration 45:
//   - Line numbers on the screen as far from the text as printed and as in Word: the numbering's distance, or a
//     quarter of an inch (auto), at the zoom (LineNumberGutter.DistancePixels); they were 8 pixels from it.
//     Within the left margin: as much of the distance as fits there (LineNumberGutter.MaxWidth).
//   - Line numbers on their lines with Show Pages: the line's top from EM_POSFROMCHAR and the baseline's step below
//     it from TOM (TOM's points alone were sometimes those of the view before it was scrolled back, by page 1's top
//     margin: the numbers a margin below their lines); painted again when they aren't where they belong.
//   - A document opened at a large zoom: the page's contents laid out at the page's new size (they stayed as for the
//     page at 100%: the text at the window's edge, the numbers out of view); the page widened for lines drawn wider
//     than the printer's by a few percent at most (it had grown to half as wide again, the footer's middle far right).
//   - The grey around the page keeps where it is scrolled to when the editor gets the focus (PageHost): WinForms
//     scrolled the editor's left edge to the window's, the line numbers and the left margin out of view.
//   - Show Pages: the last sheet's end below the text, with its footer, and grey after it (PageGaps.LastSheetEnd: the
//     last page has no gap, where typing goes on); the wheel, the scroll bar and Page Down scroll on past the text's
//     end to it (WriterEditBox.ScrollPastEnd). Its blank rest is shortened if taller than the view could scroll to.
//   - Saving with Show Pages on: the file made without its gaps at every step (HTML had them as huge margins before a
//     page's first paragraph: the steps after the library's HTML read the paragraphs' spacing from the editor).
//   - HTML: the document's styles for AhkWord at the file's end, a comment after </html> (StyleHtml.EndComment): in the
//     head they were a long line of data at the top of the HTML source view. Files with them in the head still read.
//   - OpenDocument with line numbers: the header's and footer's paragraphs not numbered (LibreOffice numbered the
//     footer's line, counting on from page to page: a second page number in the margin).
//   - Rte.PageViewReport (and the Page view report macro) lists the line numbers in view first: where each was drawn
//     and where it belongs, and when the numbers and the editor were last painted.
// Iteration 45 added to iteration 44:
//   - Show Pages with line numbers: the sheets' bands go across the numbers too (they had a white gap there:
//     LineNumberGutter.Bands), and the numbers are painted in the same pass as the editor, not a moment after it.
// Iteration 44 added to iteration 43:
//   - Show Pages: the sheets' grey bands in the page's margins scroll with the text in the same move: in Print Layout
//     the margins are the editor's own (EM_SETMARGINS: WriterEditBox.TextMargins; the header and footer text, page
//     lines and page breaks kept within the text's width), so RichEdit moves their pixels with the text's; drawn beside
//     the editor, they followed it a moment later, a flicker at the text's edges while scrolling. (With line numbers
//     the left margin is still beside them: its bands follow in the same pass as the editor's paint, FollowGapBands;
//     and the editor paints at once when the wheel, the scroll bar or a key scrolled it.)
//   - Settings: reading, writing and putting the new file in place are tried again a few times if something has the
//     file open a moment (Settings.Retried): a save failed now and then with "access denied" and was lost.
// Iteration 43 added to iteration 42:
//   - Less flicker in Show Pages and Print Layout: the editor paints through a bitmap (RichEdit's text, then the sheets'
//     bands and the page lines over it, onto the screen at once: WriterEditBox.WndProc), as the panel around it does;
//     the gaps worked out again a second after typing are put in with the window not drawing until they are in, and
//     the line with the caret stays where it was on the screen (QuietGaps).
// Iteration 42 added to iteration 41:
//   - Show Pages redesigned (rte_pagegaps.cs): where a page starts with a paragraph, its gap is space before that
//     paragraph (ordinary paragraph formatting: the caret and typing behave as anywhere); only where a page starts
//     inside a paragraph is the line before it raised, and text typed in that line is raised with it (no more lifting
//     the gap while the caret is on it, which moved the text by a page); no gap after the last page (where typing goes
//     on). Enter at a page's first paragraph leaves the gap on one paragraph (PageGaps.Tidy); formats read and set
//     there see the paragraph's own space before (FirstParagraphAsIs, ExtraAt).
// Iteration 41 added to iteration 40:
//   - AhkWord's icon (images\icon\ahkword.ico: a W whose strokes are A, H and K; the W alone at 16 to 24 pixels),
//     drawn by images\icon\make_icon.py: the window's and the taskbar's (WriterForm.AppIcon), the tray's (AhkWord.ahk).
//   - AhkWord is an app of its own on the taskbar (TaskbarIdentity): its app ID "AhkWord" for the process and its
//     windows, and the windows' relaunch properties (the icon, the name, AutoHotkey with AhkWord.ahk): the taskbar shows
//     AhkWord's icon, groups its windows apart from other AutoHotkey scripts, and pinned starts AhkWord.
// Iteration 40 added to iteration 39:
//   - Home > Styles > Copy Styles from Document (Word's Organizer): another document's styles, ticked, come into this
//     one, and into your defaults for new documents if "Also for new documents" is (StyleImport.From: AhkWord's RTF and
//     HTML whole; a .docx's, its numbered styles numbered as Word's definitions say, DocxStyles.Numbering; other RTF's
//     style sheet, RtfStyles.Definitions). A style this document has takes the copied definition (its paragraphs too).
//     Scripts: Rte.StylesIn(path), Rte.CopyStyles(path, "Law Paragraph", 1).
// Iteration 39 added to iteration 38:
//   - The name AhkWord (RTE Writer until iteration 38): the window, dialogs and messages; the files are AhkWord.ahk
//     and AhkWord.cs from now on (an iteration's own copy: make_version.ps1). Its folders are %APPDATA%\AhkWord and
//     %LOCALAPPDATA%\AhkWord, the old ones' files copied into them the first time (AppFolders); AHKWORD_SETTINGS and
//     AHKWORD_CONVERTERS name the environment variables (the old RTE_WRITER_ names still work). Scripts are unchanged:
//     the Rte object, RteClient2.ahk (which finds AhkWord's windows, and RTE Writer's), RTE_WRITER_HWND and the event
//     message keep their names.
// Iteration 38 added to iteration 37:
//   - List levels' indents: Paragraph > List levels has Number at and Text at for each level (points from the margin,
//     Word's number position and text indent; empty: automatic, as before). The levels are the document's
//     (WriterForm.Levels; {\*\rtewriterlevels} in its RTF; a .docx's from where its list paragraphs are drawn,
//     OfficeConverters.ListLevelIndents), and new documents' unless "Also for new documents" is unticked. Lists made,
//     nested or given their levels, legal numbering, typed numbers made lists, and the Word and OpenDocument lists all
//     take them (LegalLists.IndentsFor). Scripts: Rte.SetListLevel(1, "Decimal", ".", 1, 36, 72).
//   - Automatic numbering style per level is List levels' checkbox now, no longer a menu item right under List
//     Levels (one click turned it off unnoticed). Scripts: Rte.AutomaticNumberingPerLevel := 0.
//   - Fixes: ordinary lists lost their indents through Word (back at 18/0 through LibreOffice; through the built-in
//     converter made a numbered style of Word's List Paragraph); new lists were drawn at the margin until saved;
//     ending a list left its hanging indent.
// Iteration 37 added to iteration 36:
//   - Sub-levels in numbered styles: Tab and Shift+Tab move a numbered style's paragraph to a sub-level under the
//     one before it, numbered by ParagraphStyle.FormatAt (1. / a. / i. / (1) / (a) / (i) ... from the style's own
//     format, or ParagraphStyle.SubLevels, Modify Style > Sub-levels), its indent a level further in
//     (ParagraphStyle.SubLevelIndent, 36 pt at first; Modify Style > Sub-level indent). Tab
//     at the start of an unnumbered paragraph of a numbered style makes it a sub-level item, removing a number
//     typed in front (so does Tab over several selected, and a numbered style given to a paragraph that starts with
//     a typed number of one of its levels: StyleEngine.TypedLevel). Kept through style changes, saving (the rtewriterlists group now has start:level) and Word
//     and OpenDocument files (a format per list level); a Word style numbered over several levels comes in with
//     them as its sub-levels.
//   - Numbering off stays with the paragraph (as Word's numId 0): a numbered style's unnumbered paragraph keeps it and
//     its own indents through other styles, Normal too, and back (LegalLists.SetNumberingOff: ListType NoNumber on a
//     paragraph in no list; saved in {\*\rtewriteroff count:i,j,...}). Changing a style and back gives the paragraphs
//     as they were.
// Iteration 36 added to iteration 35:
//   - Numbered styles: a paragraph style can number its paragraphs (ParagraphStyle.Numbering "1.", "(a)", "i." ...
//     or "1.1" legal, NumberLevel 1-9; Modify Style > Numbering; the built-in Numbered Paragraph). They are legal
//     numbering's items (rte_legal.cs: numbers worked out from position, one numbering through the document, restarts
//     only where asked), numbered in their style's format and keeping their style's indents. Pasted numbered
//     paragraphs join the numbering where they land (StyleEngine.JoinPasted); a Word document's numbered style comes in
//     as one (StyleEngine.NumberAfterLoad); exports give Word each level's own format; AhkWord's RTF keeps where
//     the numbering starts again ({\*\rtewriterlists}). A paragraph of a numbered style can have its numbering turned
//     off (No numbering; a Word paragraph whose numbering is off): it stays so, its style the same. Items of numbered
//     styles have no room for a RichEdit number (ListTab 0), so that their first lines show where their indents say.
//   - Show Pages is off by default: a preview, not reliable while typing yet (to be redesigned). The line with the
//     caret has no gap (PageGaps.Lift): the caret at the end of a raised line showed, and typed text went, at the
//     bottom of the page.
//   - Csharp-in-AHK and Csx-in-AHK restore from nuget.org when no NuGet source is enabled; the README says how to
//     install .NET 10 (winget) and that the first start takes a while.
// Iteration 35 added to iteration 34:
//   - Page view (rte_pagegaps.cs): in Print Layout, View > Show Pages shows each printed page as a sheet of its own,
//     its header and footer in its margins and a grey band between the sheets, as in Word, while the text is edited as
//     before (a page that ends inside a table gets a "Page N" line instead). The sheets are the view's alone: saving,
//     copying, printing and the other formats see the document as it is. Scripts: Rte.PageView, and Rte.PageViewReport
//     (the macro Page view report) for a bug report: where the pages and gaps are, none of the text.
//   - Word's lists as Word numbers them: paragraphs in no list between numbered ones (a heading, a quotation) are no
//     longer numbered (Word's RTF gives them an old-style "item without a number", which RichEdit numbered and the legal
//     lists took for theirs), and the numbering goes on after them as in Word (OfficeConverters.WithoutUnlistedNumbering,
//     ContinueWordLists).
//   - No numbering (the Numbering menu, List Options > List style; Rte.RemoveList): the selected paragraphs' lists end,
//     whatever their numbering, and the Numbering button ends a list numbered "1." as Word's are.
// Iteration 34 added to iteration 33:
//   - Styles in Word and OpenDocument files, whichever converter: the built-in one (rte_docxstyles.cs) writes the
//     document's styles (whole definitions) and each paragraph's style and line and page breaks into the .docx, and
//     reads a .docx's styles (as Word works them out) into the editor; LibreOffice reads .odt files that way too (and
//     goes first for .odt); Word's RTF style names are read right. A document's own styles show as "Styles: The
//     document's"; Home > Styles > Apply My Styles gives it yours (Rte.ApplyMyStyles()).
//   - View > Structured / Free-form retired: any font size, always; headings and the like are styles.
// Iteration 33 added to iteration 32:
//   - The Paragraph dialog (rte_paragraph.cs; Home > Paragraph's Paragraph settings button, and Line Spacing
//     Options), as Word's: Indents and Spacing (alignment, left and right indents, a first-line or hanging indent, the space
//     before and after, line spacing) and Line and Page Breaks (widow/orphan control, keep with next, keep lines
//     together, page break before, suppress line numbers), with a preview; Tabs (the paragraphs' tab stops and the
//     document's default ones); Set As Default (the paragraph's style gets these, for this document or new ones too).
//     Modify Style has it too (its Paragraph button), with "Don't add space between paragraphs of the same style".
//   - Printing, Print Preview, Export as PDF, Multiple Pages and line numbers restarting on each page now keep those
//     rules (RichEdit's own page layout ignores them). Scripts: Rte.ParagraphFormat(), Rte.SetParagraphFormat(...),
//     Rte.Tabs, Rte.SetTabs(...), Rte.DefaultTabStop.
// Iteration 32 added to iteration 31:
//   - Paragraph styles (rte_styles.cs): each paragraph has a style (Normal, Title, Subtitle, Heading 1 to 6, Quote,
//     Code, and your own), chosen in the Style box (Ctrl+0 Normal, Ctrl+1 to Ctrl+6 the headings, Ctrl+7 Code). Home >
//     Styles: Modify Style (font, size, colour, alignment, indents, spacing, line spacing, keep with next, heading level,
//     the style after Enter), New Style (from the paragraph at the caret), Update Style to Match (the style made like the
//     paragraph), Delete Style, Clear Formatting (Ctrl+Space); changing a style changes its paragraphs (formatting of their own stays, unless it covers most of a paragraph, as in Word). Copy and paste keep the paragraphs' styles. Three style
//     sets, none of them Word's: Classic (Times New Roman 12, black headings: the default), Plain (the look before
//     styles), Modern; Save Styles as My Default gives new documents this document's styles.
//   - Headings are what their style says, not guessed from fonts: the outline, Markdown's #, HTML's h1-h6 (a bold 12 pt
//     line is no longer Heading 4). Styles are kept in RTF as Word and LibreOffice read them (a style sheet), so Word
//     files written through Word or LibreOffice have them too; a document's own look is kept when it is opened (its
//     styles are defined from its own paragraphs). Scripts: Rte.Style := "Heading 1", Rte.Styles(), Rte.ModifyStyle(...),
//     Rte.NewStyle(...), Rte.StyleSet := "Modern", Rte.SaveStylesAsDefault().
// Iteration 31 added to iteration 30:
//   - Several windows: File > New Window and File > Open in New Window start another AhkWord (a little down and to
//     the right of this one), to work on two documents at once; each window is a whole editor of its own (its recovery
//     copies, its macros, its automation). The windows share the settings file: each saves only what it changed, so a
//     window closing doesn't undo the other's (recent documents, macro hotkeys and remembered values are merged item
//     by item). Scripts started outside the editor talk to the AhkWord window used last. Scripts: Rte.NewWindow().
// Iteration 30 added to iteration 29:
//   - Templates (rte_templates.cs): File > New from Template (and the start page) makes an untitled document from a
//     template: a document in the templates folder beside the program (Letter, Memo, Report, Court Paper, Meeting Notes)
//     or in your own Templates folder (beside the settings). Its fields, {{Your name}} or {{Salutation|Sir or Madam}},
//     are asked for in a dialog, and the values of your own (Your name, Your address ...) remembered; a field named with the word
//     "date" ({{Hearing date}}) gets a calendar beside its box (dates aren't remembered). Insert > Insert from Template puts a template's content at the caret instead, its fields
//     filled in: the parts in a Parts folder are pieces for that (Caption, Signature Block, Certificate of Service).
//     Insert > Fill In Fields fills in the fields left in the document. Scripts: Rte.NewFromTemplate("Letter",
//     "Recipient=Ms. Rivera"), Rte.InsertTemplate("Signature Block"), Rte.FillFields(), Rte.Templates().
//   - The start page, shown at the start (without a file to open; "Show this page when AhkWord starts") and by
//     File > Start Page: a blank document, the templates, the recent documents and the macros. Esc goes to the blank
//     document. Scripts: Rte.StartPage := 1.
//   - The macros gallery is on the Home tab too, and in the Quick Access Toolbar's list.
//   - More macros in rte_macros: Clean up pasted text, My signature, Collect from anywhere (a hotkey in any
//     program), PDF on save, and Folder to PDF.
// Iteration 29 added to iteration 28:
//   - Autosave and recovery (rte_recovery.cs): while the document has unsaved changes, a recovery copy of it (RTF, with
//     its header, footer and line numbering) is written every minute, beside the settings, and removed when it is
//     saved, closed or discarded, and when AhkWord closes. If AhkWord ended without closing (a crash, a power cut,
//     Windows restarting), the next start offers the copy: recover it (as unsaved changes to the document it was, which
//     Save then replaces), discard it, or decide next time. The document's own file is never written by autosave.
//     Scripts: Rte.AutosaveSeconds := 120 (0: off), Rte.RecoveryFolder.
//   - Portable mode: with a file named "portable" or "portable.txt" beside the program, the settings, the recovery copies
//     and the converters' state (LibreOffice's profile, the log) are kept in a "data" folder beside it instead of
//     %APPDATA% and %LOCALAPPDATA%, so that the folder can be carried to another computer with them. Scripts: Rte.Portable.
// Iteration 28 added to iteration 27:
//   - File > Close (Ctrl+W): closes the document (asking to save its changes) and starts an empty one; the window stays
//     open, as in Word. Scripts: Rte.Close().
//   - Export as PDF on the ribbon (Layout, next to Print Preview) and in the Quick Access Toolbar's list, as well as in
//     File > Export: through Microsoft Print to PDF, the default printer staying as it is (DefaultPrinterGuard).
// Iteration 27 added to iteration 26:
//   - Multiple Pages (rte_pageview.cs): View > Zoom > Multiple Pages shows the printed pages side by side, as many as
//     fit across the window (the printer's own pages: the same page breaks, header, footer and line numbers), read-only.
//     Zoom In / Out / 100% and Ctrl+wheel show them larger or smaller; the arrow keys move between them; double-click a
//     page (or Enter) to edit the document there; Esc, or the button again, goes back to where the caret was. A command
//     that changes the text goes back to editing first. Scripts: Rte.MultiplePages := 1.
//   - Print Preview is a window of its own (PagePreviewForm) fitted to the screen's working area (it reached under the
//     taskbar), with Print, First / Previous / "Page 2 of 6" / Next / Last, Zoom Out / In / Whole Page and Close, the
//     page view's keys, and double-click a page to edit it there. Print Preview's pages no longer cut off the ends of
//     long lines (DocumentPrinter.PreviewPage records the text for the printer first).
// Iteration 26 added to iteration 25:
//   - Table cells keep their paragraphs' indents, tab stops and spacing in HTML (rte_linespacing2.cs writes them for
//     the cells' paragraphs too), as they already did in RTF, Word and OpenDocument files.
// Iteration 25 added to iteration 24:
//   - Paragraph spacing (rte_linespacing2.cs): space before and after paragraphs, in the Line Spacing menu (Add Space
//     Before / After Paragraph, 12 points) and dialog, kept in RTF, Word and OpenDocument files and as CSS margins in
//     HTML. Scripts: Rte.SpaceBefore := 12, Rte.SpaceAfter := 6.
//   - Word files open spaced as in Word: each paragraph's line spacing and space before and after are worked out from
//     the .docx as Word does it (document defaults, default style, a table's style winning over it in cells, the
//     paragraph's style and its own format, and "don't add space between paragraphs of the same style"), whichever
//     converter read it; the built-in reader also gets the paragraphs' indents and tab stops, which RichEditorNET's HTML
//     import dropped (DocxLineSpacing).
// Iteration 24 added to iteration 23:
//   - Print Layout (rte_printlayout.cs), on by default: lines break where they will on the printed page, as in Word's
//     Print Layout and LibreOffice (Word and OpenDocument files opened with very long lines before). RichEdit lays its
//     lines out for the printer (EM_SETTARGETDEVICE with its information context and Page Setup's text width), so the
//     editor's lines, and its line numbers, are the printed ones. The page is shown on grey, its margins as on paper.
//     Its zoom, as Word's: Page Width (the default: the page's width fills the window, as it is resized), One Page (a
//     page's height fits), or manual (Zoom In / Out / 100%, Ctrl+wheel). View > Print Layout; off: Word wrap as before.
//     Scripts: Rte.PrintLayout := 0, Rte.ZoomMode := "OnePage".
// Iteration 23 added to iteration 22:
//   - Character formats at the caret, as in Word: with nothing selected, the font box (font, size, bold, italic,
//     underline, strikethrough, superscript, subscript, colour, highlight), the Font dialog and Ctrl+B / I / U set the
//     format of what is typed next, and on an empty paragraph its paragraph mark's too, so the empty line keeps the
//     format and shows it when the caret comes back (EditorApi.AtCaret; RichEditorNET's commands changed only selected
//     text). The font box shows the caret's format (TOM's insertion point) rather than the character at the caret's.
//     Scripts get the same: Rte.SetFontName, ToggleBold, ToggleItalic, ToggleUnderline, ToggleStrikeThrough,
//     SetFontColor, SetBackgroundColor are EditorApi's now.
// Iteration 22 added to iteration 21:
//   - Line spacing (rte_linespacing.cs): Home > Paragraph > Line Spacing sets the selected paragraphs' line spacing,
//     1.0 to 3.0 lines, or Line Spacing Options for single, 1.5, double, a multiple, or at least / exactly so many points
//     (TOM's SetLineSpacing; one undo step). RichEdit keeps it in RTF, so Word and LibreOffice files have it; HTML gets
//     it as CSS line-height (also read back); the built-in .docx converter gets the exact values both ways (rte_office7.cs
//     and DocxLineSpacing, which also keeps its reader from making every paragraph 1.08 lines); Markdown has no line
//     spacing. Scripts: Rte.SetLineSpacing("1.5"), Rte.SetLineSpacing("exactly 14pt"), Rte.LineSpacing.
// Iteration 21 added to iteration 20:
//   - The Markdown and HTML views edit in Scintilla (rte_sourcebox.cs, through the Scintilla5.NET package) instead of a
//     plain-text RichEdit: Markdown and HTML syntax colours (Lexilla's lexers), line numbers (View > Source line
//     numbers), Ctrl+wheel zoom. SourceBox keeps the text box members the views use, with character positions and "\n"
//     line ends, so the views' commands are unchanged. Csharp-in-AHK runs packages from the NuGet folder, so SourceBox
//     tells Scintilla5.NET where the package's native Scintilla.dll and Lexilla.dll are.
// Iteration 20 added to iteration 19:
//   - Line numbers are printed (Print, Print Preview, Export as PDF), numbering the printed lines: rte_linenumbers2.cs
//     lays the document out as the printer does (PrintedLineNumbers: a hidden RichEdit with the printer as its target
//     device gives the lines, EM_FORMATRANGE measuring gives each line's top) and DocumentPrinter draws them left of the
//     text. Scripts: Rte.PrintedLineAt(position) gives "page:line" as printed, for citing a page and line.
//   - Export as PDF leaves the default printer alone: while Windows manages the default printer, printing to "Microsoft
//     Print to PDF" made it the default (DefaultPrinterGuard).
// Iteration 19 added to iteration 18:
//   - Line numbers (rte_linenumbers.cs), as in Word: Layout > Line Numbers numbers the lines in the left margin,
//     continuously or restarting on each page (where the printer's layout starts them), with a first number, a count-by
//     (every 5th ...) and a distance from the text (Line Numbering Options). Lines in tables and page breaks don't count,
//     nor paragraphs whose numbers are suppressed (Suppress for Current Paragraph: TOM's NoLineNumber, RTF's \noline).
//     The numbers follow the editor's lines (LineNumberGutter, a strip left of the text, drawn after the editor paints).
//     Kept in RTF (\linemod ...), Word (w:lnNumType) and OpenDocument files, and in HTML and Markdown as a comment-like
//     setting. Not printed by AhkWord yet. Scripts: Rte.LineNumbers := "continuous", Rte.SetLineNumbers("page", 1, 5),
//     Rte.SuppressLineNumbers(1), Rte.LineNumberAt(position).
// Iteration 18 added to iteration 17:
//   - Table properties (rte_tables2.cs): Table Tools > Properties (and the table's right-click menu) sets the column
//     widths, the table's alignment on the page, cell padding, row height, the borders (none, all, outside, inside ...,
//     line width and color), shading and vertical alignment of the selected cells or the whole table; Borders and
//     Shading are on the ribbon too. They are kept in RTF (and so Word and OpenDocument files) and in HTML (as CSS, also
//     read back); Markdown tables have no room for them. Word's cell backgrounds now show (RichEdit ignored them).
//     Scripts: Rte.SetColumnWidths("2in, 3in"), SetCellBorders, SetCellShading, SetTableAlignment, SetRowHeight,
//     SetCellPadding, SetCellVerticalAlignment.
// Iteration 17 added to iteration 16:
//   - Markdown and HTML views (View > Document View): a plain-text source editor (SourceBox) beside or instead of the
//     formatted text, in a SplitContainer inside the window, so the ribbon, shortcuts and automation keep working.
//     The formatted document stays the one the commands work on: entering a source view fills it from the document
//     (pictures as embedded-1.png ... instead of their data: URIs, put back when loading); edits there are loaded into
//     the document after a pause in side-by-side, when leaving the view, and before any command or save. Edits in the
//     formatted text refresh the source. Saving a .md / .html file whose document came from its source (opened from it,
//     or edited in that view since) writes the source exactly as typed, not regenerated (VerbatimSource).
//   - Line breaks (Shift+Enter) are <br> in HTML, as in other editors: in paragraphs, headings, list items and table
//     cells (rte_linebreaks.cs; RichEditorNET's HTML made each line a paragraph, and a table with one fell apart), and
//     <br> is read as one. Markdown writes them as <br> in headings and table cells too (rte_markdown6.cs).
//   - Markdown tables are written with their columns lined up (MarkdownConverter.AlignTable); in the Markdown view,
//     Tab / Shift+Tab in a table line it up and move between cells, and the Table button tidies it (FormatSourceTable);
//     tables with long cells are written compactly. In the formatted text, Tab / Shift+Tab in any table go through
//     NextCell (RichEdit's own Tab doesn't move between the cells of every table from Word), adding a row at the end.
// Iteration 16 added to iteration 15:
//   - A ribbon instead of the menu bar and toolbar: the Windows Ribbon (UIRibbon, drawn by Windows: DPI-aware, KeyTips
//     with Alt, screen readers), through the WindowsRibbon NuGet package. Its markup is ribbon\RteRibbon.xml, compiled by
//     ribbon\build_ribbon.ahk into ribbon\RteRibbon.ribbon (a resource-only DLL, loaded from the file). The ribbon's
//     buttons run the same commands the menus and toolbar did (Run(action, macro), so macro recording is unchanged);
//     the keyboard shortcuts the menus had are in a table of their own (Shortcuts, checked in ProcessCmdKey). The font
//     group is the ribbon's own font control; Table Tools appears while the caret is in a table; recent documents are
//     in the File menu, macros in a gallery on the Tools tab.
// Iteration 15 added to iteration 14:
//   - Pictures' alt text can be seen and changed: resting the mouse on a picture shows it (or "no alt text"); right-click
//     a picture > Alt text... changes it (RichEditorNET can't change it, so the picture is replaced by the same picture
//     with the new alt text, in one undo step); Insert > Image asks for it (the file name was used without asking).
// Iteration 14 added to iteration 13:
//   - The document's header and footer are shown in strips above and below the text (BandStrip), filled in as on the
//     first page ({page} is 1, {pages} the number of pages it prints on, worked out once typing pauses); clicking one
//     opens File > Header and Footer. View > Header and footer shows or hides the strips.
//   - Pictures keep their alt text through RTF (rte_office6.cs): the RTF AhkWord writes (files, Word, LibreOffice)
//     has it as each picture's description, and pictures read from RTF get it back.
// Iteration 13 added to iteration 12:
//   - Headers and footers per document (rte_headerfooter.cs): File > Header and Footer sets the document's own (and,
//     if asked, the default for new documents, which documents without one of their own get). They are printed as
//     before, and saved: as real headers and footers in RTF, and so in Word and OpenDocument files (made by Word,
//     LibreOffice or the built-in converter), with fields for {page} {pages} {date} {time}, tab stops for the left,
//     center and right parts, and the paper and margins of Page Setup; as a <meta> in HTML and a comment in Markdown.
//     Headers and footers made or changed in Word or LibreOffice are read back.
// Iteration 12 added to iteration 11:
//   - Manual page breaks (rte_pagebreaks.cs): Insert > Page break (Ctrl+Enter) puts one at the caret, as a paragraph of
//     its own, drawn as a dashed line with "Page break" (RichEdit shows nothing there). Printing starts a new page
//     after it. RTF keeps it as \page, and so Word and LibreOffice files; HTML and Markdown (rte_markdown5.cs) have it
//     as <div style="page-break-after: always; break-after: page"></div>, which browsers obey when printing, and the
//     built-in .docx converter as a Word page break (rte_office4.cs). Page breaks in documents from elsewhere (Word's
//     "\page" in the middle of a paragraph, HTML page-break styles, form feeds in text) become such paragraphs.
// Iteration 11 added to iteration 10:
//   - Exports through Word and LibreOffice (.docx, .odt) give legal lists real list numbering, so they renumber in Word
//     and LibreOffice when edited (rte_office3.cs: the RTF they get has the legal items as an RTF list).
// Iteration 10 added to iteration 9:
//   - Legal numbering (rte_legal.cs): Paragraph > Legal numbering (Ctrl+Shift+G, or the 1.1 button) numbers items 1.,
//     1.1, 1.1.1 by their list level; Tab / Shift+Tab, Enter, Backspace and Delete renumber them in the same undo step;
//     Restart / Continue numbering work on them; Paragraph > List levels sets how the numbers look (1. 1.1 / 1. 1.1. /
//     1 1.1). RichEdit cannot show compound numbers, so they are text in front of each item, kept up to date. They are
//     saved as <ol class="rte-legal"> in HTML (rte_htmllists2.cs), as [type=legal] lists in Markdown
//     (rte_markdown4.cs), as Word legal numbering by the built-in .docx converter (rte_office2.cs), and as the numbers
//     themselves in RTF (and so by Word and LibreOffice). Paragraphs that start with typed numbers such as "2.1 " take
//     them as their level when they are made legal, and opened documents with them become legal lists.
// Iteration 9 added to iteration 8:
//   - Tables (rte_tables.cs): Table > Insert table, insert and delete rows and columns, delete the table; Tab and
//     Shift+Tab move between cells; right-click in a table for the table commands; Ctrl+Shift+Insert inserts a row
//     below, Ctrl+Shift+Delete deletes the row, and Delete or Backspace deletes rows that are selected whole. They are
//     real RichEdit tables, and are saved as <table> in HTML and as GitHub tables in Markdown (rte_markdown3.cs), and
//     read back from both.
//   - Rich Text Format: open and save .rtf, which keeps everything the editor has.
//   - Word and OpenDocument (rte_office.cs): open .docx, .doc and .odt; save as .docx or .odt; File > Export as Word /
//     Export as OpenDocument (the current file stays as it is). Done through an invisible Word (only when Word isn't
//     running), LibreOffice (headless, with its own profile), the built-in Open XML converter (.docx only, NuGet
//     packages) or pandoc, whichever is available.
// Iteration 8 added to iteration 7:
//   - Markdown keeps alignment, colors, highlight, fonts, list styles and image sizes, as HTML and M-ArkDown
//     extensions (rte_markdown2.cs).
//   - Paragraph > List levels: the numbering style, punctuation and start number of each list level 1-9, used for
//     numbered lists and when nesting with Tab, and applied to the current list on request.
//   - Paragraph > Restart numbering / Continue numbering.
//   - HTML keeps list start numbers, restarts and punctuation (rte_htmllists.cs), which RichEditorNET's HTML drops.
// Iteration 7 replaced RichEditorNET's Markdown support with rte_markdown.cs:
//   - Opening and saving .md files, View > Markdown source and Rte.Markdown / GetMarkdown / SetMarkdown use real
//     Markdown: blank lines between paragraphs, list numbers, code, quotes, rules, tables and images (saved as files in
//     a "<name>_images" folder next to the document).
// Iteration 6 added macro power to iteration 5:
//   - Tools > Record macro (Ctrl+Shift+R): typing, cursor keys and the formatting commands of the menus and toolbar are
//     recorded as RteClient2 calls; stopping offers to save them as an AHK script in rte_macros, with a hotkey.
//     Mouse clicks are not recorded (positions differ from document to document).
//   - Tools > Macro hotkeys: a hotkey for any macro, active while the editor window is active.
// Iteration 5 added writing tools to iteration 4:
//   - Spelling: right-click a misspelled word (red underline) for suggestions, "delete repeated word" and "add to
//     dictionary"; F7 goes to the next spelling error. Uses the Windows spell checker, like the underlines.
//   - Outline: a panel listing the headings (View > Outline, Ctrl+Shift+O); click one to go there; right-click to
//     promote or demote it. It follows the caret and updates as you type.
// Iteration 4 added editable print headers and footers to iteration 3:
//   - File > Header and Footer: show or hide each, with left, center and right parts; the fields {title} {file} {page}
//     {pages} {date} {time} are filled in on each page. Headers and footers are drawn in the page margins.
// Iteration 3 added comfort features to iteration 2:
//   - File > Open Recent; the window size and position, view settings, last folder and page setup are remembered in
//     %APPDATA%\AhkWord\settings.json (or the file named by the AHKWORD_SETTINGS environment variable, or RTE_WRITER_SETTINGS; in
//     portable mode, data\settings.json beside the program).
//   - Drag and drop: drop a document to open it, or images to insert them where they are dropped.
//   - Page Setup, Print Preview, Print, and Export as PDF (through the "Microsoft Print to PDF" printer).
// Iteration 2 added events and fire-and-forget calls to iteration 1:
//   - Events: AHK scripts subscribe with Rte.On(Event, Callback) (RteClient2.ahk). The editor posts each event to the
//     subscriber's window as the registered message "RTE_WRITER_EVENT" (wParam = event number, lParam = event data),
//     so it never waits for a script. TextChanged and SelectionChanged are sent once typing / moving pauses briefly.
//   - Rte.Post(Name, args...) queues a call and returns at once; queued calls run in order, before any later call.
//     Errors in posted calls go to the status bar, Rte.LastAsyncError and the AsyncError event.
// Iteration 1:
//   - Documents: HTML (.html/.htm), Markdown (.md, GitHub flavor) and plain text (.txt).
//   - Formatting toolbar and menus; Tab / Shift+Tab nest list items, with automatic numbering styles per level.
//   - Automation: the public members of EditorApi, plus the main RichEditBox members, can be called by name from AHK
//     (see RteClient.ahk):
//       * from other AHK processes, through a hidden window "AhkWord automation server <pid>" (WM_COPYDATA);
//       * from the hosting AHK process, through the delegate in AppContext.GetData("rte.dispatch").
//     Request: verb, member name and arguments separated by U+001F; verb is "call", "get" or "set".
//     Reply: "0" + result as text, or "1" + error message.
//     Tools > Macros runs the AHK scripts in the rte_macros folder (with RTE_WRITER_HWND set to this editor's server).
// args: [0] AutoHotkey exe (A_AhkPath), [1] folder of AhkWord.ahk, [2] optional file to open.
using System.Collections;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.TextObjectModel2;
using RibbonLib;
using RibbonLib.Controls;
using RibbonLib.Controls.Events;
using RibbonLib.Interop;

// Runs on AutoHotkey's (STA) thread. ShowDialog runs the WinForms message loop until the window is closed;
// AHK hotkeys and timers keep running meanwhile and can call the editor (see RteClient.ahk).
// (Iteration 88: built as AhkWord.exe the program starts on an MTA thread, which WinForms can't use (drag and drop,
// the clipboard): it then runs on a new STA thread. Under AutoHotkey the thread is already STA.)
args = AppFolders.CompleteArgs?.Invoke(args) ?? args;
if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) {
	var sta = new Thread(() => RunAhkWord(args));
	sta.SetApartmentState(ApartmentState.STA);
	sta.Start();
	sta.Join();
	return;
}
RunAhkWord(args);

static void RunAhkWord(string[] args) {
	Application.EnableVisualStyles();
	AppFolders.ProgramFolder = args.Length > 1 ? args[1] : Environment.CurrentDirectory;    // (portable mode: before the settings)
	AppFolders.HostScript = args.Length > 3 && args[3] != "" ? args[3] : null;    // (AhkWord.ahk: New Window starts it again)
	var startSettings = Settings.Load();
	AppTheme.Named(startSettings.Theme).SetColorMode();    // (before anything of WinForms is made: see SetColorMode)
	Lang.Load(AppFolders.ProgramFolder, startSettings.Language);    // (AhkWord-es: before any string is made: T, rte_strings.cs)
	using var form = new WriterForm(
		args.Length > 0 ? args[0] : "",
		args.Length > 1 ? args[1] : Environment.CurrentDirectory,
		args.Length > 2 && args[2] != "" ? args[2] : null);
	form.ShowDialog();
	AHK.Return(form.FileName ?? "");
}

// The command Ids of ribbon\RteRibbon.xml.
enum RibbonCommand : uint
{
	TabHome = 100, TabInsert = 101, TabLayout = 102, TabReview = 103, TabView = 104, TabTools = 105,
	TableTools = 106, TabTableLayout = 107, GroupClipboard = 110, GroupFont = 111, GroupText = 112, GroupParagraph = 113,
	GroupStyles = 114, GroupEditing = 115, GroupPages = 116, GroupTables = 117, GroupIllustrations = 118, GroupLinks = 119, GroupContents = 135,
	GroupInsertText = 120, GroupPage = 121, GroupProofing = 122, GroupShow = 124, GroupZoom = 125, GroupDocumentView = 126,
	GroupMacros = 127, GroupScripts = 128, GroupAutomation = 129, GroupExtensions = 136, GroupRowsColumns = 130, GroupDelete = 131, GroupTableFormat = 132, GroupMerge = 137, GroupTableData = 138,
	AppMenu = 140, Qat = 141, Recent = 142, Help = 143, ExportMenu = 144, PrintMenu = 145,
	New = 200, Open = 201, Save = 202, SaveAs = 203, Print = 204, PrintPreview = 205,
	PageSetup = 206, ExportPdf = 207, ExportWord = 208, ExportOdt = 209, Exit = 210, NewFromTemplate = 212,
	NewWindow = 216, OpenInNewWindow = 217, StartPage = 215, InsertFromTemplate = 214, FillInFields = 213, GroupTemplates = 133,
	GroupHomeMacros = 134, Close = 211, Undo = 220, Redo = 221, Paste = 222, Cut = 223,
	Copy = 224, SelectAll = 225, Find = 226, Replace = 227, Font = 230, FontDialog = 231,
	ChangeCase = 232, CaseUpper = 233, CaseLower = 234, CaseTitle = 235, CaseHeadline = 397, UnderlineStyle = 236, UnderlineSingle = 237,
	UnderlineDouble = 238, UnderlineDotted = 239, UnderlineDashed = 240, UnderlineWavy = 241, Bullets = 250, Numbering = 251,
	Legal = 252, Outdent = 253, Indent = 254, AlignLeft = 255, AlignCenter = 256, AlignRight = 257,
	Justify = 258, LineSpacing = 360, LineSpacing10 = 361, LineSpacing115 = 362, LineSpacing150 = 363, LineSpacing200 = 364,
	LineSpacing250 = 365, LineSpacing300 = 366, LineSpacingOptions = 367, ParagraphDialog = 390, SpaceBefore = 368, SpaceAfter = 369,
	ListOptions = 259, ListStyle = 260, ListBullet = 261, ListDecimal = 262, ListLowerAlpha = 263, ListUpperAlpha = 264,
	ListLowerRoman = 265, ListUpperRoman = 266, ListLevels = 267, AutoLevel = 268, RestartNumbering = 269, ContinueNumbering = 270,
	ConvertTyped = 271, IndentParagraph = 272, OutdentParagraph = 273, NumberingMenu = 274, BlockStyle = 275, StylesMenu = 380,
	ModifyStyle = 381, NewStyle = 382, UpdateStyle = 383, ClearFormatting = 389, DeleteStyle = 384, StyleSetClassic = 385,
	StyleSetPlain = 386, StyleSetModern = 387, SaveStyles = 388, ApplyMyStyles = 391, CopyStyles = 394, PageBreak = 280, Table = 281,
	Picture = 282, Thumbnail = 283, Link = 284, DateTime = 285, CrossReference = 286, UpdateReferences = 287, Footnote = 288, TableOfContents = 289, Snippets = 291, HorizontalLine = 292, HeaderFooter = 290, NextSpelling = 295,
	AddToDictionary = 296, SpellCheck = 297, WordWrap = 302, PrintLayout = 370, ShowPages = 392, ListNone = 393, ZoomPageWidth = 371,
	ZoomOnePage = 372, ExportPdfButton = 374, ZoomMultiplePages = 373, ShowBands = 303, Outline = 304, ZoomIn = 305,
	ZoomOut = 306, Zoom100 = 307, ViewRich = 330, ViewMarkdown = 331, ViewHtml = 332, SideBySide = 333,
	SourceWordWrap = 334, SourceLineNumbers = 335, Macros = 310, RecordMacro = 311, MacroHotkeys = 312, RunScript = 313, EventHooks = 317, CommandPalette = 318, Extensions = 395, SectionBreak = 396, Columns = 398, ContinuousBreak = 399, Theme = 401, FormatPainter = 402, Ruler = 403, Language = 405, FontName = 410, FontSize = 411, FontBold = 412, FontItalic = 413, FontUnderline = 414, FontStrike = 415, FontSub = 416, FontSuper = 417, FontGrow = 418, FontShrink = 419, FontColor = 420, FontHighlight = 421,
	MacrosFolder = 314, AutomationApi = 315, About = 316, InsertAbove = 320, InsertBelow = 321, InsertLeft = 322,
	InsertRight = 323, DeleteRow = 324, DeleteColumn = 325, DeleteTable = 326, MergeCells = 327, SplitCells = 328, DeleteCells = 329, UpdateTable = 422, SaveTableToSource = 423, LinkTable = 424, UnlinkTable = 425, TableSelect = 426, SelectCell = 427, SelectColumn = 428, SelectRow = 429, SelectTable = 430, BordersTop = 431, BordersBottom = 432, BordersLeft = 433, BordersRight = 434, TableProperties = 340, TableBorders = 341,
	BordersAll = 342, BordersOutside = 343, BordersInside = 344, BordersNone = 345, TableShading = 346, LineNumbers = 350,
	LineNumbersNone = 351, LineNumbersContinuous = 352, LineNumbersEachPage = 353, SuppressLineNumbers = 354, LineNumberOptions = 355,
}

// Iteration 17: HTML with line breaks and indentation where they don't change what it shows: between blocks
// (paragraphs, headings, list items, table rows and cells), and after <br>. Text inside a block stays on its line
// (a line break there would be a space), and <pre> is left as it is.
static class PrettyHtml
{
	static readonly HashSet<string> Containers = new(StringComparer.OrdinalIgnoreCase) { "ul", "ol", "table", "thead", "tbody", "tfoot", "tr", "blockquote", "div" };
	static readonly HashSet<string> Leaves = new(StringComparer.OrdinalIgnoreCase) { "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "td", "th", "hr", "caption" };
	static readonly Regex Tag = new(@"<(/?)([a-zA-Z][a-zA-Z0-9]*)\b[^>]*?(/?)>", RegexOptions.Compiled);

	public static string Format(string html) {
		var output = new StringBuilder();
		int depth = 0, position = 0;
		bool lineStart = true;
		void NewLine() { if (!lineStart) { output.Append("\r\n"); lineStart = true; } }
		void Indent() { if (lineStart) { output.Append(' ', depth * 2); lineStart = false; } }
		foreach (Match tag in Tag.Matches(html)) {
			if (tag.Index < position) continue;    // (inside a <pre>, already copied)
			string text = html[position..tag.Index];
			if (text.Trim() != "") { Indent(); output.Append(text); }
			else if (!lineStart) output.Append(text);
			position = tag.Index + tag.Length;
			string name = tag.Groups[2].Value;
			bool closing = tag.Groups[1].Value == "/", selfClosing = tag.Groups[3].Value == "/";
			if (name.Equals("pre", StringComparison.OrdinalIgnoreCase) && !closing) {    // as it is, up to </pre>
				int end = html.IndexOf("</pre>", position, StringComparison.OrdinalIgnoreCase);
				end = end < 0 ? html.Length : end + "</pre>".Length;
				NewLine(); Indent();
				output.Append(html, tag.Index, end - tag.Index);
				position = end;
				NewLine();
				continue;
			}
			if (Containers.Contains(name) && !closing && Regex.Match(html[position..], @"^\s*</" + name + @"\s*>", RegexOptions.IgnoreCase) is { Success: true } empty) {
				NewLine(); Indent(); output.Append(tag.Value).Append(empty.Value.Trim()); NewLine();    // an empty one (a page break) on one line
				position += empty.Length;
			} else if (Containers.Contains(name)) {
				if (closing) { depth = Math.Max(0, depth - 1); NewLine(); Indent(); output.Append(tag.Value); NewLine(); }
				else { NewLine(); Indent(); output.Append(tag.Value); NewLine(); if (!selfClosing) depth++; }
			} else if (Leaves.Contains(name)) {
				if (closing) { Indent(); output.Append(tag.Value); NewLine(); }
				else { NewLine(); Indent(); output.Append(tag.Value); if (selfClosing || name.Equals("hr", StringComparison.OrdinalIgnoreCase)) NewLine(); }
			} else {
				Indent();
				output.Append(tag.Value);
				if (name.Equals("br", StringComparison.OrdinalIgnoreCase)) NewLine();
			}
		}
		string rest = html[position..];
		if (rest.Trim() != "") { Indent(); output.Append(rest); }
		return output.ToString().TrimEnd();
	}
}

class WriterForm : Form
{
	public const string AppName = "AhkWord";
	public const int Iteration = 94;    // (the About box; make_version.ps1)
	public readonly WriterEditBox Editor = new() { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 11f), BorderStyle = BorderStyle.None };
	public string? FileName { get; private set; }
	public bool AutoLevelStyles = true;

	readonly string ahkPath, appDir;
	readonly EditorApi api;
	readonly AutomationServer server;
	readonly ToolStripStatusLabel statusText = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
	readonly ToolStripStatusLabel positionLabel = new(), wordsLabel = new(), modeLabel = new(), automationLabel = new();
	readonly System.Windows.Forms.Timer wordCountTimer = new() { Interval = 400 };

	// Iteration 16: the ribbon (built in BuildRibbon; UpdateUi makes it follow the selection) and the keyboard shortcuts
	// that were the menus' (ProcessCmdKey).
	readonly Ribbon ribbon = new();
	readonly Dictionary<Keys, (string Name, Action Action)> shortcuts = new();
	readonly Dictionary<RibbonCommand, RibbonToggleButton> toggleButtons = new();
	readonly Dictionary<RibbonCommand, RibbonCheckBox> checkBoxes = new();
	RibbonFontControl fontControl = null!;
	RibbonInRibbonGallery blockCombo = null!;
	RibbonDropDownGallery macrosGallery = null!;
	RibbonDropDownGallery extensionsGallery = null!;
	RibbonRecentItems recentItems = null!;
	RibbonTabGroup tableTools = null!;
	RibbonDropDownColorPicker shadingPicker = null!;
	RibbonButton undoButton = null!, redoButton = null!, cutButton = null!, copyButton = null!;
	List<string> galleryMacros = new();

	// Iteration 17: the Markdown and HTML views. The source editor is beside (side by side) or instead of the page; the
	// formatted document stays the one the commands work on, and the source is loaded into it (ApplySource).
	internal enum SourceKind { None, Markdown, Html }
	readonly SplitContainer sourceSplit = new() { Dock = DockStyle.Fill, Panel1Collapsed = true, SplitterWidth = 6 };
	readonly SourceBox sourceBox = new() { Dock = DockStyle.Fill };
	readonly StatusStrip statusBar = new();
	readonly Label sourceTitle = new() { Dock = DockStyle.Top, AutoEllipsis = true, ForeColor = SystemColors.GrayText, BackColor = SystemColors.Control };
	readonly System.Windows.Forms.Timer sourceTimer = new() { Interval = 600 }, richToSourceTimer = new() { Interval = 600 };
	SourceKind view = SourceKind.None, shownKind = SourceKind.None;    // the view; what the source editor holds
	bool sideBySide, sourceDirty, filling, loadingSource, inCommand, sourceHadFocus;
	// For each kind, the source text the document corresponds to (opened from it, or loaded from or shown in the source
	// view), and what the document regenerated as that kind was then: while it still is, the document is unchanged
	// since, and the text is shown, and saved, as it is (a .md file keeps the user's own Markdown).
	// Own: the text is the opened file's or typed by the user (saved as it is), not regenerated from the document.
	readonly Dictionary<SourceKind, (string Text, string Snapshot, bool Own)> sources = new();
	string sourceNewline = "\n";
	// Pictures: their data: URIs in the source view are names (embedded-1.png ...), put back when loading.
	readonly Dictionary<string, string> embeddedData = new(), embeddedName = new();
	// Formatting in the source editor: Markdown or HTML for ribbon commands, and for keys that are not commands' own.
	readonly Dictionary<RibbonCommand, Action> sourceCommands = new();
	readonly Dictionary<Keys, RibbonCommand> keyCommands = new();
	readonly Dictionary<Keys, Action> sourceKeys = new();
	FindReplaceForm? findForm;
	bool updatingUi, discardOnClose, lastModified;

	// Events for AHK scripts, and calls queued by Rte.Post
	internal new readonly EventHub Events = new();

	// Iteration 3: settings, recent files, drag and drop
	internal readonly Settings Settings = Settings.Load();
	// (the recovery copies are kept beside the settings)
	readonly Recovery recovery = new(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(Settings.FilePath))!, "Recovery"));
	readonly System.Windows.Forms.Timer autosaveTimer = new();

	// Iteration 5: spelling and outline
	internal readonly Spelling? Spelling = Spelling.Create();

	// Iteration 6: macro recording and hotkeys
	readonly MacroRecorder recorder = new();
	readonly ToolStripStatusLabel recordingLabel = new() { ForeColor = AppTheme.Warning, Visible = false, IsLink = false, ToolTipText = T("ui.recordingTip", "Recording a macro - click to stop") };
	readonly SplitContainer outlineSplit = new() { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
	readonly TreeView outlineTree = new() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, HideSelection = false, FullRowSelect = true, ShowLines = false, ShowPlusMinus = true };
	readonly Label outlineHint = new() { Dock = DockStyle.Fill, ForeColor = SystemColors.GrayText, Padding = new Padding(6), Visible = false,
		Text = T("ui.outlineEmpty", "No headings yet.\n\nGive paragraphs a heading style (Home > Styles, Ctrl+1 to Ctrl+6) and they appear here.") };
	readonly System.Windows.Forms.Timer outlineTimer = new() { Interval = 400 };
	List<(int Level, int Start, string Text)> outline = new();
	bool updatingOutline, outlineBuilt;
	FileDropTarget? dropTarget;
	readonly System.Windows.Forms.Timer textEventTimer = new() { Interval = 200 }, selectionEventTimer = new() { Interval = 120 };
	readonly Queue<string[]> postedCalls = new();

	// Iteration 10: legal numbering
	readonly System.Windows.Forms.Timer legalTimer = new() { Interval = 300 };

	// Iteration 19: line numbers, in a strip left of the text (in the page, between its padding and the editor)
	readonly LineNumberGutter lineGutter;
	readonly Panel pagePanel;
	List<int>? pageStarts;          // where printed pages start, for numbering that restarts on each page
	bool pageStartsStale = true;    // the text or the page changed since they were worked out
	bool skipLegalCheck, guardingCaret, guardPending;
	int lastCaret = -1;

	/// Iteration 41: AhkWord's icon (images\icon\ahkword.ico, made by images\icon\make_icon.py); null if it isn't there
	/// or can't be read (the window then has .NET's).
	internal static Icon? AppIcon(string appDir) {
		string path = Path.Combine(appDir, "images", "icon", "ahkword.ico");
		try { return File.Exists(path) ? new Icon(path) : null; }
		catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException) { return null; }
	}

	// (its window: AhkWord's own on the taskbar, its icon, and pinned it starts AhkWord: TaskbarIdentity)
	protected override void OnHandleCreated(EventArgs e) {
		base.OnHandleCreated(e);
		string icon = Path.Combine(appDir, "images", "icon", "ahkword.ico");
		string script = AppFolders.HostScript is { } host && File.Exists(host) ? host : Path.Combine(appDir, "AhkWord.ahk");
		TaskbarIdentity.SetWindow(Handle, AppFolders.ProgramExe is { } exe ? $"\"{exe}\" \"{ahkPath}\" \"{appDir}\"" : $"\"{ahkPath}\" \"{script}\"",
			File.Exists(icon) ? icon : null);
	}

	public WriterForm(string ahkPath, string appDir, string? startFile) {
		AppTheme.OnPaper(Editor);
		AppTheme.OnPaper(headerStrip);
		AppTheme.OnPaper(footerStrip);
		recordingLabel.ForeColor = AppTheme.Warning;
		TaskbarIdentity.SetProcess();    // (before its first window)
		this.ahkPath = ahkPath;
		this.appDir = appDir;
		DefaultPrinterGuard.Restore(Settings);    // (if a PDF export was interrupted)
		Bands = DefaultBands();
		Text = AppName;
		if (AppIcon(appDir) is { } icon) Icon = icon;    // (the title bar's and the taskbar's)
		StartPosition = FormStartPosition.CenterScreen;
		Size = new Size(LogicalToDeviceUnits(1150), LogicalToDeviceUnits(820));
		Editor.EnableHtmlFontSizing = false;    // (any font size: styles do what "structured documents" did, iteration 34)
		// Markdown: what a paragraph is, by its style (Code; a heading's level; else body text), not by its font
		MarkdownConverter.StyleNameOf = (doc, position) => StyleEngine.StyleAt(doc, Styles, position) is { Numbered: true } numbered ? numbered.Name : null;    // (iteration 50)
		MarkdownConverter.BlockOf = (doc, position) => string.Equals(StyleEngine.StyleAt(doc, Styles, position).Name, "Code", StringComparison.OrdinalIgnoreCase) ? (int)BlockStyle.Preformatted : StyleEngine.LevelAt(doc, Styles, position);
		Editor.WordWrap = Settings.WordWrap;
		Editor.EnableSpellCheck = Settings.SpellCheck;
		AutoLevelStyles = Settings.AutoLevelStyles;
		ApplyWindowSettings();

		api = new EditorApi(this);
		lineGutter = new LineNumberGutter(Editor);
		var page = pagePanel = new BufferedPanel { Dock = DockStyle.Fill, BackColor = AppTheme.Paper };
		pageGaps = new PageGaps(Editor);
		Editor.Gaps = pageGaps;
		Editor.BandParts = GapBandParts;
		Editor.BandStyleOf = (header, page) => header ? (Bands.HeaderFontOn(page), Bands.HeaderFromEdge, Bands.HeaderLineOn(page)) : (Bands.FooterFontOn(page), Bands.FooterFromEdge, Bands.FooterLineOn(page));
		pagePanel.Paint += (s, e) => PaintGapBands(e.Graphics);
		Editor.Painted += FollowGapBands;
		FitPagePadding();
		page.Controls.Add(Editor);
		page.Controls.Add(lineGutter);     // (docked before the editor, after the header and footer)
		page.Controls.Add(headerStrip);    // (added after the editor, so docked before it)
		page.Controls.Add(footerStrip);
		Editor.Painted += lineGutter.EditorPainted;
		Editor.Painted += () => editorPainted = DateTime.Now;
		// (Show Pages' bands across the line numbers too: the panel's, moved to the numbers' own place)
		lineGutter.BandColor = PrintLayoutPage.Surround;
		lineGutter.Bands = () => PanelBands().Select(b => new Rectangle(b.X - lineGutter.Left, b.Y - lineGutter.Top, b.Width, b.Height));
		lineGutter.SizeChanged += (s, e) => FitPagePadding();
		// (iteration 87) the sections' own line numbering
		lineGutter.SectionNumbering = () => Sections.Resolve(Editor.TextDocument, Bands, DocumentColumns).Select(s => (s.Start, s.LineNumbers)).ToList();
		// Print Layout: the page follows the window, the zoom (Ctrl+wheel too) and a new editor window
		pageHost.Resize += (s, e) => FitPagePadding();
		Editor.Painted += () => {
			if (pageLayout == null) return;
			if (Math.Abs(EditorZoom() - pageZoom) > 0.001f) {
				// a zoom not made by Page Width or One Page (Ctrl+wheel): the zoom is the user's from now on
				if (Settings.ZoomMode != "Manual" && Math.Abs(EditorZoom() - autoZoom) > 0.01f) Settings.ZoomMode = "Manual";
				pageOverhang = 0; overhangStale = true; FitPagePadding(); UpdateUi();
			}
			CheckPageOverhang();
		};
		Editor.TextChanged += (s, e) => overhangStale = true;
		Editor.VScroll += (s, e) => overhangStale = true;
		Editor.HandleCreated += (s, e) => { if (!Editor.Renewing) BeginInvoke(() => ApplyPrintLayout()); };    // (wider numbers: the text stays where it was, as far as it can; FreshEditor applies it itself)
		headerStrip.Click += (s, e) => Run(HeaderFooterCommand);
		footerStrip.Click += (s, e) => Run(HeaderFooterCommand);
		Editor.TextChanged += (s, e) => {
			bool typed = Editor.Typing && (typedOnly || !bandPagesStale);
			bandPagesStale = pageStartsStale = true;
			typedOnly = typed;
			lastEdit = DateTime.Now;
		};
		bandTimer.Tick += (s, e) => UpdateBands();
		bandTimer.Start();
		var outlineTitle = new Label { Text = T("ui.outline", "Outline"), Dock = DockStyle.Top, Height = LogicalToDeviceUnits(26), Padding = new Padding(6, 6, 0, 0),
			Font = new Font(Font, FontStyle.Bold), ForeColor = SystemColors.GrayText };
		outlineSplit.Panel1.BackColor = outlineTree.BackColor = outlineHint.BackColor = SystemColors.Control;
		outlineSplit.Panel1.Controls.Add(outlineTree);
		outlineSplit.Panel1.Controls.Add(outlineHint);
		outlineSplit.Panel1.Controls.Add(outlineTitle);
		var sourcePanel = new Panel { Dock = DockStyle.Fill };
		sourcePanel.Controls.Add(sourceBox);
		sourcePanel.Controls.Add(sourceTitle);    // (added after the editor, so docked before it)
		sourceTitle.Height = LogicalToDeviceUnits(24);
		sourceTitle.Padding = new Padding(LogicalToDeviceUnits(6), LogicalToDeviceUnits(4), 0, 0);
		sourceBox.WordWrap = Settings.SourceWordWrap;
		sourceBox.LineNumbers = Settings.SourceLineNumbers;
		sourceSplit.Panel1.Controls.Add(sourcePanel);
		pageHost.Controls.Add(page);
		sourceSplit.Panel2.Controls.Add(pageHost);
		// (iteration 78) the ruler above the page, outside what scrolls down; it follows the page sideways and the zoom
		ruler.Height = LogicalToDeviceUnits(22);
		ruler.Visible = Settings.ShowRuler;
		sourceSplit.Panel2.Controls.Add(ruler);    // (added after the page, so docked before it)
		Editor.Painted += () => FollowRuler();
		pagePanel.LocationChanged += (s, e) => FollowRuler();
		pagePanel.SizeChanged += (s, e) => FollowRuler();
		Editor.SizeChanged += (s, e) => FollowRuler();
		ruler.VisibleChanged += (s, e) => FollowRuler(true);
		ruler.TextClicked += RulerTab;    // (iteration 83)
		ruler.PartDoubleClicked += RulerDoubleClicked;
		ruler.Dragged += RulerDragged;    // (iteration 84)
		ruler.Dragging += RulerDragging;
		outlineSplit.Panel2.Controls.Add(sourceSplit);
		sourceBox.TextChanged += (s, e) => SourceTextChanged();
		sourceBox.SelectionChanged += (s, e) => UpdateUi();
		sourceBox.Enter += (s, e) => { sourceHadFocus = true; UpdateUi(); };
		Editor.Enter += (s, e) => { sourceHadFocus = false; UpdateUi(); };
		sourceTimer.Tick += (s, e) => { sourceTimer.Stop(); if (sideBySide) ApplySource(); };
		richToSourceTimer.Tick += (s, e) => { richToSourceTimer.Stop(); RefreshSource(); };
		Editor.TextChanged += (s, e) => {
			// edits in the formatted text (typed, or made by a command or a script) refresh the source after a pause;
			// changes made while loading the source, or by the editor's own timers, don't
			if (view != SourceKind.None && !loadingSource && !filling && (Editor.Focused || inCommand)) Restart(richToSourceTimer);
		};
		sourceSplit.SplitterMoved += (s, e) => { if (sideBySide && sourceSplit.Width > 0) Settings.SourceSplit = (int)Math.Round(sourceSplit.SplitterDistance * 100.0 / sourceSplit.Width); };
		outlineSplit.Panel1Collapsed = !Settings.ShowOutline;
		var status = statusBar;
		status.Items.AddRange(new ToolStripItem[] { recordingLabel, statusText, extensionLabel, positionLabel, wordsLabel, modeLabel, automationLabel });
		recordingLabel.Font = new Font(status.Font, FontStyle.Bold);
		recordingLabel.Click += (s, e) => Run(ToggleRecording);
		BuildRibbon();
		Controls.Add(outlineSplit);    // docking order: the last control added is docked first
		Controls.Add(ribbon);
		Controls.Add(status);
		if (AppTheme.DarkWindows) statusBar.Renderer = new FlatStripRenderer();
		PaintFrame(AppTheme.Named(Settings.Theme));
		Activated += (s, e) => FillMacrosGallery();    // macros may have been added or removed meanwhile

		// Automation: in-process (AppContext) and cross-process (hidden window) entry points.
		AppContext.SetData("rte.dispatch", (Func<string[], string>)DispatchFromAutomation);
		server = new AutomationServer(DispatchFromAutomation);
		automationLabel.Text = T("ui.automationReady", "Automation: ready");
		automationLabel.ToolTipText = T("ui.automationTip", "AHK scripts can control this editor (RteClient2.ahk). Server window 0x{0:X}", server.Handle.ToInt64());

		Editor.SelectionChanged += (s, e) => {
			if (pageGaps.Applied && !quietGaps) pageGaps.TypingFormat();    // (typed in a raised line: raised as it is)
			// the caret guard looks once the command that moved the caret has finished (commands move it on the way)
			if (!guardPending && IsHandleCreated) { guardPending = true; BeginInvoke(() => { guardPending = false; GuardLegalCaret(); }); }
			UpdateUi();
			Restart(selectionEventTimer);
		};
		Editor.TextChanged += (s, e) => {
			wordCountTimer.Stop(); wordCountTimer.Start(); UpdateTitle(); Restart(textEventTimer); Restart(outlineTimer); Restart(legalTimer);
			// (Enter at a page's first paragraph gave both paragraphs its gap: the first keeps it)
			if (pageGaps.Applied && !quietGaps && !gapsOut && pageGaps.Tidy(check: true)) QuietGaps(() => pageGaps.Tidy());
		};
		legalTimer.Tick += (s, e) => { legalTimer.Stop(); CheckLegalNumbers(); };
		LegalLists.Format = Enum.TryParse<LegalLists.Periods>(Settings.LegalFormat, true, out var legalFormat) ? legalFormat : LegalLists.Periods.Top;
		LegalLists.StyleNumbering = id => Styles.ById(id) is { Numbered: true } numbered ? (numbered.Numbering, Math.Clamp(numbered.NumberLevel, 1, 9) - 1) : null;
		LegalLists.StyleFormat = (id, level) => Styles.ById(id) is { Numbered: true } numbered ? numbered.FormatAt(level - (Math.Clamp(numbered.NumberLevel, 1, 9) - 1)) : null;    // (and its sub-levels)    // (numbered styles)
		// (the document's list levels' indents, for lists, legal numbering and the exports' lists)
		Levels = DefaultLevels();
		LegalLists.LevelIndents = level => level < Levels.Count ? (Levels[level].NumberAt, Levels[level].TextAt) : (null, null);
		Editor.ContextMenuHook = ShowContextMenu;
		Editor.OpenLink = OpenLinkCommand;
		// (iteration 50) cross-references: their numbers follow the renumbering, in the same undo step; Ctrl+click goes to
		// the paragraph referred to
		// (iteration 51: and the numbered styles' sub-levels' spacing, ParagraphStyle.SubLevelSpace)
		LegalLists.AfterRenumber = doc => StyleEngine.FitSubLevelSpacing(doc, Styles) | CrossRefs.Update(doc) > 0;
		// (iteration 51) footnotes: the note in a tooltip over its number; Ctrl+click it to change it; each sheet's notes
		// at its foot (from the footnotes, read again once the text changed, and the pages as last worked out)
		Editor.NoteAt = position => DocumentFootnotes().FirstOrDefault(f => position >= f.Start && position <= f.End)?.Text is { } note ? Footnotes.Plain(note) : null;
		Editor.TextChanged += (s, e) => { footnotesRead = null; leadersRead = null; framesRead = null; };
		Editor.ClearCellBlock = () => Run(() => api.ClearCells(), M("ClearCells"));    // (iteration 92)
		Editor.Frames = () => framesRead ??= ParagraphFrames.All(Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "");    // (iteration 87)
		// (iteration 52) the tables of contents' leaders, which RichEdit doesn't draw
		Editor.Leaders = () => leadersRead ??= Toc.Leaders(Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "");
		Editor.NotesOn = page => page >= 1 && page <= sheetNotes.Count && sheetNotes[page - 1].Count > 0 ? (sheetNotes[page - 1], sheetNoteFont.Name, sheetNoteFont.Size) : null;
		Editor.FollowReference = GoToReferenceTarget;
		Editor.AltTextOf = position => api.PictureAltText(position);
		MarkdownConverter.AltTextAt = position => { try { return api.PictureAltText(position); } catch (Exception) { return ""; } };    // (iteration 55: Insert > Horizontal Line's rule)
		Editor.KeyDown += (s, e) => { if (recorder.Recording) RecordKeyDown(e); };
		Editor.KeyPress += (s, e) => {
			if (recorder.Recording && (e.KeyChar >= ' ' || e.KeyChar == '\t')) { recorder.Type(e.KeyChar); UpdateRecordingLabel(); }
		};
		Editor.MouseDown += (s, e) => {
			if (recorder.Recording && e.Button == MouseButtons.Left && painting == Painting.Off)
				SetStatus(T("status.mouseNotRecorded", "Mouse clicks are not recorded: move with the keyboard (arrows, Home/End, Ctrl to jump, Shift to select)."));
		};
		Editor.MouseUp += (s, e) => { if (painting != Painting.Off && e.Button == MouseButtons.Left && ModifierKeys != Keys.Control) PaintSelection(); };
		outlineTimer.Tick += (s, e) => { outlineTimer.Stop(); RefreshOutline(); };
		outlineTree.NodeMouseClick += (s, e) => { if (e.Button == MouseButtons.Left && e.Node?.Tag is int start) Run(() => GoToPosition(start)); };
		outlineTree.KeyDown += (s, e) => {
			if (e.KeyCode == Keys.Enter && outlineTree.SelectedNode?.Tag is int start) { e.SuppressKeyPress = true; Run(() => GoToPosition(start)); }
		};
		outlineTree.NodeMouseClick += (s, e) => { if (e.Button == MouseButtons.Right && e.Node != null) ShowOutlineMenu(e.Node, outlineTree.PointToScreen(e.Location)); };
		textEventTimer.Tick += (s, e) => { textEventTimer.Stop(); Events.Raise(EditorEvent.TextChanged, Editor.TextLength); };
		selectionEventTimer.Tick += (s, e) => {
			selectionEventTimer.Stop();
			HighlightOutline();
			Events.Raise(EditorEvent.SelectionChanged, (uint)Editor.SelectionStart | ((long)Editor.SelectionLength << 32));
		};
		Editor.InsertHyperlinkClicked += (s, e) => InsertHyperlinkDialog();
		Editor.InsertImageClicked += (s, e) => InsertImageDialog();
		Editor.InsertLinkedThumbnailClicked += (s, e) => InsertThumbnailDialog();
		wordCountTimer.Tick += (s, e) => { wordCountTimer.Stop(); UpdateWordCount(); };

		// Drag and drop: on the text, our drop target sits in front of RichEdit's (which would embed a dropped file as an
		// icon); elsewhere in the window, the form handles the drop.
		Editor.HandleCreated += (s, e) => { if (!Editor.Renewing) Editor.BeginInvoke(RegisterDropTarget); };
		AllowDrop = true;
		DragEnter += (s, e) => e.Effect = e.Data?.GetData(DataFormats.FileDrop) is string[] files && CanDrop(files) ? DragDropEffects.Copy : DragDropEffects.None;
		DragDrop += (s, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] files) DropFiles(files, null); };

		Shown += (s, e) => {
			try { Editor.ZoomFactor = Settings.Zoom; } catch (ArgumentException) { }
			ApplyPrintLayout();
			outlineSplit.SplitterDistance = Math.Clamp(LogicalToDeviceUnits(Settings.OutlineWidth), LogicalToDeviceUnits(100), Math.Max(LogicalToDeviceUnits(100), outlineSplit.Width / 2));
			RefreshOutline();
			if (startFile != null) Run(() => api.Open(startFile));
			// (iteration 55) else the first document as File > New makes one: in the Normal style (its font: Times New
			// Roman in the Classic set, not the editor's own Segoe UI, which new settings showed)
			else { StartingUp = true; try { Run(api.New); } finally { StartingUp = false; } }
			Editor.Modified = false;
			Editor.Focus();
			AhkWordExtensions.Started(this);    // (iteration 60: what a build with more features adds; nothing in AhkWord)
			BeginInvoke(() => RunHook(HookEvent.Started));    // (iteration 56: once the window is ready)
			if (Lang.TakeError() is string langError)    // (AhkWord-es: a language file that couldn't be read)
				BeginInvoke(() => MessageBox.Show(this, langError, "AhkWord", MessageBoxButtons.OK, MessageBoxIcon.Warning));
			UpdateUi();
			UpdateTitle();
			UpdateWordCount();
			UpdateRecentItems();
			if (Enum.TryParse<SourceKind>(Settings.SourceView, out var lastView) && lastView != SourceKind.None) SetView(lastView, Settings.SideBySide);
			SetStatus(startFile != null ? "" : T("status.newDocument", "New document. Press Alt for the ribbon's KeyTips. View > Document View edits it as Markdown or HTML; the Tools tab runs AHK macros."));
			autosaveTimer.Tick += (s, e) => Autosave();
			StartAutosave();
			BeginInvoke(() => {
				if (!OfferRecovery(startFile != null) && startFile == null && Settings.ShowStartPage) ShowStart(true);
			});
		};
		FormClosing += OnFormClosing;
		FormClosed += (s, e) => {
			autosaveTimer.Stop();
			ClearRecovery();    // (closed as the user wanted: saved, or the changes discarded)
			Events.Raise(EditorEvent.Closed);
			AppContext.SetData("rte.dispatch", null);
			server.Dispose();
			findForm?.Dispose();
		};
	}

	// ---------------------------------------------------------------- Iteration 16: the ribbon

	string RibbonFile => RibbonFor(AppTheme.Named(Settings.Theme));

	// (iteration 70) The Dark and Color themes have a ribbon of their own, with light icons.
	string RibbonFor(AppTheme theme) {
		string dark = Path.Combine(appDir, "ribbon", "RteRibbonDark.ribbon");
		return theme.Dark && File.Exists(dark) ? dark : Path.Combine(appDir, "ribbon", "RteRibbon.ribbon");
	}

	// Connects the commands of ribbon\RteRibbon.xml to what they do. Each runs through Run (so it is recorded in a macro
	// as the menus' and toolbar's were); keys are the shortcuts the menus had, handled in ProcessCmdKey.
	void BuildRibbon() {
		if (!File.Exists(RibbonFile))
			throw new FileNotFoundException($"The ribbon is missing: {RibbonFile}. Build it with ribbon\\build_ribbon.ahk (it needs the Windows SDK and the Visual Studio C++ tools).");
		ribbon.ViewCreated += (s, e) => PaintTheme();
		ribbonLoaded = RibbonFile;
		ribbon.ResourceName = "file://" + ribbonLoaded;

		// File menu
		Button(RibbonCommand.New, () => { if (ConfirmDiscard()) api.New(); }, Keys.Control | Keys.N);
		Button(RibbonCommand.Open, OpenDialog, Keys.Control | Keys.O);
		Button(RibbonCommand.Save, () => SaveCommand(), Keys.Control | Keys.S);
		Button(RibbonCommand.SaveAs, () => SaveAsDialog(), Keys.Control | Keys.Shift | Keys.S);
		Button(RibbonCommand.Print, PrintCommand, Keys.Control | Keys.P);
		Button(RibbonCommand.PrintPreview, PrintPreviewCommand);
		Button(RibbonCommand.PageSetup, PageSetupCommand);
		Button(RibbonCommand.ExportPdf, ExportPdfDialog);
		Button(RibbonCommand.ExportPdfButton, ExportPdfDialog);    // (the same, on the Layout tab and the Quick Access Toolbar)
		Button(RibbonCommand.ExportWord, () => ExportOfficeDialog("docx"));
		Button(RibbonCommand.ExportOdt, () => ExportOfficeDialog("odt"));
		Button(RibbonCommand.Exit, Close);
		Button(RibbonCommand.Close, CloseDocumentCommand, Keys.Control | Keys.W);
		Button(RibbonCommand.NewFromTemplate, NewFromTemplateCommand);
		Button(RibbonCommand.StartPage, () => ShowStart(true));
		Button(RibbonCommand.NewWindow, () => { StartWindow(null); SetStatus(T("status.openingNewWindow", "Opening a new AhkWord window...")); });
		Button(RibbonCommand.OpenInNewWindow, OpenInNewWindowDialog);
		Button(RibbonCommand.InsertFromTemplate, InsertTemplateCommand);
		Button(RibbonCommand.FillInFields, FillFieldsCommand);
		Button(RibbonCommand.Help, ShowAbout);
		recentItems = new RibbonRecentItems(ribbon, (uint)RibbonCommand.Recent);
		recentItems.ExecuteEvent += (s, e) => {
			if (RecentItemsEventArgs.Create(recentItems, e).SelectedItem is { } item && item.SelectedItemIndex >= 0 && item.SelectedItemIndex < Settings.RecentFiles.Count) {
				string path = Settings.RecentFiles[item.SelectedItemIndex];
				BeginInvoke(() => Run(() => OpenRecent(path)));
			}
		};

		// Home: clipboard and editing
		// (these act on the source editor while it is the one in use; in it, their keys are its own)
		undoButton = Button(RibbonCommand.Undo, OnText(() => { skipLegalCheck = true; Editor.Undo(); }, sourceBox.Undo), Keys.Control | Keys.Z, M("Undo"));
		redoButton = Button(RibbonCommand.Redo, OnText(() => { skipLegalCheck = true; Editor.Redo(); }, sourceBox.Redo), Keys.Control | Keys.Y, M("Redo"));
		// (copy and paste keep the paragraphs' styles: StyleClipboard)
		// (iteration 93: a block of cells is copied as a table of its own; cut, its cells emptied)
		Action cut = OnText(() => { if (CopyCellBlock()) { api.ClearCells(); return; } StyleClipboard.Copy(Editor, Styles, () => WithoutGapsChanging(Editor.Cut)); }, sourceBox.Cut),
			copy = OnText(() => { if (!CopyCellBlock()) StyleClipboard.Copy(Editor, Styles, () => WithoutGaps(Editor.Copy)); }, sourceBox.Copy),
			paste = OnText(PasteText, () => sourceBox.Paste(DataFormats.GetFormat(DataFormats.UnicodeText)));
		(cutAction, copyAction, pasteAction) = (cut, copy, paste);    // (the right-click menu's too)
		cutButton =Button(RibbonCommand.Cut, cut, Keys.Control | Keys.X, M("Cut"));
		copyButton = Button(RibbonCommand.Copy, copy, Keys.Control | Keys.C, M("Copy"));
		Button(RibbonCommand.Paste, paste, Keys.Control | Keys.V, M("Paste"));
		Shortcut(Keys.Shift | Keys.Delete, "Cut", cut, M("Cut"));
		Shortcut(Keys.Control | Keys.Insert, "Copy", copy, M("Copy"));
		Shortcut(Keys.Shift | Keys.Insert, "Paste", paste, M("Paste"));
		pastePlainAction = OnText(PastePlainText, () => sourceBox.Paste(DataFormats.GetFormat(DataFormats.UnicodeText)));
		Shortcut(Keys.Control | Keys.Shift | Keys.V, "Paste as Plain Text", pastePlainAction, M("PastePlainText"));
		Button(RibbonCommand.SelectAll, OnText(Editor.SelectAll, sourceBox.SelectAll), Keys.Control | Keys.A, M("SelectAll"));
		Button(RibbonCommand.Find, () => ShowFind(false), Keys.Control | Keys.F);
		Button(RibbonCommand.Replace, () => ShowFind(true), Keys.Control | Keys.H);
		Shortcut(Keys.F3, "Find Next", () => FindNext());

		// Home: font. The font control's B, I, U ... buttons toggle, as the toolbar's did.
		BuildDarkFont();
		fontControl = new RibbonFontControl(ribbon, (uint)RibbonCommand.Font);
		fontControl.ExecuteEvent += (s, e) => {
			var changes = FontControlEventArgs.Create(fontControl, e).ChangedFontValues;
			BeginInvoke(() => { ribbonSent.Clear(); LeaveOverlays(); ApplyFontChanges(changes); Editor.Focus(); });    // (ribbonSent: Command)
		};
		Shortcut(Keys.Control | Keys.B, "Bold", api.ToggleBold, M("ToggleBold"));
		Shortcut(Keys.Control | Keys.I, "Italic", api.ToggleItalic, M("ToggleItalic"));
		Shortcut(Keys.Control | Keys.U, "Underline", () => api.ToggleUnderline(), M("ToggleUnderline"));
		Shortcut(Keys.Control | Keys.Shift | Keys.Oemplus, "Superscript", api.ToggleSuperscript, M("ToggleSuperscript"));
		Shortcut(Keys.Control | Keys.Oemplus, "Subscript", api.ToggleSubscript, M("ToggleSubscript"));
		Shortcut(Keys.Control | Keys.OemCloseBrackets, "Increase font size", () => api.IncreaseFontSize(), M("IncreaseFontSize"));
		Shortcut(Keys.Control | Keys.OemOpenBrackets, "Decrease font size", () => api.DecreaseFontSize(), M("DecreaseFontSize"));
		Button(RibbonCommand.FontDialog, FontDialogCommand);
		foreach (var (id, mode) in new[] { (RibbonCommand.CaseUpper, "upper"), (RibbonCommand.CaseLower, "lower"), (RibbonCommand.CaseTitle, "title"), (RibbonCommand.CaseHeadline, "headline") })
			Button(id, () => api.ChangeCase(mode), Keys.None, M("ChangeCase", mode));
		foreach (var (id, style) in new[] { (RibbonCommand.UnderlineSingle, UnderlineStyle.Single), (RibbonCommand.UnderlineDouble, UnderlineStyle.Double),
			(RibbonCommand.UnderlineDotted, UnderlineStyle.Dotted), (RibbonCommand.UnderlineDashed, UnderlineStyle.Dashed), (RibbonCommand.UnderlineWavy, UnderlineStyle.Wavy) })
			Button(id, () => api.ToggleUnderline(style), Keys.None, M("ToggleUnderline", style));

		// Home: paragraph
		Toggle(RibbonCommand.Bullets, () => api.ToggleList(ListStyle.Bullet), Keys.Control | Keys.Shift | Keys.L, M("ToggleList", ListStyle.Bullet));
		Toggle(RibbonCommand.Numbering, () => api.NumberedList(), Keys.Control | Keys.Shift | Keys.N, M("NumberedList"));
		Toggle(RibbonCommand.Legal, () => api.LegalNumbering(), Keys.Control | Keys.Shift | Keys.G, M("LegalNumbering"));
		Button(RibbonCommand.Outdent, () => { if (api.InList()) { api.Nest(-1); Record(M("Nest", -1)); } else { Editor.Outdent(); Record(M("Outdent")); } });
		Button(RibbonCommand.Indent, () => { if (api.InList()) { api.Nest(1); Record(M("Nest", 1)); } else { Editor.Indent(); Record(M("Indent")); } });
		foreach (var (id, alignment, key) in new[] { (RibbonCommand.AlignLeft, ParagraphAlignment.Left, Keys.L), (RibbonCommand.AlignCenter, ParagraphAlignment.Center, Keys.E),
			(RibbonCommand.AlignRight, ParagraphAlignment.Right, Keys.R), (RibbonCommand.Justify, ParagraphAlignment.Justify, Keys.J) })
			Toggle(id, () => Editor.SetAlignment(alignment), Keys.Control | key, M("SetAlignment", alignment));
		foreach (var (id, spacing) in new[] { (RibbonCommand.LineSpacing10, "1"), (RibbonCommand.LineSpacing115, "1.15"), (RibbonCommand.LineSpacing150, "1.5"),
			(RibbonCommand.LineSpacing200, "2"), (RibbonCommand.LineSpacing250, "2.5"), (RibbonCommand.LineSpacing300, "3") })
			Toggle(id, () => api.SetLineSpacing(spacing), Keys.None, M("SetLineSpacing", spacing));
		Button(RibbonCommand.LineSpacingOptions, ParagraphCommand);
		Button(RibbonCommand.ParagraphDialog, ParagraphCommand);
		// (Word's Add / Remove Space Before / After Paragraph: 12 points, or none)
		Toggle(RibbonCommand.SpaceBefore, () => { float points = api.SpaceBefore > 0 ? 0 : 12; api.SpaceBefore = points; Record(M("SpaceBefore", points)); });
		Toggle(RibbonCommand.SpaceAfter, () => { float points = api.SpaceAfter > 0 ? 0 : 12; api.SpaceAfter = points; Record(M("SpaceAfter", points)); });
		foreach (var (id, style) in new[] { (RibbonCommand.ListBullet, ListStyle.Bullet), (RibbonCommand.ListDecimal, ListStyle.Decimal), (RibbonCommand.ListLowerAlpha, ListStyle.LowerAlpha),
			(RibbonCommand.ListUpperAlpha, ListStyle.UpperAlpha), (RibbonCommand.ListLowerRoman, ListStyle.LowerRoman), (RibbonCommand.ListUpperRoman, ListStyle.UpperRoman) })
			Button(id, () => api.ToggleList(style), Keys.None, M("ToggleList", style));
		Button(RibbonCommand.ListNone, () => api.RemoveList(), Keys.None, M("RemoveList"));
		Button(RibbonCommand.ListLevels, ListLevelsCommand);
		// (iteration 38: automatic numbering style per level is List levels' checkbox, not a menu item one click from it)
		Button(RibbonCommand.RestartNumbering, RestartNumberingCommand);
		Button(RibbonCommand.ContinueNumbering, () => api.ContinueNumbering(), Keys.None, M("ContinueNumbering"));
		Button(RibbonCommand.ConvertTyped, () => {
			int count = api.ConvertTypedNumbers();
			SetStatus(count == 0 ? T("status.noTypedNumbers", "No typed list numbers found (paragraphs that start with \"1.\", \"a)\", \"(i)\", \"•\" ... and a tab or space, two or more in a row).")
				: TN("status.becameListItems", "{0} paragraph became list items.", "{0} paragraphs became list items.", count));
		}, Keys.None, M("ConvertTypedNumbers"));
		Button(RibbonCommand.IndentParagraph, () => Editor.Indent(), Keys.Control | Keys.M, M("Indent"));
		Button(RibbonCommand.OutdentParagraph, () => Editor.Outdent(), Keys.Control | Keys.Shift | Keys.M, M("Outdent"));

		// Home: block styles (structured documents)
		blockCombo = new RibbonInRibbonGallery(ribbon, (uint)RibbonCommand.BlockStyle);    // (a combo box until iteration 72)
		blockCombo.ItemsSourceReady += (s, e) => FillStyleBox(force: true);
		blockCombo.ExecuteEvent += (s, e) => {
			if (updatingUi || GalleryItemEventArgs.Create(blockCombo, e).SelectedItem is not { SelectedItemIndex: >= 0 } item || item.SelectedItemIndex >= styleBoxNames.Count) return;
			string name = styleBoxNames[item.SelectedItemIndex];
			BeginInvoke(() => {
				ribbonSent.Clear();    // (Command)
				LeaveOverlays();
				if (SourceActive) {
					if (SourceBlockOf(name) is BlockStyle block) RunSource(() => SourceBlockStyle(block));
					else SetStatus(T("status.styleRichTextOnly", "The {0} style works in the Rich Text view (View > Rich Text); Markdown and HTML have headings, normal text and code.", name));
					return;
				}
				Run(() => api.SetStyle(name), M("SetStyle", name));
				Editor.Focus();
			});
		};
		// Ctrl+0 Normal, Ctrl+1 to Ctrl+6 the headings, Ctrl+7 Code
		for (int key = 0; key <= 7; key++) {
			string name = key == 0 ? "Normal" : key == 7 ? "Code" : "Heading " + key;
			Shortcut(Keys.Control | (Keys.D0 + key), name, () => api.SetStyle(name), M("SetStyle", name));
		}
		// Home > Styles (its label: the document's style set)
		stylesMenu = new RibbonDropDownButton(ribbon, (uint)RibbonCommand.StylesMenu);
		Button(RibbonCommand.ModifyStyle, ModifyStyleCommand);
		Button(RibbonCommand.NewStyle, NewStyleCommand);
		Button(RibbonCommand.UpdateStyle, UpdateStyleCommand);
		Button(RibbonCommand.DeleteStyle, DeleteStyleCommand);
		Button(RibbonCommand.ClearFormatting, () => api.ClearFormatting(), Keys.Control | Keys.Space, M("ClearFormatting"));
		Toggle(RibbonCommand.FormatPainter, FormatPainterCommand);
		foreach (var (id, set) in new[] { (RibbonCommand.StyleSetClassic, "Classic"), (RibbonCommand.StyleSetPlain, "Plain"), (RibbonCommand.StyleSetModern, "Modern") })
			Toggle(id, () => { api.StyleSet = set; Record(M("StyleSet", set)); });
		Button(RibbonCommand.ApplyMyStyles, () => { int count = api.ApplyMyStyles(); Record(M("ApplyMyStyles")); SetStatus(TN("status.myStylesApplied", "Your styles are this document's now ({0} paragraph restyled; formatting of their own stayed).", "Your styles are this document's now ({0} paragraphs restyled; formatting of their own stayed).", count)); });
		Button(RibbonCommand.CopyStyles, CopyStylesCommand);
		Button(RibbonCommand.SaveStyles, () => { api.SaveStylesAsDefault(); SetStatus(T("status.stylesSaved", "New documents get these styles now (Styles > Classic, Plain or Modern, then this, to change them again).")); });

		// Insert
		Button(RibbonCommand.PageBreak, () => api.InsertPageBreak(), Keys.None, M("InsertPageBreak"));    // Ctrl+Enter: ProcessCmdKey
		Button(RibbonCommand.SectionBreak, () => api.InsertSectionBreak(), Keys.None, M("InsertSectionBreak"));    // (iteration 63)
		Button(RibbonCommand.ContinuousBreak, () => api.InsertContinuousBreak(), Keys.None, M("InsertContinuousBreak"));    // (iteration 67)
		Button(RibbonCommand.Table, InsertTableDialog);
		Button(RibbonCommand.Picture, InsertImageDialog);
		Button(RibbonCommand.Thumbnail, InsertThumbnailDialog);
		Button(RibbonCommand.CrossReference, CrossReferenceCommand);
		Button(RibbonCommand.TableOfContents, TableOfContentsCommand);
		Button(RibbonCommand.Snippets, SnippetsCommand);
		Button(RibbonCommand.HorizontalLine, () => Run(() => api.InsertHorizontalLine(), M("InsertHorizontalLine")));
		Button(RibbonCommand.Footnote, FootnoteCommand, Keys.Control | Keys.Alt | Keys.F);
		Button(RibbonCommand.UpdateReferences, UpdateReferencesCommand, Keys.F9, M("UpdateCrossReferences"));
		Button(RibbonCommand.Link, InsertHyperlinkDialog, Keys.Control | Keys.K);
		Button(RibbonCommand.DateTime, () => { string now = DateTime.Now.ToString("f"); api.InsertText(now); Record(M("InsertText", now)); });

		// Layout, Review
		Button(RibbonCommand.HeaderFooter, HeaderFooterCommand);
		Button(RibbonCommand.Columns, ColumnsCommand);    // (iteration 67)
		foreach (var (id, mode) in new[] { (RibbonCommand.LineNumbersNone, LineNumbering.Modes.None), (RibbonCommand.LineNumbersContinuous, LineNumbering.Modes.Continuous),
			(RibbonCommand.LineNumbersEachPage, LineNumbering.Modes.EachPage) })
			Toggle(id, () => api.LineNumbers = LineNumbering.ModeName(mode), Keys.None, M("LineNumbers", LineNumbering.ModeName(mode)));
		Toggle(RibbonCommand.SuppressLineNumbers, () => { bool suppress = !api.LineNumbersSuppressed(); api.SuppressLineNumbers(suppress); Record(M("SuppressLineNumbers", suppress)); });
		Button(RibbonCommand.LineNumberOptions, LineNumbersCommand);
		Button(RibbonCommand.NextSpelling, NextSpellingErrorCommand, Keys.F7);
		Button(RibbonCommand.AddToDictionary, () => {
			string word = Editor.SelectedText.Trim();
			if (word == "" || word.Any(char.IsWhiteSpace)) throw new InvalidOperationException("Select a single word first.");
			api.AddToDictionary(word);
			SetStatus(T("status.addedToDictionary", "\"{0}\" was added to the Windows dictionary.", word));
		});
		Check(RibbonCommand.SpellCheck, () => Editor.EnableSpellCheck = !Editor.EnableSpellCheck);

		// View
		Check(RibbonCommand.WordWrap, () => Editor.WordWrap = !Editor.WordWrap);
		Check(RibbonCommand.PrintLayout, () => api.PrintLayout = !Settings.PrintLayout);
		Check(RibbonCommand.ShowPages, () => api.PageView = !Settings.PageView);    // (a check box: the ribbon showed the toggle button's state wrongly)
		Check(RibbonCommand.ShowBands, () => api.HeaderFooterVisible = !Settings.ShowBands);
		Check(RibbonCommand.Outline, () => api.OutlineVisible = outlineSplit.Panel1Collapsed, Keys.Control | Keys.Shift | Keys.O);
		Check(RibbonCommand.Ruler, () => api.RulerVisible = !Settings.ShowRuler);
		Button(RibbonCommand.ZoomIn, () => { if (PagesShown) pageView!.Zoom(1.1f); else ManualZoom(Math.Min(5f, EditorZoom() + 0.1f)); });
		Button(RibbonCommand.ZoomOut, () => { if (PagesShown) pageView!.Zoom(1 / 1.1f); else ManualZoom(Math.Max(0.5f, EditorZoom() - 0.1f)); });
		Button(RibbonCommand.Zoom100, () => { if (PagesShown) pageView!.Zoom(0); else ManualZoom(1f); });
		Toggle(RibbonCommand.ZoomPageWidth, () => api.ZoomMode = "PageWidth", Keys.None, M("ZoomMode", "PageWidth"));
		Toggle(RibbonCommand.ZoomOnePage, () => api.ZoomMode = "OnePage", Keys.None, M("ZoomMode", "OnePage"));
		themeGallery = new RibbonDropDownGallery(ribbon, (uint)RibbonCommand.Theme);
		themeGallery.ItemsSourceReady += (s, e) => FillThemeGallery();
		themeGallery.ExecuteEvent += (s, e) => {
			if (GalleryItemEventArgs.Create(themeGallery, e).SelectedItem is { } item && item.SelectedItemIndex >= 0 && item.SelectedItemIndex < AppTheme.All.Length) {
				string name = AppTheme.All[item.SelectedItemIndex].Name;
				BeginInvoke(() => Command(RibbonCommand.Theme, () => api.Theme = name, M("Theme", name)));
			}
		};
		languageGallery = new RibbonDropDownGallery(ribbon, (uint)RibbonCommand.Language);    // (iteration 80)
		languageGallery.ItemsSourceReady += (s, e) => FillLanguageGallery();
		languageGallery.ExecuteEvent += (s, e) => {
			if (GalleryItemEventArgs.Create(languageGallery, e).SelectedItem is { } item && item.SelectedItemIndex >= 0 && item.SelectedItemIndex < languages.Count) {
				string code = languages[(int)item.SelectedItemIndex].Code;
				BeginInvoke(() => Command(RibbonCommand.Language, () => ChooseLanguage(code), M("Language", code)));
			}
		};
		Toggle(RibbonCommand.ZoomMultiplePages, () => { bool show = !PagesShown; api.MultiplePages = show; Record(M("MultiplePages", show)); });
		Toggle(RibbonCommand.ViewRich, () => SetView(SourceKind.None, false));
		Toggle(RibbonCommand.ViewMarkdown, () => SetView(SourceKind.Markdown, sideBySide));
		Toggle(RibbonCommand.ViewHtml, () => SetView(SourceKind.Html, sideBySide));
		Toggle(RibbonCommand.SideBySide, () => SetView(view == SourceKind.None ? SourceKind.Markdown : view, !sideBySide));
		Check(RibbonCommand.SourceWordWrap, () => sourceBox.WordWrap = !sourceBox.WordWrap);
		Check(RibbonCommand.SourceLineNumbers, () => sourceBox.LineNumbers = !sourceBox.LineNumbers);

		// Tools
		macrosGallery = new RibbonDropDownGallery(ribbon, (uint)RibbonCommand.Macros);
		macrosGallery.ItemsSourceReady += (s, e) => FillMacrosGallery(true);
		macrosGallery.ExecuteEvent += (s, e) => {
			if (GalleryItemEventArgs.Create(macrosGallery, e).SelectedItem is { } item && item.SelectedItemIndex >= 0 && item.SelectedItemIndex < galleryMacros.Count) {
				string file = galleryMacros[item.SelectedItemIndex];
				BeginInvoke(() => { LeaveOverlays(); Run(() => RunMacroFromUi(file)); });
			}
		};
		Toggle(RibbonCommand.RecordMacro, ToggleRecording, Keys.Control | Keys.Shift | Keys.R);
		Button(RibbonCommand.MacroHotkeys, MacroHotkeysCommand);
		Button(RibbonCommand.EventHooks, EventHooksCommand);
		Button(RibbonCommand.CommandPalette, CommandPaletteCommand, Keys.Control | Keys.Shift | Keys.P);
		Button(RibbonCommand.RunScript, RunScriptDialog);
		Button(RibbonCommand.MacrosFolder, OpenMacrosFolder);
		Button(RibbonCommand.AutomationApi, () => ShowText("Automation API (RteClient2.ahk: Rte.<Name>(...))", api.Help()));
		Button(RibbonCommand.About, ShowAbout);
		extensionsGallery = new RibbonDropDownGallery(ribbon, (uint)RibbonCommand.Extensions);
		extensionsGallery.ItemsSourceReady += (s, e) => FillExtensionsGallery();
		extensionsGallery.ExecuteEvent += (s, e) => {
			if (GalleryItemEventArgs.Create(extensionsGallery, e).SelectedItem is { } item && item.SelectedItemIndex >= 0 && item.SelectedItemIndex < extensionCommands.Count) {
				var command = extensionCommands[item.SelectedItemIndex];
				BeginInvoke(() => { LeaveOverlays(); Run(command.Run); });
			}
		};

		BuildSourceCommands();

		// Table Tools (shown while the caret is in a table)
		tableTools = new RibbonTabGroup(ribbon, (uint)RibbonCommand.TableTools);
		Button(RibbonCommand.InsertAbove, () => api.InsertRow(false), Keys.None, M("InsertRow", false));
		Button(RibbonCommand.InsertBelow, () => api.InsertRow(true), Keys.None, M("InsertRow", true));    // Ctrl+Shift+Insert: ProcessCmdKey
		Button(RibbonCommand.InsertLeft, () => api.InsertColumn(false), Keys.None, M("InsertColumn", false));
		Button(RibbonCommand.InsertRight, () => api.InsertColumn(true), Keys.None, M("InsertColumn", true));
		Button(RibbonCommand.DeleteRow, () => api.DeleteRow(), Keys.None, M("DeleteRow"));    // Ctrl+Shift+Delete: ProcessCmdKey
		Button(RibbonCommand.DeleteColumn, () => api.DeleteColumn(), Keys.None, M("DeleteColumn"));
		Button(RibbonCommand.DeleteTable, () => api.DeleteTable(), Keys.None, M("DeleteTable"));
		Button(RibbonCommand.DeleteCells, DeleteCellsCommand);    // (iteration 92)
		Button(RibbonCommand.MergeCells, () => api.MergeCells(), Keys.None, M("MergeCells"));
		Button(RibbonCommand.SplitCells, SplitCellsCommand);
		Button(RibbonCommand.UpdateTable, UpdateTableCommand);    // (iteration 94)
		foreach (var (id, part) in new[] { (RibbonCommand.SelectCell, "cell"), (RibbonCommand.SelectColumn, "column"), (RibbonCommand.SelectRow, "row"), (RibbonCommand.SelectTable, "table") })
			Button(id, () => api.SelectTable(part), Keys.None, M("SelectTable", part));
		Button(RibbonCommand.SaveTableToSource, SaveTableToSourceCommand);
		Button(RibbonCommand.LinkTable, LinkTableCommand);
		Button(RibbonCommand.UnlinkTable, () => { api.UnlinkTable(); SetStatus(T("status.table.unlinked", "The table is no longer linked to a file")); }, Keys.None, M("UnlinkTable"));
		Button(RibbonCommand.TableProperties, TablePropertiesCommand);
		foreach (var (id, preset) in new[] { (RibbonCommand.BordersAll, "all"), (RibbonCommand.BordersOutside, "outside"), (RibbonCommand.BordersInside, "inside"), (RibbonCommand.BordersNone, "none"),
			(RibbonCommand.BordersTop, "top"), (RibbonCommand.BordersBottom, "bottom"), (RibbonCommand.BordersLeft, "left"), (RibbonCommand.BordersRight, "right") })
			Button(id, () => api.SetCellBorders(preset), Keys.None, M("SetCellBorders", preset));
		shadingPicker = new RibbonDropDownColorPicker(ribbon, (uint)RibbonCommand.TableShading);
		shadingPicker.ExecuteEvent += (s, e) => {
			var picked = ColorPickerEventArgs.Create(shadingPicker, e);
			string color = picked.ColorType == SwatchColorType.RGB && picked.RGBColor is Color c ? $"#{c.R:X2}{c.G:X2}{c.B:X2}" : "none";
			BeginInvoke(() => Command(RibbonCommand.TableShading, () => api.SetCellShading(color), M("SetCellShading", color)));
		};

		// (AhkWord-es: the ribbon is compiled without its strings, which every command is given here, in the language of
		// the setting Language (or AHKWORD_LANG); rte_strings.cs. Loaded at the start already: this only when it wasn't)
		Lang.Load(appDir, Settings.Language);
		RibbonStrings.Give(ribbon, Path.Combine(appDir, "ribbon"));
	}

	// Ribbon commands run after the ribbon's callback has returned (BeginInvoke): dialogs opened from inside it would run
	// while the ribbon is still handling the click.
	RibbonButton Button(RibbonCommand id, Action action, Keys keys = Keys.None, string? macro = null) {
		var button = new RibbonButton(ribbon, (uint)id);
		button.ExecuteEvent += (s, e) => BeginInvoke(() => Command(id, action, macro));
		if (keys != Keys.None) { Shortcut(keys, CommandName(id), action, macro); keyCommands[keys] = id; }
		paletteCommands[id] = (() => Command(id, action, macro), () => button.Enabled, null);
		return button;
	}

	void Toggle(RibbonCommand id, Action action, Keys keys = Keys.None, string? macro = null) {
		var button = new RibbonToggleButton(ribbon, (uint)id);
		button.ExecuteEvent += (s, e) => BeginInvoke(() => Command(id, action, macro));
		if (keys != Keys.None) { Shortcut(keys, CommandName(id), action, macro); keyCommands[keys] = id; }
		toggleButtons[id] = button;
		paletteCommands[id] = (() => Command(id, action, macro), () => button.Enabled, () => button.BooleanValue);
	}

	void Check(RibbonCommand id, Action action, Keys keys = Keys.None) {
		var box = new RibbonCheckBox(ribbon, (uint)id);
		box.ExecuteEvent += (s, e) => BeginInvoke(() => Command(id, action, null));
		if (keys != Keys.None) { Shortcut(keys, CommandName(id), action); keyCommands[keys] = id; }
		checkBoxes[id] = box;
		paletteCommands[id] = (() => Command(id, action, null), () => box.Enabled, () => box.BooleanValue);
	}

	// ---- Themes (iteration 70, rte_theme.cs)

	// ---- Format Painter (iteration 75): Home > Clipboard, as in Word. A click copies the formatting at the caret
	// (EditorApi.CopyFormatting: the caret's paragraph's, and its characters') and the next text selected with the mouse
	// takes it (PasteFormatting; a click without dragging: that paragraph, and the word there); two clicks within the
	// double-click time keep it painting, until Esc or a click on it.
	enum Painting { Off, Once, Sticky }
	Painting painting;
	long painterClicked;

	void FormatPainterCommand() {
		long now = Environment.TickCount64;
		bool twice = now - painterClicked <= SystemInformation.DoubleClickTime;
		painterClicked = now;
		if (painting == Painting.Once && twice) {
			painting = Painting.Sticky;
			SetStatus(T("status.painterStays", "Format Painter stays on: select text to paint it; Esc or a click on Format Painter stops."));
		} else if (painting != Painting.Off) {
			StopPainting();
			SetStatus(T("status.painterOff", "Format Painter off."));
		} else {
			api.CopyFormatting();
			Record(M("CopyFormatting"));
			painting = Painting.Once;
			Editor.PaintCursor = PaintCursor.Value;
			SetStatus(T("status.painterSelect", "Click in a paragraph to give it this formatting (and the word there), or select text. Click Format Painter twice to keep painting; Esc stops."));
		}
		toggleButtons[RibbonCommand.FormatPainter].BooleanValue = painting != Painting.Off;    // (the ribbon flips it on each click)
		Editor.Focus();
	}

	void PaintSelection() {
		if (painting == Painting.Once) StopPainting();
		Run(() => api.PasteFormatting(), M("PasteFormatting"));
	}

	void StopPainting() {
		painting = Painting.Off;
		Editor.PaintCursor = null;
		toggleButtons[RibbonCommand.FormatPainter].BooleanValue = false;
	}

	// The pointer while painting: an I-beam with a brush beside it (its hot spot the I-beam's middle)
	static readonly Lazy<Cursor> PaintCursor = new(() => {
		int size = Math.Max(32, SystemInformation.CursorSize.Width);
		float k = size / 32f;
		using var bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
		using (var g = Graphics.FromImage(bitmap)) {
			g.SmoothingMode = SmoothingMode.AntiAlias;
			g.ScaleTransform(k, k);
			PointF[][] beam = { new PointF[] { new(2, 7), new(8, 7) }, new PointF[] { new(5, 7), new(5, 27) }, new PointF[] { new(2, 27), new(8, 27) } };
			using (var halo = new Pen(Color.White, 3f)) foreach (var line in beam) g.DrawLines(halo, line);
			using (var ink = new Pen(Color.Black, 1f)) foreach (var line in beam) g.DrawLines(ink, line);
			using var outline = new Pen(Color.Black, 1f);
			void Part(Color color, RectangleF box) { using var fill = new SolidBrush(color); g.FillRectangle(fill, box); g.DrawRectangle(outline, box.X, box.Y, box.Width, box.Height); }
			Part(Color.FromArgb(0x8A, 0x5A, 0x2B), new RectangleF(19, 1, 4, 10));    // handle
			Part(Color.Silver, new RectangleF(16, 11, 10, 4));    // ferrule
			Part(Color.FromArgb(0xE8, 0xA3, 0x17), new RectangleF(16, 15, 10, 7));    // bristles
		}
		return WriterEditBox.CursorFrom(bitmap, (int)(5 * k), (int)(17 * k));
	});

	string ribbonLoaded = "";    // (the file the ribbon came from)
	RibbonDropDownGallery themeGallery = null!;
	RibbonDropDownGallery languageGallery = null!;

	// The dark ribbon's Font group (ribbon\RteRibbonDarkFont.xml). Windows' font control doesn't follow a dark ribbon: its
	// B, I, U ... stay dark navy, and its font list near-black, whatever the colors. These do what it does, by the same way
	// (ApplyFontChanges), so the views and the macros behave alike.
	bool DarkRibbon => ribbonLoaded.EndsWith("RteRibbonDark.ribbon", StringComparison.OrdinalIgnoreCase);
	RibbonComboBox darkFontName = null!, darkFontSize = null!;
	readonly Dictionary<RibbonCommand, RibbonToggleButton> darkFontToggles = new();
	static readonly string[] FontSizes = { "8", "9", "10", "10.5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72" };
	List<string> darkFontNames = new();

	internal void ChangeFont(Dictionary<FontPropertiesEnum, object> changes) => BeginInvoke(() => { ribbonSent.Clear(); LeaveOverlays(); ApplyFontChanges(changes); Editor.Focus(); });

	void BuildDarkFont() {
		if (!DarkRibbon) return;
		RibbonComboBox Box(RibbonCommand id, string widest, Action<string> chosen, Func<IEnumerable<string>> list) {
			var box = new RibbonComboBox(ribbon, (uint)id);
			box.RepresentativeString = widest;
			box.ItemsSourceReady += (s, e) => {
				if (box.ItemsSource is not { } items) return;
				items.Clear();
				foreach (var item in list()) items.Add(new GalleryItemPropertySet { Label = item, CategoryID = Constants.UI_Collection_InvalidIndex });
			};
			box.ExecuteEvent += (s, e) => {
				if (updatingUi) return;
				string text = GalleryItemEventArgs.Create(box, e).SelectedItem is { SelectedItemIndex: >= 0 } item && item.SelectedItemIndex < list().Count() ? list().ElementAt(item.SelectedItemIndex) : box.StringValue;
				if (!string.IsNullOrWhiteSpace(text)) chosen(text.Trim());
			};
			return box;
		}
		darkFontName = Box(RibbonCommand.FontName, "Times New Roman W", name => ChangeFont(new() { [FontPropertiesEnum.Family] = name }), () => {
			if (darkFontNames.Count == 0) {
				using var fonts = new System.Drawing.Text.InstalledFontCollection();
				darkFontNames = fonts.Families.Select(f => f.Name).ToList();
			}
			return darkFontNames;
		});
		darkFontSize = Box(RibbonCommand.FontSize, "10.5", text => {
			if (decimal.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var size) && size > 0) ChangeFont(new() { [FontPropertiesEnum.Size] = size });
			else ribbonSent.Clear();    // (the box shows the size again)
		}, () => FontSizes);
		void Toggles(RibbonCommand id, Func<Dictionary<FontPropertiesEnum, object>> changes) {
			var button = new RibbonToggleButton(ribbon, (uint)id);
			button.ExecuteEvent += (s, e) => ChangeFont(changes());
			darkFontToggles[id] = button;
		}
		Toggles(RibbonCommand.FontBold, () => new() { [FontPropertiesEnum.Bold] = true });
		Toggles(RibbonCommand.FontItalic, () => new() { [FontPropertiesEnum.Italic] = true });
		Toggles(RibbonCommand.FontUnderline, () => new() { [FontPropertiesEnum.Underline] = true });
		Toggles(RibbonCommand.FontStrike, () => new() { [FontPropertiesEnum.Strikethrough] = true });
		Toggles(RibbonCommand.FontSub, () => new() { [FontPropertiesEnum.VerticalPositioning] = api.SelectionFont().Subscript == tomConstants.tomTrue ? FontVerticalPosition.NotSet : FontVerticalPosition.SubScript });
		Toggles(RibbonCommand.FontSuper, () => new() { [FontPropertiesEnum.VerticalPositioning] = api.SelectionFont().Superscript == tomConstants.tomTrue ? FontVerticalPosition.NotSet : FontVerticalPosition.SuperScript });
		new RibbonButton(ribbon, (uint)RibbonCommand.FontGrow).ExecuteEvent += (s, e) => ChangeFont(new() { [FontPropertiesEnum.DeltaSize] = FontDeltaSize.Grow });
		new RibbonButton(ribbon, (uint)RibbonCommand.FontShrink).ExecuteEvent += (s, e) => ChangeFont(new() { [FontPropertiesEnum.DeltaSize] = FontDeltaSize.Shrink });
		void Picker(RibbonCommand id, FontPropertiesEnum kind, FontPropertiesEnum color) {
			var picker = new RibbonDropDownColorPicker(ribbon, (uint)id);
			picker.ExecuteEvent += (s, e) => {
				var picked = ColorPickerEventArgs.Create(picker, e);
				var changes = new Dictionary<FontPropertiesEnum, object> { [kind] = picked.ColorType };
				if (picked.RGBColor is Color c) changes[color] = c;
				ChangeFont(changes);
			};
		}
		Picker(RibbonCommand.FontColor, FontPropertiesEnum.ForegroundColorType, FontPropertiesEnum.ForegroundColor);
		Picker(RibbonCommand.FontHighlight, FontPropertiesEnum.BackgroundColorType, FontPropertiesEnum.BackgroundColor);
	}

	void ShowDarkFont(ITextFont font) {
		Push(darkFontName, "StringValue", font.Name ?? "", v => darkFontName.StringValue = v);
		Push(darkFontSize, "StringValue", font.Size > 0 ? Math.Round(font.Size, 1).ToString(System.Globalization.CultureInfo.InvariantCulture) : "", v => darkFontSize.StringValue = v);
		void On(RibbonCommand id, bool on) => Push(darkFontToggles[id], "BooleanValue", on, v => darkFontToggles[id].BooleanValue = v);
		On(RibbonCommand.FontBold, font.Bold == tomConstants.tomTrue);
		On(RibbonCommand.FontItalic, font.Italic == tomConstants.tomTrue);
		On(RibbonCommand.FontUnderline, font.Underline > 0);
		On(RibbonCommand.FontStrike, font.StrikeThrough == tomConstants.tomTrue);
		On(RibbonCommand.FontSub, font.Subscript == tomConstants.tomTrue);
		On(RibbonCommand.FontSuper, font.Superscript == tomConstants.tomTrue);
	}

	void FillThemeGallery() {
		if (!ribbon.Initialized || themeGallery.ItemsSource is not { } items) return;
		items.Clear();
		foreach (var theme in AppTheme.All)
			using (var swatch = theme.Swatch(LogicalToDeviceUnits(16)))
				items.Add(new GalleryItemPropertySet { Label = theme.Label, CategoryID = Constants.UI_Collection_InvalidIndex, ItemImage = ribbon.ConvertToUIImage(swatch) });
		ShowThemeChosen();
	}

	void ShowThemeChosen() => themeGallery.SelectedItem = (uint)Array.IndexOf(AppTheme.All, AppTheme.Named(Settings.Theme));

	// Review > Language (iteration 80): the languages there are files for, each by its own name; the one chosen in the
	// settings ticked. A choice is saved for the next start: the interface's strings are made as the window is.
	List<(string Code, string Name)> languages = new();
	void FillLanguageGallery() {
		if (!ribbon.Initialized || languageGallery.ItemsSource is not { } items) return;
		languages = Lang.Available(appDir);
		items.Clear();
		foreach (var (_, name) in languages)
			items.Add(new GalleryItemPropertySet { Label = name, CategoryID = Constants.UI_Collection_InvalidIndex });
		ShowLanguageChosen();
	}

	void ShowLanguageChosen() {
		if (!ribbon.Initialized) return;
		int index = languages.FindIndex(l => l.Code.Equals(Lang.Chosen(appDir, Settings.Language), StringComparison.OrdinalIgnoreCase));
		languageGallery.SelectedItem = (uint)Math.Max(0, index);
	}

	/// Review > Language and the command palette: the setting Language, and when AhkWord isn't in that language now, when
	/// it will be.
	internal void ChooseLanguage(string code) {
		api.Language = code;
		ShowLanguageChosen();
		string chosen = Lang.Chosen(appDir, Settings.Language);
		string name = Lang.NameOf(chosen);
		if (Environment.GetEnvironmentVariable("AHKWORD_LANG") is { Length: > 0 } env)
			MessageBox.Show(this, T("msg.languageEnvironment", "AhkWord will be in {0} when it starts without the environment variable AHKWORD_LANG, which sets its language to \"{1}\" now.", name, env),
				"AhkWord", MessageBoxButtons.OK, MessageBoxIcon.Information);
		else if (!chosen.Equals(Lang.Code, StringComparison.OrdinalIgnoreCase))
			MessageBox.Show(this, T("msg.languageNextStart", "AhkWord will be in {0} the next time it starts.", name), "AhkWord", MessageBoxButtons.OK, MessageBoxIcon.Information);
		else SetStatus(T("status.languageKept", "AhkWord is in {0}.", name));
	}

	// The theme's colors, on the ribbon as soon as it is there.
	void PaintTheme() {
		var theme = AppTheme.Named(Settings.Theme);
		theme.Paint(ribbon);
		theme.PaintTitleBar(Handle);
		PaintFrame(theme);
	}

	// A darker theme's colors on what is around the page. The dialogs stay the neutral dark of the system colors.
	void PaintFrame(AppTheme theme) {
		if (!AppTheme.DarkWindows) return;
		statusBar.BackColor = theme.Background;
		outlineSplit.BackColor = sourceSplit.BackColor = sourceTitle.BackColor = theme.Panel;
		outlineSplit.Panel1.BackColor = outlineTree.BackColor = outlineHint.BackColor = theme.Panel;
		PrintLayoutPage.Surround = lineGutter.BandColor = theme.Surround;
		if (pageHost.BackColor != AppTheme.Paper) pageHost.BackColor = theme.Surround;
		pageHost.Invalidate(true);
		Editor.Invalidate();
	}

	// A change of theme while running. The darker themes share a ribbon and change at once. Light's is another file, which loads
	// only at the start: reloading the ribbon in a running window lost the font box, the Style box and its label (seen in
	// iteration 70), so that change waits; the setting is saved for the next window.
	internal void ChangeTheme() {
		Settings.Save();
		if (!ribbon.Initialized) return;
		ShowThemeChosen();
		if (RibbonFile == ribbonLoaded) PaintTheme();
		else SetStatus(T("status.themeNextWindow", "The {0} theme starts with the next AhkWord window: File > New Window, or the next time AhkWord starts.", Settings.Theme));
	}

	// ---- The command palette (iteration 57, rte_palette.cs)

	// each ribbon command: run as its button runs it, whether it can run now, and (toggles, check boxes) whether it is on
	readonly Dictionary<RibbonCommand, (Action Run, Func<bool> Enabled, Func<bool>? On)> paletteCommands = new();
	Dictionary<uint, RibbonLabels.Label>? ribbonLabels;

	/// What the palette lists: the ribbon's commands (but the palette), the macros, the styles, the snippets, the
	/// recent files.
	internal List<PaletteItem> PaletteItems() {
		ribbonLabels ??= RibbonLabels.Load(Path.Combine(appDir, "ribbon", "RteRibbon.xml"));
		var keysOf = keyCommands.GroupBy(k => k.Value).ToDictionary(g => g.Key, g => ShortcutText(g.First().Key));
		var items = new List<PaletteItem>();
		foreach (var (id, command) in paletteCommands) {
			if (id == RibbonCommand.CommandPalette) continue;
			var label = ribbonLabels.TryGetValue((uint)id, out var l) ? l : null;
			string name = label?.Text is { Length: > 0 } text ? text : CommandName(id);
			if (command.On != null) {
				bool? on = null;
				try { on = command.On(); } catch (Exception) { }
				if (on is bool b) name += "  " + (b ? T("palette.on", "(on)") : T("palette.off", "(off)"));
			}
			items.Add(new PaletteItem(name, label?.Where ?? "", label?.Tooltip ?? "", keysOf.TryGetValue(id, out var k) ? k : "", command.Enabled, command.Run, "cmd:" + id));
		}
		items.Add(new PaletteItem(T("palette.pastePlain", "Paste as Plain Text"), T("palette.homeClipboard", "Home > Clipboard"), T("palette.pastePlainTip", "The clipboard's text only, taking the formatting where it lands (a web page's or a chat's fonts and spacing left behind)."),
			ShortcutText(Keys.Control | Keys.Shift | Keys.V), () => true, () => Run(pastePlainAction, M("PastePlainText")), "paste:plain"));
		// (the Font block is one ribbon control, not buttons: its commands, as their keys run them, in the source views too)
		foreach (var (fontName, keys, run, call) in new (string, Keys, Action, string)[] {
			("Bold", Keys.Control | Keys.B, api.ToggleBold, M("ToggleBold")), ("Italic", Keys.Control | Keys.I, api.ToggleItalic, M("ToggleItalic")),
			("Underline", Keys.Control | Keys.U, () => api.ToggleUnderline(), M("ToggleUnderline")), ("Strikethrough", Keys.None, api.ToggleStrikeThrough, M("ToggleStrikeThrough")),
			("Superscript", Keys.Control | Keys.Shift | Keys.Oemplus, api.ToggleSuperscript, M("ToggleSuperscript")), ("Subscript", Keys.Control | Keys.Oemplus, api.ToggleSubscript, M("ToggleSubscript")),
			("Increase Font Size", Keys.Control | Keys.OemCloseBrackets, () => api.IncreaseFontSize(), M("IncreaseFontSize")),
			("Decrease Font Size", Keys.Control | Keys.OemOpenBrackets, () => api.DecreaseFontSize(), M("DecreaseFontSize")) }) {
			var (key, action, macro) = (keys, run, call);
			items.Add(new PaletteItem(fontName switch {
					"Bold" => T("palette.bold", "Bold"), "Italic" => T("palette.italic", "Italic"), "Underline" => T("palette.underline", "Underline"),
					"Strikethrough" => T("palette.strikethrough", "Strikethrough"), "Superscript" => T("palette.superscript", "Superscript"), "Subscript" => T("palette.subscript", "Subscript"),
					"Increase Font Size" => T("palette.increaseFontSize", "Increase Font Size"), "Decrease Font Size" => T("palette.decreaseFontSize", "Decrease Font Size"), _ => fontName,
				}, T("palette.homeFont", "Home > Font"), "", key == Keys.None ? "" : ShortcutText(key), () => key != Keys.None || !SourceActive, () => {
				LeaveOverlays();
				if (SourceActive && key != Keys.None && sourceKeys.TryGetValue(key, out var inSource)) inSource();
				else Run(action, macro);
			}, "font:" + fontName));
		}
		foreach (var file in MacroPaths())
			items.Add(new PaletteItem(T("palette.runMacro", "Run macro: {0}", Path.GetFileNameWithoutExtension(file)), T("palette.toolsMacros", "Tools > Macros"), "", MacroHotkeyFor(file), () => true,
				() => { LeaveOverlays(); Run(() => RunMacroFromUi(file)); }, "macro:" + Path.GetFileName(file)));
		foreach (var style in Styles.Styles) {
			string styleName = style.Name;
			items.Add(new PaletteItem(T("palette.style", "Style: {0}", styleName), T("palette.homeStyles", "Home > Styles"), "", "", () => !SourceActive,
				() => Run(() => api.SetStyle(styleName), M("SetStyle", styleName)), "style:" + styleName));
		}
		foreach (var (code, languageName) in Lang.Available(appDir)) {
			string chosen = code;
			items.Add(new PaletteItem(T("palette.language", "Language: {0}", languageName), T("palette.reviewLanguage", "Review > Language"), "", "", () => true,
				() => Run(() => ChooseLanguage(chosen), M("Language", chosen)), "lang:" + code));
		}
		foreach (var snippet in SnippetList()) {
			string abbreviation = snippet.Abbreviation;
			items.Add(new PaletteItem(T("palette.snippet", "Snippet: {0}  →  {1}", abbreviation, snippet.Part), T("palette.insertSnippets", "Insert > Snippets"), "", "", () => true,
				() => Run(() => api.InsertSnippet(abbreviation, "", true), M("InsertSnippet", abbreviation, "", true)), "snippet:" + abbreviation));
		}
		foreach (var command in extensionCommands)
			items.Add(new PaletteItem(command.Name, T("palette.toolsExtensions", "Tools > Extensions"), command.Description, "", () => true, () => Run(command.Run), "ext:" + command.Name));
		foreach (var path in Settings.RecentFiles) {
			string file = path;
			items.Add(new PaletteItem(T("palette.open", "Open: {0}", Path.GetFileName(file)), Path.GetDirectoryName(file) ?? "", "", "", () => File.Exists(file),
				() => Run(() => OpenRecent(file)), "open:" + file));
		}
		return items;
	}

	// the macros in the macros folder (their files), and a macro's hotkey ("" if none)
	List<string> MacroPaths() => Directory.Exists(MacrosDir) ? Directory.GetFiles(MacrosDir, "*.ahk").OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList() : new();
	string MacroHotkeyFor(string file) => Settings.MacroHotkeys.FirstOrDefault(h => string.Equals(h.Key, Path.GetFileName(file), StringComparison.OrdinalIgnoreCase)).Value ?? "";

	static string ShortcutText(Keys keys) {
		var parts = new List<string>();
		if (keys.HasFlag(Keys.Control)) parts.Add("Ctrl");
		if (keys.HasFlag(Keys.Shift)) parts.Add("Shift");
		if (keys.HasFlag(Keys.Alt)) parts.Add("Alt");
		var key = keys & Keys.KeyCode;
		parts.Add(key switch {
			>= Keys.D0 and <= Keys.D9 => ((char)('0' + (key - Keys.D0))).ToString(),
			Keys.Oemplus => "+", Keys.OemMinus => "-", Keys.Oemcomma => ",", Keys.OemPeriod => ".",
			Keys.Next => "Page Down", Keys.Prior => "Page Up", Keys.Return => "Enter",
			_ => key.ToString()
		});
		return string.Join("+", parts);
	}

	PaletteForm? palette;

	// Ctrl+Shift+P: the palette over the window; what is chosen runs once it has closed (a dialog it opens belongs to
	// the window, not to the palette)
	void CommandPaletteCommand() {
		if (palette is { IsDisposed: false }) { palette.Activate(); return; }
		palette = new PaletteForm(PaletteItems(), Settings.PaletteRecent);
		palette.FormClosed += (s, e) => {
			var chosen = palette?.Chosen;
			palette?.Dispose();
			palette = null;
			if (chosen == null) return;
			Settings.PaletteRecent = Settings.PaletteRecent.Where(k => k != chosen.Key).Prepend(chosen.Key).Take(30).ToList();
			Settings.Save();
			BeginInvoke(() => { Activate(); Editor.Focus(); chosen.Run(); });
		};
		palette.ShowOver(this);
	}

	// A ribbon command. While the source editor is in use, formatting commands insert Markdown or HTML there (or say
	// that they need the rich text view) instead of changing the formatted text out of sight.
	bool zoomClicked = true;    // (Print Layout, Page Width, One Page or Multiple Pages clicked, or the ribbon new: UpdateRibbon pushes their states to it)

	void Command(RibbonCommand id, Action action, string? macro) {
		ribbonSent.Clear();    // (a click may have changed what the ribbon shows: UpdateRibbon sends all again)
		if (id is RibbonCommand.PrintLayout or RibbonCommand.ZoomPageWidth or RibbonCommand.ZoomOnePage or RibbonCommand.ZoomMultiplePages) zoomClicked = true;
		if (PagesShown && !PageViewCommands.Contains(id)) ShowPages(false);
		if (StartShown && id != RibbonCommand.StartPage) ShowStart(false);
		if (SourceActive && !SourceDocumentCommands.Contains(id)) {
			if (sourceCommands.TryGetValue(id, out var markup)) RunSource(markup);
			else { SetStatus(T("status.richTextOnly", "{0} works in the Rich Text view (View > Rich Text).", CommandName(id))); UpdateUi(); }
			return;
		}
		Run(action, macro);
	}

	void Shortcut(Keys keys, string name, Action action, string? macro = null) => shortcuts[keys] = (name, () => Run(action, macro));

	// An editing command that acts on the source editor when it is the one in use, else on the formatted text.
	Action OnText(Action rich, Action source) => () => { if (SourceActive) source(); else rich(); };

	// "SaveAs" -> "Save As": the command's name for the macro hotkey warnings.
	static string CommandName(RibbonCommand id) => Regex.Replace(id.ToString(), "(?<=[a-z])(?=[A-Z0-9])", " ");

	// macro: the call recorded while a macro is being recorded (made with M), or null for commands that are not recorded.
	static string M(string member, params object?[] args) => MacroRecorder.Call(member, args);

	static string BlockName(BlockStyle style) => style switch {
		BlockStyle.Paragraph => "Normal text",
		BlockStyle.Preformatted => "Preformatted",
		_ => "Heading " + (int)style
	};

	void ShowAbout() => MessageBox.Show(this, T("msg.about", "{0} (iteration {1})\n\n{2}\n\nA word processor built on RichEditorNET, hosted in AutoHotkey with Csharp-in-AHK. The ribbon is the Windows Ribbon, through WindowsRibbon.\n\nConverters for .docx and .odt:\n{3}",
		AppName, Iteration, api.Version, OfficeConverters.Describe()), AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);

	// Applies what the user changed in the font control (the ribbon lists only those). Colors come with a type:
	// "Automatic" text color and "No Color" highlight are not RGB values (the color itself then reads as black).
	void ApplyFontChanges(Dictionary<FontPropertiesEnum, object>? changes) {
		if (changes == null || changes.Count == 0) return;
		if (SourceActive) { RunSource(() => SourceFont(changes)); return; }
		// WindowsRibbon boxes some values as other enum types, so they are read by number.
		uint Number(FontPropertiesEnum key) => Convert.ToUInt32(changes[key]);
		bool Has(FontPropertiesEnum key) => changes.ContainsKey(key);
		var font = api.SelectionFont();
		if (Has(FontPropertiesEnum.Family) && changes[FontPropertiesEnum.Family] is string name && name.Trim() != "")
			Run(() => api.SetFontName(name.Trim()), M("SetFontName", name.Trim()));
		if (Has(FontPropertiesEnum.Size) && changes[FontPropertiesEnum.Size] is decimal size && size > 0) {
			Run(() => api.SetFontSize((float)size), M("SetFontSize", (float)size));
		} else if (Has(FontPropertiesEnum.DeltaSize)) {
			if ((FontDeltaSize)Number(FontPropertiesEnum.DeltaSize) == FontDeltaSize.Grow) Run(() => api.IncreaseFontSize(), M("IncreaseFontSize"));
			else Run(() => api.DecreaseFontSize(), M("DecreaseFontSize"));
		}
		if (Has(FontPropertiesEnum.Bold)) Run(api.ToggleBold, M("ToggleBold"));
		if (Has(FontPropertiesEnum.Italic)) Run(api.ToggleItalic, M("ToggleItalic"));
		if (Has(FontPropertiesEnum.Underline)) Run(() => api.ToggleUnderline(), M("ToggleUnderline"));
		if (Has(FontPropertiesEnum.Strikethrough)) Run(api.ToggleStrikeThrough, M("ToggleStrikeThrough"));
		if (Has(FontPropertiesEnum.VerticalPositioning)) {
			var position = (FontVerticalPosition)Number(FontPropertiesEnum.VerticalPositioning);
			bool super = font.Superscript == tomConstants.tomTrue, sub = font.Subscript == tomConstants.tomTrue;
			if (position == FontVerticalPosition.SuperScript != super) Run(api.ToggleSuperscript, M("ToggleSuperscript"));
			if (position == FontVerticalPosition.SubScript != sub) Run(api.ToggleSubscript, M("ToggleSubscript"));
		}
		if (Has(FontPropertiesEnum.ForegroundColorType) || Has(FontPropertiesEnum.ForegroundColor)) {
			if (Has(FontPropertiesEnum.ForegroundColorType) && (SwatchColorType)Number(FontPropertiesEnum.ForegroundColorType) != SwatchColorType.RGB)
				Run(api.ClearFontColor, M("ClearFontColor"));
			else if (changes.GetValueOrDefault(FontPropertiesEnum.ForegroundColor) is Color color)
				Run(() => SetFontColor(color));
		}
		if (Has(FontPropertiesEnum.BackgroundColorType) || Has(FontPropertiesEnum.BackgroundColor)) {
			if (Has(FontPropertiesEnum.BackgroundColorType) && (SwatchColorType)Number(FontPropertiesEnum.BackgroundColorType) != SwatchColorType.RGB)
				Run(api.ClearHighlight, M("ClearHighlight"));
			else if (changes.GetValueOrDefault(FontPropertiesEnum.BackgroundColor) is Color color)
				Run(() => SetHighlight(color));
		}
	}

	void SetFontColor(Color c) {
		api.SetFontColor(c);
		Record(M("SetFontColor", c));
	}

	void SetHighlight(Color c) {
		api.SetBackgroundColor(c);
		Record(M("SetBackgroundColor", c));
	}

	// The recent documents in the File menu (Settings.RecentFiles, newest first).
	void UpdateRecentItems() {
		if (recentItems == null) return;
		recentItems.RecentItems = Settings.RecentFiles.Select(path => new RecentItemsPropertySet { Label = Path.GetFileName(path), LabelDescription = path, Pinned = false }).ToList();
	}

	// The macros in rte_macros, with their hotkeys, in the Tools tab's gallery. Refilled when something changed.
	void FillMacrosGallery(bool force = false) {
		if (macrosGallery == null || !ribbon.Initialized || macrosGallery.ItemsSource is not { } items) return;
		var files = Directory.Exists(MacrosDir) ? Directory.GetFiles(MacrosDir, "*.ahk").OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList() : new List<string>();
		var labels = files.Select(f => Path.GetFileNameWithoutExtension(f) + (HotkeyOf(Path.GetFileName(f)) is Keys keys ? $"    ({TKeys(HotkeyBox.Describe(keys))})" : "")).ToList();
		if (!force && files.SequenceEqual(galleryMacros) && labels.SequenceEqual(galleryLabels)) return;
		(galleryMacros, galleryLabels) = (files, labels);
		items.Clear();
		foreach (var label in labels) items.Add(new GalleryItemPropertySet { Label = label, CategoryID = Constants.UI_Collection_InvalidIndex });
		if (labels.Count == 0) items.Add(new GalleryItemPropertySet { Label = "(no .ahk files in rte_macros)", CategoryID = Constants.UI_Collection_InvalidIndex });
	}

	List<string> galleryLabels = new();

	// ---- Extensions (iteration 61: hooks for what rte_extensions.cs loads; AhkWord itself adds none)

	public sealed record ExtensionCommand(string Name, string Description, Action Run);
	readonly List<ExtensionCommand> extensionCommands = new();
	readonly ToolStripStatusLabel extensionLabel = new() { Visible = false };

	/// Values for template fields (a field's name -> its value, or null if it isn't known), asked before the Fill In
	/// Fields dialog: a field given a value here isn't asked for.
	public Func<string, string?>? FieldSource { get; set; }

	/// An extension's own template language (iteration 62): given RTF (a template's, or the document's for Fill In
	/// Fields) and whether it may ask for values, returns it filled in, or null when it isn't in that language (AhkWord's
	/// own {{fields}} left in it are filled in afterwards). An error it throws shows in the status bar.
	public Func<string, bool, string?>? RenderTemplate { get; set; }

	/// The same for a header's or footer's part (plain text): filled in with the values the template's text had, or null
	/// when it isn't in that language (iteration 62).
	public Func<string, string?>? RenderTemplateText { get; set; }

	// the header's and footer's parts (each page's, the first page's)
	static IEnumerable<string[]> BandParts(PageBands bands) {
		yield return bands.Header;
		yield return bands.Footer;
		if (bands.FirstHeader != null) yield return bands.FirstHeader;
		if (bands.FirstFooter != null) yield return bands.FirstFooter;
	}

	// (iteration 63) the document's header and footer, and those of the sections that have their own
	IEnumerable<string[]> BandParts() => BandParts(Bands).Concat(Sections.Breaks(Editor.TextDocument)
		.Where(b => !b.Settings.Same && b.Settings.Bands != null).SelectMany(b => BandParts(b.Settings.Bands!)));

	/// The text whose fields are asked for: the document's (iteration 63: without the section breaks' markers, whose
	/// headers and footers are the bands'), then the header's and footer's parts.
	string FieldText() => Sections.Mark.Replace(Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "", "") + "\n" + BandsText();

	/// The header's and footer's parts as one text (their fields asked for with the document's).
	string BandsText() => string.Join("\n", BandParts().SelectMany(parts => parts));

	/// Fills in the header's and footer's parts (iteration 62): an extension's template language, then AhkWord's own
	/// fields (value: a field's value, or null); (iteration 63) the sections' own too. Returns how many parts changed.
	int FillBands(Func<string, string?> value) {
		int Fill(PageBands bands) {
			int changed = 0;
			foreach (var parts in BandParts(bands))
				for (int i = 0; i < parts.Length; i++) {
					string part = parts[i] ?? "";
					string filled = Templates.FillText(RenderTemplateText?.Invoke(part) ?? part, value);
					if (filled == part) continue;
					parts[i] = filled;
					changed++;
				}
			return changed;
		}
		int count = Fill(Bands);
		var doc = Editor.TextDocument;
		foreach (var b in Sections.Breaks(doc).Where(b => !b.Settings.Same && b.Settings.Bands != null).Reverse()) {    // (the last first: the markers before stay where they are)
			int changed = Fill(b.Settings.Bands!);
			if (changed == 0) continue;
			Sections.Write(doc, b, b.Settings);
			count += changed;
		}
		if (count > 0) { Editor.Modified = true; BandsChangedNow(); }
		return count;
	}

	// the document filled in by the extension's template language, as one undo step; false if it isn't in it
	bool RenderDocument(bool ask) {
		// (the page view's gaps left out: they would become the document's own space and raised lines, in the wrong places
		// once the text had moved)
		if (RenderTemplate?.Invoke(WithoutGaps(() => Editor.Rtf) ?? "", ask) is not string rendered) return false;
		var doc = Editor.TextDocument;
		WithoutGapsChanging(() => {
			doc.BeginEditCollection();
			try {
				Editor.SelectAll();
				Editor.SelectedRtf = rendered;
				RtfTables.DropStrayMergeMarks(doc);    // (iteration 92)
				Editor.Select(0, 0);
			} finally {
				doc.EndEditCollection();
			}
		});
		return true;
	}

	/// Adds a command to Tools > Extensions (shown once there is one) and to the command palette; it runs as a ribbon
	/// command does (an error shows in the status bar).
	public void AddExtensionCommand(string name, string description, Action run) {
		extensionCommands.RemoveAll(c => c.Name == name);
		extensionCommands.Add(new ExtensionCommand(name, description, run));
		if (ribbon.Initialized) ribbon.SetModes(0, 1);    // (the Extensions group is in mode 1)
		FillExtensionsGallery();
	}

	/// An extension's own place in the status bar, beside the messages ("" hides it).
	public void SetExtensionStatus(string text) {
		extensionLabel.Text = text;
		extensionLabel.Visible = text != "";
	}

	void FillExtensionsGallery() {
		if (extensionsGallery == null || !ribbon.Initialized || extensionsGallery.ItemsSource is not { } items) return;
		items.Clear();
		foreach (var command in extensionCommands) items.Add(new GalleryItemPropertySet { Label = command.Name, CategoryID = Constants.UI_Collection_InvalidIndex });
	}

	// A cheap test for the Table Tools tab, made on every selection change: RichEdit marks paragraphs in table rows
	// (api.InTable() reads the whole document).
	bool CaretInTable() {
		var pf = new PARAFORMAT2 { cbSize = Marshal.SizeOf<PARAFORMAT2>(), rgxTabs = new int[32] };
		SendMessage(Editor.Handle, EM_GETPARAFORMAT, 0, ref pf);
		return (pf.wEffects & PFE_TABLE) != 0;
	}

	const int EM_GETPARAFORMAT = 0x43D;
	const ushort PFE_TABLE = 0x4000;

	[StructLayout(LayoutKind.Sequential)]
	struct PARAFORMAT2
	{
		public int cbSize; public uint dwMask; public ushort wNumbering, wEffects; public int dxStartIndent, dxRightIndent, dxOffset;
		public ushort wAlignment; public short cTabCount;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] public int[] rgxTabs;
		public int dySpaceBefore, dySpaceAfter, dyLineSpacing; public short sStyle; public byte bLineSpacingRule, bOutlineLevel;
		public ushort wShadingWeight, wShadingStyle, wNumberingStart, wNumberingStyle, wNumberingTab, wBorderSpace, wBorderWidth, wBorders;
	}

	[DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, ref PARAFORMAT2 lParam);

	// ---------------------------------------------------------------- Commands

	// Runs a UI command. Errors (e.g. a feature that is unavailable in the current mode) go to the status bar.
	void Run(Action action, string? macro = null) {
		try {
			ApplySource();    // what was typed in the source view goes into the document first
			inCommand = true;
			action();
			if (macro != null) Record(macro);
		} catch (Exception ex) {
			if (ex is TargetInvocationException { InnerException: { } inner }) ex = inner;
			SetStatus("⚠ " + ex.Message);
			System.Media.SystemSounds.Beep.Play();
		} finally {
			inCommand = false;
		}
		RefreshSource();    // and what the command changed shows in the source
		lineGutter.DocumentChanged();
		UpdateUi();
		UpdateTitle();
		Restart(outlineTimer);
	}

	public void SetStatus(string text) => statusText.Text = text;

	/// (iteration 58) Progress in the status bar, painted at once: while a document prints, no messages are handled.
	internal void ShowProgress(string text) {
		statusText.Text = text;
		statusText.GetCurrentParent()?.Update();
	}

	/// Whether Esc was pressed since the last time asked, with AhkWord's window in front (a long print or export: Esc
	/// stops it).
	internal bool EscPressed() {
		bool pressed = (GetAsyncKeyState(0x1B) & 0x8001) != 0;    // VK_ESCAPE: down now, or pressed since
		return pressed && GetForegroundWindow() == Handle;
	}
	/// Whether a print or export should stop: Esc (tests: PrintStopOverride, asked instead).
	internal bool PrintStop() => PrintStopOverride?.Invoke() ?? EscPressed();
	internal Func<bool>? PrintStopOverride = null;    // (set by tests)
	[DllImport("user32.dll")] static extern short GetAsyncKeyState(int key);
	[DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

	static void Restart(System.Windows.Forms.Timer timer) {
		timer.Stop();
		timer.Start();
	}

	internal void SetFileName(string? path) {
		FileName = path;
		UpdateTitle();
	}

	// ---- Autosave and recovery (iteration 29)

	internal string RecoveryFolder => recovery.Folder;

	/// Writes recovery copies every AutosaveSeconds while there are unsaved changes (0: none).
	internal void StartAutosave() {
		autosaveTimer.Stop();
		if (Settings.AutosaveSeconds <= 0) { ClearRecovery(); return; }
		autosaveTimer.Interval = Math.Max(10, Settings.AutosaveSeconds) * 1000;
		autosaveTimer.Start();
	}

	bool autosaveFailed;

	// The recovery copy: written while there are unsaved changes (if they changed since), removed once there are none.
	internal void Autosave() {
		// (iteration 87: not while the editor's window is made again, New and Open's since iteration 86: its text can't be
		// read then; the next tick writes the copy)
		if (Editor.Renewing || !Editor.IsHandleCreated) return;
		if (!Editor.Modified) { ClearRecovery(); return; }
		try {
			recovery.Write(api.RecoveryRtf(), FileName, DocumentName);
			autosaveFailed = false;
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
			// (iteration 87) the text not readable for the moment: RichTextBox's "Text is not saved", RichEditorNET's "The
			// control handle has not been created", TOM's errors; on the timer, they ended in the crash dialog
			or InvalidOperationException or COMException) {
			if (!autosaveFailed) SetStatus("⚠ " + T("status.recoveryNotWritten", "The recovery copy could not be written: {0}", ex.Message));
			autosaveFailed = true;
		}
	}

	internal void ClearRecovery() {
		try { recovery.Clear(); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
	}

	// At the start: copies that an AhkWord which didn't close normally left behind are offered, the newest first (one
	// a start; the others the next times). Started with a file, they wait for a start without one.
	// Returns whether a copy was recovered.
	bool OfferRecovery(bool startedWithFile) {
		List<Recovery.Orphan> orphans;
		try { orphans = Recovery.Orphans(recovery.Folder); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
		if (orphans.Count == 0) return false;
		var newest = orphans[0];
		string when = T("ui.timeOnDate", "{0} on {1}", newest.Info.Saved.ToShortTimeString(), newest.Info.Saved.ToShortDateString());
		if (startedWithFile) {
			SetStatus(T("status.recoverLater", "AhkWord didn't close normally last time: start it without a file to recover \"{0}\" (autosaved at {1}).", newest.Info.Name, when));
			return false;
		}
		string more = orphans.Count > 1 ? "\n\n" + TN("msg.recoverMore", "({0} more recovery copy is offered at the next start.)", "({0} more recovery copies are offered at the next start.)", orphans.Count - 1) : "";
		var answer = MessageBox.Show(this, T("msg.recover", "AhkWord didn't close normally last time. It had autosaved \"{0}\" at {1}.\n\n"
			+ "Yes: recover it (then save it to keep it).\nNo: discard the recovery copy.\nCancel: decide at the next start.", newest.Info.Name, when) + more,
			AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
		if (answer == DialogResult.Yes) { Recover(newest); return true; }
		if (answer == DialogResult.No) Recovery.Discard(newest);
		return false;
	}

	/// Restores a copy left behind, as unsaved changes to the document it was (Save replaces its file), and removes it.
	internal void Recover(Recovery.Orphan orphan) {
		try {
			api.Recover(Recovery.Read(orphan), orphan.Info.FileName);
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
			MessageBox.Show(this, T("msg.recoveryUnreadable", "The recovery copy could not be read: {0}", ex.Message), AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}
		Recovery.Discard(orphan);
		Autosave();    // (a copy of its own at once)
		string file = orphan.Info.FileName is { Length: > 0 } f ? " " + T("status.recoveredHasFile", "Save replaces {0}; Save As keeps both.", Path.GetFileName(f)) : " " + T("status.recoveredNoFile", "Save it to keep it.");
		SetStatus(T("status.recovered", "Recovered \"{0}\" as autosaved at {1}.", orphan.Info.Name, orphan.Info.Saved.ToShortTimeString()) + file);
	}

	// ---- Templates and the start page (iteration 30)

	/// The templates beside the program (shipped) and the user's own (beside the settings).
	internal string ShippedTemplatesFolder => Path.Combine(appDir, "templates");
	internal string OwnTemplatesFolder => Path.Combine(Path.GetDirectoryName(Path.GetFullPath(Settings.FilePath))!, "Templates");
	internal List<TemplateInfo> TemplateList() => Templates.Find(ShippedTemplatesFolder, OwnTemplatesFolder);

	// Insert from Template's list: the parts (pieces made to be inserted) first, then the documents.
	internal List<TemplateInfo> InsertChoices() => TemplateList().OrderBy(t => !t.Part).ThenBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase).ToList();

	StartPage? startPage;

	/// Whether the start page is showing (in the editor's place).
	internal bool StartShown => startPage is { Visible: true };

	// Leaves the start page and Multiple Pages (for a command that acts on the document).
	void LeaveOverlays() {
		if (PagesShown) ShowPages(false);
		if (StartShown) ShowStart(false);
	}

	/// Shows the start page (show) or goes back to the document.
	internal void ShowStart(bool show) {
		if (!show) {
			if (!StartShown) return;
			startPage!.Visible = false;
			Editor.Focus();
			UpdateUi();
			return;
		}
		if (PagesShown) ShowPages(false);
		if (startPage == null) {
			startPage = new StartPage { Dock = DockStyle.Fill, Visible = false };
			startPage.BlankChosen += () => { ShowStart(false); if (Editor.TextLength > 0 || FileName != null) Run(() => { if (ConfirmDiscard()) api.New(); }); };
			startPage.Closing += () => ShowStart(false);
			startPage.TemplateChosen += template => { ShowStart(false); Run(() => NewFromTemplate(template, null, ask: true)); };
			startPage.RecentChosen += path => { ShowStart(false); Run(() => { if (ConfirmDiscard()) OpenRecent(path); }); };
			startPage.OpenChosen += () => { ShowStart(false); Run(OpenDialog); };
			startPage.MacroChosen += path => { ShowStart(false); Run(() => RunMacroFromUi(path)); };
			startPage.TemplatesFolderChosen += () => Run(() => {
				Directory.CreateDirectory(OwnTemplatesFolder);
				System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(OwnTemplatesFolder) { UseShellExecute = true });
				SetStatus(T("status.templatesFolder", "Your templates folder: put documents there (with {{Field}} fields if you like) to use them as templates."));
			});
			startPage.ShowAtStartChanged += on => { Settings.ShowStartPage = on; Settings.Save(); };
			sourceSplit.Parent!.Controls.Add(startPage);
		}
		var macros = new List<(string, string, string?)>();
		if (Directory.Exists(MacrosDir))
			foreach (var file in Directory.GetFiles(MacrosDir, "*.ahk").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
				macros.Add((Path.GetFileNameWithoutExtension(file), file, Settings.MacroHotkeys.TryGetValue(Path.GetFileName(file), out var keys) ? keys : null));
		startPage.Visible = true;
		startPage.BringToFront();
		startPage.Fill(TemplateList().Where(t => !t.Part).ToList(), Settings.RecentFiles, macros, Settings.ShowStartPage);
		(startPage.FirstTile ?? (Control)startPage).Focus();
		UpdateUi();
	}

	/// Makes an untitled document from a template: its fields are filled from values (name -> value), and, if ask, the
	/// others are asked for (with the values typed last time). Returns false if the user kept the document they had.
	internal bool NewFromTemplate(TemplateInfo template, IDictionary<string, string>? values, bool ask) {
		if (ask && !ConfirmDiscard()) return false;
		api.OpenTemplate(template.Path);
		if (RenderTemplate?.Invoke(Editor.Rtf ?? "", ask) is string rendered) Editor.Rtf = rendered;    // (iteration 62: an extension's template language)
		var doc = Editor.TextDocument;
		var given = FieldValues(T("dlg.fields.titleNew", "New from {0}", template.Name), FieldText(), values, ask);
		int filled = Templates.Fill(doc, name => given.TryGetValue(name, out var value) ? value : null);
		filled += FillBands(name => given.TryGetValue(name, out var value) ? value : null);    // (iteration 62)
		Editor.ClearUndo();
		Editor.Select(0, 0);
		Editor.Modified = filled > 0;
		SetStatus(T("status.newFromTemplate", "New document from the {0} template", template.Name) + (filled > 0 ? " " + TN("status.fieldsFilledIn", "({0} field filled in).", "({0} fields filled in).", filled) : "."));
		return true;
	}

	// The values for the fields in text: those given, and for the others, if ask, what the dialog was given (the values
	// of fields about the user, "Your ...", remembered), else for those the values typed last time (the others stay to
	// fill in).
	Dictionary<string, string> FieldValues(string title, string text, IDictionary<string, string>? values, bool ask) {
		var given = new Dictionary<string, string>(values ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);
		if (FieldSource != null)    // (iteration 61: an extension's values, such as a matter's, before asking)
			foreach (var (name, _) in Templates.Fields(text))
				if (!given.ContainsKey(name) && FieldSource(name) is { } known) given[name] = known;
		var asked = Templates.Fields(text).Where(f => !given.ContainsKey(f.Name)).ToList();
		if (ask && asked.Count > 0) {
			using var dlg = new TemplateFieldsDialog(title, asked, Settings.TemplateValues, () => { Settings.TemplateValues.Clear(); Settings.Save(); });
			if (dlg.ShowDialog(this) == DialogResult.OK) {
				foreach (var (name, value) in dlg.Values) {
					given[name] = value;
					if (value != "" && Templates.IsRemembered(name)) Settings.TemplateValues[name] = value;    // (only "Your ..." fields)
				}
				// (values kept by earlier versions for other fields, such as a paper's parties, go)
				foreach (var old in Settings.TemplateValues.Keys.Where(k => !Templates.IsRemembered(k)).ToList()) Settings.TemplateValues.Remove(old);
				Settings.Save();
			}
		} else {
			foreach (var (name, _) in asked)
				if (Templates.IsRemembered(name) && Settings.TemplateValues.TryGetValue(name, out var last)) given[name] = last;
		}
		return given;
	}

	// ---- Snippets (iteration 55, rte_snippets.cs)

	/// The snippets' list (beside the user's own templates).
	internal string SnippetsFile => Snippets.FileIn(OwnTemplatesFolder);
	internal List<Snippet> SnippetList() => Snippets.Load(SnippetsFile);

	/// A snippet's part (an RTF template in a Parts folder, or any RTF template of that name), or null.
	internal TemplateInfo? SnippetPart(Snippet snippet) => TemplateList().OrderBy(t => !t.Part)
		.FirstOrDefault(t => string.Equals(t.Name, snippet.Part, StringComparison.OrdinalIgnoreCase) && t.Path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase));

	// Insert > Snippets (iteration 55: a snippet's text typed or pasted in its own dialog, kept as a part)
	void SnippetsCommand() {
		var doc = Editor.TextDocument;
		string? selection = !SourceActive && doc.Selection2.End > doc.Selection2.Start ? WithoutGaps(() => Editor.SelectedRtf ?? "") : null;
		using var font = new Font(Styles.Normal.Font, Styles.Normal.Size);
		using var dlg = new SnippetsDialog(SnippetList(), TemplateList, selection, LoadPart, SavePart, font);
		dlg.Changed += list => Snippets.Save(SnippetsFile, list);
		dlg.Export += path => Run(() => SetStatus(T("status.snippetsExported", "The snippets were exported to {0}. Run it with AutoHotkey v2 to have them in any program.", api.ExportSnippets(path))));
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		if (dlg.ToInsert is { } insert) {
			string abbreviation = insert.Abbreviation;
			Run(() => api.InsertSnippet(abbreviation, "", true), M("InsertSnippet", abbreviation, "", true));
		}
	}

	/// A part's RTF by its name (null: there is none).
	internal string? LoadPart(string name) => SnippetPart(new Snippet("", name)) is { } part && File.Exists(part.Path) ? File.ReadAllText(part.Path, Encoding.Latin1) : null;

	/// Saves a part's RTF: over the part of that name if there is one, else a new one in your own templates' Parts folder.
	internal void SavePart(string name, string rtf) {
		string path = SnippetPart(new Snippet("", name))?.Path ?? Path.Combine(OwnTemplatesFolder, "Parts", name + ".rtf");
		Directory.CreateDirectory(Path.GetDirectoryName(path)!);
		File.WriteAllText(path + ".tmp", rtf, Encoding.Latin1);
		File.Move(path + ".tmp", path, overwrite: true);
	}

	/// A snippet's hotstring fired (AhkWord.ahk, with the ending character typed: "" for one that fires at once): once
	/// the document (or the source view) has the abbreviation and that character before the caret (the keys that typed
	/// them may still be on their way), they are taken out and the snippet's part goes in (its fields asked for), as one
	/// undo step; in the Markdown or HTML view, the part's plain text. Gives up after half a second (the abbreviation
	/// stays as typed). Returns at once.
	/// (target: "rich" or "source", the document or the source view, whichever has the focus or not: tests.)
	internal void FireSnippet(string abbreviation, string ending, string? target = null) {
		if (SnippetList().FirstOrDefault(s => string.Equals(s.Abbreviation, abbreviation, s.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase)) is not { } snippet) return;
		if (SnippetPart(snippet) is not { } part) { SetStatus("⚠ " + T("status.snippetPartGone", "The snippet \"{0}\" inserts the part {1}, which isn't in the templates' Parts folders any more.", abbreviation, snippet.Part)); return; }
		bool source = target == "source" || target == null && view != SourceKind.None && sourceBox.Focused;
		if (target == null && !source && !Editor.Focused) return;
		int tries = 0;
		var timer = new System.Windows.Forms.Timer { Interval = 15 };
		timer.Tick += (s, e) => {
			timer.Stop();
			if (TryFireSnippet(snippet, part, ending, source) || ++tries >= 33) timer.Dispose();
			else timer.Start();
		};
		timer.Start();
	}

	// The snippet in place of what typed it, if the text before the caret is that now.
	bool TryFireSnippet(Snippet snippet, TemplateInfo part, string ending, bool source) {
		int kept = snippet.KeepEnding && !snippet.Instant && ending != "" ? 1 : 0;
		if (source) {
			// (the view's lines end in "\n": made "\r", as Typed has Enter)
			int caret = sourceBox.CurrentPosition, from = Math.Max(0, caret - 60);
			string before = sourceBox.GetTextRange(from, caret - from).Replace('\n', '\r');
			int length = Snippets.Typed(before, before.Length, snippet, ending);
			if (length < 0) return false;
			string text = SnippetPlainText(part);
			var given = FieldValues(T("dlg.fields.titleInsert", "Insert {0}", part.Name), text, null, ask: true);
			text = Templates.FillText(text, name => given.TryGetValue(name, out var value) ? value : null).Replace("\r\n", "\n").Replace('\r', '\n');
			int begin = caret - length;
			sourceBox.DeleteRange(begin, length - kept);
			sourceBox.InsertText(begin, text);
			sourceBox.GotoPosition(begin + text.Length + kept);
			return true;
		}
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		if (selection.End > selection.Start) return false;
		int at = selection.Start;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int typedLength = Snippets.Typed(story, at, snippet, ending);
		if (typedLength < 0) return false;
		int start = at - typedLength;
		doc.BeginEditCollection();
		try {
			doc.Range2(start, at - kept).Text = "";
			doc.Selection2.SetRange(start, start);
			InsertTemplate(part, null, ask: true, inline: true);
			if (kept > 0) { int after = doc.Selection2.Start; doc.Selection2.SetRange(after + kept, after + kept); }
		} finally {
			doc.EndEditCollection();
		}
		return true;
	}

	/// A part's plain text (its RTF read by a RichTextBox of its own; markers left out).
	internal static string SnippetPlainText(TemplateInfo part) {
		using var box = AppTheme.OnPaper(new RichTextBox());
		box.Rtf = File.ReadAllText(part.Path, Encoding.Latin1);
		return CrossRefs.PlainText(box.Text).TrimEnd('\n', '\r');
	}

	// File > New from Template: the documents (not the parts) to choose from.
	void NewFromTemplateCommand() {
		using var dlg = new TemplatePickerDialog(T("dlg.templates.titleNew", "New from Template"), TemplateList().Where(t => !t.Part).ToList(), OwnTemplatesFolder);
		if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Chosen is { } template) NewFromTemplate(template, null, ask: true);
	}

	// Insert > Insert from Template: the parts first, then the documents.
	void InsertTemplateCommand() {
		using var dlg = new TemplatePickerDialog(T("dlg.templates.titleInsert", "Insert from Template"), InsertChoices(), OwnTemplatesFolder, insertAsSection);
		if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Chosen is not { } template) return;
		insertAsSection = dlg.AsSection;    // (remembered while AhkWord runs)
		InsertTemplate(template, null, ask: true, asSection: dlg.AsSection);
	}
	bool insertAsSection;

	/// Inserts a template's content at the caret (replacing the selection), in one undo step, and fills in its fields
	/// (only the inserted ones) as NewFromTemplate does. Returns how many were filled in.
	/// (inline, iteration 55: a snippet's part goes in where the caret is, without the paragraph mark its text ends with,
	/// which RichEdit writes for the last paragraph: inserted, that started a new paragraph after every snippet)
	internal int InsertTemplate(TemplateInfo template, IDictionary<string, string>? values, bool ask, bool inline = false, bool asSection = false) {
		if (!template.Path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
			throw new InvalidOperationException($"Only RTF templates can be inserted; open {Path.GetFileName(template.Path)} and save it as Rich Text Format.");
		string file = File.ReadAllText(template.Path, Encoding.Latin1);
		string rtf = HeaderFooter.StripFromRtf(file);    // (the header and footer stay the document's)
		if (RenderTemplate?.Invoke(rtf, ask) is string rendered) rtf = rendered;    // (iteration 62)
		if (inline) rtf = Snippets.WithoutFinalParagraph(rtf);
		// (iteration 63) as a new section: its own header and footer, the template's (the section before's if it has none)
		var bands = asSection ? HeaderFooter.FromRtf(file) : null;
		var doc = Editor.TextDocument;
		doc.BeginEditCollection();
		int filled;
		try {
			if (asSection) Sections.Insert(Editor, new SectionSettings { Same = bands == null, Bands = bands, Restart = true, Start = 1 });
			int start = doc.Selection2.Start;
			Editor.SelectedRtf = rtf;
			int end = doc.Selection2.Start;
			if (RtfTables.DropStrayMergeMarks(doc) > 0) end = doc.Selection2.Start;    // (iteration 92)
			string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
			string asked = story[Math.Min(start, story.Length)..Math.Min(end, story.Length)];
			if (bands != null) asked += "\n" + string.Join("\n", BandParts(bands).SelectMany(parts => parts));    // (its header's and footer's fields too)
			var given = FieldValues(T("dlg.fields.titleInsert", "Insert {0}", template.Name), asked, values, ask);
			int before = Editor.TextLength;
			filled = Templates.Fill(doc, name => given.TryGetValue(name, out var value) ? value : null, start, end);
			Editor.Select(Math.Max(start, end + Editor.TextLength - before), 0);    // (after what was inserted)
			if (bands != null) filled += FillBands(name => given.TryGetValue(name, out var value) ? value : null);
		} finally {
			doc.EndEditCollection();
		}
		if (asSection) BandsChangedNow();
		SetStatus((asSection ? T("status.insertedTemplateSection", "Inserted the {0} template as a new section", template.Name) : T("status.insertedTemplate", "Inserted the {0} template", template.Name))
			+ (filled > 0 ? " " + TN("status.fieldsFilledIn", "({0} field filled in).", "({0} fields filled in).", filled) : "."));
		return filled;
	}

	// Insert > Fill In Fields
	void FillFieldsCommand() {
		var doc = Editor.TextDocument;
		bool rendered = RenderDocument(ask: true);    // (iteration 62: an extension's template language first)
		if (Templates.Fields(FieldText()).Count == 0) {
			bool bands = FillBands(_ => null) > 0;    // (an extension's language in the header or footer)
			SetStatus(rendered || bands ? T("status.templateFilled", "The template was filled in.") : T("status.noFields", "There are no {{fields}} to fill in in this document (fields are written {{Name}} or {{Name|default}})."));
			return;
		}
		int filled = FillFields(null, ask: true, rendered: true);
		SetStatus(rendered
			? (filled > 0 ? TN("status.templateFilledFields", "The template was filled in; {0} field filled in.", "The template was filled in; {0} fields filled in.", filled) : T("status.templateFilledNoFields", "The template was filled in; no fields were filled in."))
			: (filled > 0 ? TN("status.fieldsFilled", "{0} field filled in.", "{0} fields filled in.", filled) : T("status.noFieldsFilled", "no fields were filled in.")));
	}

	/// Fills in the document's fields (one undo step): values given, then (ask) the dialog, or the values typed last time;
	/// an extension's template language first (rendered: done already).
	internal int FillFields(IDictionary<string, string>? values, bool ask, bool rendered = false) {
		if (!rendered) RenderDocument(ask);
		var doc = Editor.TextDocument;
		var given = FieldValues(T("dlg.fields.title", "Fill In Fields"), FieldText(), values, ask);
		int filled;
		doc.BeginEditCollection();
		try { filled = Templates.Fill(doc, name => given.TryGetValue(name, out var value) ? value : null); }
		finally { doc.EndEditCollection(); }
		return filled + FillBands(name => given.TryGetValue(name, out var value) ? value : null);    // (iteration 62: the header and footer too)
	}

	// The recent list without a template (a document made from it is a new one)
	internal void ForgetRecent(string path) {
		if (Settings.RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase)) == 0) return;
		Settings.Save();
		UpdateRecentItems();
	}

	// ---- Paragraph styles (iteration 32)

	/// The document's styles.
	internal StyleSheet Styles = StyleSets.Get("Classic");

	/// The styles new documents get: yours (Save Styles as My Default), else your style set's.
	internal StyleSheet DefaultStyles() => Settings.MyStyles?.Copy() is { Styles.Count: > 0 } mine ? mine : StyleSets.Get(Settings.StyleSetName);

	/// The editor's base font (text without a font of its own: a new document's, HTML's and Markdown's): Normal's. Set
	/// before a document is loaded (it applies to all the text there is).
	internal void UseNormalFont() {
		var normal = Styles.Normal;
		try { Editor.Font = new Font(normal.Font, normal.Size); } catch (ArgumentException) { }
	}

	/// After a document is loaded (its paragraphs' styles set): its style sheet. defined: the styles the file defined
	/// (AhkWord's own RTF), else the default ones. The styles its paragraphs use that weren't defined are defined from
	/// the first paragraph with each (the document keeps its look), or, with takeLook (Markdown, plain text, which have
	/// no look of their own), its paragraphs are given the styles' look.
	/// Iteration 38: the document's list levels (Paragraph > List levels): numbering, start and indents of each level,
	/// saved in AhkWord's RTF ({\*\rtewriterlevels}); a new document, or one from elsewhere, has the default ones
	/// (Settings.ListLevels). Always 9.
	internal List<ListLevel> Levels = new();
	static bool LevelsSame(List<ListLevel> a, List<ListLevel> b) => a.Count == b.Count && a.Zip(b).All(p =>
		p.First.Style == p.Second.Style && p.First.Punctuation == p.Second.Punctuation && p.First.Start == p.Second.Start
		&& p.First.NumberAt == p.Second.NumberAt && p.First.TextAt == p.Second.TextAt);

	/// The document's list levels changed: legal numbering's items take their levels' indents (they all follow them,
	/// one numbering); the document has changed.
	internal void LevelsChanged() {
		var doc = Editor.TextDocument;
		if (LegalLists.Paragraphs(doc).Any(p => p.Legal && !p.Styled)) LegalLists.Grouped(Editor, () => { });
		Editor.Modified = true;
	}

	internal List<ListLevel> DefaultLevels() => Enumerable.Range(0, 9).Select(i => (i < Settings.ListLevels.Count ? Settings.ListLevels[i] : global::ListLevel.Defaults()[i]).Copy()).ToList();

	/// The numbered items' ListStart values the RTF being loaded kept (LegalLists.StartsFrom), for StylesAfterLoad.
	internal List<(int Start, int Level)>? LoadedListStarts;

	internal void StylesAfterLoad(StyleSheet? defined, bool takeLook, bool styledLook = false) {
		var sheet = defined?.Copy() ?? DefaultStyles();
		foreach (var usual in DefaultStyles().Styles) if (sheet.Find(usual.Name) == null) sheet.Styles.Add(usual.Copy());    // (the usual styles are always there)
		Styles = sheet;
		var doc = Editor.TextDocument;
		var firsts = new Dictionary<int, int>();
		foreach (var (start, _) in StyleEngine.Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)) {
			var para = doc.Range2(start, start).Para;
			if (StyleEngine.NameOf(para.Style) == null) StyleEngine.Tag(para, -1);    // (a number no style has: Normal)
			firsts.TryAdd(para.Style, start);
		}
		bool own = defined is { Set: "" };    // (the built-in reader's: a .docx's own styles)
		foreach (var (id, start) in firsts) {
			var style = sheet.Ensure(StyleEngine.NameOf(id)!);
			if (takeLook) continue;
			if (defined?.Find(style.Name) != null) continue;
			int i = sheet.Styles.IndexOf(style);
			sheet.Styles[i] = StyleEngine.Capture(doc, start, style.Name, style.Level, style.Next);
			own = true;
		}
		// (styles taken from the document's look: they are the document's own, not a set's)
		if (own && !takeLook) sheet.Set = "";
		// (styledLook: the paragraphs with a style other than Normal get its characters' look: HTML's readers give
		// headings sizes of their own; their paragraph formats are the file's, and so are Normal paragraphs)
		if (styledLook && !takeLook) {
			doc.Freeze();
			doc.BeginEditCollection();
			try {
				foreach (var (start, _) in StyleEngine.Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)) {
					var style = StyleEngine.StyleAt(doc, sheet, start);
					if (style != sheet.Normal) StyleEngine.ApplyCharacters(doc, sheet, style, start);
				}
			} finally {
				doc.EndEditCollection();
				doc.Unfreeze();
			}
		}
		if (takeLook) {
			doc.BeginEditCollection();
			try {
				foreach (var (start, _) in StyleEngine.Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)) {
					var style = StyleEngine.StyleAt(doc, sheet, start);
					var para = doc.Range2(start, start).Para;
					// (Normal paragraphs: their characters; lists and quotes keep their indents)
					if (style == sheet.Normal || (para.ListType & 0xFFFF) != tomConstants.tomListNone) StyleEngine.ApplyCharacters(doc, sheet, style, start);
					else StyleEngine.Apply(doc, sheet, style, start, start);
				}
			} finally {
				doc.EndEditCollection();
			}
		}
		// numbered styles: their paragraphs items of their numbering (a Word document's numbered styles become ones), and
		// the items' restarts as AhkWord's RTF kept them
		if (StyleEngine.NumberAfterLoad(doc, sheet, LoadedListStarts)) Editor.ClearUndo();
		LoadedListStarts = null;
		FillStyleBox(force: true);
		styleSetShown = null;
		ShowStyleSet();
	}

	List<string> styleBoxNames = new();
	RibbonDropDownButton? stylesMenu;
	string? styleSetShown;

	/// The Styles drop-down's label: the document's style set, and whether its styles have been changed from the set's.
	internal string StyleSetLabel() {
		if (Styles.Set == "") return T("ui.stylesDocument", "Styles: The document's");    // (its own, from the file: Apply My Styles gives it yours)
		var set = StyleSets.Get(Styles.Set);
		bool changed = set.Styles.Any(s => Styles.Find(s.Name) is not { } mine || System.Text.Json.JsonSerializer.Serialize(mine) != System.Text.Json.JsonSerializer.Serialize(s));
		// (the set's name in the language: "Clásico" in Spanish)
		string name = set.Set switch {
			"Classic" => T("styleSet.classic", "Classic"), "Plain" => T("styleSet.plain", "Plain"), "Modern" => T("styleSet.modern", "Modern"), var other => other
		};
		return changed ? T("ui.stylesSetChanged", "Styles: {0} (changed)", name) : T("ui.stylesSet", "Styles: {0}", name);
	}

	// (only when the styles may have changed: StylesChanged, a document loaded, a set chosen)
	void ShowStyleSet(bool force = false) {
		if (stylesMenu == null || !ribbon.Initialized) return;
		if (!force && styleSetShown != null) return;
		string label = StyleSetLabel();
		if (label == styleSetShown) return;
		styleSetShown = label;
		try { SetRibbonLabel(stylesMenu, RibbonCommand.StylesMenu, label); } catch (COMException) { }
	}

	// Iteration 48: a ribbon control's label changed without RibbonLib's Label setter. The setter tells the ribbon with
	// a pointer to a key it has just made (PropertyKeyRef.From); the ribbon queues it for its timer, and a garbage
	// collection in between moved or freed the key: the timer read freed memory, and AhkWord ended with an access
	// violation in UIRibbon.dll (COfficeSpaceUser::OnUpdate > InvalidateUICommandInternal), now and then after opening
	// a long document. Here the label goes where RibbonLib keeps it, and the ribbon is told with a key pinned for good.
	static readonly GCHandle pinnedLabelKey = GCHandle.Alloc(PropertyKeyRef.From(RibbonProperties.Label), GCHandleType.Pinned);
	void SetRibbonLabel(RibbonDropDownButton control, RibbonCommand id, string label) {
		const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
		var provider = typeof(RibbonDropDownButton).GetField("_labelPropertiesProvider", Private)?.GetValue(control);
		var field = provider?.GetType().GetField("_label", Private);
		if (field == null || !ribbon.Initialized) { control.Label = label; return; }    // (another RibbonLib: its own setter)
		field.SetValue(provider, label);
		ribbon.Framework.InvalidateUICommand((uint)id, Invalidations.Property, (PropertyKeyRef)pinnedLabelKey.Target!);
	}

	// The Styles gallery: the document's styles, in the sheet's order, a tile for each.
	internal void FillStyleBox(bool force = false) {
		if (blockCombo == null || !ribbon.Initialized || blockCombo.ItemsSource is not { } items) return;
		var names = Styles.Styles.Select(s => s.Name).ToList();
		if (!force && names.SequenceEqual(styleBoxNames)) return;    // (StylesChanged forces it: a tile shows the style's font)
		styleBoxNames = names;
		items.Clear();
		foreach (var style in Styles.Styles)
			using (var tile = StyleTile(style, LogicalToDeviceUnits(72), LogicalToDeviceUnits(52)))
				items.Add(new GalleryItemPropertySet { Label = style.Name, CategoryID = Constants.UI_Collection_InvalidIndex, ItemImage = ribbon.ConvertToUIImage(tile) });
		// (a style sent before the list was there is lost: the box stayed empty on the dark ribbon, which fills it later)
		ribbonSent.Remove((blockCombo, "SelectedItem"));
		BeginInvoke(UpdateUi);
	}

	// A style's tile, as in Word's gallery: a sample in its font, size (made smaller to fit), weight and color, and its name,
	// on white in every theme (a preview shows the page as it prints).
	static Bitmap StyleTile(ParagraphStyle style, int width, int height) {
		var tile = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
		using var g = Graphics.FromImage(tile);
		g.Clear(Color.White);
		// (no TextRenderingHint of its own: AntiAliasGridFit here left the whole window's text without smoothing, iteration
		// 72-73; on white, the usual ClearType is right)
		using (var edge = new Pen(Color.FromArgb(0xD0, 0xD0, 0xD0))) g.DrawRectangle(edge, 0, 0, width - 1, height - 1);
		var look = (style.Bold ? FontStyle.Bold : 0) | (style.Italic ? FontStyle.Italic : 0) | (style.Underline ? FontStyle.Underline : 0);
		Color ink = Color.Black;
		if (style.Color != "") try { ink = ColorTranslator.FromHtml(style.Color); } catch (Exception) { }
		const string Sample = "AaBbCc";
		float pixels = Math.Clamp(style.Size, 8f, 16f) * height / 52f * 96f / 72f;    // (the tile is 52 logical pixels high)
		int nameHeight = height * 15 / 52;
		for (; ; pixels -= 1f) {
			using var font = new Font(style.Font, pixels, look, GraphicsUnit.Pixel);
			var size = g.MeasureString(Sample, font);
			if (pixels > 7f && (size.Width > width - 6 || size.Height > height - nameHeight - 2)) continue;
			using var brush = new SolidBrush(ink);
			g.DrawString(Sample, font, brush, (width - size.Width) / 2, Math.Max(1, (height - nameHeight - size.Height) / 2));
			break;
		}
		using (var nameFont = new Font("Segoe UI", height * 10f / 52f, GraphicsUnit.Pixel))
		using (var format = new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
			g.DrawString(style.Name, nameFont, Brushes.DimGray, new RectangleF(2, height - nameHeight, width - 4, nameHeight - 1), format);
		return tile;
	}

	/// The style of the paragraph at the caret ("" if the selection has several).
	internal string CaretStyleName() {
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		var paragraphs = StyleEngine.Paragraphs(doc, selection.Start, selection.End);
		var names = paragraphs.Select(p => StyleEngine.StyleAt(doc, Styles, p.Start).Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
		return names.Count == 1 ? names[0] : "";
	}

	// The style the source views have for a style: the headings, Normal and Code (null: none).
	static BlockStyle? SourceBlockOf(string name) =>
		string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase) ? BlockStyle.Paragraph
		: string.Equals(name, "Code", StringComparison.OrdinalIgnoreCase) ? BlockStyle.Preformatted
		: ParagraphStyle.HeadingNumber(name) is int n and > 0 ? (BlockStyle)n : null;

	// Enter at the end of a paragraph: the style the new paragraph is to get (null: the same, which RichEdit gives it).
	string? NextStyleAfterEnter() {
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		if (selection.Start != selection.End) return null;
		var (start, end) = StyleEngine.Paragraphs(doc, selection.Start, selection.Start)[0];
		if (selection.Start != end - 1) return null;    // (not at the paragraph's end)
		var style = StyleEngine.StyleAt(doc, Styles, start);
		return style.Next != "" && !string.Equals(style.Next, style.Name, StringComparison.OrdinalIgnoreCase) && Styles.Find(style.Next) != null ? style.Next : null;
	}

	void ApplyStyleAtCaret(string name) {
		if (Styles.Find(name) is not { } style) return;
		using var _ = FirstParagraphAsIs(changes: true);
		var doc = Editor.TextDocument;
		int caret = doc.Selection2.Start;
		StyleEngine.Apply(doc, Styles, style, caret, caret);
		StyleEngine.SetTyping(doc.Selection2.Font, style);
		UpdateUi();
	}

	internal void StyleSetMayHaveChanged() { styleSetShown = null; ShowStyleSet(); }

	// Styles changed: the box, the outline, the line numbers.
	internal void StylesChanged() {
		Editor.Modified = true;
		styleSetShown = null;    // (the label shows the set again, below)
		// (the caret in an empty paragraph: what is typed next in its style, which RichEdit keeps apart from the paragraph)
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		if (selection.Start == selection.End && StyleEngine.Paragraphs(doc, selection.Start, selection.Start)[0] is var (start, end) && end - start <= 1)
			StyleEngine.SetTyping(selection.Font, StyleEngine.StyleAt(doc, Styles, start));
		FillStyleBox(force: true);
		RefreshOutline();
		lineGutter.DocumentChanged();
		UpdateUi();
	}

	/// Changes a style (old to updated) in the sheet and in its paragraphs, in one undo step; returns how many paragraphs
	/// have it.
	internal int ChangeStyle(ParagraphStyle old, ParagraphStyle updated) {
		BeginInvoke(() => RunHook(HookEvent.StyleChanged, updated.Name));    // (iteration 56: once it is done)
		using var _ = FirstParagraphAsIs(changes: true, all: true);
		int index = Styles.Styles.FindIndex(s => string.Equals(s.Name, old.Name, StringComparison.OrdinalIgnoreCase));
		if (index < 0) throw new ArgumentException($"There is no style \"{old.Name}\".");
		if (!string.Equals(old.Name, updated.Name, StringComparison.OrdinalIgnoreCase))
			foreach (var other in Styles.Styles) if (string.Equals(other.Next, old.Name, StringComparison.OrdinalIgnoreCase)) other.Next = updated.Name;
		Styles.Styles[index] = updated;
		var doc = Editor.TextDocument;
		doc.BeginEditCollection();
		int count;
		try { count = StyleEngine.Restyle(doc, Styles, old, updated); StyleEngine.FitSubLevelSpacing(doc, Styles); } finally { doc.EndEditCollection(); }    // (iteration 51: the sub-levels' spacing as it says now)
		StylesChanged();
		return count;
	}

	/// Gives the document a style set: each of its styles becomes the set's (its paragraphs with it); your own styles stay.
	internal void UseStyleSet(string name) => UseStyles(StyleSets.Get(name));

	/// Gives the document a sheet's styles (a style set, or your own): each of its styles becomes the sheet's (its
	/// paragraphs with it; formatting of their own stays); styles the sheet hasn't got stay. Returns how many
	/// paragraphs were restyled.
	internal int UseStyles(StyleSheet set) {
		BeginInvoke(() => RunHook(HookEvent.StyleChanged, "(style set)"));    // (iteration 56)
		using var _ = FirstParagraphAsIs(changes: true, all: true);
		var doc = Editor.TextDocument;
		int count = 0;
		doc.BeginEditCollection();
		try {
			foreach (var style in set.Styles) {
				var old = Styles.Find(style.Name);
				if (old == null) { Styles.Styles.Add(style.Copy()); continue; }
				Styles.Styles[Styles.Styles.IndexOf(old)] = style.Copy();
				count += StyleEngine.Restyle(doc, Styles, old, style);
			}
			Styles.Set = set.Set;
			StyleEngine.FitSubLevelSpacing(doc, Styles);
		} finally {
			doc.EndEditCollection();
		}
		StylesChanged();
		return count;
	}

	// Home > Styles > Modify Style: the caret paragraph's style.
	void ModifyStyleCommand() {
		var style = StyleEngine.StyleAt(Editor.TextDocument, Styles, Editor.TextDocument.Selection2.Start);
		using var dlg = new StyleDialog(style.Copy(), Styles, isNew: false);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		int count = ChangeStyle(style, dlg.Result);
		Record(M("ModifyStyle", style.Name, StyleSettings(dlg.Result)));
		SetStatus(TN("status.styleChanged", "The {1} style changed ({0} paragraph). Styles > Save Styles as My Default keeps it for new documents.", "The {1} style changed ({0} paragraphs). Styles > Save Styles as My Default keeps it for new documents.", count, dlg.Result.Name));
	}

	// Home > Styles > New Style: from the caret paragraph, which gets it.
	void NewStyleCommand() {
		var doc = Editor.TextDocument;
		int n = 1;
		while (Styles.Find("Style " + n) != null) n++;
		var start = StyleEngine.Capture(doc, doc.Selection2.Start, "Style " + n, 0, "");
		using var dlg = new StyleDialog(start, Styles, isNew: true);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		api.NewStyle(dlg.Result.Name, StyleSettings(dlg.Result));
		Record(M("NewStyle", dlg.Result.Name, StyleSettings(dlg.Result)));
		SetStatus(TN("status.newStyle", "New style {1}, given to the paragraph at the caret.", "New style {1}, given to the paragraphs at the caret.", StyleEngine.Paragraphs(doc, doc.Selection2.Start, doc.Selection2.End).Count, dlg.Result.Name));
	}

	// Home > Styles > Update Style to Match: the caret paragraph's style becomes like it.
	void UpdateStyleCommand() {
		string name = api.UpdateStyleToMatch();
		Record(M("UpdateStyleToMatch"));
		SetStatus(T("status.styleFromParagraph", "The {0} style now looks like the paragraph at the caret, and so do its other paragraphs.", name));
	}

	// Home > Styles > Delete Style: the caret paragraph's style, if it is one of the user's.
	void DeleteStyleCommand() {
		string name = CaretStyleName();
		if (name == "" || Styles.Find(name) is not { } style) { SetStatus(T("status.deleteStyleCaret", "Put the caret in a paragraph with the style to delete.")); return; }
		if (StyleSets.Get("Classic").Find(name) != null) { SetStatus(T("status.usualStyle", "{0} is one of the usual styles, which stay (Modify Style changes it).", name)); return; }
		if (MessageBox.Show(this, T("msg.deleteStyle", "Delete the {0} style? Its paragraphs become Normal.", name), AppName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
		api.DeleteStyle(name);
		Record(M("DeleteStyle", name));
	}

	// A style's settings as Rte.ModifyStyle takes them ("Font=Georgia\nSize=12 ...").
	internal static string StyleSettings(ParagraphStyle style) => string.Join("\n",
		typeof(ParagraphStyle).GetProperties().Where(p => p.Name != "Name").Select(p => p.Name + "=" + Convert.ToString(p.GetValue(style), CultureInfo.InvariantCulture)));

	// ---- Several windows (iteration 31)

	/// What New Window runs: AutoHotkey with AhkWord's own script (the one this window was started with, else
	/// AhkWord.ahk beside the program), and the document to open (null: a new one).
	internal System.Diagnostics.ProcessStartInfo NewWindowStart(string? file) {
		if (AppFolders.ProgramExe is { } exe) {    // (the installed AhkWord.exe: the exe again, with the same arguments)
			var again = new System.Diagnostics.ProcessStartInfo(exe) { UseShellExecute = false, WorkingDirectory = appDir };
			again.ArgumentList.Add(ahkPath);
			again.ArgumentList.Add(appDir);
			if (file != null) again.ArgumentList.Add(Path.GetFullPath(file));
			again.Environment["RTE_WRITER_CASCADE"] = "1";
			again.Environment.Remove("RTE_WRITER_HWND");
			return again;
		}
		if (ahkPath == "" || !File.Exists(ahkPath)) throw new InvalidOperationException("AutoHotkey was not found; start AhkWord with its AhkWord.ahk.");
		string script = AppFolders.HostScript is { } host && File.Exists(host) ? host
			: File.Exists(Path.Combine(appDir, "AhkWord.ahk")) ? Path.Combine(appDir, "AhkWord.ahk")
			: throw new FileNotFoundException("AhkWord's script (AhkWord.ahk) was not found in " + appDir + ".");
		var start = new System.Diagnostics.ProcessStartInfo(ahkPath) { UseShellExecute = false, WorkingDirectory = appDir };
		start.ArgumentList.Add(script);
		if (file != null) start.ArgumentList.Add(Path.GetFullPath(file));
		start.Environment["RTE_WRITER_CASCADE"] = "1";
		start.Environment.Remove("RTE_WRITER_HWND");    // (a window of its own, not this one's macros' editor)
		return start;
	}

	/// Starts another AhkWord window, with a new document or the file given.
	internal void StartWindow(string? file) {
		using var process = System.Diagnostics.Process.Start(NewWindowStart(file));
	}

	// File > Open in New Window
	void OpenInNewWindowDialog() {
		using var dlg = new OpenFileDialog { Title = T("file.openInNewWindow", "Open in New Window"), Filter = Filter(OpenTypes), InitialDirectory = StartFolder };
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		StartWindow(dlg.FileName);
		SetStatus(T("status.openingInNewWindow", "Opening {0} in a new window...", Path.GetFileName(dlg.FileName)));
	}

	/// A window's place a step down and to the right of bounds, back at the top left of the working area once it would
	/// run off it.
	internal static Rectangle Cascaded(Rectangle bounds, Rectangle area, int step) {
		var moved = new Rectangle(bounds.X + step, bounds.Y + step, bounds.Width, bounds.Height);
		if (moved.Right > area.Right || moved.Bottom > area.Bottom) moved.Location = new Point(area.Left + step, area.Top + step);
		return moved;
	}

	// File > Close: the document goes (its changes saved, if wanted), and an empty one takes its place.
	void CloseDocumentCommand() {
		string name = DocumentName;
		if (!ConfirmDiscard()) return;
		api.Close();
		SetStatus(T("status.closed", "Closed {0}.", name));
	}

	bool ConfirmDiscard() {
		if (!Editor.Modified) return true;
		var answer = MessageBox.Show(this, T("msg.saveChanges", "Save changes to {0}?", DocumentName), AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
		return answer == DialogResult.No || (answer == DialogResult.Yes && SaveCommand());
	}

	string DocumentName => FileName != null ? Path.GetFileName(FileName) : T("ui.untitled", "Untitled");

	// File types for the Open, Save As and Export dialogs. Word and OpenDocument have one entry per converter as well as
	// the automatic one (Via ""), so that a particular converter can be chosen.
	sealed record FileType(string Label, string Patterns, string Via = "");

	static readonly FileType[] OpenTypes = {
		new("Documents", "*.html;*.htm;*.md;*.markdown;*.rtf;*.docx;*.doc;*.odt;*.txt"),
		new("HTML", "*.html;*.htm"), new("Markdown", "*.md;*.markdown"), new("Rich Text Format", "*.rtf"),
		new("Word", "*.docx;*.doc"),
		new("Word, read with Docxodus", "*.docx", OfficeConverters.Docxodus), new("Word, built-in converter", "*.docx", OfficeConverters.OpenXml), new("Word, converted by Word", "*.docx;*.doc", OfficeConverters.Word),
		new("Word, converted by LibreOffice", "*.docx;*.doc", OfficeConverters.LibreOffice), new("Word, converted by pandoc", "*.docx", OfficeConverters.Pandoc),
		new("OpenDocument text", "*.odt"),
		new("OpenDocument, converted by LibreOffice", "*.odt", OfficeConverters.LibreOffice), new("OpenDocument, converted by Word", "*.odt", OfficeConverters.Word),
		new("OpenDocument, converted by pandoc", "*.odt", OfficeConverters.Pandoc),
		new("Text", "*.txt"), new("All files", "*.*")
	};

	static readonly FileType[] WordTypes = {
		new("Word", "*.docx"), new("Word, built-in converter", "*.docx", OfficeConverters.OpenXml),
		new("Word, converted by Word", "*.docx", OfficeConverters.Word), new("Word, converted by LibreOffice", "*.docx", OfficeConverters.LibreOffice),
		new("Word, converted by pandoc", "*.docx", OfficeConverters.Pandoc)
	};

	static readonly FileType[] OdtTypes = {
		new("OpenDocument text", "*.odt"), new("OpenDocument, converted by LibreOffice", "*.odt", OfficeConverters.LibreOffice),
		new("OpenDocument, converted by Word", "*.odt", OfficeConverters.Word), new("OpenDocument, converted by pandoc", "*.odt", OfficeConverters.Pandoc)
	};

	static readonly FileType[] SaveTypes = new[] { new FileType("HTML, pictures in a folder", "*.html"), new FileType("HTML, pictures inside (one file)", "*.html", EditorApi.PicturesInside),
			new FileType("Markdown", "*.md"), new FileType("Rich Text Format", "*.rtf") }
		.Concat(WordTypes).Concat(OdtTypes).Append(new FileType("Text", "*.txt")).ToArray();

	// (a type's name in the interface's language)
	static string TypeLabel(string label) => label switch {
		"Documents" => T("file.documents", "Documents"), "All files" => T("file.all", "All files"), "Text" => T("file.text", "Text"),
		"HTML, pictures in a folder" => T("file.htmlFolder", "HTML, pictures in a folder"), "HTML, pictures inside (one file)" => T("file.htmlInside", "HTML, pictures inside (one file)"),
		"OpenDocument text" => T("file.odt", "OpenDocument text"), "Word, built-in converter" => T("file.wordBuiltIn", "Word, built-in converter"), "Word, read with Docxodus" => T("file.wordDocxodus", "Word, read with Docxodus"),
		"Word, converted by Word" => T("file.wordByWord", "Word, converted by Word"), "Word, converted by LibreOffice" => T("file.wordByLibreOffice", "Word, converted by LibreOffice"),
		"Word, converted by pandoc" => T("file.wordByPandoc", "Word, converted by pandoc"), "OpenDocument, converted by LibreOffice" => T("file.odtByLibreOffice", "OpenDocument, converted by LibreOffice"),
		"OpenDocument, converted by Word" => T("file.odtByWord", "OpenDocument, converted by Word"), "OpenDocument, converted by pandoc" => T("file.odtByPandoc", "OpenDocument, converted by pandoc"),
		"Rich Text Format" => T("file.rtf", "Rich Text Format"), _ => label,
	};
	static string Filter(IEnumerable<FileType> types) => string.Join("|", types.Select(t => $"{TypeLabel(t.Label)} ({t.Patterns.Replace(";*.htm;", ";").Replace(";*.markdown", "")})|{t.Patterns}"));

	// The dialog's FilterIndex (from 1) for a file name and converter: the first type with its extension and converter.
	static int FilterIndex(FileType[] types, string? fileName, string via) {
		string extension = Path.GetExtension(fileName ?? "").ToLowerInvariant();
		int index = Array.FindIndex(types, t => t.Patterns.Split(';').Contains("*" + extension) && t.Via == via);
		if (index < 0) index = Array.FindIndex(types, t => t.Patterns.Split(';').Contains("*" + extension));
		return Math.Max(0, index) + 1;
	}

	static string ViaOf(FileType[] types, int filterIndex) => filterIndex >= 1 && filterIndex <= types.Length ? types[filterIndex - 1].Via : "";

	void OpenDialog() {
		if (!ConfirmDiscard()) return;
		using var dlg = new OpenFileDialog { Title = T("file.open", "Open"), Filter = Filter(OpenTypes), InitialDirectory = StartFolder };
		if (dlg.ShowDialog(this) == DialogResult.OK) {
			Cursor = Cursors.WaitCursor;
			try { api.Open(dlg.FileName, ViaOf(OpenTypes, dlg.FilterIndex)); }
			finally { Cursor = Cursors.Default; }
			SetStatus(T("status.opened", "Opened {0}", dlg.FileName) + (api.LastConverter != "" ? " " + T("status.convertedBy", "(converted by {0})", api.LastConverter) : "") + api.OpenNotice());
		}
	}

	bool SaveCommand() {
		if (FileName == null) return SaveAsDialog();
		try {
			api.Save();
			SetStatus(T("status.saved", "Saved {0}", FileName));
			return true;
		} catch (Exception ex) {
			MessageBox.Show(this, ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		}
	}

	bool SaveAsDialog() {
		using var dlg = new SaveFileDialog {
			Title = T("file.saveAs", "Save As"), Filter = Filter(SaveTypes), FilterIndex = FilterIndex(SaveTypes, FileName ?? "Untitled.html", api.FileConverter),
			FileName = FileName != null ? Path.GetFileName(FileName) : "Untitled.html",
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return false;
		Cursor = Cursors.WaitCursor;
		try {
			api.SaveAs(dlg.FileName, ViaOf(SaveTypes, dlg.FilterIndex));
			SetStatus(T("status.saved", "Saved {0}", dlg.FileName) + (api.LastConverter != "" ? " " + T("status.convertedBy", "(converted by {0})", api.LastConverter) : ""));
			return true;
		} catch (Exception ex) {
			MessageBox.Show(this, ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		} finally {
			Cursor = Cursors.Default;
		}
	}

	string StartFolder => Settings.LastFolder is { } folder && Directory.Exists(folder) ? folder : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

	// ---------------------------------------------------------------- Iteration 3: settings and recent files

	void ApplyWindowSettings() {
		if (Settings.Window is { Length: 4 } w && w[2] >= 400 && w[3] >= 300) {
			var bounds = new Rectangle(w[0], w[1], w[2], w[3]);
			// (started by New Window: not on top of the window that started it)
			if (Environment.GetEnvironmentVariable("RTE_WRITER_CASCADE") == "1")
				bounds = Cascaded(bounds, Screen.FromRectangle(bounds).WorkingArea, SystemInformation.CaptionHeight + LogicalToDeviceUnits(8));
			if (Screen.AllScreens.Any(screen => screen.WorkingArea.IntersectsWith(bounds))) {
				StartPosition = FormStartPosition.Manual;
				Bounds = bounds;
			}
		}
		if (Settings.Maximized) WindowState = FormWindowState.Maximized;
	}

	void SaveSettings() {
		var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
		Settings.Window = new[] { bounds.X, bounds.Y, bounds.Width, bounds.Height };
		Settings.Maximized = WindowState == FormWindowState.Maximized;
		Settings.WordWrap = Editor.WordWrap;
		Settings.SpellCheck = Editor.EnableSpellCheck;
		Settings.AutoLevelStyles = AutoLevelStyles;
		Settings.ShowOutline = !outlineSplit.Panel1Collapsed;
		Settings.SourceView = view.ToString();
		Settings.SideBySide = sideBySide;
		Settings.SourceWordWrap = sourceBox.WordWrap;
		Settings.SourceLineNumbers = sourceBox.LineNumbers;
		if (!outlineSplit.Panel1Collapsed) Settings.OutlineWidth = (int)Math.Round(outlineSplit.SplitterDistance * 96.0 / DeviceDpi);
		try { Settings.Zoom = Editor.ZoomFactor; } catch (ObjectDisposedException) { }
		Settings.Save();
	}

	// Called whenever a document is opened or saved.
	internal void AddRecent(string path) {
		Settings.RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
		Settings.RecentFiles.Insert(0, path);
		if (Settings.RecentFiles.Count > Settings.MaxRecentFiles) Settings.RecentFiles.RemoveRange(Settings.MaxRecentFiles, Settings.RecentFiles.Count - Settings.MaxRecentFiles);
		Settings.LastFolder = Path.GetDirectoryName(path);
		Settings.Save();
		UpdateRecentItems();
	}


	void OpenRecent(string path) {
		if (!File.Exists(path)) {
			Settings.RecentFiles.Remove(path);
			Settings.Save();
			throw new FileNotFoundException($"{path} no longer exists, so it was removed from the list.");
		}
		if (!ConfirmDiscard()) return;
		api.Open(path);
		SetStatus(T("status.opened", "Opened {0}", path) + api.OpenNotice());
	}

	// ---------------------------------------------------------------- Iteration 3: drag and drop

	static readonly string[] DocumentExtensions = { ".html", ".htm", ".md", ".markdown", ".rtf", ".docx", ".doc", ".odt", ".txt" };
	static readonly string[] ImageExtensions = { ".png", ".jpg", ".jpeg", ".gif" };
	static bool IsDocument(string path) => DocumentExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());
	static bool IsImage(string path) => ImageExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());
	static bool CanDrop(string[] files) => files.Any(f => IsDocument(f) || IsImage(f));

	internal string DropTargetStatus { get; private set; } = "not registered";

	void RegisterDropTarget() {
		if (!Editor.IsHandleCreated) return;
		dropTarget = FileDropTarget.Register(Editor.Handle, CanDrop, DropFiles, out string status);
		DropTargetStatus = status;
	}

	// A dropped document is opened; dropped images are inserted at the drop point (screen coordinates), or at the caret.
	// This runs after the drop has completed, so a "save changes?" question does not hold up the program dragged from.
	internal void DropFiles(string[] files, Point? screenPoint) {
		BeginInvoke(() => Run(() => {
			if (files.FirstOrDefault(IsDocument) is { } document) {
				if (!ConfirmDiscard()) return;
				if (StartShown) ShowStart(false);    // (dropped on the start page: the document shown, as when opened from it)
				api.Open(document);
				SetStatus(T("status.opened", "Opened {0}", document) + api.OpenNotice());
				return;
			}
			if (screenPoint is { } point) Editor.Select(Editor.GetCharIndexFromPosition(Editor.PointToClient(point)), 0);
			int count = 0;
			foreach (var image in files.Where(IsImage)) {
				api.InsertImageFile(image, Path.GetFileNameWithoutExtension(image));
				count++;
			}
			SetStatus(TN("status.insertedImages", "Inserted {0} image.", "Inserted {0} images.", count));
		}));
	}

	// ---------------------------------------------------------------- Iteration 3: printing and PDF

	internal DocumentPrinter CreatePrinter(string? printerName = null, bool dialog = false, bool paperOnly = false) {
		// (laid out and printed as the document is: the page view's gaps out while the printer is in use; not for its paper
		// alone)
		bool gaps = !paperOnly && pageGaps.Applied && !gapsOut && !quietGaps;
		// (iteration 55: the window's painting held until the gaps are back, then painted once: each recovery copy and save
		// showed the text without its gaps for a moment, a flicker and a jump every minute or so. Not while a dialog shows
		// or a PDF is made (dialog): the window is painted meanwhile.)
		var held = gaps && !dialog ? HoldGaps() : null;
		if (gaps) { QuietGaps(pageGaps.Remove); gapsOut = true; }
		try {
			var printer = new DocumentPrinter(Editor, DocumentName, FileName ?? "", Settings, Bands, Numbering, printerName, PageMargins, PageLandscape, DocumentPaper ?? Settings.Paper) {    // (no paper of its own: the printer's)
				NoteStyle = NoteStyle, PageSections = Sections.Resolve(Editor.TextDocument, Bands, DocumentColumns)    // (iteration 63)
			};
			if (gaps) printer.Disposed += () => { try { gapsOut = false; QuietGaps(pageGaps.Reapply); } finally { held?.Dispose(); } };
			return printer;
		} catch {
			if (gaps) { gapsOut = false; QuietGaps(pageGaps.Reapply); }
			held?.Dispose();
			throw;
		}
	}

	/// The document's header and footer.
	internal PageBands Bands = new();

	// ---- the header and footer shown above and below the text (iteration 14)

	readonly BandStrip headerStrip = new(true), footerStrip = new(false);
	readonly System.Windows.Forms.Timer bandTimer = new() { Interval = 500 };
	List<int>? bandStarts;          // where the printed pages start, for {page} (the page in view) and {pages}; null: not known
	// The text changed since it was counted. (typedOnly: only by typing, so that the pages may well be as they were: UpdateBands)
	bool bandPagesStale { get => pagesStale; set { pagesStale = value; typedOnly = false; } }
	bool pagesStale = true, typedOnly;
	DateTime lastEdit = DateTime.MinValue;
	string bandsShown = "";

	// Shows the strips as the document's header and footer are now (twice a second: cheap unless something changed). The
	// pages are laid out again (with the printer's layout) once typing has paused for a second; {page} is the page in
	// view (the one in the middle of the window), so it changes as the document is scrolled.
	void UpdateBands() {
		bool visible = Settings.ShowBands;
		var bands = Bands;
		static bool Numbered(string[] parts) => PageBands.Parts(parts).Any(p => p.Contains("{page", StringComparison.OrdinalIgnoreCase));
		bool needPages = (visible && ((bands.ShowHeader && Numbered(bands.Header)) || (bands.ShowFooter && Numbered(bands.Footer))
				|| (bands.DifferentFirst && (bands.ShowHeader || bands.ShowFooter))))    // (iteration 49: which page is in view)
			|| (Settings.PrintLayout && !PagesShown);    // (and Print Layout's page boundaries)
		bool pageView = PageViewActive;
		if (!pageView && pageGaps.Applied && !gapsOut) { QuietGaps(pageGaps.Remove); pageGaps.Forget(); }
		// (the page view showing without its gaps, as back from a source view, which took them out: worked out again)
		// (not when they were worked out and there are none, as with one page that starts with a table: that would be again
		// at every tick, the window repainted each time)
		if (pageView && !pageGaps.Applied && !gapsOut && !quietGaps && !noGapsNeeded) bandPagesStale = true;
		if (pageGaps.Lost) { pageGaps.Lost = false; bandPagesStale = true; }    // (gaps editing took apart: worked out again)
		// (iteration 55) a header or footer that now reaches further or less far: the sheets laid out again at once, not
		// after the next edit (meanwhile the text was under the header)
		bool marginsMoved = pageView && pageGaps.Applied && !gapsOut && !quietGaps && pageLayout is { } laidOut && GapMargins(laidOut) != gapsMargins;
		if (marginsMoved) bandPagesStale = true;
		if ((needPages || marginsMoved) && bandPagesStale && !gapsOut && (marginsMoved || (DateTime.Now - lastEdit).TotalMilliseconds > 1000)) {
			bool samePages = typedOnly && pageView && !marginsMoved && pagesStamp != null && pagesStamp == PagesStamp();
			bandPagesStale = false;
			if (samePages) bandStarts = pageRanges.Select(p => p.Start).ToList();    // (where typing has moved them)
			else if (pageView) RecomputeGaps();
			else {
				try { using var printer = CreatePrinter(); bandStarts = printer.PageStarts(); }
				catch (Exception) { bandStarts = null; }
			}
		}
		Editor.PageStarts = Settings.PrintLayout ? bandStarts : null;    // (until they are worked out again: the ones before)
		// (iteration 46) the line numbers drawn somewhere else than their lines are now: painted again
		if (lineGutter.Visible && !quietGaps && !gapsOut && !lineGutter.DrawnWhereTheyBelong()) lineGutter.Invalidate();
		int viewPage = 1;
		if (needPages && bandStarts is { Count: > 0 } starts && Editor.IsHandleCreated) {
			int middle = Editor.GetCharIndexFromPosition(new Point(Editor.ClientSize.Width / 2, Editor.ClientSize.Height / 2));
			viewPage = Math.Max(1, starts.FindLastIndex(s => s <= middle) + 1);
		}
		if (lineGutter.Visible && lineGutter.NeedsPageStarts && pageStartsStale && (DateTime.Now - lastEdit).TotalMilliseconds > 1000) {
			pageStartsStale = false;
			try { using var printer = CreatePrinter(); pageStarts = printer.PageStarts(); }
			catch (Exception) { pageStarts = null; }
			lineGutter.PageStarts = pageStarts;
		}
		var (pageBands, on, shownPage, sectionPages) = PageOf(viewPage);    // (iteration 63: the page's section's)
		bands = pageBands;
		string Fill(string text) => Regex.Replace(text, @"\{(title|file|page|pages|sectionpages|date|time)\}", m => m.Groups[1].Value.ToLowerInvariant() switch {
			"title" => DocumentName,
			"file" => FileName ?? DocumentName,
			"page" => shownPage.ToString(CultureInfo.CurrentCulture),
			"pages" => bandStarts?.Count.ToString(CultureInfo.CurrentCulture) ?? "?",
			"sectionpages" => sectionPages.ToString(CultureInfo.CurrentCulture),
			"date" => DateTime.Now.ToShortDateString(),
			_ => DateTime.Now.ToShortTimeString()
		}, RegexOptions.IgnoreCase);
		var header = PageBands.Parts(bands.HeaderOn(on)).Select(Fill).ToArray();
		var footer = PageBands.Parts(bands.FooterOn(on)).Select(Fill).ToArray();
		bool headerOn = bands.HeaderOn(on) != null, footerOn = bands.FooterOn(on) != null;
		string shown = $"{visible}|{pageView}|{headerOn}|{string.Join("\t", header)}|{footerOn}|{string.Join("\t", footer)}|{bands.HeaderFontOn(on)}|{bands.FooterFontOn(on)}|{bands.HeaderLine}|{bands.FooterLine}";
		if (shown == bandsShown) return;
		bandsShown = shown;
		// (in page view each sheet has its own: not the strips)
		headerStrip.Display(visible && headerOn && !pageView, header, bands.HeaderFontOn(on));
		footerStrip.Display(visible && footerOn && !pageView, footer, bands.FooterFontOn(on));
		if (pageView) Editor.Invalidate();
	}    // (the default, once the settings are loaded: in the constructor)

	// ---- Page view (iteration 35): the printed pages as sheets of their own (rte_pagegaps.cs)

	PageGaps pageGaps = null!;
	bool quietGaps, gapsOut;

	/// Whether the page view shows now: Print Layout, Show Pages, the text being edited (not Multiple Pages or a source view).
	internal bool PageViewActive => Settings.PrintLayout && Settings.PageView && pageLayout != null && !PagesShown && !SourceActive && Editor.IsHandleCreated;

	internal PageGaps Gaps => pageGaps;

	const int EM_GETSCROLLPOS = 0x04DD, EM_SETSCROLLPOS = 0x04DE, EM_SETEVENTMASK_GAPS = 0x0445;
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendPoint(IntPtr hWnd, int msg, IntPtr wParam, ref Point lParam);
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendMask(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	// Puts the gaps in or takes them out: without undo steps or change notifications, the document's modified flag as
	// it was, the line with the caret where it was on the screen (or, the caret out of view, the text at the top of the
	// window), the window not drawing until it is all done (iteration 43: no in-between state shown).
	void QuietGaps(Action change) {
		if (!Editor.IsHandleCreated) { pageGaps.Forget(); return; }    // (the window closing: the gaps go with its text)
		if (quietGaps) { change(); return; }
		quietGaps = true;
		var doc = Editor.TextDocument;
		bool modified = Editor.Modified;
		int top = Editor.GetCharIndexFromPosition(new Point(1, 1));
		int topY = Editor.GetPositionFromCharIndex(top).Y;
		// (the caret's line's bottom, the next line's top: a gap before the caret's line or a raise of it leaves its text
		// there, where its top moves)
		var caretLine = doc.Range2(doc.Selection2.Start, doc.Selection2.Start);
		caretLine.Expand(5);    // tomLine
		if (caretLine.End < doc.Range2(0, 0).StoryLength - 1 && Editor.GetPositionFromCharIndex(caretLine.End).Y is int bottom
			&& bottom > 0 && bottom <= Editor.ClientSize.Height)
			(top, topY) = (caretLine.End, bottom);
		// (the editor's painting held until the gaps are in: a paint asked for meanwhile is left to be done after, not done
		// in between. Iteration 48: LockWindowUpdate did this from iteration 43, but a paint while locked drew nothing and
		// took the window's need of one with it, so the window kept what it showed before (a table row where it had been,
		// after opening). WM_SETREDRAW off gave positions as if scrolled to the top.)
		bool hold = Editor.Visible && !heldGaps;    // (held already: painted once, after it all: HeldGaps)
		if (hold) Editor.HoldPaint = true;
		doc.Undo(-9999995);    // tomSuspend
		IntPtr mask = SendMask(Editor.Handle, EM_SETEVENTMASK_GAPS, IntPtr.Zero, IntPtr.Zero);
		doc.Freeze();
		try {
			change();
		} finally {
			doc.Unfreeze();
			SendMask(Editor.Handle, EM_SETEVENTMASK_GAPS, IntPtr.Zero, mask);
			doc.Undo(-9999994);    // tomResume
			Editor.Modified = modified;
			// scrolled back by as much as it moved (iteration 43: EM_SETSCROLLPOS takes the window's pixels, not unzoomed ones,
			// as was taken before: zoomed, the text was left a line or so off, which showed as a jump); what a step moved the
			// text by is measured, and any rest scrolled with that
			double perStep = 1;
			for (int tries = 0; tries < 4 && !openingSheets; tries++) {
				int moved = Editor.GetPositionFromCharIndex(top).Y - topY;
				int step = (int)Math.Round(moved / perStep);
				if (step == 0) break;
				var scroll = new Point();
				SendPoint(Editor.Handle, EM_GETSCROLLPOS, IntPtr.Zero, ref scroll);
				scroll.Y += step;
				SendPoint(Editor.Handle, EM_SETSCROLLPOS, IntPtr.Zero, ref scroll);
				int done = moved - (Editor.GetPositionFromCharIndex(top).Y - topY);
				if (done == 0) break;    // (at the top or the end: as far as it goes)
				perStep = done / (double)step;
			}
			quietGaps = false;
			if (hold) { Editor.HoldPaint = false; RepaintAfterGaps(); }
		}
	}

	// Runs body, which takes the gaps out and puts them back (or works them out again), with the editor's painting held
	// until it is all done, then paints once. Iteration 51: each QuietGaps painted when it was done,
	// so the window showed the text without its gaps, scrolled elsewhere, for a moment in between (a flash at Ctrl+C,
	// and every few minutes at the recovery copy, which reads the document without them).
	void HeldGaps(Action body) {
		using (HoldGaps()) body();
	}

	// HeldGaps's hold, until disposed.
	IDisposable HoldGaps() {
		if (heldGaps || !Editor.IsHandleCreated) return new OnDispose(() => { });
		heldGaps = true;
		bool hold = Editor.Visible;
		if (hold) Editor.HoldPaint = true;
		// (iteration 52: not frozen as well, TOM's Freeze: frozen, RichEdit gives where characters were before, and
		// QuietGaps' scrolling back, which measures where the text went, then did nothing)
		return new OnDispose(() => {
			heldGaps = false;
			if (hold && Editor.IsHandleCreated) { Editor.HoldPaint = false; RepaintAfterGaps(); }
		});
	}
	bool heldGaps;

	/// Iteration 68: while a document opens: the busy cursor and a word in the status bar, the editor's painting held
	/// (HoldGaps), painted once after; the keyboard's focus off the editor (with it there, Windows' text services follow each
	/// of the many changes a file's formats are put in with). OpenTime: how long it took.
	internal IDisposable HoldOpening(string name) {
		var clock = System.Diagnostics.Stopwatch.StartNew();
		bool focused = Editor.Focused;
		if (focused) ActiveControl = null;
		UseWaitCursor = true;
		Cursor.Current = Cursors.WaitCursor;
		ShowProgress($"Opening {name}...");
		var held = HoldGaps();
		return new OnDispose(() => {
			held.Dispose();
			UseWaitCursor = false;
			Cursor.Current = Cursors.Default;
			if (focused && Editor.CanFocus) Editor.Focus();
			OpenTime = clock.Elapsed;
		});
	}
	internal TimeSpan OpenTime;

	/// The new document's sheets worked out now, not a moment after it shows (when they would move the text), and its
	/// start in view with page 1's top margin.
	internal void SheetsNow() {
		if (!PageViewActive || gapsOut) return;
		measuredPages = null;
		ReferencesNow();    // (they can change the pages: after it shows, the sheets would be worked out again, and move it)
		var measured = measuredPages;    // (the pages the references laid out, nothing changed since)
		measuredPages = null;
		openingSheets = true;
		try { RecomputeGaps(measured); }
		finally { openingSheets = false; }
		bandPagesStale = false;
	}
	/// Runs change with no scrolling back after the gaps: the document is seen from its top, its top margin in view.
	internal void FromTheTop(Action change) {
		openingSheets = true;
		try { change(); }
		finally { openingSheets = false; }
	}
	bool openingSheets;    // (nothing to scroll back to: the document starts at its top, and scrolling draws at once, held or not)

	// The window painted after the gaps went in or out.
	void RepaintAfterGaps() {
		// (painted now, all of it: only asked for, the paint was sometimes not done, the window keeping what it showed
		// before the gaps went in)
		Editor.Invalidate();
		Editor.Update();
		pagePanel.Invalidate();
		// (iteration 46: the line numbers too, where they are now that the text is back in its place; drawn while the
		// gaps went in, they were left where the text had been moved to before it)
		lineGutter.Invalidate();
		// (iteration 49) and once more when the messages waiting are done: RichEdit sometimes lays the text out again
		// only then, and the paint just now drew it as it was before (after a recovery copy, whose RTF is read without
		// the gaps: the text without page 1's top margin, over the header drawn where it belongs), until scrolled
		if (!repaintPending && IsHandleCreated) {
			repaintPending = true;
			BeginInvoke(() => {
				repaintPending = false;
				if (quietGaps || heldGaps || !Editor.IsHandleCreated) return;
				Editor.Invalidate();
				Editor.Update();
				lineGutter.Invalidate();
			});
		}
	}

	bool repaintPending;

	// Rte.PageViewReport: positions and sizes only.
	internal string PageViewReport() {
		var shownNumbers = Numbering.On ? lineGutter.ShownNumbers() : new();    // (first: see below)
		var drawnNumbers = lineGutter.LastDrawn;
		var r = new StringBuilder();
		var doc = Editor.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var paragraphs = StyleEngine.Paragraphs(doc, 0, storyEnd).Select(p => p.Start).ToList();
		int ParagraphOf(int at) => paragraphs.FindLastIndex(s => s <= at) + 1;
		r.AppendLine($"Program: {AppFolders.HostScript ?? "(not known)"}");
		r.AppendLine($"Print Layout {Settings.PrintLayout}, Show Pages {Settings.PageView} (active {PageViewActive}), zoom {Editor.ZoomFactor:0.###} ({Settings.ZoomMode}), {Editor.PixelsPerPoint:0.###} px/pt, window {Editor.ClientSize.Width}x{Editor.ClientSize.Height}");
		r.AppendLine($"Characters {storyEnd}, paragraphs {paragraphs.Count}, layout stale {bandPagesStale}, gaps out {gapsOut}");
		// (iteration 46) the line numbers in view: where each was drawn last, where it belongs now (worked out before
		// anything else here, which can take the gaps out and put them back, and so paint everything again), against
		// its line's top
		if (Numbering.On) {
			r.AppendLine($"Line numbers {Numbering.Describe()}; screen {DeviceDpi} dpi; numbers' column {lineGutter.Bounds}, editor {Editor.Bounds}, text margins {Editor.TextMargins}; " +
				$"painted {(DateTime.Now - lineGutter.LastPainted).TotalMilliseconds:0} ms ago, the editor {(DateTime.Now - editorPainted).TotalMilliseconds:0} ms ago; page starts for restarts {lineGutter.PageStarts?.Count.ToString() ?? "none"}; " +
				$"drawn where they belong: {drawnNumbers.Count == shownNumbers.Count && drawnNumbers.Zip(shownNumbers).All(p => p.First.Number == p.Second.Number && p.First.Top == p.Second.Top)}");
			for (int i = 0; i < Math.Max(shownNumbers.Count, drawnNumbers.Count); i++) {
				var shown = i < shownNumbers.Count ? shownNumbers[i] : null;
				var drawn = i < drawnNumbers.Count ? drawnNumbers[i] : null;
				string where = shown == null ? "" : $"line {shown.Line} at {shown.Start} (p{ParagraphOf(shown.Start)}), line top {Editor.GetPositionFromCharIndex(shown.Start).Y}, text's baseline {shown.Baseline + lineGutter.Top - Editor.Top}, belongs at top {shown.Top + lineGutter.Top - Editor.Top}, space before {doc.Range2(shown.Start, shown.Start).Para.SpaceBefore:0.#} pt";
				r.AppendLine($"  number {shown?.Number ?? drawn!.Number}: {where}; drawn at top {(drawn == null ? "-" : (drawn.Top + lineGutter.Top - Editor.Top).ToString())}{(drawn != null && shown != null && drawn.Number != shown.Number ? $" (as number {drawn.Number})" : "")}");
			}
		}		try {
			using var printer = CreatePrinter();
			r.AppendLine($"Printer {printer.Document.PrinterSettings.PrinterName}; printed pages start at {string.Join(" ", printer.PageHeights().Select(p => $"{p.Start} (p{ParagraphOf(p.Start)}, {p.Height:0} pt)"))}");
		} catch (Exception ex) { r.AppendLine("Printed pages: " + ex.Message); }
		r.AppendLine($"Gaps {pageGaps.Gaps.Count} for {pageGaps.Pages} pages ({pageGaps.Gaps.Count(g => g.Paragraph)} as space before a paragraph), page 1's top {pageGaps.TopExtra:0.#} pt");
		int top = Editor.GetCharIndexFromPosition(new Point(1, 1));
		foreach (var gap in pageGaps.Gaps) {
			var para = doc.Range2(gap.LineStart, gap.LineStart).Para;
			var (footer, band, header) = pageGaps.Areas(gap, Editor.ClientSize.Width, Editor.PixelsPerPoint);
			r.AppendLine($"  page {gap.Page}: p{ParagraphOf(gap.LineStart)} {(gap.Paragraph ? $"space before (its own {gap.Original:0.#} pt, now {gap.Applied:0.#})" : $"line {gap.LineStart}-{gap.LineEnd} raised {gap.Raise:0.##} pt")}, next {gap.Next}, height {gap.Height:0.#} pt, spacing rule {para.LineSpacingRule} {para.LineSpacing:0.##}; " +
				$"line at y {Editor.GetPositionFromCharIndex(gap.LineStart).Y}, next at y {(gap.Next >= 0 ? Editor.GetPositionFromCharIndex(gap.Next).Y : -1)}, band {band.Top:0}-{band.Bottom:0}");
		}
		var run = doc.Range2(0, 0);
		int untracked = 0;
		while (run.Start < storyEnd) {
			run.MoveEnd(13, 1);    // tomCharFormat
			if (run.End <= run.Start) break;
			float position = run.Font.Position;
			if (Math.Abs(position) > 0.01f && !pageGaps.Gaps.Any(g => g.LineStart <= run.Start && run.Start < g.LineEnd)) {
				untracked++;
				r.AppendLine($"  raised but not a gap: {run.Start}-{run.End} (p{ParagraphOf(run.Start)}) by {position:0.##} pt");
			}
			if (run.End >= storyEnd) break;
			run.SetRange(run.End, run.End);
		}
		r.AppendLine($"Raised text that isn't a gap's: {untracked}. In view from character {top} (p{ParagraphOf(top)}).");
		return r.ToString();
	}

	// (iteration 55) the sheets' top and bottom margins (points), as printed: a header or footer reaching past its
	// margin pushes the text in; and those the gaps were put in with
	(float Top, float Bottom, float FirstTop, float FirstBottom) GapMargins(PageLayout layout) => (Math.Max(layout.Top / 20f, Bands.HeaderReach() * 0.72f), Math.Max(layout.Bottom / 20f, Bands.FooterReach() * 0.72f),
		Math.Max(layout.Top / 20f, Bands.HeaderReach(1) * 0.72f), Math.Max(layout.Bottom / 20f, Bands.FooterReach(1) * 0.72f));
	(float Top, float Bottom, float FirstTop, float FirstBottom) gapsMargins;

	// The pages laid out again (as printed: without gaps) and their gaps put in. (Rte.PageViewReport and tests: how many
	// times.)
	internal int GapsRecomputed;
	void RecomputeGaps((List<(int Start, float Height)> Pages, List<List<Footnotes.NoteLine>> Notes, (string Name, float Size) NoteFont)? measured = null) {
		GapsRecomputed++;
		// (not now: the gaps are out for something else, or the page isn't laid out yet: a moment later, as after typing,
		// rather than a page view left without its gaps)
		if (pageLayout is not { } layout || gapsOut) { bandPagesStale = true; return; }
		QuietGaps(() => {
			pageGaps.Remove();
			List<(int Start, float Height)>? pages = null;
			if (measured is { } m) (pages, sheetNotes, sheetNoteFont) = (m.Pages, m.Notes, m.NoteFont);
			else try {
				using var printer = CreatePrinter();
				pages = printer.PageHeights();
				(sheetNotes, sheetNoteFont) = (printer.LastNotes, printer.NoteFontUsed);    // (iteration 51: the sheets' notes, as printed)
			}
			catch (Exception) { bandPagesStale = true; lastEdit = DateTime.Now; }    // (the printer busy: again in a second)
			bandStarts = pages?.Select(p => p.Start).ToList();
			if (pages != null && PageViewActive)
			{
				var (top, bottom, firstTop, firstBottom) = gapsMargins = GapMargins(layout);
				// (iteration 63) each page's own margins: its section's, with its header and footer, the section's first page its own
				Func<int, (float, float, float)>? marginsOf = null;
				var sections = Sections.Resolve(Editor.TextDocument, Bands);
				if (sections.Count > 1) {
					var numbers = Sections.Number(pages.Select(p => p.Start).ToList(), sections);
					marginsOf = k => {
						var n = numbers[Math.Min(k, numbers.Count - 1)];
						static int Hundredths(int twips) => (int)Math.Round(twips / 14.4);
						var (t, bo) = Sections.MarginsOn(sections[n.Section], n.InSection == 1 ? 1 : 2, Hundredths(layout.Top), Hundredths(layout.Bottom), Hundredths(layout.Height));
						return (layout.Height / 20f - t * 0.72f - bo * 0.72f, bo * 0.72f, t * 0.72f);
					};
				}
				pageGaps.Apply(pages, layout.Height / 20f - top - bottom, bottom, top, layout.Height / 20f - firstTop - firstBottom, firstBottom, firstTop, marginsOf);
				noGapsNeeded = !pageGaps.Applied;
			}
		});
		pagesStamp = bandStarts == null ? null : PagesStamp();
		pageRanges = (bandStarts ?? new()).Select(s => Editor.TextDocument.Range2(s, s)).ToList();
	}

	// What says that the pages are as they were worked out, without taking the gaps out to see: as many lines, as tall, and
	// each gap on its line. Worked out only when they aren't: taking the gaps out and putting them back moves the text, and with
	// a table at or below the caret the window showed that, a second after typing stopped. Null: a gap that can't be said of.
	string? PagesStamp() {
		var doc = Editor.TextDocument;
		int end = Math.Max(0, doc.Range2(0, 0).StoryLength - 1);
		var stamp = new StringBuilder($"{SendMask(Editor.Handle, EM_GETLINECOUNT, IntPtr.Zero, IntPtr.Zero)} {pageGaps.PositionOf(end).Y - pageGaps.PositionOf(0).Y}");
		foreach (var gap in pageGaps.Gaps) {
			if (gap.Track is not { } track) return null;
			stamp.Append($" {SendMask(Editor.Handle, EM_EXLINEFROMCHAR, IntPtr.Zero, track.Start)}");
			if (gap.Paragraph) continue;
			var line = doc.Range2(track.Start, track.Start);
			line.Expand(5);    // tomLine
			if (line.Start != track.Start || line.End - track.End is not (0 or 1)) return null;
		}
		return stamp.ToString();
	}
	const int EM_GETLINECOUNT = 0x00BA, EM_EXLINEFROMCHAR = 0x0436;
	string? pagesStamp;
	List<ITextRange2> pageRanges = new();    // (where the pages start: they move with the text)

	bool noGapsNeeded;    // (the page view worked out for this text, with no gaps to put: see UpdateBands)

	/// Runs read with the document as it is, without the page view's gaps, which are then put back as they were (read
	/// mustn't change the text: see WithoutGapsChanging).
	internal T WithoutGaps<T>(Func<T> read) {
		if (gapsOut || quietGaps || gapsLeft || !pageGaps.Applied) return read();
		T result = default!;
		HeldGaps(() => {
			QuietGaps(pageGaps.Remove);
			gapsOut = true;
			try { result = read(); }
			finally { gapsOut = false; QuietGaps(pageGaps.Reapply); }
		});
		return result;
	}

	internal void WithoutGaps(Action read) => WithoutGaps(() => { read(); return 0; });

	/// The RTF of the document as it is without the gaps, read leaving the window alone: they are taken out of a hidden copy
	/// (the recovery copy, once a minute while typing: taken out of the editor, that moved the text). Else as WithoutGaps.
	internal string RtfLeavingGaps() {
		if (gapsOut || quietGaps || !pageGaps.Applied) return WithoutGaps(() => Editor.Rtf ?? "");
		return pageGaps.OutOf(Editor.Rtf ?? "") ?? WithoutGaps(() => Editor.Rtf ?? "");
	}

	/// Runs read, which doesn't depend on the gaps, without WithoutGaps taking them out for it.
	internal T LeavingGaps<T>(Func<T> read) {
		bool was = gapsLeft;
		gapsLeft = true;
		try { return read(); }
		finally { gapsLeft = was; }
	}
	bool gapsLeft;

	/// Runs change with the document as it is, without the page view's gaps, which are worked out again after it.
	internal void WithoutGapsChanging(Action change) {
		if (gapsOut || quietGaps || !pageGaps.Applied) { change(); return; }
		HeldGaps(() => {
			QuietGaps(pageGaps.Remove);
			gapsOut = true;
			try { change(); }
			finally { gapsOut = false; pageGaps.Forget(); RecomputeGaps(); }
		});
	}

	/// For a command that reads or changes paragraph formats: page 1's top margin (space before the first paragraph)
	/// taken out while it runs if the selection is in the first paragraph (or all: a style changed), and the gaps put
	/// back after it (worked out again if it changed something).
	internal IDisposable FirstParagraphAsIs(bool changes, bool all = false) {
		if (gapsOut || quietGaps || !pageGaps.Applied) return new OnDispose(() => { });
		var doc = Editor.TextDocument;
		if (!all && !pageGaps.AnyParagraphGapIn(doc.Selection2.Start, doc.Selection2.End)) return new OnDispose(() => { });
		var held = HoldGaps();    // (painted once, after the gaps are back)
		QuietGaps(pageGaps.Remove);
		gapsOut = true;
		return new OnDispose(() => {
			try {
				gapsOut = false;
				if (changes) { pageGaps.Forget(); RecomputeGaps(); } else QuietGaps(pageGaps.Reapply);
			} finally { held.Dispose(); }
		});
	}

	/// Show Pages changed (or Print Layout): the gaps worked out again at once, or taken out.
	internal void PageViewChanged() {
		bandPagesStale = true;
		lastEdit = DateTime.MinValue;
		if (!PageViewActive && pageGaps.Applied) { QuietGaps(pageGaps.Remove); pageGaps.Forget(); }
		bandsShown = "";
		UpdateBands();
		UpdateUi();
	}

	// The header's or footer's parts for a page, filled in (null: not shown).
	string[]? GapBandParts(bool header, int page) {
		var (bands, on, shown, sectionPages) = PageOf(page);    // (iteration 63: the page's section's)
		if (!Settings.ShowBands || (header ? bands.HeaderOn(on) : bands.FooterOn(on)) == null) return null;    // (iteration 55: the page's own)
		int pages = pageGaps.Pages > 0 ? pageGaps.Pages : bandStarts?.Count ?? 1;
		string Fill(string text) => Regex.Replace(text, @"\{(title|file|page|pages|sectionpages|date|time)\}", m => m.Groups[1].Value.ToLowerInvariant() switch {
			"title" => DocumentName,
			"file" => FileName ?? DocumentName,
			"page" => shown.ToString(CultureInfo.CurrentCulture),
			"pages" => pages.ToString(CultureInfo.CurrentCulture),
			"sectionpages" => sectionPages.ToString(CultureInfo.CurrentCulture),
			"date" => DateTime.Now.ToShortDateString(),
			_ => DateTime.Now.ToShortTimeString()
		}, RegexOptions.IgnoreCase);
		return (header ? bands.HeaderOn(on) : bands.FooterOn(on))!.Select(Fill).ToArray();
	}

	// (iteration 63) a page's (from 1) section's header and footer, its place in the section (1: the section's first),
	// its number as shown and its section's pages; worked out again when the text or the pages change
	(PageBands Bands, int InSection, int Shown, int SectionPages) PageOf(int page) {
		var starts = bandStarts;
		// (the window closing, a timer's last call after the editor is gone, or before it has its handle: the document's
		// own header and footer, without asking the editor)
		if (Editor.IsDisposed || !Editor.IsHandleCreated) return (Bands, page, page, starts?.Count ?? page);
		int length = Editor.TextLength;
		if (pageOfCache == null || !ReferenceEquals(pageOfStarts, starts) || pageOfLength != length || !ReferenceEquals(pageOfBands, Bands)) {
			(pageOfStarts, pageOfLength, pageOfBands) = (starts, length, Bands);
			var sections = Sections.Resolve(Editor.TextDocument, Bands);
			pageOfCache = (sections, starts is { Count: > 0 } ? Sections.Number(starts, sections) : null);
		}
		var (all, numbers) = pageOfCache.Value;
		if (numbers == null || numbers.Count == 0) return (Bands, page, page, starts?.Count ?? page);
		var n = numbers[Math.Clamp(page - 1, 0, numbers.Count - 1)];
		return (all[n.Section].Bands, n.InSection, n.Shown, n.SectionPages);
	}
	(List<PageSection>, List<Sections.PageNumber>?)? pageOfCache;
	List<int>? pageOfStarts;
	int pageOfLength;
	PageBands? pageOfBands;

	// The bands between the sheets where they cross the page's margins (outside the editor, which draws its part).
	void PaintGapBands(Graphics g) {
		var bands = PanelBands();
		if (bands.Count == 0) return;
		using var grey = new SolidBrush(PrintLayoutPage.Surround);
		foreach (var band in bands) g.FillRectangle(grey, band);
	}

	DateTime editorPainted;    // (Rte.PageViewReport)
	readonly PageHost pageHost = new() { Dock = DockStyle.Fill };

	// The sheets' bands in the panel's margins, as FollowGapBands last had them painted (the panel's coordinates).
	List<Rectangle> panelBands = new();

	// The bands across the panel, beside the editor's (in view only).
	List<Rectangle> PanelBands() {
		var bands = new List<Rectangle>();
		if (!pageGaps.Applied || !Editor.IsHandleCreated) return bands;
		float ppp = Editor.PixelsPerPoint;
		foreach (var gap in pageGaps.Gaps) {
			if (gap.Page == 0) continue;    // (page 1's top: no band)
			var (_, band, _) = pageGaps.Areas(gap, Editor.ClientSize.Width, ppp);
			float y = Editor.Top + band.Y;
			if (y + band.Height < 0 || y > pagePanel.Height) continue;
			int top = (int)Math.Floor(Math.Max(Editor.Top, y)), bottom = (int)Math.Ceiling(Math.Min(band.Bottom + Editor.Top, Editor.Bottom));
			if (bottom > top) bands.Add(new Rectangle(0, top, pagePanel.Width, bottom - top));
		}
		// (iteration 46) the grey after the last sheet
		if (Editor.LastSheetEnd() is { } end && Editor.Top + end.Bottom < Editor.Bottom) {
			int top = (int)Math.Floor(Math.Max(Editor.Top, Editor.Top + end.Bottom));
			bands.Add(new Rectangle(0, top, pagePanel.Width, Editor.Bottom - top));
		}
		return bands;
	}

	// Iteration 44: after the editor painted (it scrolled, or the text changed), the margins' bands follow at once, in the
	// same pass: painted later, as Windows gets round to it, they lagged a moment behind the bands in the page, which
	// showed as a flicker while scrolling. Only where a band was or is now, and not at all if none moved.
	void FollowGapBands() {
		var now = PanelBands();
		if (now.SequenceEqual(panelBands)) return;
		foreach (var band in panelBands.Concat(now)) pagePanel.Invalidate(band, false);
		panelBands = now;
		pagePanel.Update();
	}

	/// The header and footer new documents get (File > Header and Footer can make a document's the default).
	internal PageBands DefaultBands() => new() {
		ShowHeader = Settings.ShowHeader, Header = Settings.Parts(Settings.Header), ShowFooter = Settings.ShowFooter, Footer = Settings.Shown(Settings.Parts(Settings.Footer)),
		DifferentFirst = Settings.DifferentFirst, FirstHeader = Settings.Parts(Settings.FirstHeader), FirstFooter = Settings.Shown(Settings.Parts(Settings.FirstFooter)),
		HeaderFont = Settings.HeaderFont, FooterFont = Settings.FooterFont, HeaderDistance = Settings.HeaderDistance, FooterDistance = Settings.FooterDistance,
		HeaderLine = Settings.HeaderLine, FooterLine = Settings.FooterLine,
		ShowFirstHeader = Settings.DifferentFirst ? Settings.ShowFirstHeader : null, ShowFirstFooter = Settings.DifferentFirst ? Settings.ShowFirstFooter : null,
		FirstHeaderFont = Settings.DifferentFirst ? Settings.FirstHeaderFont : null, FirstFooterFont = Settings.DifferentFirst ? Settings.FirstFooterFont : null,
		FirstHeaderLine = Settings.DifferentFirst ? Settings.FirstHeaderLine : null, FirstFooterLine = Settings.DifferentFirst ? Settings.FirstFooterLine : null
	};

	internal void MakeDefault(PageBands bands) {
		(Settings.ShowHeader, Settings.Header, Settings.ShowFooter, Settings.Footer) = (bands.ShowHeader, Settings.Parts(bands.Header), bands.ShowFooter, Settings.Neutral(Settings.Parts(bands.Footer)));
		(Settings.DifferentFirst, Settings.FirstHeader, Settings.FirstFooter) = (bands.DifferentFirst, Settings.Parts(bands.FirstHeader), Settings.Neutral(Settings.Parts(bands.FirstFooter)));
		(Settings.HeaderFont, Settings.FooterFont, Settings.HeaderDistance, Settings.FooterDistance) = (bands.HeaderFont, bands.FooterFont, bands.HeaderDistance, bands.FooterDistance);
		(Settings.HeaderLine, Settings.FooterLine) = (bands.HeaderLine, bands.FooterLine);
		(Settings.ShowFirstHeader, Settings.ShowFirstFooter, Settings.FirstHeaderFont, Settings.FirstFooterFont, Settings.FirstHeaderLine, Settings.FirstFooterLine)
			= (bands.ShowFirstHeader, bands.ShowFirstFooter, bands.FirstHeaderFont, bands.FirstFooterFont, bands.FirstHeaderLine, bands.FirstFooterLine);
		Settings.Save();
	}

	/// (iteration 63) The document's own page setup: the margins (1/100 inch: left, right, top, bottom) and orientation its
	/// file gave it, or Page Setup did (null: none; Page Setup's last, the settings', then).
	internal int[]? DocumentMargins;
	internal bool? DocumentLandscape;
	internal SectionColumns? DocumentColumns;    // (iteration 67: the first section's columns; null: one)
	internal int[] PageMargins => DocumentMargins ?? Settings.Margins;
	internal bool PageLandscape => DocumentLandscape ?? Settings.Landscape;
	/// (iteration 82) The document's own paper (1/100 inch, width and height upright: the file's, or Page Setup's; null:
	/// none, the default's then), and the paper its pages are on: the document's, the default (Page Setup's Make default),
	/// else the default printer's (Letter or A4 without one).
	internal int[]? DocumentPaper;
	internal int[] PagePaper => DocumentPaper ?? Settings.Paper ?? PrinterPaper();

	int[] PrinterPaper() {
		try {
			using var printer = new System.Drawing.Printing.PrintDocument();
			var paper = printer.DefaultPageSettings.PaperSize;
			return Papers.Upright(paper.Width, paper.Height, twips: false) ?? Papers.RegionDefault;
		} catch (Exception) {
			return Papers.RegionDefault;
		}
	}

	/// The document's own page setup changed (opened, new): its pages laid out again if they are different now.
	internal void SetDocumentPageSetup(int[]? margins, bool? landscape, SectionColumns? columns = null, int[]? paper = null) {
		var (oldMargins, oldLandscape, oldColumns, oldPaper) = (PageMargins, PageLandscape, DocumentColumns, PagePaper);
		(DocumentMargins, DocumentLandscape, DocumentColumns, DocumentPaper) = (margins, landscape, columns, paper);
		if (PageMargins.SequenceEqual(oldMargins) && PageLandscape == oldLandscape && columns == oldColumns && PagePaper.SequenceEqual(oldPaper)) return;
		pageStartsStale = bandPagesStale = true;
		ApplyPrintLayout();
	}

	/// The page that other programs lay the document out on (twips): (iteration 82) the document's paper (PagePaper) and
	/// Page Setup's margins and orientation.
	internal PageLayout PageLayout() {
		var (width, height) = (PagePaper[0], PagePaper[1]);    // 1/100 inch
		if (PageLandscape) (width, height) = (height, width);
		int[] m = PageMargins is { Length: 4 } margins ? margins : new[] { 100, 100, 100, 100 };
		static int Twips(int hundredths) => (int)Math.Round(hundredths * 14.4);
		return new PageLayout(Twips(width), Twips(height), Twips(m[0]), Twips(m[1]), Twips(m[2]), Twips(m[3]), PageLandscape, Twips(Bands.HeaderFromEdge), Twips(Bands.FooterFromEdge), DocumentColumns);
	}

	// ---- Cross-references (iteration 50, rte_crossrefs.cs)

	/// Brings the cross-references' numbers up to date, as one undo step; returns how many changed. (Iteration 52) The
	/// tables of contents too: made again from the headings (tables, F9), or only their page numbers (before saving and
	/// printing).
	internal int UpdateReferences(bool tables = false) {
		var doc = Editor.TextDocument;
		// (iteration 57) the markers' case put back first (Change Case made "[[xreF:_TOC1]]" of them before it left hidden
		// text alone)
		if (Editor.IsHandleCreated && Editor.Text.Contains("[[", StringComparison.Ordinal)) {
			doc.BeginEditCollection();
			try { CrossRefs.RepairCase(doc); }
			finally { doc.EndEditCollection(); }
		}
		if (!Editor.IsHandleCreated || !CrossRefs.Any(Editor.Text)) return 0;
		int changed;
		doc.BeginEditCollection();
		try { changed = Footnotes.Update(doc) + CrossRefs.Update(doc); }    // (iteration 51: the footnotes' numbers too)
		finally { doc.EndEditCollection(); }
		return changed + UpdateContents(pagesOnly: !tables);
	}

	void UpdateReferencesCommand() {
		bool contents = Toc.Any(Editor.Text);
		int changed = UpdateReferences(tables: true);
		SetStatus(contents ? T("status.allUpToDate", "The cross-references, footnotes and tables of contents are up to date.")
			: changed == 0 ? T("status.crossRefsUpToDate", "The cross-references are up to date.") : TN("status.crossRefsUpdated", "{0} cross-reference updated.", "{0} cross-references updated.", changed));
	}

	// ---- Tables of contents (iteration 52, rte_toc.cs)

	/// Makes every table of contents again from the headings (pagesOnly: only their page numbers brought up to date), as
	/// one undo step; returns how many tables were made again (pagesOnly: how many page numbers changed).
	internal int UpdateContents(bool pagesOnly = false) {
		if (!Editor.IsHandleCreated) return 0;
		var doc = Editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (!Toc.Any(story) || pagesOnly && !story.Contains("[[toc=", StringComparison.Ordinal)) return 0;
		int changed = 0, styles = Styles.Styles.Count;
		WithoutGapsChanging(() => {
			doc.BeginEditCollection();
			try {
				if (!pagesOnly)
					foreach (var table in Enumerable.Reverse(Toc.Tables(story))) {
						Toc.Rebuild(doc, Styles, table, table.Options, TextWidthPoints());
						changed++;
					}
				int numbers = Toc.UpdatePages(doc, PageOf());
				if (pagesOnly) changed = numbers;
			} finally { doc.EndEditCollection(); }
		});
		if (Styles.Styles.Count != styles) StylesChanged();
		return changed;
	}

	/// Inserts a table of contents at the caret's paragraph (Toc.Insert), or, with the caret in one, makes that one
	/// again with these options; its page numbers as printed. One undo step.
	internal void InsertContents(Toc.Options options, string title) {
		var doc = Editor.TextDocument;
		int styles = Styles.Styles.Count;
		WithoutGapsChanging(() => {
			doc.BeginEditCollection();
			try {
				string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
				int caret = doc.Selection2.Start;
				int at = Toc.TableAt(story, caret) is { } table ? Toc.Rebuild(doc, Styles, table, options, TextWidthPoints())
					: Toc.Insert(doc, Styles, caret, options, title, TextWidthPoints());
				Toc.UpdatePages(doc, PageOf());
				// the caret after the table
				story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
				if (at >= 0 && Toc.Tables(story).FirstOrDefault(t => t.Start == at) is { } made) {
					int after = story.IndexOf('\r', Math.Max(made.Start, made.End - 1));
					after = after < 0 ? story.Length : after + 1;
					doc.Selection2.SetRange(Math.Min(after, story.Length - 1), Math.Min(after, story.Length - 1));
				}
			} finally { doc.EndEditCollection(); }
		});
		if (Styles.Styles.Count != styles) StylesChanged();
	}

	// The page each position is printed on (1: the first), as the pages are laid out now.
	Func<int, int> PageOf() {
		List<int> starts;
		using (var printer = CreatePrinter()) {
			var pages = printer.PageHeights();
			starts = pages.Select(p => p.Start).ToList();
			measuredPages = pageGaps.Applied ? null : (pages, printer.LastNotes, printer.NoteFontUsed);
		}
		return at => Math.Max(1, starts.FindLastIndex(s => s <= at) + 1);
	}
	// The pages PageOf laid out (without gaps), for the sheets that follow when nothing changed in between (SheetsNow:
	// laying out a 100-page document's pages again took up to 0.8 s of its opening)
	(List<(int Start, float Height)> Pages, List<List<Footnotes.NoteLine>> Notes, (string Name, float Size) NoteFont)? measuredPages;

	// The text's width on the page (points): where a table of contents' page numbers go.
	float TextWidthPoints() => api.PageTextWidth() / 20f;

	void TableOfContentsCommand() {
		var doc = Editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var existing = Toc.TableAt(story, doc.Selection2.Start);
		using var dlg = new TableOfContentsDialog(existing?.Options ?? new Toc.Options(), existing == null);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var o = dlg.Options;
		string title = dlg.TitleText;
		Run(() => api.InsertTableOfContents(o.Levels, title, o.PageNumbers, o.RightAlign, o.Leader), M("InsertTableOfContents", o.Levels, title, o.PageNumbers, o.RightAlign, o.Leader));
		SetStatus(existing != null ? T("status.tocRemade", "The table of contents was made again with these options.") : T("status.tocInserted", "Table of contents inserted. Ctrl+click an entry to go to its heading; F9 (or right-click it > Update table) brings it up to date."));
	}

	// Ctrl+click on a cross-reference: the caret to the paragraph it refers to (on a footnote's number: its note, to
	// change).
	bool GoToReferenceTarget(int position) {
		var doc = Editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		if (Footnotes.At(story, position) is { } note) {
			BeginInvoke(() => Run(() => EditFootnote(note)));
			return true;
		}
		// (iteration 52) a table of contents' entry: its heading
		if (Toc.EntryAt(story, position) is { } heading) {
			int target = CrossRefs.TargetPosition(story, heading);
			if (target < 0) { SetStatus("⚠ " + T("status.tocHeadingGone", "This entry's heading is gone: update the table of contents (F9).")); return true; }
			var line = doc.Range2(target, target);
			line.StartOf(tomConstants.tomParagraph, 0);
			Editor.Select(line.Start, 0);
			Editor.ScrollToCaret();
			SetStatus(T("status.tocHeading", "The heading of the table of contents' entry."));
			return true;
		}
		if (!CrossRefs.Any(story) || CrossRefs.ReferenceAt(story, position) is not { } name) return false;
		int at = CrossRefs.TargetPosition(story, name);
		if (at < 0) { SetStatus("⚠ " + T("status.crossRefGone", "The paragraph this cross-reference referred to is gone.")); return true; }
		var paragraph = doc.Range2(at, at);
		paragraph.StartOf(tomConstants.tomParagraph, 0);
		Editor.Select(paragraph.Start, 0);
		Editor.ScrollToCaret();
		SetStatus(T("status.crossRefTarget", "The paragraph the cross-reference refers to."));
		return true;
	}

	// ---- Footnotes (iteration 51, rte_footnotes.cs)

	List<Footnotes.Footnote>? footnotesRead;    // (the document's footnotes until its text changes)
	List<Toc.Leader>? leadersRead;    // (iteration 52: its tables of contents' leaders, likewise)
	List<ParagraphFrames.Framed>? framesRead;    // (iteration 87: its paragraphs' borders and shading, likewise)

	/// The document's Footnote Text style's font ("": the text's) and size (null: it has none: NoteSize).
	internal (string Font, float Size)? NoteStyle => Styles.Find(Footnotes.StyleName) is { Size: > 0 } style ? (style.Font, style.Size) : null;

	/// The size the document's notes have now (points): its Footnote Text style's, else NoteSize of its text's.
	internal float NoteSizeNow() => Footnotes.NoteFontOf(Editor.TextDocument, DocumentFootnotes(), Editor.Font.Name, Editor.Font.SizeInPoints, NoteStyle).Size;

	/// Gives the document's notes a size (its Footnote Text style, made if it has none; its font the text's).
	internal void SetNoteStyle(float size) {
		if (Styles.Find(Footnotes.StyleName) is { } style) style.Size = size;
		else Styles.Styles.Add(new ParagraphStyle {
			Name = Footnotes.StyleName, Size = size,
			Font = Footnotes.NoteFontOf(Editor.TextDocument, DocumentFootnotes(), Editor.Font.Name, Editor.Font.SizeInPoints).Name    // (the notes' text's)
		});
		StylesChanged();
		bandPagesStale = true;    // (the pages laid out again: the notes take more or less room)
		lastEdit = DateTime.MinValue;
		Editor.Invalidate();
	}
	List<List<Footnotes.NoteLine>> sheetNotes = new();    // each sheet's notes' lines and their font, as the pages were last laid out
	(string Name, float Size) sheetNoteFont = ("Segoe UI", 9);
	List<Footnotes.Footnote> DocumentFootnotes() =>
		footnotesRead ??= Footnotes.All(Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "");

	// Insert > Footnote: a new footnote at the caret, or, with the caret on one, its note to change.
	void FootnoteCommand() {
		var doc = Editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int caret = doc.Selection2.Start;
		if (Footnotes.At(story, caret) is { } existing) { EditFootnote(existing); return; }
		using var dlg = new FootnoteDialog(T("dlg.footnote.titleInsert", "Insert Footnote"), "", NoteSizeNow());
		if (dlg.ShowDialog(this) != DialogResult.OK || Footnotes.Plain(dlg.Note).Trim() == "") return;
		string note = Footnotes.Paragraphs(Footnotes.ToTags(dlg.Note));
		NoteSizeFrom(dlg);
		Run(() => api.InsertFootnote(note), M("InsertFootnote", note));
		SetStatus(T("status.footnoteInserted", "Footnote inserted. Its note is at the foot of the page (Show Pages, printing); hover over its number to see it, Ctrl+click to change it."));
	}

	// The notes' size the dialog was given another (all the document's notes: its Footnote Text style).
	void NoteSizeFrom(FootnoteDialog dlg) {
		if (Math.Abs(dlg.NoteSize - NoteSizeNow()) < 0.05f) return;
		float size = dlg.NoteSize;
		Run(() => api.FootnoteSize = size, M("FootnoteSize", size));
	}

	void EditFootnote(Footnotes.Footnote footnote) {
		using var dlg = new FootnoteDialog(T("dlg.footnote.title", "Footnote {0}", footnote.Number), footnote.Text, NoteSizeNow());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		NoteSizeFrom(dlg);
		int index = Footnotes.All(Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "").FindIndex(f => f.Start == footnote.Start) + 1;
		if (index <= 0) return;
		string note = Footnotes.Paragraphs(Footnotes.ToTags(dlg.Note));
		if (Footnotes.Plain(dlg.Note).Trim() == "") Run(() => api.DeleteFootnote(index), M("DeleteFootnote", index));
		else Run(() => api.SetFootnote(index, note), M("SetFootnote", index, note));
	}

	void CrossReferenceCommand() {
		var targets = CrossRefs.Targets(Editor.TextDocument);
		if (targets.Count == 0) {
			MessageBox.Show(this, T("msg.noNumberedParagraphs", "There are no numbered paragraphs to refer to yet. Number the paragraphs first (Numbering, Legal, or a numbered style such as Numbered Paragraph)."), AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
			return;
		}
		int caret = Editor.TextDocument.Selection2.Start;
		using var dlg = new CrossReferenceDialog(targets, Settings.CrossReferenceFull, Math.Max(0, targets.FindLastIndex(t => t.Start <= caret)));
		if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Chosen is not int index) return;
		if (Settings.CrossReferenceFull != dlg.Full) { Settings.CrossReferenceFull = dlg.Full; Settings.Save(); }
		string show = dlg.Full ? "full" : "number";
		Run(() => api.InsertCrossReference(index + 1, show), M("InsertCrossReference", index + 1, show));
		SetStatus(T("status.crossRefInserted", "Cross-reference inserted: its number follows the paragraph's. Ctrl+click it to go there; F9 updates them all."));
	}

	/// The header or footer changed: shown at once (iteration 49), not at the strips' next look, half a second on.
	internal void BandsChangedNow() {
		Editor.SectionsChanged();
		bandsShown = "";
		pageStartsStale = true;    // (iteration 55: a taller header or footer moves the page breaks)
		if (!PageViewActive) bandPagesStale = true;
		UpdateBands();
	}

	// (iteration 67) Layout > Columns: the selected paragraphs' (with continuous breaks around them), or the caret's section's
	void ColumnsCommand() {
		var doc = Editor.TextDocument;
		var (at, to) = (doc.Selection2.Start, doc.Selection2.End);
		var sections = Sections.Resolve(doc, Bands, DocumentColumns);
		int index = Sections.IndexAt(sections, at);
		using var dialog = new ColumnsDialog(sections[index].Columns, index + 1, sections.Count > 1, to > at, (int)(TextWidthPoints() * 20));
		if (dialog.ShowDialog(this) != DialogResult.OK) return;
		var columns = dialog.Result;
		int gap = (int)Math.Round((columns?.Gap ?? 720) / 14.4);
		string widths = columns?.Widths is { } own ? string.Join(",", own.Select(w => (int)Math.Round(w / 14.4))) : "";
		if (to > at) api.SetColumnsOn(at, to, columns?.Count ?? 1, gap, columns?.Line ?? false, widths);
		else api.SetColumns(at, columns?.Count ?? 1, gap, columns?.Line ?? false, widths);
		SetStatus(columns == null ? T("status.oneColumn", "One column.") : T("status.columns", "{0} columns, on the printed pages: you go on editing in one column, and Print Preview shows them.", columns.Count));
	}

	void HeaderFooterCommand() {
		// (iteration 63) the caret in a section after a break: that section's
		var doc = Editor.TextDocument;
		if (Sections.BreakFor(doc, doc.Selection2.Start) is { } sectionBreak) {
			var breaks = Sections.Breaks(doc);
			int index = breaks.FindIndex(b => b.MarkStart == sectionBreak.MarkStart) + 1;
			var resolved = Sections.Resolve(breaks, Bands);
			using var sectionDialog = new HeaderFooterDialog(Sections.CopyBands(resolved[index].Bands), DefaultBands(), sectionBreak.Settings, index + 1, new[] { PageMargins[2], PageMargins[3] });
			if (sectionDialog.ShowDialog(this) != DialogResult.OK) return;
			doc.BeginEditCollection();
			try { Sections.Write(doc, sectionBreak, sectionDialog.SectionResult()!); }
			finally { doc.EndEditCollection(); }
			Editor.Modified = true;
			BandsChangedNow();
			if (sectionDialog.MakeDefault) MakeDefault(sectionDialog.Result());
			SetStatus(T("status.headerFooterSection", "Header, footer and page numbers set for section {0}.", index + 1));
			return;
		}
		using var dlg = new HeaderFooterDialog(Bands, DefaultBands());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var bands = dlg.Result();
		if (!bands.SameAs(Bands)) {
			Bands = bands;
			Editor.Modified = true;    // (saved with the document)
			BandsChangedNow();
		}
		if (dlg.MakeDefault) MakeDefault(bands);
		SetStatus((dlg.MakeDefault ? T("status.headerFooterNew", "Header and footer set for this document and new ones.") : T("status.headerFooter", "Header and footer set for this document."))
			+ " " + T("status.headerFooterWhere", "File > Print Preview shows them; Word and OpenDocument files get them too."));
	}

	// ---- line spacing (iteration 22)

	// ---- the Paragraph dialog (iteration 33; Line Spacing Options opens it too)

	void ParagraphCommand() {
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		ParagraphFormat format;
		using (FirstParagraphAsIs(changes: false)) format = global::ParagraphFormat.Read(doc, Styles, start, end);
		var style = StyleEngine.StyleAt(doc, Styles, start);
		var font = doc.Range2(start, Math.Min(start + 1, doc.Range2(0, 0).StoryLength)).Font;
		using var dlg = new ParagraphDialog(format, forStyle: false, style.Name, font.Name ?? style.Font, font.Size > 0 ? font.Size : style.Size, doc.DefaultTabStop);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		string settings = dlg.Result.Describe();
		doc.BeginEditCollection();    // (one undo step)
		try {
			if (settings != "") { api.SetParagraphFormat(settings); Record(M("SetParagraphFormat", settings)); }
			if (dlg.DefaultTabStop is float spacing) { api.DefaultTabStop = spacing; Record(M("DefaultTabStop", spacing)); }
		} finally {
			doc.EndEditCollection();
		}
		if (dlg.SetAsDefault) {
			// the style gets what the dialog showed (what the paragraphs share, with the changes)
			var shown = format.Copy();
			var result = dlg.Result;
			foreach (var field in typeof(global::ParagraphFormat).GetFields()) if (field.GetValue(result) is { } value) field.SetValue(shown, value);
			var updated = style.Copy();
			shown.ToStyle(updated);
			string styleSettings = StyleSettings(updated);
			api.ModifyStyle(style.Name, styleSettings);
			Record(M("ModifyStyle", style.Name, styleSettings));
			if (dlg.NewDocumentsToo) { api.SaveStylesAsDefault(); Record(M("SaveStylesAsDefault")); }
			SetStatus(dlg.NewDocumentsToo ? T("status.styleParagraphsNew", "The {0} style has these paragraph settings now, in new documents too.", style.Name)
			: T("status.styleParagraphs", "The {0} style has these paragraph settings now (Styles > Save Styles as My Default keeps them for new documents).", style.Name));
		}
	}

	// ---- line numbers (iteration 19)

	/// The document's line numbering (saved with it); SetLineNumbering changes it.
	internal LineNumbering Numbering { get; private set; } = new();

	internal LineNumberGutter LineGutter => lineGutter;

	/// Sets the document's line numbering; changesDocument: it is a change to the document (not one being opened).
	internal void SetLineNumbering(LineNumbering numbering, bool changesDocument = true) {
		var value = numbering.Copy().Normalized();
		if (changesDocument && !value.SameAs(Numbering)) Editor.Modified = true;
		Numbering = value;
		lineGutter.Numbering = value;
		pageStartsStale = true;
		if (!lineGutter.NeedsPageStarts) lineGutter.PageStarts = null;
		FitPagePadding();
		UpdateUi();
	}

	// The page: in Print Layout, a sheet as wide as the paper at the zoom (its margins either side of the text, the line
	// numbers in the left one), centred on grey, which scrolls sideways if the window is narrower; Page Width and One
	// Page set the zoom first. Otherwise the whole width, the line numbers taking the padding's place left of the text
	// (the text moves as little as it can).
	void FitPagePadding() {
		if (fittingPage) return;
		fittingPage = true;
		try {
			var host = pageHost;
			if (pageLayout == null) {
				host.AutoScroll = false;
				host.BackColor = AppTheme.Paper;
				pagePanel.Dock = DockStyle.Fill;
				int left = LogicalToDeviceUnits(28);
				if (lineGutter.On) left = Math.Max(LogicalToDeviceUnits(4), left - Math.Max(lineGutter.Width, LogicalToDeviceUnits(30)));
				pagePanel.Padding = new Padding(left, LogicalToDeviceUnits(14), LogicalToDeviceUnits(16), 0);
				Editor.TextMargins = headerStrip.Inset = footerStrip.Inset = (0, 0);
				lineGutter.MaxWidth = int.MaxValue;
				return;
			}
			if (Settings.ZoomMode is "PageWidth" or "OnePage") FitZoom(host);
			pageZoom = EditorZoom();
			var (marginLeft, text, marginRight) = PrintLayoutPage.Pixels(pageLayout, pageZoom, DeviceDpi);
			lineGutter.MaxWidth = Math.Max(LogicalToDeviceUnits(20), marginLeft - LogicalToDeviceUnits(4));    // (the numbers within the left margin)
			// (the lines are broken with the printer's measurements but drawn with the screen's, whose hinted letters are a
			// little wider: 3% more, and more if a line drawn still reaches the edge, PageOverhang)
			int gutter = lineGutter.On ? lineGutter.Width : 0, bar = SystemInformation.VerticalScrollBarWidth;
			int slack = LogicalToDeviceUnits(6) + (int)Math.Ceiling(text * 0.03) + pageOverhang;
			int padLeft = Math.Max(LogicalToDeviceUnits(4), marginLeft - gutter), padRight = Math.Max(LogicalToDeviceUnits(4), marginRight - bar - slack);
			int width = padLeft + gutter + text + bar + slack + padRight;
			pagePanel.Dock = DockStyle.None;
			// (the margins are the editor's own, inside it (TextMargins), so that Show Pages' bands in them scroll with the
			// text in the same move; the left one beside the line numbers is the panel's, left of them; the scroll bar at
			// the page's right edge)
			int panelLeft = lineGutter.On ? padLeft : 0;
			pagePanel.Padding = new Padding(panelLeft, LogicalToDeviceUnits(14), 0, 0);
			Editor.TextMargins = headerStrip.Inset = footerStrip.Inset = (padLeft - panelLeft, padRight);
			host.BackColor = PrintLayoutPage.Surround;
			host.AutoScrollMinSize = new Size(width, 0);
			host.AutoScroll = width > host.ClientSize.Width;
			int x = Math.Max(0, (host.ClientSize.Width - width) / 2);
			pagePanel.SetBounds(x + host.AutoScrollPosition.X, 0, width, host.ClientSize.Height);
			// (iteration 46: what it holds laid out at its new size: a document opened at a large zoom left the editor
			// and the line numbers as for the page at 100%, the text at the window's edge, the footer off its middle)
			pagePanel.PerformLayout();
		} finally {
			fittingPage = false;
		}
	}

	bool fittingPage;
	float autoZoom = -1;    // the zoom Page Width or One Page set last

	// Page Width: the zoom at which the page (paper width) fills the window, a little space either side; One Page: at
	// which a page's height fits.
	void FitZoom(Control host) {
		// a zoom this didn't set (Ctrl+wheel, a script's ZoomFactor): the user's from now on, not undone
		if (autoZoom > 0 && Math.Abs(EditorZoom() - autoZoom) > 0.01f) { Settings.ZoomMode = "Manual"; return; }
		var (paperWidth, paperHeight) = PrintLayoutPage.Paper(pageLayout!, DeviceDpi);
		int gap = LogicalToDeviceUnits(12);
		double zoom = Settings.ZoomMode == "OnePage"
			? (host.ClientSize.Height - pagePanel.Padding.Top - (headerStrip.Visible ? headerStrip.Height : 0) - (footerStrip.Visible ? footerStrip.Height : 0) - gap) / paperHeight
			: (host.ClientSize.Width - 2 * gap) / paperWidth;
		if (host.ClientSize.Width <= 0 || host.ClientSize.Height <= 0) return;
		float value = (float)Math.Clamp(Math.Round(zoom, 3), 0.3, 4.0);
		if (Math.Abs(value - EditorZoom()) < 0.005f) { autoZoom = EditorZoom(); return; }
		autoZoom = value;
		try { Editor.ZoomFactor = value; } catch (ArgumentException) { }
		autoZoom = EditorZoom();
	}

	// Zoom In / Out / 100%: the zoom is the user's (not Page Width's or One Page's).
	void ManualZoom(float zoom) {
		Settings.ZoomMode = "Manual";
		try { Editor.ZoomFactor = zoom; } catch (ArgumentException) { }
		FitPagePadding();
	}

	// ---- Print Layout (iteration 24)

	PrintLayoutTarget? pageTarget;
	PageLayout? pageLayout;    // Page Setup's paper and margins while Print Layout is on (null: off)

	/// The page's text width in twips while Print Layout is on (null: off).
	internal int? PrintTextWidth => pageLayout?.TextWidth;
	float pageZoom = -1;
	int pageOverhang;          // pixels the editor was widened by, for lines drawn wider than the text width
	bool overhangStale = true; // the text or layout changed since the drawn lines were measured

	// After the editor painted: whether a line it drew reaches its right edge (wider letters than the printer's); if so,
	// the page widens to fit (only ever widens, until the zoom or the page changes).
	void CheckPageOverhang() {
		if (pageLayout == null || !overhangStale || !Editor.IsHandleCreated) return;
		overhangStale = false;
		const int EM_GETLINECOUNT = 0x00BA, EM_LINEINDEX = 0x00BB, EM_GETFIRSTVISIBLELINE = 0x00CE;
		var doc = Editor.TextDocument;
		int count = (int)SendMessage(Editor.Handle, EM_GETLINECOUNT, IntPtr.Zero, IntPtr.Zero);
		int first = (int)SendMessage(Editor.Handle, EM_GETFIRSTVISIBLELINE, IntPtr.Zero, IntPtr.Zero), widest = 0;
		try {
			for (int line = first; line < Math.Min(count, first + 60); line++) {
				int start = (int)SendMessage(Editor.Handle, EM_LINEINDEX, (IntPtr)line, IntPtr.Zero);
				int next = line + 1 < count ? (int)SendMessage(Editor.Handle, EM_LINEINDEX, (IntPtr)(line + 1), IntPtr.Zero) : Editor.TextLength;
				if (start < 0 || next <= start) continue;
				int last = next - 1;
				while (last > start && (doc.Range2(last, last + 1).Text ?? "") is "\r" or " " or "\v" or "\u0007") last--;
				doc.Range2(last, last + 1).GetPoint(0x0A | 0x200, out int x, out int y);    // tomEnd | tomAllowOffClient (screen)
				if (Editor.PointToClient(new Point(x, y)).Y > Editor.ClientSize.Height) break;
				widest = Math.Max(widest, Editor.PointToClient(new Point(x, y)).X);
			}
		} catch (COMException) { return; }
		int room = Editor.ClientSize.Width - Editor.TextMargins.Right - LogicalToDeviceUnits(2);
		if (widest <= room) return;
		// (iteration 46: at most a few percent of the text's width, as the screen's letters are wider than the printer's;
		// measured while the page was being laid out afresh (a document opened at a large zoom), a line could seem a
		// page wider than it is, and the page stayed that wide, its footer's middle far to the right)
		int cap = Math.Max(LogicalToDeviceUnits(8), (int)Math.Ceiling((pageLayout?.TextWidth ?? 0) / 1440f * DeviceDpi * EditorZoom() * 0.04f));
		int wider = Math.Min(cap, pageOverhang + widest - room + LogicalToDeviceUnits(4));
		if (wider <= pageOverhang) return;
		pageOverhang = wider;
		FitPagePadding();
	}

	float EditorZoom() {
		try { return Editor.ZoomFactor; } catch (Exception) { return 1f; }
	}

	/// Lays the editor's lines out as on the printed page (Print Layout), or at the window's width (Word wrap) or not
	/// wrapped. Again after Page Setup, and when the editor gets a new window.
	internal void ApplyPrintLayout() {
		if (!Editor.IsHandleCreated || IsDisposed) return;
		autoZoom = -1;    // (Page Width or One Page fit the zoom afresh, whatever it is now)
		var old = pageTarget;
		if (Settings.PrintLayout) {
			pageLayout = PageLayout();
			string? printer = null;
			try { var settings = new PrinterSettings(); if (settings.IsValid) printer = settings.PrinterName; } catch (Exception) { }
			pageTarget = old != null && old.IsPrinter && old.Device == printer ? old : new PrintLayoutTarget(printer);
			pageTarget.Apply(Editor, pageLayout.TextWidth);
		} else {
			pageLayout = null;
			pageTarget = null;
			PrintLayoutTarget.Release(Editor, Editor.WordWrap);
		}
		if (old != null && old != pageTarget) old.Dispose();    // (after the editor has let go of it)
		pageOverhang = 0;
		overhangStale = true;
		FitPagePadding();
		lineGutter.DocumentChanged();
		pageStartsStale = bandPagesStale = true;
		if (PagesShown) RefreshPages();
		UpdateUi();
	}

	// ---- Multiple Pages (iteration 27)

	PageOverview? pageView;

	/// Whether Multiple Pages is showing (the printed pages, read-only, in the editor's place).
	internal bool PagesShown => pageView is { Visible: true };

	// The commands that leave the pages showing: viewing, printing and saving the document; any other goes back to
	// editing first (so it doesn't change the text out of sight).
	static readonly HashSet<RibbonCommand> PageViewCommands = new() {
		RibbonCommand.ZoomMultiplePages, RibbonCommand.ZoomIn, RibbonCommand.ZoomOut, RibbonCommand.Zoom100, RibbonCommand.Print,
		RibbonCommand.PrintPreview, RibbonCommand.ExportPdf, RibbonCommand.ExportPdfButton, RibbonCommand.ExportWord, RibbonCommand.ExportOdt, RibbonCommand.Save,
		RibbonCommand.SaveAs, RibbonCommand.PageSetup, RibbonCommand.HeaderFooter, RibbonCommand.Columns, RibbonCommand.ShowBands, RibbonCommand.Outline, RibbonCommand.Ruler,
		RibbonCommand.Help, RibbonCommand.About, RibbonCommand.AutomationApi, RibbonCommand.Exit, RibbonCommand.NewWindow, RibbonCommand.OpenInNewWindow,
		RibbonCommand.Theme, RibbonCommand.Language,
	};

	/// Shows the printed pages in the editor's place (show) or goes back to editing where the caret was.
	internal void ShowPages(bool show) {
		if (show == PagesShown) { if (show) RefreshPages(); return; }
		if (show) {
			if (pageView == null) {
				pageView = new PageOverview { Dock = DockStyle.Fill, Visible = false };
				pageView.PageChosen += (page, start) => {
					ShowPages(false);
					Editor.Select(Math.Min(start, Editor.TextLength), 0);
					Editor.ScrollToCaret();
					SetStatus(T("status.page", "Page {0}.", page + 1));
				};
				pageView.Closing += () => ShowPages(false);
				sourceSplit.Parent!.Controls.Add(pageView);
			}
			Cursor = Cursors.WaitCursor;
			try { RefreshPages(); } finally { Cursor = Cursors.Default; }
			pageView.Visible = true;
			pageView.BringToFront();
			pageView.Focus();
			SetStatus(TN("status.multiplePages", "Multiple Pages: {0} page, read-only. Double-click a page to edit it there; Esc goes back. Zoom or Ctrl+wheel shows them larger or smaller.", "Multiple Pages: {0} pages, read-only. Double-click a page to edit it there; Esc goes back. Zoom or Ctrl+wheel shows them larger or smaller.", pageView.PageCount));
		} else {
			pageView!.Visible = false;
			pageView.Clear();
			if (SourceActive) sourceBox.Focus(); else Editor.Focus();
			SetStatus("");
		}
		UpdateUi();
	}

	// The pages again, as they print now (after Page Setup, the header and footer, a script's change).
	void RefreshPages() {
		if (pageView == null) return;
		List<int> starts;
		using (var measuring = CreatePrinter()) starts = measuring.PageStarts();
		using var printer = CreatePrinter();
		int caret = Editor.TextDocument.Selection2.Start;
		pageView.Show(printer.Document, starts, Math.Max(0, starts.FindLastIndex(start => start <= caret)));
	}

	void LineNumbersCommand() {
		using var dlg = new LineNumbersDialog(Numbering);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var numbering = dlg.Result(Numbering);
		if (numbering.SameAs(Numbering)) return;
		SetLineNumbering(numbering);
		Record(numbering.On ? M("SetLineNumbers", LineNumbering.ModeName(numbering.Mode), numbering.Start, numbering.CountBy, (float)(numbering.Distance / 20.0))
			: M("LineNumbers", "none"));
		SetStatus(numbering.On ? T("status.lineNumbers", "Line numbers: {0}. Saved with the document (RTF, Word, OpenDocument, HTML, Markdown).", numbering.Describe()) : T("status.lineNumbersOff", "Line numbers are off."));
	}

	// (iteration 82) AhkWord's own Page Setup (rte_pagesetup.cs): this document's, and, ticked, new documents' too
	void PageSetupCommand() {
		using var dlg = new PageSetupForm(PageMargins, PageLandscape, PagePaper, Bands.HeaderFromEdge, Bands.FooterFromEdge, Bands.DifferentFirst);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var c = dlg.Result;
		SetPageSetup(c.Margins, c.Landscape, c.Paper, c.HeaderDistance, c.FooterDistance, c.DifferentFirst, c.MakeDefault);
		Record(M("SetPageSetup", c.Margins[0] / 100.0, c.Margins[1] / 100.0, c.Margins[2] / 100.0, c.Margins[3] / 100.0, c.Landscape, c.Paper[0] / 100.0, c.Paper[1] / 100.0));
		SetStatus(c.MakeDefault ? T("status.pageSetupNew", "Page setup set for this document and new ones.") : T("status.pageSetup", "Page setup set for this document (saved with it)."));
	}

	/// (iteration 82) The document's page setup: margins (1/100 inch: left, right, top, bottom), orientation, paper (1/100
	/// inch, upright), the header's and footer's distances from the edges (1/100 inch; null: as they are) and the different
	/// first page (null: as it is); makeDefault: new documents' too (not the different first page: the Header and Footer
	/// dialog's default). Its pages laid out again.
	internal void SetPageSetup(int[] margins, bool landscape, int[] paper, int? headerDistance = null, int? footerDistance = null, bool? differentFirst = null, bool makeDefault = false) {
		bool pageChanged = !PageMargins.SequenceEqual(margins) || PageLandscape != landscape || !PagePaper.SequenceEqual(paper);
		(DocumentMargins, DocumentLandscape, DocumentPaper) = (margins.ToArray(), landscape, paper.ToArray());
		static int? Distance(int d) => d == PageBands.DefaultDistance ? null : d;
		var (header, footer) = (headerDistance is int h ? Distance(h) : Bands.HeaderDistance, footerDistance is int f ? Distance(f) : Bands.FooterDistance);
		bool first = differentFirst ?? Bands.DifferentFirst;
		bool bandsChanged = header != Bands.HeaderDistance || footer != Bands.FooterDistance || first != Bands.DifferentFirst;
		(Bands.HeaderDistance, Bands.FooterDistance, Bands.DifferentFirst) = (header, footer, first);
		if (makeDefault) {
			(Settings.Margins, Settings.Landscape, Settings.Paper) = (margins.ToArray(), landscape, paper.ToArray());
			(Settings.HeaderDistance, Settings.FooterDistance) = (header, footer);
			Settings.Save();
		}
		if (pageChanged || bandsChanged) Editor.Modified = true;    // (saved with the document)
		pageStartsStale = bandPagesStale = true;
		if (bandsChanged) BandsChangedNow();
		ApplyPrintLayout();    // (the page's new width)
	}

	// Print Preview: the printed pages in a window fitted to the screen's working area (PagePreviewForm), from the page
	// the caret is on; double-clicking a page edits the document there.
	void PrintPreviewCommand() {
		UpdateReferences();
		List<int> starts;
		using (var measuring = CreatePrinter()) starts = measuring.PageStarts();
		using var printer = CreatePrinter(dialog: true);
		using var preview = new PagePreviewForm(DocumentName, this);
		int caret = Editor.TextDocument.Selection2.Start;
		preview.Load += (s, e) => preview.Pages.Show(printer.Document, starts, Math.Max(0, starts.FindLastIndex(start => start <= caret)));
		preview.PrintRequested += () => PrintCommand(preview);
		if (preview.ShowDialog(this) == DialogResult.OK && preview.ChosenStart is int start) {
			if (PagesShown) ShowPages(false);
			Editor.Select(Math.Min(start, Editor.TextLength), 0);
			Editor.ScrollToCaret();
			Editor.Focus();
		}
	}

	void PrintCommand() => PrintCommand(this);

	void PrintCommand(IWin32Window owner) {
		UpdateReferences();
		using var printer = CreatePrinter(dialog: true);
		// (iteration 58) page by page in the status bar; Esc stops it
		printer.Progress = (page, pages) => ShowProgress(page < pages ? T("status.printingPage", "Printing: page {0} of {1}. Esc stops it.", page, pages) : T("status.printingHanding", "Printing: handing the pages to the printer…"));
		printer.CancelRequested = PrintStop;
		using var dlg = new PrintDialog { Document = printer.Document, UseEXDialog = true };
		if (dlg.ShowDialog(owner) != DialogResult.OK) return;
		EscPressed();    // (an Esc from before, the dialog's: not this one's)
		printer.Document.Print();
		SetStatus(printer.Cancelled ? T("status.printingStopped", "Printing stopped (Esc): the pages printed so far may come out.") : T("status.sentTo", "Sent to {0}", printer.Document.PrinterSettings.PrinterName));
	}

	void ExportPdfDialog() {
		using var dlg = new SaveFileDialog {
			Title = T("file.exportPdf", "Export as PDF"), Filter = "PDF (*.pdf)|*.pdf",
			FileName = Path.ChangeExtension(DocumentName, ".pdf"),
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Cursor = Cursors.WaitCursor;
		try {
			SetStatus(T("status.exported", "Exported {0}", api.ExportPdf(dlg.FileName)));
		} catch (OperationCanceledException ex) {
			SetStatus(ex.Message);
		} finally {
			Cursor = Cursors.Default;
		}
	}

	// ---------------------------------------------------------------- Links

	// Ctrl+click on a link opens it: web and mail addresses at once, anything else (a file or a program could be
	// behind it) after asking.
	void OpenLinkCommand(string url) {
		// (iteration 87) a link to a place in the document ("#name", Word's bookmarks): the caret goes to its target there
		if (url.StartsWith('#') && url.Length > 1) {
			var doc = Editor.TextDocument;
			string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
			int at = story.IndexOf("[[xref:" + url[1..] + "]]", StringComparison.Ordinal);
			if (at < 0) at = story.IndexOf("[[toc:" + url[1..] + "]]", StringComparison.Ordinal);
			if (at < 0) { SetStatus("⚠ " + T("status.noLinkTarget", "This link's place ({0}) is not in the document.", url[1..])); return; }
			// (past the target's hidden markers: the caret at its first character that shows)
			while (story.IndexOf("]]", at, StringComparison.Ordinal) is int close and >= 0 && (story.AsSpan(at).StartsWith("[[xref:") || story.AsSpan(at).StartsWith("[[toc:")))
				at = close + 2;
			Editor.Select(at, 0);
			Editor.ScrollToCaret();
			SetStatus(T("status.wentToLinkTarget", "Went to {0}.", url[1..]));
			return;
		}
		bool web = Regex.IsMatch(url, @"^(https?|mailto):", RegexOptions.IgnoreCase);
		if (!web && MessageBox.Show(this, T("msg.openLink", "Open this link?\n\n{0}", url), AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
		try {
			System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true })?.Dispose();
			SetStatus(T("status.opened", "Opened {0}", url));
		} catch (Exception ex) {
			SetStatus("⚠ " + T("status.couldNotOpen", "Could not open {0}: {1}", url, ex.Message));
		}
	}

	// Paste: the text, its paragraphs' styles (StyleClipboard), and numbered paragraphs joining the numbering where they
	// land (StyleEngine.JoinPasted), as one undo step. paste: RichEdit's own Paste (tests give one of their own, so as to
	// leave the user's clipboard alone).
	void PasteText() => PasteText(Editor.Paste);

	// Paste as Plain Text (iteration 65; Ctrl+Shift+V, as Word and browsers have it): the clipboard's text only, taking
	// the formatting where it lands (a web page's or a chat's fonts, sizes, spacing and rules left behind); numbered
	// paragraphs join the numbering where they land; one undo step. RichEdit's own paste of the clipboard's text (as
	// Ctrl+V's is RichEdit's own); paste: tests give one of their own, so as to leave the user's clipboard alone.
	internal void PastePlainText() => PastePlainText(PlainPasteForTests ?? (() => Editor.Paste(DataFormats.GetFormat(DataFormats.UnicodeText))));
	internal Action? PlainPasteForTests = null;    // (set by tests)

	internal void PastePlainText(Action paste) {
		var doc = Editor.TextDocument;
		doc.BeginEditCollection();
		try {
			BeforeNumber();
			int start = doc.Selection2.Start;
			paste();
			StyleEngine.JoinPasted(doc, Styles, start, doc.Selection2.Start);
		} finally {
			doc.EndEditCollection();
		}
		Editor.ScrollToCaret();
		RefreshOutline();
	}

	internal void PasteText(Action paste) {
		var doc = Editor.TextDocument;
		bool added;
		doc.BeginEditCollection();
		try {
			// whole paragraphs pasted at the start of a numbered paragraph's text (the caret can't go before its number):
			// before the paragraph, as Word puts them, so that each is an item of its own; else a pasted paragraph's own
			// number stayed in front of its text after the item's ("2.<tab>4.<tab>Delta")
			BeforeNumber();
			added = StyleClipboard.Paste(Editor, Styles, paste);
			var (start, end) = StyleClipboard.Pasted;
			if (RtfTables.DropStrayMergeMarks(doc) > 0) end = doc.Selection2.Start;    // (iteration 92: a table with merged and nested cells)
			StyleEngine.JoinPasted(doc, Styles, start, end);
		} finally {
			doc.EndEditCollection();
		}
		if (added) FillStyleBox();
		RefreshOutline();
	}

	// The selection's start at a numbered (legal or numbered style's) paragraph's text, and what is pasted whole
	// paragraphs (its text ends with a paragraph mark): the start moved before the paragraph's number.
	void BeforeNumber() {
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		if (LegalLists.Touched(doc, selection.Start, selection.Start).FirstOrDefault() is not { Legal: true, PrefixLength: > 0 } p) return;
		if (selection.Start != p.Start + p.PrefixLength) return;
		string text;
		try { text = StyleClipboard.Read()?.GetData(DataFormats.UnicodeText) as string ?? ""; }
		catch (ExternalException) { return; }
		if (!text.EndsWith("\n") && !text.EndsWith("\r")) return;
		selection.SetRange(p.Start, Math.Max(p.Start, selection.End));
	}

	// ---------------------------------------------------------------- Iteration 10: legal numbering

	// The caret never stays inside a legal item's number (typing there would break it): it goes after the number, or,
	// moving left from the start of the text, to the end of the paragraph before.
	void GuardLegalCaret() {
		if (guardingCaret || !Editor.IsHandleCreated) return;
		var selection = Editor.TextDocument.Selection2;
		int caret = selection.Start;
		if (selection.End != caret) { lastCaret = -1; return; }
		var para = Editor.TextDocument.Range2(caret, caret);
		para.Expand(tomConstants.tomParagraph);
		if (LegalLists.IsLegal(para.Para.ListType)) {
			int textStart = para.Start + LegalLists.Prefix.Match(para.Text ?? "").Length;
			if (caret < textStart) {
				int target = lastCaret == textStart && para.Start > 0 ? para.Start - 1 : textStart;
				guardingCaret = true;
				try { selection.SetRange(target, target); } finally { guardingCaret = false; }
				caret = target;
			}
		}
		lastCaret = caret;
	}

	// After changes the commands don't cover (pasting, dropping, RichEdit's own editing), legal numbers are put right, as
	// an undo step of their own. Not after Undo / Redo, which restore numbers along with the text.
	void CheckLegalNumbers() {
		if (skipLegalCheck) { skipLegalCheck = false; return; }
		try {
			if (LegalLists.AnyLegal(Editor.TextDocument)) LegalLists.Renumber(Editor);
		} catch (Exception ex) { SetStatus("⚠ " + ex.Message); }
	}

	// ---------------------------------------------------------------- Iteration 9: Word, OpenDocument and tables

	// Writes a copy of the document as .docx or .odt; the current file stays the one being edited.
	void ExportOfficeDialog(string format) {
		var types = format == "docx" ? WordTypes : OdtTypes;
		using var dlg = new SaveFileDialog {
			Title = format == "docx" ? "Export as Word" : "Export as OpenDocument",
			Filter = Filter(types),
			FileName = Path.ChangeExtension(DocumentName, "." + format),
			InitialDirectory = FileName != null ? Path.GetDirectoryName(FileName) : StartFolder
		};
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Cursor = Cursors.WaitCursor;
		SetStatus(T("status.converting", "Converting..."));
		try {
			string via = ViaOf(types, dlg.FilterIndex);
			string path = format == "docx" ? api.ExportDocx(dlg.FileName, via) : api.ExportOdt(dlg.FileName, via);
			SetStatus(T("status.exported", "Exported {0}", path) + " " + T("status.convertedBy", "(converted by {0})", api.LastConverter));
		} finally {
			Cursor = Cursors.Default;
		}
	}

	void InsertTableDialog() {
		using var dlg = new InsertTableDialog(Settings.TableRows, Settings.TableColumns, Settings.TableHeader, fromFile: true);
		var result = dlg.ShowDialog(this);
		if (result is not (DialogResult.OK or DialogResult.Yes)) return;
		(Settings.TableRows, Settings.TableColumns, Settings.TableHeader) = (dlg.Rows, dlg.Columns, dlg.Header);
		Settings.Save();
		if (result == DialogResult.Yes) { InsertTableFromFileCommand(dlg.Header); return; }
		Run(() => api.InsertTable(dlg.Rows, dlg.Columns, dlg.Header, true), M("InsertTable", dlg.Rows, dlg.Columns, dlg.Header, true));
	}

	// (iteration 91) Insert > Table > From File: a workbook, CSV, TSV or JSON file as a table; a workbook's sheet and
	// range asked for.
	void InsertTableFromFileCommand(bool header) {
		using var open = new OpenFileDialog { Title = T("dlg.tableFromFile.title", "Insert table from file"), InitialDirectory = StartFolder,
			Filter = T("dlg.tableFromFile.types", "Excel workbooks, CSV and JSON") + " (*.xlsx;*.xlsm;*.csv;*.tsv;*.txt;*.json)|*.xlsx;*.xlsm;*.csv;*.tsv;*.txt;*.json|"
				+ T("file.all", "All files") + " (*.*)|*.*" };
		if (open.ShowDialog(this) != DialogResult.OK) return;
		string path = open.FileName, sheet = "", range = "";
		if (Path.GetExtension(path).ToLowerInvariant() is ".xlsx" or ".xlsm" or ".xltx" or ".xltm") {
			List<string> sheets;
			try { sheets = TabularData.SheetNames(path); }
			catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException or InvalidDataException) {
				SetStatus("⚠ " + T("status.tableFromFile.unreadable", "{0} could not be read: {1}", Path.GetFileName(path), ex.Message));
				System.Media.SystemSounds.Beep.Play();
				return;
			}
			using var choose = new TableFromFileDialog(Path.GetFileName(path), sheets, header);
			if (choose.ShowDialog(this) != DialogResult.OK) return;
			(sheet, range, header) = (sheets.Count > 1 ? choose.Sheet : "", choose.Range, choose.Header);
		}
		int rows = 0;
		Run(() => rows = api.InsertTableFromFile(path, header, sheet, range), M("InsertTableFromFile", path, header, sheet, range));
		if (rows > 0) SetStatus(T("status.tableFromFile.done", "A table of {0} rows from {1}", rows, Path.GetFileName(path)));
	}

	// (iteration 94) Table Layout > Data: a table linked to its file (EditorApi.LinkTable, rte_tabledata.cs TableSources).
	void UpdateTableCommand() {
		int rows = 0;
		string source = "";
		Run(() => { source = api.TableSource(); rows = api.UpdateTable(); }, M("UpdateTable"));
		if (rows > 0) SetStatus(T("status.table.updated", "Table updated from {0}: {1} rows", Path.GetFileName(source.Split('|')[0]), rows));
	}

	void SaveTableToSourceCommand() {
		int count = -1;
		string source = "";
		string? warning = null;
		Run(() => warning = api.SaveTableCheck());    // (iteration 94: past the link's range, or over rows below the data)
		if (warning == null) return;
		if (warning != "") {
			var answer = MessageBox.Show(this, warning + "\n\n" + T("msg.saveTable.ask", "Write the table to the file anyway?"), "AhkWord",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
			if (answer != DialogResult.Yes) { SetStatus(T("status.table.notSaved", "The table was not written to its file")); return; }
		}
		bool force = warning != "";
		Run(() => { source = api.TableSource(); count = api.SaveTableToSource(0, force); }, M("SaveTableToSource"));
		if (count < 0) return;
		string file = source.Split('|')[0];
		SetStatus(Path.GetExtension(file).ToLowerInvariant() is ".xlsx" or ".xlsm"
			? T("status.table.savedCells", "{0} cells changed in {1}", count, Path.GetFileName(file))
			: T("status.table.savedRows", "{0} rows written to {1}", count, Path.GetFileName(file)));
	}

	void LinkTableCommand() {
		if (!api.InTable()) { SetStatus("⚠ " + T("status.table.caretFirst", "Put the caret in a table first.")); return; }
		using var open = new OpenFileDialog { Title = T("dlg.linkTable.title", "Link the table to a file"), InitialDirectory = FileName is { Length: > 0 } f ? Path.GetDirectoryName(f) : StartFolder,
			Filter = T("dlg.tableFromFile.types", "Excel workbooks, CSV and JSON") + " (*.xlsx;*.xlsm;*.csv;*.tsv;*.txt;*.json)|*.xlsx;*.xlsm;*.csv;*.tsv;*.txt;*.json|"
				+ T("file.all", "All files") + " (*.*)|*.*" };
		if (open.ShowDialog(this) != DialogResult.OK) return;
		string path = open.FileName, sheet = "", range = "";
		if (Path.GetExtension(path).ToLowerInvariant() is ".xlsx" or ".xlsm" or ".xltx" or ".xltm") {
			List<string> sheets;
			try { sheets = TabularData.SheetNames(path); }
			catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException or InvalidDataException) {
				SetStatus("⚠ " + T("status.tableFromFile.unreadable", "{0} could not be read: {1}", Path.GetFileName(path), ex.Message));
				System.Media.SystemSounds.Beep.Play();
				return;
			}
			using var choose = new TableFromFileDialog(Path.GetFileName(path), sheets, true);
			if (choose.ShowDialog(this) != DialogResult.OK) return;
			(sheet, range) = (sheets.Count > 1 ? choose.Sheet : "", choose.Range);
		}
		bool done = false;
		Run(() => { api.LinkTable(path, 0, sheet, range); done = true; }, M("LinkTable", path, 0, sheet, range));
		if (done) SetStatus(T("status.table.linked", "Table linked to {0}", Path.GetFileName(path)));
	}

	// Table Tools > Properties: the dialog, its changes made in one step (one undo step), and recorded as the calls they are.
	void TablePropertiesCommand() {
		var (rows, selected) = api.TableFormats();
		using var dlg = new TablePropertiesDialog(rows, selected.Count, api.PageTextWidth());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		var changes = dlg.Changes;
		api.ApplyTableChanges(changes);
		foreach (string call in changes.Calls()) Record(call);
	}

	// (iteration 92) Table Tools > Split Cells: how many columns and rows, the cell's own span to start with.
	void SplitCellsCommand() {
		var (columns, rows) = api.CellSpan();
		using var dlg = new SplitCellsDialog(columns, rows);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Run(() => api.SplitCells(dlg.Columns, dlg.Rows), M("SplitCells", dlg.Columns, dlg.Rows));
	}

	// (iteration 92) Table Tools > Delete Cells: the cells after them move left or the cells below them up, or their
	// rows or columns go, as Word asks.
	void DeleteCellsCommand() {
		if (!api.InTable()) throw new InvalidOperationException("Put the caret in a table first.");
		using var dlg = new DeleteCellsDialog();
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Run(() => api.DeleteCells(dlg.How), M("DeleteCells", dlg.How));
	}

	// ---------------------------------------------------------------- Iteration 5: spelling

	// The Home tab's Cut, Copy and Paste (styles kept, the page view's gaps left out), for the right-click menu too.
	Action cutAction = () => { }, copyAction = () => { }, pasteAction = () => { }, pastePlainAction = () => { };

	// The right-click menu, as Word's: spelling suggestions on a misspelled word; Cut, Copy and Paste; the footnote's
	// note on a footnote; Font and Paragraph; in a table, the table commands (which then act on the cell clicked).
	// Iteration 51: always this menu (it was only on a misspelled word or in a table, else the library's formatting
	// toolbar, or RichEdit's own spelling menu, which came up in turns), the caret moved by WriterEditBox's right-button
	// press as RichEdit's own hit-testing has it.
	bool ShowContextMenu(Point screenPoint, int position) {
		// a picture: selected (as in Word), with Alt text... in the menu
		int picture = Editor.PictureAt(screenPoint);
		var selected = Editor.TextDocument.Selection2;
		if (picture < 0 && selected.End == selected.Start + 1 && selected.Text == "\uFFFC") picture = selected.Start;
		if (picture >= 0) {
			selected.SetRange(picture, picture + 1);
			var pictureMenu = new ContextMenuStrip { ShowImageMargin = false };
			pictureMenu.Items.Add(new ToolStripMenuItem(T("menu.cut", "Cu&t"), null, (s, e) => Run(Editor.Cut, M("Cut"))));
			pictureMenu.Items.Add(new ToolStripMenuItem(T("menu.copy", "&Copy"), null, (s, e) => Run(Editor.Copy, M("Copy"))));
			pictureMenu.Items.Add(new ToolStripMenuItem(T("menu.paste", "&Paste"), null, (s, e) => Run(Editor.Paste, M("Paste"))));
			pictureMenu.Items.Add(new ToolStripSeparator());
			pictureMenu.Items.Add(new ToolStripMenuItem(T("menu.altText", "&Alt text..."), null, (s, e) => Run(() => AltTextCommand(picture))));
			pictureMenu.Closed += (s, e) => BeginInvoke(pictureMenu.Dispose);
			pictureMenu.Show(screenPoint);
			return true;
		}
		var menu = TextMenu(position);
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
		return true;
	}

	// (iteration 93) A block of table cells put on the clipboard as a table of its own (RTF) and as tab-separated text
	// (Excel takes it as cells); false: no block (the selection is copied as it is).
	bool CopyCellBlock() {
		if (RtfTables.BlockData(Editor) is not { } data) return false;
		var obj = new DataObject();
		obj.SetData(DataFormats.Rtf, data.Rtf);
		obj.SetData(DataFormats.UnicodeText, data.Text);
		Clipboard.SetDataObject(obj, true);
		return true;
	}

	// The right-click menu for text (position: the character clicked).
	internal ContextMenuStrip TextMenu(int position) {
		var tom = Editor.TextDocument.Selection2;    // (WinForms' SelectionLength is short across table rows)
		// (a click outside the selection moved the caret there already, as in Word: WriterEditBox.OnMouseDown)
		bool inTable = RtfTables.At(Editor, position) != null;
		var error = Editor.EnableSpellCheck && Spelling != null ? api.SpellingErrorAt(position) : null;
		// (the suggestions only with no selection, or the misspelled word's own, as in Word: right-clicked inside a
		// selection of more, the menu is for the selection)
		if (error != null && tom.End > tom.Start && (tom.Start < error.Start || tom.End > error.Start + error.Length)) error = null;
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		if (error != null) {
			AddSpellingItems(menu, error);
			menu.Items.Add(new ToolStripSeparator());
		}
		bool selection = tom.End > tom.Start || RtfTables.HasBlock(Editor);
		menu.Items.Add(new ToolStripMenuItem(T("menu.cut", "Cu&t"), null, (s, e) => Run(cutAction, M("Cut"))) { Enabled = selection, ShortcutKeyDisplayString = "Ctrl+X" });
		menu.Items.Add(new ToolStripMenuItem(T("menu.copy", "&Copy"), null, (s, e) => Run(copyAction, M("Copy"))) { Enabled = selection, ShortcutKeyDisplayString = "Ctrl+C" });
		menu.Items.Add(new ToolStripMenuItem(T("menu.paste", "&Paste"), null, (s, e) => Run(pasteAction, M("Paste"))) { ShortcutKeyDisplayString = "Ctrl+V" });
		menu.Items.Add(new ToolStripMenuItem(T("menu.pastePlain", "Paste as Plain Te&xt"), null, (s, e) => Run(pastePlainAction, M("PastePlainText"))) { ShortcutKeyDisplayString = TKeys("Ctrl+Shift+V") });
		menu.Items.Add(new ToolStripSeparator());
		void Add(string text, Action action, string macro, string keys = "") =>
			menu.Items.Add(new ToolStripMenuItem(text, null, (s, e) => Run(action, macro)) { ShortcutKeyDisplayString = TKeys(keys) });
		string story = Editor.TextDocument.Range2(0, Editor.TextDocument.Range2(0, 0).StoryLength).Text ?? "";
		if ((Footnotes.At(story, position) ?? Footnotes.At(story, Math.Max(0, position - 1))) is { } footnote) {
			menu.Items.Add(new ToolStripMenuItem(T("menu.editFootnote", "Edit &footnote {0}...", footnote.Number), null, (s, e) => Run(() => EditFootnote(footnote))));
			menu.Items.Add(new ToolStripSeparator());
		}
		if (Toc.TableAt(story, position) != null) {
			menu.Items.Add(new ToolStripMenuItem(T("menu.updateToc", "&Update table"), null, (s, e) => Run(() => api.UpdateTableOfContents(false), M("UpdateTableOfContents", false))) { ShortcutKeyDisplayString = "F9" });
			menu.Items.Add(new ToolStripMenuItem(T("menu.updateTocPages", "Update page &numbers only"), null, (s, e) => Run(() => api.UpdateTableOfContents(true), M("UpdateTableOfContents", true))));
			menu.Items.Add(new ToolStripSeparator());
		}
		menu.Items.Add(new ToolStripMenuItem(T("menu.font", "&Font..."), null, (s, e) => Run(FontDialogCommand)));
		menu.Items.Add(new ToolStripMenuItem(T("menu.paragraph", "P&aragraph..."), null, (s, e) => Run(ParagraphCommand)));
		if (!inTable) return menu;
		menu.Items.Add(new ToolStripSeparator());
		Add(T("menu.insertRowAbove", "Insert row &above"), () => api.InsertRow(false), M("InsertRow", false));
		Add(T("menu.insertRowBelow", "Insert row &below"), () => api.InsertRow(true), M("InsertRow", true), "Ctrl+Shift+Ins");
		Add(T("menu.insertColumnLeft", "Insert column &left"), () => api.InsertColumn(false), M("InsertColumn", false));
		Add(T("menu.insertColumnRight", "Insert column &right"), () => api.InsertColumn(true), M("InsertColumn", true));
		menu.Items.Add(new ToolStripSeparator());
		int selectedRows = api.SelectedRows();
		if (selectedRows > 1) Add(T("menu.deleteRows", "&Delete {0} rows", selectedRows), () => api.DeleteSelectedRows(), M("DeleteSelectedRows"), "Del");
		else Add(T("menu.deleteRow", "&Delete row"), () => api.DeleteRow(), M("DeleteRow"), "Ctrl+Shift+Del");
		Add(T("menu.deleteColumn", "Delete colu&mn"), () => api.DeleteColumn(), M("DeleteColumn"));
		Add(T("menu.deleteTable", "Delete &table"), () => api.DeleteTable(), M("DeleteTable"));
		// (iteration 92) the cells: Merge when the selection reaches over more than one
		menu.Items.Add(new ToolStripSeparator());
		if (api.SelectionSpansCells()) Add(T("menu.mergeCells", "Mer&ge cells"), () => api.MergeCells(), M("MergeCells"));
		menu.Items.Add(new ToolStripMenuItem(T("menu.splitCells", "Spl&it cells..."), null, (s, e) => Run(SplitCellsCommand)));
		menu.Items.Add(new ToolStripMenuItem(T("menu.deleteCells", "Delete c&ells..."), null, (s, e) => Run(DeleteCellsCommand)));
		menu.Items.Add(new ToolStripSeparator());
		menu.Items.Add(new ToolStripMenuItem(T("menu.tableProperties", "Table &properties..."), null, (s, e) => Run(TablePropertiesCommand)));
		return menu;
	}

	// The right-click menu on a misspelled word: suggestions, "delete repeated word" and "add to dictionary".
	bool ShowSpellingMenu(Point screenPoint, int position) =>
		Editor.EnableSpellCheck && Spelling != null && api.SpellingErrorAt(position) is { } error && ShowSpellingMenu(screenPoint, error);

	bool ShowSpellingMenu(Point screenPoint, Spelling.Error error) {
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		AddSpellingItems(menu, error);
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
		return true;
	}

	void AddSpellingItems(ContextMenuStrip menu, Spelling.Error error) {
		var bold = new Font(menu.Font, FontStyle.Bold);
		menu.Disposed += (s, e) => bold.Dispose();
		if (error.Action == Spelling.DeleteAction) {
			menu.Items.Add(new ToolStripMenuItem(T("menu.deleteRepeated", "Delete repeated word \"{0}\"", error.Word), null, (s, e) => Run(() => api.DeleteRepeatedWord(error))) { Font = bold });
		} else {
			var suggestions = error.Action == Spelling.ReplaceAction && error.Replacement is { Length: > 0 } replacement ? new List<string> { replacement } : Spelling?.Suggest(error.Word) ?? new List<string>();
			foreach (var suggestion in suggestions)
				menu.Items.Add(new ToolStripMenuItem(suggestion, null, (s, e) => Run(() => api.ReplaceRange(error.Start, error.Length, suggestion))) { Font = bold });
			if (suggestions.Count == 0) menu.Items.Add(new ToolStripMenuItem(T("menu.noSuggestions", "(no suggestions)")) { Enabled = false });
			menu.Items.Add(new ToolStripSeparator());
			menu.Items.Add(new ToolStripMenuItem(T("menu.addToDictionary", "Add \"{0}\" to dictionary", error.Word), null, (s, e) => Run(() => {
				api.AddToDictionary(error.Word);
				SetStatus(T("status.addedToDictionary", "\"{0}\" was added to the Windows dictionary.", error.Word));
			})));
		}
	}

	void NextSpellingErrorCommand() {
		if (Spelling == null) throw new InvalidOperationException("The Windows spell checker is not available.");
		if (api.NextSpellingError() == "") {
			SetStatus(T("status.noSpellingErrors", "No spelling errors found."));
			return;
		}
		// Open the suggestions just below the word.
		var point = Editor.GetPositionFromCharIndex(Editor.SelectionStart);
		point.Offset(0, Editor.Font.Height + LogicalToDeviceUnits(4));
		ShowSpellingMenu(Editor.PointToScreen(point), Editor.SelectionStart);
	}

	// ---------------------------------------------------------------- Iteration 5: outline

	internal void SetOutlineVisible(bool visible) {
		outlineSplit.Panel1Collapsed = !visible;
		if (visible) RefreshOutline();
	}

	internal bool OutlineVisible => !outlineSplit.Panel1Collapsed;

	// ---------------------------------------------------------------- Iteration 78: the ruler (rte_ruler.cs)

	readonly RulerView ruler = new();

	internal void SetRulerVisible(bool visible) {
		Settings.ShowRuler = visible;
		ruler.Visible = visible && !PagesShown;
		UpdateUi();
	}

	/// Shows the page's place and, with paragraph, the caret's paragraph's indents and tab stops (only on a selection
	/// change or a command: the page's place on every paint, cheaply).
	void FollowRuler(bool paragraph = false) {
		if (!ruler.Visible || !Editor.IsHandleCreated || !ruler.IsHandleCreated) return;
		var area = Editor.TextArea;
		int textLeft = ruler.PointToClient(Editor.PointToScreen(new Point(area.Left, 0))).X, textRight, paperLeft, paperRight;
		if (pageLayout != null) {
			var (marginLeft, text, marginRight) = PrintLayoutPage.Pixels(pageLayout, EditorZoom(), DeviceDpi);
			textRight = textLeft + text; paperLeft = textLeft - marginLeft; paperRight = textRight + marginRight;
		} else {
			textRight = textLeft + area.Width; paperLeft = textLeft; paperRight = textRight;
		}
		var frame = new RulerView.Frame(paperLeft, textLeft, textRight, paperRight, Editor.PixelsPerPoint);
		ruler.Show(frame, paragraph ? RulerParagraph() : ruler.ShownParagraph);
	}

	// The paragraph at the selection's start (Word's ruler shows the first one's); none in a table.
	RulerView.Paragraph? RulerParagraph() {
		try {
			if (CaretInTable()) return null;
			var doc = Editor.TextDocument;
			var start = doc.Selection2.Start;
			var p = doc.Range2(start, start).Para;
			var tabs = new (float, int)[p.TabCount];
			for (int i = 0; i < tabs.Length; i++) { p.GetTab(i, out float position, out int alignment, out _); tabs[i] = (position, alignment); }
			return new(p.FirstLineIndent, p.LeftIndent, p.RightIndent, tabs, doc.DefaultTabStop);
		} catch (Exception) { return null; }
	}

	// (iteration 83) A click on the ruler's white strip: a left tab stop there for the selected paragraphs.
	void RulerTab(float points) {
		string tab = points.ToString("0.##", CultureInfo.InvariantCulture);
		api.AddTab(tab);
		Record(M("AddTab", tab));
		FollowRuler(true);
		SetStatus(T("status.rulerTab", "Left tab stop at {0} (double-click it for Tabs).", RulerMeasure(points)));
	}

	// A measure on the ruler as its scale shows it: inches, or centimetres.
	string RulerMeasure(float points) => ruler.Metric ? (points / 72 * 2.54).ToString("0.##", CultureInfo.CurrentCulture) + " cm" : (points / 72).ToString("0.###", CultureInfo.CurrentCulture) + "\"";

	// (iteration 84) A marker dragged on the ruler and let go: a tab stop moved or cleared, or an indent (the shown
	// paragraph's values, given to every selected paragraph, as the Paragraph dialog does); one undo step, recorded.
	void RulerDragged(RulerView.Handle handle, float from, float? to) {
		if (ruler.ShownParagraph is not { } p) return;
		static string P(float v) => v.ToString("0.##", CultureInfo.InvariantCulture);
		if (handle == RulerView.Handle.Tab) {
			api.MoveTab(P(from), to is float t ? P(t) : "");
			Record(M("MoveTab", P(from), to is float t2 ? P(t2) : ""));
			FollowRuler(true);
			SetStatus(to is float at ? T("status.rulerTabMoved", "Tab stop moved to {0}.", RulerMeasure(at)) : T("status.rulerTabCleared", "Tab stop cleared."));
			return;
		}
		if (to is not float x) return;
		string settings = handle switch {
			RulerView.Handle.FirstLine => "FirstLine=" + P(x - p.Left),
			RulerView.Handle.Hanging => "LeftIndent=" + P(x) + "; FirstLine=" + P(p.Left + p.First - x),
			RulerView.Handle.LeftIndent => "LeftIndent=" + P(x),
			RulerView.Handle.RightIndent => "RightIndent=" + P(ruler.TextWidth - x),
			_ => "",
		};
		if (settings == "") return;
		api.SetParagraphFormat(settings);
		Record(M("SetParagraphFormat", settings));
		FollowRuler(true);
	}

	// While a marker is dragged: where it would go, on the status bar.
	void RulerDragging(RulerView.Handle handle, float? at) {
		if (handle == RulerView.Handle.None) return;
		if (at is not float x) { SetStatus(T("status.rulerDragOff", "Let go to clear the tab stop.")); return; }
		var p = ruler.ShownParagraph;
		string text = handle switch {
			RulerView.Handle.FirstLine when p != null => T("status.rulerFirstLine", "First line indent: {0}", RulerMeasure(x - p.Left)),
			RulerView.Handle.Hanging => T("status.rulerHanging", "Hanging indent: {0}", RulerMeasure(x)),
			RulerView.Handle.LeftIndent => T("status.rulerLeftIndent", "Left indent: {0}", RulerMeasure(x)),
			RulerView.Handle.RightIndent => T("status.rulerRightIndent", "Right indent: {0}", RulerMeasure(ruler.TextWidth - x)),
			_ => T("status.rulerTabAt", "Tab stop: {0}", RulerMeasure(x)),
		};
		SetStatus(text);
	}

	// A double-click on the ruler: Tabs on a tab stop or the white strip (not in a table), Paragraph on an indent marker,
	// Page Setup on the margins (as Word's).
	void RulerDoubleClicked(RulerView.Part part) {
		switch (part) {
			case RulerView.Part.Tab or RulerView.Part.Text when ruler.ShownParagraph != null: TabsCommand(); break;
			case RulerView.Part.Indent: ParagraphCommand(); break;
			case RulerView.Part.Margin: PageSetupCommand(); break;
		}
		FollowRuler(true);
	}

	// The Tabs dialog by itself (the Paragraph dialog's, for the selected paragraphs).
	void TabsCommand() {
		var doc = Editor.TextDocument;
		var selection = doc.Selection2;
		global::ParagraphFormat format;
		using (FirstParagraphAsIs(changes: false)) format = global::ParagraphFormat.Read(doc, Styles, selection.Start, selection.End);
		using var dlg = new TabsDialog(format.Tabs, doc.DefaultTabStop);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		doc.BeginEditCollection();    // (one undo step)
		try {
			if (dlg.TabsChanged) { string tabs = string.Join(",", dlg.Tabs); api.SetTabs(tabs); Record(M("SetTabs", tabs)); }
			if (dlg.DefaultTabStop is float spacing && Math.Abs(spacing - doc.DefaultTabStop) > 0.01f) { api.DefaultTabStop = spacing; Record(M("DefaultTabStop", spacing)); }
		} finally {
			doc.EndEditCollection();
		}
	}

	internal List<(int Level, int Start, string Text)> CurrentOutline() {
		RefreshOutline();
		return outline;
	}

	// Rebuilds the tree when the headings have changed.
	/// The outline of a document just opened, with it, not a moment after.
	internal void OutlineNow() {
		outlineTimer.Stop();
		RefreshOutline();
	}

	void RefreshOutline() {
		if (!Editor.IsHandleCreated) return;
		List<(int Level, int Start, string Text)> headings;
		try {
			headings = api.Headings();
		} catch (COMException) {
			return;    // the document is being replaced; the next change refreshes again
		}
		if (outlineBuilt && headings.SequenceEqual(outline)) {
			HighlightOutline();
			return;
		}
		outline = headings;
		updatingOutline = true;
		outlineTree.BeginUpdate();
		try {
			outlineTree.Nodes.Clear();
			var parents = new TreeNode?[7];
			foreach (var (level, start, text) in headings) {
				string shown = CrossRefs.PlainText(text);    // (a numbered heading can be a cross-reference's target)
				var node = new TreeNode(shown.Length > 80 ? shown[..80] + "…" : shown) { Tag = start, ToolTipText = shown };
				TreeNode? parent = null;
				for (int l = level - 1; l >= 1 && parent == null; l--) parent = parents[l];
				(parent?.Nodes ?? outlineTree.Nodes).Add(node);
				parents[level] = node;
				for (int l = level + 1; l <= 6; l++) parents[l] = null;
			}
			outlineTree.ExpandAll();
		} finally {
			outlineTree.EndUpdate();
			updatingOutline = false;
		}
		outlineBuilt = true;
		outlineHint.Visible = headings.Count == 0;
		outlineTree.Visible = headings.Count > 0;
		HighlightOutline();
	}

	// Selects the heading of the section the caret is in.
	void HighlightOutline() {
		if (updatingOutline || outlineSplit.Panel1Collapsed || outline.Count == 0) return;
		int caret = Editor.SelectionStart;
		TreeNode? best = null;
		foreach (var node in AllNodes(outlineTree.Nodes))
			if (node.Tag is int start && start <= caret && (best == null || start >= (int)best.Tag!)) best = node;
		if (best != null && outlineTree.SelectedNode != best) outlineTree.SelectedNode = best;
		if (best == null) outlineTree.SelectedNode = null;
	}

	static IEnumerable<TreeNode> AllNodes(TreeNodeCollection nodes) {
		foreach (TreeNode node in nodes) {
			yield return node;
			foreach (var child in AllNodes(node.Nodes)) yield return child;
		}
	}

	// Puts the caret at a position and scrolls so that its line is at the top of the editor.
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	internal void GoToPosition(int position) {
		position = Math.Clamp(position, 0, Editor.TextLength);
		Editor.Select(position, 0);
		const int EM_GETFIRSTVISIBLELINE = 0x00CE, EM_LINESCROLL = 0x00B6;
		int first = (int)SendMessage(Editor.Handle, EM_GETFIRSTVISIBLELINE, IntPtr.Zero, IntPtr.Zero);
		int line = Editor.GetLineFromCharIndex(position);
		SendMessage(Editor.Handle, EM_LINESCROLL, IntPtr.Zero, (IntPtr)(line - first));
		Editor.Focus();
	}

	void ShowOutlineMenu(TreeNode node, Point screenPoint) {
		if (node.Tag is not int start) return;
		var level = outline.FirstOrDefault(h => h.Start == start).Level;
		var menu = new ContextMenuStrip { ShowImageMargin = false };
		menu.Items.Add(new ToolStripMenuItem(T("menu.goToHeading", "Go to heading"), null, (s, e) => Run(() => GoToPosition(start))));
		menu.Items.Add(new ToolStripSeparator());
		menu.Items.Add(new ToolStripMenuItem(T("menu.promote", "Promote (to Heading {0})", Math.Max(1, level - 1)), null, (s, e) => Run(() => api.SetHeadingLevel(start, level - 1))) { Enabled = level > 1 });
		menu.Items.Add(new ToolStripMenuItem(T("menu.demote", "Demote (to Heading {0})", Math.Min(6, level + 1)), null, (s, e) => Run(() => api.SetHeadingLevel(start, level + 1))) { Enabled = level < 6 });
		menu.Items.Add(new ToolStripMenuItem(T("menu.makeNormal", "Make normal text"), null, (s, e) => Run(() => api.SetHeadingLevel(start, 0))));
		menu.Closed += (s, e) => BeginInvoke(menu.Dispose);
		menu.Show(screenPoint);
	}

	// ---------------------------------------------------------------- Iteration 40: copy styles from a document

	// Home > Styles > Copy Styles from Document: a document, then which of its styles come into this one (and into
	// your defaults for new documents, if ticked).
	void CopyStylesCommand() {
		using var open = new OpenFileDialog { Title = T("file.copyStyles", "Copy Styles from Document"), InitialDirectory = StartFolder,
			Filter = T("file.withStyles", "Documents with styles") + " (*.rtf;*.docx;*.html;*.htm)|*.rtf;*.docx;*.html;*.htm|" + T("file.all", "All files") + " (*.*)|*.*" };
		if (open.ShowDialog(this) != DialogResult.OK) return;
		StyleSheet? source;
		try { source = StyleImport.From(open.FileName); }
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or System.Xml.XmlException) {
			MessageBox.Show(this, T("msg.fileUnreadable", "{0} could not be read: {1}", Path.GetFileName(open.FileName), ex.Message), AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}
		if (source == null) {
			MessageBox.Show(this, T("msg.noStylesInFile", "{0} has no paragraph styles that AhkWord can read (AhkWord's own RTF and HTML, Word's .docx, other RTF).", Path.GetFileName(open.FileName)), AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
			return;
		}
		using var dlg = new CopyStylesDialog(open.FileName, source, Styles);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		string names = string.Join(", ", dlg.Chosen);
		int count = api.CopyStyles(open.FileName, names, dlg.ForNewDocuments);
		Record(M("CopyStyles", open.FileName, names, dlg.ForNewDocuments));
		SetStatus(dlg.ForNewDocuments ? TN("status.copiedStylesNew", "Copied {0} style ({1}) from {2}, for new documents too.", "Copied {0} styles ({1}) from {2}, for new documents too.", count, names, Path.GetFileName(open.FileName))
			: TN("status.copiedStyles", "Copied {0} style ({1}) from {2}.", "Copied {0} styles ({1}) from {2}.", count, names, Path.GetFileName(open.FileName)));
	}

	// ---------------------------------------------------------------- Iteration 8: list levels

	void ListLevelsCommand() {
		using var dlg = new ListLevelsDialog(Levels, api.InList(), LegalLists.Format, AutoLevelStyles);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		if (dlg.AutomaticPerLevel != AutoLevelStyles) {
			AutoLevelStyles = Settings.AutoLevelStyles = dlg.AutomaticPerLevel;    // (an app setting, as before)
			Settings.Save();
			Record(M("AutomaticNumberingPerLevel", AutoLevelStyles));
		}
		// (the document's levels; and new documents' unless told otherwise)
		bool changed = !LevelsSame(Levels, dlg.Levels);
		Levels = dlg.Levels;
		if (dlg.ForNewDocuments) {
			Settings.ListLevels = dlg.Levels.Select(l => l.Copy()).ToList();
			Settings.Save();
		}
		if (changed) LevelsChanged();
		for (int i = 0; changed && i < 9; i++) {
			var l = Levels[i];
			Record(M("SetListLevel", i + 1, l.Style, l.Punctuation, l.Start, l.NumberAt ?? -1f, l.TextAt ?? -1f, dlg.ForNewDocuments));
		}
		if (dlg.LegalFormat != LegalLists.Format) {
			api.LegalFormat = dlg.LegalFormat.ToString();
			Record(M("LegalFormat", dlg.LegalFormat.ToString()));
		}
		if (dlg.ApplyToCurrentList) {
			api.ApplyListLevels();
			Record(M("ApplyListLevels"));
			SetStatus(T("status.listLevelsApplied", "List levels saved and applied to the current list."));
		} else {
			SetStatus(T("status.listLevelsSaved", "List levels saved. They are used for new numbered lists and when nesting with Tab."));
		}
	}

	void RestartNumberingCommand() {
		if (!api.InList()) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		if (Prompt(T("dlg.restartNumbering.title", "Restart numbering"), (T("dlg.restartNumbering.startAt", "Start at"), "1")) is not { } values) return;
		if (!int.TryParse(values[0], out int start) || start < 0) throw new ArgumentException("The start number must be 0 or more.");
		Run(() => api.RestartNumbering(start), M("RestartNumbering", start));
	}

	void OnFormClosing(object? sender, FormClosingEventArgs e) {
		RunHook(HookEvent.Closing, FileName ?? "");    // (iteration 56: waits for its macro, which can still save)
		if (!discardOnClose && !ConfirmDiscard()) e.Cancel = true;
		discardOnClose = false;
		if (!e.Cancel) SaveSettings();    // here, while the editor still reports its zoom (not after the window has closed)
	}

	// Called by automation: closes after the reply has been sent, so the calling script does not wait for a dialog.
	internal void CloseLater(bool discardChanges) {
		discardOnClose = discardChanges;
		BeginInvoke(Close);
	}

	void FontDialogCommand() {
		var font = api.SelectionFont();
		var tom = Editor.TextDocument.Selection2.Font;
		// (what the selection has; mixed counts as off, so that only what is changed in the dialog changes)
		var style = (tom.Bold == -1 ? FontStyle.Bold : 0) | (tom.Italic == -1 ? FontStyle.Italic : 0) | (tom.Underline == 1 ? FontStyle.Underline : 0) | (tom.StrikeThrough == -1 ? FontStyle.Strikeout : 0);
		var color = tom.ForeColor is >= 0 and <= 0xFFFFFF ? ColorTranslator.FromWin32(tom.ForeColor) : Color.Black;
		using var dlg = new FontDialog { ShowEffects = true, ShowColor = true, Color = color };
		try { dlg.Font = new Font(string.IsNullOrEmpty(font.Name) ? Editor.Font.Name : font.Name, font.Size > 0 ? font.Size : Editor.Font.Size, style); } catch (ArgumentException) { }
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		ApplyFontChoice(dlg.Font, dlg.Color, style, color);
	}

	// The Font dialog's choice, one undo step: its font and size, and what it changes of the styles and the colour.
	// (RichEdit's undo keeps each format change apart, edit collection or not: so all of it goes on at once)
	internal void ApplyFontChoice(Font chosen, Color chosenColor, FontStyle oldStyle, Color oldColor) {
		int Change(FontStyle flag) => (chosen.Style & flag) == (oldStyle & flag) ? -1 : chosen.Style.HasFlag(flag) ? 1 : 0;
		var (bold, italic, underline, strike) = (Change(FontStyle.Bold), Change(FontStyle.Italic), Change(FontStyle.Underline), Change(FontStyle.Strikeout));
		bool newColor = chosenColor.ToArgb() != oldColor.ToArgb();
		// (typed legal numbers take a new size through SetFontSize, which sizes them with their text)
		bool legal = LegalLists.Paragraphs(Editor.TextDocument).Any(p => p.PrefixLength > 0);
		api.SetFontLook(bold, italic, underline, strike, newColor ? chosenColor : null, chosen.Name, Editor.EnableFontSize && !legal ? chosen.SizeInPoints : null);
		if (Editor.EnableFontSize && legal) api.SetFontSize(chosen.SizeInPoints);
		Record(M("SetFontName", chosen.Name));
		if (Editor.EnableFontSize) Record(M("SetFontSize", chosen.SizeInPoints));
		else SetStatus(T("status.fontSizesFreeform", "Font applied. Font sizes need free-form mode (View > Free-form)."));
		if (bold >= 0 || italic >= 0 || underline >= 0 || strike >= 0) Record(M("SetFontStyle", bold, italic, underline, strike));
		if (newColor) Record(M("SetFontColor", chosenColor));
	}

	void InsertHyperlinkDialog() {
		string selected = Editor.SelectedText.Trim();
		if (Prompt(T("dlg.hyperlink.title", "Insert hyperlink"), (T("dlg.hyperlink.address", "Address"), "https://"), (T("dlg.hyperlink.display", "Text to display"), selected != "" ? selected : "")) is not { } values) return;
		string url = values[0], display = values[1] != "" ? values[1] : values[0];
		// (the library's two steps, the text and then its link, frozen and as one undo step: each was drawn, a flicker)
		var doc = Editor.TextDocument;
		Run(() => {
			doc.Freeze();
			doc.BeginEditCollection();
			try {
				Editor.InsertHyperlink(url, display);
			} finally {
				doc.EndEditCollection();
				doc.Unfreeze();
			}
		}, M("InsertHyperlink", url, display));
	}

	void InsertImageDialog() {
		if (ChooseImage() is not { } path) return;
		if (Prompt(T("dlg.image.title", "Insert image"), (T("dlg.image.altText", "Alt text: what the picture shows, for screen readers (empty for a decorative picture)"), Path.GetFileNameWithoutExtension(path))) is not { } values) return;
		Run(() => api.InsertImageFile(path, values[0]), M("InsertImageFile", path, values[0]));
	}

	// Right-click a picture > Alt text...
	void AltTextCommand(int position) {
		string current = api.PictureAltText(position);
		if (Prompt(T("dlg.altText.title", "Alt text"), (T("dlg.image.altText", "Alt text: what the picture shows, for screen readers (empty for a decorative picture)"), current)) is not { } values) return;
		if (values[0] == current) return;
		Run(() => api.SetPictureAltText(position, values[0]), M("SetPictureAltText", position, values[0]));
	}

	void InsertThumbnailDialog() {
		if (ChooseImage() is not { } path) return;
		if (Prompt(T("dlg.thumbnail.title", "Insert linked thumbnail"), (T("dlg.thumbnail.link", "Link to the full-size image"), new Uri(path).AbsoluteUri), (T("dlg.altText.title", "Alt text"), Path.GetFileNameWithoutExtension(path))) is not { } values) return;
		Run(() => api.InsertLinkedThumbnailFile(values[0], path, values[1]), M("InsertLinkedThumbnailFile", values[0], path, values[1]));
	}

	string? ChooseImage() {
		using var dlg = new OpenFileDialog { Title = T("dlg.image.title", "Insert image"), Filter = "Images (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif" };
		return dlg.ShowDialog(this) == DialogResult.OK ? dlg.FileName : null;
	}

	// ---------------------------------------------------------------- Find and replace

	void ShowFind(bool replace) {
		findForm ??= new FindReplaceForm(this);
		var (length, selected) = SourceActive ? (sourceBox.SelectionLength, sourceBox.SelectedText) : (Editor.SelectionLength, Editor.SelectedText);
		findForm.ShowFor(replace, length is > 0 and < 100 && !selected.Contains('\n') ? selected : null);
	}

	internal bool FindNext() {
		if (findForm == null || findForm.FindText == "") {
			ShowFind(false);
			return false;
		}
		if (SourceActive) return FindInSource();
		int index = api.Find(findForm.FindText, findForm.MatchCase, findForm.WholeWord);
		Record(M("Find", findForm.FindText, findForm.MatchCase, findForm.WholeWord));
		SetStatus(index < 0 ? T("status.notFound", "\"{0}\" was not found.", findForm.FindText) : "");
		return index >= 0;
	}

	internal void ReplaceNext() {
		if (findForm == null) return;
		var comparison = findForm.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
		if (SourceActive) {
			if (sourceBox.SelectionLength > 0 && string.Equals(sourceBox.SelectedText, findForm.FindText, comparison)) sourceBox.SelectedText = findForm.ReplaceText;
			FindInSource();
			return;
		}
		if (Editor.SelectionLength > 0 && string.Equals(Editor.SelectedText, findForm.FindText, comparison))
			Run(() => Editor.SelectedText = findForm.ReplaceText, M("InsertText", findForm.ReplaceText));
		FindNext();
	}

	internal void ReplaceAll() {
		if (findForm == null || findForm.FindText == "") return;
		int count = 0;
		if (SourceActive) {
			var pattern = new Regex((findForm.WholeWord ? @"\b" : "") + Regex.Escape(findForm.FindText) + (findForm.WholeWord ? @"\b" : ""), findForm.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
			string replacement = findForm.ReplaceText;
			string text = pattern.Replace(sourceBox.Text, m => { count++; return replacement; });
			if (count > 0) { sourceBox.SelectAll(); sourceBox.SelectedText = text; }    // one undo step
			SetStatus(TN("status.replacements", "{0} replacement made.", "{0} replacements made.", count));
			return;
		}
		Run(() => count = api.Replace(findForm.FindText, findForm.ReplaceText, findForm.MatchCase, findForm.WholeWord),
			M("Replace", findForm.FindText, findForm.ReplaceText, findForm.MatchCase, findForm.WholeWord));
		SetStatus(TN("status.replacements", "{0} replacement made.", "{0} replacements made.", count));
	}

	// ---------------------------------------------------------------- Keyboard: Tab / Shift+Tab in lists

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		if (keyData == Keys.Escape && painting != Painting.Off) { StopPainting(); SetStatus(T("status.painterOff", "Format Painter off.")); return true; }
		// The start page: a shortcut leaves it first (Ctrl+O, Ctrl+N ... then act on the document as usual)
		if (StartShown && startPage!.ContainsFocus && shortcuts.ContainsKey(keyData)) ShowStart(false);
		// Multiple Pages: its own keys; a shortcut that changes the text goes back to editing first
		if (PagesShown && pageView!.IsHandleCreated && msg.HWnd == pageView.Handle) {
			if (keyData is Keys.Escape or Keys.Enter) return base.ProcessCmdKey(ref msg, keyData);
			if (shortcuts.TryGetValue(keyData, out var pageShortcut) && !PageViewCommands.Any(c => CommandName(c) == pageShortcut.Name)) ShowPages(false);
		}
		// The source editor: its own editing keys, the document's (save, open, find ...); not the formatting ones, which
		// would act on the formatted text out of sight (or let RichEdit format the plain text, as Ctrl+E would)
		if (sourceBox.IsHandleCreated && msg.HWnd == sourceBox.Handle) {
			if (MacroForKeys(keyData) is { } sourceMacro) { Run(() => RunMacroFromUi(sourceMacro)); return true; }
			if ((keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab)) && InMarkdown && SourceTableAt(sourceBox.SelectionStart) != null) {
				RunSource(() => FormatSourceTable(keyData == Keys.Tab ? 1 : -1));
				return true;
			}
			if (sourceKeys.TryGetValue(keyData, out var markupKey)) { RunSource(markupKey); return true; }
			if (keyCommands.TryGetValue(keyData, out var keyCommand) && sourceCommands.TryGetValue(keyCommand, out var markupCommand)) { RunSource(markupCommand); return true; }
			if (SourceDocumentKeys.Contains(keyData) && shortcuts.TryGetValue(keyData, out var documentShortcut)) { documentShortcut.Action(); return true; }
			if (SourceEditKeys.Contains(keyData)) return base.ProcessCmdKey(ref msg, keyData);
			if (keyData == (Keys.Control | Keys.V) || keyData == (Keys.Shift | Keys.Insert)) return base.ProcessCmdKey(ref msg, keyData);
			if (shortcuts.ContainsKey(keyData) || SourceBox.FormattingKeys.Contains(keyData)) return true;
			return base.ProcessCmdKey(ref msg, keyData);
		}
		// Ctrl+Enter: a page break
		if (keyData == (Keys.Control | Keys.Enter) && Editor.IsHandleCreated && msg.HWnd == Editor.Handle) {
			Run(() => api.InsertPageBreak(), M("InsertPageBreak"));
			return true;
		}
		// Tables: row shortcuts, and Delete / Backspace on whole rows (RichEdit would only empty their cells)
		var rowKeys = Keys.Control | Keys.Shift;
		if (Editor.IsHandleCreated && msg.HWnd == Editor.Handle && (keyData == (rowKeys | Keys.Insert) || keyData == (rowKeys | Keys.Delete) || keyData is Keys.Delete or Keys.Back)) {
			if (keyData == (rowKeys | Keys.Insert) && api.InTable()) { Run(() => api.InsertRow(true), M("InsertRow", true)); return true; }
			if (keyData == (rowKeys | Keys.Delete) && api.InTable()) { Run(() => api.DeleteRow(), M("DeleteRow")); return true; }
			if (keyData is Keys.Delete or Keys.Back && api.SelectedRows() > 0) { Run(() => api.DeleteSelectedRows(), M("DeleteSelectedRows")); return true; }
		}
		// Legal lists: Enter, Backspace and Delete renumber in the same undo step (RichEdit's own would leave the numbers,
		// which are text, behind)
		if (Editor.IsHandleCreated && msg.HWnd == Editor.Handle && keyData is Keys.Enter or Keys.Back or Keys.Delete && api.TouchesLegalList()) {
			bool handled = false;
			Run(() => handled = keyData switch { Keys.Enter => api.LegalEnter(), Keys.Back => api.LegalBackspace(), _ => api.LegalDelete() },
				M("Key", keyData switch { Keys.Enter => "Enter", Keys.Back => "Backspace", _ => "Delete" }));
			if (handled) return true;
		}
		// Enter at the end of a paragraph whose style is followed by another (a heading by Normal): the new paragraph gets it
		if (keyData == Keys.Enter && Editor.IsHandleCreated && msg.HWnd == Editor.Handle && NextStyleAfterEnter() is { } nextStyle)
			BeginInvoke(() => ApplyStyleAtCaret(nextStyle));
		// (Enter in a style without space between its paragraphs: the two halves get it)
		else if (keyData == Keys.Enter && Editor.IsHandleCreated && msg.HWnd == Editor.Handle && Styles.Styles.Any(s => s.Contextual))
			BeginInvoke(() => { int caret = Editor.TextDocument.Selection2.Start; StyleEngine.Contextual(Editor.TextDocument, Styles, caret, caret); });
		// Tab / Shift+Tab in a table: the next or previous cell (RichEdit's own Tab doesn't in every table, such as
		// some from Word), and a new row from the last cell
		if ((keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab)) && Editor.IsHandleCreated && msg.HWnd == Editor.Handle && CaretInTable() && api.InTable()) {
			string key = keyData == Keys.Tab ? "Tab" : "Shift+Tab";
			Run(() => api.NextCell(keyData == Keys.Tab ? 1 : -1), M("Key", key));
			return true;
		}
		if ((keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab)) && Editor.IsHandleCreated && msg.HWnd == Editor.Handle && !api.InTable()
			&& (api.InList() || keyData == Keys.Tab && api.TabMakesSubLevel())) {    // (an unnumbered paragraph of a numbered style: a sub-level item)
			int delta = keyData == Keys.Tab ? 1 : -1;
			Run(() => api.Nest(delta), M("Nest", delta));
			return true;
		}
		// Macro hotkeys (Tools > Macro Hotkeys) come before the other shortcuts.
		if (MacroForKeys(keyData) is { } macro) {
			Run(() => RunMacroFromUi(macro));
			return true;
		}
		// The menus' shortcuts (Shortcut in BuildRibbon), except while typing in the ribbon (its font name and size boxes).
		if (shortcuts.TryGetValue(keyData, out var shortcut) && !ribbon.ContainsFocus) {
			shortcut.Action();
			return true;
		}
		return base.ProcessCmdKey(ref msg, keyData);
	}

	// ---------------------------------------------------------------- Macros (AHK scripts run as separate processes)

	string MacrosDir => Path.Combine(appDir, "rte_macros");
	internal string MacrosFolderPath => MacrosDir;

	void RunScriptDialog() {
		using var dlg = new OpenFileDialog { Title = T("file.runScript", "Run AHK script"), Filter = T("file.ahkScripts", "AutoHotkey scripts") + " (*.ahk)|*.ahk", InitialDirectory = Directory.Exists(MacrosDir) ? MacrosDir : appDir };
		if (dlg.ShowDialog(this) == DialogResult.OK) Run(() => RunMacroFromUi(dlg.FileName));
	}

	internal System.Diagnostics.Process? RunScript(string path, params string[] arguments) {
		if (ahkPath == "" || !File.Exists(ahkPath)) throw new InvalidOperationException("AutoHotkey was not found; start AhkWord with AhkWord.ahk.");
		var start = new System.Diagnostics.ProcessStartInfo(ahkPath) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(path)! };
		start.ArgumentList.Add(path);
		foreach (var argument in arguments) start.ArgumentList.Add(argument);
		start.Environment["RTE_WRITER_HWND"] = server.Handle.ToInt64().ToString(CultureInfo.InvariantCulture);
		// (iteration 56: a hook's event and what it was about, also as A_Args)
		if (arguments.Length > 0) start.Environment["AHKWORD_EVENT"] = arguments[0];
		if (arguments.Length > 1) start.Environment["AHKWORD_EVENT_DETAIL"] = arguments[1];
		var process = System.Diagnostics.Process.Start(start);
		SetStatus(T("status.running", "Running {0}", Path.GetFileName(path)));
		return process;
	}

	// ---------------------------------------------------------------- Iteration 56: event hooks (rte_hooks.cs)

	/// While the first document is made at the start (its hook is Started, not NewDocument).
	internal bool StartingUp;

	// the hooks' macros running now, by event (an event doesn't start its macro again while it runs)
	readonly Dictionary<HookEvent, System.Diagnostics.Process> hooksRunning = new();
	/// (tests) the hooks started, "Event: macro", and whether a waiting one timed out
	internal readonly List<string> HooksStarted = new();

	/// Runs the macro hooked to an event, if there is one (detail: the file, style, printer). BeforeSave and Closing wait
	/// for it to finish (at most EventHooks.WaitSeconds; meanwhile the window takes no input, but the macro's Rte calls
	/// are done); the others don't.
	internal void RunHook(HookEvent e, string detail = "") {
		if (EventHooks.MacroFor(Settings.EventHooks, e) is not { } macro || !EventHooks.IsOn(Settings.EventHooksOff, e)) return;
		RunHookMacro(e, macro, detail);
	}

	internal void RunHookMacro(HookEvent e, string macro, string detail) {
		if (hooksRunning.TryGetValue(e, out var running)) {
			if (!HasEnded(running)) return;    // (its own event, from the macro itself: not again)
			hooksRunning.Remove(e);
		}
		string path = EventHooks.PathOf(macro, MacrosDir);
		if (!File.Exists(path)) { SetStatus("⚠ " + T("status.hookMacroGone", "The {0} hook's macro, {1}, isn't there any more (Tools > Event Hooks).", e, macro)); return; }
		System.Diagnostics.Process? process;
		try { process = RunScript(path, e.ToString(), detail); }
		catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception) { SetStatus("⚠ " + T("status.hookFailed", "The {0} hook couldn't run: {1}", e, ex.Message)); return; }
		if (process == null) return;
		hooksRunning[e] = process;
		HooksStarted.Add($"{e}: {macro}");
		SetStatus(T("status.hookRunning", "{0}: running the {1} macro.", e, Path.GetFileNameWithoutExtension(macro)));
		if (!EventHooks.Waits(e)) return;
		// (waits: the macro's calls are answered meanwhile, the window takes no typing or clicks)
		bool enabled = Enabled, focused = Editor.Focused;
		Enabled = false;
		UseWaitCursor = true;
		var watch = System.Diagnostics.Stopwatch.StartNew();
		try {
			while (!HasEnded(process) && watch.Elapsed.TotalSeconds < EventHooks.WaitSeconds) {
				Application.DoEvents();
				process.WaitForExit(15);
			}
		} finally {
			UseWaitCursor = false;
			Enabled = enabled;
			if (focused && !IsDisposed) Editor.Focus();
		}
		if (!HasEnded(process)) {
			HooksStarted.Add($"{e}: {macro} (still running after {EventHooks.WaitSeconds} s)");
			SetStatus("⚠ " + T("status.hookStillRunning", "The {0} hook's macro, {1}, was still running after {2} seconds: AhkWord went on without waiting.", e, Path.GetFileNameWithoutExtension(macro), EventHooks.WaitSeconds));
		} else SetStatus(T("status.hookDone", "{0}: the {1} macro is done.", e, Path.GetFileNameWithoutExtension(macro)));
	}

	static bool HasEnded(System.Diagnostics.Process process) {
		try { return process.HasExited; } catch (InvalidOperationException) { return true; }
	}

	// Tools > Event Hooks
	void EventHooksCommand() {
		using var dlg = new EventHooksDialog(Settings.EventHooks, Settings.EventHooksOff, MacrosDir, (e, macro) => RunHookMacro(e, macro, e switch {
			HookEvent.StyleChanged => Styles.Normal.Name, HookEvent.Printed => "", HookEvent.Started or HookEvent.NewDocument => "", _ => FileName ?? ""
		}));
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Settings.EventHooks = dlg.Result;
		Settings.EventHooksOff = dlg.Off;
		Settings.Save();
		int count = Settings.EventHooks.Count;
		SetStatus(count == 0 ? T("status.noHooks", "No event hooks.") : TN("status.hooks", "{0} event hook: {1}.", "{0} event hooks: {1}.", count, string.Join(", ", Settings.EventHooks.Select(h => T("status.hookRuns", "{0} runs {1}", h.Key, Path.GetFileNameWithoutExtension(h.Value))
			+ (EventHooks.IsOn(Settings.EventHooksOff, EventHooks.Parse(h.Key)) ? "" : " " + T("status.hookOff", "(off)"))))));
	}

	// ---------------------------------------------------------------- Iteration 6: recording macros

	internal bool IsRecording => recorder.Recording;

	// Records a call while a macro is being recorded (from menus, toolbar and dialogs; automation calls are not recorded).
	void Record(string call) {
		if (!recorder.Recording) return;
		recorder.Add(call);
		UpdateRecordingLabel();
	}

	// Runs a macro chosen in the UI (Macros menu, hotkey, Run AHK script). While recording, this is recorded as a
	// RunWait of that script: a replay waits for it to finish before going on, so its output lands in the right place
	// (macros are separate processes and would otherwise run alongside each other).
	void RunMacroFromUi(string path) {
		RunScript(path);
		if (!recorder.Recording) return;
		static string Quoted(string s) => s.Replace("`", "``").Replace("'", "`'");
		string name = Path.GetFileNameWithoutExtension(path);
		string script = string.Equals(Path.GetDirectoryName(Path.GetFullPath(path)), Path.GetFullPath(MacrosDir), StringComparison.OrdinalIgnoreCase)
			? $"' A_ScriptDir '\\{Quoted(Path.GetFileName(path))}"    // recorded macros are saved in rte_macros too
			: Quoted(Path.GetFullPath(path));
		recorder.AddLine($"RunWait('\"' A_AhkPath '\" \"{script}\"')    ; runs the \"{name}\" macro and waits for it to finish");
		UpdateRecordingLabel();
		// Macros that react to events (Rte.On) keep running until the editor closes, so a replay would wait for ever.
		bool keepsRunning = File.ReadAllText(path).Contains("Rte.On(", StringComparison.OrdinalIgnoreCase);
		SetStatus(keepsRunning
			? "⚠ " + T("status.recordedRunKeeps", "Recorded: run the \"{0}\" macro - but it keeps running, so a replay would wait for it forever. Remove its RunWait line when you save.", name)
			: T("status.recordedRun", "Recorded: run the \"{0}\" macro. A replay waits for it to finish before going on.", name));
	}

	static readonly Dictionary<Keys, string> RecordedKeys = new() {
		[Keys.Left] = "Left", [Keys.Right] = "Right", [Keys.Up] = "Up", [Keys.Down] = "Down",
		[Keys.Home] = "Home", [Keys.End] = "End", [Keys.Back] = "Backspace", [Keys.Delete] = "Delete"
	};

	// Keys that reach the editor (menu shortcuts and Tab in lists are recorded as commands, before this).
	void RecordKeyDown(KeyEventArgs e) {
		if (e.KeyCode == Keys.Enter && !e.Control && !e.Alt) {
			recorder.Type('\n');
		} else if (RecordedKeys.TryGetValue(e.KeyCode, out var name) && !e.Alt) {
			// Backspace right after typing just takes back the last typed character.
			if (e.KeyCode == Keys.Back && !e.Control && Editor.SelectionLength == 0 && recorder.Backspace()) { UpdateRecordingLabel(); return; }
			recorder.AddKey((e.Control ? "Ctrl+" : "") + (e.Shift ? "Shift+" : "") + name);
		} else {
			return;
		}
		UpdateRecordingLabel();
	}

	void UpdateRecordingLabel() {
		recordingLabel.Visible = recorder.Recording;
		recordingLabel.Text = "● " + TN("ui.recordingSteps", "Recording ({0} step)", "Recording ({0} steps)", recorder.Steps);
		if (toggleButtons.TryGetValue(RibbonCommand.RecordMacro, out var record) && ribbon.Initialized) record.BooleanValue = recorder.Recording;
	}

	internal void StartRecording() {
		recorder.Start();
		UpdateRecordingLabel();
		SetStatus(T("status.recording", "Recording: type and use the menus and toolbar. Mouse clicks are not recorded. Ctrl+Shift+R stops."));
	}

	internal string StopRecording() {
		var lines = recorder.Stop();
		UpdateRecordingLabel();
		return lines.Count == 0 ? "" : MacroScript(lines);
	}

	void ToggleRecording() {
		if (!recorder.Recording) {
			StartRecording();
			return;
		}
		string script = StopRecording();
		if (script == "") {
			SetStatus(T("status.nothingRecorded", "Nothing was recorded."));
			return;
		}
		using var dlg = new SaveMacroDialog(UniqueMacroName(), script, MenuShortcuts(), this);
		if (dlg.ShowDialog(this) != DialogResult.OK) {
			SetStatus(T("status.recordingDiscarded", "The recording was discarded."));
			return;
		}
		Directory.CreateDirectory(MacrosDir);
		string path = Path.Combine(MacrosDir, dlg.MacroName + ".ahk");
		if (File.Exists(path) && MessageBox.Show(this, T("msg.replaceMacro", "Replace the existing macro \"{0}\"?", dlg.MacroName), AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
			SetStatus(T("status.recordingNotSaved", "The recording was not saved."));
			return;
		}
		File.WriteAllText(path, dlg.Script.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
		SetMacroHotkey(Path.GetFileName(path), dlg.Hotkey);
		FillMacrosGallery();
		SetStatus(dlg.Hotkey != Keys.None ? T("status.savedMacroHotkey", "Saved macro \"{0}\" ({1}). It is in Tools > Macros.", dlg.MacroName, TKeys(HotkeyBox.Describe(dlg.Hotkey)))
			: T("status.savedMacro", "Saved macro \"{0}\". It is in Tools > Macros.", dlg.MacroName));
	}

	static string MacroScript(List<string> lines) =>
		"#Requires AutoHotkey v2.0\r\n#Include \"..\\RteClient2.ahk\"\r\n"
		+ $"; Recorded in AhkWord on {DateTime.Now:g}. Each line is an AhkWord call (Tools > Automation API lists them).\r\n"
		+ string.Join("\r\n", lines) + "\r\n";

	string UniqueMacroName() {
		for (int i = 1; ; i++) {
			string name = $"Recorded macro {i}";
			if (!File.Exists(Path.Combine(MacrosDir, name + ".ahk"))) return name;
		}
	}

	// ---------------------------------------------------------------- Iteration 6: macro hotkeys

	internal Keys? HotkeyOf(string macroFile) =>
		Settings.MacroHotkeys.TryGetValue(macroFile, out var text) && HotkeyBox.Parse(text) is Keys keys && keys != Keys.None ? keys : null;

	internal void SetMacroHotkey(string macroFile, Keys keys) {
		// one macro per hotkey
		foreach (var other in Settings.MacroHotkeys.Where(p => HotkeyBox.Parse(p.Value) == keys).Select(p => p.Key).ToList()) Settings.MacroHotkeys.Remove(other);
		if (keys == Keys.None) Settings.MacroHotkeys.Remove(macroFile);
		else Settings.MacroHotkeys[macroFile] = HotkeyBox.Describe(keys);
		Settings.Save();
	}

	string? MacroForKeys(Keys keys) {
		if (keys == Keys.None || Settings.MacroHotkeys.Count == 0) return null;
		foreach (var (file, text) in Settings.MacroHotkeys)
			if (HotkeyBox.Parse(text) == keys && File.Exists(Path.Combine(MacrosDir, file))) return Path.Combine(MacrosDir, file);
		return null;
	}

	internal string MacroPath(string name) {
		string file = name.EndsWith(".ahk", StringComparison.OrdinalIgnoreCase) ? name : name + ".ahk";
		string path = Path.Combine(MacrosDir, file);
		return File.Exists(path) ? path : throw new FileNotFoundException($"There is no macro \"{name}\" in {MacrosDir}.");
	}

	internal List<string> MacroFiles() =>
		Directory.Exists(MacrosDir) ? Directory.GetFiles(MacrosDir, "*.ahk").Select(Path.GetFileName).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList()! : new List<string>();

	// The editor's keyboard shortcuts, to warn about hotkeys that would hide them.
	internal Dictionary<Keys, string> MenuShortcuts() {
		var names = shortcuts.ToDictionary(p => p.Key, p => p.Value.Name);
		names.TryAdd(Keys.Control | Keys.Enter, "Page Break");
		names.TryAdd(Keys.Tab, "Tab");
		names.TryAdd(Keys.Shift | Keys.Tab, "Shift+Tab");
		return names;
	}

	void MacroHotkeysCommand() {
		var files = MacroFiles();
		if (files.Count == 0) throw new InvalidOperationException("There are no macros in " + MacrosDir + " yet.");
		using var dlg = new MacroHotkeysDialog(files, f => HotkeyOf(f) ?? Keys.None, MenuShortcuts());
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		Settings.MacroHotkeys.Clear();
		foreach (var (file, keys) in dlg.Hotkeys) if (keys != Keys.None) Settings.MacroHotkeys[file] = HotkeyBox.Describe(keys);
		Settings.Save();
		FillMacrosGallery();
		SetStatus(TN("status.macroHotkeys", "{0} macro hotkey set. They work while the editor is active.", "{0} macro hotkeys set. They work while the editor is active.", Settings.MacroHotkeys.Count));
	}

	void OpenMacrosFolder() {
		Directory.CreateDirectory(MacrosDir);
		System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + MacrosDir + "\""));
	}

	// ---------------------------------------------------------------- Automation entry point

	string DispatchFromAutomation(string[] request) {
		if (request.Length > 1 && request[0] == "post") {
			postedCalls.Enqueue(request[1..]);
			if (postedCalls.Count == 1) BeginInvoke(RunPostedCalls);
			return "0";
		}
		RunPostedCalls();    // calls run in the order they were made, posted or not
		Restart(outlineTimer);
		string reply;
		try {
			if (request.Length < 2 || !request[1].Equals("SourceText", StringComparison.OrdinalIgnoreCase)) ApplySource();
			inCommand = true;
			reply = Automation.Dispatch(api, Editor, request);
		} finally {
			inCommand = false;
		}
		RefreshSource();
		if (request.Length > 1) automationLabel.Text = T("ui.automation", "Automation: {0}", request[1]);
		UpdateUi();
		UpdateTitle();
		return reply;
	}

	void RunPostedCalls() {
		if (postedCalls.Count == 0) return;
		while (postedCalls.TryDequeue(out var call)) {
			string reply = Automation.Dispatch(api, Editor, new[] { "call" }.Concat(call).ToArray());
			if (reply.StartsWith('1')) {
				api.LastAsyncError = $"{call[0]}: {reply[1..]}";
				SetStatus("⚠ " + api.LastAsyncError);
				Events.Raise(EditorEvent.AsyncError);
			}
			automationLabel.Text = T("ui.automationPosted", "Automation: {0} (posted)", call[0]);
		}
		UpdateUi();
		UpdateTitle();
	}

	// ---------------------------------------------------------------- UI state

	void UpdateTitle() {
		string title = $"{DocumentName}{(Editor.Modified ? " *" : "")} - {AppName}";
		if (Text != title) Text = title;
		if (Editor.Modified != lastModified) {
			lastModified = Editor.Modified;
			Events.Raise(EditorEvent.ModifiedChanged, lastModified ? 1 : 0);
		}
	}

	void UpdateWordCount() => wordsLabel.Text = TN("ui.words", "{0} words", "{0} words", api.WordCount());    // (English: "1 words", as before)

	void UpdateUi() {
		if (!Editor.IsHandleCreated || updatingUi) return;
		updatingUi = true;
		try {
			var range = api.SelectionRange();
			if (ribbon.Initialized) UpdateRibbon(range);
			if (ruler.Visible != (Settings.ShowRuler && !PagesShown)) ruler.Visible = Settings.ShowRuler && !PagesShown;    // (not over Multiple Pages)
			FollowRuler(true);

			int line, column;
			if (SourceActive) {
				line = sourceBox.LineOf(sourceBox.SelectionStart);
				column = sourceBox.SelectionStart - sourceBox.LineStart(line);
			} else {
				line = Editor.GetLineFromCharIndex(Editor.SelectionStart);
				column = Editor.SelectionStart - Editor.GetFirstCharIndexFromLine(line);
			}
			positionLabel.Text = T("ui.position", "Ln {0}, Col {1}", line + 1, column + 1);
			modeLabel.Text = view != SourceKind.None ? ViewName(view) + (sideBySide ? " | " + T("ui.sideBySide", "side by side") : "") : "";
		} catch (COMException) {
			// The document can be briefly unavailable while it is being replaced.
		} finally {
			updatingUi = false;
		}
	}

	// The ribbon shows the selection's format (properties that differ within it are blank, as in Word) and the view's
	// settings; Table Tools appears while the caret is in a table.
	// Iteration 48: what was last sent to the ribbon, each control's property: a value is sent again only when it changes.
	// UpdateRibbon runs at every caret move and sent some 45 values each time, nearly all unchanged (the font group's
	// each made the ribbon read all its properties again); the ribbon queues them for its timer, and its timer has been
	// seen to crash (UIRibbon.dll, an access violation) after the window was busy for a while: the fewer queued, the
	// fewer chances. Forgotten when a ribbon command runs (Command): a click changes a toggle's state in the ribbon itself.
	readonly Dictionary<(object Control, string Property), object?> ribbonSent = new();
	void Push<T>(object control, string property, T value, Action<T> send) {
		if (ribbonSent.TryGetValue((control, property), out var sent) && Equals(sent, value)) return;
		ribbonSent[(control, property)] = value;
		send(value);
	}

	void UpdateRibbon(ITextRange2 range) {
		var font = api.SelectionFont();    // (with nothing selected: what is typed next, as Word shows it)
		FontProperties State(int value) => value == tomConstants.tomTrue ? FontProperties.Set : value == tomConstants.tomFalse ? FontProperties.NotSet : FontProperties.NotAvailable;
		if (DarkRibbon) ShowDarkFont(font);
		else {
		Push(fontControl, "Family", font.Name ?? "", v => fontControl.Family = v);
		Push(fontControl, "Size", font.Size > 0 ? (decimal)Math.Round(font.Size, 1) : 0, v => fontControl.Size = v);
		Push(fontControl, "Bold", State(font.Bold), v => fontControl.Bold = v);
		Push(fontControl, "Italic", State(font.Italic), v => fontControl.Italic = v);
		Push(fontControl, "Underline", font.Underline == tomConstants.tomUndefined ? FontUnderline.NotAvailable : font.Underline > 0 ? FontUnderline.Set : FontUnderline.NotSet, v => fontControl.Underline = v);
		Push(fontControl, "Strikethrough", State(font.StrikeThrough), v => fontControl.Strikethrough = v);
		Push(fontControl, "VerticalPositioning", font.Superscript == tomConstants.tomTrue ? FontVerticalPosition.SuperScript
			: font.Subscript == tomConstants.tomTrue ? FontVerticalPosition.SubScript
			: font.Superscript == tomConstants.tomUndefined || font.Subscript == tomConstants.tomUndefined ? FontVerticalPosition.NotAvailable : FontVerticalPosition.NotSet, v => fontControl.VerticalPositioning = v);
		if (font.ForeColor >= 0) Push(fontControl, "ForegroundColor", ColorTranslator.FromWin32(font.ForeColor), v => fontControl.ForegroundColor = v);
		}

		var para = range.Para;
		int listType = para.ListType;
		bool legal = listType > 0 && (listType & 0x40000) != 0;
		int style = listType > 0 ? listType & 0xFFFF : listType;
		Push(toggleButtons[RibbonCommand.Bullets], "BooleanValue", style == (int)ListStyle.Bullet, v => toggleButtons[RibbonCommand.Bullets].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.Numbering], "BooleanValue", style > 0 && style != (int)ListStyle.Bullet && !legal, v => toggleButtons[RibbonCommand.Numbering].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.Legal], "BooleanValue", legal, v => toggleButtons[RibbonCommand.Legal].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.AlignLeft], "BooleanValue", para.Alignment == (int)ParagraphAlignment.Left, v => toggleButtons[RibbonCommand.AlignLeft].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.AlignCenter], "BooleanValue", para.Alignment == (int)ParagraphAlignment.Center, v => toggleButtons[RibbonCommand.AlignCenter].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.AlignRight], "BooleanValue", para.Alignment == (int)ParagraphAlignment.Right, v => toggleButtons[RibbonCommand.AlignRight].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.Justify], "BooleanValue", para.Alignment == (int)ParagraphAlignment.Justify, v => toggleButtons[RibbonCommand.Justify].BooleanValue = v);
		string spacing = LineSpacing.Of(range);
		foreach (var (id, value) in new[] { (RibbonCommand.LineSpacing10, "1"), (RibbonCommand.LineSpacing115, "1.15"), (RibbonCommand.LineSpacing150, "1.5"),
			(RibbonCommand.LineSpacing200, "2"), (RibbonCommand.LineSpacing250, "2.5"), (RibbonCommand.LineSpacing300, "3") })
			Push(toggleButtons[id], "BooleanValue", spacing == value, v => toggleButtons[id].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.FormatPainter], "BooleanValue", painting != Painting.Off, v => toggleButtons[RibbonCommand.FormatPainter].BooleanValue = v);
		// (page 1's top margin, the page view's space before the first paragraph, isn't the paragraph's)
		float before = LineSpacing.SpaceOf(range, true);
		if (pageGaps.Applied) before -= pageGaps.ExtraAt(range.Start);    // (a page's first paragraph: the gap isn't its own space)
		Push(toggleButtons[RibbonCommand.SpaceBefore], "BooleanValue", before > 0.05f, v => toggleButtons[RibbonCommand.SpaceBefore].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.SpaceAfter], "BooleanValue", LineSpacing.SpaceOf(range, false) > 0, v => toggleButtons[RibbonCommand.SpaceAfter].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.LineNumbersNone], "BooleanValue", !Numbering.On, v => toggleButtons[RibbonCommand.LineNumbersNone].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.LineNumbersContinuous], "BooleanValue", Numbering.Mode == LineNumbering.Modes.Continuous, v => toggleButtons[RibbonCommand.LineNumbersContinuous].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.LineNumbersEachPage], "BooleanValue", Numbering.Mode == LineNumbering.Modes.EachPage, v => toggleButtons[RibbonCommand.LineNumbersEachPage].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.SuppressLineNumbers], "BooleanValue", para.NoLineNumber == tomConstants.tomTrue, v => toggleButtons[RibbonCommand.SuppressLineNumbers].BooleanValue = v);
		FillStyleBox();
		int styleIndex = styleBoxNames.FindIndex(n => string.Equals(n, CaretStyleName(), StringComparison.OrdinalIgnoreCase));
		Push(blockCombo, "SelectedItem", styleIndex >= 0 ? (uint)styleIndex : Constants.UI_Collection_InvalidIndex, v => blockCombo.SelectedItem = v);
		foreach (var (id, set) in new[] { (RibbonCommand.StyleSetClassic, "Classic"), (RibbonCommand.StyleSetPlain, "Plain"), (RibbonCommand.StyleSetModern, "Modern") })
			Push(toggleButtons[id], "BooleanValue", string.Equals(Styles.Set, set, StringComparison.OrdinalIgnoreCase), v => toggleButtons[id].BooleanValue = v);
		ShowStyleSet();
		Push(toggleButtons[RibbonCommand.RecordMacro], "BooleanValue", recorder.Recording, v => toggleButtons[RibbonCommand.RecordMacro].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.ViewRich], "BooleanValue", view == SourceKind.None, v => toggleButtons[RibbonCommand.ViewRich].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.ViewMarkdown], "BooleanValue", view == SourceKind.Markdown, v => toggleButtons[RibbonCommand.ViewMarkdown].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.ViewHtml], "BooleanValue", view == SourceKind.Html, v => toggleButtons[RibbonCommand.ViewHtml].BooleanValue = v);
		Push(toggleButtons[RibbonCommand.SideBySide], "BooleanValue", sideBySide, v => toggleButtons[RibbonCommand.SideBySide].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.SourceWordWrap], "BooleanValue", sourceBox.WordWrap, v => checkBoxes[RibbonCommand.SourceWordWrap].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.SourceLineNumbers], "BooleanValue", sourceBox.LineNumbers, v => checkBoxes[RibbonCommand.SourceLineNumbers].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.WordWrap], "BooleanValue", Editor.WordWrap, v => checkBoxes[RibbonCommand.WordWrap].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.WordWrap], "Enabled", !Settings.PrintLayout, v => checkBoxes[RibbonCommand.WordWrap].Enabled = v);    // (Print Layout breaks lines as on paper)
		Push(checkBoxes[RibbonCommand.PrintLayout], "BooleanValue", Settings.PrintLayout, v => checkBoxes[RibbonCommand.PrintLayout].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.ShowPages], "BooleanValue", Settings.PageView, v => checkBoxes[RibbonCommand.ShowPages].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.ShowPages], "Enabled", Settings.PrintLayout, v => checkBoxes[RibbonCommand.ShowPages].Enabled = v);
		// (Page Width, One Page and Multiple Pages: one of them at a time; each pushed to the ribbon even when it seems to
		// have it already, as a click checks a button in the ribbon before the command has run)
		void Show(RibbonCommand id, bool on) { var button = toggleButtons[id]; if (zoomClicked && button.BooleanValue == on) button.BooleanValue = !on; button.BooleanValue = on; }
		Show(RibbonCommand.ZoomPageWidth, Settings.PrintLayout && Settings.ZoomMode == "PageWidth" && !PagesShown);
		Show(RibbonCommand.ZoomOnePage, Settings.PrintLayout && Settings.ZoomMode == "OnePage" && !PagesShown);
		Push(toggleButtons[RibbonCommand.ZoomPageWidth], "Enabled", toggleButtons[RibbonCommand.ZoomOnePage].Enabled = Settings.PrintLayout, v => toggleButtons[RibbonCommand.ZoomPageWidth].Enabled = v);
		Show(RibbonCommand.ZoomMultiplePages, PagesShown);
		zoomClicked = false;
		Push(checkBoxes[RibbonCommand.ShowBands], "BooleanValue", Settings.ShowBands, v => checkBoxes[RibbonCommand.ShowBands].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.Outline], "BooleanValue", !outlineSplit.Panel1Collapsed, v => checkBoxes[RibbonCommand.Outline].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.Ruler], "BooleanValue", Settings.ShowRuler, v => checkBoxes[RibbonCommand.Ruler].BooleanValue = v);
		Push(checkBoxes[RibbonCommand.SpellCheck], "BooleanValue", Editor.EnableSpellCheck, v => checkBoxes[RibbonCommand.SpellCheck].BooleanValue = v);
		Push(undoButton, "Enabled", Editor.CanUndo, v => undoButton.Enabled = v);
		Push(redoButton, "Enabled", Editor.CanRedo, v => redoButton.Enabled = v);
		Push(cutButton, "Enabled", Editor.SelectionLength > 0, v => cutButton.Enabled = v);
		Push(copyButton, "Enabled", Editor.SelectionLength > 0, v => copyButton.Enabled = v);
		var tools = CaretInTable() ? ContextAvailability.Available : ContextAvailability.NotAvailable;
		if (tableTools.ContextAvailable != tools) tableTools.ContextAvailable = tools;
	}

	// ---------------------------------------------------------------- Iteration 17: Markdown and HTML views

	internal SourceKind View => view;
	internal bool SideBySide => sideBySide;

	static string ViewName(SourceKind kind) => kind switch { SourceKind.Markdown => "Markdown", SourceKind.Html => "HTML", _ => T("ui.richText", "Rich text") };

	// The source editor is the one in use: it has the focus, or had it last (a ribbon click doesn't take it away).
	bool SourceActive => view != SourceKind.None && (sourceBox.Focused || sourceHadFocus && !Editor.Focused);

	// Commands that work the same in the source editor (the document, the view, editing the source itself).
	static readonly HashSet<RibbonCommand> SourceDocumentCommands = new() {
		RibbonCommand.New, RibbonCommand.NewFromTemplate, RibbonCommand.StartPage, RibbonCommand.NewWindow, RibbonCommand.OpenInNewWindow, RibbonCommand.Open, RibbonCommand.Save, RibbonCommand.SaveAs, RibbonCommand.Close, RibbonCommand.Print, RibbonCommand.PrintPreview,
		RibbonCommand.PageSetup, RibbonCommand.ExportPdf, RibbonCommand.ExportPdfButton, RibbonCommand.ExportWord, RibbonCommand.ExportOdt, RibbonCommand.Exit,
		RibbonCommand.Help, RibbonCommand.About, RibbonCommand.AutomationApi, RibbonCommand.RecordMacro, RibbonCommand.MacroHotkeys, RibbonCommand.EventHooks, RibbonCommand.CommandPalette,
		RibbonCommand.RunScript, RibbonCommand.MacrosFolder, RibbonCommand.Snippets, RibbonCommand.ViewRich, RibbonCommand.ViewMarkdown, RibbonCommand.ViewHtml,
		RibbonCommand.SideBySide, RibbonCommand.SourceWordWrap, RibbonCommand.SourceLineNumbers, RibbonCommand.WordWrap, RibbonCommand.PrintLayout, RibbonCommand.ZoomPageWidth, RibbonCommand.ZoomOnePage,
		RibbonCommand.ZoomMultiplePages, RibbonCommand.ShowBands, RibbonCommand.Outline, RibbonCommand.Ruler, RibbonCommand.ZoomIn, RibbonCommand.ZoomOut, RibbonCommand.Zoom100,
		RibbonCommand.HeaderFooter, RibbonCommand.Columns, RibbonCommand.LineNumbersNone, RibbonCommand.LineNumbersContinuous, RibbonCommand.LineNumbersEachPage,
		RibbonCommand.LineNumberOptions, RibbonCommand.Undo, RibbonCommand.Redo, RibbonCommand.Cut, RibbonCommand.Copy, RibbonCommand.Paste,
		RibbonCommand.SelectAll, RibbonCommand.Find, RibbonCommand.Replace, RibbonCommand.Theme, RibbonCommand.Language,
	};

	// Keys the source editor handles itself (RichEdit's own editing), and the document's commands it lets through.
	static readonly HashSet<Keys> SourceEditKeys = new() {
		Keys.Control | Keys.Z, Keys.Control | Keys.Y, Keys.Control | Keys.X, Keys.Control | Keys.C, Keys.Control | Keys.A,
	};
	static readonly HashSet<Keys> SourceDocumentKeys = new() {
		Keys.Control | Keys.N, Keys.Control | Keys.O, Keys.Control | Keys.S, Keys.Control | Keys.Shift | Keys.S, Keys.Control | Keys.P, Keys.Control | Keys.W,
		Keys.Control | Keys.F, Keys.Control | Keys.H, Keys.F3, Keys.Control | Keys.Shift | Keys.O,
	};

	/// Shows the formatted text (None), or the document as Markdown or HTML, alone or side by side with the formatted text.
	internal void SetView(SourceKind kind, bool side) {
		ApplySource();
		side = side && kind != SourceKind.None;
		bool wasHidden = view == SourceKind.None;
		view = kind;
		sideBySide = side;
		if (kind != SourceKind.None && (wasHidden || shownKind != kind || !SourceCurrent(kind))) FillSource();
		if (kind == SourceKind.None) {
			sourceSplit.Panel2Collapsed = false;
			sourceSplit.Panel1Collapsed = true;
		} else {
			sourceSplit.Panel1Collapsed = false;
			sourceSplit.Panel2Collapsed = !side;
			if (side) sourceSplit.SplitterDistance = Math.Clamp(sourceSplit.Width * Settings.SourceSplit / 100, 50, Math.Max(50, sourceSplit.Width - 50));
		}
		sourceTitle.Text = kind == SourceKind.None ? "" : (kind == SourceKind.Markdown
			? (side ? T("ui.sourceMarkdownSide", "Markdown: the formatted text follows as you type") : T("ui.sourceMarkdown", "Markdown: View > Rich Text shows it formatted"))
			: side ? T("ui.sourceHtmlSide", "HTML: the formatted text follows as you type; tags and styles it doesn't support are dropped") : ViewName(kind))
			+ (embeddedData.Count > 0 ? ". " + T("ui.sourcePictures", "Pictures appear as embedded-1.png ...") : "");
		if (kind == SourceKind.None) Editor.Focus(); else sourceBox.Focus();
		UpdateUi();
	}

	// Fills the source editor from the document: with the source text it came from, if it is unchanged since; else
	// regenerated (pictures as embedded-1.png ...).
	void FillSource() {
		var kind = view;
		string regenerated = api.SourceOf(kind);
		if (!sources.TryGetValue(kind, out var source) || source.Snapshot != regenerated)
			sources[kind] = source = (Embedded(regenerated).Replace("\r\n", "\n"), regenerated, false);
		filling = true;
		try {
			sourceBox.Language = kind == SourceKind.Markdown ? SourceBox.Languages.Markdown : SourceBox.Languages.Html;
			sourceBox.Text = source.Text;
			sourceBox.ClearUndo();
			sourceBox.Select(0, 0);
		} finally {
			filling = false;
		}
		shownKind = kind;
		sourceDirty = false;
	}

	// The source again after the formatted text changed, unless edits typed in the source are waiting.
	void RefreshSource() {
		if (view == SourceKind.None || sourceDirty || loadingSource) return;
		if (shownKind != view || !SourceCurrent(view)) {
			int caret = sourceBox.SelectionStart;
			FillSource();
			sourceBox.Select(Math.Min(caret, sourceBox.TextLength), 0);
			sourceBox.ScrollToCaret();
		}
	}

	void SourceTextChanged() {
		if (filling) return;
		sourceDirty = true;
		Editor.Modified = true;
		UpdateTitle();
		if (sideBySide) Restart(sourceTimer);
	}

	/// Loads what was typed in the source view into the document (if anything was).
	internal void ApplySource() {
		sourceTimer.Stop();
		if (!sourceDirty || loadingSource) return;
		string text = sourceBox.Text;
		double where = sourceBox.TextLength > 0 ? (double)sourceBox.SelectionStart / sourceBox.TextLength : 0;
		bool hadFocus = sourceBox.Focused;
		loadingSource = true;
		try {
			api.LoadSource(shownKind, WithPictures(text));
		} finally {
			loadingSource = false;
		}
		sources[shownKind] = (text, api.SourceOf(shownKind), true);
		sourceDirty = false;
		Editor.Modified = true;
		// the formatted text shows about where the caret is in the source
		Editor.Select((int)(where * Editor.TextLength), 0);
		Editor.ScrollToCaret();
		if (hadFocus) sourceBox.Focus();
		UpdateTitle();
		Restart(outlineTimer);
		Restart(wordCountTimer);
	}

	/// Before New or Open puts its document in: the editor emptied into a new RichEdit window (WriterEditBox.Renew: the
	/// ranges UI Automation clients left in the old one slowed every edit), what was kept of the old one let go of first.
	internal void FreshEditor() {
		if (!Editor.IsHandleCreated || StartingUp) return;
		pageGaps.Forget();    // (their ranges: the old window's; the document goes anyway)
		pageRanges.Clear();
		bandStarts = null;
		Editor.PageStarts = null;
		Editor.Renew();
		RegisterDropTarget();
		ApplyPrintLayout();
	}

	/// After New or Open: the source the document came from (a .md or .html file's text), if any.
	internal void DocumentLoaded(SourceKind kind, string? text) {
		ClearRecovery();    // (the document before it went, as the user wanted)
		bandStarts = null;    // (the pages before were the document before's)
		pageGaps.Forget();
		bandPagesStale = true;
		Editor.PageStarts = null;
		sourceTimer.Stop();
		richToSourceTimer.Stop();
		embeddedData.Clear();
		embeddedName.Clear();
		sourceDirty = false;
		sources.Clear();
		// (a file's own pictures as data: URIs show as names too, and go back when it is saved)
		if (text != null && kind != SourceKind.None) sources[kind] = (Embedded(text).Replace("\r\n", "\n"), api.SourceOf(kind), true);
		sourceNewline = text != null && text.Contains("\r\n") ? "\r\n" : kind == SourceKind.Html ? "\r\n" : "\n";
		if (view != SourceKind.None) FillSource();
		if (IsHandleCreated && CrossRefs.Any(Editor.Text)) {
			referencesPending = true;
			BeginInvoke(ReferencesNow);    // (opened in the page view: already done, SheetsNow)
		}
	}

	bool referencesPending;
	void ReferencesNow() {
		if (!referencesPending) return;
		referencesPending = false;
		QuietReferenceUpdate();
	}

	// (iteration 50) Once a document is loaded: its cross-references' numbers as the editor numbers the paragraphs (a
	// file from Word or LibreOffice has their numbers; RichEdit gives a list its numbers only once it has laid it out),
	// as part of the loading: no undo step, nothing to save.
	// (iteration 52) And its tables of contents' page numbers as AhkWord prints the pages (a file's are as another program
	// laid them out, or, from Word's own converting, all 1).
	void QuietReferenceUpdate() {
		if (!Editor.IsHandleCreated || Editor.CanUndo) return;    // (something done since: F9, or saving, updates them)
		var doc = Editor.TextDocument;
		bool modified = Editor.Modified;
		bool contents = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text?.Contains("[[toc=", StringComparison.Ordinal) == true;
		int styles = Styles.Styles.Count;
		doc.Undo(-9999995);    // tomSuspend
		try {
			Footnotes.Update(doc);
			CrossRefs.Update(doc);
			// (the tables' look and tabs before their pages are laid out: they don't depend on the pages, and the pages are
			// then as the sheets that follow would lay them out, which take them as they are when nothing else changed)
			if (contents && Toc.Restyle(doc, Styles, TextWidthPoints()) + Toc.FitTabs(doc, TextWidthPoints()) > 0) bandPagesStale = true;
		}
		finally { doc.Undo(-9999994); Editor.Modified = modified; }    // tomResume
		if (contents) {
			var pageOf = PageOf();    // (out of the suspended undo: laying them out takes the page view's gaps out and puts them back)
			doc.Undo(-9999995);
			try { if (Toc.UpdatePages(doc, pageOf) > 0) { bandPagesStale = true; measuredPages = null; } }    // (the numbers' new widths may move the pages)
			finally { doc.Undo(-9999994); Editor.Modified = modified; }
		}
		// (the TOC's styles a file without them was given: part of the loading too, so the document isn't changed by them;
		// a .md file asked to be saved, unchanged, before 2026-11-08)
		if (Styles.Styles.Count != styles) BeginInvoke(() => { bool before = Editor.Modified; StylesChanged(); Editor.Modified = before; });
	}

	/// The text to save a .md / .html file with, if the document came from such source text (the file's own, or typed in
	/// the source view) and is unchanged since (null: save the document as usual). Pictures shown as names are data:
	/// URIs again. Waiting source edits are loaded first.
	internal string? VerbatimSource(string extension) {
		ApplySource();
		var kind = extension is ".md" or ".markdown" ? SourceKind.Markdown : extension is ".html" or ".htm" ? SourceKind.Html : SourceKind.None;
		if (kind == SourceKind.None || !sources.TryGetValue(kind, out var source) || !source.Own) return null;
		if (api.SourceOf(kind) != source.Snapshot) return null;
		string text = WithPictures(source.Text);
		return sourceNewline == "\n" ? text : text.Replace("\n", sourceNewline);
	}

	// Find in the source view: from the selection on, wrapping round once.
	bool FindInSource() {
		var options = (findForm!.MatchCase ? RichTextBoxFinds.MatchCase : RichTextBoxFinds.None) | (findForm.WholeWord ? RichTextBoxFinds.WholeWord : RichTextBoxFinds.None);
		int start = Math.Min(sourceBox.SelectionStart + sourceBox.SelectionLength, sourceBox.TextLength);
		int index = sourceBox.Find(findForm.FindText, start, sourceBox.TextLength, options);
		if (index < 0 && start > 0) index = sourceBox.Find(findForm.FindText, 0, sourceBox.TextLength, options);
		if (index >= 0) sourceBox.ScrollToCaret();
		SetStatus(index < 0 ? T("status.notFound", "\"{0}\" was not found.", findForm.FindText) : "");
		return index >= 0;
	}

	// Whether the remembered source text of this kind still matches the document.
	bool SourceCurrent(SourceKind kind) => sources.TryGetValue(kind, out var source) && api.SourceOf(kind) == source.Snapshot;

	/// The text in the source view ("" when the formatted text is shown).
	internal string SourceText => view == SourceKind.None ? "" : sourceBox.Text;

	/// Replaces the text in the source view and loads it into the document.
	internal void SetSourceText(string text) {
		if (view == SourceKind.None) throw new InvalidOperationException("Show the Markdown or HTML view first (View = \"Markdown\").");
		sourceBox.Text = text.Replace("\r\n", "\n");
		ApplySource();
	}

	internal static readonly Regex DataUri = new(@"data:image/(?<type>[a-zA-Z0-9.+-]+);base64,[A-Za-z0-9+/=]+", RegexOptions.Compiled);
	static readonly Regex EmbeddedName = new(@"embedded-\d+\.[a-z]+", RegexOptions.Compiled);

	// Pictures' data: URIs as short names (the same picture keeps its name), and back.
	string Embedded(string text) => DataUri.Replace(text, m => {
		if (!embeddedName.TryGetValue(m.Value, out var name)) {
			string type = m.Groups["type"].Value.ToLowerInvariant();
			string extension = type switch { "jpeg" => "jpg", "svg+xml" => "svg", "x-icon" or "vnd.microsoft.icon" => "ico", _ => type };
			name = $"embedded-{embeddedData.Count + 1}.{extension}";
			embeddedName[m.Value] = name;
			embeddedData[name] = m.Value;
		}
		return name;
	});

	string WithPictures(string text) => embeddedData.Count == 0 ? text : EmbeddedName.Replace(text, m => embeddedData.TryGetValue(m.Value, out var data) ? data : m.Value);

	// ---------------------------------------------------------------- Iteration 17: Markdown and HTML from the ribbon and keys

	bool InMarkdown => shownKind == SourceKind.Markdown;

	void BuildSourceCommands() {
		void Both(RibbonCommand id, Action markdown, Action html) => sourceCommands[id] = () => { if (InMarkdown) markdown(); else html(); };
		Both(RibbonCommand.Bullets, () => MarkdownList(false), () => HtmlList(false, ""));
		Both(RibbonCommand.Numbering, () => MarkdownList(true), () => HtmlList(true, ""));
		Both(RibbonCommand.ListBullet, () => MarkdownList(false), () => HtmlList(false, ""));
		Both(RibbonCommand.ListDecimal, () => MarkdownList(true), () => HtmlList(true, ""));
		foreach (var (id, type) in new[] { (RibbonCommand.ListLowerAlpha, "a"), (RibbonCommand.ListUpperAlpha, "A"), (RibbonCommand.ListLowerRoman, "i"), (RibbonCommand.ListUpperRoman, "I") })
			Both(id, () => MarkdownList(true, type), () => HtmlList(true, $" type=\"{type}\""));
		Both(RibbonCommand.Legal, () => MarkdownList(true, "legal"), () => HtmlList(true, " class=\"rte-legal\""));
		sourceCommands[RibbonCommand.Indent] = () => ListLevel(1);
		sourceCommands[RibbonCommand.Outdent] = () => ListLevel(-1);
		sourceCommands[RibbonCommand.IndentParagraph] = () => ListLevel(1);
		sourceCommands[RibbonCommand.OutdentParagraph] = () => ListLevel(-1);
		sourceCommands[RibbonCommand.AlignLeft] = () => Align("left");
		sourceCommands[RibbonCommand.AlignCenter] = () => Align("center");
		sourceCommands[RibbonCommand.AlignRight] = () => Align("right");
		sourceCommands[RibbonCommand.Justify] = () => Align("justify");
		sourceCommands[RibbonCommand.UnderlineSingle] = () => Wrap("<u>", "</u>");
		foreach (var (id, mode) in new[] { (RibbonCommand.CaseUpper, "upper"), (RibbonCommand.CaseLower, "lower"), (RibbonCommand.CaseTitle, "title"), (RibbonCommand.CaseHeadline, "headline") })
			sourceCommands[id] = () => ChangeSourceCase(mode);
		sourceCommands[RibbonCommand.PageBreak] = () => InsertBlock("<div style=\"page-break-after: always; break-after: page\"></div>");
		sourceCommands[RibbonCommand.DateTime] = () => sourceBox.SelectedText = DateTime.Now.ToString("f");
		sourceCommands[RibbonCommand.Link] = SourceLink;
		sourceCommands[RibbonCommand.Picture] = SourcePicture;
		sourceCommands[RibbonCommand.Table] = () => {
			if (InMarkdown && SourceTableAt(sourceBox.SelectionStart) != null) { FormatSourceTable(0); SetStatus(T("status.tableLinedUp", "The table's columns are lined up. Tab and Shift+Tab move between cells and do the same.")); }
			else SourceTable();
		};
		sourceKeys[Keys.Control | Keys.B] = () => Wrap(InMarkdown ? "**" : "<strong>", InMarkdown ? "**" : "</strong>");
		sourceKeys[Keys.Control | Keys.I] = () => Wrap(InMarkdown ? "*" : "<em>", InMarkdown ? "*" : "</em>");
		sourceKeys[Keys.Control | Keys.U] = () => Wrap("<u>", "</u>");
		sourceKeys[Keys.Control | Keys.Shift | Keys.Oemplus] = () => Wrap("<sup>", "</sup>");
		sourceKeys[Keys.Control | Keys.Oemplus] = () => Wrap("<sub>", "</sub>");
		foreach (var style in Enum.GetValues<BlockStyle>())
			sourceKeys[style == BlockStyle.Preformatted ? Keys.Control | Keys.D7 : Keys.Control | (Keys.D0 + (int)style)] = () => SourceBlockStyle(style);
	}

	// Runs a formatting command in the source editor (an edit there, loaded into the document as typing is).
	void RunSource(Action action) {
		try {
			action();
		} catch (Exception ex) {
			if (ex is TargetInvocationException { InnerException: { } inner }) ex = inner;
			SetStatus("⚠ " + ex.Message);
			System.Media.SystemSounds.Beep.Play();
		}
		sourceBox.Focus();
		UpdateUi();
	}

	// Replaces part of the source text (one undo step) and selects a part of the new text.
	void ReplaceSource(int start, int length, string text, int selectFrom, int selectLength) {
		sourceBox.Select(start, length);
		sourceBox.SelectedText = text;
		sourceBox.Select(start + selectFrom, selectLength);
	}

	// The whole lines the selection touches: where they start and end (without the last line's line break), and their text.
	(int Start, int End, string[] Lines) SelectedLines() {
		string text = sourceBox.Text;
		int from = sourceBox.SelectionStart, to = from + sourceBox.SelectionLength;
		if (to > from && text[to - 1] == '\n') to--;    // a selection of whole lines ends after the last one's line break
		int start = from == 0 ? 0 : text.LastIndexOf('\n', from - 1) + 1;
		int end = text.IndexOf('\n', to);
		if (end < 0) end = text.Length;
		return (start, end, text[start..end].Split('\n'));
	}

	void ReplaceLines(int start, int end, IEnumerable<string> lines) {
		string text = string.Join("\n", lines);
		ReplaceSource(start, end - start, text, 0, text.Length);
	}

	// Puts the selection between two markers, or takes them away if it already is (Ctrl+B: **...** or <strong>...</strong>).
	void Wrap(string open, string close) {
		string text = sourceBox.Text, selected = sourceBox.SelectedText;
		int start = sourceBox.SelectionStart, length = sourceBox.SelectionLength;
		bool boldNotItalic(string s) => open == "*" && s.StartsWith("**") && !s.StartsWith("***");
		if (selected.Length >= open.Length + close.Length && selected.StartsWith(open) && selected.EndsWith(close) && !boldNotItalic(selected)) {
			ReplaceSource(start, length, selected[open.Length..^close.Length], 0, length - open.Length - close.Length);
			return;
		}
		if (start >= open.Length && start + length + close.Length <= text.Length && text.Substring(start - open.Length, open.Length) == open
			&& text.Substring(start + length, close.Length) == close && !boldNotItalic(text[Math.Max(0, start - 2)..])) {
			ReplaceSource(start - open.Length, length + open.Length + close.Length, selected, 0, length);
			return;
		}
		ReplaceSource(start, length, open + selected + close, open.Length, length);
	}

	static readonly Regex MarkdownHeading = new(@"^#{1,6}\s+"), HtmlBlock = new(@"^\s*<(h[1-6]|p)(\s[^>]*)?>(.*)</\1>\s*$", RegexOptions.IgnoreCase),
		ListMarker = new(@"^(\s*)([-*+]|\d+[.)])\s+"), ListType = new(@"\s*\[type=[^\]]*\]\s*$");

	// Ctrl+0 ... Ctrl+7 and the Style list: # headings / <h1> ..., normal text, ``` or <pre> for preformatted.
	void SourceBlockStyle(BlockStyle style) {
		var (start, end, lines) = SelectedLines();
		if (style == BlockStyle.Preformatted) {
			bool fenced = InMarkdown ? lines.Length >= 2 && lines[0].StartsWith("```") && lines[^1].StartsWith("```")
				: lines[0].TrimStart().StartsWith("<pre", StringComparison.OrdinalIgnoreCase) && lines[^1].TrimEnd().EndsWith("</pre>", StringComparison.OrdinalIgnoreCase);
			if (fenced && InMarkdown) ReplaceLines(start, end, lines[1..^1]);
			else if (fenced) ReplaceSource(start, end - start, Regex.Replace(Regex.Replace(sourceBox.Text[start..end], @"^\s*<pre[^>]*>", "", RegexOptions.IgnoreCase), @"</pre>\s*$", "", RegexOptions.IgnoreCase), 0, 0);
			else ReplaceLines(start, end, InMarkdown ? lines.Prepend("```").Append("```") : new[] { "<pre>" + string.Join("\n", lines) + "</pre>" });
			return;
		}
		int level = (int)style;    // 0: normal text
		ReplaceLines(start, end, lines.Select(line => {
			if (line.Trim() == "") return line;
			if (InMarkdown) {
				string text = MarkdownHeading.Replace(line, "");
				return level > 0 ? new string('#', level) + " " + text : text;
			}
			var block = HtmlBlock.Match(line);
			string inner = block.Success ? block.Groups[3].Value : line.Trim();
			return level > 0 ? $"<h{level}>{inner}</h{level}>" : $"<p>{inner}</p>";
		}));
	}

	// Bullets or numbers on the selected lines, or off if they all have them already; type: M-ArkDown's [type=a] ... or legal.
	void MarkdownList(bool ordered, string? type = null) {
		var (start, end, lines) = SelectedLines();
		var items = lines.Where(line => line.Trim() != "").ToList();
		bool already = items.Count > 0 && items.All(line => ListMarker.Match(line) is { Success: true } m && char.IsDigit(m.Groups[2].Value[0]) == ordered);
		int number = 1;
		bool first = true;
		ReplaceLines(start, end, lines.Select(line => {
			if (line.Trim() == "") return line;
			var marker = ListMarker.Match(line);
			string indent = marker.Success ? marker.Groups[1].Value : "", content = marker.Success ? line[marker.Length..] : line;
			if (already && type == null) return indent + ListType.Replace(content, "");
			content = ListType.Replace(content, "");
			if (first && type != null) content += $" [type={type}]";
			first = false;
			return indent + (ordered ? $"{number++}." : "-") + " " + content;
		}));
		RenumberMarkdownList();
	}

	// The selected lines as an <ul> / <ol> of <li>, or back to lines if they are one already.
	void HtmlList(bool ordered, string attributes) {
		var (start, end, lines) = SelectedLines();
		if (lines.Length >= 2 && Regex.IsMatch(lines[0].Trim(), @"^<(ul|ol)\b", RegexOptions.IgnoreCase) && Regex.IsMatch(lines[^1].Trim(), @"^</(ul|ol)>$", RegexOptions.IgnoreCase)) {
			ReplaceLines(start, end, lines[1..^1].Select(line => Regex.Replace(line.Trim(), @"^<li[^>]*>(.*)</li>$", "$1", RegexOptions.IgnoreCase)));
			return;
		}
		string tag = ordered ? "ol" : "ul";
		ReplaceLines(start, end, lines.Where(line => line.Trim() != "").Select(line => $"<li>{Regex.Replace(line.Trim(), @"^<p>(.*)</p>$", "$1")}</li>")
			.Prepend($"<{tag}{attributes}>").Append($"</{tag}>"));
	}

	// Tab / Shift+Tab for Markdown lists: list items move under the item above (to its text's column), or back out.
	void ListLevel(int delta) {
		if (!InMarkdown) throw new InvalidOperationException("In HTML, nest a list by putting an <ul> or <ol> inside an <li>.");
		var (start, end, lines) = SelectedLines();
		string before = sourceBox.Text[..start];
		var above = before.Split('\n').Reverse().Select(line => ListMarker.Match(line)).Where(m => m.Success).ToList();
		var firstMarker = ListMarker.Match(lines[0]);
		int current = firstMarker.Success ? firstMarker.Groups[1].Value.Length : lines[0].Length - lines[0].TrimStart().Length;
		int target;
		if (delta > 0) {
			var parent = above.FirstOrDefault(m => m.Groups[1].Value.Length <= current);
			target = parent != null ? Math.Max(current + 1, parent.Groups[1].Value.Length + parent.Groups[2].Value.Length + 1) : current + 4;
		} else {
			var parent = above.FirstOrDefault(m => m.Groups[1].Value.Length < current);
			target = parent != null ? parent.Groups[1].Value.Length : 0;
		}
		int shift = target - current;
		ReplaceLines(start, end, lines.Select(line => line.Trim() == "" ? line
			: shift >= 0 ? new string(' ', shift) + line : line[Math.Min(-shift, line.Length - line.TrimStart(' ').Length)..]));
		RenumberMarkdownList();
	}

	// Numbers the items of the Markdown list around the selection 1, 2, 3 ... per level (keeping the top level's start
	// number, and each item's "." or ")"), as they would be typed: after a level change, or making lines a list.
	void RenumberMarkdownList() {
		string text = sourceBox.Text;
		int selectionStart = sourceBox.SelectionStart, selectionLength = sourceBox.SelectionLength;
		var lines = text.Split('\n');
		int caretLine = text[..selectionStart].Count(c => c == '\n');
		int first = caretLine, last = caretLine;
		while (first > 0 && lines[first - 1].Trim() != "") first--;
		while (last < lines.Length - 1 && lines[last + 1].Trim() != "") last++;
		var counters = new SortedDictionary<int, int>();    // indent -> the last number used at that level
		int topIndent = Enumerable.Range(first, last - first + 1).Select(i => ListMarker.Match(lines[i])).Where(m => m.Success)
			.Select(m => m.Groups[1].Value.Length).DefaultIfEmpty(0).Min();
		bool changed = false;
		for (int i = first; i <= last; i++) {
			var marker = ListMarker.Match(lines[i]);
			if (!marker.Success) continue;
			int indent = marker.Groups[1].Value.Length;
			foreach (int deeper in counters.Keys.Where(k => k > indent).ToList()) counters.Remove(deeper);
			string value = marker.Groups[2].Value;
			if (!char.IsDigit(value[0])) { counters.Remove(indent); continue; }
			int number = counters.TryGetValue(indent, out int previous) ? previous + 1
				: indent == topIndent ? int.Parse(value[..^1], CultureInfo.InvariantCulture) : 1;    // a nested run starts at 1
			counters[indent] = number;
			string renumbered = number.ToString(CultureInfo.InvariantCulture) + value[^1];
			if (renumbered != value) {
				lines[i] = marker.Groups[1].Value + renumbered + lines[i][(marker.Groups[1].Length + value.Length)..];
				changed = true;
			}
		}
		if (!changed) return;
		int blockStart = lines.Take(first).Sum(line => line.Length + 1), blockEnd = blockStart + lines[first..(last + 1)].Sum(line => line.Length + 1) - 1;
		string block = string.Join("\n", lines[first..(last + 1)]);
		sourceBox.Select(blockStart, blockEnd - blockStart);
		sourceBox.SelectedText = block;
		sourceBox.Select(Math.Min(selectionStart, sourceBox.TextLength), 0);
	}

	// Alignment: the selected lines in <div align="..."> (with blank lines inside in Markdown, which keeps it Markdown).
	void Align(string alignment) {
		var (start, end, lines) = SelectedLines();
		var body = lines.ToList();
		if (body.Count >= 2 && Regex.IsMatch(body[0].Trim(), @"^<div\s+align=", RegexOptions.IgnoreCase) && body[^1].Trim().Equals("</div>", StringComparison.OrdinalIgnoreCase)) {
			body = body.Skip(1).SkipLast(1).ToList();
			while (body.Count > 0 && body[0].Trim() == "") body.RemoveAt(0);
			while (body.Count > 0 && body[^1].Trim() == "") body.RemoveAt(body.Count - 1);
		}
		if (alignment != "left") {
			body.Insert(0, $"<div align=\"{alignment}\">");
			body.Add("</div>");
			if (InMarkdown) { body.Insert(1, ""); body.Insert(body.Count - 1, ""); }
		}
		ReplaceLines(start, end, body);
	}

	void ChangeSourceCase(string mode) {
		string text = sourceBox.SelectedText;
		if (text == "") throw new InvalidOperationException("Select the text first.");
		string changed = mode switch {
			"upper" => text.ToUpper(CultureInfo.CurrentCulture),
			"lower" => text.ToLower(CultureInfo.CurrentCulture),
			"headline" => HeadlineCase.Apply(text),
			_ => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLower(CultureInfo.CurrentCulture))
		};
		ReplaceSource(sourceBox.SelectionStart, text.Length, changed, 0, changed.Length);
	}

	// A block (page break, table) on lines of its own, after the line with the caret.
	void InsertBlock(string block) {
		string text = sourceBox.Text;
		int at = text.IndexOf('\n', sourceBox.SelectionStart + sourceBox.SelectionLength);
		at = at < 0 ? text.Length : at;
		string insert = (at > 0 && text[..at].TrimEnd(' ') != "" ? "\n" : "") + (InMarkdown ? "\n" : "") + block + (InMarkdown ? "\n" : "");
		ReplaceSource(at, 0, insert, insert.Length, 0);
	}

	void SourceLink() {
		string selected = sourceBox.SelectedText.Trim();
		if (Prompt(T("dlg.hyperlink.title", "Insert hyperlink"), (T("dlg.hyperlink.address", "Address"), "https://"), (T("dlg.hyperlink.display", "Text to display"), selected)) is not { } values) return;
		string url = values[0], display = values[1] != "" ? values[1] : values[0];
		sourceBox.SelectedText = InMarkdown ? $"[{display}]({url})" : $"<a href=\"{WebUtility.HtmlEncode(url)}\">{WebUtility.HtmlEncode(display)}</a>";
	}

	// The picture travels with the document: it goes in as a data: URI, shown as embedded-N.png.
	void SourcePicture() {
		if (ChooseImage() is not { } path) return;
		if (Prompt(T("dlg.image.title", "Insert image"), (T("dlg.image.altText", "Alt text: what the picture shows, for screen readers (empty for a decorative picture)"), Path.GetFileNameWithoutExtension(path))) is not { } values) return;
		string type = Path.GetExtension(path).ToLowerInvariant() switch { ".jpg" or ".jpeg" => "jpeg", ".gif" => "gif", _ => "png" };
		string name = Embedded($"data:image/{type};base64,{Convert.ToBase64String(File.ReadAllBytes(path))}");
		sourceBox.SelectedText = InMarkdown ? $"![{values[0]}]({name})" : $"<img src=\"{name}\" alt=\"{WebUtility.HtmlEncode(values[0])}\">";
	}

	void SourceTable() {
		using var dlg = new InsertTableDialog(Settings.TableRows, Settings.TableColumns, Settings.TableHeader);
		if (dlg.ShowDialog(this) != DialogResult.OK) return;
		(Settings.TableRows, Settings.TableColumns, Settings.TableHeader) = (dlg.Rows, dlg.Columns, dlg.Header);
		Settings.Save();
		var columns = Enumerable.Range(1, dlg.Columns);
		string table;
		if (InMarkdown) {    // Markdown tables always have a header row
			var rows = new List<string> { "| " + string.Join(" | ", columns.Select(c => $"Column {c}")) + " |", "|" + string.Concat(columns.Select(c => " --- |")) };
			for (int r = 1; r < dlg.Rows; r++) rows.Add("|" + string.Concat(columns.Select(c => "   |")));
			table = string.Join("\n", rows);
		} else {
			var rows = new List<string> { "<table>" };
			for (int r = 0; r < dlg.Rows; r++) {
				rows.Add("<tr>" + string.Concat(columns.Select(c => r == 0 && dlg.Header ? $"<th>Column {c}</th>" : "<td></td>")) + "</tr>");
			}
			rows.Add("</table>");
			table = string.Join("\n", rows);
		}
		InsertBlock(table);
	}

	// The Markdown table the position is in (its rows and delimiter row, as lines next to each other with "|"), or null.
	(int FirstLine, List<string> Lines)? SourceTableAt(int position) {
		var lines = sourceBox.Text.Split('\n');
		int line = sourceBox.Text[..Math.Min(position, sourceBox.TextLength)].Count(c => c == '\n');
		bool IsRow(int i) => i >= 0 && i < lines.Length && lines[i].Trim() != "" && lines[i].Contains('|');
		if (!IsRow(line)) return null;
		int first = line, last = line;
		while (IsRow(first - 1)) first--;
		while (IsRow(last + 1)) last++;
		var table = lines[first..(last + 1)].ToList();
		return table.Count >= 2 && MarkdownConverter.IsTableDelimiter(table[1]) ? (first, table) : null;
	}

	// The cell a position on a table line is in: unescaped "|" before it, less the one the row starts with.
	static int CellIndex(string line, int column) {
		int pipes = 0;
		for (int i = 0; i < Math.Min(column, line.Length); i++) {
			if (line[i] == '\\') { i++; continue; }
			if (line[i] == '|') pipes++;
		}
		return Math.Max(0, line.TrimStart().StartsWith('|') ? pipes - 1 : pipes);
	}

	// Lines up the columns of the Markdown table with the caret and moves the caret move cells on (Tab: 1, Shift+Tab:
	// -1, 0: stays), as Markdown-aware editors do. Tab in the last cell adds a row.
	void FormatSourceTable(int move) {
		if (SourceTableAt(sourceBox.SelectionStart) is not var (first, lines)) return;
		string text = sourceBox.Text;
		int caretLine = text[..sourceBox.SelectionStart].Count(c => c == '\n');
		int lineStart = caretLine == 0 ? 0 : text.LastIndexOf('\n', sourceBox.SelectionStart - 1) + 1;
		int row = caretLine - first, cell = CellIndex(lines[row], sourceBox.SelectionStart - lineStart);
		var aligned = MarkdownConverter.AlignTable(lines);
		int columns = aligned[0].Count(c => c == '|') - 1;
		cell = Math.Min(cell, columns - 1) + move;
		if (cell >= columns) { cell = 0; row++; }
		if (cell < 0) { cell = columns - 1; row--; }
		if (row == 1) row += move < 0 ? -1 : 1;    // over the delimiter row
		if (row < 0) { row = 0; cell = 0; }
		if (row >= aligned.Count) {    // past the last cell: a new, empty row
			aligned.Add("|" + string.Concat(Enumerable.Repeat(" |", columns)));
			aligned = MarkdownConverter.AlignTable(aligned);
		}
		int start = text.Split('\n').Take(first).Sum(l => l.Length + 1);
		int length = lines.Sum(l => l.Length + 1) - 1;
		string table = string.Join("\n", aligned);
		// the caret at the start of the cell's text (past the padding of a right-aligned column), or after "| " if it is empty
		int offset = aligned.Take(row).Sum(l => l.Length + 1), pipe = -1;
		for (int i = 0; i <= cell; i++) pipe = aligned[row].IndexOf('|', pipe + 1);
		int next = aligned[row].IndexOf('|', pipe + 1), at = pipe + 2;
		while (at < next - 1 && aligned[row][at] == ' ') at++;
		if (at >= next - 1) at = pipe + 2;
		ReplaceSource(start, length, table, Math.Min(offset + at, table.Length), 0);
	}

	// The font control: bold, italic ... as their markers; font, size and colors as <span style="...">.
	void SourceFont(Dictionary<FontPropertiesEnum, object> changes) {
		bool Has(FontPropertiesEnum key) => changes.ContainsKey(key);
		uint Number(FontPropertiesEnum key) => Convert.ToUInt32(changes[key]);
		if (Has(FontPropertiesEnum.Bold)) sourceKeys[Keys.Control | Keys.B]();
		if (Has(FontPropertiesEnum.Italic)) sourceKeys[Keys.Control | Keys.I]();
		if (Has(FontPropertiesEnum.Underline)) Wrap("<u>", "</u>");
		if (Has(FontPropertiesEnum.Strikethrough)) Wrap(InMarkdown ? "~~" : "<s>", InMarkdown ? "~~" : "</s>");
		if (Has(FontPropertiesEnum.VerticalPositioning)) {
			var position = (FontVerticalPosition)Number(FontPropertiesEnum.VerticalPositioning);
			if (position == FontVerticalPosition.SuperScript) Wrap("<sup>", "</sup>");
			else if (position == FontVerticalPosition.SubScript) Wrap("<sub>", "</sub>");
		}
		var style = new List<string>();
		if (Has(FontPropertiesEnum.Family) && changes[FontPropertiesEnum.Family] is string family && family.Trim() != "") style.Add($"font-family: {family.Trim()}");
		if (Has(FontPropertiesEnum.Size) && changes[FontPropertiesEnum.Size] is decimal size && size > 0) style.Add($"font-size: {size.ToString(CultureInfo.InvariantCulture)}pt");
		string Hex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";
		if (Has(FontPropertiesEnum.ForegroundColor) && changes[FontPropertiesEnum.ForegroundColor] is Color fore
			&& !(Has(FontPropertiesEnum.ForegroundColorType) && (SwatchColorType)Number(FontPropertiesEnum.ForegroundColorType) != SwatchColorType.RGB))
			style.Add($"color: {Hex(fore)}");
		if (Has(FontPropertiesEnum.BackgroundColor) && changes[FontPropertiesEnum.BackgroundColor] is Color back
			&& !(Has(FontPropertiesEnum.BackgroundColorType) && (SwatchColorType)Number(FontPropertiesEnum.BackgroundColorType) != SwatchColorType.RGB))
			style.Add($"background-color: {Hex(back)}");
		if (style.Count > 0) Wrap($"<span style=\"{string.Join("; ", style)}\">", "</span>");
		if (Has(FontPropertiesEnum.DeltaSize) && !Has(FontPropertiesEnum.Size)) SetStatus(T("status.growShrinkRichText", "Grow and shrink font work in the Rich Text view; choose a size instead."));
	}

	// ---------------------------------------------------------------- Small dialogs

	void ShowText(string title, string text) {
		using var dlg = new Form { Text = title, StartPosition = FormStartPosition.CenterParent, ShowInTaskbar = false, MinimizeBox = false,
			Size = new Size(LogicalToDeviceUnits(760), LogicalToDeviceUnits(560)) };
		var box = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
			Font = new Font("Consolas", 10f), Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n"), BackColor = SystemColors.Window };
		var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(4) };
		var close = Dialogs.Cancel(this, T("dlg.close", "Close"));
		var copy = Dialogs.Button(this, T("dlg.copy", "Copy"));
		copy.Click += (s, e) => { if (text != "") Clipboard.SetText(text); };
		buttons.Controls.AddRange(new Control[] { close, copy });
		dlg.Controls.Add(box);
		dlg.Controls.Add(buttons);
		dlg.CancelButton = close;
		box.Select(0, 0);
		dlg.ShowDialog(this);
	}

	string[]? Prompt(string title, params (string Label, string Value)[] fields) {
		using var dlg = new Form();
		Dialogs.Setup(dlg, title);
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
		var boxes = new List<TextBox>();
		foreach (var (label, value) in fields) {
			table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left });
			var box = new TextBox { Text = value, Width = LogicalToDeviceUnits(380) };
			table.Controls.Add(box);
			boxes.Add(box);
		}
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		dlg.Controls.Add(table);
		dlg.AcceptButton = ok;
		dlg.CancelButton = cancel;
		return dlg.ShowDialog(this) == DialogResult.OK ? boxes.Select(b => b.Text).ToArray() : null;
	}
}

// The automation surface. Every public member can be called from AHK by name (Rte.<Name> in RteClient.ahk), along with
// the RichEditBox members listed in Automation.EditorMembers. A member here wins over a RichEditBox member of the same name.
// These methods never show dialogs, so a calling script is never left waiting for the user.
public class EditorApi
{
	readonly WriterForm form;
	RichEditBox ed => building ?? form.Editor;
	RichEditBox? building;    // (iteration 68: the hidden editor a document is being built in: SetHtml)

	internal EditorApi(WriterForm form) {
		this.form = form;
		form.Editor.TextChanged += (s, e) => altCache = null;
	}

	public string Version => $"{WriterForm.AppName} 0.11, RichEditorNET " +
		(typeof(RichEditBox).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0]
			?? typeof(RichEditBox).Assembly.GetName().Version?.ToString());

	// ---- Documents

	/// Starts a new, empty document (unsaved changes are discarded).
	public void New() {
		form.Styles = form.DefaultStyles();
		form.Levels = form.DefaultLevels();
		form.UseNormalFont();
		FreshEditor();
		ed.Html = "";
		(ed as WriterEditBox)?.AdvancedLayout(false);
		StyleEngine.Apply(ed.TextDocument, form.Styles, form.Styles.Normal, 0, 0);
		StyleEngine.SetTyping(ed.TextDocument.Selection2.Font, form.Styles.Normal);
		form.FillStyleBox(force: true);
		form.StyleSetMayHaveChanged();
		form.Bands = form.DefaultBands();
		form.SetLineNumbering(new LineNumbering(), changesDocument: false);
		form.SetDocumentPageSetup(null, null);    // (iteration 63: Page Setup's)
		ed.Modified = false;
		FileConverter = "";
		form.SetFileName(null);
		form.DocumentLoaded(WriterForm.SourceKind.None, null);
		form.Events.Raise(EditorEvent.Opened);
		if (!form.StartingUp) form.RunHook(HookEvent.NewDocument);    // (iteration 56; the first document at the start: Started)
	}

	/// Seconds between the recovery copies written while the document has unsaved changes (0: none; at least 10).
	/// Remembered.
	public int AutosaveSeconds {
		get => form.Settings.AutosaveSeconds;
		set {
			if (value < 0) throw new ArgumentException("AutosaveSeconds is 0 (off) or more.");
			form.Settings.AutosaveSeconds = value;
			form.Settings.Save();
			form.StartAutosave();
		}
	}

	/// The folder of the recovery copies (beside the settings).
	public string RecoveryFolder => form.RecoveryFolder;

	/// Whether AhkWord runs in portable mode (a file named "portable" or "portable.txt" beside it): its settings,
	/// recovery copies and the converters' state are in its data folder.
	public bool Portable => AppFolders.Portable;

	// The recovery copy's content: the document as RTF, with its header, footer and line numbering.
	// (iteration 47: the list indents fitted before the styles go in: fitting them goes through RichEdit, which writes
	// only its own styles, Normal and the headings, and so lost the others' paragraphs, a Numbered Paragraph's)
	internal string RecoveryRtf() {
		string rtf = OfficeConverters.ListIndentsRtf(form.RtfLeavingGaps());
		return form.LeavingGaps(() => WithBands(StyledRtf(rtf), form.FileName ?? ""));
	}

	// A recovery copy back in the editor: as unsaved changes to the document it was (fileName; null: a new one).
	internal void Recover(string rtf, string? fileName) {
		var bands = HeaderFooter.FromRtf(rtf);
		var numbering = LineNumberFormats.FromRtf(rtf);
		form.Styles = form.DefaultStyles();
		form.Levels = form.DefaultLevels();
		var (margins, landscape, paper) = PageSetupOf(rtf);
		form.SetDocumentPageSetup(margins, landscape, Sections.FirstColumnsFromRtf(rtf), paper);
		rtf = Sections.ImportRtf(rtf);    // (iteration 63)
		Rtf = rtf;
		form.StylesAfterLoad(StylesFromRtf(rtf), takeLook: false);
		OfficeConverters.NewNumbersInto(ed, loadedNewNumbers);    // (iteration 87: the lists started again)
		loadedNewNumbers = null;
		form.Bands = bands ?? form.DefaultBands();
		form.SetLineNumbering(numbering ?? new LineNumbering(), changesDocument: false);
		MarksFollowText();
		FileConverter = "";
		LastConverter = "";
		form.SetFileName(fileName is { Length: > 0 } ? fileName : null);
		form.DocumentLoaded(WriterForm.SourceKind.None, null);
		ed.Modified = true;
		form.Events.Raise(EditorEvent.Opened);
	}

	/// The templates' names, one a line: the documents (File > New from Template), or with parts = true the parts (pieces
	/// for Insert from Template, in the Parts folders).
	public string Templates(bool parts = false) => string.Join("\n", form.TemplateList().Where(t => t.Part == parts).Select(t => t.Name));

	/// Inserts a template's content at the caret (a part, by name, such as "Signature Block", or a document, or a file's
	/// path), its fields filled in: fields as for NewFromTemplate. Returns how many fields were filled in.
	public int InsertTemplate(string template, string fields = "", bool ask = false) =>
		form.InsertTemplate(TemplateNamed(template, parts: true), FieldLines(fields), ask);

	// ---- Snippets (iteration 55)

	/// The snippets, one a line: the abbreviation, a tab, the part it inserts, a tab, how it fires ("ending": on Space,
	/// Enter, Tab or punctuation; "ending kept": the character stays after it; "instant": at once), a tab, "case" if the
	/// abbreviation's case must match (else "any").
	public string Snippets => string.Join("\n", form.SnippetList().OrderBy(s => s.Abbreviation, StringComparer.OrdinalIgnoreCase)
		.Select(s => s.Abbreviation + "\t" + s.Part + "\t" + (s.Instant ? "instant" : s.KeepEnding ? "ending kept" : "ending") + "\t" + (s.CaseSensitive ? "case" : "any")));

	/// Gives a part (by name: Rte.Templates(true) lists them; an RTF template) an abbreviation: typed in the document,
	/// then Space, Enter, Tab or punctuation (instant = 1: at once), it becomes the part. keepEnding = 1: the ending
	/// character stays after it; caseSensitive = 1: the abbreviation fires only typed with its case. An abbreviation it
	/// had already is changed.
	public void AddSnippet(string abbreviation, string part, bool instant = false, bool keepEnding = false, bool caseSensitive = false) {
		abbreviation = abbreviation.Trim();
		var list = form.SnippetList();
		string problem = global::Snippets.Problem(abbreviation, list, abbreviation);
		if (problem != "") throw new ArgumentException(problem);
		var snippet = new Snippet(abbreviation, part.Trim(), instant, keepEnding && !instant, caseSensitive);
		if (form.SnippetPart(snippet) == null) throw new ArgumentException($"There is no part \"{part}\" (an RTF template: Rte.Templates(true) lists them).");
		list.RemoveAll(s => string.Equals(s.Abbreviation, abbreviation, StringComparison.OrdinalIgnoreCase));
		list.Add(snippet);
		global::Snippets.Save(form.SnippetsFile, list);
	}

	/// Takes a snippet's abbreviation away (its part stays). Returns whether there was one.
	public bool RemoveSnippet(string abbreviation) {
		var list = form.SnippetList();
		bool removed = list.RemoveAll(s => string.Equals(s.Abbreviation, abbreviation.Trim(), StringComparison.OrdinalIgnoreCase)) > 0;
		if (removed) global::Snippets.Save(form.SnippetsFile, list);
		return removed;
	}

	/// Inserts a snippet's part at the caret (as Insert from Template does): fields as for InsertTemplate. Returns how many
	/// fields were filled in.
	public int InsertSnippet(string abbreviation, string fields = "", bool ask = false) {
		var snippet = form.SnippetList().FirstOrDefault(s => string.Equals(s.Abbreviation, abbreviation.Trim(), StringComparison.OrdinalIgnoreCase))
			?? throw new ArgumentException($"There is no snippet \"{abbreviation}\" (Rte.Snippets lists them).");
		var part = form.SnippetPart(snippet) ?? throw new FileNotFoundException($"The snippet's part \"{snippet.Part}\" isn't in the templates' Parts folders.");
		return form.InsertTemplate(part, FieldLines(fields), ask, inline: true);
	}

	/// Saves the selection, formatting and all, as a new part (in your own templates' Parts folder, as part.rtf), and
	/// gives it an abbreviation (as AddSnippet).
	public void SaveSelectionAsSnippet(string part, string abbreviation, bool instant = false, bool keepEnding = false, bool caseSensitive = false) {
		part = part.Trim();
		if (part == "" || part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) throw new ArgumentException("A part's name can't be empty, nor have \\ / : * ? \" < > | in it.");
		var selection = ed.TextDocument.Selection2;
		if (selection.End <= selection.Start) throw new InvalidOperationException("Select the text to save as a snippet first.");
		string problem = global::Snippets.Problem(abbreviation.Trim(), form.SnippetList(), abbreviation.Trim());
		if (problem != "") throw new ArgumentException(problem);
		string folder = Path.Combine(form.OwnTemplatesFolder, "Parts");
		string file = Path.Combine(folder, part + ".rtf");
		if (File.Exists(file)) throw new IOException($"There is a part \"{part}\" already ({file}): choose another name, or give it the abbreviation (AddSnippet).");
		string rtf = form.WithoutGaps(() => ((RichTextBox)ed).SelectedRtf ?? "");
		Directory.CreateDirectory(folder);
		File.WriteAllText(file, rtf, Encoding.Latin1);
		AddSnippet(abbreviation, part, instant, keepEnding, caseSensitive);
	}

	/// Writes the snippets as an AutoHotkey v2 script of plain-text hotstrings, to use them in any program (their fields
	/// as their defaults or [their names]). Returns its full path.
	public string ExportSnippets(string path) {
		string full = Path.GetFullPath(path);
		File.WriteAllText(full, global::Snippets.Script(form.SnippetList(), s => form.SnippetPart(s) is { } part ? WriterForm.SnippetPlainText(part) : null), new UTF8Encoding(true));
		return full;
	}

	/// For AhkWord.ahk: the snippets' hotstrings, one a line (their AutoHotkey options, a tab, the abbreviation), and the
	/// file they are kept in (read again when it changes).
	public string SnippetHotstrings => global::Snippets.Definitions(form.SnippetList());
	public string SnippetsFile => form.SnippetsFile;

	/// For AhkWord.ahk: a snippet's hotstring fired, with the ending character typed after it ("" for one that fires at
	/// once). Returns at once; the snippet goes in once the editor has what was typed.
	public void FireSnippet(string abbreviation, string ending = "") => form.FireSnippet(abbreviation, ending);

	/// Fills in the {{fields}} left in the document: fields as for NewFromTemplate. Returns how many were filled in.
	public int FillFields(string fields = "", bool ask = false) => form.FillFields(FieldLines(fields), ask);

	// A template by name (the parts or the documents first), or a file's path.
	TemplateInfo TemplateNamed(string template, bool parts) =>
		form.TemplateList().OrderBy(t => t.Part != parts).FirstOrDefault(t => string.Equals(t.Name, template, StringComparison.OrdinalIgnoreCase))
		?? (File.Exists(template) ? new TemplateInfo(Path.GetFileNameWithoutExtension(template), Path.GetFullPath(template), false)
			: throw new FileNotFoundException($"There is no template \"{template}\" (Rte.Templates() and Rte.Templates(true) list them)."));

	// "Name=Value" lines (\n in a value: a new line)
	static Dictionary<string, string> FieldLines(string fields) {
		var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		foreach (var line in fields.Split('\n', StringSplitOptions.RemoveEmptyEntries)) {
			int equals = line.IndexOf('=');
			if (equals > 0) values[line[..equals].Trim()] = line[(equals + 1)..].Trim().Replace("\\n", "\n");
		}
		return values;
	}

	/// Your own templates' folder (beside the settings); the shipped ones are in the templates folder beside AhkWord.
	public string TemplatesFolder => form.OwnTemplatesFolder;

	/// Makes an untitled document from a template (by name, or a file's path); unsaved changes are discarded. fields:
	/// "Name=Value" lines for its {{Name}} fields; the others get the values typed last time (or their defaults). With
	/// ask = true, the fields not given are asked for in the dialog instead.
	public void NewFromTemplate(string template, string fields = "", bool ask = false) {
		var info = TemplateNamed(template, parts: false);
		ed.Modified = false;
		form.NewFromTemplate(info, FieldLines(fields), ask);
	}

	/// The start page: showing (true) or the document (false).
	public bool StartPage {
		get => form.StartShown;
		set => form.ShowStart(value);
	}

	// A template as a new document: opened, then untitled (and not a recent document).
	internal void OpenTemplate(string path) {
		Open(path);
		FileConverter = "";
		form.SetFileName(null);
		form.ForgetRecent(Path.GetFullPath(path));
		ed.Modified = false;
	}

	/// The document's styles' names, one a line (the Style box's list).
	public string Styles() => string.Join("\n", form.Styles.Styles.Select(s => s.Name));

	/// The style of the paragraphs the selection is in ("" if they differ); setting it gives them the style.
	public string Style {
		get => form.CaretStyleName();
		set => SetStyle(value);
	}

	/// Gives the paragraphs the selection is in a style (by name: Rte.Styles() lists them), in one undo step: their
	/// formatting becomes the style's, except formatting given to part of a paragraph (a bold word).
	public void SetStyle(string name) {
		using var _ = form.FirstParagraphAsIs(changes: true);
		var style = form.Styles.Find(name)
			?? (ParagraphStyle.BuiltIn(name) || StyleSets.Get("Classic").Find(name) != null ? form.Styles.Ensure(name) : throw new ArgumentException($"There is no style \"{name}\" (Rte.Styles() lists them)."));
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		doc.BeginEditCollection();
		try {
			StyleEngine.Apply(doc, form.Styles, style, selection.Start, selection.End);
			// (an empty paragraph: what is typed next gets the style)
			var paragraph = doc.Range2(selection.Start, selection.Start);
			paragraph.Expand(tomConstants.tomParagraph);
			if (selection.Start == selection.End && (paragraph.Text ?? "").Trim('\r', '\n') == "") StyleEngine.SetTyping(selection.Font, style);
		} finally {
			doc.EndEditCollection();
		}
		form.StylesChanged();
	}

	// A style's settings changed: "Name=Value" lines (Font, Size, Bold, Italic, Underline, Color, Align, LeftIndent,
	// RightIndent, FirstLine, Before, After, LineSpacing, KeepWithNext, KeepTogether, WidowControl, PageBreakBefore,
	// Contextual, Level, Next, Numbering ("1.", "(a)", "1.1" ...; nothing: none), NumberLevel (1-9), SubLevels ("a., i., (1)";
	// nothing: the usual ones after Numbering), SubLevelIndent (points each sub-level further in, 36 at first), SubLevelSpace
	// (None, AfterParagraph or Each: the space after a paragraph with sub-levels and between them)).
	static ParagraphStyle WithSettings(ParagraphStyle style, string settings) {
		var result = style.Copy();
		foreach (var line in settings.Split('\n', StringSplitOptions.RemoveEmptyEntries)) {
			int equals = line.IndexOf('=');
			if (equals <= 0) continue;
			string key = line[..equals].Trim(), value = line[(equals + 1)..].Trim();
			var property = typeof(ParagraphStyle).GetProperty(key, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)
				?? throw new ArgumentException($"A style has no \"{key}\" (Font, Size, Bold, Italic, Underline, Color, Align, LeftIndent, RightIndent, FirstLine, Before, After, LineSpacing, KeepWithNext, KeepTogether, WidowControl, PageBreakBefore, Contextual, Level, Next, Numbering, NumberLevel, SubLevels, SubLevelIndent, SubLevelSpace).");
			if (property.Name == "Name" || !property.CanWrite) continue;
			object converted = property.PropertyType == typeof(bool) ? value is "1" or "true" or "True" or "yes"
				: property.PropertyType == typeof(float) ? float.Parse(value, CultureInfo.InvariantCulture)
				: property.PropertyType == typeof(int) ? int.Parse(value, CultureInfo.InvariantCulture) : value;
			property.SetValue(result, converted);
		}
		if (ParagraphStyle.BuiltIn(style.Name)) result.Level = style.Level;
		global::LineSpacing.Parse(result.LineSpacing);    // (throws for one that isn't)
		if (result.Numbering != "" && !NumberFormats.Valid(result.Numbering))
			throw new ArgumentException($"\"{result.Numbering}\" is not a numbering: {string.Join(", ", NumberFormats.All)}, or nothing for none.");
		if (result.NumberLevel is < 1 or > 9) throw new ArgumentException("NumberLevel is 1 (the top level) to 9.");
		if (result.SubLevelIndent is < 0 or > 144) throw new ArgumentException("SubLevelIndent is 0 to 144 points.");
		if (ParagraphStyle.SubLevelSpaces.FirstOrDefault(s => s.Equals(result.SubLevelSpace, StringComparison.OrdinalIgnoreCase)) is not { } space)
			throw new ArgumentException("SubLevelSpace is None (no space after a paragraph with sub-levels, nor between them), AfterParagraph or Each.");
		result.SubLevelSpace = space;
		foreach (var part in result.SubLevels.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
			if (!NumberFormats.Valid(part) || part == NumberFormats.Legal) throw new ArgumentException($"\"{part}\" is not a sub-level's format (SubLevels: \"a., i., (1)\", or nothing for the usual ones).");
		result.SubLevels = string.Join(",", result.SubLevels.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
		return result;
	}

	/// Changes a style: settings as "Name=Value" lines ("Font=Georgia\nSize=12\nBefore=6"); its paragraphs change too
	/// (one undo step). Returns how many paragraphs have it.
	public int ModifyStyle(string name, string settings) {
		var style = form.Styles.Find(name) ?? throw new ArgumentException($"There is no style \"{name}\" (Rte.Styles() lists them).");
		return form.ChangeStyle(style, WithSettings(style, settings));
	}

	/// A new style, from the paragraph at the caret (with settings, as ModifyStyle takes them, on top), given to the
	/// paragraphs the selection is in.
	public void NewStyle(string name, string settings = "") {
		form.BeginInvoke(() => form.RunHook(HookEvent.StyleChanged, name));    // (iteration 56)
		using var _ = form.FirstParagraphAsIs(changes: true);
		name = name.Trim();
		if (name == "" || name.Contains('@') || name.Contains(';') || name.Contains('{') || name.Contains('}')) throw new ArgumentException("A style's name can't be empty, nor have @ ; { } in it.");
		if (form.Styles.Find(name) != null) throw new ArgumentException($"There is a style \"{name}\" already.");
		var doc = ed.TextDocument;
		var style = WithSettings(StyleEngine.Capture(doc, doc.Selection2.Start, name, 0, ""), settings);
		style.Name = name;
		form.Styles.Styles.Add(style);
		SetStyle(name);
	}

	/// The caret paragraph's style made to look like the paragraph; its other paragraphs change too. Returns its name.
	public string UpdateStyleToMatch() {
		using var _ = form.FirstParagraphAsIs(changes: true, all: true);
		var doc = ed.TextDocument;
		var style = StyleEngine.StyleAt(doc, form.Styles, doc.Selection2.Start);
		form.ChangeStyle(style, StyleEngine.Capture(doc, doc.Selection2.Start, style.Name, style.Level, style.Next));
		return style.Name;
	}

	/// Clear Formatting (Ctrl+Space): the selected text back to its paragraphs' style (all formatting of its own goes); with
	/// no selection, the caret's whole paragraph, its paragraph formatting (spacing, indents ...) too. One undo step.
	public void ClearFormatting() {
		using var _ = form.FirstParagraphAsIs(changes: true);
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		doc.BeginEditCollection();
		try { StyleEngine.ResetFormatting(doc, form.Styles, selection.Start, selection.End, paragraphs: selection.Start == selection.End); }
		finally { doc.EndEditCollection(); }
		form.StylesChanged();
	}

	// ---- Format Painter (iteration 75; as Word's, iteration 77)

	ITextFont? copiedFont;
	ParagraphStyle? copiedStyle;
	ParaLook? copiedPara;    // (null: characters only, a selection without a paragraph mark)

	// A paragraph's formatting, as the Format Painter copies it: alignment, indents, spacing, keep options, tabs and its
	// numbering (list, level, series); its style is copiedStyle.
	sealed record ParaLook(int Alignment, float First, float Left, float Right, float Before, float After, int Rule, float Line,
		int KeepWithNext, int KeepTogether, int WidowControl, int PageBreakBefore, int ListType, int ListLevel, int ListStart,
		float ListTab, (float Position, int Alignment, int Leader)[] Tabs) {
		public bool Legal => LegalLists.IsLegal(ListType);

		public static ParaLook Of(ITextPara p) {
			var tabs = new List<(float, int, int)>();
			for (int i = 0; i < p.TabCount; i++) { p.GetTab(i, out float position, out int alignment, out int leader); tabs.Add((position, alignment, leader)); }
			return new(p.Alignment, p.FirstLineIndent, p.LeftIndent, p.RightIndent, p.SpaceBefore, p.SpaceAfter, p.LineSpacingRule,
				p.LineSpacing, p.KeepWithNext, p.KeepTogether, p.WidowControl, p.PageBreakBefore, p.ListType, p.ListLevelIndex,
				p.ListStart, p.ListTab, tabs.ToArray());
		}

		public void PutOn(ITextPara p) {
			// (the list first: RichEdit gives a paragraph that becomes a list item indents of its own)
			if (p.ListType != ListType) p.ListType = ListType;
			if ((ListType & 0xFFFF) != tomConstants.tomListNone) {
				if (p.ListLevelIndex != ListLevel) p.ListLevelIndex = ListLevel;
				if (p.ListStart != ListStart) p.ListStart = ListStart;
				if (Math.Abs(p.ListTab - ListTab) > 0.01f) p.ListTab = ListTab;
			}
			p.Alignment = Alignment;
			p.SetIndents(First, Left, Right);
			p.SpaceBefore = Before;
			p.SpaceAfter = After;
			p.SetLineSpacing(Rule, global::LineSpacing.Held(Rule, Line));    // (iteration 87: 1.08 held as 1.05, saved as 1.08)
			p.KeepWithNext = KeepWithNext;
			p.KeepTogether = KeepTogether;
			p.WidowControl = WidowControl;
			p.PageBreakBefore = PageBreakBefore;
			p.ClearAllTabs();
			foreach (var (position, alignment, leader) in Tabs) p.AddTab(position, alignment, leader);
		}
	}

	/// Copy Formatting (Home > Format Painter), as Word's: with no selection, the caret's paragraph formatting (its style,
	/// numbering, indents, spacing, alignment, tabs) and the character formatting at the caret (of the character before
	/// it; at a paragraph's start, of the first one); with a selection, the formatting most of the selected text has (a
	/// paragraph with a bold lead-in: its plain text's), and its paragraph formatting too only if the selection takes in a
	/// paragraph mark (or runs to the document's end). Kept for PasteFormatting.
	public void CopyFormatting() {
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		bool before = end == start && start > 0 && StyleEngine.Paragraphs(doc, start, start)[0].Start < start;
		copiedFont = (end > start ? MostFont(doc, start, end) : null)
			?? (before ? doc.Range2(start - 1, start) : doc.Range2(start, start + 1)).Font.Duplicate;
		bool paragraph = end == start || (doc.Range2(start, end).Text ?? "").Contains('\r') || end >= doc.Range2(0, 0).StoryLength - 1;
		copiedStyle = StyleEngine.StyleAt(doc, form.Styles, start);
		copiedPara = paragraph ? ParaLook.Of(doc.Range2(start, start).Para) : null;
	}

	// The formatting that covers most of the text from start to end (a paragraph with a bold lead-in: its plain text's),
	// hidden text and paragraph marks not counted; null when there is no such text.
	static ITextFont? MostFont(ITextDocument2 doc, int start, int end) {
		var counts = new List<(string Key, ITextFont Font, int Length)>();
		for (int at = start; at < end; ) {
			var run = doc.Range2(at, at);
			run.MoveEnd(13, 1);    // (tomCharFormat: to where the formatting changes)
			int to = Math.Min(run.End, end);
			if (to <= at) break;
			var range = doc.Range2(at, to);
			at = to;
			var font = range.Font;
			if (font.Hidden == tomConstants.tomTrue) continue;
			int length = (range.Text ?? "").Count(c => c != '\r' && c != '\n');
			if (length == 0) continue;
			string key = string.Join("|", font.Name, font.Size, font.Bold, font.Italic, font.Underline, font.StrikeThrough,
				font.ForeColor, font.BackColor, font.Superscript, font.Subscript, font.SmallCaps, font.AllCaps, font.Spacing);
			int i = counts.FindIndex(c => c.Key == key);
			if (i < 0) counts.Add((key, font.Duplicate, length));
			else counts[i] = (key, counts[i].Font, counts[i].Length + length);
		}
		return counts.Count == 0 ? null : counts.Aggregate((a, b) => b.Length > a.Length ? b : a).Font;
	}

	// Paragraphs by number (0: the first), which stays the same while numbers written in front of items change.
	static int ParagraphIndex(ITextDocument2 doc, int position) {
		var range = doc.Range2(position, position);
		range.Expand(tomConstants.tomParagraph);
		if (range.Start == 0) return 0;
		var at = doc.Range2(range.Start, range.Start);
		return -at.Move(tomConstants.tomParagraph, -int.MaxValue);
	}

	static (int Start, int End) ParagraphNumbered(ITextDocument2 doc, int index) {
		var range = doc.Range2(0, 0);
		if (index > 0) range.Move(tomConstants.tomParagraph, index);
		range.Expand(tomConstants.tomParagraph);
		return (range.Start, range.End);
	}

	/// Paste Formatting, as Word's: a click (no selection) gives the caret's paragraph the copied paragraph formatting
	/// (style, numbering, indents ...; its text keeps its bold, italic ...), and the word clicked in the copied character
	/// formatting; a selection takes the character formatting, and its paragraphs the paragraph formatting. Paragraph
	/// formatting goes on only if CopyFormatting took it (no selection, or one with a paragraph mark). One undo step.
	public void PasteFormatting() {
		if (copiedFont == null || copiedStyle == null) throw new InvalidOperationException("No formatting copied yet: CopyFormatting first.");
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		var paragraphs = StyleEngine.Paragraphs(doc, start, end);
		if (end == start) {
			// (the word the caret is in, without the space after it; none there: what is typed next)
			var word = doc.Range2(start, start);
			word.Expand(tomConstants.tomWord);
			string text = word.Text ?? "";
			int length = text.TrimEnd().Length;
			if (length > 0 && !char.IsPunctuation(text[0]) && !char.IsWhiteSpace(text[0]) && word.Start <= start && word.Start + length > start)
				(start, end) = (word.Start, word.Start + length);
		}
		// (the characters to paint, as paragraph numbers and distances from their paragraphs' ends: numbers written in
		// front of items change the positions, not those)
		int first = ParagraphIndex(doc, paragraphs[0].Start), count = paragraphs.Count;
		int last = Math.Max(start, end - 1);
		var startAt = (Index: ParagraphIndex(doc, start), FromEnd: StyleEngine.Paragraphs(doc, start, start)[0].End - start);
		var endAt = (Index: ParagraphIndex(doc, last), FromEnd: StyleEngine.Paragraphs(doc, last, last)[0].End - end);
		bool characters = end > start;
		using var _ = form.FirstParagraphAsIs(changes: true);
		doc.BeginEditCollection();
		try {
			if (copiedPara is { } look) {
				var style = form.Styles.Find(copiedStyle.Name) ?? form.Styles.Normal;
				bool legal = look.Legal;
				for (int i = first; i < first + count; i++) {
					int pStart = ParagraphNumbered(doc, i).Start;
					if (StyleEngine.StyleAt(doc, form.Styles, pStart).Name != style.Name) {
						StyleEngine.Apply(doc, form.Styles, style, pStart, pStart);
						pStart = ParagraphNumbered(doc, i).Start;
					}
					if (LegalLists.IsLegal(doc.Range2(pStart, pStart).Para.ListType)) {
						legal = true;
						// (an item that won't be one any more: its written number goes)
						if (!look.Legal) {
							int pEnd = ParagraphNumbered(doc, i).End;
							int prefix = LegalLists.Prefix.Match(doc.Range2(pStart, pEnd).Text ?? "").Length;
							if (prefix > 0) doc.Range2(pStart, pStart + prefix).Text = "";
						}
					}
					look.PutOn(doc.Range2(pStart, pStart).Para);
				}
				if (legal) LegalLists.Renumber(ed);    // (an item's number written in front of it, or the others' changed)
			}
			if (characters) {
				int from = ParagraphNumbered(doc, startAt.Index).End - startAt.FromEnd, to = ParagraphNumbered(doc, endAt.Index).End - endAt.FromEnd;
				// (hidden text, a footnote's note or a link's field, stays hidden)
				var hidden = new List<(int, int)>();
				for (int at = from; at < to; ) {
					var run = doc.Range2(at, at);
					run.MoveEnd(13, 1);    // (tomCharFormat: to where the formatting changes)
					int stop = Math.Min(run.End, to);
					if (stop <= at) break;
					if (doc.Range2(at, stop).Font.Hidden == tomConstants.tomTrue) hidden.Add((at, stop));
					at = stop;
				}
				if (to > from) doc.Range2(from, to).Font = copiedFont;
				foreach (var (a, b) in hidden) doc.Range2(a, b).Font.Hidden = tomConstants.tomTrue;
			} else if (copiedPara == null) selection.Font = copiedFont;
		} finally {
			doc.EndEditCollection();
		}
		if (copiedPara != null) form.StylesChanged();
	}

	/// Deletes one of your own styles; its paragraphs become Normal (one undo step).
	public void DeleteStyle(string name) {
		form.BeginInvoke(() => form.RunHook(HookEvent.StyleChanged, name));    // (iteration 56)
		var style = form.Styles.Find(name) ?? throw new ArgumentException($"There is no style \"{name}\".");
		if (StyleSets.Get("Classic").Find(style.Name) != null) throw new ArgumentException($"{style.Name} is one of the usual styles, which stay.");
		var doc = ed.TextDocument;
		int id = StyleEngine.IdOf(style.Name);
		doc.BeginEditCollection();
		try {
			// (from the last paragraph: a numbered style's numbers go with it, which moves the text after them)
			foreach (var (start, _) in Enumerable.Reverse(StyleEngine.Paragraphs(doc, 0, doc.Range2(0, 0).StoryLength)))
				if (doc.Range2(start, start).Para.Style == id) StyleEngine.Apply(doc, form.Styles, form.Styles.Normal, start, start);
			foreach (var other in form.Styles.Styles) if (string.Equals(other.Next, style.Name, StringComparison.OrdinalIgnoreCase)) other.Next = "";
			form.Styles.Styles.Remove(style);
		} finally {
			doc.EndEditCollection();
		}
		form.StylesChanged();
	}

	/// The document's style set ("Classic", "Plain", "Modern"; "" for a document's own styles, from a file made
	/// elsewhere: ApplyMyStyles gives it yours); setting it gives the document that set's styles (your
	/// own styles stay). New documents keep theirs: SaveStylesAsDefault.
	public string StyleSet {
		get => form.Styles.Set;
		set {
			if (!StyleSets.Names.Contains(value.Trim(), StringComparer.OrdinalIgnoreCase)) throw new ArgumentException("The style sets are Classic, Plain and Modern.");
			form.UseStyleSet(value.Trim());
		}
	}

	/// Iteration 40: the paragraph styles a document has (AhkWord's own RTF and HTML, a Word .docx, other RTF), one name
	/// a line, without opening it.
	public string StylesIn(string path) =>
		string.Join("\n", (StyleImport.From(Path.GetFullPath(path)) ?? throw new InvalidOperationException($"No styles found in {Path.GetFileName(path)}.")).Styles.Select(s => s.Name));

	/// Copy Styles from Document: the styles named (by commas or lines; empty: all of them) come from another document
	/// into this one. A style this document has takes the copied definition and its paragraphs the new look (formatting
	/// of their own stays); others are added. With forNewDocuments they go into your default styles too (the ones new
	/// documents get), the others of those staying as they are. Returns how many styles were copied.
	public int CopyStyles(string path, string names = "", bool forNewDocuments = false) {
		form.BeginInvoke(() => form.RunHook(HookEvent.StyleChanged, names == "" ? "(styles from " + Path.GetFileName(path) + ")" : names));    // (iteration 56)
		path = Path.GetFullPath(path);
		var source = StyleImport.From(path) ?? throw new InvalidOperationException($"No styles found in {Path.GetFileName(path)}.");
		var wanted = names.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
			.Select(n => source.Find(n) ?? throw new ArgumentException($"{Path.GetFileName(path)} has no style \"{n}\" (Rte.StylesIn lists them).")).ToList();
		if (wanted.Count == 0) wanted = source.Styles.ToList();
		// (a style after Enter that neither has: the same style again)
		ParagraphStyle Fitted(ParagraphStyle style, StyleSheet into) {
			var copy = style.Copy();
			if (copy.Next != "" && into.Find(copy.Next) == null && !wanted.Any(w => string.Equals(w.Name, copy.Next, StringComparison.OrdinalIgnoreCase))) copy.Next = "";
			return copy;
		}
		foreach (var style in wanted) {
			var copy = Fitted(style, form.Styles);
			if (form.Styles.Find(copy.Name) is { } mine) { copy.Name = mine.Name; if (ParagraphStyle.BuiltIn(mine.Name)) copy.Level = mine.Level; form.ChangeStyle(mine, copy); }
			else form.Styles.Styles.Add(copy);
		}
		form.StylesChanged();
		form.FillStyleBox(force: true);
		if (forNewDocuments) {
			var defaults = form.DefaultStyles();
			foreach (var style in wanted) {
				var copy = Fitted(style, defaults);
				int at = defaults.Styles.FindIndex(s => string.Equals(s.Name, copy.Name, StringComparison.OrdinalIgnoreCase));
				if (at >= 0) { copy.Name = defaults.Styles[at].Name; if (ParagraphStyle.BuiltIn(copy.Name)) copy.Level = defaults.Styles[at].Level; defaults.Styles[at] = copy; }
				else defaults.Styles.Add(copy);
			}
			form.Settings.MyStyles = defaults;
			form.Settings.Save();
		}
		return wanted.Count;
	}

	/// New documents get this document's styles (your changes included).
	public void SaveStylesAsDefault() {
		form.Settings.MyStyles = form.Styles.Copy();
		if (form.Styles.Set != "") form.Settings.StyleSetName = form.Styles.Set;
		form.Settings.Save();
	}

	/// Apply My Styles: the document gets your styles, the ones new documents get (saved with SaveStylesAsDefault, else
	/// your style set's): each of its styles that you have takes your definition, and its paragraphs your look
	/// (formatting of their own stays); styles of the document's own that you haven't got stay. One undo step. Returns
	/// how many paragraphs were restyled.
	public int ApplyMyStyles() => form.UseStyles(form.DefaultStyles());

	/// Opens another AhkWord window (File > New Window), with a new document or the file given. The new window is a
	/// separate editor: scripts reach it once it is the AhkWord used last (or with its RTE_WRITER_HWND).
	public void NewWindow(string path = "") => form.StartWindow(path == "" ? null : path);

	/// Closes the document and starts an empty one, as File > Close does (unsaved changes are discarded, as by New; the
	/// window stays open).
	public void Close() => New();

	// (iteration 63) A document's own page setup, as its RTF (or .docx) gives it (null: none, Page Setup's).
	// (iteration 82: the paper too)
	static (int[]?, bool?, int[]?) PageSetupOf(string rtf) => Sections.PageSetupFromRtf(rtf) is { } setup ? (setup.Margins, setup.Landscape, Sections.PaperFromRtf(rtf)) : (null, null, Sections.PaperFromRtf(rtf));
	static (int[]?, bool?, int[]?) PageSetupOfDocx(string path) => Sections.PageSetupFromDocx(path) is { } setup ? (setup.Margins, setup.Landscape, Sections.PaperFromDocx(path)) : (null, null, Sections.PaperFromDocx(path));

	/// Opens an .html/.htm, .md/.markdown, .rtf, .docx/.doc, .odt or .txt file (unsaved changes are discarded).
	/// Word and OpenDocument files are converted by Word, openxml, LibreOffice or pandoc (see Converters and
	/// LastConverter); via names the converter(s) to use, e.g. "openxml" (Save then uses them too).
	public void Open(string path, string via = "") {
		string full = Path.GetFullPath(path);
		// (iteration 68: not painted until its sheets are in, so it shows once, where it stays)
		using (form.HoldOpening(Path.GetFileName(full))) {
			Load(full, via);
			ReleaseTom();
			form.SheetsNow();
			form.OutlineNow();
		}
		form.Events.Raise(EditorEvent.Opened);
		form.RunHook(HookEvent.Opened, full);    // (iteration 56)
	}

	void Load(string full, string via) {
		string extension = Path.GetExtension(full).ToLowerInvariant();
		LastConverter = "";
		(ed as WriterEditBox)?.AdvancedLayout(false);    // (the document's RTF turns it on if it needs it)
		// the header and footer: the document's own; documents that can have them but don't have none, others get the default
		PageBands? bands = null;
		LineNumbering? numbering = null;
		bool canHaveBands = extension is ".rtf" or ".docx" or ".doc" or ".odt";
		string? sourceText = null;    // a .md / .html file's own text, for the source view (and saving it as it was)
		form.Styles = form.DefaultStyles();
		form.Levels = form.DefaultLevels();    // (AhkWord's RTF: its own, as the Rtf setter finds them)
		form.UseNormalFont();    // (text without a font of its own, as HTML's and Markdown's: Normal's)
		(int[]? Margins, bool? Landscape, int[]? Paper) pageSetup = (null, null, null);    // (iteration 63: the file's, where it gives one)
		SectionColumns? fileColumns = null;                            // (iteration 67: its first section's)
		StyleSheet? fileStyles = null;
		bool takeLook = false, styledLook = false;
		using (Quietly()) {
		switch (extension) {
			case ".md": case ".markdown": {
				string markdown = sourceText = File.ReadAllText(full);
				bands = HeaderFooter.FromMarkdown(markdown);
				numbering = LineNumberFormats.FromMarkdown(markdown);
				FreshEditor();
				LoadMarkdown(markdown, Path.GetDirectoryName(full));
				takeLook = true;    // (Markdown has no look of its own: its styles')
				break;
			}
			case ".txt": { string text = File.ReadAllText(full).Replace("\r\n", "\n"); FreshEditor(); ed.Text = text; takeLook = true; break; }
			case ".rtf": {
				string rtf = File.ReadAllText(full, Encoding.Latin1);
				bands = HeaderFooter.FromRtf(rtf);
				numbering = LineNumberFormats.FromRtf(rtf);
				pageSetup = PageSetupOf(rtf);
				fileColumns = Sections.FirstColumnsFromRtf(rtf);
				rtf = Sections.ImportRtf(rtf);    // (iteration 63: Word's sections as AhkWord's section breaks)
				FreshEditor();
				Rtf = rtf;
				fileStyles = StylesFromRtf(rtf);
				LegalLists.AdoptTypedNumbers(ed);
				LinkFonts(ed);
				break;
			}
			case ".docx": case ".doc": case ".odt": {
				var (kind, content, converter) = OfficeConverters.Import(full, via);
				bands = extension == ".docx" ? HeaderFooter.FromDocx(full) : extension == ".odt" ? HeaderFooter.FromOdt(full) : kind == "rtf" ? HeaderFooter.FromRtf(content) : null;
				numbering = extension == ".docx" ? LineNumberFormats.FromDocx(full) : extension == ".odt" ? LineNumberFormats.FromOdt(full) : kind == "rtf" ? LineNumberFormats.FromRtf(content) : null;
				// (iteration 63) the file's page setup, and its sections
				pageSetup = kind == "rtf" ? PageSetupOf(content) : extension == ".docx" ? PageSetupOfDocx(full) : (null, null, null);
				fileColumns = kind == "rtf" ? Sections.FirstColumnsFromRtf(content) : extension == ".docx" ? HeaderFooter.FirstColumnsFromDocx(full) : null;
				FreshEditor();
				if (kind == "rtf") {
					content = Sections.ImportRtf(content);
					Rtf = content;
					fileStyles = StylesFromRtf(content);
					// (a .docx's styles as the file defines them, where AhkWord's own aren't there: else a style's look
					// would be its first paragraph's, as the converter left it, such as a numbered style's first item
					// with the list's indents)
					if (fileStyles == null && extension == ".docx")
						try {
							if (DocxStyles.Read(full).Sheet is { Styles.Count: > 0 } read) {
								// (in the order the document uses them, as styles taken from its paragraphs are)
								var used = StyleEngine.Paragraphs(ed.TextDocument, 0, ed.TextDocument.Range2(0, 0).StoryLength)
									.Select(p => StyleEngine.NameOf(ed.TextDocument.Range2(p.Start, p.Start).Para.Style)).Distinct().ToList();
								read.Styles = read.Styles.OrderBy(s => used.FindIndex(n => string.Equals(n, s.Name, StringComparison.OrdinalIgnoreCase)) is var at && at < 0 ? int.MaxValue : at).ToList();
								fileStyles = read;
							}
						} catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException) { }
				}
				else {
					SetHtml(content);
					OfficeConverters.HiddenFromMarkers(ed);    // (iteration 87: hidden text, which HTML can't carry into the editor)
					if (extension == ".docx" && converter is OfficeConverters.OpenXml or OfficeConverters.Docxodus) Sections.FromDocxMarkers(ed, HeaderFooter.SectionsFromDocx(full));
					// (the built-in reader's: the .docx's styles, DocxStyles; its text has its own look already, every run's
					// font as the file has it, which a style's look mustn't take over)
					fileStyles = StyleHtml.FromMeta(content);
				}
				LegalLists.ConvertTypedLists(ed, 0, ed.TextDocument.Range2(0, 0).StoryLength, requireTab: true,
					converter is OfficeConverters.OpenXml or OfficeConverters.Docxodus || kind == "html" ? OfficeConverters.TypedIndents : null);    // numbers a converter left as text
				LastConverter = converter;
				// tables' borders, shading, padding ... from the file itself, where the converter went through HTML
				if (extension == ".docx" && converter is OfficeConverters.OpenXml or OfficeConverters.Pandoc or OfficeConverters.Docxodus) RtfTables.TablesFromDocx(ed, full);
				// paragraphs spaced as Word spaces them (RTF from Word or LibreOffice: RichEdit can't leave out the space
				// between paragraphs of the same style, nor give cells their table style's spacing)
				if (extension == ".docx" && kind == "rtf" && !DocxLineSpacing.IntoEditor(ed, full))
					form.SetStatus(T("status.spacingNotMatched", "The paragraphs' spacing could not be matched to the file's; it is as the converter made it."));
				if (kind == "rtf") LinkFonts(ed);
				// (a .docx's list levels: where its lists put their numbers and text, from its most used list definition;
				// the document's lists and legal numbering take them)
				if ((extension == ".docx" ? OfficeConverters.ListLevelIndents(full) : extension == ".odt" && kind == "html" ? OfficeConverters.ConvertedListIndents : null) is { } indents) {
					for (int i = 0; i < 9; i++) (form.Levels[i].NumberAt, form.Levels[i].TextAt) = indents[i];
					OfficeConverters.FitListIndents(ed.TextDocument);
					if (LegalLists.Paragraphs(ed.TextDocument).Any(p => p.Legal && !p.Styled)) LegalLists.Renumber(ed);
				}
				if (kind != "rtf") OfficeConverters.ListMarksFromText(ed);    // (iteration 87: list numbers in their text's font)
				if (extension == ".docx") ParagraphFrames.FromDocx(ed, full);    // (iteration 87: paragraph borders and shading)
				if (extension == ".docx") ListSuffixes.FromDocx(ed, full);    // (iteration 87: a space or nothing after list numbers)
				break;
			}
			default: {
				string html = File.ReadAllText(full);
				if (extension is ".html" or ".htm") sourceText = html;
				bands = HeaderFooter.FromHtml(html);
				numbering = LineNumberFormats.FromHtml(html);
				FreshEditor();
				SetHtml(html);
				if (StyleHtml.LevelsIn(html) is { Count: 9 } levels) {    // (AhkWord's HTML: where its lists put their numbers and text)
					form.Levels = levels;
					OfficeConverters.FitListIndents(ed.TextDocument);
				}
				fileStyles = StyleHtml.FromMeta(html);    // (AhkWord's HTML: its styles, which its styled paragraphs are given again)
				styledLook = fileStyles != null;
				break;
			}
		}
		ReleaseTom();    // (the document is in: what reading it made)
		form.Bands = bands ?? (canHaveBands ? new PageBands() : form.DefaultBands());
		form.SetLineNumbering(numbering ?? new LineNumbering(), changesDocument: false);
		form.SetDocumentPageSetup(pageSetup.Margins, pageSetup.Landscape, fileColumns, pageSetup.Paper);
		form.StylesAfterLoad(fileStyles, takeLook, styledLook);
		ReleaseTom();    // (a range and a paragraph format, and more, for every paragraph)
		PageBreaks.Normalize(ed);    // (form feeds in plain text, for one)
		MarksFollowText();
		if (LegalLists.AnyLegal(ed.TextDocument)) LegalLists.Renumber(ed);    // (legal items' indents fitted to their numbers' size)
		// (iteration 50) the cross-references' targets where they belong now that the lists are made, and their numbers
		// as the editor numbers the paragraphs (Word's and LibreOffice's REF fields come with theirs)
		if (CrossRefs.AfterLoad(ed.TextDocument) > 0) CrossRefs.Update(ed.TextDocument);
		// (iteration 51) the notes' size the file had (Word's, LibreOffice's): the document's Footnote Text style
		if (Footnotes.LoadedSize is float loaded && loaded >= 4 && form.Styles.Find(Footnotes.StyleName) == null && Footnotes.Any(Story())) form.SetNoteStyle(loaded);
		Footnotes.LoadedSize = null;
		ListSuffixes.ToStarts(ed.TextDocument);    // (iteration 87: the suffix markers where they belong, before the restarts)
		OfficeConverters.RestartsFromMarkers(ed);    // (iteration 87: lists right after lists, numbered from their own start)
		OfficeConverters.NewNumbersInto(ed, loadedNewNumbers);    // (and as AhkWord's .rtf kept them)
		loadedNewNumbers = null;
		ListSuffixes.Apply(ed.TextDocument);    // (iteration 87: list paragraphs with a space or nothing after their numbers, as AhkWord's .rtf keeps them)
		}
		ed.ClearUndo();
		ed.Select(0, 0);
		ed.Modified = false;
		FileConverter = extension is ".docx" or ".doc" or ".odt" ? via : "";
		// an HTML file with its pictures inside it is saved that way again
		if (extension is ".html" or ".htm" && sourceText != null && sourceText.Contains("data:image/", StringComparison.OrdinalIgnoreCase)) FileConverter = PicturesInside;
		form.SetFileName(full);
		form.AddRecent(full);
		form.DocumentLoaded(sourceText == null ? WriterForm.SourceKind.None : extension is ".md" or ".markdown" ? WriterForm.SourceKind.Markdown : WriterForm.SourceKind.Html, sourceText);
	}

	// While a document is loaded: no undo steps (they are cleared once it is), and no change notifications, which would
	// update the window (the word count, the outline, the line numbers ...) after each of the many changes a file's
	// formats are put in with; one notification for the whole once it is loaded.
	IDisposable Quietly() {
		quiet = true;
		ed.TextDocument.Undo(-9999995);    // tomSuspend
		IntPtr mask = SendQuiet(ed.Handle, EM_SETEVENTMASK, IntPtr.Zero, IntPtr.Zero);
		return new Restore(() => {
			quiet = false;
			SendQuiet(ed.Handle, EM_SETEVENTMASK, IntPtr.Zero, mask);
			ed.TextDocument.Undo(-9999994);    // tomResume (the editor's RichEdit now: FreshEditor may have made a new one)
			typeof(Control).GetMethod("OnTextChanged", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(ed, new object[] { EventArgs.Empty });
		});
	}

	sealed class Restore(Action restore) : IDisposable { public void Dispose() => restore(); }

	// (2026-11-08) The TOM objects (ranges, fonts, paragraph formats) that opening a document made, let go of now. .NET
	// releases them only when it collects, and until then each range is one RichEdit moves on every change to the text:
	// a long document's loading made tens of thousands, and the steps after it slowed (the first .md open of the 105-page
	// dissertation 8-9 s, against 2.6 s after); when a collection then came, releasing them all at once took seconds more
	// (sampled: CTxtRange::~CTxtRange, CTxtRange::OnPostReplaceRange). Let go of at a few points while they are few,
	// they cost 10-100 ms (a .NET collection of a small heap).
	static void ReleaseTom() {
		GC.Collect();
		GC.WaitForPendingFinalizers();
		Marshal.CleanupUnusedObjectsInCurrentContext();    // (the editor's thread's, released here)
	}
	const int EM_SETEVENTMASK = 0x0445;
	bool quiet;

	// (2026-11-08) The window's editor emptied into a new RichEdit (WriterForm.FreshEditor), just before Load or New puts
	// the document in: after the file has been read and converted, so a file that can't be leaves the document as it was.
	// Inside Quietly the new one is quiet too.
	void FreshEditor() {
		if (ed != form.Editor) return;
		form.FreshEditor();
		(ed as WriterEditBox)?.AdvancedLayout(false);    // (the document's RTF turns it on if it needs it)
		if (!quiet) return;
		ed.TextDocument.Undo(-9999995);    // tomSuspend
		SendQuiet(ed.Handle, EM_SETEVENTMASK, IntPtr.Zero, IntPtr.Zero);
	}
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendQuiet(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	/// Saves to the current file and returns its path.
	public string Save() =>
		SaveAs(form.FileName ?? throw new InvalidOperationException("The document has not been saved yet; use SaveAs(path)."), FileConverter);

	/// The converter(s) chosen for the current Word or OpenDocument file when it was opened or saved ("" = automatic);
	/// Save uses them.
	public string FileConverter { get; private set; } = "";

	/// Saves as HTML, Markdown, RTF, Word, OpenDocument or text, chosen by the extension (.html, .md, .rtf, .docx, .odt,
	/// .txt), and returns the full path. .docx and .odt are converted (see Converters and LastConverter); via names the
	/// converter(s) to use, e.g. "openxml". HTML and Markdown pictures are saved as files in a <name>_images folder;
	/// via = "inside" keeps an HTML file's pictures inside it (as data: URIs), as one file.
	public string SaveAs(string path, string via = "") {
		string full = Path.GetFullPath(path);
		form.RunHook(HookEvent.BeforeSave, full);    // (iteration 56: waits for its macro)
		string extension = Path.GetExtension(full).ToLowerInvariant();
		LastConverter = "";
		// HTML and Markdown keep their pictures in document_images\ next to the file, except HTML saved with "inside"
		bool picturesInside = extension is ".html" or ".htm" && via == PicturesInside;
		// a .md / .html file whose document came from its source text, unchanged since, gets that text as it was typed
		// (iteration 46: the file made from the document as it is, without the page view's gaps, whatever step reads it)
		form.UpdateReferences();    // (iteration 50: the cross-references' numbers as they are now)
		form.WithoutGaps(() => ListSuffixes.WithFileIndents(ed, () => WriteFile(full, extension, via, picturesInside)));    // (iteration 87: list paragraphs' own indents)
		ed.Modified = false;
		form.ClearRecovery();
		FileConverter = extension is ".docx" or ".odt" || picturesInside ? via : "";
		form.SetFileName(full);
		form.AddRecent(full);
		form.Events.Raise(EditorEvent.Saved);
		form.RunHook(HookEvent.Saved, full);    // (iteration 56)
		return full;
	}

	void WriteFile(string full, string extension, string via, bool picturesInside) {
		if (form.VerbatimSource(extension) is { } verbatim)
			File.WriteAllText(full, picturesInside ? verbatim : PicturesToFiles(verbatim, full), new UTF8Encoding(false));
		else if (extension is ".docx" or ".odt") Export(full, via);
		// (ordinary lists with hanging indents that fit their numbers: RichEdit's own RTF indents them so that Word and
		// LibreOffice put the numbers too far in)
		else if (extension == ".rtf")    // (and pictures as PNG, and the header and footer)
			File.WriteAllText(full, global::LineSpacing.WordLines(WithBands(OfficeConverters.PngPictures(OfficeConverters.WithAltTexts(StyledRtf(OfficeConverters.ListIndentsRtf(RawRtf)), AltTexts())), full)), Encoding.Latin1);    // (list indents before the styles: RecoveryRtf; iteration 87: Word's 1.08 lines again)
		else {
			string content = extension switch {
				".md" or ".markdown" => form.WithoutGaps(() => Footnotes.ToMarkdown(CrossRefs.ToMarkdown(MarkdownConverter.ToMarkdown(ed, ImageFiles(full))))) + (OwnBands ? HeaderFooter.MarkdownLine(form.Bands) + "\n" : "") + LineNumberFormats.MarkdownLine(form.Numbering),
				".txt" => CrossRefs.PlainText(ed.Text).Replace("\n", "\r\n"),
				_ => picturesInside ? HtmlFile(Path.GetFileNameWithoutExtension(full), pretty: true)
					: PicturesToFiles(HtmlFile(Path.GetFileNameWithoutExtension(full), pretty: true), full)
			};
			File.WriteAllText(full, content, new UTF8Encoding(false));
		}
	}

	/// The full path of the current file ("" for a new document).
	public string FileName => form.FileName ?? "";

	/// The document as HTML, with list numbering that RichEditorNET's HTML leaves out (start numbers, restarts,
	/// punctuation such as "a)"; the punctuation needs the <style> element at the start).
	public string GetHtml() {
		string body = HtmlBody(out string css);
		return HtmlLists.StyleElement(css) + body;
	}

	/// Replaces the document with HTML (a whole page or a part of one), including the list numbering and tables GetHtml writes.
	public void SetHtml(string html) {
		// (iteration 68: built in a hidden editor, which lays nothing out for the page and has no spell checker or text services
		// following it, and handed over at once: the many changes a file's formats take were up to minutes in the window's own)
		using var hidden = AppTheme.OnPaper(new WriterEditBox { WordWrap = false, Width = 600, Height = 400, Font = ed.Font });
		_ = hidden.Handle;
		hidden.EnableFontSize = ed.EnableFontSize;
		hidden.SetDefaultSize();
		hidden.TextDocument.Undo(tomConstants.tomSuspend);    // (no undo kept while it is built: about a second of a 100-page document's opening)
		building = hidden;
		try { BuildHtml(html); }
		finally { building = null; }
		building = hidden;
		MarksFollowText();    // (here: thousands of looks at single characters, seconds in the window)
		OfficeConverters.FitListIndents(hidden.TextDocument);    // (HTML's lists come without them: number and text apart, as RTF's and Word's)
		building = null;
		marksFollowed = true;
		// (the paragraphs' styles and list levels aren't all in RichEdit's RTF; by paragraph: the markers are still in this text)
		var styled = new Dictionary<int, int>();
		var listed = new Dictionary<int, (int Type, int Start, int Level, float Tab)>();
		var breaks = new Dictionary<int, string>();
		EachParagraph(hidden, (index, paragraph) => {
			var para = paragraph.Para;
			if (para.Style is int style and not -1) styled[index] = style;
			if (global::LineSpacing.BreaksIn(paragraph.Text ?? "") is { } kept) breaks[index] = kept;
			if ((para.ListType & 0xFFFF) != tomConstants.tomListNone) listed[index] = (para.ListType, para.ListStart, para.ListLevelIndex, para.ListTab);
		});
		string rtf = ((RichTextBox)hidden).Rtf ?? "";
		StreamRtf(global::LineSpacing.MarkersToRtf(rtf));
		global::LineSpacing.FromMarkers(ed);    // (any that didn't come through in one piece, split by a format)
		var lostNumbers = new List<(int Start, int Length)>();
		if (listed.Count + breaks.Count > 0) EachParagraph(ed, (index, paragraph) => {
			var para = paragraph.Para;
			if (breaks.TryGetValue(index, out var kept)) global::LineSpacing.SetBreaks(para, kept);
			if (!listed.TryGetValue(index, out var list)) return;
			// (iteration 87) a list paragraph whose list RichEdit's RTF lost on the way (after legal items: its \pn group),
			// which kept its number as text instead ("(A)<tab>"): the text goes, the list comes back below
			if ((para.ListType & 0xFFFF) == tomConstants.tomListNone && LostListNumber.Match(paragraph.Text ?? "") is { Success: true } typed)
				lostNumbers.Add((paragraph.Start, typed.Length));    // (taken out after the loop, which reads the story's end once)
			if (para.ListType != list.Type) para.ListType = list.Type;
			if (para.ListStart != list.Start) para.ListStart = list.Start;
			if (para.ListLevelIndex != list.Level) para.ListLevelIndex = list.Level;
			if (Math.Abs(para.ListTab - list.Tab) > 0.01f) para.ListTab = list.Tab;
		});
		foreach (var (start, length) in Enumerable.Reverse(lostNumbers)) ed.TextDocument.Range2(start, start + length).Text = "";
		if (styled.Count > 0) EachParagraph(ed, (index, paragraph) => { if (styled.TryGetValue(index, out int style)) {
			// (a built-in style puts its own spacing on the paragraph: the paragraph's stay)
			var para = paragraph.Para;
			var (before, after, rule, line, next, together, widow) = (para.SpaceBefore, para.SpaceAfter, para.LineSpacingRule, para.LineSpacing, para.KeepWithNext, para.KeepTogether, para.WidowControl);
			var (alignment, first, left, right) = (para.Alignment, para.FirstLineIndent, para.LeftIndent, para.RightIndent);
			para.Style = style;
			para.Alignment = alignment;
			para.SetIndents(first, left, right);
			(para.SpaceBefore, para.SpaceAfter, para.KeepWithNext, para.KeepTogether, para.WidowControl) = (before, after, next, together, widow);
			para.SetLineSpacing(rule, line);
		} });
	}

	static readonly Regex LostListNumber = new(@"^(?:\(?(?:\d{1,4}|[a-zA-Z]{1,6})(?:\.\d{1,4})*[.)]?|[•▪o·])	");

	// RichEdit's own RTF into the window's editor, as the whole document. (2026-11-08) A one-paragraph document goes in
	// first: once the editor's display had been woken up (a document opened through Word, the second time in a session),
	// it followed the RTF reader's every insertion, laying out its lines again each time (msftedit's
	// CDisplayML::OnPreReplaceRange; for the printer, in Print Layout): 1.4-1.6 s of a 105-page document's RTF against
	// 0.1 s. Streamed in just before, the small document leaves the display as a fresh editor's, which lays the lines
	// out once, afterwards (measured: 0.1 s again, for 10-20 ms). Neither TOM's Freeze nor emptying the text helped.
	void StreamRtf(string rtf) {
		var box = (RichTextBox)ed;
		box.Rtf = @"{\rtf1\par}";
		box.Rtf = rtf;
	}

	static void EachParagraph(RichEditBox editor, Action<int, ITextRange2> each) {
		var doc = editor.TextDocument;
		int end = doc.Range2(0, 0).StoryLength, index = 0;
		var paragraph = doc.Range2(0, 0);
		paragraph.Expand(tomConstants.tomParagraph);
		while (true) {
			each(index++, paragraph);
			if (paragraph.End >= end) break;
			paragraph.SetRange(paragraph.End, paragraph.End);
			paragraph.Expand(tomConstants.tomParagraph);
		}
	}

	void BuildHtml(string html) {
		// (frozen: the markers put into the text, and taken out one by one, aren't shown)
		ed.TextDocument.Freeze();
		try {
			RtfTables.LoadHtml(ed, BodyOf(LineBreaks.MarkHtml(PageBreaks.MarkHtml(global::LineSpacing.MarkHtml(StyleHtml.MarkHtml(CrossRefs.FromHtml(Footnotes.FromHtml(html))))))), HtmlLists.Load, HtmlLists.Load, TableWidth());
			RtfTables.DropStrayMergeMarks(ed.TextDocument);    // (iteration 92)
			CrossRefs.AfterLoad(ed.TextDocument);
			StyleHtml.FromMarkers(ed);    // (headings and styles from their tags, not their fonts)
			if (LineBreaks.FromMarkers(ed) + PageBreaks.FromMarkers(ed) + PageBreaks.Normalize(ed) + (building == null ? global::LineSpacing.FromMarkers(ed) : 0) > 0) ed.ClearUndo();
		} finally {
			ed.TextDocument.Unfreeze();
		}
	}

	// A link's hidden field instruction (HYPERLINK "...") in its text's font: RTF from Word and LibreOffice gives it the
	// document's default size, and the HTML (and so the built-in converter's .docx) took the link's size from it.
	static void LinkFonts(RichEditBox editor) {
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		foreach (Match m in Regex.Matches(story, "HYPERLINK\\s+\"[^\"]*\"\\s?")) {
			int end = m.Index + m.Length;
			if (end >= story.Length || doc.Range2(m.Index, m.Index + 1).Font.Hidden != tomConstants.tomTrue) continue;
			var text = doc.Range2(end, end + 1).Font;
			var instruction = doc.Range2(m.Index, end).Font;
			if (instruction.Size != text.Size) instruction.Size = text.Size;
			if (text.Name is { Length: > 0 } name && instruction.Name != name) instruction.Name = name;
		}
	}

	// The library's HTML, with page breaks, real tables (it writes their cells as paragraphs) and list numbering.
	// (iteration 46: all of it without the page view's gaps: the steps after the library's HTML read the paragraphs'
	// spacing from the editor, and wrote Show Pages' gaps as huge margins before a page's first paragraph)
	string HtmlBody(out string css) {
		string listsCss = "";
		string body = form.WithoutGaps(() => {
			string html = HtmlLists.Export(ed, RtfTables.ExportHtml(ed, PageBreaks.ExportHtml(ed, global::LineSpacing.ExportHtml(ed, StyleHtml.Export(ed, form.Styles, LineBreaks.ExportHtml(ed, ed.Html))))), out var found);
			listsCss = found;
			return MarkdownConverter.RulesToHtml(Footnotes.ToHtml(CrossRefs.ToHtml(html)));
		});
		css = ParagraphCss + listsCss;
		return body;
	}

	// Paragraphs as the editor shows them: browsers put a gap between paragraphs (and headings), and give an empty
	// paragraph no height, where the editor has neither gaps nor vanishing blank lines.
	const string ParagraphCss = "p, h1, h2, h3, h4, h5, h6 { margin: 0; }\r\np:empty::before { content: \"\\00a0\"; }\r\n";

	/// The document as HTML (get), or replaced by HTML (set); see GetHtml and SetHtml.
	public string Html {
		get => GetHtml();
		set => SetHtml(value);
	}

	// The document as an HTML file. pretty: one block per line, lists and tables indented (the HTML view and .html files;
	// conversions use it as it comes).
	string HtmlFile(string title, bool pretty = false) {
		string body = HtmlBody(out string css);
		return HtmlDocument(pretty ? PrettyHtml.Format(body) : body, title, css, (OwnBands ? HeaderFooter.HtmlMeta(form.Bands) : "") + LineNumberFormats.HtmlMeta(form.Numbering), StyleHtml.EndComment(form.Styles) + StyleHtml.LevelsComment(form.Levels));
	}

	/// The "via" for an HTML file that keeps its pictures inside it.
	public const string PicturesInside = "inside";

	// Pictures as data: URIs (HTML src, Markdown links) saved as files in document_images\ and referred to from there.
	static string PicturesToFiles(string text, string path) {
		var file = ImageFiles(path);
		int number = 0;
		return WriterForm.DataUri.Replace(text, m => {
			string type = m.Groups["type"].Value.ToLowerInvariant();
			string extension = type switch { "jpeg" => "jpg", "svg+xml" => "svg", _ => type };
			return file(++number, Convert.FromBase64String(m.Value[(m.Value.IndexOf(',') + 1)..]), extension);
		});
	}

	// Whether the document's header and footer differ from the default (HTML and Markdown keep them only then: a
	// document that has the default gets whatever the default is when it is opened).
	bool OwnBands => !form.Bands.SameAs(form.DefaultBands());

	// RTF with the document's header, footer, page layout and line numbering, for path.
	// (iteration 63: its section breaks as Word's sections)
	string WithBands(string rtf, string path) {
		var layout = form.PageLayout();
		rtf = HeaderFooter.AddToRtf(Footnotes.ToWordRtf(CrossRefs.ToWordRtf(ListSuffixes.ToWordRtf(ParagraphFrames.ToWordRtf(rtf))), ed.Font.Name, ed.Font.SizeInPoints, form.NoteStyle?.Size ?? 0), form.Bands, layout, Path.GetFileName(path), path, ed.Font.SizeInPoints);
		rtf = Sections.ExportRtf(rtf, layout, Path.GetFileName(path), path, ed.Font.SizeInPoints, form.Numbering);
		return LineNumberFormats.AddToRtf(rtf, form.Numbering);
	}

	// ---- Rich Text Format, Word and OpenDocument

	// RichEdit's own RTF (without the pictures' alt texts, for changes that go through a RichEdit).
	string RawRtf => form.WithoutGaps(() => OfficeConverters.WithNewNumbers(((RichTextBox)ed).Rtf ?? "", ed.TextDocument));    // (iteration 87: with the lists started again, for the steps that go through a RichEdit)    // (the page view's gaps aren't the document's)

	// RTF with the document's styles: the style sheet and each paragraph's (RichEdit keeps only the headings').
	// (and the numbered items' ListStart values: where their numbering starts again)
	string StyledRtf(string rtf) {
		string styled = RtfStyles.Inject(rtf, StyleEngine.Names(ed.TextDocument, form.Styles), form.Styles, StyleEngine.WidowControls(ed.TextDocument));
		styled = form.WithoutGaps(() => OfficeConverters.WithNewNumbers(styled, ed.TextDocument));    // (iteration 87: the lists started again)
		// (and the paragraphs whose numbering is off: LegalLists.OffGroup)
		string starts = form.WithoutGaps(() => LegalLists.StartsGroup(ed.TextDocument) + LegalLists.OffGroup(ed.TextDocument)) + ListLevelsGroup();
		int end = styled.LastIndexOf('}');
		return starts == "" || end < 0 ? styled : styled.Insert(end, starts);
	}

	// Iteration 38: the document's list levels in its RTF, {\*\rtewriterlevels base64-JSON} (others ignore it).
	string ListLevelsGroup() =>
		@"{\*\rtewriterlevels " + Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(form.Levels))) + "}";

	static readonly Regex ListLevelsGroupRx = new(@"\{\\\*\\rtewriterlevels\s+([A-Za-z0-9+/=]+)\}");

	/// The list levels an RTF kept (null: none, or not readable); always 9.
	static List<ListLevel>? ListLevelsFrom(string rtf) {
		var m = ListLevelsGroupRx.Match(rtf);
		if (!m.Success) return null;
		try {
			var levels = JsonSerializer.Deserialize<List<ListLevel>>(Encoding.UTF8.GetString(Convert.FromBase64String(m.Groups[1].Value)));
			if (levels == null) return null;
			while (levels.Count < 9) levels.Add(ListLevel.Defaults()[levels.Count]);
			return levels.Take(9).ToList();
		} catch (Exception ex) when (ex is FormatException or JsonException) { return null; }
	}

	// The styles an RTF just loaded has (Word's, LibreOffice's, AhkWord's): each paragraph's set from its \sN; returns
	// AhkWord's own definitions (null if the file has none). An RTF without any styles (written by AhkWord before it
	// had them, or by WordPad) gets its headings from their fonts, as before, the large ones only (bold 24, 18 and 14 pt:
	// a bold 12 pt line is a heading's look only by chance).
	List<(int Index, int Start)>? loadedNewNumbers;

	StyleSheet? StylesFromRtf(string rtf) {
		form.LoadedListStarts = LegalLists.StartsFrom(rtf);    // (StylesAfterLoad gives them back to the numbered items)
		loadedNewNumbers = OfficeConverters.NewNumbersFrom(rtf);    // (iteration 87: the lists started again; given back last of all)
		var (sheet, names, paragraphs, any) = RtfStyles.Read(rtf);
		var doc = ed.TextDocument;
		LegalLists.RestoreOff(doc, rtf);    // (the paragraphs whose numbering is off, as AhkWord's RTF kept them)
		if (any) {
			string Name(int number) => names.TryGetValue(number, out var name) ? (ParagraphStyle.HeadingNumber(name) is int n and > 0 ? "Heading " + n : name) : "Normal";
			StyleEngine.SetNames(doc, paragraphs.Select(Name).ToList());
			return sheet;
		}
		ForEachParagraph((start, text) => {
			if (text.Trim() == "" || text.IndexOfAny(new[] { '\u0007', '\uFFF9', '\uFFFB' }) >= 0) return;
			var font = doc.Range2(start, start + 1).Font;
			bool bold = font.Bold == tomConstants.tomTrue;
			int level = !bold ? 0 : Math.Abs(font.Size - 24) < 0.1 ? 1 : Math.Abs(font.Size - 18) < 0.1 ? 2 : Math.Abs(font.Size - 14) < 0.1 ? 3 : 0;
			if (level > 0) StyleEngine.Tag(doc.Range2(start, start).Para, -(level + 1));
			else if (!bold && Math.Abs(font.Size - 10) < 0.1 && string.Equals(font.Name, "Courier New", StringComparison.OrdinalIgnoreCase)) StyleEngine.Tag(doc.Range2(start, start).Para, StyleEngine.IdOf("Code"));
		});
		return null;
	}

	/// The document as RTF (get), or replaced by RTF (set). RTF keeps everything the editor has.
	public string Rtf {
		get => OfficeConverters.WithAltTexts(RawRtf, AltTexts());
		set {
			if (ListLevelsFrom(value) is { } levels) form.Levels = levels;    // (the document's list levels, as AhkWord's RTF kept them)
			value = OfficeConverters.SimplePictures(value);    // (Word's picture wrappers, which RichEdit can drop)
			value = OfficeConverters.WithoutUnlistedNumbering(value, out var wordStarts);    // (Word's paragraphs in no list, which RichEdit would number)
			value = HeaderFooter.StripFromRtf(value);             // (the header and footer are the document's Bands)
			value = CrossRefs.FromWordRtf(value);                 // (iteration 50: Word's REF fields and bookmarks, which RichEdit drops)
			value = Footnotes.FromWordRtf(value);                 // (iteration 51: Word's footnotes, which RichEdit shows in the text)
			value = RtfTables.WordShading(value);                 // (Word's cell backgrounds, which RichEdit ignores)
			StreamRtf(OfficeConverters.WithoutAltTexts(value));    // (RichEdit would put their accents into the text)
			OfficeConverters.ContinueWordLists(ed.TextDocument, wordStarts);    // (the numbers after them, as Word has them)
			PageBreaks.Normalize(ed);    // (Word's \page can be in the middle of a paragraph)
			RestorePictures(value);       // (with the alt texts, from the RTF as it was)
			LegalLists.LoadedIndents.Clear();
			LegalLists.LevelsFromNumbers(ed.TextDocument);    // (RichEdit's RTF loses the levels of legal items)
			RtfTables.DropStrayMergeMarks(ed.TextDocument);    // (iteration 92: a nested table's first row a line too tall)
			CrossRefs.AfterLoad(ed.TextDocument);
			ed.ClearUndo();
			ed.Select(0, 0);
		}
	}

	// Each embedded object's alt text (an editor picture's; "" for others), in document order. RichEditorNET gives the
	// selected picture's, so each is selected in turn (and the selection put back).
	List<string> AltTexts() {
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var alts = new List<string>();
		if (story.IndexOf('\uFFFC') < 0) return alts;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		bool startActive = (selection.Flags & 1) != 0;    // tomSelStartActive: the caret at the selection's start
		doc.Freeze();
		try {
			for (int i = story.IndexOf('\uFFFC'); i >= 0; i = story.IndexOf('\uFFFC', i + 1)) {
				string alt = "";
				try { selection.SetRange(i, i + 1); alt = ed.GetImageAltText() ?? ""; }
				catch (Exception ex) when (ex is COMException or InvalidOperationException or NullReferenceException) { }
				alts.Add(alt.Trim());
			}
		} finally {
			selection.SetRange(start, end);
			if (startActive && end > start) selection.Flags |= 1;
			doc.Unfreeze();
		}
		return alts;
	}

	// Pictures that come from RTF are plain RichEdit pictures, which RichEditorNET's HTML (and so Markdown and the
	// built-in .docx converter) leaves out: each becomes an editor picture again, made from the image in the RTF, at
	// its size and with its link.
	void RestorePictures(string rtf) {
		var pictures = OfficeConverters.Pictures(rtf);
		if (pictures.Count == 0) return;
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var positions = story.Select((c, i) => (c, i)).Where(x => x.c == '\uFFFC').Select(x => x.i).ToList();
		if (positions.Count != pictures.Count) return;    // other embedded objects as well: they can't be told apart
		int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
		try {
			for (int k = positions.Count - 1; k >= 0; k--) {
				if (pictures[k] is not { } picture) continue;
				int at = positions[k];
				string url = Regex.Match(doc.Range2(at, at + 1).URL ?? "", @"""?([^""]+)""?").Groups[1].Value;
				int before = doc.Range2(0, 0).StoryLength;
				var selection = doc.Selection2;
				selection.SetRange(at, at + 1);
				selection.Text = "";
				// a picture that was a link's only text takes the link's hidden address with it: back to where the link began
				at -= Math.Max(0, before - doc.Range2(0, 0).StoryLength - 1);
				selection.SetRange(at, at);
				ed.DefaultImageWidth = picture.Width;
				ed.DefaultImageHeight = picture.Height;
				using var stream = new MemoryStream(picture.Image);
				ed.InsertImage(stream, picture.Alt);
				if (url != "") doc.Range2(at, at + 1).URL = "\"" + url.Replace("\"", "%22") + "\"";
			}
		} catch (Exception ex) {
			form.SetStatus(T("status.picturesNotEditable", "Some pictures could not be made editor pictures again: {0}", ex.Message));
		} finally {
			ed.DefaultImageWidth = oldWidth;
			ed.DefaultImageHeight = oldHeight;
		}
	}

	/// Writes a copy of the document as Word (.docx); the current file name does not change. via names the converters
	/// to try, e.g. "openxml" or "word,pandoc" (default: Word if it isn't running, openxml, LibreOffice, pandoc).
	/// Returns the full path.
	public string ExportDocx(string path, string via = "") => Export(Path.ChangeExtension(Path.GetFullPath(path), ".docx"), via);

	/// Writes a copy of the document as OpenDocument text (.odt), like ExportDocx (default order: LibreOffice, Word if
	/// it isn't running, pandoc).
	public string ExportOdt(string path, string via = "") => Export(Path.ChangeExtension(Path.GetFullPath(path), ".odt"), via);

	string Export(string full, string? via) {
		OfficeConverters.DefaultFont = (ed.Font.Name, ed.Font.SizeInPoints);
		var alts = AltTexts();
		OfficeConverters.LegalItems = LegalLists.Items(ed.TextDocument);    // (numbered styles' formats and levels, which the RTF has lost)
		try {
			LastConverter = OfficeConverters.Export(full, RawRtf, () => HtmlFile(Path.GetFileNameWithoutExtension(full)), via,
				rtf => WithBands(OfficeConverters.WithAltTexts(StyledRtf(rtf), alts), full));    // (alt texts last: RichEdit, which some changes go through, mangles them; styles too)
		} finally {
			OfficeConverters.LegalItems = null;
		}
		// the header and footer: parts of their own if the converter had no RTF to make them from, and the exact copy
		try {
			if (full.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) HeaderFooter.AddToDocx(full, form.Bands, form.PageLayout(), Path.GetFileName(full), full, parts: true);
			else HeaderFooter.AddToOdt(full, form.Bands, Path.GetFileName(full), full);
			// the line numbering, the same whichever converter made the file (pandoc's and the built-in one's have none)
			if (full.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) LineNumberFormats.AddToDocx(full, form.Numbering, Sections.Breaks(ed.TextDocument).Select(b => b.Settings.LineNumbers).ToList());    // (iteration 87: sections' own)
			else LineNumberFormats.AddToOdt(full, form.Numbering);
			OfficeConverters.AltTextsToFile(full, AltTexts());    // (pictures' alt texts that the converter left out)
			// tables' borders, shading, padding ... in Word's own terms, where the converter went through HTML
			if (full.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) && LastConverter is OfficeConverters.OpenXml or OfficeConverters.Pandoc) {
				RtfTables.TablesToDocx(full, RtfTables.Find(ed.TextDocument).Select(t => RtfTables.Formats(ed, t)).ToList());
				DocxLineSpacing.FromEditor(full, ed);    // (HtmlToOpenXml writes "at least" as exactly)
				DocxStyles.IntoDocx(full, ed, form.Styles);    // (the styles, and each paragraph's)
				// (iteration 50) the cross-references' targets as bookmarks (the converter drops the HTML's empty anchors)
				// (iteration 52: the headings' too, the tables of contents' entries going to them, in the document's order)
				CrossRefs.BookmarksIntoDocx(full, CrossRefs.Targets(ed.TextDocument).Where(t => t.Name != null).Select(t => (t.Start, Name: t.Name!, t.Text))
					.Concat(Toc.Headings(ed.TextDocument, form.Styles, 6).Where(h => h.Name != null).Select(h => (h.Start, Name: h.Name!, Text: h.Bare)))
					.GroupBy(t => t.Name).Select(g => g.First()).OrderBy(t => t.Start).Select(t => (t.Name, t.Text)).ToList());
				Toc.IntoDocx(full);    // (Word's TOC field around the entries)
				Footnotes.IntoDocx(full, Footnotes.All(Story()).Select(f => f.Text).ToList(), form.NoteSizeNow());    // (iteration 51: real footnotes)
			}
		} catch (Exception ex) {
			form.SetStatus("⚠ " + T("status.headerFooterNotAdded", "The header and footer could not be added to {0}: {1}", Path.GetFileName(full), ex.Message));
		}
		return full;
	}

	/// The converter used by the last Open, Save, SaveAs or Export of a Word or OpenDocument file ("word",
	/// "libreoffice", "openxml" or "pandoc"; "" if none was needed).
	public string LastConverter { get; private set; } = "";

	/// Which converters for Word and OpenDocument files are available.
	public string Converters() => OfficeConverters.Describe();

	// ---- Print Layout (iteration 24)

	/// Print Layout (1): lines break where they will on the printed page, the page shown on grey; 0: lines wrap at the
	/// window's width (WordWrap). Remembered.
	public bool PrintLayout {
		get => form.Settings.PrintLayout;
		set {
			form.Settings.PrintLayout = value;
			form.Settings.Save();
			form.ApplyPrintLayout();
		}
	}

	/// Page view (View > Show Pages): in Print Layout, each printed page as a sheet of its own, with its header and footer
	/// and a band between the sheets (1); 0: one long page, a line where each page begins. Remembered.
	public bool PageView {
		get => form.Settings.PageView;
		set { form.Settings.PageView = value; form.Settings.Save(); form.PageViewChanged(); }
	}

	/// The page view as it is now, for a bug report: the program's folder, the zoom, where the printed pages start,
	/// each gap (positions, line spacing) and any raised text that isn't a gap's. Numbers only: none of the text.
	public string PageViewReport() => form.PageViewReport();

	/// Multiple Pages: the printed pages side by side, read-only, in the editor's place (true), or editing (false).
	public bool MultiplePages {
		get => form.PagesShown;
		set => form.ShowPages(value);
	}

	/// Print Layout's zoom: "PageWidth" (the page's width fills the window), "OnePage" (a page's height fits) or
	/// "Manual" (ZoomFactor, as Zoom In / Out / 100% and Ctrl+wheel set it). Remembered.
	/// <summary>The window's theme: "Light", "Dark", "Teal", "Navy", "Forest", "Plum", "Burgundy" or "Bronze". The document page stays white in all of them.</summary>
	public string Theme {
		get => form.Settings.Theme;
		set {
			form.Settings.Theme = AppTheme.Find(value)?.Name ?? throw new ArgumentException($"No theme \"{value}\": {string.Join(", ", AppTheme.All.Select(t => t.Name))}.");
			form.ChangeTheme();
		}
	}

	public string ZoomMode {
		get => form.Settings.ZoomMode;
		set {
			string mode = value.Trim().ToLowerInvariant() switch {
				"pagewidth" or "page width" or "width" => "PageWidth",
				"onepage" or "one page" or "page" => "OnePage",
				"manual" or "none" or "" => "Manual",
				_ => throw new ArgumentException($"\"{value}\" is not a zoom mode (PageWidth, OnePage or Manual).")
			};
			form.Settings.ZoomMode = mode;
			form.Settings.Save();
			form.ApplyPrintLayout();
		}
	}

	// ---- Line spacing (iteration 22)

	/// Sets the line spacing of the selected paragraphs (every paragraph the selection touches): a number of lines
	/// ("1", "1.15", "1.5", "2"), or "exactly 14pt" / "at least 12pt". One undo step.
	public void SetLineSpacing(string spacing) {
		var (rule, value) = global::LineSpacing.Parse(spacing);
		global::LineSpacing.Apply(SelectionRange(), rule, value);
		form.LineGutter.DocumentChanged();
	}

	/// The line spacing of the selected paragraphs: "1.5" (lines), "exactly 14pt", "at least 12pt", or "mixed".
	public string LineSpacing => global::LineSpacing.Of(SelectionRange());

	/// The space before the selected paragraphs, in points (-1 if they differ); setting it sets it for them (one undo step).
	public float SpaceBefore {
		get { using var _ = form.FirstParagraphAsIs(changes: false); return global::LineSpacing.SpaceOf(SelectionRange(), true); }
		set { using var _ = form.FirstParagraphAsIs(changes: true); global::LineSpacing.ApplySpace(SelectionRange(), true, value); form.LineGutter.DocumentChanged(); }
	}

	/// The space after the selected paragraphs, in points (-1 if they differ), as SpaceBefore.
	public float SpaceAfter {
		get => global::LineSpacing.SpaceOf(SelectionRange(), false);
		set { global::LineSpacing.ApplySpace(SelectionRange(), false, value); form.LineGutter.DocumentChanged(); }
	}

	// ---- The Paragraph dialog's formats (iteration 33)

	/// The selected paragraphs' format as "Key=Value" lines, what they share (a key they differ in is left out): Align,
	/// LeftIndent, RightIndent, FirstLine (negative: hanging), Before, After (points), LineSpacing, KeepWithNext,
	/// KeepTogether, WidowControl, PageBreakBefore, NoLineNumber (1 or 0), Tabs (72/right/dots,...), and their style's
	/// Contextual (no space between paragraphs of the same style) and Level (0: body text, 1-6: a heading).
	public string ParagraphFormat() {
		using var _ = form.FirstParagraphAsIs(changes: false);
		var selection = ed.TextDocument.Selection2;
		return global::ParagraphFormat.Read(ed.TextDocument, form.Styles, selection.Start, selection.End).Describe();
	}

	/// Gives the selected paragraphs a format: "Key=Value" lines as ParagraphFormat gives them (or with ; between them),
	/// measures in points or with a unit: Rte.SetParagraphFormat("LeftIndent=0.5in; FirstLine=-0.5in; KeepTogether=1").
	/// What isn't given stays as each paragraph has it. One undo step.
	public void SetParagraphFormat(string settings) {
		using var _ = form.FirstParagraphAsIs(changes: true);
		var format = global::ParagraphFormat.Parse(settings);
		if (format.Contextual != null) throw new ArgumentException("Contextual is a style's: Rte.ModifyStyle(name, \"Contextual=1\").");
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		// (before the tabs: lines laid out without it stay wrapped as they were)
		if (format.Tabs?.Any(t => t.Alignment != "left" || t.Leader != "none") == true) (ed as WriterEditBox)?.AdvancedLayout(true);
		doc.BeginEditCollection();
		try { format.Apply(doc, selection.Start, selection.End); } finally { doc.EndEditCollection(); }
		form.LineGutter.DocumentChanged();
	}

	/// The selected paragraphs' tab stops: "72/left/none,216/right/dots" (positions in points from the left indent's edge;
	/// "" none; "mixed" if the paragraphs differ).
	public string Tabs {
		get {
			var format = global::ParagraphFormat.Read(ed.TextDocument, form.Styles, ed.TextDocument.Selection2.Start, ed.TextDocument.Selection2.End);
			return format.Tabs == null ? "mixed" : string.Join(",", format.Tabs);
		}
	}

	/// Gives the selected paragraphs these tab stops (and no others): "72/right/dots,3in/decimal" ("" clears them).
	public void SetTabs(string tabs) => SetParagraphFormat("Tabs=" + tabs);

	/// Adds a tab stop to each selected paragraph, keeping its others (one at the same place gives way): "2in",
	/// "144/right/dots" (as SetTabs takes them). One undo step. (A click on the ruler adds a left one.)
	public void AddTab(string tab) {
		var t = ParagraphTab.Parse(tab);
		using var _ = form.FirstParagraphAsIs(changes: true);
		if (t.Alignment != "left" || t.Leader != "none") (ed as WriterEditBox)?.AdvancedLayout(true);
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		doc.BeginEditCollection();
		try {
			foreach (var (start, _) in StyleEngine.Paragraphs(doc, selection.Start, selection.End))
				global::LineSpacing.AddTab(doc.Range2(start, start).Para, t.Position, Array.IndexOf(ParagraphTab.Alignments, t.Alignment), Math.Max(0, Array.IndexOf(ParagraphTab.Leaders, t.Leader)));
		} finally {
			doc.EndEditCollection();
		}
		form.LineGutter.DocumentChanged();
	}

	/// Moves the selected paragraphs' tab stop at from ("2in", 144) to to, keeping its alignment and leader ("": clears
	/// it); paragraphs without one there are left as they are. One undo step. (Dragging a tab stop on the ruler.)
	public void MoveTab(string from, string to = "") {
		float old = ParagraphTab.Parse(from).Position;
		float? at = to == "" ? null : ParagraphTab.Parse(to).Position;
		using var _ = form.FirstParagraphAsIs(changes: true);
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		doc.BeginEditCollection();
		try {
			foreach (var (start, _) in StyleEngine.Paragraphs(doc, selection.Start, selection.End)) {
				var para = doc.Range2(start, start).Para;
				for (int i = 0; i < para.TabCount; i++) {
					para.GetTab(i, out float position, out int alignment, out int leader);
					if (Math.Abs(position - old) > 0.5f) continue;
					para.DeleteTab(position);
					if (at is float x) global::LineSpacing.AddTab(para, x, alignment, leader);
					break;
				}
			}
		} finally {
			doc.EndEditCollection();
		}
		form.LineGutter.DocumentChanged();
	}

	/// Clears the selected paragraphs' tab stop at ("2in", 144), keeping their others. One undo step.
	public void ClearTab(string at) => MoveTab(at, "");

	/// The spacing of the document's default tab stops (where Tab goes past the paragraph's own), in points.
	public float DefaultTabStop {
		get => ed.TextDocument.DefaultTabStop;
		set {
			if (value < 1 || value > 1584) throw new ArgumentException("Default tab stops are 1 to 1584 points apart.");
			ed.TextDocument.DefaultTabStop = value;
			ed.Modified = true;
		}
	}

	// ---- Line numbers (iteration 19)

	/// The document's line numbering, as in Word's Layout > Line Numbers: "none", "continuous" (counting through the
	/// document) or "page" (restarting on each page). Setting it keeps the start, count-by and distance. Saved with the
	/// document.
	public string LineNumbers {
		get => LineNumbering.ModeName(form.Numbering.Mode);
		set {
			var numbering = form.Numbering.Copy();
			numbering.Mode = LineNumbering.ParseMode(value);
			form.SetLineNumbering(numbering);
		}
	}

	/// Sets the line numbering: mode as LineNumbers; start: the first line's number; countBy: show only the numbers that
	/// are multiples of it (5: 5, 10, 15 ...); distance: from the text in points, for Word and LibreOffice (0 = automatic).
	public void SetLineNumbers(string mode = "continuous", int start = 1, int countBy = 1, float distance = 0) {
		if (start < 1 || countBy < 1 || countBy > 100 || distance < 0 || distance > 1584)
			throw new ArgumentException("SetLineNumbers needs start >= 1, countBy 1 to 100 and distance 0 to 1584 points.");
		form.SetLineNumbering(new LineNumbering { Mode = LineNumbering.ParseMode(mode), Start = start, CountBy = countBy, Distance = (int)Math.Round(distance * 20) });
	}

	/// The line numbering's settings: "none", or e.g. "continuous, start 1, count by 1, distance auto".
	public string LineNumberSettings() => form.Numbering.Describe();

	/// Leaves the selected paragraphs' lines without numbers, and not counted (suppress = 1), or numbers them again (0),
	/// as Word's Suppress for Current Paragraph. One undo step.
	public void SuppressLineNumbers(bool suppress = true) {
		SelectionRange().Para.NoLineNumber = suppress ? tomConstants.tomTrue : tomConstants.tomFalse;
		form.LineGutter.DocumentChanged();
	}

	/// Whether the line numbers of the selected paragraphs are suppressed (all of them).
	public bool LineNumbersSuppressed() => SelectionRange().Para.NoLineNumber == tomConstants.tomTrue;

	/// The number of the line with the character at position (-1: the caret's line), whether or not count-by shows it;
	/// 0 if the line has none (line numbers off, or a line in a table, a page break or a suppressed paragraph).
	public int LineNumberAt(int position = -1) => form.LineGutter.NumberOf(form.LineGutter.LineOf(position < 0 ? ed.SelectionStart : position));

	/// Where the character at position (-1: the caret) is printed, as "page:line", e.g. "3:12": its page, and its line's
	/// printed number (0 if the line has none, or line numbers are off). Laid out for the printer, as Print does (the
	/// editor's own lines can wrap elsewhere).
	public string PrintedLineAt(int position = -1) {
		using var printer = form.CreatePrinter();
		var (page, line) = printer.PrintedLineAt(position < 0 ? ed.SelectionStart : position);
		return $"{page}:{line}";
	}

	// ---- Tables

	/// Inserts a table at the caret (headerRow = 1 makes the first row bold) and puts the caret in its first cell.
	/// roomAbove = 1, in an empty document: an empty paragraph stays above the table, as the Insert Table dialog leaves it:
	/// RichEdit can't space a table's first row, so a table first in a document has no top margin in Show Pages, and nothing
	/// above it for Page Up to go to.
	public void InsertTable(int rows = 3, int columns = 3, bool headerRow = true, bool roomAbove = false) {
		if (rows < 1 || columns < 1 || rows > 500 || columns > 63) throw new ArgumentException("A table has 1 to 500 rows and 1 to 63 columns.");
		InsertTableWith(rows, columns, headerRow, roomAbove, null);
	}

	// (iteration 91) InsertTable's work, with the cells' text (data: rows of cells, the columns fitted to it) or empty.
	void InsertTableWith(int rows, int columns, bool headerRow, bool roomAbove, List<List<string>>? data, Action? then = null) {
		// (gaps out first: the paragraph a gap is on would end up after the table, its space with it)
		var doc = ed.TextDocument;
		bool room = roomAbove && doc.Range2(0, 0).StoryLength <= 1;
		void Insert() => form.WithoutGapsChanging(() => {
			doc.BeginEditCollection();    // (one step to undo)
			try {
				if (room) {
					doc.Range2(0, 0).Text = "\r";
					doc.Selection2.SetRange(1, 1);
				}
				int width = WidthAt(doc.Selection2.Start);
				if (data == null) RtfTables.Insert(ed, rows, columns, headerRow, width);
				else RtfTables.Insert(ed, rows, columns, headerRow, width, (r, c) => c < data[r].Count ? data[r][c] : null, RtfTables.FittedWidths(data, width));
				then?.Invoke();    // (iteration 94: in the same step to undo)
			}
			finally { doc.EndEditCollection(); }
		});
		if (room) form.FromTheTop(Insert);
		else Insert();
	}

	/// Whether the caret is in a table.
	public bool InTable() => RtfTables.At(ed, ed.SelectionStart) != null;

	/// Inserts an empty row below (below = 1) or above (0) the caret's row.
	public void InsertRow(bool below = true) { RtfTables.InsertRow(ed, below); Redraw(); }

	/// Inserts an empty column right (right = 1) or left (0) of the caret's column; the table keeps its width.
	public void InsertColumn(bool right = true) { RtfTables.InsertColumn(ed, right); Redraw(); }

	/// Deletes the caret's row (and the table, if it was the only row).
	public void DeleteRow() { RtfTables.DeleteRow(ed); Redraw(); }

	/// Deletes the caret's column (and the table, if it was the only column).
	public void DeleteColumn() { RtfTables.DeleteColumn(ed); Redraw(); }

	/// Deletes the table the caret is in.
	public void DeleteTable() { RtfTables.DeleteTable(ed); Redraw(); }

	/// (iteration 92) Merges the selected cells into one, as Word does: their text goes into it, each cell's as paragraphs
	/// of its own. The selected cells are those from the cell where the selection starts to the one where it ends (the
	/// rows between, and the columns between the two cells), widened to whole merged cells. Cells merged across columns
	/// make one wider cell; cells merged down rows show as one cell (RichEdit and Word draw them so; the text is the top
	/// cell's).
	public void MergeCells() { RtfTables.MergeCells(ed); Redraw(); }

	/// (iteration 92) Splits the caret's cell (or the selected cells, merged first) into columns and rows, as Word does:
	/// columns of equal width; a cell merged down into a number of rows that divides the rows it spans (all of them
	/// undoes the merge); a cell of one row into rows adds rows, the row's other cells merged down over them. Its text
	/// stays in the first cell.
	public void SplitCells(int columns = 2, int rows = 1) { RtfTables.SplitCells(ed, columns, rows); Redraw(); }

	/// (iteration 92) Deletes the selected cells (the caret's), as Word's Delete Cells does: how = "left" (the cells
	/// after them in their rows move left), "up" (the cells below them move up: their text; empty cells at the bottom),
	/// "row" (their rows) or "column" (their columns).
	public void DeleteCells(string how = "left") { RtfTables.DeleteCells(ed, how); Redraw(); }

	/// (iteration 92) Selects the cells from row1, cell1 to row2, cell2 (counted from 1) of the table at the caret (the
	/// document's first table if the caret isn't in one) as a block, as dragging over them with the mouse does: the rows
	/// between, and the columns between the two cells.
	/// MergeCells, SplitCells, DeleteCells and ClearCells then act on them (RichEdit itself can only select whole rows
	/// across rows).
	public void SelectCells(int row1, int cell1, int row2, int cell2) => RtfTables.SelectCells(ed, row1, cell1, row2, cell2);

	/// (iteration 94) Selects part of the table at the caret, as Word's Table Layout > Select does: part = "table",
	/// "row", "column" (the cells under the caret's, by position) or "cell". Borders, shading and the other formats then
	/// act on all of them (a block of cells, as SelectCells makes).
	public void SelectTable(string part = "table") => RtfTables.SelectTablePart(ed, part);

	/// (iteration 92) Empties the selected cells' text (as Delete does on a block of cells), keeping their formats.
	public void ClearCells() { RtfTables.ClearCells(ed); Redraw(); }

	// (iteration 92) The caret's cell's span (columns, rows): Split Cells' dialog starts with it.
	internal (int Columns, int Rows) CellSpan() => RtfTables.CellSpan(ed);

	// (iteration 92) Whether the selection reaches from one table cell into another (the right-click menu's Merge cells).
	internal bool SelectionSpansCells() {
		if (RtfTables.HasBlock(ed)) return true;
		var selection = ed.TextDocument.Selection2;
		if (selection.End <= selection.Start) return false;
		var a = RtfTables.At(ed, selection.Start);
		var b = RtfTables.At(ed, Math.Max(selection.Start, selection.End - 1));
		return a is { } x && b is { } y && x.Table.Start == y.Table.Start && (x.Row != y.Row || x.Cell != y.Cell);
	}

	/// The number of table rows the selection covers completely (0 if it is not whole rows of one table).
	public int SelectedRows() => RtfTables.SelectedRows(ed).Rows.Count;

	/// Deletes the table rows the selection covers completely (the whole table if that is all of them), as the Delete
	/// key does; returns how many were deleted.
	public int DeleteSelectedRows() { int count = RtfTables.DeleteSelectedRows(ed); Redraw(); return count; }

	// RichEdit doesn't repaint where a removed table row was (its borders stayed on screen): the whole editor is redrawn.
	void Redraw() => ed.Invalidate();

	// ---- Iteration 91: tables from data, and data from tables (rte_tabledata.cs)

	/// Inserts a table at the caret made from data: CSV or TSV text, or JSON (an array of rows, each an array of values,
	/// or an array of objects, whose names make the first row); format "csv", "tsv", "json" or "" (told from the text:
	/// JSON if it starts with [ or {; CSV's delimiter is a tab, ";" or ",", whichever its first line has most of).
	/// headerRow = 1 makes the first row bold. The columns are as wide as their text needs, filling the line. Returns
	/// the number of rows (at most 10000; at most 63 columns).
	public int InsertTableFromData(string data, string format = "", bool headerRow = true) =>
		InsertDataTable(TabularData.Parse(data, format), headerRow);

	/// Inserts a table at the caret made from a file: an Excel workbook (.xlsx, .xlsm; no Excel needed, and it may be
	/// open in Excel), CSV or TSV (.csv, .tsv, .txt), or JSON (.json), as InsertTableFromData. In a workbook: sheet is a
	/// sheet's name or number ("" for the first), range one such as "A1:D20" ("" for the cells in use); values come as
	/// Excel shows them (number formats, dates, percentages; formulas' last results). Returns the number of rows.
	/// (iteration 94) link = 1 (the default): the table is linked to the file (LinkTable), for UpdateTable and
	/// SaveTableToSource.
	public int InsertTableFromFile(string path, bool headerRow = true, string sheet = "", string range = "", bool link = true) {
		string full = FullPath(path);
		var source = link ? new TableSources.Source(TableSources.Relative(full, DocumentFolder), (sheet ?? "").Trim(), (range ?? "").Trim()) : null;
		return InsertDataTable(TabularData.ReadFile(full, sheet, range), headerRow,
			source == null ? null : () => { if (RtfTables.At(ed, ed.TextDocument.Selection2.Start) is { } at) PutSource(at.Table, source); });
	}

	int InsertDataTable(List<List<string>> rows, bool headerRow, Action? then = null) {
		if (rows.Count == 0 || rows[0].Count == 0) throw new ArgumentException("The data has no rows.");
		InsertTableWith(rows.Count, rows[0].Count, headerRow, true, rows, then);
		return rows.Count;
	}

	/// The names of an Excel workbook's sheets, one per line.
	public string WorkbookSheets(string path) => string.Join("\n", TabularData.SheetNames(FullPath(path)));

	static string FullPath(string path) {
		if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Give a file's path.");
		return Path.GetFullPath(Environment.ExpandEnvironmentVariables(path.Trim().Trim('"')));
	}

	// A table: 0 the caret's, n the document's nth (from 1).
	RtfTables.Table TableNumber(int table) {
		if (table == 0) return CurrentTable();
		var all = RtfTables.Find(ed.TextDocument);
		if (table < 1 || table > all.Count) throw new ArgumentException($"There is no table {table}: the document has {all.Count} table{(all.Count == 1 ? "" : "s")}.");
		return all[table - 1];
	}

	// A cell's text: its paragraphs and line breaks as \n.
	string CellOf(RtfTables.Cell cell) {
		string text = TableSources.Marker.Replace(ed.TextDocument.Range2(cell.Start, cell.End).Text ?? "", "");    // (iteration 94: not a link's marker)
		return text.Replace(RtfTables.MergedMark.ToString(), "").Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n');    // (iteration 93: not a merged cell's mark)
	}

	// The table's cells' text, row by row.
	List<List<string>> CellsOf(RtfTables.Table table) => table.Rows.Select(r => r.Cells.Select(CellOf).ToList()).ToList();

	/// A table's cells as text: format "csv", "csv;" (semicolons, as Excel saves CSV where the decimal mark is a
	/// comma), "tsv", "json" (an array of rows) or "objects" (an array of objects named by the first row's cells).
	/// table: 0 for the caret's, or a number from 1 (the document's first, second ... table).
	public string TableData(int table = 0, string format = "csv") => TabularData.Write(CellsOf(TableNumber(table)), format);

	/// (iteration 93) Writes a table's cells back to a file, the way InsertTableFromFile and SetTableDataFromFile read
	/// it: an Excel workbook (.xlsx, .xlsm; it must exist and be closed in Excel) gets the cells into sheet (a name or
	/// number, "" for the first) from range's top left ("" for where its cells in use start), only the cells whose value
	/// changed, the cells left over emptied, formulas recalculated when Excel opens it; .json an array of objects if the
	/// file holds one (named by the first row), else an array of rows; other files CSV or TSV (format "csv", "csv;",
	/// "tsv", or "" for the file's own delimiter), UTF-8. table: 0 for the caret's, or a number from 1. Returns the number
	/// of cells changed (a workbook) or of rows written.
	public int SaveTableData(string path, int table = 0, string sheet = "", string range = "", string format = "") =>
		TabularData.WriteFile(FullPath(path), CellsOf(TableNumber(table)), sheet, range, format);

	// ---- Iteration 94: tables linked to their files (TableSources, rte_tabledata.cs)

	string? DocumentFolder => form.FileName is { Length: > 0 } file ? Path.GetDirectoryName(Path.GetFullPath(file)) : null;

	// A table's link: its marker's place in the first cell, and what it says (null: none).
	(int Start, int End, TableSources.Source? Source) SourceOf(RtfTables.Table table) {
		if (table.Rows.Count == 0 || table.Rows[0].Cells.Count == 0) return (-1, -1, null);
		var cell = table.Rows[0].Cells[0];
		var m = TableSources.Marker.Match(ed.TextDocument.Range2(cell.Start, cell.End).Text ?? "");
		return m.Success ? (cell.Start + m.Index, cell.Start + m.Index + m.Length, TableSources.Source.Decode(m.Value)) : (-1, -1, null);
	}

	// The table's marker replaced (source null: taken out), inside the caller's step to undo.
	void PutSource(RtfTables.Table table, TableSources.Source? source) {
		var (start, end, _) = SourceOf(table);
		var doc = ed.TextDocument;
		if (start >= 0) doc.Range2(start, end).Text = "";
		if (source != null && table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0) CrossRefs.Hidden(doc, table.Rows[0].Cells[0].Start, source.Encode());
	}

	void SetSource(RtfTables.Table table, TableSources.Source? source) {
		var doc = ed.TextDocument;
		int caret = doc.Selection2.Start, end = doc.Selection2.End;
		form.WithoutGapsChanging(() => {
			doc.BeginEditCollection();
			try { PutSource(table, source); }
			finally { doc.EndEditCollection(); }
		});
		doc.Selection2.SetRange(caret, end);
	}

	/// Links a table (0: the caret's, or a number from 1) to a file its data come from, as InsertTableFromFile does: an
	/// Excel workbook (sheet: a name or number, "" for the first; range such as "A1:D20", "" for the cells in use), CSV,
	/// TSV or JSON. The link is a hidden marker in the table's first cell: it goes wherever the table goes and is saved
	/// with the document (RTF, Word, HTML, Markdown). The path is kept relative to the document's folder when the file
	/// is in it or below it.
	public void LinkTable(string path, int table = 0, string sheet = "", string range = "") =>
		SetSource(TableNumber(table), new TableSources.Source(TableSources.Relative(FullPath(path), DocumentFolder), (sheet ?? "").Trim(), (range ?? "").Trim()));

	/// Takes away a table's link to its file (the table stays as it is).
	public void UnlinkTable(int table = 0) => SetSource(TableNumber(table), null);

	/// A table's link as "path|sheet|range" (the path as kept: maybe relative to the document's folder), or "" if it
	/// has none.
	public string TableSource(int table = 0) => SourceOf(TableNumber(table)).Source?.ToString() ?? "";

	// A linked table's file, read.
	List<List<string>> ReadSource(TableSources.Source source) => TabularData.ReadFile(TableSources.Resolve(source.Path, DocumentFolder), source.Sheet, source.Range);

	TableSources.Source LinkOf(RtfTables.Table table) =>
		SourceOf(table).Source ?? throw new InvalidOperationException(T("status.table.notLinked", "This table isn't linked to a file: insert it from a file, or link it (Table Layout > Link to File)."));

	/// Fills a linked table again from its file (all its rows, the file's first row too), keeping its formats: rows
	/// are added or deleted to fit. Returns the number of rows.
	public int UpdateTable(int table = 0) {
		var target = TableNumber(table);
		return FillTable(target, ReadSource(LinkOf(target)), 1, null);
	}

	/// Fills every linked table of the document again from its file; returns how many. A file that can't be read
	/// leaves its table as it was; the errors come at the end, one per line.
	public int UpdateTables() {
		var doc = ed.TextDocument;
		var linked = RtfTables.Find(doc).Select(t => (t.Start, SourceOf(t).Source)).Where(x => x.Source != null).ToList();
		int done = 0;
		var errors = new List<string>();
		foreach (var (start, source) in Enumerable.Reverse(linked)) {    // (from the end: the earlier tables stay where they are)
			try {
				var table = RtfTables.Find(doc).First(t => t.Start == start);
				FillTable(table, ReadSource(source!), 1, null);
				done++;
			}
			catch (Exception ex) when (ex is IOException or ArgumentException or InvalidOperationException or UnauthorizedAccessException) { errors.Add(Path.GetFileName(source!.Path) + ": " + ex.Message); }
		}
		if (errors.Count > 0) throw new InvalidOperationException(T("status.table.updateFailed", "{0} of {1} linked tables updated; not:", done, linked.Count) + "\n" + string.Join("\n", errors));
		return done;
	}

	/// Writes a linked table back to its file, as SaveTableData does (a workbook: only the cells that changed).
	/// Returns the number of cells changed (a workbook) or of rows written. (iteration 94) A workbook's table that has
	/// grown past the link's range, or would write over cells of the sheet that aren't empty beyond the area it was
	/// read from (rows below the data), is not written unless force = 1 (SaveTableCheck says what it would do); written
	/// past the range, the link's range grows to take the table in, so that UpdateTable reads all of it.
	public int SaveTableToSource(int table = 0, bool force = false) {
		var target = TableNumber(table);
		var source = LinkOf(target);
		string path = TableSources.Resolve(source.Path, DocumentFolder);
		var cells = CellsOf(target);
		var reach = TabularData.WriteReach(path, cells, source.Sheet, source.Range);
		if (!force && ReachWarning(source, reach) is { Length: > 0 } warning)
			throw new InvalidOperationException(warning + " " + T("status.table.notWritten", "The table was not written (SaveTableToSource with force = 1 writes it anyway)."));
		int count = TabularData.WriteFile(path, cells, source.Sheet, source.Range);
		if (reach.Past) SetSource(target, source with { Range = reach.Area });
		return count;
	}

	/// (iteration 94) What SaveTableToSource would do that it doesn't do without force = 1, as a warning: "" if
	/// nothing (the table fits in the link's range, and writes over nothing beyond the data it was read from).
	public string SaveTableCheck(int table = 0) {
		var target = TableNumber(table);
		var source = LinkOf(target);
		return ReachWarning(source, TabularData.WriteReach(TableSources.Resolve(source.Path, DocumentFolder), CellsOf(target), source.Sheet, source.Range));
	}

	string ReachWarning(TableSources.Source source, (string Area, bool Past, List<int> Filled) reach) {
		var parts = new List<string>();
		if (reach.Past)
			parts.Add(T("msg.saveTable.past", "The table goes past the linked range {0}: it would take {1} (the link's range then becomes {1}).", source.Range.ToUpperInvariant(), reach.Area));
		if (reach.Filled.Count > 0) {
			string rows = string.Join(", ", reach.Filled.Take(8)) + (reach.Filled.Count > 8 ? ", …" : "");
			parts.Add(reach.Filled.Count == 1
				? T("msg.saveTable.filledOne", "Row {0} of the sheet is not empty where the table would write: its cells would be written over.", rows)
				: T("msg.saveTable.filled", "Rows {0} of the sheet are not empty where the table would write: their cells would be written over.", rows));
		}
		return string.Join(" ", parts);
	}

	/// The text of a cell (row and column from 1) of the caret's table (table = 0) or of the document's nth table.
	public string CellText(int row, int column, int table = 0) => CellOf(CellAt(TableNumber(table), row, column));

	RtfTables.Cell CellAt(RtfTables.Table table, int row, int column) {
		if (row < 1 || row > table.Rows.Count) throw new ArgumentException($"There is no row {row}: the table has {table.Rows.Count}.");
		var cells = table.Rows[row - 1].Cells;
		if (column < 1 || column > cells.Count) throw new ArgumentException($"There is no column {column} in row {row}: it has {cells.Count} cell{(cells.Count == 1 ? "" : "s")}.");
		return cells[column - 1];
	}

	/// Replaces a cell's text (row and column from 1; `n for a line break), keeping its format.
	public void SetCellText(int row, int column, string text, int table = 0) {
		var target = TableNumber(table);
		var cell = CellAt(target, row, column);
		var doc = ed.TextDocument;
		var source = row == 1 && column == 1 ? SourceOf(target).Source : null;    // (iteration 94: a linked table's marker stays)
		int start = target.Start;
		form.WithoutGapsChanging(() => {
			doc.BeginEditCollection();
			try {
				doc.Range2(cell.Start, cell.End).Text = CellValue(text);
				if (source != null) PutSource(RtfTables.Find(doc).First(t => t.Start == start), source);
			}
			finally { doc.EndEditCollection(); }
		});
		Redraw();
	}

	static string CellValue(string? text) => (text ?? "").Replace("\r\n", "\v").Replace('\n', '\v').Replace('\r', '\v');

	/// Fills a table with data (as InsertTableFromData's), keeping its formats: from row fromRow on (1: all of it; 2: the
	/// first row, a header, stays), the table gets the data's rows, its last row's format copied for rows added and the
	/// rows left over deleted. With fromRow 2 or more: a first data row the same as the table's header row is left out,
	/// and JSON objects go in the columns their names head (in their own order if a name heads none). A row with more
	/// values than cells is an error. table: 0 for the caret's, or a number from 1. Returns the number of rows filled.
	public int SetTableData(string data, string format = "", int table = 0, int fromRow = 1) {
		var target = TableNumber(table);
		var names = fromRow > 1 ? TabularData.JsonNames(data, format) : null;
		var rows = TabularData.Parse(data, format, keysRow: names == null);
		return FillTable(target, rows, fromRow, names);
	}

	/// SetTableData from a file, read as InsertTableFromFile reads it (an Excel workbook's sheet and range, CSV, TSV or
	/// JSON).
	public int SetTableDataFromFile(string path, int table = 0, int fromRow = 1, string sheet = "", string range = "") {
		var target = TableNumber(table);
		string full = FullPath(path);
		var names = fromRow > 1 && Path.GetExtension(full).Equals(".json", StringComparison.OrdinalIgnoreCase) ? TabularData.JsonNames(TabularData.ReadText(full), "json") : null;
		var rows = names != null ? TabularData.Parse(TabularData.ReadText(full), "json", keysRow: false) : TabularData.ReadFile(full, sheet, range);
		return FillTable(target, rows, fromRow, names);
	}

	int FillTable(RtfTables.Table table, List<List<string>> rows, int fromRow, List<string>? names) {
		if (fromRow < 1 || fromRow > table.Rows.Count) throw new ArgumentException($"fromRow is 1 to {table.Rows.Count} (the table's rows).");
		int keep = fromRow - 1;
		if (keep > 0) {
			var header = table.Rows[0].Cells.Select(c => CellOf(c).Trim()).ToList();
			// (JSON objects: their values in the columns their names head)
			if (names != null && names.All(n => header.Contains(n.Trim(), StringComparer.OrdinalIgnoreCase)))
				rows = rows.Select(r => header.Select(h => names.FindIndex(n => n.Trim().Equals(h, StringComparison.OrdinalIgnoreCase)) is int i && i >= 0 && i < r.Count ? r[i] : "").ToList()).ToList();
			// (a first row that repeats the header: the data's own header)
			else if (rows.Count > 0 && rows[0].Count >= header.Count && rows[0].Take(header.Count).Select(t => t.Trim()).SequenceEqual(header, StringComparer.OrdinalIgnoreCase)
				&& rows[0].Skip(header.Count).All(t => t.Trim() == ""))
				rows = rows.Skip(1).ToList();
		}
		if (keep == 0 && rows.Count == 0) throw new ArgumentException("The data has no rows (a table keeps at least one).");
		if (keep + rows.Count > TabularData.MaxRows) throw new ArgumentException($"A table can have at most {TabularData.MaxRows} rows.");
		for (int r = 0; r < rows.Count; r++) {    // (checked before anything changes)
			int last = rows[r].FindLastIndex(v => v != "");
			int cells = table.Rows[Math.Min(keep + r, table.Rows.Count - 1)].Cells.Count;
			if (last >= cells) throw new ArgumentException($"Data row {r + 1} has {last + 1} values; table row {keep + r + 1} has {cells} cells.");
		}
		var doc = ed.TextDocument;
		int start = table.Start, caret = doc.Selection2.Start;
		var source = SourceOf(table).Source;    // (iteration 94: a linked table's marker, in its first cell, stays)
		form.WithoutGapsChanging(() => {
			doc.BeginEditCollection();    // (one step to undo)
			doc.Freeze();
			try {
				RtfTables.ResizeRows(ed, start, keep + rows.Count);
				var now = RtfTables.Find(doc).First(t => t.Start == start);
				for (int r = now.Rows.Count - 1; r >= keep; r--) {
					var values = rows[r - keep];
					var cells = now.Rows[r].Cells;
					for (int c = cells.Count - 1; c >= 0; c--) {    // (from the end: the earlier cells stay where they are)
						string value = CellValue(c < values.Count ? values[c] : "");
						var range = doc.Range2(cells[c].Start, cells[c].End);
						if ((range.Text ?? "") != value) range.Text = value;
					}
				}
				if (source != null) PutSource(RtfTables.Find(doc).First(t => t.Start == start), source);
			}
			finally {
				doc.Unfreeze();
				doc.EndEditCollection();
			}
		});
		var filled = RtfTables.Find(doc).FirstOrDefault(t => t.Start == start);
		if (filled != null) ed.Select(caret < filled.Start ? caret : filled.Rows[Math.Min(keep, filled.Rows.Count - 1)].Cells[0].Start, 0);
		Redraw();
		return rows.Count;
	}

	// ---- Page breaks

	// ---- Pictures' alt text (iteration 15)

	List<string>? altCache;    // the pictures' alt texts in document order, from RichEditorNET's HTML; null: to be read again

	/// The alt text of the picture at a position ("" if it has none).
	public string PictureAltText(int position) {
		var doc = ed.TextDocument;
		if (doc.Range2(position, position + 1).Text != "\uFFFC") throw new ArgumentException("There is no picture at " + position + ".");
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		int index = story[..position].Count(c => c == '\uFFFC');
		altCache ??= Regex.Matches(ed.Html, "<img [^>]*>", RegexOptions.IgnoreCase)
			.Select(m => WebUtility.HtmlDecode(Regex.Match(m.Value, "\\balt=\"([^\"]*)\"").Groups[1].Value)).ToList();
		int pictures = story.Count(c => c == '\uFFFC');
		return altCache.Count == pictures && index < altCache.Count ? altCache[index] : "";
	}

	/// Gives the picture at a position a new alt text ("" for none). RichEditorNET can't change a picture's alt text, so
	/// the picture is replaced by the same picture (image, size and link) with the new one, in one undo step.
	public void SetPictureAltText(int position, string alt) {
		var doc = ed.TextDocument;
		if (doc.Range2(position, position + 1).Text != "\uFFFC") throw new ArgumentException("There is no picture at " + position + ".");
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		selection.SetRange(position, position + 1);
		var pictures = OfficeConverters.Pictures(((RichTextBox)ed).SelectedRtf ?? "");
		if (pictures.Count != 1 || pictures[0] is not { } picture) {
			selection.SetRange(start, end);
			throw new InvalidOperationException("This picture can't be read, so its alt text can't be changed.");
		}
		doc.BeginEditCollection();
		try {
			string url = Regex.Match(doc.Range2(position, position + 1).URL ?? "", @"""?([^""]+)""?").Groups[1].Value;
			int before = doc.Range2(0, 0).StoryLength;
			selection.Text = "";
			// a picture that was a link's only text takes the link's hidden address with it: back to where the link began
			int at = position - Math.Max(0, before - doc.Range2(0, 0).StoryLength - 1);
			selection.SetRange(at, at);
			int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
			try {
				ed.DefaultImageWidth = picture.Width;
				ed.DefaultImageHeight = picture.Height;
				using var stream = new MemoryStream(picture.Image);
				ed.InsertImage(stream, alt);
			} finally {
				ed.DefaultImageWidth = oldWidth;
				ed.DefaultImageHeight = oldHeight;
			}
			if (url != "") doc.Range2(at, at + 1).URL = "\"" + url.Replace("\"", "%22") + "\"";
		} finally {
			doc.EndEditCollection();
		}
		altCache = null;
		selection.SetRange(start, end);
	}

	/// Iteration 63: inserts a section break at the caret (replacing the selection): a page break after which a new
	/// section starts, with its own header and footer (at first the section before's: "same as previous") and its page
	/// numbers starting again at 1 (Layout > Header & Footer changes them, where the caret is).
	public void InsertSectionBreak() {
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try { Sections.Insert(ed); }
		finally { doc.EndEditCollection(); }
		form.BandsChangedNow();
		form.SetStatus(T("status.sectionBreak", "Section break inserted: the next page starts a new section, its page numbers from 1. Layout > Header & Footer (with the caret in it) gives it its own footer."));
	}

	/// The sections: where each starts (characters), whether its page numbers start again and at what, whether its
	/// header and footer are the section before's ("same as previous"), and any top and bottom margins of its own (1/100
	/// inch; "-" for one that is the document's). The first section is the document's own.
	public string SectionList() {
		var lines = new List<string> { "0 restart 1 own" };
		var breaks = Sections.Breaks(ed.TextDocument);
		var sections = Sections.Resolve(breaks, form.Bands);
		foreach (var (b, section) in breaks.Zip(sections.Skip(1)))
			lines.Add($"{b.SectionStart} {(b.Settings.Restart ? "restart " + b.Settings.Start : "continue")} {(b.Settings.Same ? "same" : "own")}{(section.Top == null && section.Bottom == null ? "" : $" {section.Top?.ToString() ?? "-"} {section.Bottom?.ToString() ?? "-"}")}");
		return string.Join("\n", lines);
	}

	/// Gives the section a position is in (not the first: the document's own, Page Setup's) top and bottom margins of its
	/// own (1/100 inch; -1: the document's).
	public void SetSectionMargins(int position, int top, int bottom) {
		var doc = ed.TextDocument;
		if (Sections.BreakFor(doc, position) is not { } section) throw new ArgumentException("The first section has the document's margins (Page Setup).");
		static int? Twips(int hundredths) => hundredths < 0 ? null : (int)Math.Round(hundredths * 14.4);
		var settings = section.Settings.Copy();
		settings.Page = SectionPage.WithMargins(settings.Page, Twips(top), Twips(bottom));
		doc.BeginEditCollection();
		try { Sections.Write(doc, section, settings); }
		finally { doc.EndEditCollection(); }
		ed.Modified = true;
		form.BandsChangedNow();
	}

	/// Inserts a continuous section break at the caret: the section after it starts on the same page, with the settings of
	/// the section it splits, so that the two can have different columns.
	public void InsertContinuousBreak() {
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try { Sections.InsertContinuous(ed, doc.Selection2.End, firstColumns: form.DocumentColumns); }
		finally { doc.EndEditCollection(); }
		ed.Modified = true;
		form.BandsChangedNow();
		form.SetStatus(T("status.continuousBreak", "Continuous section break inserted: the next section starts on this page. Layout > Columns, with the caret in it, gives it columns."));
	}

	/// Puts the paragraphs from start to end in columns, as Word does: continuous breaks around them where there are none,
	/// the text around them keeping its columns. The arguments as SetColumns'. Returns where their section starts.
	public int SetColumnsOn(int start, int end, int count, int gap = 50, bool line = false, string widths = "") {
		var doc = ed.TextDocument;
		var first = doc.Range2(start, start);
		first.Expand(tomConstants.tomParagraph);
		var last = doc.Range2(Math.Max(start, end - 1), Math.Max(start, end - 1));
		last.Expand(tomConstants.tomParagraph);
		int from = first.Start, to = last.End;
		bool Starts(int at) => at == 0 || Sections.Breaks(doc).Any(b => b.SectionStart == at);
		doc.BeginEditCollection();
		try {
			if (to < doc.Range2(0, 0).StoryLength && !Starts(to)) Sections.InsertContinuous(ed, to, firstColumns: form.DocumentColumns);
			if (!Starts(from)) from = Sections.InsertContinuous(ed, from, firstColumns: form.DocumentColumns);
			int index = Sections.Breaks(doc).FindIndex(b => b.SectionStart == from);
			WriteColumns(from, count, gap, line, widths);
			if (index >= 0) from = Sections.Breaks(doc)[index].SectionStart;    // (its marker grew)
		}
		finally { doc.EndEditCollection(); }
		doc.Selection2.SetRange(from, from);
		ed.Modified = true;
		form.BandsChangedNow();
		return from;
	}

	/// Sets the columns of the section a position is in, on the printed pages (the text is edited in one column): 1 to 6
	/// of equal width (1: none), the gap between them (1/100 inch), a line in the gap; widths: each column's own (1/100
	/// inch, with commas between them; on the page in proportion to the text's width), "": equal.
	public void SetColumns(int position, int count, int gap = 50, bool line = false, string widths = "") {
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try { WriteColumns(position, count, gap, line, widths); }
		finally { doc.EndEditCollection(); }
		ed.Modified = true;
		form.BandsChangedNow();
	}

	void WriteColumns(int position, int count, int gap, bool line, string widths) {
		if (count < 1 || count > 6) throw new ArgumentException("A section has 1 to 6 columns.");
		if (gap < 0) throw new ArgumentException("The gap can't be negative.");
		var doc = ed.TextDocument;
		int[]? own = null;
		if (!string.IsNullOrWhiteSpace(widths)) {
			var parts = widths.Split(',', StringSplitOptions.TrimEntries);
			if (parts.Length != count || parts.Any(p => !int.TryParse(p, out int w) || w <= 0)) throw new ArgumentException("Give each column's width, 1/100 inch, with commas between them.");
			own = parts.Select(p => (int)Math.Round(int.Parse(p) * 14.4)).ToArray();
		}
		var columns = SectionColumns.Of(count, (int)Math.Round(gap * 14.4), line, own);
		int index = Sections.IndexAt(Sections.Resolve(doc, form.Bands, form.DocumentColumns), position);
		if (Sections.BreakFor(doc, position) is { } section) {
			var settings = section.Settings.Copy();
			settings.Columns = columns;
			Sections.Write(doc, section, settings);
		}
		else form.DocumentColumns = columns;
		if (columns != null) FitTables(index, columns.NarrowestWidth(TableWidth()));
	}

	// The tables of section index, those wider than width made as wide, their columns in proportion: a table doesn't wrap as
	// text does, and would print across the next column. (As wide as the narrowest column: it can land in any of them.)
	void FitTables(int index, int width) {
		var doc = ed.TextDocument;
		var sections = Sections.Resolve(doc, form.Bands, form.DocumentColumns);
		int start = sections[index].Start, end = index + 1 < sections.Count ? sections[index + 1].Start : int.MaxValue;
		foreach (var table in RtfTables.Find(doc).Where(t => t.Start >= start && t.Start < end).Reverse().ToList())
			RtfTables.Format(ed, table.Start, rows => {
				foreach (var row in rows) {
					int total = row.Cells.Sum(c => c.Width);
					if (total > width)
						foreach (var cell in row.Cells) cell.Width = Math.Max(144, cell.Width * width / total);
				}
			});
	}

	// How wide a new table at a position is: the page's text, or the narrowest of its section's columns.
	int WidthAt(int position) {
		var sections = Sections.Resolve(ed.TextDocument, form.Bands, form.DocumentColumns);
		int width = TableWidth();
		return sections[Sections.IndexAt(sections, position)].Columns is { } columns && columns.FitIn(width) ? Math.Min(width, columns.NarrowestWidth(width)) : width;
	}

	/// The columns of the section a position is in: how many, the gap (1/100 inch), 1 if a line is drawn in it ("1 0 0": one),
	/// then, if they aren't of equal width, each one's (1/100 inch, with commas between them).
	public string ColumnsAt(int position) {
		var sections = Sections.Resolve(ed.TextDocument, form.Bands, form.DocumentColumns);
		return sections[Sections.IndexAt(sections, position)].Columns is { } c ? $"{c.Count} {(int)Math.Round(c.Gap / 14.4)} {(c.Line ? 1 : 0)}" + (c.Widths is { } own ? " " + string.Join(",", own.Select(w => (int)Math.Round(w / 14.4))) : "") : "1 0 0";
	}

	/// What the status line adds on opening a document whose sections have a page setup that is kept but not shown, or
	/// columns, which show only on the printed pages ("": none).
	internal string OpenNotice() {
		var breaks = Sections.Breaks(ed.TextDocument);
		string notice = Sections.KeptOnly(breaks) is { Count: > 0 } kept
			? $". Section{(kept.Count == 1 ? "" : "s")} {string.Join(", ", kept)}: paper, orientation or left and right margins of {(kept.Count == 1 ? "its" : "their")} own are printed, and Print Preview shows them; here {(kept.Count == 1 ? "it is" : "they are")} edited at the first section's width"
			: "";
		if (form.DocumentColumns != null || breaks.Any(b => b.Settings.Columns != null)) notice += ". It has columns: they are printed, and Print Preview shows them; here it is edited in one column";
		var openTime = form.OpenTime;
		if (openTime.TotalSeconds >= 2) notice += $" ({openTime.TotalSeconds:0.0} seconds)";
		return notice;
	}

	/// Inserts a page break at the caret (replacing the selection): the text after it starts a new page when printed,
	/// and in Word and LibreOffice. A numbered list it splits goes on with its numbers after it. Ctrl+Enter.
	public void InsertPageBreak() {
		var doc = ed.TextDocument;
		string hint = "";
		doc.BeginEditCollection();
		try {
			PageBreaks.Insert(ed);
			// the list items after the break go on from those before it (legal numbers always do)
			var para = ParaAt(doc.Selection2.Start);
			int type = para.ListType & 0xFFFF;
			if (type != tomConstants.tomListNone && type != tomConstants.tomListBullet && !LegalLists.IsLegal(para.ListType)) {
				if (para.ListLevelIndex == 0) {
					bool continued = true;
					try { ContinueNumbering(); } catch (InvalidOperationException) { continued = false; }
					if (continued) hint = T("status.pageBreakListGoesOn", "Page break inserted. The numbering after it goes on from the list before it, but stays as it is: if you add or remove items "
						+ "before the break, put the caret in the first item after it and use Home > List options > Continue numbering.");
				} else {
					hint = T("status.pageBreakSubList", "Page break inserted. The sub-list after it is numbered from the start again (a break before or after the list keeps its numbering).");
				}
			}
			if (LegalLists.AnyLegal(doc)) LegalLists.Renumber(ed);
		} finally {
			doc.EndEditCollection();
		}
		Redraw();
		if (hint != "") form.SetStatus(hint);
	}

	/// The number of page breaks in the document.
	public int PageBreakCount() => PageBreaks.Positions(ed.TextDocument).Count;

	/// The caret's table position as "row,column" (from 1), or "" outside tables.
	public string TableCell() => RtfTables.At(ed, ed.SelectionStart) is { } at ? $"{at.Row + 1},{at.Cell + 1}" : "";

	/// The number of tables in the document.
	public int TableCount => RtfTables.Find(ed.TextDocument).Count;

	/// (iteration 93) The caret's table's number in the document (from 1, as TableData's table counts them), or 0 outside
	/// tables.
	public int TableNumberAtCaret() => RtfTables.At(ed, ed.SelectionStart) is { } at ? RtfTables.Find(ed.TextDocument).FindIndex(t => t.Start == at.Table.Start) + 1 : 0;

	// ---- Iteration 18: table properties (rte_tables2.cs)

	RtfTables.Table CurrentTable() => RtfTables.At(ed, ed.SelectionStart)?.Table ?? throw new InvalidOperationException("Put the caret in a table first.");

	// Changes the formats of the caret's table's rows, in one undo step.
	void FormatTable(Action<RtfTables.Table, List<RtfTables.RowFormat>> change) {
		var table = CurrentTable();
		var doc = ed.TextDocument;
		// (iteration 94) the selected cells taken before the table's RTF is read (reading it moves the selection, which
		// ends a block of cells), and the block put back after, for the next format
		var writer = ed as WriterEditBox;
		var block = writer?.CellBlock;
		formatCells = RtfTables.SelectedCells(ed, table);
		doc.BeginEditCollection();
		try {
			RtfTables.Format(ed, table.Start, formats => change(table, formats));
		} finally {
			doc.EndEditCollection();
			formatCells = null;
		}
		if (writer != null && block is { } b && RtfTables.At(ed, b.From) is { } x && RtfTables.At(ed, b.To) is { } y && x.Table.Start == y.Table.Start)
			writer.SelectBlock(b.From, b.To);
		Redraw();
	}

	// The cells a format is for while FormatTable runs (null: the selection's, as it is now).
	List<(int Row, int Cell)>? formatCells;

	/// The formats of the caret's table's rows, and the cells (row, column) the selection is in.
	internal (List<RtfTables.RowFormat> Rows, List<(int Row, int Cell)> Selected) TableFormats() {
		var table = CurrentTable();
		var selected = RtfTables.SelectedCells(ed, table);    // (iteration 94: before the RTF is read, which ends a block)
		var writer = ed as WriterEditBox;
		var block = writer?.CellBlock;
		var formats = RtfTables.Formats(ed, table);
		if (writer != null && block is { } b && writer.CellBlock == null) writer.SelectBlock(b.From, b.To);
		return (formats, selected);
	}

	/// The width the text has on the page, in twips (new tables are as wide).
	internal int PageTextWidth() => TableWidth();

	// The cells a change is for: those the selection is in, or all of the table.
	List<(int Row, int Cell)> CellsFor(RtfTables.Table table, List<RtfTables.RowFormat> rows, bool wholeTable) =>
		wholeTable ? rows.SelectMany((row, r) => row.Cells.Select((cell, c) => (r, c))).ToList() : formatCells ?? RtfTables.SelectedCells(ed, table);

	/// Puts the caret's table at the "left", "center" or "right" of the page.
	public void SetTableAlignment(string alignment) {
		int value = alignment.Trim().ToLowerInvariant() switch { "left" => 0, "center" or "centre" => 1, "right" => 2, _ => throw new ArgumentException("The alignment is \"left\", \"center\" or \"right\".") };
		FormatTable((table, rows) => rows.ForEach(row => row.Alignment = value));
	}

	/// Sets the space left and right of the text in the cells of the caret's table, in points.
	public void SetCellPadding(float points) {
		if (points is < 0 or > 72) throw new ArgumentException("Cell padding is 0 to 72 points.");
		FormatTable((table, rows) => rows.ForEach(row => row.Padding = (int)Math.Round(points * 20)));
	}

	/// Sets the height of the rows the selection is in (all rows with wholeTable = 1), in points: 0 is automatic (as high
	/// as their text), otherwise at least that high, or exactly that high (exact = 1).
	public void SetRowHeight(float points, bool exact = false, bool wholeTable = false) {
		if (points is < 0 or > 1440) throw new ArgumentException("A row height is 0 (automatic) to 1440 points.");
		int height = (int)Math.Round(points * 20) * (exact ? -1 : 1);
		FormatTable((table, rows) => { foreach (int r in CellsFor(table, rows, wholeTable).Select(c => c.Row).Distinct()) rows[r].Height = height; });
	}

	/// Sets the widths of the columns of the caret's table, left to right: a list such as "2in, 1.5in, 3cm" (pt, in, cm,
	/// mm; points without a unit). Rows with another number of cells keep theirs.
	public void SetColumnWidths(string widths) {
		var list = widths.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
			.Select(w => RtfTables.CssLength(Regex.IsMatch(w, @"^[\d.]+$") ? w + "pt" : w.Replace(" ", "")) is int t && t >= 144 ? t
				: throw new ArgumentException($"\"{w}\" is not a column width (at least 0.1in; units pt, in, cm, mm).")).ToList();
		FormatTable((table, rows) => {
			foreach (var row in rows)
				if (row.Cells.Count == list.Count)
					for (int c = 0; c < list.Count; c++) row.Cells[c].Width = list[c];
		});
	}

	/// Sets the borders of the selected cells (the whole table with wholeTable = 1): "all", "outside", "inside", "none",
	/// "top", "bottom", "left" or "right", a line width in points and a color (a name such as "black", or "#RRGGBB").
	public void SetCellBorders(string borders = "all", float width = 0.5f, string color = "", bool wholeTable = false) {
		if (width is <= 0 or > 6) throw new ArgumentException("A border is 0.25 to 6 points wide.");
		var border = new RtfTables.Border(Math.Max(5, (int)Math.Round(width * 20)), ParseColor(color));
		FormatTable((table, rows) => RtfTables.SetBorders(rows, CellsFor(table, rows, wholeTable), borders, border));
	}

	/// Sets the background color of the selected cells (the whole table with wholeTable = 1); "" or "none" removes it.
	public void SetCellShading(string color = "", bool wholeTable = false) {
		Color? shading = color.Trim() is "" || color.Trim().Equals("none", StringComparison.OrdinalIgnoreCase) ? null : ParseColor(color);
		FormatTable((table, rows) => { foreach (var (r, c) in CellsFor(table, rows, wholeTable)) rows[r].Cells[c].Shading = shading; });
	}

	/// Sets where the text sits in the selected cells (the whole table with wholeTable = 1): "top", "center" or "bottom".
	public void SetCellVerticalAlignment(string alignment = "top", bool wholeTable = false) {
		int value = alignment.Trim().ToLowerInvariant() switch { "top" => 0, "center" or "middle" => 1, "bottom" => 2, _ => throw new ArgumentException("The alignment is \"top\", \"center\" or \"bottom\".") };
		FormatTable((table, rows) => { foreach (var (r, c) in CellsFor(table, rows, wholeTable)) rows[r].Cells[c].VerticalAlign = value; });
	}

	/// The caret's table's format, one line per row: alignment, height and padding, then each cell's width, borders,
	/// shading and vertical alignment (lengths in points).
	public string TableFormat() {
		var (rows, _) = TableFormats();
		static string Pt(int twips) => (twips / 20.0).ToString("0.##", CultureInfo.InvariantCulture);
		static string Col(Color? c) => c is Color k ? $"#{k.R:X2}{k.G:X2}{k.B:X2}" : "auto";
		return string.Join("\n", rows.Select((row, r) => $"row {r + 1}: align={new[] { "left", "center", "right" }[row.Alignment]} height={(row.Height == 0 ? "auto" : (row.Height < 0 ? "exactly " : "at least ") + Pt(Math.Abs(row.Height)))} padding={Pt(row.Padding)}; "
			+ string.Join("; ", row.Cells.Select((cell, c) => $"cell {c + 1}: width={Pt(cell.Width)} borders={string.Join(",", cell.Borders.Select(b => b.Width == 0 ? "none" : Pt(b.Width) + " " + Col(b.Color)))} shading={(cell.Shading is Color ? Col(cell.Shading) : "none")} valign={new[] { "top", "center", "bottom" }[cell.VerticalAlign]}"))));
	}

	// A color from a name ("black", "LightBlue") or #RRGGBB; null for "" or "auto" (the automatic color).
	static Color? ParseColor(string color) {
		color = color.Trim();
		if (color == "" || color.Equals("auto", StringComparison.OrdinalIgnoreCase) || color.Equals("automatic", StringComparison.OrdinalIgnoreCase)) return null;
		if (RtfTables.CssColor(color) is Color css) return css;
		var named = Color.FromName(color);
		if (named.IsKnownColor) return Color.FromArgb(named.R, named.G, named.B);
		throw new ArgumentException($"\"{color}\" is not a color: use a name such as \"black\" or \"#RRGGBB\".");
	}

	// The Table Properties dialog's changes, in one step.
	internal void ApplyTableChanges(TableChanges changes) =>
		FormatTable((table, rows) => {
			var cells = CellsFor(table, rows, changes.WholeTable);
			if (changes.Alignment is int alignment) rows.ForEach(row => row.Alignment = alignment);
			if (changes.Padding is int padding) rows.ForEach(row => row.Padding = padding);
			if (changes.Height is int height) foreach (int r in cells.Select(c => c.Row).Distinct()) rows[r].Height = height;
			if (changes.Widths is { } widths)
				foreach (var row in rows.Where(row => row.Cells.Count == widths.Count))
					for (int c = 0; c < widths.Count; c++) row.Cells[c].Width = widths[c];
			if (changes.Borders is { } borders) RtfTables.SetBorders(rows, cells, borders, changes.Border);
			if (changes.SetShading) foreach (var (r, c) in cells) rows[r].Cells[c].Shading = changes.Shading;
			if (changes.VerticalAlign is int valign) foreach (var (r, c) in cells) rows[r].Cells[c].VerticalAlign = valign;
		});

	// Selects the next (delta 1) or previous (-1) cell, as Tab and Shift+Tab do; Tab in the last cell adds a row, as in Word.
	internal void NextCell(int delta) {
		if (RtfTables.At(ed, ed.SelectionStart) is not { } at) return;
		var cells = at.Table.Rows.SelectMany(r => r.Cells).ToList();
		int index = at.Table.Rows.Take(at.Row).Sum(r => r.Cells.Count) + at.Cell + delta;
		if (index >= cells.Count && delta > 0) {
			InsertRow(true);
			if (RtfTables.At(ed, ed.SelectionStart) is not { } again) return;
			cells = again.Table.Rows.SelectMany(r => r.Cells).ToList();
		}
		if (index < 0 || index >= cells.Count) return;
		ed.Select(cells[index].Start, cells[index].End - cells[index].Start);
	}

	// After opening: paragraph and cell marks take the size of the text before them, as in Word. HTML has no paragraph
	// marks (they came back at the default size, and list numbers, drawn in the mark's format, with them), and
	// LibreOffice's RTF gives table marks its own default (12 pt, making rows taller). Table row markers take the size of
	// their row's first and last cell.
	bool marksFollowed;    // (done as the document was built: SetHtml)
	void MarksFollowText() {
		if (building == null && marksFollowed) { marksFollowed = false; return; }
		var doc = ed.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		static bool Mark(char c) => c is '\r' or '\u0007' or '\uFFF9' or '\uFFFB' or '\v';
		float SizeAt(int i) => i >= 0 && i < story.Length ? doc.Range2(i, i + 1).Font.Size : -1;
		void SetSize(int i, float size) {
			if (size <= 0) return;
			var r = doc.Range2(i, i + 1);
			if (Math.Abs(r.Font.Size - size) > 0.01) r.Font.Size = size;
		}
		// paragraph and cell marks after text
		for (int i = 1; i < story.Length; i++)
			if (story[i] is '\r' or '\u0007' && !Mark(story[i - 1]) && story[i - 1] != '\uFFFC') SetSize(i, SizeAt(i - 1));
		// table rows: U+FFF9 and its paragraph mark take the first cell's text size, U+FFFB and its mark the last cell's mark
		for (int i = 0; i < story.Length; i++) {
			if (story[i] == '\uFFF9' && i + 2 < story.Length) { float size = SizeAt(i + 2); SetSize(i, size); if (story[i + 1] == '\r') SetSize(i + 1, size); }
			else if (story[i] == '\uFFFB' && i > 0) { float size = SizeAt(i - 1); SetSize(i, size); if (i + 1 < story.Length && story[i + 1] == '\r') SetSize(i + 1, size); }
		}
	}

	// ---- Font sizes (RichEditorNET's set the size of the WinForms selection, which is short in tables, and do sums
	// with the "undefined" size of a selection with several sizes: Decrease made such text 1 point)

	/// Sets the font size of the selection (free-form mode). Where the selection covers a paragraph's whole text, its
	/// paragraph mark gets the size too, so that its list number (and a legal number) follows, as in Word.
	public void SetFontSize(float size) {
		if (size <= 0 || size > 1638) throw new ArgumentOutOfRangeException(nameof(size), "Font sizes are from 1 to 1638 points.");
		ChangeFontSize(_ => size);
	}

	/// Makes each part of the selection points larger (keeping their differences).
	public void IncreaseFontSize(float points = 1) => ChangeFontSize(size => size + points);

	/// Makes each part of the selection points smaller (keeping their differences, down to 1 point).
	public void DecreaseFontSize(float points = 1) => ChangeFontSize(size => size - points);

	void ChangeFontSize(Func<float, float> change) {
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		int start = selection.Start, end = selection.End;
		float New(float size) => Math.Clamp(MathF.Round(change(size > 0 ? size : ed.Font.SizeInPoints) * 2, MidpointRounding.AwayFromZero) / 2, 1, 1638);
		if (start == end) {    // the caret: the size of what is typed next (and of an empty line)
			float size = New(selection.Font.Size);
			AtCaret(font => font.Size = size);
			return;
		}
		var ranges = new List<(int Start, int End)> { (start, end) };
		foreach (var p in LegalLists.Paragraphs(doc)) {
			int textStart = p.Start + p.PrefixLength, textEnd = p.Text.Length > 0 ? p.Start + p.Text.Length : p.Start;
			if (textEnd <= textStart || start > textStart || end < textEnd || p.End <= start || p.Start >= end) continue;
			if (p.PrefixLength > 0) ranges.Add((p.Start, textStart));    // a legal number
			if (p.End > textEnd) ranges.Add((textEnd, p.End));           // the paragraph mark (list number, cell mark)
		}
		int storyEnd = doc.Range2(0, 0).StoryLength;
		if (end >= storyEnd - 1) ranges.Add((storyEnd - 1, storyEnd));    // Select all: the document's last mark too (what is typed at the end)
		doc.BeginEditCollection();
		try {
			foreach (var (from, to) in ranges) {
				// run by run, so that each keeps its own size when sizes are changed by points
				int at = from;
				while (at < to) {
					var run = doc.Range2(at, at);
					run.MoveEnd(tomConstants.tomCharFormat, 1);
					int runEnd = Math.Min(Math.Max(run.End, at + 1), to);
					var part = doc.Range2(at, runEnd);
					part.Font.Size = New(part.Font.Size);
					at = runEnd;
				}
			}
			if (LegalLists.AnyLegal(doc)) LegalLists.Renumber(ed);    // legal numbers' indents fit their new size
		} finally {
			doc.EndEditCollection();
			selection.SetRange(start, end);
		}
	}

	// ---- Iteration 10: legal numbering

	/// Makes the selected paragraphs a legal numbered list (1., 1.1, 1.1.1 by list level), or, if the caret is in one,
	/// ends it for the selected paragraphs. Paragraphs that start with a typed number such as "2.1 " take its level.
	public void LegalNumbering() {
		var range = SelectionRange();
		if (InLegalList()) LegalLists.Remove(ed, range.Start, range.End);
		else LegalLists.Make(ed, range.Start, range.End);
	}

	/// Paragraphs that begin with typed list numbers or bullets ("1. ", "a) ", "(iv) ", "• ", or with a tab after them),
	/// two or more in a row, become real list items: numbering style, punctuation and start from the numbers, levels
	/// from the indents; compound numbers ("1.2") become legal items. Works on the selected paragraphs, or the whole
	/// document when nothing is selected. Returns how many paragraphs changed.
	public int ConvertTypedNumbers() {
		var range = SelectionRange();
		bool all = range.Start == range.End;
		return LegalLists.ConvertTypedLists(ed, all ? 0 : range.Start, all ? ed.TextDocument.Range2(0, 0).StoryLength : range.End, requireTab: false);
	}

	/// True if the caret is in a legal numbered item.
	public bool InLegalList() => LegalLists.IsLegal(ed.TextDocument.Range2(ed.SelectionStart, ed.SelectionStart).Para.ListType);

	/// True if the selection includes a legal numbered item.
	public bool TouchesLegalList() {
		var range = SelectionRange();
		return LegalLists.Touched(ed.TextDocument, range.Start, range.End).Any(p => p.Legal);
	}

	/// Puts every legal number right (the editor does this itself; for scripts that change the text directly).
	/// Returns 1 if anything changed.
	public bool Renumber() {
		bool changed = false;
		LegalLists.Grouped(ed, () => changed = LegalLists.Renumber(ed));
		return changed;
	}

	/// How legal numbers look: "Top" (1. 1.1 1.1.1), "All" (1. 1.1. 1.1.1.) or "None" (1 1.1 1.1.1). Setting it
	/// renumbers the document's legal lists and keeps it in the settings.
	public string LegalFormat {
		get => LegalLists.Format.ToString();
		set {
			if (!Enum.TryParse<LegalLists.Periods>(value, true, out var format)) throw new ArgumentException("LegalFormat is Top, All or None.");
			LegalLists.Format = format;
			form.Settings.LegalFormat = format.ToString();
			form.Settings.Save();
			if (LegalLists.AnyLegal(ed.TextDocument)) LegalLists.Grouped(ed, () => { });
		}
	}

	/// Makes paragraphs that start with typed compound numbers ("1.2<tab>Text") a legal list; returns how many changed.
	/// Opening a document does this by itself.
	public int AdoptTypedNumbers() => LegalLists.AdoptTypedNumbers(ed);

	/// Enter in a legal list: a new item, renumbered; in an empty item, the list ends there. Returns 0 if the caret is
	/// not in a legal item.
	public bool LegalEnter() {
		var selection = ed.TextDocument.Selection2;
		var p = ParagraphAt(selection.Start);
		if (p == null || !p.Legal) return false;
		if (selection.Start == selection.End && p.Text.Length == p.PrefixLength) {
			LegalLists.Remove(ed, p.Start, p.Start);
			return true;
		}
		LegalLists.Grouped(ed, () => selection.TypeText("\r"));
		return true;
	}

	/// Backspace in a legal list: at the start of an item's text, its number goes (it becomes a normal paragraph); with a
	/// selection, the selection goes and the list is renumbered. Returns 0 when RichEdit's own Backspace is right.
	public bool LegalBackspace() {
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) return DeleteLegalSelection();
		var p = ParagraphAt(selection.Start);
		if (p == null || !p.Legal || selection.Start > p.Start + p.PrefixLength) return false;
		LegalLists.Remove(ed, p.Start, p.Start);
		return true;
	}

	/// Delete in a legal list: at the end of an item, the next paragraph joins it (without its number); with a selection,
	/// the selection goes and the list is renumbered. Returns 0 when RichEdit's own Delete is right.
	public bool LegalDelete() {
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) return DeleteLegalSelection();
		var p = ParagraphAt(selection.Start);
		if (p == null || selection.Start != p.End - 1 || p.End >= ed.TextDocument.Range2(0, 0).StoryLength) return false;
		var next = ParagraphAt(p.End);
		if (next == null || (!p.Legal && !next.Legal)) return false;
		LegalLists.Grouped(ed, () => {
			var doc = ed.TextDocument;
			doc.Range2(p.End - 1, next.Start + next.PrefixLength).Text = "";
			KeepFormat(p);
			doc.Selection2.SetRange(p.End - 1, p.End - 1);
		});
		return true;
	}

	// Deletes a selection that includes legal items; the paragraph it ends in keeps the first paragraph's list format
	// (RichEdit would give it the last one's).
	bool DeleteLegalSelection() {
		var selection = ed.TextDocument.Selection2;
		var first = ParagraphAt(selection.Start);
		if (first == null) return false;
		LegalLists.Grouped(ed, () => {
			selection.Text = "";
			KeepFormat(first);
		});
		return true;
	}

	void KeepFormat(LegalLists.Paragraph first) {
		var para = ed.TextDocument.Range2(first.Start, first.Start).Para;
		if (first.Legal) {
			if (!LegalLists.IsLegal(para.ListType) || para.ListLevelIndex != first.Level) LegalLists.SetLegal(para, first.Level, first.ListStart);
		} else if (LegalLists.IsLegal(para.ListType)) {
			para.ListType = LegalLists.WithoutRestart(first.ListType);
			para.SetIndents(0, 0, para.RightIndent);
			var now = ParagraphAt(first.Start);
			if (now != null && LegalLists.Prefix.Match(now.Text) is { Success: true } m && !first.Text.StartsWith(m.Value))
				ed.TextDocument.Range2(now.Start, now.Start + m.Length).Text = "";
		}
	}

	LegalLists.Paragraph? ParagraphAt(int position) =>
		LegalLists.Paragraphs(ed.TextDocument).FirstOrDefault(p => position >= p.Start && position < Math.Max(p.End, p.Start + 1));

	// New tables are as wide as the editor shows text (at most 6.5 inches, the width of a Letter page's text).
	// A new table's width: in Print Layout the page's text width (as in Word; the editor's width depends on the zoom),
	// otherwise what the window shows, at most 6.5 inches.
	int TableWidth() => form.PrintTextWidth is int text ? text
		: ed.ClientSize.Width > 200 ? Math.Min(9360, (ed.ClientSize.Width - 40) * 1440 / Math.Max(96, ed.DeviceDpi)) : 9360;
	/// The document as Markdown. With embedImages = 1 (the default) images are data: URIs, so the text is complete on
	/// its own; with 0 they are written as image1.png, image2.png ... (as when viewing the source).
	public string GetMarkdown(bool embedImages = true) =>
		form.WithoutGaps(() => Footnotes.ToMarkdown(CrossRefs.ToMarkdown(MarkdownConverter.ToMarkdown(ed, embedImages ? null : (number, data, extension) => $"image{number}.{extension}"))));

	// Markdown into the editor, with its cross-references (iteration 50).
	void LoadMarkdown(string markdown, string? folder) {
		MarkdownConverter.Load(ed, CrossRefs.FromMarkdown(Footnotes.FromMarkdown(markdown)), folder);
		CrossRefs.AfterLoad(ed.TextDocument);
	}

	/// Replaces the document with Markdown. Relative image paths are resolved from the current file's folder.
	public void SetMarkdown(string markdown) {
		LoadMarkdown(markdown, form.FileName != null ? Path.GetDirectoryName(form.FileName) : Environment.CurrentDirectory);
		form.StylesAfterLoad(form.Styles, takeLook: true);
	}

	/// The document view: "Rich" (the formatted text), "Markdown" or "HTML" (the source view, see SideBySide).
	public string View {
		get => form.View switch { WriterForm.SourceKind.Markdown => "Markdown", WriterForm.SourceKind.Html => "HTML", _ => "Rich" };
		set {
			var kind = value.Trim().ToLowerInvariant() switch {
				"rich" or "richtext" or "rich text" or "formatted" or "" => WriterForm.SourceKind.None,
				"markdown" or "md" => WriterForm.SourceKind.Markdown,
				"html" or "htm" => WriterForm.SourceKind.Html,
				_ => throw new ArgumentException("View is \"Rich\", \"Markdown\" or \"HTML\".")
			};
			form.SetView(kind, form.SideBySide);
		}
	}

	/// Whether the Markdown or HTML view shows the formatted text beside it (setting it on in the rich text view shows Markdown).
	public bool SideBySide {
		get => form.SideBySide;
		set => form.SetView(form.View == WriterForm.SourceKind.None ? (value ? WriterForm.SourceKind.Markdown : WriterForm.SourceKind.None) : form.View, value);
	}

	/// The Markdown or HTML in the source view ("" in the rich text view); setting it replaces that text and loads it.
	public string SourceText {
		get => form.SourceText;
		set => form.SetSourceText(value);
	}

	// The document as the source view shows it (and a .md / .html file gets it): Markdown with the header and footer
	// line, or a whole HTML page; pictures as data: URIs.
	internal string SourceOf(WriterForm.SourceKind kind) => kind == WriterForm.SourceKind.Markdown
		? form.WithoutGaps(() => Footnotes.ToMarkdown(CrossRefs.ToMarkdown(MarkdownConverter.ToMarkdown(ed, null)))) + (OwnBands ? HeaderFooter.MarkdownLine(form.Bands) + "\n" : "") + LineNumberFormats.MarkdownLine(form.Numbering)
		: HtmlFile(form.FileName != null ? Path.GetFileNameWithoutExtension(form.FileName) : "Untitled", pretty: true);

	// Replaces the document with Markdown or HTML from the source view, as Open does with a file's.
	internal void LoadSource(WriterForm.SourceKind kind, string text) {
		if (kind == WriterForm.SourceKind.Markdown) {
			form.Bands = HeaderFooter.FromMarkdown(text) ?? form.DefaultBands();
			form.SetLineNumbering(LineNumberFormats.FromMarkdown(text) ?? new LineNumbering());
			LoadMarkdown(text, form.FileName != null ? Path.GetDirectoryName(form.FileName) : Environment.CurrentDirectory);
		} else {
			form.Bands = HeaderFooter.FromHtml(text) ?? form.DefaultBands();
			form.SetLineNumbering(LineNumberFormats.FromHtml(text) ?? new LineNumbering());
			SetHtml(text);
		}
		PageBreaks.Normalize(ed);
		MarksFollowText();
		if (LegalLists.AnyLegal(ed.TextDocument)) LegalLists.Renumber(ed);
	}

	/// The document as Markdown (get), or replaced by Markdown (set); see GetMarkdown and SetMarkdown.
	public string Markdown {
		get => GetMarkdown();
		set => SetMarkdown(value);
	}

	// Saving document.md writes its images to document_images\image1.png ... and refers to them from the Markdown.
	// Images left from an earlier save are removed first.
	static Func<int, byte[], string, string> ImageFiles(string markdownPath) {
		string name = Path.GetFileNameWithoutExtension(markdownPath) + "_images";
		string folder = Path.Combine(Path.GetDirectoryName(markdownPath)!, name);
		if (Directory.Exists(folder))
			foreach (var old in Directory.GetFiles(folder).Where(f => Regex.IsMatch(Path.GetFileName(f), @"^image\d+\.(png|jpg|gif|bmp|tiff)$", RegexOptions.IgnoreCase)))
				File.Delete(old);
		return (number, data, extension) => {
			Directory.CreateDirectory(folder);
			File.WriteAllBytes(Path.Combine(folder, $"image{number}.{extension}"), data);
			return Uri.EscapeDataString(name) + $"/image{number}.{extension}";
		};
	}

	// ---- Text and selection

	/// Pastes the clipboard's text only (iteration 65), taking the formatting where it lands, as Ctrl+Shift+V does.
	public void PastePlainText() => form.PastePlainText();

	/// Inserts text at the caret, replacing the selection; the caret ends up after the new text.
	public void InsertText(string text) {
		ed.SelectedText = text;
		ed.ScrollToCaret();
	}

	/// Selects length characters from start (0-based).
	public void Select(int start, int length = 0) {
		ed.Select(Math.Clamp(start, 0, ed.TextLength), Math.Max(0, length));
		ed.ScrollToCaret();
	}

	/// Moves the caret; a negative position goes to the end of the document.
	public void GoTo(int position) => Select(position < 0 ? ed.TextLength : position);

	/// Finds the next occurrence after the selection (wrapping round), selects it and returns its position, or -1.
	public int Find(string text, bool matchCase = false, bool wholeWord = false) {
		if (string.IsNullOrEmpty(text)) throw new ArgumentException("There is nothing to find.");
		var options = (matchCase ? RichTextBoxFinds.MatchCase : 0) | (wholeWord ? RichTextBoxFinds.WholeWord : 0);
		int from = ed.SelectionStart + ed.SelectionLength;
		int index = from < ed.TextLength ? FindVisible(text, from, options) : -1;
		if (index < 0) index = FindVisible(text, 0, options);
		if (index >= 0) { ed.Select(index, text.Length); ed.ScrollToCaret(); }
		return index;
	}

	// The next match from a position that is all in text you see (-1: none), not selected. (Iteration 57) A match in
	// hidden text, the markers of cross-references, tables of contents and footnotes and a footnote's note, isn't one:
	// Replace All of "toc" took the table of contents apart, changing its markers.
	int FindVisible(string text, int from, RichTextBoxFinds options) {
		while (from < ed.TextLength) {
			int index = ed.Find(text, from, options | RichTextBoxFinds.NoHighlight);
			if (index < 0) return -1;
			if (ed.TextDocument.Range2(index, index + text.Length).Font.Hidden == tomConstants.tomFalse) return index;
			from = index + 1;
		}
		return -1;
	}

	/// Replaces every occurrence and returns the number of replacements.
	public int Replace(string text, string replacement, bool matchCase = false, bool wholeWord = false) {
		if (string.IsNullOrEmpty(text)) throw new ArgumentException("There is nothing to find.");
		var options = (matchCase ? RichTextBoxFinds.MatchCase : 0) | (wholeWord ? RichTextBoxFinds.WholeWord : 0);
		int count = 0, position = 0;
		// (iteration 57: the text you see only, not the hidden markers; and as one undo step)
		ed.TextDocument.BeginEditCollection();
		try {
			while (position < ed.TextLength) {
				int index = FindVisible(text, position, options);
				if (index < 0) break;
				ed.Select(index, text.Length);
				ed.SelectedText = replacement;
				count++;
				position = index + replacement.Length;
			}
		} finally {
			ed.TextDocument.EndEditCollection();
		}
		return count;
	}

	public int WordCount() => Regex.Matches(CrossRefs.PlainText(ed.Text), @"[\p{L}\p{N}][\p{L}\p{N}'’\-]*").Count;    // (not the cross-references' hidden markers)
	public int CharacterCount() => ed.TextLength;

	/// Changes the case of the selection, keeping its formatting: "upper", "lower", "title", "headline", "sentence" or
	/// "toggle". Headline Case: small words lower case, abbreviations as written (rte_headline.cs).
	public void ChangeCase(string mode) {
		mode = mode.ToLowerInvariant();
		int type = mode switch {
			"upper" => tomConstants.tomUpperCase,
			"lower" => tomConstants.tomLowerCase,
			"title" => tomConstants.tomTitleCase,
			"headline" => 0,
			"sentence" => tomConstants.tomSentenceCase,
			"toggle" => tomConstants.tomToggleCase,
			_ => throw new ArgumentException("Case must be upper, lower, title, headline, sentence or toggle.")
		};
		// (iteration 57: TOM's ChangeCase is no undo step, so Ctrl+Z undid what came before it. The new text is worked out
		// with it while undo is held, the text put back, and then the characters that change put in as one undo step,
		// each keeping its own formatting)
		var doc = ed.TextDocument;
		var range = SelectionRange();
		int start = range.Start;
		string before = range.Text ?? "", after;
		if (before.Length == 0) return;
		if (mode == "headline") after = HeadlineCase.Apply(before, HiddenMask(start, start + before.Length));
		else {
			doc.Undo(-9999995);    // tomSuspend
			try {
				range.ChangeCase(type);
				after = range.Text ?? "";
				if (after.Length != before.Length) return;    // (a letter that became two: as TOM made it, without an undo step)
				PutText(doc, start, after, before);    // (back as it was)
			} finally {
				doc.Undo(-9999994);    // tomResume
			}
		}
		doc.BeginEditCollection();
		try { PutText(doc, start, before, after, skipHidden: true); }
		finally { doc.EndEditCollection(); }
		doc.Selection2.SetRange(start, start + before.Length);
	}

	// The text from start, now as it is, made as it should be: the characters that differ, one by one (each keeps its
	// formatting). skipHidden: hidden text left as it is (the cross-references', tables of contents' and footnotes'
	// markers, a footnote's note: a marker in another case is no marker, as Change Case made "[[xreF:_TOC1]]")
	static void PutText(ITextDocument2 doc, int start, string now, string want, bool skipHidden = false) {
		for (int i = 0; i < want.Length; i++) {
			if (want[i] == now[i] || want[i] == '\r' || now[i] == '\r') continue;
			var range = doc.Range2(start + i, start + i + 1);
			if (skipHidden && range.Font.Hidden == tomConstants.tomTrue) continue;
			range.Text = want[i].ToString();
		}
	}

	// Which characters from start to end are hidden text (by runs of the same formatting, TOM tomCharFormat).
	bool[] HiddenMask(int start, int end) {
		var mask = new bool[end - start];
		var run = ed.TextDocument.Range2(start, start);
		while (run.Start < end) {
			if (run.MoveEnd(tomConstants.tomCharFormat, 1) == 0 || run.End <= run.Start) break;
			if (run.Font.Hidden == tomConstants.tomTrue)
				for (int i = Math.Max(run.Start, start); i < Math.Min(run.End, end); i++) mask[i - start] = true;
			run.SetRange(run.End, run.End);
		}
		return mask;
	}

	// ---- Formatting (these wrap RichEditBox methods of the same name)

	/// RichEditorNET 1.0.11: ToggleSuperscript sets Superscript and then clears Subscript, and RichEdit clears both,
	/// so turning it on has no effect. If that happens, the format is set through TOM (fixed upstream in the PR).
	public void ToggleSuperscript() => ToggleScript(true);
	public void ToggleSubscript() => ToggleScript(false);

	void ToggleScript(bool superscript) {
		var caret = ed.TextDocument.Selection2;
		if (caret.Start == caret.End) {    // the caret: what is typed next (and an empty line)
			bool on = (superscript ? caret.Font.Superscript : caret.Font.Subscript) != tomConstants.tomTrue;
			AtCaret(font => { if (superscript) font.Superscript = on ? tomConstants.tomTrue : tomConstants.tomFalse; else font.Subscript = on ? tomConstants.tomTrue : tomConstants.tomFalse; });
			return;
		}
		int IsOn() { var font = SelectionRange().Font; return superscript ? font.Superscript : font.Subscript; }
		bool wasOn = IsOn() == tomConstants.tomTrue;
		if (superscript) ed.ToggleSuperscript(); else ed.ToggleSubscript();
		if (wasOn || IsOn() == tomConstants.tomTrue) return;
		var range = SelectionRange().Font;
		if (superscript) range.Superscript = tomConstants.tomTrue; else range.Subscript = tomConstants.tomTrue;
	}

	/// Toggles a list; numbered lists start at 1 (RichEditorNET 1.0.11 starts new decimal lists at 0).
	public void ToggleList(ListStyle style) {
		if (InLegalList()) {    // a legal list becomes a list of this style (its numbers go)
			LegalLists.Remove(ed, SelectionRange().Start, SelectionRange().End);
		}
		else if ((SelectionRange().Para.ListType & 0xFFFF) == (int)style) {    // the same style again ends the list, whatever its punctuation
			RemoveList();
			return;
		}
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			ToggleFromTopLevel(() => ed.ToggleList(style));
			var para = SelectionRange().Para;
			if (style != ListStyle.Bullet && (para.ListType & 0xFFFF) == (int)style && para.ListStart == 0) para.ListStart = 1;
			FitLists();    // (the items' indents: their level's, Paragraph > List levels)
			LegalLists.BreakChains(doc);    // (right after legal items: numbered from 1, not on from them)
		} finally {
			doc.EndEditCollection();
		}
	}

	/// Iteration 38: the selected list items get their level's indents (the document's list levels: Number at, Text at).
	void FitLists() {
		var range = SelectionRange();
		var paragraphs = ed.TextDocument.Range2(range.Start, range.End);
		paragraphs.Expand(tomConstants.tomParagraph);
		OfficeConverters.FitListIndents(ed.TextDocument, paragraphs.Start, Math.Max(paragraphs.Start, paragraphs.End - 1));
	}

	/// Sets the block style of the selected paragraphs: their style (Normal, Heading 1 to 6, Code), as SetStyle.
	public void SetBlockStyle(BlockStyle style) {
		SetStyle(style switch { BlockStyle.Paragraph => "Normal", BlockStyle.Preformatted => "Code", _ => "Heading " + (int)style });
	}

	// (the block styles as RichEditorNET had them, before styles)
	void SetBlockStyleAsBefore(BlockStyle style) {
		ed.SetBlockStyle(style);
		// RichEditorNET formats only the paragraph's text, so on an empty paragraph the style would be lost: apply it to
		// the paragraph mark and the caret too, so that what is typed next gets it.
		var para = ed.TextDocument.Range2(ed.SelectionStart, ed.SelectionStart);
		para.Expand(tomConstants.tomParagraph);
		if (ed.SelectionLength == 0 && (para.Text ?? "").Trim('\r', '\n') == "") {
			ApplyBlockFont(ed.TextDocument.Range2(para.Start, para.End).Font, style);
			ApplyBlockFont(ed.TextDocument.Selection2.Font, style);
		}
	}

	// The library's block style fonts (internal), or the same values if they cannot be reached.
	static readonly Type? blockStyleHelper = typeof(RichEditBox).Assembly.GetType("EllipticBit.RichEditorNET.Formatting.BlockStyleHelper");
	static T BlockConstant<T>(string name, T fallback) =>
		blockStyleHelper?.GetField(name, BindingFlags.NonPublic | BindingFlags.Static)?.GetValue(null) is T value ? value : fallback;

	void ApplyBlockFont(ITextFont font, BlockStyle style) {
		switch (style) {
			case BlockStyle.Preformatted:
				font.Bold = tomConstants.tomFalse;
				font.Size = BlockConstant("PreformattedPointSize", 10f);
				font.Name = BlockConstant("PreformattedFontName", "Courier New");
				break;
			case BlockStyle.Paragraph:
				font.Bold = tomConstants.tomFalse;
				font.Size = BlockConstant("ParagraphPointSize", 12f);
				font.Name = ed.Font.Name;
				break;
			default:
				font.Bold = tomConstants.tomTrue;
				font.Size = BlockConstant("HeadingPointSizes", new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f })[(int)style];
				font.Name = ed.Font.Name;
				break;
		}
	}

	public void ClearFontColor() => Format(() => SelectionRange().Font.ForeColor = tomConstants.tomAutoColor, font => font.ForeColor = tomConstants.tomAutoColor);
	public void ClearHighlight() => Format(() => SelectionRange().Font.BackColor = tomConstants.tomAutoColor, font => font.BackColor = tomConstants.tomAutoColor);

	// ---- Character formats (iteration 23): with nothing selected, for what is typed next (and an empty line), as in Word

	/// Toggles bold: of the selected text; with nothing selected, of what is typed next (and of an empty line).
	public void ToggleBold() => Format(ed.ToggleBold, Toggled(font => font.Bold, (font, value) => font.Bold = value));
	/// Toggles italic, as ToggleBold.
	public void ToggleItalic() => Format(ed.ToggleItalic, Toggled(font => font.Italic, (font, value) => font.Italic = value));
	/// Toggles strikethrough, as ToggleBold.
	public void ToggleStrikeThrough() => Format(ed.ToggleStrikeThrough, Toggled(font => font.StrikeThrough, (font, value) => font.StrikeThrough = value));

	/// Toggles an underline style (Single, Double, Dotted, Dashed, Wavy), as ToggleBold.
	public void ToggleUnderline(UnderlineStyle style = UnderlineStyle.Single) {
		int tom = style switch { UnderlineStyle.Double => 3, UnderlineStyle.Dotted => 4, UnderlineStyle.Dashed => 5, UnderlineStyle.Wavy => 8, _ => 1 };    // tomSingle, tomDouble ...
		int now = ed.TextDocument.Selection2.Font.Underline;
		Format(() => ed.ToggleUnderline(style), font => font.Underline = now == tom ? 0 : tom);    // (0: tomNone)
	}

	/// Sets bold, italic, single underline and strikethrough, as SetFontColor: 1 on, 0 off, -1 as each part is.
	public void SetFontStyle(int bold = -1, int italic = -1, int underline = -1, int strikeThrough = -1) => SetFontLook(bold, italic, underline, strikeThrough, null);

	// (a selection's changed at once, through a duplicate of its font: one undo step)
	internal void SetFontLook(int bold, int italic, int underline, int strikeThrough, Color? color, string? name = null, float? size = null) {
		void Apply(ITextFont font) {
			if (name != null) font.Name = name;
			if (size is { } points) font.Size = points;
			if (color is { } c) font.ForeColor = ColorTranslator.ToWin32(c);
			if (bold >= 0) font.Bold = bold > 0 ? -1 : 0;
			if (italic >= 0) font.Italic = italic > 0 ? -1 : 0;
			if (underline >= 0) font.Underline = underline > 0 ? 1 : 0;
			if (strikeThrough >= 0) font.StrikeThrough = strikeThrough > 0 ? -1 : 0;
		}
		Format(() => {
			var selection = ed.TextDocument.Selection2;
			var look = selection.Font.Duplicate;
			Apply(look);
			selection.Font = look;
		}, Apply);
	}

	/// Sets the font colour: of the selected text; with nothing selected, of what is typed next (and of an empty line).
	public void SetFontColor(Color color) => Format(() => ed.SetFontColor(color), font => font.ForeColor = ColorTranslator.ToWin32(color));
	/// Sets the highlight (background) colour, as SetFontColor.
	public void SetBackgroundColor(Color color) => Format(() => ed.SetBackgroundColor(color), font => font.BackColor = ColorTranslator.ToWin32(color));

	// A format: RichEditorNET's (or ours) for selected text, AtCaret's with nothing selected.
	void Format(Action selected, Action<ITextFont> caret) {
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) selected(); else AtCaret(caret);
	}

	// An on/off format at the caret: on unless what is typed next has it (the mark of an empty line gets the same).
	Action<ITextFont> Toggled(Func<ITextFont, int> get, Action<ITextFont, int> set) {
		int value = get(ed.TextDocument.Selection2.Font) == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
		return font => set(font, value);
	}

	/// True if the caret is in a list item.
	public bool InList() => (SelectionRange().Para.ListType & 0xFFFF) != tomConstants.tomListNone;

	/// Nests (positive) or un-nests (negative) the selected list items by that many levels, like Tab / Shift+Tab.
	/// With automatic level styles, numbered items take the style of their new level: 1. -> a. -> i. -> 1. ...
	public void Nest(int levels = 1) {
		if (NestStyled(levels)) return;    // (numbered styles' paragraphs: their sub-levels)
		if (TouchesLegalList()) {    // legal items: their new levels' indents and numbers, in one undo step
			var range = SelectionRange();
			int start = range.Start, end = range.End;
			LegalLists.Grouped(ed, () => {
				for (int i = 0; i < Math.Abs(levels); i++) NestOnce(Math.Sign(levels));
				LegalLists.Reindent(ed.TextDocument, start, Math.Max(start, end));
			});
			return;
		}
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			for (int i = 0; i < Math.Abs(levels); i++) NestOnce(Math.Sign(levels));
			FitLists();    // (their new level's indents)
		} finally {
			doc.EndEditCollection();
		}
	}


	// Numbered styles' sub-levels: the selected items of numbered styles go levels down (or up) within their style,
	// never above its own level, each level the style's SubLevelIndent further in; going down (Tab), a paragraph of a numbered style
	// whose numbering is off becomes an item a level below its style's, a letter or number typed in front of it taken
	// out (its number is the style's now). One undo step. False if the selection has no paragraph of a numbered style.
	bool NestStyled(int levels) {
		var doc = ed.TextDocument;
		var range = SelectionRange();
		var touched = LegalLists.Touched(doc, range.Start, range.End).Where(p => form.Styles.ById(p.Style) is { Numbered: true }).ToList();
		if (touched.Count == 0 || touched.All(p => !p.Legal) && levels < 0) return false;
		doc.BeginEditCollection();
		try {
			foreach (var p in Enumerable.Reverse(touched)) {    // (from the last: a typed number that goes moves the text after it)
				var style = form.Styles.ById(p.Style)!;
				int top = Math.Clamp(style.NumberLevel, 1, 9) - 1;
				var para = doc.Range2(p.Start, p.Start).Para;
				int level;
				if (p.Legal) {
					level = Math.Clamp(p.Level + levels, top, 8);
					if (level == p.Level) continue;
					para.ListLevelIndex = level;
				} else {
					if (levels < 0) continue;
					if (TypedNumber(p.Text) is > 0 and var typed) doc.Range2(p.Start, p.Start + typed).Text = "";
					level = Math.Min(8, top + levels);
					para = doc.Range2(p.Start, p.Start).Para;
					LegalLists.SetStyled(para, level);
				}
				para.SetIndents(style.FirstLine, style.LeftIndent + (level - top) * style.SubLevelIndent, para.RightIndent);
				LegalLists.NoNumberSpace(para);
			}
			StyleEngine.Renumber(doc);
		} finally {
			doc.EndEditCollection();
		}
		return true;
	}

	/// The caret at the start of an unnumbered paragraph of a numbered style (or in a number typed in front of it):
	/// Tab makes it a sub-level item there, as on a numbered paragraph, instead of typing a tab. So too a selection over
	/// several paragraphs, all of numbered styles (the a.-e. under a "1." item, their numbering off).
	public bool TabMakesSubLevel() {
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		if (selection.Start != selection.End) {
			var touched = LegalLists.Touched(doc, selection.Start, selection.End);
			return touched.Count > 1 && touched.All(p => form.Styles.ById(p.Style) is { Numbered: true });
		}
		var p = LegalLists.Touched(doc, selection.Start, selection.Start).FirstOrDefault();
		if (p == null || p.Legal || form.Styles.ById(p.Style) is not { Numbered: true }) return false;
		return selection.Start <= p.Start + TypedNumber(p.Text);
	}

	/// The length of a number typed in front of a paragraph ("a. ", "(ii) ", "3) ", with its spaces or tab), or 0: a
	/// number, one letter or a roman numeral, with its punctuation (not just any word with a space after it).
	static int TypedNumber(string text) {
		var m = LegalLists.TypedMarker.Match(text);
		if (!m.Success || m.Groups["bullet"].Success || !m.Groups["close"].Success) return 0;
		string n = m.Groups["n"].Value;
		bool roman = n.All(ch => "ivxlcdm".Contains(ch)) || n.All(ch => "IVXLCDM".Contains(ch));
		return char.IsDigit(n[0]) || n.Length == 1 || roman ? m.Length : 0;
	}

	void NestOnce(int delta) {
		int selStart = ed.SelectionStart, selLength = ed.SelectionLength, selEnd = selStart + selLength;
		var starts = new List<int>();
		var r = ed.TextDocument.Range2(selStart, selStart);
		r.StartOf(tomConstants.tomParagraph, 0);
		do starts.Add(r.Start);
		while (selLength > 0 && r.Move(tomConstants.tomParagraph, 1) != 0 && r.Start < selEnd && r.Start < ed.TextLength);
		try {
			foreach (int start in starts) {
				ed.Select(start, 0);
				var para = SelectionRange().Para;
				if ((para.ListType & 0xFFFF) == tomConstants.tomListNone) continue;
				int before = para.ListLevelIndex;
				if (delta > 0) ed.IncreaseListLevel(); else ed.DecreaseListLevel();
				para = SelectionRange().Para;
				int level = para.ListLevelIndex;
				// numbered items take the numbering of their new level from Paragraph > List levels
				if (level != before && form.AutoLevelStyles && (para.ListType & 0xFFFF) != tomConstants.tomListBullet) ApplyLevel(para, level);
			}
		} finally {
			ed.Select(selStart, selLength);
		}
	}

	// ---- Images

	/// Inserts a PNG, JPEG or GIF image at the caret. Without a width and height it keeps the image's own size,
	/// scaled down to at most 600 pixels wide; with just one of them, the other follows the image's proportions.
	/// (RichEditorNET shows inserted images at DefaultImageWidth x DefaultImageHeight, so those are set for the insert.)
	public void InsertImageFile(string path, string altText = "", int width = 0, int height = 0) {
		byte[] data = File.ReadAllBytes(path);
		using (var buffer = new MemoryStream(data))
		using (var image = Image.FromStream(buffer)) {
			double aspect = (double)image.Height / image.Width;
			if (width <= 0 && height <= 0) {
				width = Math.Min(image.Width, 600);
				height = (int)Math.Round(width * aspect);
			} else if (height <= 0) {
				height = (int)Math.Round(width * aspect);
			} else if (width <= 0) {
				width = (int)Math.Round(height / aspect);
			}
		}
		int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
		ed.DefaultImageWidth = Math.Max(1, width);
		ed.DefaultImageHeight = Math.Max(1, height);
		try {
			using var stream = new MemoryStream(data);
			ed.InsertImage(stream, altText);
		} finally {
			ed.DefaultImageWidth = oldWidth;
			ed.DefaultImageHeight = oldHeight;
		}
	}

	public void InsertLinkedThumbnailFile(string imageUrl, string path, string altText = "") {
		using var stream = File.OpenRead(path);
		ed.InsertLinkedThumbnail(imageUrl, stream, altText);
	}

	// ---- Events (RteClient2.ahk: Rte.On / Rte.Off subscribe A_ScriptHwnd with these)

	/// Sends an event to an AHK window (A_ScriptHwnd) from now on.
	public void Subscribe(string eventName, long window) => form.Events.Subscribe(ParseEvent(eventName), (IntPtr)window);

	/// Stops sending an event ("*" for all events) to an AHK window.
	public void Unsubscribe(string eventName, long window) =>
		form.Events.Unsubscribe(eventName == "*" ? null : ParseEvent(eventName), (IntPtr)window);

	/// The event names, one per line.
	public string EventNames => string.Join("\n", Enum.GetNames<EditorEvent>());

	/// The error of the last call made with Rte.Post that failed.
	public string LastAsyncError { get; internal set; } = "";

	static EditorEvent ParseEvent(string name) =>
		Enum.TryParse<EditorEvent>(name, true, out var e) && Enum.IsDefined(e) ? e
			: throw new ArgumentException($"Unknown event \"{name}\". Events: {string.Join(", ", Enum.GetNames<EditorEvent>())}.");

	// ---- Iteration 3: recent files, settings, printing

	/// The recently opened or saved files, newest first, one per line.
	public string RecentFiles => string.Join("\n", form.Settings.RecentFiles);

	/// The settings file (window, view settings, recent files, page setup).
	public string SettingsFile => Settings.FilePath;

	/// Exports the document as a PDF file through the "Microsoft Print to PDF" printer, and returns its full path.
	/// The page setup (File > Page Setup) applies. This can take a few seconds. bookmarks (1, as at first): a bookmark
	/// for each heading, under the one of the level above it, and each table of contents' entry a link to its heading
	/// (iteration 53); 0: the pages only.
	public string ExportPdf(string path, bool bookmarks = true) {
		string pdf = ExportPdfFile(path, bookmarks);
		form.BeginInvoke(() => form.RunHook(HookEvent.Printed, pdf));    // (iteration 56)
		return pdf;
	}

	string ExportPdfFile(string path, bool bookmarks) {
		const string pdfPrinter = "Microsoft Print to PDF";
		if (!PrinterSettings.InstalledPrinters.Cast<string>().Contains(pdfPrinter))
			throw new InvalidOperationException($"The \"{pdfPrinter}\" printer is not installed (Windows Features: Microsoft Print to PDF).");
		string full = Path.GetFullPath(path);
		if (File.Exists(full)) File.Delete(full);
		form.UpdateReferences();
		using var guard = new DefaultPrinterGuard(form.Settings);    // (the PDF printer doesn't become the default)
		(List<PdfMarks.Bookmark> Bookmarks, List<PdfMarks.Link> Links) places = (new(), new());
		using (var printer = form.CreatePrinter(pdfPrinter, dialog: true)) {
			var document = printer.Document;
			document.PrinterSettings.PrintToFile = true;
			document.PrinterSettings.PrintFileName = full;
			document.PrintController = new StandardPrintController();    // no "Printing..." window
			// (iteration 53) where the headings and the tables of contents' entries are printed, as this printer lays out
			// the pages (the gaps of the page view out, as for printing)
			if (bookmarks) places = printer.PdfPlaces(Toc.Headings(ed.TextDocument, form.Styles, 6));
			// (iteration 58) page by page in the status bar; Esc stops it
			form.EscPressed();    // (an Esc from before: not this one's)
			// (after the last page, the printer puts the file together, which takes a while for a long document: said)
			printer.Progress = (page, pages) => form.ShowProgress(page < pages ? T("status.pdfPage", "Exporting as PDF: page {0} of {1}. Esc stops it.", page, pages)
				: TN("status.pdfFinishing", "Exporting as PDF: Windows is finishing the file ({0} page)…", "Exporting as PDF: Windows is finishing the file ({0} pages)…", pages));
			printer.CancelRequested = form.PrintStop;
			document.Print();
			if (printer.Cancelled) {
				// (the printer writes what it had in the background: the part file deleted once it is let go of, or comes)
				var wait = System.Diagnostics.Stopwatch.StartNew();
				while (wait.ElapsedMilliseconds < 8000) {
					if (File.Exists(full)) {
						try { File.Delete(full); if (!File.Exists(full)) break; } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
					} else if (wait.ElapsedMilliseconds > 2000) break;
					Thread.Sleep(100);
				}
				throw new OperationCanceledException(T("status.pdfStopped", "The PDF export was stopped (Esc): no file was made."));
			}
		}
		// The printer writes the file in the background; wait until it is complete.
		var clock = System.Diagnostics.Stopwatch.StartNew();
		bool done = false;
		while (clock.ElapsedMilliseconds < 20000 && !done) {
			try {
				if (File.Exists(full)) {
					using var file = File.Open(full, FileMode.Open, FileAccess.Read, FileShare.None);
					done = file.Length > 0;
				}
			} catch (IOException) { }
			if (!done) Thread.Sleep(100);
		}
		if (!done) throw new TimeoutException(T("status.pdfTimeout", "The PDF printer did not finish the file within 20 s."));
		if (places.Bookmarks.Count + places.Links.Count > 0) {
			form.ShowProgress(T("status.pdfMarks", "Exporting as PDF: adding the bookmarks and links…"));
			try { PdfMarks.Add(full, places.Bookmarks, places.Links); }
			catch (Exception ex) { form.SetStatus("⚠ " + T("status.pdfNoBookmarks", "The PDF was made, but its bookmarks couldn't be put in: {0}", ex.Message)); }
		}
		return full;
	}

	/// Prints without a dialog, to the default printer or the printer named. The page setup applies.
	public void Print(string printerName = "") {
		form.UpdateReferences();
		using var printer = form.CreatePrinter(printerName == "" ? null : printerName);
		printer.Document.PrintController = new StandardPrintController();
		printer.Document.Print();
		form.BeginInvoke(() => form.RunHook(HookEvent.Printed, printer.Document.PrinterSettings.PrinterName));    // (iteration 56)
	}

	// ---- Iteration 8: list levels (Paragraph > List levels)

	/// Ends the lists of the selected paragraphs, whatever their numbering (bullets, numbers, legal numbers, or a
	/// Word document's numbered style): their numbers go, their text and indents stay. One undo step.
	public void RemoveList() {
		var range = SelectionRange();
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			// a numbered style's items: their numbers go, their style and indents stay (as a Word paragraph's numbering
			// turned off; they stay unnumbered when the document is opened again)
			foreach (var p in Enumerable.Reverse(LegalLists.Touched(doc, range.Start, range.End)).Where(p => p.Legal && p.Styled)) {
				if (p.PrefixLength > 0) doc.Range2(p.Start, p.Start + p.PrefixLength).Text = "";
				doc.Range2(p.Start, p.Start).Para.ListType = tomConstants.tomListNone;
			}
			range = SelectionRange();
			if (TouchesLegalList()) LegalLists.Remove(ed, range.Start, range.End);
			range = SelectionRange();
			// (ordinary items with their level's indents, the number where the level has it: their style's again, as
			// before they were in the list)
			foreach (var p in LegalLists.Touched(doc, range.Start, range.End).Where(p => (p.ListType & 0xFFFF) != tomConstants.tomListNone && !LegalLists.IsLegal(p.ListType))) {
				var para = doc.Range2(p.Start, p.Start).Para;
				var (first, left) = LegalLists.IndentsFor(p.Level, 18);
				if (Math.Abs(para.LeftIndent + para.FirstLineIndent - (left + first)) > 0.5f) continue;    // (the number's place)
				var style = StyleEngine.StyleAt(doc, form.Styles, p.Start);
				para.SetIndents(style.FirstLine, style.LeftIndent, para.RightIndent);
			}
			var paragraphs = doc.Range2(range.Start, range.End);
			paragraphs.Expand(tomConstants.tomParagraph);
			paragraphs.Para.ListType = tomConstants.tomListNone;
			StyleEngine.Renumber(doc);    // (the items after them)
		} finally {
			doc.EndEditCollection();
		}
		form.StylesChanged();
	}

	/// Makes the selected paragraphs a numbered list (or ends it), numbered as List levels sets for their levels.
	public void NumberedList() {
		if (InLegalList()) {    // a legal list becomes an ordinary numbered list, keeping its levels
			var range = SelectionRange();
			LegalLists.Grouped(ed, () => {
				var doc = ed.TextDocument;
				foreach (var p in Enumerable.Reverse(LegalLists.Touched(doc, range.Start, range.End)).Where(p => p.Legal)) {
					if (p.PrefixLength > 0) doc.Range2(p.Start, p.Start + p.PrefixLength).Text = "";
					var para = doc.Range2(p.Start, p.Start).Para;
					para.ListType = tomConstants.tomListNumberAsArabic | tomConstants.tomListPeriod;
					ApplyLevel(para, p.Level);
				}
			});
			return;
		}
		int type = SelectionRange().Para.ListType & 0xFFFF;
		if (type != tomConstants.tomListNone && type != tomConstants.tomListBullet) {
			// numbered already: the list ends (RichEditorNET's toggle ends it only for its own plain style, so a list with
			// "1." or "(1)", as Word's documents have, would only have changed its punctuation)
			RemoveList();
			return;
		}
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			ToggleFromTopLevel(() => ed.ToggleList(ListStyle.Decimal));
			ForEachSelectedParagraph(para => ApplyLevel(para, para.ListLevelIndex));
			LegalLists.BreakChains(doc);    // (right after legal items: numbered from 1, not on from them)
		} finally {
			doc.EndEditCollection();
		}
	}

	// Runs a list toggle so that paragraphs that were not list items start at the top level. RichEdit keeps the level of
	// a list item that Enter ended (on an empty item), so a list started there later took that level's numbering style,
	// without its indent.
	void ToggleFromTopLevel(Action toggle) {
		var notInList = new List<bool>();
		ForEachSelectedParagraph(para => notInList.Add((para.ListType & 0xFFFF) == tomConstants.tomListNone));
		toggle();
		int i = 0;
		ForEachSelectedParagraph(para => {
			if (i < notInList.Count && notInList[i++] && (para.ListType & 0xFFFF) != tomConstants.tomListNone && para.ListLevelIndex != 0) para.ListLevelIndex = 0;
		});
	}

	/// Applies List levels to the whole list the caret is in: each item gets the numbering and indents of its level
	/// (legal numbering's items, all of them: one numbering).
	public void ApplyListLevels() {
		if (!InList()) throw new InvalidOperationException("Put the caret in a list first.");
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			ApplyListLevelsHere();
			if (LegalLists.Paragraphs(doc).Any(p => p.Legal)) LegalLists.Renumber(ed);    // (their indents)
		} finally {
			doc.EndEditCollection();
		}
	}

	void ApplyListLevelsHere() {
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// from the caret's paragraph back and forward over the neighbouring list paragraphs
		var first = doc.Range2(ed.SelectionStart, ed.SelectionStart);
		first.Expand(tomConstants.tomParagraph);
		while (first.Start > 0) {
			var before = doc.Range2(first.Start - 1, first.Start - 1);
			before.Expand(tomConstants.tomParagraph);
			if ((before.Para.ListType & 0xFFFF) == tomConstants.tomListNone) break;
			first = before;
		}
		var range = doc.Range2(first.Start, first.Start);
		range.Expand(tomConstants.tomParagraph);
		while ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) {
			var para = range.Para;
			ApplyLevel(para, para.ListLevelIndex);
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		OfficeConverters.FitListIndents(doc, first.Start, Math.Max(first.Start, range.Start));    // (their indents)
	}

	/// Starts the numbering again at start, from the list item with the caret (that item and the rest of its level).
	/// RichEdit numbers on through items that share a start value, so two lists that start at the same number are kept
	/// apart by an empty paragraph (at the top level; a nested list cannot be split that way).
	public void RestartNumbering(int start = 1) {
		if (InLegalList()) { LegalLists.Restart(ed, ed.SelectionStart, start); return; }
		var items = ListParagraphs(out int index);
		var para = ParaAt(items[index]);
		int type = para.ListType & 0xFFFF;
		if (type == tomConstants.tomListNone || type == tomConstants.tomListBullet) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		int level = para.ListLevelIndex;
		var before = PreviousAtLevel(items, index, level);
		if (before is int previous && ParaAt(previous).ListStart == start && ParaAt(previous).ListType == para.ListType) {
			if (level > 0) throw new InvalidOperationException($"The items above also start at {start}, and RichEdit numbers on through items with the same start; a nested list cannot be split. Choose another start number.");
			// split the list with an empty, normal paragraph
			int at = items[index];
			ed.TextDocument.Range2(at, at).Text = "\r";
			var separator = ParaAt(at);
			separator.ListType = tomConstants.tomListNone;
			separator.SetIndents(0, 0, 0);
			ed.Select(at + 1, 0);    // back in the item (now after the separator)
			items = ListParagraphs(out index);
		}
		foreach (int item in ItemsOfRun(items, index, level)) ParaAt(item).ListStart = start;
	}

	/// Continues the numbering of the previous list of the same kind, for the list item with the caret and the rest of
	/// its list. Empty paragraphs between the two lists are removed, so that they become one list.
	public void ContinueNumbering() {
		if (InLegalList()) { LegalLists.Continue(ed, ed.SelectionStart); return; }
		var items = ListParagraphs(out int index);
		var para = ParaAt(items[index]);
		int type = para.ListType & 0xFFFF;
		if (type == tomConstants.tomListNone || type == tomConstants.tomListBullet) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		int level = para.ListLevelIndex;
		// the first item of the caret's run
		int first = index;
		for (int k = index - 1; k >= 0; k--) {
			var other = ParaAt(items[k]);
			if (other.ListLevelIndex < level) break;
			if (other.ListLevelIndex > level) continue;
			if (other.ListStart != para.ListStart || other.ListType != para.ListType) break;
			first = k;
		}
		var doc = ed.TextDocument;
		if (PreviousAtLevel(items, first, level) is int inSameList) {
			// the item just before belongs to the same block of list paragraphs: take its start value, and the two runs join
			var previous = ParaAt(inSameList);
			if (previous.ListType != para.ListType) throw new InvalidOperationException("The list before has another numbering style; List options > List levels can make them the same.");
			foreach (int item in ItemsOfRun(items, first, level)) ParaAt(item).ListStart = previous.ListStart;
			return;
		}
		if (level > 0) throw new InvalidOperationException("There is no earlier item at this level in this list to continue from.");
		// an earlier top-level list, before non-list paragraphs
		int position = items[0];
		var range = doc.Range2(position, position);
		range.Expand(tomConstants.tomParagraph);
		bool onlyEmpty = true;
		int? found = null;
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			range = back;
			var p = range.Para;
			if ((p.ListType & 0xFFFF) == tomConstants.tomListNone) {
				if ((range.Text ?? "").Trim('\r', '\n', ' ', '\t') != "") onlyEmpty = false;
				continue;
			}
			if (p.ListLevelIndex == 0) { found = range.Start; break; }
		}
		if (found is not int previousItem || (ParaAt(previousItem).ListType & 0xFFFF) != type)
			throw new InvalidOperationException("There is no earlier numbered list of the same kind to continue.");
		var last = ParaAt(previousItem);
		if (last.ListType != para.ListType) throw new InvalidOperationException("The list before has another numbering style; List options > List levels can make them the same.");
		if (onlyEmpty) {
			// only empty paragraphs between: remove them, so that the two become one list
			var lastRange = doc.Range2(previousItem, previousItem);
			lastRange.Expand(tomConstants.tomParagraph);
			doc.Range2(lastRange.End, items[0]).Text = "";
			items = ListParagraphs(out index);
			// the runs join when they share the start value
			int startValue = last.ListStart;
			int firstOfRun = items.FindIndex(i => i >= lastRange.End);
			foreach (int item in ItemsOfRun(items, Math.Max(0, firstOfRun), 0)) ParaAt(item).ListStart = startValue;
			return;
		}
		int next = DisplayedNumber(previousItem) + 1;
		foreach (int item in ItemsOfRun(items, first, level)) ParaAt(item).ListStart = next;
	}

	ITextPara ParaAt(int position) => ed.TextDocument.Range2(position, position).Para;

	// The starts of the list paragraphs around the caret (a list ends at a paragraph that is not a list item), and the
	// index of the caret's paragraph among them.
	List<int> ListParagraphs(out int index) {
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var range = doc.Range2(ed.SelectionStart, ed.SelectionStart);
		range.Expand(tomConstants.tomParagraph);
		if ((range.Para.ListType & 0xFFFF) == tomConstants.tomListNone) throw new InvalidOperationException("Put the caret in a numbered list item first.");
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			if ((back.Para.ListType & 0xFFFF) == tomConstants.tomListNone) break;
			range = back;
		}
		var starts = new List<int>();
		while ((range.Para.ListType & 0xFFFF) != tomConstants.tomListNone) {
			starts.Add(range.Start);
			if (range.End >= storyEnd || range.End <= range.Start) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
		index = Math.Max(0, starts.FindLastIndex(start => start <= ed.SelectionStart));
		return starts;
	}

	// The item at the same level before items[index] in the same sublist (deeper items are skipped), or null.
	int? PreviousAtLevel(List<int> items, int index, int level) {
		for (int k = index - 1; k >= 0; k--) {
			int otherLevel = ParaAt(items[k]).ListLevelIndex;
			if (otherLevel < level) return null;
			if (otherLevel == level) return items[k];
		}
		return null;
	}

	// items[index] and the items after it at the same level, up to the end of their sublist.
	IEnumerable<int> ItemsOfRun(List<int> items, int index, int level) {
		for (int k = index; k < items.Count; k++) {
			int otherLevel = ParaAt(items[k]).ListLevelIndex;
			if (otherLevel < level) yield break;
			if (otherLevel == level) yield return items[k];
		}
	}

	// The number RichEdit shows for a list item: its start value plus the items before it with the same start value.
	int DisplayedNumber(int position) {
		var doc = ed.TextDocument;
		var para = ParaAt(position);
		int level = para.ListLevelIndex, startValue = para.ListStart, count = 0;
		var range = doc.Range2(position, position);
		range.Expand(tomConstants.tomParagraph);
		while (range.Start > 0) {
			var back = doc.Range2(range.Start - 1, range.Start - 1);
			back.Expand(tomConstants.tomParagraph);
			range = back;
			var p = range.Para;
			if ((p.ListType & 0xFFFF) == tomConstants.tomListNone || p.ListLevelIndex < level) break;
			if (p.ListLevelIndex > level) continue;
			if (p.ListStart != startValue || p.ListType != para.ListType) break;
			count++;
		}
		return startValue + count;
	}

	/// The document's list levels, one per line: level (1-9), style, punctuation, start number, and where the number
	/// and the text start (points from the margin, or "auto"), separated by tabs.
	public string ListLevels => string.Join("\n", form.Levels.Select((l, i) =>
		$"{i + 1}\t{l.Style}\t{l.Punctuation}\t{l.Start}\t{Points(l.NumberAt)}\t{Points(l.TextAt)}"));
	static string Points(float? value) => value is float v ? v.ToString("0.##", CultureInfo.InvariantCulture) : "auto";

	/// Whether Tab and Shift+Tab give a list item its new level's numbering style (List levels' "automatic numbering
	/// style per level"; 0: items keep theirs). Remembered.
	public bool AutomaticNumberingPerLevel {
		get => form.AutoLevelStyles;
		set { form.AutoLevelStyles = form.Settings.AutoLevelStyles = value; form.Settings.Save(); }
	}

	/// Sets a list level (1-9) of the document, and of new documents too unless forNewDocuments is 0: style Decimal,
	/// LowerAlpha, UpperAlpha, LowerRoman, UpperRoman or Bullet; punctuation ".", ")", "()" or ""; the start number;
	/// where the number and the text start (points from the margin; -1: automatic). New items and nested ones get the
	/// level; ApplyListLevels gives it to the list with the caret.
	public void SetListLevel(int level, string style, string punctuation = ".", int start = 1, float numberAt = -1, float textAt = -1, bool forNewDocuments = true) {
		if (level < 1 || level > 9) throw new ArgumentOutOfRangeException(nameof(level), "Levels are 1 to 9.");
		var parsed = Enum.Parse<ListStyle>(style, ignoreCase: true);
		if (!ListLevel.Punctuations.Contains(punctuation)) throw new ArgumentException("Punctuation must be \".\", \")\", \"()\" or \"\".");
		if (numberAt > 1584 || textAt > 1584 || numberAt is < 0 and not -1 || textAt is < 0 and not -1) throw new ArgumentException("numberAt and textAt are 0 to 1584 points, or -1 for automatic.");
		var value = new ListLevel { Style = parsed.ToString(), Punctuation = punctuation, Start = Math.Max(0, start),
			NumberAt = numberAt < 0 ? null : numberAt, TextAt = textAt < 0 ? null : textAt };
		form.Levels[level - 1] = value.Copy();
		form.LevelsChanged();
		if (!forNewDocuments) return;
		var levels = form.Settings.ListLevels.ToList();
		while (levels.Count < 9) levels.Add(ListLevel.Defaults()[levels.Count]);
		levels[level - 1] = value;
		form.Settings.ListLevels = levels;
		form.Settings.Save();
	}

	// One list paragraph gets the numbering of a level from the document's list levels.
	void ApplyLevel(ITextPara para, int level) {
		if (LegalLists.IsLegal(para.ListType)) return;    // legal items keep their numbering
		var levels = form.Levels;
		var scheme = level < levels.Count ? levels[level] : ListLevel.Defaults()[Math.Min(level, 8)];
		para.ListType = scheme.ListType;
		para.ListStart = scheme.Start;
	}

	void ForEachSelectedParagraph(Action<ITextPara> action) {
		int selStart = ed.SelectionStart, selEnd = selStart + ed.SelectionLength;
		var range = ed.TextDocument.Range2(selStart, selStart);
		range.Expand(tomConstants.tomParagraph);
		while (true) {
			action(range.Para);
			if (range.End >= selEnd || range.End <= range.Start || range.End >= ed.TextLength) break;
			range.SetRange(range.End, range.End);
			range.Expand(tomConstants.tomParagraph);
		}
	}

	// ---- Iteration 6: keys and macros

	/// Presses a key, as recorded macros do: Left, Right, Up, Down, Home, End, Backspace, Delete, Enter or Tab, optionally
	/// with Ctrl+ and/or Shift+ ("Ctrl+Shift+Left"). Shift extends the selection; Ctrl moves by words, paragraphs or the
	/// whole document (Home/End). count repeats the key.
	public void Key(string key, int count = 1) {
		var parts = key.Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
		bool ctrl = parts.Any(p => p.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || p.Equals("Control", StringComparison.OrdinalIgnoreCase));
		bool shift = parts.Any(p => p.Equals("Shift", StringComparison.OrdinalIgnoreCase));
		string name = parts.Length > 0 ? parts[^1].ToLowerInvariant() : "";
		var selection = ed.TextDocument.Selection2;
		int extend = shift ? tomConstants.tomExtend : tomConstants.tomMove;
		for (int i = 0; i < Math.Max(1, count); i++) {
			switch (name) {
				case "left": selection.MoveLeft(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1, extend); break;
				case "right": selection.MoveRight(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1, extend); break;
				case "up": selection.MoveUp(ctrl ? tomConstants.tomParagraph : tomConstants.tomLine, 1, extend); break;
				case "down": selection.MoveDown(ctrl ? tomConstants.tomParagraph : tomConstants.tomLine, 1, extend); break;
				case "home": selection.HomeKey(ctrl ? tomConstants.tomStory : tomConstants.tomLine, extend); break;
				case "end": selection.EndKey(ctrl ? tomConstants.tomStory : tomConstants.tomLine, extend); break;
				case "backspace": if (ctrl || !TouchesLegalList() || !LegalBackspace()) selection.Delete(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, selection.End > selection.Start ? 1 : -1); break;
				case "delete": if (ctrl || !TouchesLegalList() || !LegalDelete()) selection.Delete(ctrl ? tomConstants.tomWord : tomConstants.tomCharacter, 1); break;
				case "enter": if (!TouchesLegalList() || !LegalEnter()) selection.TypeText("\r"); break;
				case "tab": if (InTable()) NextCell(shift ? -1 : 1); else if (InList()) Nest(shift ? -1 : 1); else selection.TypeText("\t"); break;
				default: throw new ArgumentException($"Unknown key \"{key}\". Keys: Left, Right, Up, Down, Home, End, Backspace, Delete, Enter, Tab, with Ctrl+ and/or Shift+.");
			}
		}
		ed.ScrollToCaret();
	}

	/// Runs a macro from the rte_macros folder by name (without .ahk); it runs as a separate script.
	public void RunMacro(string name) => form.RunScript(form.MacroPath(name));

	// ---- Event hooks (iteration 56, rte_hooks.cs): macros run by AhkWord when something happens

	/// Hooks a macro (a name in the macros folder, or a full path; "" for none) to an event: Started, Opened,
	/// NewDocument, BeforeSave, Saved, StyleChanged, Printed or Closing (Tools > Event Hooks).
	public void SetEventHook(string eventName, string macro) {
		var e = EventHooks.Parse(eventName);
		var hooks = form.Settings.EventHooks.Where(h => !string.Equals(h.Key, e.ToString(), StringComparison.OrdinalIgnoreCase)).ToDictionary(h => h.Key, h => h.Value);
		if (macro.Trim() != "") {
			string name = macro.Trim();
			if (!File.Exists(EventHooks.PathOf(name, form.MacrosFolderPath)) && File.Exists(EventHooks.PathOf(name + ".ahk", form.MacrosFolderPath))) name += ".ahk";
			if (!File.Exists(EventHooks.PathOf(name, form.MacrosFolderPath))) throw new FileNotFoundException($"There is no macro \"{macro}\" in the macros folder ({form.MacrosFolderPath}).");
			hooks[e.ToString()] = name;
		}
		form.Settings.EventHooks = hooks;
		if (macro.Trim() == "") form.Settings.EventHooksOff = form.Settings.EventHooksOff.Where(o => !string.Equals(o, e.ToString(), StringComparison.OrdinalIgnoreCase)).ToList();
		form.Settings.Save();
	}

	/// Switches an event's hook on or off (off: its macro stays hooked, and doesn't run).
	public void SetEventHookOn(string eventName, bool on) {
		var e = EventHooks.Parse(eventName);
		var off = form.Settings.EventHooksOff.Where(o => !string.Equals(o, e.ToString(), StringComparison.OrdinalIgnoreCase)).ToList();
		if (!on) off.Add(e.ToString());
		form.Settings.EventHooksOff = off;
		form.Settings.Save();
	}

	/// Whether an event's hook is on (true also when it has no macro).
	public bool EventHookOn(string eventName) => EventHooks.IsOn(form.Settings.EventHooksOff, EventHooks.Parse(eventName));

	/// The macro hooked to an event ("" if none).
	public string EventHook(string eventName) => EventHooks.MacroFor(form.Settings.EventHooks, EventHooks.Parse(eventName)) ?? "";

	/// The hooks, one per line: "Event=macro" (" (off)" after one switched off).
	public string EventHookList => string.Join("\n", form.Settings.EventHooks.Select(h => h.Key + "=" + h.Value
		+ (EventHooks.IsOn(form.Settings.EventHooksOff, EventHooks.Parse(h.Key)) ? "" : " (off)")));

	/// Iteration 57: what the command palette lists (Ctrl+Shift+P), one per line: "name<tab>where<tab>shortcut".
	public string PaletteItems => string.Join("\n", form.PaletteItems().Select(i => i.Name + "\t" + i.Where + "\t" + i.Shortcut));

	/// The events a hook can run on, one per line.
	public string HookEventNames => string.Join("\n", Enum.GetNames<HookEvent>());

	/// The macros, one per line: file name and hotkey (if any), separated by a tab.
	public string Macros => string.Join("\n", form.MacroFiles().Select(f => f + "\t" + (form.HotkeyOf(f) is Keys k ? HotkeyBox.Describe(k) : "")));

	/// Sets the hotkey of a macro ("Ctrl+Alt+1", "F9", ...), or removes it with "". Works while the editor is active.
	public void SetMacroHotkey(string name, string keys) {
		string file = Path.GetFileName(form.MacroPath(name));
		var parsed = keys.Trim() == "" ? Keys.None : HotkeyBox.Parse(keys) ?? throw new ArgumentException($"\"{keys}\" is not a key combination (examples: Ctrl+Alt+1, F9, Ctrl+Shift+M).");
		form.SetMacroHotkey(file, parsed);
	}

	/// Starts recording a macro, as Tools > Record macro does.
	public void StartRecording() => form.StartRecording();

	/// Stops recording and returns the recorded AHK script ("" if nothing was recorded), without saving it.
	public string StopRecording() => form.StopRecording();

	/// Whether a macro is being recorded.
	public bool IsRecording => form.IsRecording;

	// ---- Iteration 5: spelling (the Windows spell checker, which also draws the red underlines)

	/// The spell checker's language, e.g. en-US ("" if the Windows spell checker is not available).
	public string SpellingLanguage => form.Spelling?.Language ?? "";

	/// Selects the next spelling error after the caret (wrapping round) and returns the word, or "" if there is none.
	public string NextSpellingError() {
		var spelling = RequireSpelling();
		int caret = ed.SelectionStart + ed.SelectionLength;
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		// From the caret's paragraph to the end, then from the start up to the caret.
		foreach (var (from, to) in new[] { (caret, storyEnd), (0, caret) }) {
			var para = doc.Range2(from, from);
			para.Expand(tomConstants.tomParagraph);
			while (para.Start < to) {
				int start = para.Start;
				foreach (var error in spelling.Check(para.Text ?? "", start))
					if (error.Start >= from && error.Start < to) {
						ed.Select(error.Start, error.Length);
						ed.ScrollToCaret();
						return error.Word;
					}
				if (para.End >= storyEnd || para.End <= start) break;
				para.SetRange(para.End, para.End);
				para.Expand(tomConstants.tomParagraph);
			}
		}
		return "";
	}

	/// All spelling errors, one per line: start, length and word, separated by tabs.
	public string SpellingErrors() {
		var spelling = RequireSpelling();
		var lines = new List<string>();
		ForEachParagraph((start, text) => {
			foreach (var error in spelling.Check(text, start)) lines.Add($"{error.Start}\t{error.Length}\t{error.Word}");
		});
		return string.Join("\n", lines);
	}

	/// Spelling suggestions for a word, one per line.
	public string Suggestions(string word) => string.Join("\n", RequireSpelling().Suggest(word));

	/// Adds a word to the Windows dictionary (shared with other programs), and refreshes the red underlines.
	public void AddToDictionary(string word) {
		RequireSpelling().Add(word.Trim());
		if (ed.EnableSpellCheck) {    // re-check, so the underline goes
			ed.EnableSpellCheck = false;
			ed.EnableSpellCheck = true;
		}
	}

	/// Replaces length characters at start with text, keeping the formatting of the first replaced character.
	public void ReplaceRange(int start, int length, string text) {
		var range = ed.TextDocument.Range2(start, start + length);
		range.Text = text;
		ed.Select(start + text.Length, 0);
	}

	internal Spelling.Error? SpellingErrorAt(int position) {
		if (form.Spelling == null) return null;
		var para = ed.TextDocument.Range2(position, position);
		para.Expand(tomConstants.tomParagraph);
		foreach (var error in form.Spelling.Check(para.Text ?? "", para.Start))
			if (position >= error.Start && position <= error.Start + error.Length) return error;
		return null;
	}

	internal void DeleteRepeatedWord(Spelling.Error error) {
		// Delete the word together with the space before it.
		int start = error.Start, length = error.Length;
		if (start > 0 && ed.TextDocument.Range2(start - 1, start).Text == " ") { start--; length++; }
		ReplaceRange(start, length, "");
	}

	Spelling RequireSpelling() => form.Spelling ?? throw new InvalidOperationException("The Windows spell checker is not available.");

	void ForEachParagraph(Action<int, string> action) {
		var doc = ed.TextDocument;
		int storyEnd = doc.Range2(0, 0).StoryLength;
		var para = doc.Range2(0, 0);
		para.Expand(tomConstants.tomParagraph);
		while (true) {
			int start = para.Start;
			action(start, para.Text ?? "");
			if (para.End >= storyEnd || para.End <= start) break;
			para.SetRange(para.End, para.End);
			para.Expand(tomConstants.tomParagraph);
		}
	}

	// ---- Iteration 5: outline

	/// The headings, one per line: level (1-6), start position and text, separated by tabs.
	public string Outline => string.Join("\n", form.CurrentOutline().Select(h => $"{h.Level}\t{h.Start}\t{h.Text}"));

	/// Goes to the n-th heading (1 = the first) and scrolls it to the top.
	public void GoToHeading(int number) {
		var headings = form.CurrentOutline();
		if (number < 1 || number > headings.Count) throw new ArgumentOutOfRangeException(nameof(number), $"There are {headings.Count} headings.");
		form.GoToPosition(headings[number - 1].Start);
	}

	/// Whether the outline panel is shown.
	public bool OutlineVisible {
		get => form.OutlineVisible;
		set => form.SetOutlineVisible(value);
	}

	/// <summary>The interface's language, as the setting Language has it: "en" (English) or a language AhkWord has a file
	/// for in its lang folder ("es": Spanish). Set, it is saved, and shows the next time AhkWord starts (Review > Language).</summary>
	public string Language {
		get => Lang.Chosen(AppFolders.ProgramFolder, form.Settings.Language);
		set {
			var found = Lang.Available(AppFolders.ProgramFolder).FirstOrDefault(l => l.Code.Equals(value?.Trim(), StringComparison.OrdinalIgnoreCase));
			if (found.Code == null) throw new ArgumentException($"No language \"{value}\": {string.Join(", ", Lang.Available(AppFolders.ProgramFolder).Select(l => l.Code))}.");
			form.Settings.Language = found.Code;
			form.Settings.Save();
		}
	}

	/// Whether the ruler (View > Show > Ruler) is shown. Remembered.
	public bool RulerVisible {
		get => form.Settings.ShowRuler;
		set => form.SetRulerVisible(value);
	}

	/// Changes the heading level (1-6, or 0 for normal text) of the paragraph at a position.
	public void SetHeadingLevel(int position, int level) {
		int selStart = ed.SelectionStart, selLength = ed.SelectionLength;
		ed.Select(position, 0);
		try {
			SetBlockStyle(level <= 0 ? BlockStyle.Paragraph : (BlockStyle)Math.Min(6, level));
		} finally {
			ed.Select(selStart, selLength);
		}
	}

	internal List<(int Level, int Start, string Text)> Headings() {
		var headings = new List<(int, int, string)>();
		var doc = ed.TextDocument;
		ForEachParagraph((start, text) => {
			text = text.Replace("￼", "").Trim();    // (a picture in a heading is an object character)
			// table cells are not headings, though a header row's bold text has a heading's font (cells end in U+0007)
			if (text == "" || text.IndexOfAny(new[] { '\u0007', '\uFFF9', '\uFFFB' }) >= 0) return;
			int level = StyleEngine.LevelAt(doc, form.Styles, start);    // (by the paragraph's style, not its font)
			if (level is >= 1 and <= 6) headings.Add((level, start, text));
		});
		return headings;
	}

	// The library's own rule for recognising block styles (internal), or the same sizes if it cannot be reached.
	static readonly MethodInfo? blockStyleFromFont = typeof(RichEditBox).Assembly
		.GetType("EllipticBit.RichEditorNET.Formatting.BlockStyleHelper")?.GetMethod("GetBlockStyleFromFont", BindingFlags.NonPublic | BindingFlags.Static);

	static BlockStyle? BlockStyleOf(float size, bool bold, string? name) {
		if (blockStyleFromFont != null) return (BlockStyle)blockStyleFromFont.Invoke(null, new object?[] { size, bold, name })!;
		int level = Array.IndexOf(new[] { 0f, 24f, 18f, 14f, 12f, 10f, 8f }, size);
		return bold && level > 0 ? (BlockStyle)level : null;
	}

	// ---- The document's header and footer (fields: {title} {file} {page} {pages} {date} {time}); printed, and saved
	// with the document (iteration 13)

	void BandsChanged() { ed.Modified = true; form.BandsChangedNow(); }

	/// Sets the document's header (left, center and right parts) and shows it. A part can have several lines ("`n"
	/// between them).
	public void SetHeader(string left = "", string center = "", string right = "") {
		form.Bands.Header = new[] { left, center, right };
		form.Bands.ShowHeader = true;
		BandsChanged();
	}

	/// Sets the document's footer (left, center and right parts) and shows it.
	public void SetFooter(string left = "", string center = "", string right = "") {
		form.Bands.Footer = new[] { left, center, right };
		form.Bands.ShowFooter = true;
		BandsChanged();
	}

	/// Whether the document has a header (printed, and in its files).
	public bool ShowHeader {
		get => form.Bands.ShowHeader;
		set { form.Bands.ShowHeader = value; BandsChanged(); }
	}

	/// Whether the document has a footer.
	public bool ShowFooter {
		get => form.Bands.ShowFooter;
		set { form.Bands.ShowFooter = value; BandsChanged(); }
	}

	/// The header's left, center and right parts, one per line.
	public string HeaderParts => string.Join("\n", PageBands.Parts(form.Bands.Header).Select(BandText.Plain));

	/// The footer's left, center and right parts, one per line.
	public string FooterParts => string.Join("\n", PageBands.Parts(form.Bands.Footer).Select(BandText.Plain));

	/// Whether the first page has a header and footer of its own (SetFirstHeader, SetFirstFooter). Empty
	/// parts leave the first page without one, as on a letter's or a pleading's first page.
	public bool DifferentFirstPage {
		get => form.Bands.DifferentFirst;
		set { form.Bands.DifferentFirst = value; BandsChanged(); }
	}

	/// Sets the first page's header (left, center and right parts): the first page gets its own (DifferentFirstPage),
	/// shown (ShowFirstHeader). (Iteration 55) The other pages are as they were: ShowHeader off, only the first page has
	/// one.
	public void SetFirstHeader(string left = "", string center = "", string right = "") {
		form.Bands.FirstHeader = new[] { left, center, right };
		form.Bands.FirstFooter = PageBands.Parts(form.Bands.FirstFooter);
		form.Bands.DifferentFirst = true;
		form.Bands.ShowFirstHeader = true;
		BandsChanged();
	}

	/// Sets the first page's footer (left, center and right parts), as SetFirstHeader.
	public void SetFirstFooter(string left = "", string center = "", string right = "") {
		form.Bands.FirstFooter = new[] { left, center, right };
		form.Bands.FirstHeader = PageBands.Parts(form.Bands.FirstHeader);
		form.Bands.DifferentFirst = true;
		form.Bands.ShowFirstFooter = true;
		BandsChanged();
	}

	/// Iteration 55: whether the first page has a header (with DifferentFirstPage; the other pages: ShowHeader).
	public bool ShowFirstHeader {
		get => form.Bands.ShowHeaderOn(1);
		set { form.Bands.ShowFirstHeader = value; BandsChanged(); }
	}

	/// Whether the first page has a footer (with DifferentFirstPage; the other pages: ShowFooter).
	public bool ShowFirstFooter {
		get => form.Bands.ShowFooterOn(1);
		set { form.Bands.ShowFirstFooter = value; BandsChanged(); }
	}

	/// The first page's header font, as SetHeaderFont (with DifferentFirstPage; an empty name: plain).
	public void SetFirstHeaderFont(string name = "", double size = 10, bool bold = false, bool italic = false, bool underline = false) {
		form.Bands.FirstHeaderFont = name.Trim() == "" ? BandFont.Plain : new BandFont { Name = name.Trim(), Size = (float)Math.Clamp(size, 4, 72), Bold = bold, Italic = italic, Underline = underline };
		BandsChanged();
	}

	/// The first page's footer font, as SetHeaderFont.
	public void SetFirstFooterFont(string name = "", double size = 10, bool bold = false, bool italic = false, bool underline = false) {
		form.Bands.FirstFooterFont = name.Trim() == "" ? BandFont.Plain : new BandFont { Name = name.Trim(), Size = (float)Math.Clamp(size, 4, 72), Bold = bold, Italic = italic, Underline = underline };
		BandsChanged();
	}

	/// A line below the first page's header (with DifferentFirstPage).
	public bool FirstHeaderLine {
		get => form.Bands.HeaderLineOn(1);
		set { form.Bands.FirstHeaderLine = value; BandsChanged(); }
	}

	/// A line above the first page's footer.
	public bool FirstFooterLine {
		get => form.Bands.FooterLineOn(1);
		set { form.Bands.FirstFooterLine = value; BandsChanged(); }
	}

	/// The first page's header parts, one per line (empty lines if it has none of its own).
	public string FirstHeaderParts => string.Join("\n", PageBands.Parts(form.Bands.FirstHeader));

	/// The first page's footer parts, one per line.
	public string FirstFooterParts => string.Join("\n", PageBands.Parts(form.Bands.FirstFooter));

	/// Whether the header and footer are shown above and below the text (View > Header and footer).
	public bool HeaderFooterVisible {
		get => form.Settings.ShowBands;
		set { form.Settings.ShowBands = value; form.Settings.Save(); }
	}

	/// Iteration 55: the header's font: its name, size in points, bold, italic and underline (all of its parts, the first
	/// page's too); printed, on Show Pages' sheets, and in Word and OpenDocument files. An empty name: Segoe UI 8 pt in
	/// grey, as before fonts could be chosen.
	public void SetHeaderFont(string name = "", double size = 10, bool bold = false, bool italic = false, bool underline = false) {
		form.Bands.HeaderFont = name.Trim() == "" ? null : new BandFont { Name = name.Trim(), Size = (float)Math.Clamp(size, 4, 72), Bold = bold, Italic = italic, Underline = underline };
		BandsChanged();
	}

	/// The footer's font, as SetHeaderFont.
	public void SetFooterFont(string name = "", double size = 10, bool bold = false, bool italic = false, bool underline = false) {
		form.Bands.FooterFont = name.Trim() == "" ? null : new BandFont { Name = name.Trim(), Size = (float)Math.Clamp(size, 4, 72), Bold = bold, Italic = italic, Underline = underline };
		BandsChanged();
	}

	/// Iteration 82: the document's page setup, in inches: margins, orientation and paper (its width and height upright; 0:
	/// as it is), as Page Setup sets it. Saved with the document.
	public void SetPageSetup(double left, double right, double top, double bottom, bool landscape = false, double paperWidth = 0, double paperHeight = 0) {
		static int Hundredths(double inches) => (int)Math.Round(Math.Clamp(inches, 0, 48) * 100);
		var paper = paperWidth > 0 && paperHeight > 0 ? Papers.Upright(Hundredths(paperWidth), Hundredths(paperHeight), twips: false)! : form.PagePaper;
		int[] margins = { Hundredths(left), Hundredths(right), Hundredths(top), Hundredths(bottom) };
		var (w, h) = landscape ? (paper[1], paper[0]) : (paper[0], paper[1]);
		if (w - margins[0] - margins[1] < 50 || h - margins[2] - margins[3] < 50) throw new ArgumentException("The margins leave less than half an inch for the text.");
		form.SetPageSetup(margins, landscape, paper);
	}

	/// Iteration 82: the document's margins, in inches.
	public double LeftMargin => form.PageMargins[0] / 100.0;
	public double RightMargin => form.PageMargins[1] / 100.0;
	public double TopMargin => form.PageMargins[2] / 100.0;
	public double BottomMargin => form.PageMargins[3] / 100.0;
	/// Whether the document's pages are landscape.
	public bool Landscape => form.PageLandscape;
	/// The document's paper, in inches, upright (Letter: 8.5 by 11).
	public double PaperWidth => form.PagePaper[0] / 100.0;
	public double PaperHeight => form.PagePaper[1] / 100.0;

	/// Iteration 55: from the page's top edge to the header, in inches (Word's "Header from top"; 0.5 by default). A
	/// header reaching past the top margin pushes the text down, as in Word.
	public double HeaderFromTop {
		get => form.Bands.HeaderFromEdge / 100.0;
		set { int d = (int)Math.Round(Math.Clamp(value, 0, 5) * 100); form.Bands.HeaderDistance = d == PageBands.DefaultDistance ? null : d; BandsChanged(); }
	}

	/// From the page's bottom edge to the footer's last line, in inches (Word's "Footer from bottom"; 0.5 by default).
	public double FooterFromBottom {
		get => form.Bands.FooterFromEdge / 100.0;
		set { int d = (int)Math.Round(Math.Clamp(value, 0, 5) * 100); form.Bands.FooterDistance = d == PageBands.DefaultDistance ? null : d; BandsChanged(); }
	}

	/// Iteration 55: a line below the header (Word's border below its paragraphs).
	public bool HeaderLine {
		get => form.Bands.HeaderLine;
		set { form.Bands.HeaderLine = value; BandsChanged(); }
	}

	/// A line above the footer.
	public bool FooterLine {
		get => form.Bands.FooterLine;
		set { form.Bands.FooterLine = value; BandsChanged(); }
	}

	/// Iteration 55: a horizontal line across the text, in a paragraph of its own (Insert > Horizontal Line): a picture of
	/// a line, as Word's, so that it is printed and kept in every format; Markdown gets a rule (---) and HTML an <hr>.
	public void InsertHorizontalLine() {
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		float width = Math.Max(72, PageTextWidth() / 20f - 2);
		int pixels = (int)Math.Round(width * 96 / 72), height = 8;
		using var bitmap = new Bitmap(pixels, height);
		using (var g = Graphics.FromImage(bitmap)) {
			g.Clear(Color.White);    // (not transparent: RichEdit draws that black)
			using var brush = new SolidBrush(Color.FromArgb(128, 128, 128));
			g.FillRectangle(brush, 0, 3, pixels, 2);
		}
		using var png = new MemoryStream();
		bitmap.Save(png, System.Drawing.Imaging.ImageFormat.Png);
		doc.BeginEditCollection();
		try {
			selection.Text = "";
			var paragraph = doc.Range2(selection.Start, selection.Start);
			paragraph.Expand(tomConstants.tomParagraph);
			if (selection.Start > paragraph.Start) { selection.Text = "\r"; selection.Collapse(tomConstants.tomCollapseEnd); }
			int oldWidth = ed.DefaultImageWidth, oldHeight = ed.DefaultImageHeight;
			try {
				ed.DefaultImageWidth = pixels;
				ed.DefaultImageHeight = height;
				png.Position = 0;
				ed.InsertImage(png, MarkdownConverter.RuleAlt);
			} finally {
				ed.DefaultImageWidth = oldWidth;
				ed.DefaultImageHeight = oldHeight;
			}
			int after = doc.Selection2.Start;
			string next = doc.Range2(after, after + 1).Text ?? "";
			if (next != "\r") { doc.Range2(after, after).Text = "\r"; }
			doc.Selection2.SetRange(after + 1, after + 1);
		} finally {
			doc.EndEditCollection();
		}
	}

	/// The header's font ("Times New Roman 10 pt, bold"), empty without one of its own.
	public string HeaderFont => form.Bands.HeaderFont?.Describe() ?? "";

	/// The footer's font, as HeaderFont.
	public string FooterFont => form.Bands.FooterFont?.Describe() ?? "";

	/// Makes the document's header and footer the default for new documents (and documents without their own).
	public void MakeHeaderFooterDefault() => form.MakeDefault(form.Bands);

	// ---- Cross-references to numbered paragraphs (iteration 50)

	/// The numbered paragraphs a cross-reference can refer to (legal items, numbered styles' paragraphs, numbered
	/// lists), one a line: the number, a tab, the full number (with the levels above: "3(b)"), a tab, and the start of
	/// the text.
	public string CrossReferenceTargets => string.Join("\n", CrossRefs.Targets(ed.TextDocument).Select(t => t.Number + "\t" + t.Full + "\t" + t.Text));

	/// Inserts a cross-reference at the caret (in place of the selection) to numbered paragraph index (1: the first, as
	/// CrossReferenceTargets lists them), showing its number (show "number": "(b)") or its full number ("full": "3(b)"),
	/// without a trailing period. Its number follows the paragraph's when the paragraphs are renumbered; Word and
	/// LibreOffice files get it as a REF field. One undo step.
	public void InsertCrossReference(int index, string show = "number") {
		var doc = ed.TextDocument;
		var targets = CrossRefs.Targets(doc);
		if (index < 1 || index > targets.Count) throw new ArgumentException($"There are {targets.Count} numbered paragraphs (CrossReferenceTargets): {index} isn't one of them.");
		string how = show.Trim().ToLowerInvariant() switch {
			"number" or "r" or "" => "r", "full" or "w" => "w",
			_ => throw new ArgumentException("show is \"number\" or \"full\".")
		};
		doc.BeginEditCollection();
		try {
			var selection = doc.Selection2;
			int end = CrossRefs.Insert(doc, doc.Range2(selection.Start, selection.End), targets[index - 1], how);
			selection.SetRange(end, end);
			selection.Font.Hidden = tomConstants.tomFalse;    // (typing on after it: not in the marker's hidden format)
		} finally {
			doc.EndEditCollection();
		}
	}

	/// Brings every cross-reference's number up to date (done by itself as the paragraphs are renumbered, and before
	/// saving, printing and exporting; F9). Returns how many changed.
	public int UpdateCrossReferences() => form.UpdateReferences(tables: true);

	// ---- Tables of contents (iteration 52)

	/// Inserts a table of contents at the caret's paragraph (in place of it if it is empty, else before it): the headings
	/// of levels 1 to levels (1-6), each with the page it is printed on, under a title ("": none). pageNumbers 0: no
	/// page numbers; rightAlign 0: each just after its heading's text, else at the right margin after a leader (dots,
	/// dashes, underline or none). With the caret in a table of contents: that table made again with these options.
	/// One undo step.
	public void InsertTableOfContents(int levels = 3, string title = "Contents", bool pageNumbers = true, bool rightAlign = true, string leader = "dots") {
		if (levels is < 1 or > 6) throw new ArgumentException("levels is 1 to 6: a table of contents takes the headings of levels 1 to that.");
		leader = (leader ?? "dots").Trim().ToLowerInvariant();
		if (!Toc.Options.Leaders.Contains(leader)) throw new ArgumentException("The leader is dots, dashes, underline or none.");
		form.InsertContents(new Toc.Options(levels, pageNumbers, rightAlign, leader), title ?? "");
	}

	/// Makes every table of contents again from the headings, with their pages as printed (pagesOnly = 1: only their page
	/// numbers brought up to date, as before saving and printing). Returns how many tables were made again (pagesOnly: how
	/// many page numbers changed). One undo step.
	public int UpdateTableOfContents(bool pagesOnly = false) => form.UpdateContents(pagesOnly);

	/// The (first) table of contents' entries, one a line: the level, a tab, the heading's text, a tab, the page.
	public string TableOfContents => Toc.Listing(ed.TextDocument, form.Styles);

	// ---- Footnotes (iteration 51)

	/// Inserts a footnote at the caret (after the selection) with its note (its paragraphs apart by "`n"; <b> <i> <u> tags
	/// for bold, italic and underline): a superscript number in the text, the note at the foot of the page when printed (and on
	/// Show Pages' sheets). The footnotes are numbered 1, 2, 3 ... in order. One undo step.
	public void InsertFootnote(string note) {
		note = Footnotes.FromTags(note);
		if (string.IsNullOrWhiteSpace(Footnotes.Plain(note))) throw new ArgumentException("A footnote needs a note.");
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			var selection = doc.Selection2;
			int end = Footnotes.Insert(doc, doc.Range2(selection.End, selection.End), note);
			selection.SetRange(end, end);
			selection.Font.Hidden = tomConstants.tomFalse;    // (typing on after it: not in the note's hidden format, nor superscript)
			selection.Font.Superscript = tomConstants.tomFalse;
		} finally {
			doc.EndEditCollection();
		}
	}

	/// The document's footnotes, one a line: the number, a tab and the note (its formatting as <b> <i> <u> tags; its
	/// paragraphs apart by "`r", so that each note stays on its line). FootnoteText gives one note.
	public string FootnoteList => string.Join("\n", Footnotes.All(Story()).Select(f => f.Number + "\t" + Footnotes.Paragraphs(Footnotes.ToTags(f.Text), "\r")));

	/// Footnote number's (1: the first) note: its formatting as <b> <i> <u> tags, its paragraphs apart by "`n".
	public string FootnoteText(int number) {
		var all = Footnotes.All(Story());
		if (number < 1 || number > all.Count) throw new ArgumentException($"There are {all.Count} footnotes: {number} isn't one of them.");
		return Footnotes.Paragraphs(Footnotes.ToTags(all[number - 1].Text));
	}

	/// The number of footnotes.
	public int FootnoteCount => Footnotes.All(Story()).Count;

	/// The size of the footnotes' notes (points): the document's Footnote Text style's (set: made if it has none, as
	/// a Word document opened has it), else three quarters of the text's (12 points: 9).
	public float FootnoteSize {
		get => form.NoteSizeNow();
		set {
			if (value is < 4 or > 72) throw new ArgumentException("FootnoteSize is 4 to 72 points.");
			form.SetNoteStyle(value);
			ed.Modified = true;
		}
	}

	/// Gives footnote number (1: the first) another note (as InsertFootnote's: paragraphs apart by "`n", <b> <i> <u> tags
	/// for its formatting). One undo step.
	public void SetFootnote(int number, string note) {
		var all = Footnotes.All(Story());
		if (number < 1 || number > all.Count) throw new ArgumentException($"There are {all.Count} footnotes: {number} isn't one of them.");
		note = Footnotes.FromTags(note);
		if (string.IsNullOrWhiteSpace(Footnotes.Plain(note))) throw new ArgumentException("A footnote needs a note (DeleteFootnote takes one out).");
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try { Footnotes.SetText(doc, all[number - 1], note); }
		finally { doc.EndEditCollection(); }
	}

	/// Takes footnote number (1: the first) out, its number and its note; the others are numbered again. One undo step.
	public void DeleteFootnote(int number) {
		var all = Footnotes.All(Story());
		if (number < 1 || number > all.Count) throw new ArgumentException($"There are {all.Count} footnotes: {number} isn't one of them.");
		var doc = ed.TextDocument;
		doc.BeginEditCollection();
		try {
			doc.Range2(all[number - 1].Start, all[number - 1].End).Text = "";
			Footnotes.Update(doc);
		} finally {
			doc.EndEditCollection();
		}
	}

	string Story() => ed.TextDocument.Range2(0, ed.TextDocument.Range2(0, 0).StoryLength).Text ?? "";

	// ---- Window

	/// Shows a message in the status bar.
	public void SetStatus(string text) => form.SetStatus(text);

	/// Brings the editor window to the front.
	public void Activate() {
		if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
		form.Activate();
		ed.Focus();
	}

	/// Closes the editor after replying. With discardChanges = 0, the user is asked about unsaved changes.
	public void Close(bool discardChanges = false) => form.CloseLater(discardChanges);

	/// Lists everything AHK scripts can call.
	public string Help() => Automation.Help();

	// ---- Helpers

	// The selection as TOM has it (WinForms' SelectionLength is short across table rows).
	internal ITextRange2 SelectionRange() {
		var selection = ed.TextDocument.Selection2;
		return ed.TextDocument.Range2(selection.Start, selection.End);
	}

	/// The selection's character format; with nothing selected, that of what is typed next (TOM's insertion point,
	/// which a font change at the caret sets), as Word's font box shows it.
	internal ITextFont SelectionFont() {
		var selection = ed.TextDocument.Selection2;
		return selection.Start == selection.End ? selection.Font : SelectionRange().Font;
	}

	/// Sets the font of the selected text; with nothing selected, of what is typed next, and on an empty paragraph of its
	/// paragraph mark too (the empty line keeps the font), as in Word.
	public void SetFontName(string name) {
		if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A font name is needed.");
		var selection = ed.TextDocument.Selection2;
		if (selection.Start != selection.End) { ed.SetFontName(name.Trim()); return; }
		AtCaret(font => font.Name = name.Trim());
	}

	// A character format with nothing selected: for what is typed at the caret (TOM's insertion point) and, on an empty
	// paragraph (or table cell), for its mark too, so the empty line keeps the format, and shows it when the caret comes
	// back, as in Word. (RichEditorNET's font commands change selected text only.) The mark's change is an undo step.
	void AtCaret(Action<ITextFont> set) {
		var doc = ed.TextDocument;
		var selection = doc.Selection2;
		int at = selection.Start;
		var paragraph = doc.Range2(at, at);
		paragraph.Expand(4);    // tomParagraph
		if (paragraph.End - paragraph.Start == 1 && (paragraph.Text ?? "") is "\r" or "\u0007")
			set(doc.Range2(paragraph.Start, paragraph.End).Font);
		set(selection.Font);    // (last: a change elsewhere can reset the insertion point's format)
	}

	static readonly MethodInfo? getBlockStyle = typeof(RichEditBox).GetMethod("GetBlockStyle", BindingFlags.NonPublic | BindingFlags.Instance);
	internal BlockStyle? GetBlockStyle() => getBlockStyle?.Invoke(ed, null) as BlockStyle?;

	T WithGithubMarkdown<T>(Func<T> action) {
		bool common = ed.EnableCommonMarkdown, github = ed.EnableGithubMarkdown;
		ed.EnableGithubMarkdown = true;
		try {
			return action();
		} finally {
			ed.EnableCommonMarkdown = common;
			ed.EnableGithubMarkdown = github;
		}
	}

	static string HtmlDocument(string body, string title, string css = "", string meta = "", string end = "") =>
		"<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<meta charset=\"utf-8\">\r\n" + meta + "<title>" + WebUtility.HtmlEncode(title) + "</title>\r\n"
		+ HtmlLists.StyleElement(css) + "</head>\r\n<body>\r\n"
		+ body + "\r\n</body>\r\n</html>\r\n" + end;

	static string BodyOf(string html) {
		var match = Regex.Match(html, @"<body[^>]*>(.*)</body>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
		return (match.Success ? match.Groups[1].Value : html).Trim();
	}
}

// Calls EditorApi / RichEditBox members by name with text arguments, and formats the results as text.
static class Automation
{
	// Inherited RichTextBox / TextBoxBase members that are useful to scripts (members declared on RichEditBox are always available).
	static readonly HashSet<string> EditorMembers = new(StringComparer.OrdinalIgnoreCase) {
		"Text", "TextLength", "SelectedText", "SelectionStart", "SelectionLength", "SelectAll", "Undo", "Redo", "CanUndo", "CanRedo",
		"Cut", "Copy", "Paste", "Clear", "AppendText", "ScrollToCaret", "Modified", "ReadOnly", "WordWrap", "ZoomFactor",
		"GetLineFromCharIndex", "GetFirstCharIndexFromLine", "Lines"
	};

	const BindingFlags Flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase;

	public static string Dispatch(EditorApi api, RichEditBox editor, string[] request) {
		try {
			if (request.Length < 2) throw new ArgumentException("A request needs a verb (call, get or set) and a member name.");
			string verb = request[0], name = request[1];
			string[] args = request[2..];
			var members = Members(api, editor, name).ToList();
			if (members.Count == 0) throw new MissingMemberException($"Unknown member \"{name}\". Rte.Help() lists the available members.");
			object? result = verb switch {
				"call" => Call(members, name, args),
				"get" => Get(members, name, args),
				"set" => Set(members, name, args),
				_ => throw new ArgumentException($"Unknown verb \"{verb}\".")
			};
			return "0" + Format(result);
		} catch (Exception ex) {
			while (ex is TargetInvocationException { InnerException: { } inner }) ex = inner;
			return "1" + ex.Message;
		}
	}

	static IEnumerable<(object Target, MemberInfo Member)> Members(EditorApi api, RichEditBox editor, string name) {
		foreach (var member in typeof(EditorApi).GetMember(name, Flags))
			yield return (api, member);
		foreach (var member in typeof(RichEditBox).GetMember(name, Flags))
			if (member.DeclaringType == typeof(RichEditBox) || EditorMembers.Contains(member.Name))
				yield return (editor, member);
	}

	static object? Call(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		Exception? conversionError = null;
		foreach (var (target, member) in members) {
			if (member is MethodInfo method && !method.IsSpecialName) {
				var parameters = method.GetParameters();
				int required = parameters.Count(p => !p.HasDefaultValue);
				if (args.Length < required || args.Length > parameters.Length) continue;
				object?[] values;
				try {
					values = parameters.Select((p, i) => i < args.Length ? Convert(args[i], p.ParameterType) : p.DefaultValue).ToArray();
				} catch (Exception ex) when (ex is FormatException or ArgumentException or NotSupportedException or OverflowException) {
					conversionError = ex;
					continue;
				}
				return method.Invoke(target, values);
			}
			if (member is PropertyInfo property)    // Rte.Text() / Rte.Text("value") also work
				return args.Length == 0 ? property.GetValue(target) : SetProperty(target, property, args[0]);
		}
		if (conversionError != null) throw conversionError;
		var counts = members.Select(m => m.Member).OfType<MethodInfo>().Select(m => m.GetParameters().Length).Distinct();
		throw new ArgumentException($"{name} does not take {args.Length} argument{(args.Length == 1 ? "" : "s")} (it takes {string.Join(" or ", counts)}).");
	}

	static object? Get(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		foreach (var (target, member) in members) {
			if (member is PropertyInfo { CanRead: true } property && property.GetIndexParameters().Length == 0) return property.GetValue(target);
			if (member is MethodInfo method && !method.IsSpecialName && method.GetParameters().All(p => p.HasDefaultValue))
				return Call(members, name, args);    // Rte.WordCount without parentheses
		}
		throw new MissingMemberException($"{name} is not a readable property.");
	}

	static object? Set(List<(object Target, MemberInfo Member)> members, string name, string[] args) {
		if (args.Length != 1) throw new ArgumentException("A property assignment needs exactly one value.");
		foreach (var (target, member) in members)
			if (member is PropertyInfo { CanWrite: true } property) return SetProperty(target, property, args[0]);
		throw new MissingMemberException($"{name} is not a writable property.");
	}

	static object? SetProperty(object target, PropertyInfo property, string value) {
		property.SetValue(target, Convert(value, property.PropertyType));
		return null;
	}

	static object? Convert(string text, Type type) {
		if (Nullable.GetUnderlyingType(type) is { } underlying) return text == "" ? null : Convert(text, underlying);
		if (type == typeof(string) || type == typeof(object)) return text;
		if (type == typeof(bool)) return text.Trim().ToLowerInvariant() switch {
			"1" or "true" or "yes" or "on" => true,
			"0" or "false" or "no" or "off" or "" => false,
			_ => throw new FormatException($"\"{text}\" is not a boolean (use 1 or 0).")
		};
		if (type.IsEnum) return Enum.Parse(type, text.Trim(), ignoreCase: true);
		if (type == typeof(Color)) {
			string t = text.Trim();
			if (t.StartsWith("0x", StringComparison.OrdinalIgnoreCase)) return Color.FromArgb(255, Color.FromArgb(int.Parse(t[2..], NumberStyles.HexNumber)));
			return ColorTranslator.FromHtml(t);
		}
		if (typeof(IConvertible).IsAssignableFrom(type)) return System.Convert.ChangeType(text.Trim(), type, CultureInfo.InvariantCulture);
		throw new NotSupportedException($"A {type.Name} cannot be passed from AHK.");
	}

	static bool CanConvert(Type type) {
		type = Nullable.GetUnderlyingType(type) ?? type;
		return type == typeof(string) || type == typeof(object) || type.IsEnum || type == typeof(Color) || typeof(IConvertible).IsAssignableFrom(type);
	}

	static string Format(object? value) => value switch {
		null => "",
		string s => s,
		bool b => b ? "1" : "0",
		Color c => $"#{c.R:X2}{c.G:X2}{c.B:X2}",
		Enum e => e.ToString(),
		IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
		IEnumerable items => string.Join("\n", items.Cast<object?>().Select(Format)),
		_ => value.ToString() ?? ""
	};

	public static string Help() {
		var lines = new List<string> {
			"Call any of these from AHK after #Include \"RteClient2.ahk\":",
			"  Rte.Name(args...)  calls a method      Rte.Name  reads a property      Rte.Name := value  sets one",
			"Arguments: text, numbers, 1/0 for booleans, enum names (Heading1, LowerAlpha, Wavy...), colors as #RRGGBB or names.",
			"RteClient2.ahk also has:",
			"  Rte.Post(Name, args...)  queues a call and returns at once (errors: Rte.LastAsyncError / the AsyncError event)",
			"  Rte.On(Event, Callback) / Rte.Off(Event [, Callback])  calls Callback(e) when an event happens",
			"Events and their e properties (all have e.Name):",
			"  TextChanged (e.TextLength)   SelectionChanged (e.Start, e.Length)   ModifiedChanged (e.Modified)",
			"  Saved   Opened (also New)   Closed   AsyncError",
			""
		};
		void Section(string title, Type type, Func<MemberInfo, bool> include) {
			lines.Add(title);
			var members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance)
				.Where(m => m is MethodInfo { IsSpecialName: false } || m is PropertyInfo)
				.Where(m => m.DeclaringType != typeof(object) && include(m))
				.Where(m => m is not MethodInfo method || method.GetParameters().All(p => CanConvert(p.ParameterType)))    // e.g. not InsertImage(Stream)
				.GroupBy(m => Describe(m)).Select(g => g.Key).OrderBy(s => s, StringComparer.OrdinalIgnoreCase);
			lines.AddRange(members.Select(m => "  " + m));
			lines.Add("");
		}
		Section("AhkWord (EditorApi):", typeof(EditorApi), m => m.DeclaringType == typeof(EditorApi));
		var api = typeof(EditorApi).GetMembers(BindingFlags.Public | BindingFlags.Instance).Select(m => m.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
		Section("Editor (RichEditBox):", typeof(RichEditBox), m => (m.DeclaringType == typeof(RichEditBox) || EditorMembers.Contains(m.Name)) && !api.Contains(m.Name)
			&& !(m is PropertyInfo p && p.PropertyType.Namespace == typeof(ITextDocument2).Namespace));
		return string.Join("\r\n", lines);
	}

	static string Describe(MemberInfo member) => member switch {
		MethodInfo m => $"{m.Name}({string.Join(", ", m.GetParameters().Select(p => $"{TypeName(p.ParameterType)} {p.Name}{(p.HasDefaultValue ? " = " + Format(p.DefaultValue) : "")}"))})"
			+ (m.ReturnType == typeof(void) ? "" : " -> " + TypeName(m.ReturnType)),
		PropertyInfo p => $"{p.Name} : {TypeName(p.PropertyType)}{(p.CanWrite && p.SetMethod!.IsPublic ? "" : " (read-only)")}",
		_ => member.Name
	};

	static string TypeName(Type type) => Nullable.GetUnderlyingType(type) is { } u ? TypeName(u) + "?" : type switch {
		_ when type == typeof(string) => "text",
		_ when type == typeof(bool) => "1/0",
		_ when type == typeof(int) || type == typeof(float) || type == typeof(double) => "number",
		_ when type.IsEnum => $"{type.Name} ({string.Join("|", Enum.GetNames(type))})",
		_ when type == typeof(Color) => "color",
		_ => type.Name
	};
}

// Editor events for AHK scripts. Subscribers are windows (an AHK script's A_ScriptHwnd). Each event is posted to them as
// the registered message "RTE_WRITER_EVENT" (wParam = event number, lParam = event data), so the editor never waits for
// a script; windows that no longer exist are dropped. The numbers are part of the protocol (RteClient2.ahk).
enum EditorEvent { TextChanged = 1, SelectionChanged, ModifiedChanged, Saved, Opened, Closed, AsyncError }

sealed class EventHub
{
	public static readonly int Message = RegisterWindowMessage("RTE_WRITER_EVENT");
	readonly Dictionary<EditorEvent, HashSet<IntPtr>> subscribers = new();

	public void Subscribe(EditorEvent e, IntPtr window) {
		if (!IsWindow(window)) throw new ArgumentException("That window does not exist.");
		if (!subscribers.TryGetValue(e, out var windows)) subscribers[e] = windows = new HashSet<IntPtr>();
		windows.Add(window);
	}

	public void Unsubscribe(EditorEvent? e, IntPtr window) {
		foreach (var (key, windows) in subscribers)
			if (e == null || key == e) windows.Remove(window);
	}

	public void Raise(EditorEvent e, long data = 0) {
		if (!subscribers.TryGetValue(e, out var windows)) return;
		foreach (var window in windows.ToArray())
			if (!IsWindow(window) || !PostMessage(window, Message, (IntPtr)(int)e, (IntPtr)data)) windows.Remove(window);
	}

	[DllImport("user32", CharSet = CharSet.Unicode)] static extern int RegisterWindowMessage(string name);
	[DllImport("user32")] static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	[DllImport("user32")] static extern bool IsWindow(IntPtr hWnd);
}

// Settings remembered between sessions, stored as JSON in %APPDATA%\AhkWord\settings.json (in portable mode,
// data\settings.json beside the program), or in the file named by the AHKWORD_SETTINGS (or RTE_WRITER_SETTINGS) environment variable. A
// missing or unreadable file gives the defaults.
// Iteration 41: AhkWord as an app of its own on the taskbar. It runs in AutoHotkey's process, so Windows would take it
// for AutoHotkey: its button grouped with other scripts', AutoHotkey's icon, and pinned it would start AutoHotkey alone.
// An app ID of its own ("AhkWord", for the process and its windows) and, on each window, the shell's "relaunch"
// properties (its icon, its name, the command that starts it) make the taskbar show AhkWord's icon, group AhkWord's
// windows by themselves, and start AhkWord when its pinned button is clicked.
public static class TaskbarIdentity
{
	public const string AppId = "AhkWord";

	[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
	static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
	[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
	static extern int GetCurrentProcessExplicitAppUserModelID(out IntPtr appId);
	[DllImport("shell32.dll")]
	static extern int SHGetPropertyStoreForWindow(IntPtr hwnd, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IPropertyStore store);
	[DllImport("ole32.dll")]
	static extern int PropVariantClear(ref PropVariant value);

	[ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IPropertyStore {
		int GetCount(out uint count);
		int GetAt(uint index, out PropertyKey key);
		int GetValue(ref PropertyKey key, out PropVariant value);
		int SetValue(ref PropertyKey key, ref PropVariant value);
		int Commit();
	}

	[StructLayout(LayoutKind.Sequential, Pack = 4)]
	struct PropertyKey { public Guid Format; public uint Id; public PropertyKey(Guid format, uint id) { Format = format; Id = id; } }

	// a PROPVARIANT: its type (VT_LPWSTR: 31; VT_EMPTY: 0) and, 8 bytes in, the string's pointer
	[StructLayout(LayoutKind.Explicit, Size = 24)]
	struct PropVariant { [FieldOffset(0)] public ushort Type; [FieldOffset(8)] public IntPtr Pointer; }

	static readonly Guid AppUserModel = new("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
	static readonly PropertyKey RelaunchCommand = new(AppUserModel, 2), RelaunchIcon = new(AppUserModel, 3),
		RelaunchName = new(AppUserModel, 4), Id = new(AppUserModel, 5);

	/// The process's app ID: before its first window, so that the taskbar knows it from the start.
	public static void SetProcess() {
		try { SetCurrentProcessExplicitAppUserModelID(AppId); } catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException) { }
	}

	/// The process's app ID (null: none set).
	public static string? ProcessId() {
		try {
			if (GetCurrentProcessExplicitAppUserModelID(out IntPtr id) != 0 || id == IntPtr.Zero) return null;
			string? value = Marshal.PtrToStringUni(id);
			Marshal.FreeCoTaskMem(id);
			return value;
		} catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException) { return null; }
	}

	/// A window's app ID and relaunch properties: the command that starts AhkWord, its icon (an .ico; null: none).
	public static bool SetWindow(IntPtr window, string command, string? icon) {
		var properties = new List<(PropertyKey Key, string Value)> { (Id, AppId), (RelaunchCommand, command), (RelaunchName, AppId) };
		if (icon != null) properties.Add((RelaunchIcon, icon + ",0"));
		try {
			var iid = typeof(IPropertyStore).GUID;
			if (SHGetPropertyStoreForWindow(window, ref iid, out var store) != 0) return false;
			try {
				foreach (var (key, value) in properties) {
					var k = key;
					var v = new PropVariant { Type = 31, Pointer = Marshal.StringToCoTaskMemUni(value) };
					try { store.SetValue(ref k, ref v); } finally { PropVariantClear(ref v); }
				}
				store.Commit();
			} finally { Marshal.ReleaseComObject(store); }
			return true;
		} catch (Exception ex) when (ex is COMException or DllNotFoundException or EntryPointNotFoundException or InvalidCastException) { return false; }
	}

	/// A window's properties as set (app ID, relaunch command, icon, name), for checking.
	public static (string? Id, string? Command, string? Icon, string? Name) OfWindow(IntPtr window) {
		try {
			var iid = typeof(IPropertyStore).GUID;
			if (SHGetPropertyStoreForWindow(window, ref iid, out var store) != 0) return default;
			try {
				string? Get(PropertyKey key) {
					if (store.GetValue(ref key, out var v) != 0) return null;
					try { return v.Type == 31 ? Marshal.PtrToStringUni(v.Pointer) : null; } finally { PropVariantClear(ref v); }
				}
				return (Get(Id), Get(RelaunchCommand), Get(RelaunchIcon), Get(RelaunchName));
			} finally { Marshal.ReleaseComObject(store); }
		} catch (Exception ex) when (ex is COMException or InvalidCastException) { return default; }
	}
}

public sealed class Settings
{
	public const int MaxRecentFiles = 10;
	public int[]? Window { get; set; }                  // normal (not maximized) bounds: x, y, width, height
	public bool Maximized { get; set; }
	public List<string> RecentFiles { get; set; } = new();
	public string? LastFolder { get; set; }
	public bool Structured { get; set; } = true;    // (no longer used: View > Structured / Free-form were retired in iteration 34)
	public bool WordWrap { get; set; } = true;
	public bool PrintLayout { get; set; } = true;    // View > Print Layout: lines as on the printed page
	public bool PageView { get; set; } = true;       // View > Show Pages: Print Layout's pages as sheets of their own (on from iteration 48)
	public string Theme { get; set; } = "Light";    // View > Theme: a name in AppTheme.All (iteration 70)
	public string Language { get; set; } = Lang.English;    // (AhkWord-es) the interface's language: lang\<code>.json (rte_strings.cs)
	public string ZoomMode { get; set; } = "PageWidth";    // Print Layout's zoom: PageWidth, OnePage or Manual (View > Zoom)
	public bool SpellCheck { get; set; } = true;
	public bool AutoLevelStyles { get; set; } = true;
	public float Zoom { get; set; } = 1f;
	public int[] Margins { get; set; } = { 100, 100, 100, 100 };    // left, right, top, bottom, in 1/100 inch
	public bool Landscape { get; set; }
	public int[]? Paper { get; set; }    // (iteration 82) new documents' paper: width, height upright, 1/100 inch (null: the default printer's)
	// Printed header and footer: left, center and right parts, with the fields {title} {file} {page} {pages} {date} {time}
	public bool ShowHeader { get; set; }
	public string[] Header { get; set; } = DefaultHeader;
	public bool ShowFooter { get; set; } = true;
	public bool ShowOutline { get; set; } = true;
	public bool ShowRuler { get; set; }    // View > Show > Ruler (iteration 78)
	public string SourceView { get; set; } = "None";    // View > Document View: None (rich text), Markdown or Html
	public bool SideBySide { get; set; }
	public int SourceSplit { get; set; } = 50;    // side by side: the source's share of the width, in percent
	public bool SourceWordWrap { get; set; } = true;
	public bool SourceLineNumbers { get; set; } = true;    // View > Source line numbers
	public bool ShowBands { get; set; } = true;    // View > Header and footer: the strips above and below the text
	public List<ListLevel> ListLevels { get; set; } = ListLevel.Defaults();    // Paragraph > List levels
	public int TableRows { get; set; } = 3;    // Table > Insert table
	public int TableColumns { get; set; } = 3;
	public bool TableHeader { get; set; } = true;
	public bool CrossReferenceFull { get; set; }    // (iteration 50) Insert > Cross-reference: the full number ("3(b)")
	public string LegalFormat { get; set; } = "Top";    // legal numbering: Top (1. 1.1), All (1. 1.1.) or None (1 1.1)
	public Dictionary<string, string> MacroHotkeys { get; set; } = new(StringComparer.OrdinalIgnoreCase);
	public Dictionary<string, string> EventHooks { get; set; } = new();    // (iteration 56) event name -> macro (rte_hooks.cs)
	public List<string> EventHooksOff { get; set; } = new();
	public List<string> PaletteRecent { get; set; } = new();    // (iteration 57) the palette's items used last, newest first (their keys)    // the hooks switched off (their macros kept): event names    // macro file name -> keys, e.g. "Ctrl+Alt+1"
	public int OutlineWidth { get; set; } = 220;    // at 96 dpi
	public int? PrinterModeToRestore { get; set; }
	public int AutosaveSeconds { get; set; } = 60;
	public bool ShowStartPage { get; set; } = true;
	public string StyleSetName { get; set; } = "Classic";    // the style set new documents get (unless MyStyles)
	public StyleSheet? MyStyles { get; set; }                 // Save Styles as My Default: new documents get these    // the start page at the start (without a file to open)
	public Dictionary<string, string> TemplateValues {    // a "Your ..." template field's name -> the value typed last
		get => templateValues;
		set => templateValues = new(value ?? new(), StringComparer.OrdinalIgnoreCase);    // (names ignoring case, also when read back)
	}
	Dictionary<string, string> templateValues = new(StringComparer.OrdinalIgnoreCase);    // between recovery copies while there are unsaved changes (0: none)    // Windows' LegacyDefaultPrinterMode while a PDF export changed it (-1: none)
	public string[] Footer { get; set; } = DefaultFooter;
	public bool DifferentFirst { get; set; }    // (iteration 49) the first page's own header and footer
	public string[] FirstHeader { get; set; } = { "", "", "" };
	public string[] FirstFooter { get; set; } = { "", "", "" };
	public BandFont? HeaderFont { get; set; }    // (iteration 55) null: Segoe UI 8 pt in grey
	public BandFont? FooterFont { get; set; }
	public int? HeaderDistance { get; set; }    // (iteration 55) 1/100 inch from the page's edge (null: half an inch)
	public int? FooterDistance { get; set; }
	public bool HeaderLine { get; set; }    // (iteration 55) a line below the header, above the footer
	public bool FooterLine { get; set; }
	public bool? ShowFirstHeader { get; set; }    // (iteration 55) the first page's own (null: as the other pages)
	public bool? ShowFirstFooter { get; set; }
	public BandFont? FirstHeaderFont { get; set; }
	public BandFont? FirstFooterFont { get; set; }
	public bool? FirstHeaderLine { get; set; }
	public bool? FirstFooterLine { get; set; }
	public static string[] DefaultHeader => new[] { "{title}", "", "{date}" };
	public static string[] DefaultFooter => new[] { "", PageOf, "" };
	const string PageOf = "Page {page} of {pages}";

	// The default footer's "Page {page} of {pages}" in the language for new documents ("Página {page} de {pages}"),
	// kept in English in settings.json (made default again: back to English), so that it follows a change of language.
	static string PageOfShown => T("band.pageOf", "Page {page} of {pages}");    // (PageOf: a literal for check_lang)
	public static string[] Shown(string[] parts) => parts.Select(p => p == PageOf ? PageOfShown : p).ToArray();
	public static string[] Neutral(string[] parts) => parts.Select(p => p == PageOfShown ? PageOf : p).ToArray();

	// Always three parts, whatever the file contained.
	public static string[] Parts(string[]? parts) => Enumerable.Range(0, 3).Select(i => parts != null && i < parts.Length ? parts[i] ?? "" : "").ToArray();

	// (AHKWORD_SETTINGS; RTE_WRITER_SETTINGS, its name before iteration 39, still does)
	public static string FilePath => (Environment.GetEnvironmentVariable("AHKWORD_SETTINGS") is { Length: > 0 } set ? set
		: Environment.GetEnvironmentVariable("RTE_WRITER_SETTINGS")) is { Length: > 0 } path ? path
		: Path.Combine(AppFolders.Roaming, "settings.json");

	public static Settings Load() {
		Settings settings;
		try {
			settings = File.Exists(FilePath) ? JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new Settings() : new Settings();
		} catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) {
			settings = new Settings();
		}
		settings.known = settings.Snapshot();
		return settings;
	}

	public void Save() {
		try {
			Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
			MergeFromFile();
			string json = JsonSerializer.Serialize(this, new JsonSerializerOptions {
				WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping    // "Ctrl+Alt+1", not "Ctrl+Alt+1"
			});
			// (written beside, then moved over: another window never reads half a file)
			string temporary = FilePath + "." + Environment.ProcessId + ".tmp";
			Retried(() => File.WriteAllText(temporary, json));
			Retried(() => File.Move(temporary, FilePath, true));
			known = Snapshot();
		} catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
	}

	// The file opened a moment by something else (a virus scanner, the search indexer, another window's save): tried
	// again a few times (iteration 44: putting the new file in place then failed now and then with "access denied", and
	// the save was lost, a change made in another window with it; a failed read left another window's changes out).
	static T Retried<T>(Func<T> io) {
		for (int tries = 1; ; tries++) {
			try { return io(); }
			catch (Exception ex) when (ex is IOException or UnauthorizedAccessException && tries < 5) { System.Threading.Thread.Sleep(25 * tries); }
		}
	}
	static void Retried(Action io) => Retried(() => { io(); return 0; });

	// Several AhkWord windows share this file. Each remembers the settings as it last read or wrote them (known); a
	// save first takes from the file what another window changed since and this one didn't: a setting changed only
	// there is taken over, one changed in both keeps this window's, except that lists and maps (the recent documents,
	// macro hotkeys, remembered template values) are merged item by item (this window's additions, changes and removals
	// on top of the other's).
	Dictionary<string, string>? known;
	static readonly PropertyInfo[] Kept = typeof(Settings).GetProperties(BindingFlags.Public | BindingFlags.Instance)
		.Where(p => p.CanRead && p.CanWrite && p.GetIndexParameters().Length == 0).ToArray();
	static string Json(object? value) => JsonSerializer.Serialize(value);
	Dictionary<string, string> Snapshot() => Kept.ToDictionary(p => p.Name, p => Json(p.GetValue(this)));

	void MergeFromFile() {
		if (known == null || !File.Exists(FilePath)) return;
		Settings? disk;
		try { disk = JsonSerializer.Deserialize<Settings>(Retried(() => File.ReadAllText(FilePath))); } catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return; }
		if (disk == null) return;
		foreach (var property in Kept) {
			string before = known.TryGetValue(property.Name, out var b) ? b : "", ours = Json(property.GetValue(this)), theirs = Json(property.GetValue(disk));
			if (theirs == before || theirs == ours) continue;                                  // (not changed by another window)
			if (ours == before) { property.SetValue(this, property.GetValue(disk)); continue; }    // (changed only there)
			if (property.GetValue(this) is List<string> ourList && property.GetValue(disk) is List<string> theirList) {
				// (what either window removed since stays removed; what either added is kept, this window's first)
				var baseList = JsonSerializer.Deserialize<List<string>>(before) ?? new();
				var removed = baseList.Where(x => !ourList.Contains(x, StringComparer.OrdinalIgnoreCase) || !theirList.Contains(x, StringComparer.OrdinalIgnoreCase)).ToList();
				var merged = ourList.Concat(theirList.Where(x => !ourList.Contains(x, StringComparer.OrdinalIgnoreCase)))
					.Where(x => !removed.Contains(x, StringComparer.OrdinalIgnoreCase)).ToList();
				if (property.Name == nameof(RecentFiles)) merged = merged.Take(MaxRecentFiles).ToList();
				property.SetValue(this, merged);
			} else if (property.GetValue(this) is Dictionary<string, string> ourMap && property.GetValue(disk) is Dictionary<string, string> theirMap) {
				var baseMap = JsonSerializer.Deserialize<Dictionary<string, string>>(before) ?? new();
				var merged = new Dictionary<string, string>(theirMap, ourMap.Comparer);
				foreach (var (key, value) in ourMap) if (!baseMap.TryGetValue(key, out var was) || was != value) merged[key] = value;
				foreach (var key in baseMap.Keys) if (!ourMap.ContainsKey(key)) merged.Remove(key);
				property.SetValue(this, merged);
			}
		}
	}
}

// Keeps the default printer as it is while AhkWord prints to a printer the user didn't choose (PDF export). While
// Windows manages the default printer ("Let Windows manage my default printer", on by default), printing makes the
// printer used the default, also in Windows' record for the network the computer is on; and Windows applies it again
// when the process ends, so putting the default back afterwards doesn't last. Instead, Windows is told that the user
// manages the default (LegacyDefaultPrinterMode = 1) for the print job, and its own setting is put back as soon as the
// job is done (checked: the default and Windows' record stay as they were, also after the process ends). The setting to
// put back is kept in the settings file meanwhile, and put back at the next start if AhkWord ended in between.
sealed class DefaultPrinterGuard : IDisposable
{
	const string Key = @"Software\Microsoft\Windows NT\CurrentVersion\Windows", Value = "LegacyDefaultPrinterMode";
	readonly Settings settings;
	readonly bool active;

	public DefaultPrinterGuard(Settings settings) {
		this.settings = settings;
		try {
			using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Key, true);
			if (key == null) return;
			object? mode = key.GetValue(Value);
			if (mode is int user && user == 1) return;    // the user manages the default: printing doesn't change it
			settings.PrinterModeToRestore = mode is int m ? m : -1;
			settings.Save();
			key.SetValue(Value, 1, Microsoft.Win32.RegistryValueKind.DWord);
			active = true;
		} catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException) { }
	}

	public void Dispose() {
		if (active) Restore(settings);
	}

	/// Puts back Windows' setting as a guard found it (after its job, or at the next start if AhkWord ended meanwhile).
	public static void Restore(Settings settings) {
		if (settings.PrinterModeToRestore is not int mode) return;
		try {
			using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Key, true);
			if (key != null) {
				if (mode < 0) key.DeleteValue(Value, false);
				else key.SetValue(Value, mode, Microsoft.Win32.RegistryValueKind.DWord);
			}
		} catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException) { return; }
		settings.PrinterModeToRestore = null;
		settings.Save();
	}
}

// Prints the editor's content: RichEdit lays out each page itself (EM_FORMATRANGE), and a footer shows the page number.
sealed class DocumentPrinter : IDisposable
{
	public readonly PrintDocument Document = new();
	readonly RichEditBox editor;
	readonly string title, path;
	readonly Settings settings;
	readonly PageBands bands;
	readonly LineNumbering numbering;
	List<List<PrintedLineNumbers.Mark>>? numberMarks;    // each page's line numbers (worked out with the page count)
	List<List<(float X1, float X2, float Baseline, float Size, string Kind)>>? pageLeaders;    // each page's leaders (iteration 52)
	List<int>? pageStarts;                                  // where each page starts (worked out with the page count)
	List<List<Footnotes.NoteLine>>? pageNotes;              // each page's footnotes' lines (iteration 51; with the page count)

	/// The notes' font the pages were last laid out with, and (PageHeights) each page's notes' lines (iteration 51).
	public (string Name, float Size) NoteFontUsed { get; private set; }

	/// The document's Footnote Text style's font and size (null: none; Footnotes.NoteFontOf).
	public (string Font, float Size)? NoteStyle { get; init; }
	public List<List<Footnotes.NoteLine>> LastNotes { get; private set; } = new();
	int nextChar, page, pageCount;

	// (iteration 67) a column of a page: its text and how tall it is (twips); a band across the page: its columns (null:
	// one), how far down the text area it starts (twips) and its columns' text
	sealed record ColumnPart(int Start, int End, int Height);
	sealed record PageBand(SectionColumns? Columns, int Top, List<ColumnPart> Parts)
	{
		public int Height => Parts.Max(p => p.Height);
	}
	List<List<PageBand>?>? pageLayout;    // each page's bands (null: a page of one column; with the page count)

	const int EM_FORMATRANGE = 0x0400 + 57;

	[StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
	[StructLayout(LayoutKind.Sequential)] struct CHARRANGE { public int cpMin, cpMax; }
	[StructLayout(LayoutKind.Sequential)] struct FORMATRANGE { public IntPtr hdc, hdcTarget; public RECT rc, rcPage; public CHARRANGE chrg; }
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref FORMATRANGE lParam);
	[DllImport("user32")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	/// Iteration 63: the document's sections (null: one, with the bands given): each page has its section's header and
	/// footer (the section's first page its own first-page ones), text area and page number.
	public IReadOnlyList<PageSection>? PageSections {
		get => sections;
		init => sections = value is { Count: > 0 } ? value : null;
	}
	IReadOnlyList<PageSection>? sections;
	IReadOnlyList<PageSection> AllSections => sections ?? new[] { new PageSection(0, bands, true, 1) };
	List<Sections.PageNumber>? pageNumbers;    // (each page's section and number: worked out with the page count)
	int[]? baseMargins;                        // (the page setup's margins: left, right, top, bottom)

	// (iteration 63: margins and landscape, the document's own page setup where it has one, else the settings')
	public DocumentPrinter(RichEditBox editor, string title, string path, Settings settings, PageBands bands, LineNumbering numbering, string? printerName, int[]? margins = null, bool? landscape = null, int[]? paper = null) {
		this.numbering = numbering.Copy();
		this.editor = editor;
		this.bands = bands;
		this.title = title;
		this.path = path;
		this.settings = settings;
		Document.DocumentName = title;
		if (!string.IsNullOrEmpty(printerName)) Document.PrinterSettings.PrinterName = printerName;    // before the page settings
		if (paper is { Length: 2 }) Document.DefaultPageSettings.PaperSize = Papers.ForPrinter(Document.PrinterSettings, paper[0], paper[1]);    // (iteration 82: the document's)
		Document.DefaultPageSettings.Landscape = landscape ?? settings.Landscape;
		// (iteration 55) a header or footer reaching past its margin pushes the text in, as in Word
		if ((margins ?? settings.Margins) is { Length: 4 } m) {
			int top = Math.Max(m[2], bands.HeaderReach()), bottom = Math.Max(m[3], bands.FooterReach());
			Document.DefaultPageSettings.Margins = new Margins(m[0], m[1], top, bottom);
			baseMargins = m.ToArray();
		}
		Document.BeginPrint += (s, e) => { nextChar = 0; page = 0; pageCount = 0; numberMarks = null; pageStarts = null; pageNotes = null; pageNumbers = null; sectionSettings = null; pageSizes = null; framesAll = null; };
		// (iteration 87) each page on its section's paper, in its orientation, with its left and right margins
		// (always, a section's page or the document's: a page's settings are the page before's until changed, and a page
		// after a landscape section is upright again)
		Document.QueryPageSettings += (s, e) => {
			if (sections == null) return;
			var own = SettingsAt(nextChar);
			e.PageSettings.PaperSize = own.PaperSize;
			e.PageSettings.Landscape = own.Landscape;
			e.PageSettings.Margins = own.Margins;
		};
		Document.PrintPage += PrintPage;
		Document.EndPrint += (s, e) => SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);    // frees RichEdit's cache
	}

	// (iteration 55) a page's text area against the page setup's (1/100 inch): its header or footer reaches further or
	// less far; (iteration 63) its section's, the section's first page its own first-page ones; its section's own top and
	// bottom margins
	(int Top, int Bottom) Shift(int position) {
		if (baseMargins is not { } m) return (0, 0);
		var all = AllSections;
		var section = all[Sections.IndexAt(all, position)];
		int on = position <= section.Start ? 1 : 2;    // (a page starting where its section does: the section's first)
		var (top, bottom) = Sections.MarginsOn(section, on, m[2], m[3], PageSize(position).Height);
		return (top - Document.DefaultPageSettings.Margins.Top, bottom - Document.DefaultPageSettings.Margins.Bottom);
	}

	// (iteration 87) The page settings of a page starting at position: the document's, with its section's own
	// orientation, paper and left and right margins (top and bottom: Shift). One copy for each section, made once.
	PageSettings SettingsAt(int position) {
		var all = AllSections;
		int index = Sections.IndexAt(all, position);
		if (all[index].Page is not { } own || !own.KeptOnly) return Document.DefaultPageSettings;
		sectionSettings ??= new();
		if (sectionSettings.TryGetValue(index, out var made)) return made;
		var settings = (PageSettings)Document.DefaultPageSettings.Clone();
		if (Papers.Upright(own.Width, own.Height) is { } paper) {
			var current = settings.PaperSize;
			if (Papers.Upright(current.Width, current.Height, twips: false) is not { } now || Math.Abs(now[0] - paper[0]) > 4 || Math.Abs(now[1] - paper[1]) > 4)
				settings.PaperSize = Papers.ForPrinter(Document.PrinterSettings, paper[0], paper[1]);
		}
		// (Word gives a landscape page as wider than high, its orientation too)
		if (own.Landscape is bool landscape) settings.Landscape = landscape;
		else if (own.Width is int w && own.Height is int h && w != h) settings.Landscape = w > h;
		if (own.Left != null || own.Right != null) {
			var m = settings.Margins;
			settings.Margins = new Margins(own.Left is int left ? (int)Math.Round(left / 14.4) : m.Left, own.Right is int right ? (int)Math.Round(right / 14.4) : m.Right, m.Top, m.Bottom);
		}
		return sectionSettings[index] = settings;
	}
	Dictionary<int, PageSettings>? sectionSettings;

	// (the page's size, 1/100 inch, as its orientation has it; the printer is asked once for each section's)
	Size PageSize(int position) {
		var settings = SettingsAt(position);
		pageSizes ??= new();
		if (!pageSizes.TryGetValue(settings, out var size)) pageSizes[settings] = size = settings.Bounds.Size;
		return size;
	}
	Dictionary<PageSettings, Size>? pageSizes;

	// A page's text area (twips, from the document's first page's printable corner), from the document's: its section's
	// top and bottom (Shift), and (iteration 87) its section's page and left and right margins.
	RECT AreaAt(RECT r, int position) {
		var (top, bottom) = Shift(position);
		var own = SettingsAt(position);
		if (ReferenceEquals(own, Document.DefaultPageSettings))
			return new() { Left = r.Left, Right = r.Right, Top = r.Top + (int)Math.Round(top * 14.4), Bottom = r.Bottom - (int)Math.Round(bottom * 14.4) };
		var size = PageSize(position);
		var first = Document.DefaultPageSettings;
		var firstSize = PageSize(0);
		int left = r.Left + (int)Math.Round((own.Margins.Left - first.Margins.Left) * 14.4);
		int width = (int)Math.Round((size.Width - own.Margins.Left - own.Margins.Right) * 14.4);
		int height = r.Bottom - r.Top + (int)Math.Round((size.Height - firstSize.Height) * 14.4);
		return new() { Left = left, Right = left + Math.Max(1440, width), Top = r.Top + (int)Math.Round(top * 14.4), Bottom = r.Top + height - (int)Math.Round(bottom * 14.4) };
	}

	// The columns a page starting at position has in a text area this wide (null: one).
	SectionColumns? ColumnsAt(int position, int width) {
		var all = AllSections;
		return all[Sections.IndexAt(all, position)].Columns is { } columns && columns.FitIn(width) ? columns : null;
	}

	// Where the text that has the columns position has stops (int.MaxValue: at the document's end), where the next goes on,
	// and whether that is on the same page. (A hidden paragraph joins the line after it, so a continuous break's isn't
	// laid out: the text stops before it and goes on after it.)
	(int Stop, int Next, bool Continuous) RegionEnd(int position, int width) {
		var all = AllSections;
		SectionColumns? Of(PageSection s) => s.Columns is { } columns && columns.FitIn(width) ? columns : null;
		int index = Sections.IndexAt(all, position);
		for (int i = index + 1; i < all.Count; i++)
			if (Of(all[i]) != Of(all[index])) return (all[i].Continuous ? all[i].Stop : all[i].Start, all[i].Start, all[i].Continuous);
		return (int.MaxValue, int.MaxValue, false);
	}

	// range's text drawn: in its one area, or in each band's columns. Where the text goes on.
	int DrawText(FORMATRANGE range, List<PageBand>? layout) {
		if (layout == null) return (int)SendMessage(editor.Handle, EM_FORMATRANGE, (IntPtr)1, ref range);
		foreach (var band in layout)
			for (int i = 0; i < band.Parts.Count; i++) {
				var column = range;
				column.rc.Top = range.rc.Top + band.Top;
				if (band.Columns != null) (column.rc.Left, column.rc.Right) = band.Columns.Slot(range.rc.Left, range.rc.Right, i);
				column.chrg = new CHARRANGE { cpMin = band.Parts[i].Start, cpMax = band.Parts[i].End };
				SendMessage(editor.Handle, EM_FORMATRANGE, (IntPtr)1, ref column);
			}
		return layout[^1].Parts[^1].End;
	}

	void PrintPage(object? sender, PrintPageEventArgs e) {
		page++;
		var g = e.Graphics!;
		var allMargins = e.MarginBounds;
		var (shiftTop, shiftBottom) = Shift(nextChar);    // (iteration 63: the page's section's)
		var margins = Rectangle.FromLTRB(allMargins.Left, allMargins.Top + shiftTop, allMargins.Right, allMargins.Bottom - shiftBottom);
		float hardX = e.PageSettings.HardMarginX, hardY = e.PageSettings.HardMarginY;
		int start = nextChar;
		int end = editor.TextDocument.Range2(0, 0).StoryLength - 1;    // without the final paragraph mark
		List<PageBand>? layout = null;
		IntPtr hdc = g.GetHdc();
		try {
			var range = new FORMATRANGE {
				hdc = hdc, hdcTarget = hdc,
				rc = Twips(allMargins, hardX, hardY), rcPage = Twips(e.PageBounds, hardX, hardY),
				chrg = new CHARRANGE { cpMin = nextChar, cpMax = -1 }
			};
			if (pageCount == 0) {
				pageStarts = new List<int>();
				pageNotes = new List<List<Footnotes.NoteLine>>();
				pageCount = CountPages(range, end, pageStarts, notes: pageNotes);    // for {pages}, and where each page ends
				numberMarks = Marks(pageStarts, range, e.PageSettings);
				for (int i = 0; numberMarks != null && i < numberMarks.Count && i < pageStarts.Count; i++)    // (a page's lines lower or higher: its section's)
					if (Shift(pageStarts[i]).Top is int down and not 0)
						numberMarks[i] = numberMarks[i].Select(mark => mark with { Baseline = mark.Baseline + (int)Math.Round(down * 14.4) }).ToList();
				pageNumbers = Sections.Number(pageStarts, AllSections);    // (iteration 63)
				pageLeaders = LeadersOnPages(range, pageStarts);    // (iteration 52)
				// (iteration 67: both are measured across the whole text area, so not on pages in columns)
				for (int i = 0; i < pageStarts.Count; i++)
					if (pageLayout?[i] != null) {
						if (numberMarks != null && i < numberMarks.Count) numberMarks[i] = new();
						if (pageLeaders != null && i < pageLeaders.Count) pageLeaders[i] = new();
					}
			}
			range.rc = Twips(margins, hardX, hardY);    // (page 1's own)
			range.chrg.cpMin = nextChar;
			// (the page ends where the paragraphs' rules have it: page break before, keep together, widow control, keep with next)
			range.chrg.cpMax = pageStarts != null && page < pageStarts.Count ? pageStarts[page] : -1;
			layout = pageLayout != null && page <= pageLayout.Count ? pageLayout[page - 1] : null;
			// (iteration 87) the page's paragraph borders and shading (not on pages in columns): the shading under the text
			framesOnPage = layout == null ? FramesOnPage(range, nextChar, range.chrg.cpMax < 0 ? end + 1 : range.chrg.cpMax) : null;
			if (Document.PrintController?.IsPreview == true) nextChar = PreviewPage(range, e.PageBounds, out preview, layout, e.PageSettings, framesOnPage);
			else if (framesOnPage is { Count: > 0 }) {
				// (the shading, then the text over it without the white RichEdit fills its area with first, the lines over it)
				g.ReleaseHdc(hdc);
				foreach (var place in framesOnPage) ParagraphFrames.Draw(g, Scaled(place, 1 / 14.4f), 100 / 72f, 0.5f, lines: false);
				var (recorded, drawn) = RecordPage(range, e.PageBounds, layout, e.PageSettings, framesOnPage);
				// (the page's corner and size in the printer's pixels, from its printable area's corner)
				float dpiX = g.DpiX / 100f, dpiY = g.DpiY / 100f;
				using (recorded) PlayWithoutFill(g, recorded, Rectangle.Round(new RectangleF(-hardX * dpiX, -hardY * dpiY, e.PageBounds.Width * dpiX, e.PageBounds.Height * dpiY)));
				nextChar = drawn;
				hdc = g.GetHdc();
			}
			else nextChar = DrawText(range, layout);
			nextChar = PageBreaks.AfterBreak(editor.TextDocument, nextChar);    // (a page break's paragraph mark stays with it)
		} finally {
			g.ReleaseHdc(hdc);
		}
		if (preview != null) {
			using (preview) {
				g.InterpolationMode = InterpolationMode.HighQualityBicubic;
				g.DrawImage(preview, new RectangleF(-hardX, -hardY, e.PageBounds.Width, e.PageBounds.Height));
			}
			preview = null;
		}
		// (iteration 67) the lines between the columns, as tall as the tallest
		foreach (var band in layout ?? new())
			if (band.Columns is { Line: true } lined) {
				using var pen = new Pen(Color.Black, PageBands.RuleWidth);
				float top = margins.Top - hardY + band.Top / 14.4f;
				for (int i = 0; i + 1 < band.Parts.Count; i++) {
					int across = (int)Math.Round(margins.Width * 14.4);
					float x = margins.Left - hardX + (lined.Slot(0, across, i).Right + lined.Slot(0, across, i + 1).Left) / 2f / 14.4f;
					g.DrawLine(pen, x, top, x, top + band.Height / 14.4f);
				}
			}
		// (iteration 51) the page's footnotes at the foot of its text area, under a short line
		if (pageNotes != null && page <= pageNotes.Count && pageNotes[page - 1].Count > 0) {
			using var noteFont = Footnotes.NoteFont(NoteFontUsed.Name, NoteFontUsed.Size);
			var on = pageNotes[page - 1];
			float height = Footnotes.Height(g, on, noteFont);
			Footnotes.Draw(g, on, new RectangleF(margins.Left - hardX, margins.Bottom - hardY - height, margins.Width, height), noteFont, Color.Black);
		}
		// (iteration 52) the tables of contents' leaders, which RichEdit doesn't draw
		if (pageLeaders != null && page <= pageLeaders.Count)
			foreach (var (x1, x2, baseline, size, kind) in pageLeaders[page - 1])
				Toc.DrawLeader(g, kind, margins.Left - hardX + x1 * 100 / 72f, margins.Left - hardX + x2 * 100 / 72f, margins.Top - hardY + baseline * 100 / 72f, size * 100 / 72f, Color.Black);
		// The line numbers left of the text; the header just above the text area and the footer just below it, in the margins.
		if (numberMarks != null && page <= numberMarks.Count)
			PrintedLineNumbers.Draw(g, numberMarks[page - 1], NumberingAt(start), editor.Font, margins.Left - hardX);
		// (iteration 55) the header that far from the page's top edge, the footer's last line that far from its bottom edge
		// (in the band's font: black, as Word prints it; grey without a font of its own, as before; words in their own)
		// (iteration 63) the page's section's header and footer; its place in the section (1: the section's first page)
		var number = pageNumbers != null && page <= pageNumbers.Count ? pageNumbers[page - 1] : new Sections.PageNumber(0, page, Math.Max(page, pageCount), page);
		var bands = AllSections[Math.Min(number.Section, AllSections.Count - 1)].Bands;
		int inSection = number.InSection;
		(shownPage, sectionPages) = (number.Shown, number.SectionPages);
		if (bands.HeaderOn(inSection) is { } header) {    // (iteration 49: the first page's own, if it has one)
			var texts = Settings.Parts(header).Select(Fill).ToArray();
			var font = bands.HeaderFontOn(inSection);
			float height = BandText.Draw(g, texts, font ?? BandFont.Plain, margins.Left - hardX, margins.Width, bands.HeaderFromEdge - hardY, 1, PageBands.IsPlain(font) ? Brushes.DimGray : Brushes.Black, Brushes.Black);
			if (bands.HeaderLineOn(inSection) && PageBands.LineCount(header) > 0) {    // (iteration 55: the line below it)
				float y = bands.HeaderFromEdge - hardY + height + PageBands.RuleGap;
				using var pen = new Pen(Color.Black, PageBands.RuleWidth);
				g.DrawLine(pen, margins.Left - hardX, y, margins.Right - hardX, y);
			}
		}
		if (bands.FooterOn(inSection) is { } footer) {
			var texts = Settings.Parts(footer).Select(Fill).ToArray();
			var own = bands.FooterFontOn(inSection);
			var font = own ?? BandFont.Plain;
			float height = BandText.Height(texts, font, f => f.GetHeight(g));
			float top = e.PageBounds.Height - bands.FooterFromEdge - height - hardY;
			BandText.Draw(g, texts, font, margins.Left - hardX, margins.Width, top, 1, PageBands.IsPlain(own) ? Brushes.DimGray : Brushes.Black, Brushes.Black);
			if (bands.FooterLineOn(inSection) && PageBands.LineCount(footer) > 0) {    // (iteration 55: the line above it)
				using var pen = new Pen(Color.Black, PageBands.RuleWidth);
				g.DrawLine(pen, margins.Left - hardX, top - PageBands.RuleGap, margins.Right - hardX, top - PageBands.RuleGap);
			}
		}
		e.HasMorePages = nextChar > start && nextChar < end;
		// (iteration 58) the page done, told; a stop asked for: no more pages
		Progress?.Invoke(page, Math.Max(page, pageCount));
		if (e.HasMorePages && CancelRequested?.Invoke() == true) { Cancelled = true; e.Cancel = true; e.HasMorePages = false; }
	}

	// (iteration 87) The paragraphs with borders or shading on a page (its text from start to end), laid out as range
	// lays them out: where their text is, in twips as range's rc (its lines' top and bottom, without the space before
	// and after; from the left indent, the first line's if it is further left, to the right indent).
	List<ParagraphFrames.Place> FramesOnPage(FORMATRANGE range, int start, int end) {
		var places = new List<ParagraphFrames.Place>();
		var doc = editor.TextDocument;
		framesAll ??= ParagraphFrames.All(doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "");
		var on = framesAll.Where(f => f.End >= start && f.Start < end).ToList();
		if (on.Count == 0) return places;
		int Below(int position) {
			if (position <= start) return range.rc.Top;
			var measure = range;
			measure.chrg = new CHARRANGE { cpMin = start, cpMax = position };
			SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref measure);
			return measure.rc.Bottom;
		}
		var texts = new List<RectangleF>();
		foreach (var f in on) {
			var para = doc.Range2(f.Start, f.Start).Para;
			float top = f.Start >= start ? Below(f.Start) + para.SpaceBefore * 20 : range.rc.Top;
			float bottom = f.End < end ? Below(Math.Min(end, f.End + 1)) - para.SpaceAfter * 20 : Below(end);
			float left = range.rc.Left + (para.LeftIndent + Math.Min(0, para.FirstLineIndent)) * 20, right = range.rc.Right - para.RightIndent * 20;
			texts.Add(RectangleF.FromLTRB(left, top, Math.Max(left + 20, right), Math.Max(top + 20, bottom)));
		}
		SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);    // (frees RichEdit's cache)
		for (int i = 0; i < on.Count; i++) {
			bool above = i > 0 && ParagraphFrames.Joined(on[i - 1], on[i]), below = i + 1 < on.Count && ParagraphFrames.Joined(on[i], on[i + 1]);
			var text = texts[i];
			if (below) text = RectangleF.FromLTRB(text.Left, text.Top, text.Right, Math.Max(text.Bottom, texts[i + 1].Top));    // (one box: the space between shaded too)
			places.Add(new ParagraphFrames.Place(on[i], text, above, below));
		}
		return places;
	}
	List<ParagraphFrames.Framed>? framesAll;
	List<ParagraphFrames.Place>? framesOnPage;

	static ParagraphFrames.Place Scaled(ParagraphFrames.Place place, float scale, float dx = 0, float dy = 0) =>
		place with { Text = new RectangleF((place.Text.X - dx) * scale, (place.Text.Y - dy) * scale, place.Text.Width * scale, place.Text.Height * scale) };

	/// Iteration 58: told after each page is printed (the page, the pages in all); asked after each whether to stop
	/// (Cancelled then true).
	public Action<int, int>? Progress;
	public Func<bool>? CancelRequested;
	public bool Cancelled { get; private set; }

	Bitmap? preview;
	const int PreviewScale = 2;    // the preview's picture of a page: twice the screen's resolution

	// Print Preview records each page as a metafile at the printer's resolution, and shrinking that to the screen loses
	// RichEdit's thin lines (table borders, 1/2 point): there the page is drawn into a picture instead, at twice the
	// screen's resolution. The text is first recorded in a metafile made for the printer, so that each letter is where it
	// will print (drawn straight at the screen's resolution, letters are a little wider, and the end of a line could be
	// cut off), laid out for the printer as before (so the pages break in the same places).
	int PreviewPage(FORMATRANGE range, Rectangle page, out Bitmap? picture, List<PageBand>? layout, PageSettings pageSettings, List<ParagraphFrames.Place>? frames = null) {
		int width = (int)Math.Ceiling(page.Width * 0.96 * PreviewScale), height = (int)Math.Ceiling(page.Height * 0.96 * PreviewScale);
		var (recorded, next) = RecordPage(range, page, layout, pageSettings, frames);
		using (recorded) {
			var bitmap = new Bitmap(width, height);
			picture = bitmap;
			using var g = Graphics.FromImage(bitmap);
			g.Clear(Color.White);
			g.SmoothingMode = SmoothingMode.HighQuality;
			g.DrawImage(recorded, new Rectangle(0, 0, width, height));
			// (iteration 87) the paragraphs' shading under the text: the picture multiplied by its colour there (white
			// becomes the colour, the text stays as dark); played without RichEdit's white fill instead, as on a printer, the
			// picture's letters were drawn unsmoothed
			if (frames is { Count: > 0 }) {
				g.Flush();
				float perInch = width / (page.Width / 100f);
				foreach (var place in frames)
					ParagraphFrames.Draw(g, Scaled(place, perInch / 1440f, range.rcPage.Left, range.rcPage.Top), perInch / 72f, 1, (rect, color) => ParagraphFrames.Multiply(bitmap, Rectangle.Round(rect), color), lines: false);
			}
		}
		return next;
	}

	// A page's text recorded in a metafile made for the printer (so that each letter is where it will print), laid out
	// for the printer; (iteration 87) the paragraphs' border lines over it. Where the text goes on.
	(System.Drawing.Imaging.Metafile Recorded, int Next) RecordPage(FORMATRANGE range, Rectangle page, List<PageBand>? layout, PageSettings pageSettings, List<ParagraphFrames.Place>? frames) {
		int next;
		using var measuring = Document.PrinterSettings.CreateMeasurementGraphics(pageSettings);    // (iteration 87: the page's own)
		IntPtr reference = measuring.GetHdc();
		System.Drawing.Imaging.Metafile recorded;
		try {
			recorded = new System.Drawing.Imaging.Metafile(reference, new RectangleF(0, 0, page.Width / 100f, page.Height / 100f), System.Drawing.Imaging.MetafileFrameUnit.Inch, System.Drawing.Imaging.EmfType.EmfOnly);
			using var record = Graphics.FromImage(recorded);
			int dx = range.rcPage.Left, dy = range.rcPage.Top;
			IntPtr hdc = record.GetHdc();
			try {
				// the same areas, relative to the page's corner (the picture's) rather than the printable area's
				range.hdc = hdc;
				range.hdcTarget = reference;
				range.rc = new RECT { Left = range.rc.Left - dx, Top = range.rc.Top - dy, Right = range.rc.Right - dx, Bottom = range.rc.Bottom - dy };
				range.rcPage = new RECT { Left = 0, Top = 0, Right = range.rcPage.Right - dx, Bottom = range.rcPage.Bottom - dy };
				next = DrawText(range, layout);
			} finally {
				record.ReleaseHdc(hdc);
			}
			record.PageUnit = GraphicsUnit.Inch;    // (inches from the page's corner)
			foreach (var place in frames ?? new()) ParagraphFrames.Draw(record, Scaled(place, 1 / 1440f, dx, dy), 1 / 72f, 0.005f, fill: false);
		} finally {
			measuring.ReleaseHdc(reference);
		}
		return (recorded, next);
	}

	// (iteration 87) A recorded page played onto g without the white RichEdit fills its text area with before the text
	// (a PATCOPY with a white brush, which would hide the paragraphs' shading drawn first); its other fills (table cells'
	// shading) stay.
	// (Windows' own playback, record by record, onto g's device context: the place in its device's pixels)
	static void PlayWithoutFill(Graphics g, System.Drawing.Imaging.Metafile recorded, Rectangle place) {
		IntPtr emf = recorded.GetHenhmetafile();    // (the Metafile is done with after this)
		IntPtr hdc = g.GetHdc();
		try {
			var bounds = new RECT { Left = place.Left, Top = place.Top, Right = place.Right, Bottom = place.Bottom };
			EnhMetaFileProc play = (dc, table, record, objects, data) => {
				const int EMR_BITBLT = 76, PATCOPY = 0x00F00021, OBJ_BRUSH = 2;
				if (Marshal.ReadInt32(record) == EMR_BITBLT && Marshal.ReadInt32(record, 40) == PATCOPY) {    // (type, size, bounds 16, x y cx cy 16, then the raster operation)
					var brush = new LOGBRUSH();
					if (GetObject(GetCurrentObject(dc, OBJ_BRUSH), Marshal.SizeOf<LOGBRUSH>(), ref brush) > 0 && brush.lbStyle == 0 && (brush.lbColor & 0xFFFFFF) == 0xFFFFFF) return 1;    // white: left out
				}
				PlayEnhMetaFileRecord(dc, table, record, (uint)objects);
				return 1;
			};
			EnumEnhMetaFile(hdc, emf, play, IntPtr.Zero, ref bounds);
			GC.KeepAlive(play);
		} finally {
			g.ReleaseHdc(hdc);
			DeleteEnhMetaFile(emf);
		}
	}
	delegate int EnhMetaFileProc(IntPtr hdc, IntPtr table, IntPtr record, int objects, IntPtr data);
	[StructLayout(LayoutKind.Sequential)] struct LOGBRUSH { public uint lbStyle; public uint lbColor; public IntPtr lbHatch; }
	[DllImport("gdi32.dll")] static extern bool EnumEnhMetaFile(IntPtr hdc, IntPtr emf, EnhMetaFileProc proc, IntPtr data, ref RECT bounds);
	[DllImport("gdi32.dll")] static extern bool PlayEnhMetaFileRecord(IntPtr hdc, IntPtr table, IntPtr record, uint objects);
	[DllImport("gdi32.dll")] static extern bool DeleteEnhMetaFile(IntPtr emf);
	[DllImport("gdi32.dll")] static extern IntPtr GetCurrentObject(IntPtr hdc, uint type);
	[DllImport("gdi32.dll", EntryPoint = "GetObjectW")] static extern int GetObject(IntPtr obj, int size, ref LOGBRUSH buffer);

	/// The number of pages the document prints on now (laid out for the printer, without printing).
	public int PageCount() => PageStarts().Count;

	/// Where each page the document prints on starts (character positions, the first 0), laid out for the printer.
	public List<int> PageStarts() {
		var starts = new List<int>();
		Measuring(range => CountPages(range, editor.TextDocument.Range2(0, 0).StoryLength - 1, starts));
		return starts;
	}

	/// Iteration 63: each page as laid out: where it starts, its section, its number as shown, its section's page count,
	/// and its footer's parts filled in (left|center|right; "" without one).
	public List<(int Start, int Section, int Shown, int SectionPages, string Footer)> PagePlan() {
		var starts = new List<int>();
		Measuring(range => CountPages(range, editor.TextDocument.Range2(0, 0).StoryLength - 1, starts));
		pageCount = starts.Count;
		var numbers = Sections.Number(starts, AllSections);
		var plan = new List<(int, int, int, int, string)>();
		for (int i = 0; i < starts.Count; i++) {
			var n = numbers[i];
			(shownPage, sectionPages) = (n.Shown, n.SectionPages);
			var bandsOf = AllSections[n.Section].Bands;
			string footer = bandsOf.FooterOn(n.InSection) is { } parts ? string.Join("|", Settings.Parts(parts).Select(Fill)) : "";
			plan.Add((starts[i], n.Section, n.Shown, n.SectionPages, footer));
		}
		return plan;
	}

	/// Where each printed page starts, and how tall its text is (points), laid out for the printer.
	public List<(int Start, float Height)> PageHeights() {
		var starts = new List<int>();
		var heights = new List<float>();
		var notes = new List<List<Footnotes.NoteLine>>();
		Measuring(range => CountPages(range, editor.TextDocument.Range2(0, 0).StoryLength - 1, starts, heights, notes));
		LastNotes = notes;
		return starts.Select((s, i) => (s, i < heights.Count ? heights[i] : 0f)).ToList();
	}

	/// Where a character is printed: its page (from 1) and its line's printed number (0: none, or line numbers are off).
	public (int Page, int Line) PrintedLineAt(int position) {
		var starts = PageStarts();
		int page = Math.Max(0, starts.FindLastIndex(start => start <= position));
		if (!NumberingAt(position).On) return (page + 1, 0);
		List<List<PrintedLineNumbers.Mark>>? marks = null;
		Measuring(range => marks = Marks(starts, range, Document.DefaultPageSettings, everyLine: true));
		var mark = marks?[page].FirstOrDefault(m => position >= m.Start && position < m.End);
		return (page + 1, mark?.Number ?? 0);
	}

	// Lays the pages out (without printing) with the printer's information context and the page settings' text area.
	void Measuring(Action<FORMATRANGE> measure) {
		var settings = Document.DefaultPageSettings;
		using var g = Document.PrinterSettings.CreateMeasurementGraphics(settings);
		float hardX = settings.HardMarginX, hardY = settings.HardMarginY;
		var pageBounds = settings.Bounds;
		var margins = new Rectangle(settings.Margins.Left, settings.Margins.Top,
			pageBounds.Width - settings.Margins.Left - settings.Margins.Right, pageBounds.Height - settings.Margins.Top - settings.Margins.Bottom);
		IntPtr hdc = g.GetHdc();
		try {
			measure(new FORMATRANGE { hdc = hdc, hdcTarget = hdc, rc = Twips(margins, hardX, hardY), rcPage = Twips(pageBounds, hardX, hardY), chrg = new CHARRANGE { cpMax = -1 } });
		} finally {
			g.ReleaseHdc(hdc);
		}
	}

	// Each page's tables of contents' leaders (iteration 52; null: none): from where on the screen they go across (in
	// Print Layout the lines break as printed, so the screen's points from the text's left are the page's), on the
	// baseline of the number's line as laid out for the printer (EM_FORMATRANGE up to the line's start gives its top;
	// the font's ascent below it): points from the text area's left and top.
	List<List<(float X1, float X2, float Baseline, float Size, string Kind)>>? LeadersOnPages(FORMATRANGE range, List<int> starts) {
		if (!settings.PrintLayout || editor is not WriterEditBox box || box.Leaders?.Invoke() is not { Count: > 0 } leaders || starts.Count == 0) return null;
		float ppp = box.PixelsPerPoint;
		int left = box.TextArea.Left;
		var pages = starts.Select(_ => new List<(float, float, float, float, string)>()).ToList();
		foreach (var leader in leaders) {
			if (box.LeaderLine(leader) is not { } line) continue;
			int page = starts.FindLastIndex(s => s <= leader.Number);
			if (page < 0) continue;
			// (the top of the number's line: the text laid out up to its start; the baseline the font's ascent below it)
			var numberLine = editor.TextDocument.Range2(leader.Number, leader.Number);
			numberLine.Expand(5);    // tomLine
			float top = 0;
			if (numberLine.Start > starts[page]) {
				var measure = range;
				measure.chrg = new CHARRANGE { cpMin = starts[page], cpMax = numberLine.Start };
				SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref measure);
				top = (measure.rc.Bottom - range.rc.Top) / 20f;
			}
			string fontName = editor.TextDocument.Range2(leader.Number, leader.Number + 1).Font.Name ?? editor.Font.Name;
			float ascent;
			try { using var family = new FontFamily(fontName); ascent = line.Size * family.GetCellAscent(FontStyle.Regular) / family.GetEmHeight(FontStyle.Regular); }
			catch (ArgumentException) { ascent = line.Size * 0.9f; }
			pages[page].Add(((line.X1 - left) / ppp, (line.X2 - left) / ppp, top + ascent, line.Size, leader.Kind));
		}
		SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);    // (frees RichEdit's cache)
		return pages;
	}

	/// Iteration 53: where the headings are printed, for a PDF's bookmarks (each heading's level, text, page from 0 and
	/// its top in points from the page's top), and the tables of contents' entries, for links to their headings (each
	/// entry's line across the text area, its page and area, and its heading's page and top): laid out as this printer
	/// prints the pages (EM_FORMATRANGE up to a paragraph's start gives its top on its page).
	public (List<PdfMarks.Bookmark> Bookmarks, List<PdfMarks.Link> Links) PdfPlaces(List<Toc.Heading> headings) {
		var bookmarks = new List<PdfMarks.Bookmark>();
		var links = new List<PdfMarks.Link>();
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		var entries = Toc.Entry.Matches(story).Cast<System.Text.RegularExpressions.Match>().ToList();
		if (headings.Count == 0 && entries.Count == 0) return (bookmarks, links);
		var settings = Document.DefaultPageSettings;
		float marginTop = settings.Margins.Top * 0.72f, marginLeft = settings.Margins.Left * 0.72f;
		float textWidth = (settings.Bounds.Width - settings.Margins.Left - settings.Margins.Right) * 0.72f;
		Measuring(range => {
			var starts = new List<int>();
			CountPages(range, doc.Range2(0, 0).StoryLength - 1, starts);
			if (starts.Count == 0) return;
			// a position's page, and the top (points from the page's top) of the line it starts
			(int Page, float Top) Place(int position) {
				int page = Math.Max(0, starts.FindLastIndex(s => s <= position));
				float pageTop = marginTop + Shift(starts[page]).Top * 0.72f;    // (iteration 55: page 1's own; 63: its section's)
				// (iteration 67: from the top of its column, in its band)
				int from = starts[page];
				var measure = range;
				foreach (var band in (pageLayout != null && page < pageLayout.Count ? pageLayout[page] : null) ?? new())
					for (int i = 0; i < band.Parts.Count; i++)
						if (band.Parts[i].Start <= position) {
							from = band.Parts[i].Start;
							(measure.rc.Left, measure.rc.Right) = band.Columns?.Slot(range.rc.Left, range.rc.Right, i) ?? (range.rc.Left, range.rc.Right);
							pageTop = marginTop + Shift(starts[page]).Top * 0.72f + band.Top / 20f;
						}
				if (position <= from) return (page, pageTop);
				measure.chrg = new CHARRANGE { cpMin = from, cpMax = position };
				SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref measure);
				return (page, pageTop + (measure.rc.Bottom - range.rc.Top) / 20f);
			}
			var headingAt = new Dictionary<string, (int Page, float Top)>(StringComparer.Ordinal);
			foreach (var heading in headings) {
				var (page, top) = Place(heading.Start);
				bookmarks.Add(new PdfMarks.Bookmark(heading.Level, heading.Text, page, top));
				if (heading.Name != null) headingAt[heading.Name] = (page, top);
			}
			foreach (var entry in entries) {
				if (!headingAt.TryGetValue(entry.Groups[1].Value, out var to)) continue;
				var paragraph = doc.Range2(entry.Index, entry.Index);
				paragraph.Expand(4);    // tomParagraph
				var (page, top) = Place(paragraph.Start);
				var (endPage, bottom) = Place(Math.Max(paragraph.Start + 1, paragraph.End - 1));
				// (laid out to its last character, the text ends at the bottom of its last line)
				float height = endPage == page && bottom > top ? bottom - top : 14;
				links.Add(new PdfMarks.Link(page, new RectangleF(marginLeft, top, textWidth, height), to.Page, to.Top));
			}
			SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);    // (frees RichEdit's cache)
		});
		return (bookmarks, links);
	}

	// (iteration 87) The line numbering at a position: its section's own, else the document's; and whether any section
	// has line numbers.
	LineNumbering NumberingAt(int position) {
		var all = AllSections;
		return all[Sections.IndexAt(all, position)].LineNumbers ?? numbering;
	}
	bool AnyNumbering => numbering.On || AllSections.Any(s => s.LineNumbers?.On == true);

	// Each page's line numbers (null: off), for the text area and page of range, with the printer's information context
	// (it outlives the layout; a page's own device context may be a preview's metafile).
	// (iteration 87: each page's section's own numbering, at its page's width; the top as page 1's, the pages' own
	// lower or higher tops added when printed)
	List<List<PrintedLineNumbers.Mark>>? Marks(List<int> starts, FORMATRANGE range, PageSettings settings, bool everyLine = false) {
		if (!AnyNumbering) return null;
		using var g = Document.PrinterSettings.CreateMeasurementGraphics(settings);
		IntPtr ic = g.GetHdc();
		try {
			static Rectangle Rect(RECT r) => Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
			Rectangle TextOf(int page) {
				var own = AreaAt(range.rc, starts[page]);
				return Rectangle.FromLTRB(own.Left, range.rc.Top, own.Right, range.rc.Bottom);
			}
			return PrintedLineNumbers.Compute(editor, page => NumberingAt(starts[page]), ic, TextOf, Rect(range.rcPage), starts, everyLine);
		} finally {
			g.ReleaseHdc(ic);
		}
	}

	// Lays out the whole document without drawing (EM_FORMATRANGE with wParam 0) to count the pages (and list where
	// they start), each page ending where the paragraphs' rules have it (PageKeeping).
	int CountPages(FORMATRANGE range, int end, List<int>? starts = null, List<float>? heights = null, List<List<Footnotes.NoteLine>>? notes = null) {
		int count = 0, position = 0;
		var doc = editor.TextDocument;
		string story = doc.Range2(0, doc.Range2(0, 0).StoryLength).Text ?? "";
		// (iteration 67) the printed lines of a column of that width, and the columns of a page (text laid out for each in
		// turn, the page ending early at a page break)
		var columnLines = new Dictionary<int, PrintedLines>();
		IReadOnlyList<int>? LinesIn(int width, int from, int to) {
			if (!columnLines.TryGetValue(width, out var own)) columnLines[width] = own = new PrintedLines(editor, Document.PrinterSettings.CreateMeasurementGraphics(Document.DefaultPageSettings), width);
			return own.Starts(from, to);
		}
		bool EndsPage(int at) => (at >= 2 && story[at - 2] == PageBreaks.Break) || (at >= 1 && story[at - 1] == PageBreaks.Break);
		(List<ColumnPart> Parts, int Next) LayColumns(SectionColumns columns, RECT area, int from, int limit) {
			var parts = new List<ColumnPart>();
			int next = from;
			for (int i = 0; i < columns.Count; i++) {
				int begin = next;
				var column = area;
				(column.Left, column.Right) = columns.Slot(area.Left, area.Right, i);
				range.rc = column;
				range.chrg.cpMin = begin;
				range.chrg.cpMax = limit == int.MaxValue ? -1 : limit;
				int laidOut = (int)SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
				int height = range.rc.Bottom - area.Top;
				int width = column.Right - column.Left;
				next = PageBreaks.AfterBreak(doc, laidOut >= limit ? PageKeeping.BreakBefore(doc, story, begin, laidOut) ?? laidOut : PageKeeping.Adjust(doc, story, begin, laidOut, (a, b) => LinesIn(width, a, b)));
				if (next <= begin) { next = begin; break; }
				if (next < laidOut) {    // (a rule ended it earlier)
					range.rc = column;
					range.chrg.cpMax = next;
					SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
					height = range.rc.Bottom - area.Top;
				}
				parts.Add(new ColumnPart(begin, next, height));
				if (next >= end || next >= limit || EndsPage(next)) break;
			}
			return (parts, next);
		}
		// a page's bands, from the top of area: full-width text, columns, and so on, as far as the page goes; columns that
		// end on the page with text after them there are balanced, the shortest height that holds them
		(List<PageBand> Bands, int Next) LayBands(RECT area, int from) {
			var layout = new List<PageBand>();
			int at = from, top = area.Top, width = area.Right - area.Left;
			while (top < area.Bottom) {
				var (limit, onwardAt, onward) = RegionEnd(at, width);
				if (onward && at >= limit) { at = onwardAt; continue; }
				var room = area;
				room.Top = top;
				var columns = ColumnsAt(at, width);
				List<ColumnPart> parts;
				int next;
				if (columns == null) {
					range.rc = room;
					range.chrg.cpMin = at;
					range.chrg.cpMax = limit == int.MaxValue ? -1 : limit;
					int laidOut = (int)SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
					int height = range.rc.Bottom - top;
					next = PageBreaks.AfterBreak(doc, laidOut >= limit ? PageKeeping.BreakBefore(doc, story, at, laidOut) ?? laidOut : PageKeeping.Adjust(doc, story, at, laidOut, (a, b) => LinesIn(width, a, b)));
					if (next <= at) break;
					if (next < laidOut) {
						range.rc = room;
						range.chrg.cpMax = next;
						SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
						height = range.rc.Bottom - top;
					}
					parts = new() { new ColumnPart(at, next, height) };
				}
				else {
					(parts, next) = LayColumns(columns, room, at, limit);
					if (parts.Count == 0) break;
					if (onward && next >= limit) {
						int low = 0, high = room.Bottom - room.Top;
						while (high - low > 40) {
							int middle = (low + high) / 2;
							var shorter = room;
							shorter.Bottom = room.Top + middle;
							var (tried, reached) = LayColumns(columns, shorter, at, limit);
							if (reached >= limit) { (parts, next) = (tried, reached); high = middle; } else low = middle;
						}
					}
				}
				var band = new PageBand(columns, top - area.Top, parts);
				layout.Add(band);
				top += band.Height;
				at = next;
				if (at >= end || at < limit || !onward || EndsPage(at)) break;
				at = onwardAt;
			}
			return (layout, at);
		}
		var laidBands = new List<List<PageBand>?>();
		pageLayout = laidBands;
		var otherArea = range.rc;    // (EM_FORMATRANGE gives back the height it used in rc: each page gets the whole area again)
		// (iteration 51) the footnotes: room for each page's notes at its foot, as lines (Footnotes.Lines), at most a
		// little over half the page; the lines that don't fit go on at the foot of the next page, before its own, as in
		// Word (and on the last page, in the room its text leaves, if there is enough)
		var footnotes = Footnotes.All(story);
		Graphics? measure = null;
		Font? noteFont = null;
		var (noteName, noteSize) = NoteFontUsed = Footnotes.NoteFontOf(doc, footnotes, editor.Font.Name, editor.Font.SizeInPoints, NoteStyle);
		float widthUnits = (otherArea.Right - otherArea.Left) / 14.4f;    // (1/100 inch: the measuring graphics' unit)
		var linesOf = new Dictionary<(int, float), List<Footnotes.NoteLine>>();    // each footnote's lines, by its place (and the page's width)
		List<Footnotes.NoteLine> LinesOf(List<Footnotes.Footnote> on) {
			var lines = new List<Footnotes.NoteLine>();
			if (on.Count == 0) return lines;
			measure ??= Document.PrinterSettings.CreateMeasurementGraphics(Document.DefaultPageSettings);
			noteFont ??= Footnotes.NoteFont(noteName, noteSize);
			foreach (var f in on) {
				if (!linesOf.TryGetValue((f.Start, widthUnits), out var own)) linesOf[(f.Start, widthUnits)] = own = Footnotes.Lines(measure, new() { f }, widthUnits, noteFont);
				lines.AddRange(own);
			}
			return lines;
		}
		int Twips(IReadOnlyList<Footnotes.NoteLine> lines) => lines.Count == 0 ? 0 : (int)Math.Ceiling(Footnotes.Height(measure!, lines, noteFont!) * 14.4f);
		(List<Footnotes.NoteLine> Shown, List<Footnotes.NoteLine> Left) Split(List<Footnotes.NoteLine> lines, int room) {
			if (lines.Count == 0) return (lines, new());
			int fit = Math.Max(1, Footnotes.Fitting(measure!, lines, noteFont!, room / 14.4f));
			return (lines.Take(fit).ToList(), lines.Skip(fit).ToList());
		}
		var carried = new List<Footnotes.NoteLine>();    // a note's lines from the page before
		List<Footnotes.Footnote> On(int from, int to) => footnotes.Count == 0 ? new() : Footnotes.In(footnotes, from, to >= end ? int.MaxValue : to);
		try {
			while (true) {
				count++;
				var area = AreaAt(otherArea, position);    // (iteration 55: page 1's own; 63: the page's section's)
				widthUnits = (area.Right - area.Left) / 14.4f;    // (iteration 87: its notes as wide as its text)
				starts?.Add(position);
				// (iteration 67) a page in columns, or with text that goes on in columns: the notes across its foot, the bands above them
				int across = area.Right - area.Left;
				bool banded = ColumnsAt(position, across) != null;
				if (!banded && RegionEnd(position, across) is { Continuous: true } region) {
					range.rc = area;
					range.chrg.cpMin = position;
					range.chrg.cpMax = -1;
					banded = (int)SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range) > region.Stop;
				}
				if (banded) {
					int reserved = 0, most = (area.Bottom - area.Top) / 2;
					List<PageBand> layout = new();
					int after = position;
					for (int pass = 0; pass < 4; pass++) {
						var above = area;
						above.Bottom -= reserved;
						(layout, after) = LayBands(above, position);
						int needed = Math.Min(most, Twips(carried.Concat(LinesOf(On(position, after))).ToList()));
						if (needed <= reserved) break;
						reserved = needed;
					}
					var (onPage, later) = Split(carried.Concat(LinesOf(On(position, after))).ToList(), reserved);
					notes?.Add(onPage);
					carried = later;
					laidBands.Add(layout.Count > 0 ? layout : null);
					heights?.Add(layout.Count > 0 ? (layout[^1].Top + layout[^1].Height) / 20f : 0);
					if (after <= position || after >= end) break;
					position = after;
					continue;
				}
				laidBands.Add(null);
				range.rc = area;
				range.chrg.cpMin = position;
				range.chrg.cpMax = -1;
				int laidOut = (int)SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
				int used = range.rc.Bottom - area.Top;
				int next = PageBreaks.AfterBreak(doc, PageKeeping.Adjust(doc, story, position, laidOut, (a, b) => LinesIn(area.Right - area.Left, a, b)));
				// the page's notes (the ones going on from the page before, then its own): the page ends at the last of its
				// lines (as laid out) for which the text, the notes going on and its own notes up to the last one's first line
				// fit; the notes' lines that fit below the text are shown, the others go on at the next page's foot (Word's way)
				var pageArea = area;
				var shown = new List<Footnotes.NoteLine>();
				var rest = new List<Footnotes.NoteLine>();
				if (carried.Count > 0 || On(position, next).Count > 0) {
					int TextHeight(int to) {
						range.rc = area;
						range.chrg.cpMin = position;
						range.chrg.cpMax = to;
						SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
						return range.rc.Bottom - area.Top;
					}
					List<Footnotes.NoteLine> NotesTo(int to) => carried.Concat(LinesOf(On(position, to))).ToList();
					// (what must fit: the lines going on, and each own note's lines up to the last one's first line)
					int Needed(List<Footnotes.NoteLine> noteLines, int own) {
						if (noteLines.Count == 0) return 0;
						int lastFirst = own == 0 ? Math.Min(noteLines.Count, 1) : noteLines.FindLastIndex(l => l.Number != null) + 1;
						return Twips(noteLines.Take(Math.Max(1, lastFirst)).ToList());
					}
					bool Fits(int to, bool whole) => TextHeight(to) + (whole ? Twips(NotesTo(to)) : Needed(NotesTo(to), On(position, to).Count)) <= area.Bottom - area.Top;
					var ends = (LinesIn(area.Right - area.Left, position, next) ?? Array.Empty<int>()).Where(at => at > position).Append(next).Distinct().OrderBy(at => at).ToList();
					int Furthest(bool whole) {
						int best = -1, lo = 0, hi = ends.Count - 1;
						while (lo <= hi) {
							int mid = (lo + hi) / 2;
							if (Fits(ends[mid], whole)) { best = mid; lo = mid + 1; } else hi = mid - 1;
						}
						return best;
					}
					// the notes whole on the page (Word's way: their references' lines go to the next page with them), unless
					// that leaves more than a quarter of the page empty against a note going on to the next page (a long one)
					int whole = Furthest(true), split = Furthest(false);
					int best = whole;
					if (split > whole && (whole < 0 || TextHeight(ends[split]) - TextHeight(ends[whole]) > (area.Bottom - area.Top) / 4)) best = split;
					// (not even one line with its note: the page's first line, its notes as they come)
					int chosen = best >= 0 ? ends[best] : ends[0];
					if (chosen < next) {
						int kept = PageBreaks.AfterBreak(doc, PageKeeping.Adjust(doc, story, position, chosen, (a, b) => LinesIn(area.Right - area.Left, a, b)));
						next = kept > position ? kept : chosen;
					}
					used = TextHeight(next);
					laidOut = next;
					pageArea = area;
					pageArea.Bottom = area.Top + used;
					(shown, rest) = Split(NotesTo(next), area.Bottom - area.Top - used);
				}
				notes?.Add(shown);
				carried = rest;
				if (heights != null) {
					// (the height of the text the page holds: measured again if a rule ended it earlier)
					if (next < laidOut && next > position) {
						range.rc = pageArea;
						range.chrg.cpMax = next;
						SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, ref range);
						used = range.rc.Bottom - area.Top;
					}
					heights.Add(used / 20f);
				}
				if (next <= position || next >= end) break;
				position = next;
			}
		} finally {
			foreach (var own in columnLines.Values) own.Dispose();
			measure?.Dispose();
			noteFont?.Dispose();
		}
		SendMessage(editor.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero);
		return count;
	}

	int shownPage = 1, sectionPages = 1;    // (iteration 63: the page's number as shown, its section's page count)

	string Fill(string text) {
		if (string.IsNullOrEmpty(text)) return "";
		var now = DateTime.Now;
		return Regex.Replace(text, @"\{(title|file|page|pages|sectionpages|date|time)\}", m => m.Groups[1].Value.ToLowerInvariant() switch {
			"title" => title,
			"file" => path != "" ? path : title,
			"page" => shownPage.ToString(CultureInfo.CurrentCulture),
			"pages" => pageCount.ToString(CultureInfo.CurrentCulture),
			"sectionpages" => sectionPages.ToString(CultureInfo.CurrentCulture),
			"date" => now.ToShortDateString(),
			_ => now.ToShortTimeString()
		}, RegexOptions.IgnoreCase);
	}

	// 1/100 inch (relative to the page) to twips (relative to the printable area, where the printer's HDC starts)
	static RECT Twips(Rectangle r, float hardX, float hardY) => new() {
		Left = (int)((r.Left - hardX) * 14.4f), Top = (int)((r.Top - hardY) * 14.4f),
		Right = (int)((r.Right - hardX) * 14.4f), Bottom = (int)((r.Bottom - hardY) * 14.4f)
	};

	/// After the printer is done with the editor (the page view puts its gaps back).
	public event Action? Disposed;

	public void Dispose() {
		Document.Dispose();
		var done = Disposed;
		Disposed = null;
		done?.Invoke();
	}
}

// (iteration 67) Layout > Columns: the columns of the caret's section.
sealed class ColumnsDialog : Form
{
	readonly NumericUpDown count = new() { Minimum = 1, Width = 60, AccessibleName = T("dlg.columns.countName", "Number of columns") };
	readonly NumericUpDown gap = new() { DecimalPlaces = 2, Increment = 0.05m, Width = 70, AccessibleName = T("dlg.columns.gapName", "Gap between columns") };
	readonly CheckBox line = new() { Text = T("dlg.columns.line", "Line between columns"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
	readonly CheckBox equal = new() { Text = T("dlg.columns.equal", "Equal widths"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
	readonly NumericUpDown[] widths = Enumerable.Range(1, 6).Select(i => new NumericUpDown { DecimalPlaces = 2, Increment = 0.05m, Minimum = 0.05m, Maximum = 99, Width = 62, Margin = new Padding(0, 3, 6, 3), AccessibleName = T("dlg.columns.widthName", "Width of column {0}", i) }).ToArray();
	readonly bool metric = System.Globalization.RegionInfo.CurrentRegion.IsMetric;

	int Twips(decimal shown) => (int)Math.Round(Math.Round(metric ? shown / 0.0254m : shown * 100) * 14.4m);
	new decimal Shown(int twips) => Math.Round(metric ? twips / 14.4m * 0.0254m : twips / 1440m, 2);

	/// The columns chosen (null: one).
	public SectionColumns? Result => SectionColumns.Of((int)count.Value, Twips(gap.Value), line.Checked, equal.Checked ? null : widths.Take((int)count.Value).Select(w => Twips(w.Value)).ToArray());

	public ColumnsDialog(SectionColumns? columns, int sectionNumber, bool sections, bool selection, int textWidth = 9360) {
		Dialogs.Setup(this, selection ? T("dlg.columns.titleSelection", "Columns: the selected paragraphs") : sections ? T("dlg.columns.titleSection", "Columns: Section {0}", sectionNumber) : T("dlg.columns.title", "Columns"));
		count.Maximum = Math.Max(3, columns?.Count ?? 1);
		count.Value = columns?.Count ?? 1;
		gap.Maximum = metric ? 5 : 2;
		decimal hundredths = (columns?.Gap ?? 720) / 14.4m;
		gap.Value = Math.Min(gap.Maximum, Math.Round(metric ? hundredths * 0.0254m : hundredths / 100m, 2));
		line.Checked = columns?.Line ?? false;
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
		table.Controls.Add(new Label { Text = T("dlg.columns.count", "Columns:"), AutoSize = true, Margin = new Padding(3, 7, 3, 0) }, 0, 0);
		table.Controls.Add(count, 1, 0);
		table.Controls.Add(new Label { Text = metric ? T("dlg.columns.gapCm", "Gap (cm):") : T("dlg.columns.gapIn", "Gap (in):"), AutoSize = true, Margin = new Padding(3, 7, 3, 0) }, 0, 1);
		table.Controls.Add(gap, 1, 1);
		table.Controls.Add(line, 0, 2);
		table.SetColumnSpan(line, 2);
		equal.Checked = columns?.Widths == null;
		table.Controls.Add(equal, 0, 3);
		table.SetColumnSpan(equal, 2);
		var widthRow = new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(LogicalToDeviceUnits(240), 0), Margin = new Padding(3, 0, 3, 0) };
		widthRow.Controls.AddRange(widths);
		table.Controls.Add(new Label { Text = metric ? T("dlg.columns.widthsCm", "Widths (cm):") : T("dlg.columns.widthsIn", "Widths (in):"), AutoSize = true, Margin = new Padding(3, 7, 3, 0) }, 0, 4);
		table.Controls.Add(widthRow, 1, 4);
		var note = new Label { Text = T("dlg.columns.note", "Columns are printed, and Print Preview shows them. You go on editing in one column. Widths are kept in proportion, to fill the page's text width."), AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(280), 0), Margin = new Padding(3, 6, 3, 8) };
		table.Controls.Add(note, 0, 5);
		table.SetColumnSpan(note, 2);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		table.Controls.Add(buttons, 0, 6);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
		// (the boxes: the widths the columns have, or, of equal width, what they would be)
		void Fill() {
			int shown = (int)count.Value;
			if (shown < 2) return;
			var taken = equal.Checked || columns?.Widths?.Length != shown ? SectionColumns.Of(shown, Twips(gap.Value), false) : columns;
			for (int i = 0; i < shown; i++) { var (left, right) = taken!.Slot(0, textWidth, i); widths[i].Value = Math.Clamp(Shown(right - left), widths[i].Minimum, widths[i].Maximum); }
		}
		void Enable() {
			gap.Enabled = line.Enabled = equal.Enabled = count.Value > 1;
			for (int i = 0; i < widths.Length; i++) {
				widths[i].Visible = i < count.Value && count.Value > 1;
				widths[i].Enabled = !equal.Checked;
			}
		}
		count.ValueChanged += (s, e) => { Fill(); Enable(); };
		gap.ValueChanged += (s, e) => { if (equal.Checked) Fill(); };
		equal.CheckedChanged += (s, e) => Enable();
		Fill();
		Enable();
	}
}

// Records editor actions as RteClient2 calls (one per line). Typed text is collected into InsertText calls, and
// repeated keys into Key(key, count).
sealed class MacroRecorder
{
	readonly List<string> lines = new();
	readonly StringBuilder typed = new();
	string? lastKey;
	int lastKeyCount;

	public bool Recording { get; private set; }
	public int Steps => lines.Count + (typed.Length > 0 ? 1 : 0);

	public void Start() {
		lines.Clear();
		typed.Clear();
		lastKey = null;
		Recording = true;
	}

	public void Type(char c) {
		if (!Recording) return;
		typed.Append(c);
		lastKey = null;
	}

	// Takes back the last typed character; false if there is none.
	public bool Backspace() {
		if (typed.Length == 0) return false;
		typed.Length--;
		return true;
	}

	// A line of AHK code that is not an Rte call (e.g. running another macro).
	public void AddLine(string line) {
		if (!Recording) return;
		Flush();
		lastKey = null;
		lines.Add(line);
	}

	public void Add(string call) {
		if (!Recording) return;
		Flush();
		lastKey = null;
		lines.Add("Rte." + call);
	}

	public void AddKey(string key) {
		if (!Recording) return;
		Flush();
		if (lastKey == key && lines.Count > 0) {
			lines[^1] = "Rte." + Call("Key", key, ++lastKeyCount);
		} else {
			lastKey = key;
			lastKeyCount = 1;
			lines.Add("Rte." + Call("Key", key));
		}
	}

	void Flush() {
		if (typed.Length == 0) return;
		lines.Add("Rte." + Call("InsertText", typed.ToString()));
		typed.Clear();
		lastKey = null;
	}

	public List<string> Stop() {
		Flush();
		Recording = false;
		return new List<string>(lines);
	}

	public static string Call(string member, params object?[] args) => $"{member}({string.Join(", ", args.Select(Literal))})";

	// An AHK v2 expression for a value.
	public static string Literal(object? value) => value switch {
		null => "\"\"",
		string s => "\"" + s.Replace("`", "``").Replace("\"", "`\"").Replace("\r\n", "`n").Replace("\n", "`n").Replace("\r", "`n").Replace("\t", "`t") + "\"",
		bool b => b ? "1" : "0",
		Enum e => Literal(e.ToString()),
		Color c => Literal($"#{c.R:X2}{c.G:X2}{c.B:X2}"),
		IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
		_ => Literal(value.ToString())
	};
}

// A text box that shows the key combination pressed in it (with Ctrl or Alt, or a function key).
// Backspace or Delete clears it.
sealed class HotkeyBox : TextBox
{
	public Keys Hotkey { get; private set; }
	public event EventHandler? HotkeyChanged;

	public HotkeyBox() {
		ReadOnly = true;
		BackColor = SystemColors.Window;
		ShortcutsEnabled = false;
		Text = "(none)";
	}

	public void SetHotkey(Keys keys) {
		Hotkey = keys;
		Text = keys == Keys.None ? "(none)" : Describe(keys);
		HotkeyChanged?.Invoke(this, EventArgs.Empty);
	}

	public static string Describe(Keys keys) => new KeysConverter().ConvertToInvariantString(keys) ?? "";

	public static Keys? Parse(string text) {
		try {
			return new KeysConverter().ConvertFromInvariantString(text.Trim()) is Keys keys ? keys : null;
		} catch (Exception ex) when (ex is ArgumentException or FormatException or NotSupportedException) {
			return null;
		}
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
		var code = keyData & Keys.KeyCode;
		var modifiers = keyData & Keys.Modifiers;
		if (code is Keys.ShiftKey or Keys.ControlKey or Keys.Menu or Keys.LWin or Keys.RWin) return true;
		if (modifiers == Keys.None && code is Keys.Back or Keys.Delete) { SetHotkey(Keys.None); return true; }
		if (modifiers == Keys.None && code is Keys.Tab or Keys.Enter or Keys.Escape) return base.ProcessCmdKey(ref msg, keyData);
		if ((modifiers & (Keys.Control | Keys.Alt)) != 0 || (code >= Keys.F1 && code <= Keys.F24)) SetHotkey(keyData);
		return true;
	}
}

// Shown when a recording stops: name, optional hotkey, and the script, which can be edited before it is saved.
sealed class SaveMacroDialog : Form
{
	readonly TextBox nameBox = new(), scriptBox = new();
	readonly HotkeyBox hotkeyBox = new();
	readonly Label warning = new() { AutoSize = true, ForeColor = AppTheme.Warning };

	public string MacroName => string.Concat(nameBox.Text.Trim().Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
	public string Script => scriptBox.Text;
	public Keys Hotkey => hotkeyBox.Hotkey;

	public SaveMacroDialog(string name, string script, Dictionary<Keys, string> menuShortcuts, WriterForm owner) {
		Dialogs.Setup(this, T("dlg.saveMacro.title", "Save recorded macro"), sizable: true);
		Size = new Size(LogicalToDeviceUnits(720), LogicalToDeviceUnits(520));
		MinimumSize = new Size(LogicalToDeviceUnits(480), LogicalToDeviceUnits(360));

		nameBox.Text = name;
		nameBox.Width = hotkeyBox.Width = LogicalToDeviceUnits(260);
		scriptBox.Multiline = true;
		scriptBox.ScrollBars = ScrollBars.Both;
		scriptBox.WordWrap = false;
		scriptBox.Font = new Font("Consolas", 10f);
		scriptBox.Dock = DockStyle.Fill;
		scriptBox.Text = script;
		hotkeyBox.HotkeyChanged += (s, e) => {
			warning.Text = menuShortcuts.TryGetValue(hotkeyBox.Hotkey, out var item) ? T("dlg.saveMacro.replacesShortcut", "This replaces the menu shortcut of {0} while the editor is active.", item) : "";
		};

		var fields = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Top };
		fields.Controls.Add(new Label { Text = T("dlg.saveMacro.name", "Name:"), AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
		fields.Controls.Add(nameBox, 1, 0);
		fields.Controls.Add(new Label { Text = T("dlg.saveMacro.hotkey", "Hotkey (optional):"), AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
		fields.Controls.Add(hotkeyBox, 1, 1);
		fields.Controls.Add(warning, 1, 2);
		fields.Controls.Add(new Label { Text = T("dlg.saveMacro.script", "Script (you can edit it):"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) }, 0, 3);
		fields.SetColumnSpan(fields.GetControlFromPosition(0, 3)!, 2);

		var save = Dialogs.Button(this, T("dlg.saveMacro.save", "Save"));
		var discard = Dialogs.Cancel(this, T("dlg.saveMacro.discard", "Discard"));
		save.Click += (s, e) => {
			if (MacroName == "") { MessageBox.Show(this, T("msg.macroNoName", "Give the macro a name."), Text); return; }
			DialogResult = DialogResult.OK;
		};
		var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Bottom };
		buttons.Controls.AddRange(new Control[] { discard, save });
		Controls.Add(scriptBox);
		Controls.Add(fields);
		Controls.Add(buttons);
		CancelButton = discard;
	}
}

// Tools > Macro hotkeys: one hotkey box per macro.
sealed class MacroHotkeysDialog : Form
{
	readonly Dictionary<string, HotkeyBox> boxes = new();
	readonly Label warning = new() { AutoSize = true, ForeColor = AppTheme.Warning, MaximumSize = new Size(600, 0) };
	readonly Dictionary<Keys, string> menuShortcuts;

	public IEnumerable<(string File, Keys Keys)> Hotkeys => boxes.Select(p => (p.Key, p.Value.Hotkey));

	public MacroHotkeysDialog(List<string> files, Func<string, Keys> current, Dictionary<Keys, string> menuShortcuts) {
		this.menuShortcuts = menuShortcuts;
		Dialogs.Setup(this, T("dlg.macroHotkeys.title", "Macro hotkeys"));

		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		layout.Controls.Add(new Label { AutoSize = true, Text = T("dlg.macroHotkeys.help", "Click a box and press a key combination with Ctrl or Alt, or a function key. Backspace clears it.\nThe hotkeys work while the editor window is active.") });
		var scroller = new Panel { AutoScroll = true, Width = LogicalToDeviceUnits(560), Height = LogicalToDeviceUnits(Math.Min(360, 34 * files.Count + 10)) };
		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Top };
		int row = 0;
		foreach (var file in files) {
			var box = new HotkeyBox { Width = LogicalToDeviceUnits(170) };
			box.SetHotkey(current(file));
			box.HotkeyChanged += (s, e) => Check();
			var clear = new Button { Text = T("dlg.macroHotkeys.clear", "Clear"), AutoSize = true };
			clear.Click += (s, e) => box.SetHotkey(Keys.None);
			table.Controls.Add(new Label { Text = Path.GetFileNameWithoutExtension(file), AutoSize = true, Anchor = AnchorStyles.Left, MaximumSize = new Size(LogicalToDeviceUnits(280), 0) }, 0, row);
			table.Controls.Add(box, 1, row);
			table.Controls.Add(clear, 2, row);
			boxes[file] = box;
			row++;
		}
		scroller.Controls.Add(table);
		layout.Controls.Add(scroller);
		layout.Controls.Add(warning);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		Check();
	}

	void Check() {
		var messages = new List<string>();
		foreach (var group in boxes.Where(p => p.Value.Hotkey != Keys.None).GroupBy(p => p.Value.Hotkey)) {
			string keys = TKeys(HotkeyBox.Describe(group.Key));
			if (group.Count() > 1) messages.Add(T("dlg.macroHotkeys.twice", "{0} is used by more than one macro; only one of them will run.", keys));
			if (menuShortcuts.TryGetValue(group.Key, out var item)) messages.Add(T("dlg.macroHotkeys.replacesShortcut", "{0} replaces the menu shortcut of {1}.", keys, item));
		}
		warning.Text = string.Join("\n", messages);
	}
}

// The editor control: RichEditBox with a hook for the right-click menu, so the spelling menu can come before the
// library's formatting toolbar. The hook gets the screen point and the character position, and returns true if it
// showed a menu.
// The grey around the page (Print Layout), which scrolls sideways when the page is wider than the window. Iteration 46:
// it stays where it is scrolled to when the editor gets the focus. WinForms scrolls a panel to bring a control that gets
// the focus into view, and so put the editor's left edge at the window's: with line numbers, the numbers and the left
// margin (beside the editor) went out of view.
sealed class PageHost : Panel
{
	protected override Point ScrollToControl(Control activeControl) => AutoScrollPosition;
}

// The panel around the editor (Show Pages draws the bands in its margins): painted through a bitmap, so the grey isn't
// first painted white (iteration 43).
sealed class BufferedPanel : Panel
{
	public BufferedPanel() => DoubleBuffered = true;
}

sealed class WriterEditBox : RichEditBox
{
	public Func<Point, int, bool>? ContextMenuHook;

	/// Ctrl+click on a link (text or picture): the link's address, to be opened.
	public Action<string>? OpenLink;

	/// Ctrl+click on a cross-reference (iteration 50): its position, to go to the paragraph it refers to; true if it was one.
	public Func<int, bool>? FollowReference;
	/// The pointer over the text while the Format Painter is on (null: RichEdit's own).
	public Cursor? PaintCursor;

	// ---- (iteration 92) a block of table cells (rte_tablecells.cs)
	(int From, int To)? cellBlock;
	int blockCaret = -1, dragFrom = -1, dragTo = -1;

	/// Table cells selected as a block, AhkWord's own selection: dragging from a cell in one row into another row of the
	/// same table (RichEdit selects the whole rows then), or Rte.SelectCells. Positions of a character in the first and
	/// in the last cell; the cells are shaded as selected, RichEdit's own selection a caret in the first; Merge, Split and
	/// Delete Cells, Delete Row and Delete Column act on them. Null: none (RichEdit's selection is the selection). Typing,
	/// clicking or moving the caret ends it.
	public (int From, int To)? CellBlock {
		get => cellBlock;
		set { if (cellBlock == value) return; cellBlock = value; if (IsHandleCreated) Invalidate(); }
	}

	/// Delete or Backspace with a block of cells: their text emptied (set by the form).
	public Action? ClearCellBlock;

	/// The edges of the block's table's cells, in twips (RtfTables.EdgesOf), for painting it.
	public List<List<int>>? BlockEdges { get; private set; }

	/// The block of cells from one position's cell to another's, RichEdit's selection a caret in the first.
	public void SelectBlock(int from, int to) {
		cellBlock = null;
		BlockEdges = RtfTables.EdgesOf(this, from);
		TextDocument.Selection2.SetRange(from, from);
		blockCaret = TextDocument.Selection2.Start;
		CellBlock = (from, to);
	}

	// (iteration 93) Shift+Up/Down from a table cell into another row: a block of cells grows or shrinks by a cell, as in
	// Word (RichEdit selects the whole rows); Shift+Left/Right too while there is a block. False: RichEdit's own.
	bool ExtendCellBlock(Keys key) {
		int from, to;
		if (cellBlock is { } block) (from, to) = block;
		else {
			if (key is Keys.Left or Keys.Right) return false;
			var selection = TextDocument.Selection2;
			from = selection.Start;
			to = Math.Max(selection.Start, selection.End - 1);
			if (!RtfTables.SameCell(this, from, to) && !(RtfTables.At(this, from) is { } a && RtfTables.At(this, to) is { } b && a.Table.Start == b.Table.Start && a.Row == b.Row)) return false;
		}
		int? next;
		try { next = RtfTables.NextCell(this, to, key, cellBlock != null ? BlockEdges : RtfTables.EdgesOf(this, from)); }
		catch (Exception ex) when (ex is COMException or ArgumentException) { return false; }
		if (next is not int at) return cellBlock != null;    // (at the table's edge: the block stays)
		if (RtfTables.SameCell(this, from, at)) {
			CellBlock = null;
			TextDocument.Selection2.SetRange(from, from);
		}
		else SelectBlock(from, at);
		return true;
	}

	protected override void OnKeyDown(KeyEventArgs e) {
		if (e.Shift && !e.Control && !e.Alt && e.KeyCode is Keys.Up or Keys.Down or Keys.Left or Keys.Right && ExtendCellBlock(e.KeyCode)) {
			e.Handled = e.SuppressKeyPress = true;
			return;
		}
		if (cellBlock != null) {
			if (e.KeyCode is Keys.Delete or Keys.Back && e.Modifiers == Keys.None && ClearCellBlock != null) {
				e.Handled = e.SuppressKeyPress = true;
				ClearCellBlock();
				return;
			}
			if (e.KeyCode is not (Keys.ShiftKey or Keys.ControlKey or Keys.Menu)) CellBlock = null;
		}
		base.OnKeyDown(e);
	}

	[StructLayout(LayoutKind.Sequential)]
	struct IconInfo { public bool Icon; public int HotX, HotY; public IntPtr Mask, Color; }
	[DllImport("user32.dll")] static extern bool GetIconInfo(IntPtr icon, out IconInfo info);
	[DllImport("user32.dll")] static extern IntPtr CreateIconIndirect(ref IconInfo info);
	[DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr icon);

	/// A cursor from a picture (its transparency kept), with its hot spot at x, y.
	public static Cursor CursorFrom(Bitmap picture, int x, int y) {
		IntPtr icon = picture.GetHicon();
		GetIconInfo(icon, out var info);
		DestroyIcon(icon);
		info.Icon = false;
		(info.HotX, info.HotY) = (x, y);
		IntPtr cursor = CreateIconIndirect(ref info);
		DeleteObject(info.Mask);
		DeleteObject(info.Color);
		return new Cursor(cursor);
	}

	// The right button: a press outside the selection moves the caret there, as in Word (inside it, the selection stays
	// for the menu's Cut and Copy). Iteration 51: as RichEdit's own hit-testing has it (TOM), as the menu does; the
	// library's (WinForms' character index, short of TOM's past hidden text such as footnotes' notes) put the caret
	// somewhere else, and the menu then came up for what was there.
	protected override void OnMouseDown(MouseEventArgs e) {
		// (iteration 92) the right button on a block of cells: it stays, for the menu
		if (e.Button == MouseButtons.Right && cellBlock is { } block && RtfTables.BlockRectangles(this, block.From, block.To) is { } cells && cells.Any(r => r.Contains(e.Location))) {
			if (!Focused) Focus();
			return;
		}
		if (e.Button != MouseButtons.Right) { base.OnMouseDown(e); return; }
		if (!Focused) Focus();
		var selection = TextDocument.Selection2;
		int position = PositionAt(PointToScreen(e.Location));
		if (selection.End == selection.Start || position < selection.Start || position > selection.End) selection.SetRange(position, position);
	}

	protected override void OnMouseUp(MouseEventArgs e) {
		base.OnMouseUp(e);
		if (e.Button != MouseButtons.Left || ModifierKeys != Keys.Control) return;
		int position = PositionAt(PointToScreen(e.Location));
		if (FollowReference?.Invoke(position) == true) return;
		if (OpenLink == null) return;
		// the character clicked can be the one after the point (a click on its right half): both are looked at
		foreach (int at in new[] { position, position - 1 }) {
			if (at < 0) continue;
			string url = LinkAddress(TextDocument.Range2(at, at + 1).URL);
			if (url != "") { OpenLink(url); return; }
		}
	}

	// The address in a link's field instruction: "https://..." (TOM gives it in quotes, Word's with switches after it).
	public static string LinkAddress(string? url) {
		if (string.IsNullOrWhiteSpace(url)) return "";
		var quoted = System.Text.RegularExpressions.Regex.Match(url, "\"([^\"]+)\"");
		return (quoted.Success ? quoted.Groups[1].Value : url).Trim();
	}

	/// The alt text of the picture at a position, for the tooltip (set by the form).
	public Func<int, string>? AltTextOf;

	readonly ToolTip pictureTip = new() { InitialDelay = 400, ReshowDelay = 100 };
	int tipPicture = -1;

	/// The position of the picture under a screen point, or -1.
	public int PictureAt(Point screen) {
		int position = PositionAt(screen);
		foreach (int at in new[] { position, position - 1 }) {
			if (at < 0) continue;
			try {
				var range = TextDocument.Range2(at, at + 1);
				if (range.Text != "\uFFFC") continue;
				range.GetPoint(0x20, out int left, out int top);       // tomStart, top left (screen)
				range.GetPoint(0x0A, out int right, out int bottom);   // tomEnd, bottom right
				if (screen.X >= left && screen.X <= right && screen.Y >= top && screen.Y <= bottom) return at;
			} catch (Exception ex) when (ex is COMException or ArgumentException) { }
		}
		return -1;
	}

	/// The note of the footnote whose number is at a position (iteration 51; null: none), for the tooltip.
	public Func<int, string?>? NoteAt;
	string? tipNote;

	protected override void OnMouseMove(MouseEventArgs e) {
		base.OnMouseMove(e);
		int picture = e.Button == MouseButtons.None && AltTextOf != null ? PictureAt(PointToScreen(e.Location)) : -1;
		// (a footnote's number: its note)
		if (picture < 0 && e.Button == MouseButtons.None && NoteAt != null) {
			string? note = null;
			try {
				int at = PositionAt(PointToScreen(e.Location));
				note = NoteAt(at) ?? (at > 0 ? NoteAt(at - 1) : null);
			} catch (Exception ex) when (ex is COMException or ArgumentException) { }
			if (note != tipNote) {
				tipNote = note;
				tipPicture = -1;
				if (note == null) pictureTip.Hide(this);
				else pictureTip.Show(note + "\n(Ctrl+click to change the footnote)", this, e.X + 12, e.Y + 20, 15000);
			}
			if (note != null) return;
		}
		if (picture == tipPicture) return;
		tipPicture = picture;
		if (picture < 0) { pictureTip.Hide(this); return; }
		string alt = "";
		try { alt = AltTextOf!(picture); } catch (Exception ex) when (ex is ArgumentException or COMException) { }
		pictureTip.Show((alt != "" ? "Alt text: " + alt : "No alt text") + "\n(right-click > Alt text... to change it)", this, e.X + 12, e.Y + 20, 8000);
	}

	protected override void OnMouseLeave(EventArgs e) {
		base.OnMouseLeave(e);
		tipPicture = -1;
		pictureTip.Hide(this);
	}

	int PositionAt(Point screen) {
		try { return TextDocument.RangeFromPoint2(screen.X, screen.Y, 0).Start; }
		catch (COMException) { return GetCharIndexFromPosition(PointToClient(screen)); }
	}

	// WinForms hands RichEdit its Font in pixels, so 11 pt became 11.25 pt (and 11.5 in RTF, which keeps half points):
	// new text gets the Font's own size, in points.
	protected override void OnHandleCreated(EventArgs e) {
		base.OnHandleCreated(e);
		SetDefaultSize();
		ApplyTextMargins();
	}

	/// Empties the editor and gives it a new RichEdit window, as fresh as at the start (the zoom kept). (2026-11-08) In
	/// the focused editor a UI Automation client (the touch keyboard's TextInputHost, Copilot, text services) takes text
	/// ranges as the text changes and never lets them go; RichEdit moves every live range on every insertion, so each
	/// edit cost more as the session went on: a document opened a third time took seconds longer, typing slowed too.
	/// Neither emptying the text nor UiaDisconnectAllProviders gave the time back; a new window does (the old one's
	/// ranges die with it). TOM ranges kept from before are dead afterwards: their holders drop them first.
	public void Renew() {
		if (!IsHandleCreated) return;
		int numerator = 0, denominator = 0;
		SendZoom(Handle, EM_GETZOOM, ref numerator, ref denominator);
		Text = "";    // (what WinForms carries over into the new window, as RTF: nothing)
		Renewing = true;
		try { RecreateHandle(); }
		finally { Renewing = false; }
		if (numerator > 0 && denominator > 0) SendZoom(Handle, EM_SETZOOM, (IntPtr)numerator, (IntPtr)denominator);
		pageBreaks = null;
		sectionMarks = null;
	}
	/// Renew is making the new window (HandleCreated handlers that its caller does itself can skip theirs).
	public bool Renewing { get; private set; }

	/// RichEdit's advanced layout, on or off. RichEdit takes it itself on reading RTF that has right or centre tabs or
	/// leaders; without it such tabs act as left ones and leaders aren't drawn. Only for documents that have them: typing
	/// flickers in it, with tables and Show Pages (AhkWord 68 had it off: tested by the user).
	public void AdvancedLayout(bool on) {
		if (IsHandleCreated) SendTypography(Handle, EM_SETTYPOGRAPHYOPTIONS, (IntPtr)(on ? TO_ADVANCEDTYPOGRAPHY : 0), (IntPtr)TO_ADVANCEDTYPOGRAPHY);
	}

	protected override void OnFontChanged(EventArgs e) {
		base.OnFontChanged(e);
		if (IsHandleCreated) { SetDefaultSize(); ApplyTextMargins(); }
	}

	/// The space left and right of the text inside the editor (pixels). Iteration 44: in Print Layout the page's margins
	/// are the editor's own, so that what is drawn in them (Show Pages' bands) moves with the text when RichEdit scrolls,
	/// in the same move: drawn beside the editor, they followed it a moment later, a flicker at the text's edges.
	public (int Left, int Right) TextMargins {
		get => textMargins;
		set { if (value == textMargins && appliedZoom == CurrentZoom()) return; textMargins = value; ApplyTextMargins(); Invalidate(); }
	}
	(int Left, int Right) textMargins;
	float appliedZoom = -1;
	float CurrentZoom() { try { return Math.Max(0.1f, ZoomFactor); } catch (Exception) { return 1f; } }
	// (RichEdit zooms its margins: it is given them unzoomed; set again when the zoom changes, by the page's layout)
	void ApplyTextMargins() {
		if (!IsHandleCreated) return;
		float zoom = appliedZoom = CurrentZoom();
		int left = (int)Math.Round(textMargins.Left / zoom), right = (int)Math.Round(textMargins.Right / zoom);
		SendScroll(Handle, EM_SETMARGINS, (IntPtr)3, (IntPtr)((right << 16) | (left & 0xFFFF)));    // EC_LEFTMARGIN | EC_RIGHTMARGIN
	}
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendScroll(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	const int EM_SETMARGINS = 0x00D3;

	/// Where the text goes across the window (inside TextMargins): what is drawn with it (headers and footers, page lines)
	/// stays within it.
	public Rectangle TextArea => Rectangle.FromLTRB(textMargins.Left, 0, Math.Max(textMargins.Left, ClientSize.Width - textMargins.Right), ClientSize.Height);

	public void SetDefaultSize() {
		var format = new CharFormat2 { cbSize = Marshal.SizeOf<CharFormat2>(), dwMask = 0x80000000u, yHeight = (int)Math.Round(Font.SizeInPoints * 20) };    // CFM_SIZE, in twips
		SendCharFormat(Handle, 0x0444, IntPtr.Zero, ref format);    // EM_SETCHARFORMAT, SCF_DEFAULT
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	struct CharFormat2
	{
		public int cbSize;
		public uint dwMask, dwEffects;
		public int yHeight, yOffset, crTextColor;
		public byte bCharSet, bPitchAndFamily;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string szFaceName;
		public short wWeight, sSpacing;
		public int crBackColor, lcid, dwReserved;
		public short sStyle, wKerning;
		public byte bUnderlineType, bAnimation, bRevAuthor, bUnderlineColor;
	}

	[DllImport("user32.dll", EntryPoint = "SendMessageW")]
	static extern IntPtr SendCharFormat(IntPtr hWnd, int msg, IntPtr wParam, ref CharFormat2 lParam);

	// ---- page breaks: RichEdit shows them as empty lines, so a dashed line with "Page break" is drawn over them

	List<int>? pageBreaks;    // their positions; null: to be found again

	protected override void OnTextChanged(EventArgs e) {
		base.OnTextChanged(e);
		pageBreaks = null;
		sectionMarks = null;
		CellBlock = null;    // (iteration 92: its positions are gone)
		RepaintDrawings();
	}

	protected override void OnSelectionChanged(EventArgs e) {
		base.OnSelectionChanged(e);
		if (cellBlock != null && IsHandleCreated) {
			var selection = TextDocument.Selection2;
			if (selection.Start != blockCaret || selection.End != blockCaret) CellBlock = null;    // (iteration 92: the caret moved)
		}
		RepaintDrawings();
	}

	// RichEdit redraws the lines it changes itself (typing, a format applied), not through WM_PAINT: what is drawn over
	// them (a gap's footer, band and header, a page's line) would be gone until the next full paint. So once it has
	// done that, the parts of the window with drawings are painted again, which draws them again.
	bool repaintPending;
	void RepaintDrawings() {
		if (repaintPending || !IsHandleCreated) return;
		repaintPending = true;
		BeginInvoke(() => {
			repaintPending = false;
			if (!IsHandleCreated) return;
			var client = ClientRectangle;
			if (Gaps is { Applied: true } gaps) {
				float ppp = PixelsPerPoint;
				foreach (var gap in gaps.Gaps) {
					var (footer, _, header) = gaps.Areas(gap, client.Width, ppp);
					if (header.Bottom < client.Top || footer.Top > client.Bottom) continue;
					Invalidate(Rectangle.FromLTRB(0, (int)Math.Floor(footer.Top), client.Width, (int)Math.Ceiling(header.Bottom) + 1), false);
				}
				if (LastSheetEnd() is { } end && end.Top < client.Bottom)
					Invalidate(Rectangle.FromLTRB(0, (int)Math.Floor(Math.Max(0, end.Top)), client.Width, client.Bottom), false);
			}
			if (pageStarts is { Count: > 1 } starts)
				foreach (int at in starts.Skip(1)) {
					int y = GetPositionFromCharIndex(at).Y;
					if (y > client.Bottom + 12) break;
					if (y >= client.Top - 12) Invalidate(new Rectangle(0, y - 12, client.Width, 24), false);
				}
			foreach (var (at, _) in sectionMarks ?? new()) {
				int y = GetPositionFromCharIndex(at).Y;
				if (y >= client.Top - 12 && y <= client.Bottom + 12) Invalidate(new Rectangle(0, y - 12, client.Width, 24), false);
			}
			// (iteration 87) the framed paragraphs in view
			foreach (var place in FramePlaces()) {
				var shaded = Rectangle.Round(ParagraphFrames.ShadeArea(place, PixelsPerPoint));
				shaded.Inflate(6, 6);
				Invalidate(shaded, false);
			}
		});
	}

	/// Draws the page breaks in view (after RichEdit has painted).
	public void DrawPageBreaks(Graphics g) {
		pageBreaks ??= PageBreaks.Positions(TextDocument);
		if (pageBreaks.Count == 0) return;
		var client = TextArea;
		int storyEnd = TextDocument.Range2(0, 0).StoryLength;
		int lineHeight = Math.Max(8, Font.Height);
		using var pen = new Pen(Color.FromArgb(150, 150, 150)) { DashStyle = DashStyle.Dash };
		using var label = new Font("Segoe UI", 8f);
		using var brush = new SolidBrush(Color.FromArgb(120, 120, 120));
		using var back = new SolidBrush(BackColor);
		foreach (int at in pageBreaks) {
			int top = GetPositionFromCharIndex(at).Y;
			if (top > client.Bottom) break;
			// the break's line and its paragraph mark's (an empty line) make one band, with the dashed line in its middle
			int after = Math.Min(storyEnd - 1, TextDocument.Range2(at + 1, at + 2).Text == "\r" ? at + 2 : at + 1);
			int bottom = after > at ? GetPositionFromCharIndex(after).Y : top + lineHeight;
			if (after >= storyEnd - 1 || bottom <= top) bottom = top + 2 * lineHeight;
			if (bottom < client.Top) continue;
			// (at the end of a line of text, where joining paragraphs leaves one: under that line, not through it)
			int y = at > 0 && TextDocument.Range2(at - 1, at).Text != "\r" ? bottom : (top + bottom) / 2;
			g.DrawLine(pen, client.Left + 4, y, client.Right - 4, y);
			string text = T("ui.pageBreak", "Page break");
			var size = g.MeasureString(text, label);
			float x = client.Left + (client.Width - size.Width) / 2;
			g.FillRectangle(back, x - 4, y - size.Height / 2, size.Width + 8, size.Height);
			g.DrawString(text, label, brush, x, y - size.Height / 2);
		}
	}

	// (iteration 67) continuous section breaks take no room: a thin dashed line above the section's first line, with its
	// columns at the right
	List<(int At, string Text)>? sectionMarks;
	int sectionMarksLength = -1;    // (the document's length they were found for: not all edits raise TextChanged)
	bool SectionMarksStale => sectionMarks == null || sectionMarksLength != TextDocument.Range2(0, 0).StoryLength;

	/// A section's settings changed (the document's length may not have): its lines found and drawn again.
	public void SectionsChanged() {
		sectionMarks = null;
		if (IsHandleCreated) Invalidate();
	}

	/// The lines drawn at the last paint (for tests).
	internal string SectionMarksDrawn => string.Join("; ", (sectionMarks ?? new()).Select(m => m.Text));
	void DrawSectionMarks(Graphics g) {
		if (SectionMarksStale) {
			sectionMarksLength = TextDocument.Range2(0, 0).StoryLength;
			sectionMarks = Sections.Breaks(TextDocument).Where(b => b.Settings.Continuous)
				.Select(b => (b.SectionStart, b.Settings.Columns is { } c ? $"Section: {c.Count} columns" : "Section: one column")).ToList();
		}
		if (sectionMarks!.Count == 0) return;
		var client = TextArea;
		using var pen = new Pen(Color.FromArgb(150, 150, 150)) { DashStyle = DashStyle.Dot };
		using var label = new Font("Segoe UI", 7.5f);
		using var brush = new SolidBrush(Color.FromArgb(120, 120, 120));
		using var back = new SolidBrush(BackColor);
		foreach (var (at, text) in sectionMarks) {
			int y = GetPositionFromCharIndex(at).Y;
			if (y < client.Top || y > client.Bottom) continue;
			g.DrawLine(pen, client.Left + 4, y, client.Right - 4, y);
			var size = g.MeasureString(text, label);
			float x = client.Right - size.Width - 8;
			g.FillRectangle(back, x, y - size.Height / 2, size.Width, size.Height);
			g.DrawString(text, label, brush, x, y - size.Height / 2);
		}
	}

	/// Where the printed pages start (character positions; null: none shown): Print Layout draws a line where each
	/// page begins, with its number (the editor's lines are the printed ones there, but RichEdit has no pages).
	public List<int>? PageStarts {
		get => pageStarts;
		set { if (value == null ? pageStarts == null : pageStarts != null && value.SequenceEqual(pageStarts)) return; pageStarts = value; Invalidate(); }
	}
	List<int>? pageStarts;

	/// The page view's gaps (rte_pagegaps.cs), and the header's and footer's parts for a page (null: not shown).
	public PageGaps? Gaps;
	public Func<bool, int, string[]?>? BandParts;
	/// (iteration 55) the header's (true) or footer's font (null: Segoe UI in grey, as before) and its distance from the
	/// sheet's edge (1/100 inch).
	public Func<bool, int, (BandFont? Font, int Distance, bool Line)>? BandStyleOf;

	/// Where a character is (its line's top), from RichEdit itself, in view or not. RichTextBox's own gives (0, 0) for
	/// a character past its Text's length, which can be shorter than the document (a Word document's last page): what
	/// is drawn there (a page's line, a footer) would then stay at the window's top while the text scrolls under it.
	/// It also reads the whole text for each call.
	public new Point GetPositionFromCharIndex(int index) {
		if (!IsHandleCreated) return Point.Empty;
		int storyEnd = TextDocument.Range2(0, 0).StoryLength;
		var point = new Point();
		SendPosition(Handle, 0x00D6, ref point, (IntPtr)Math.Clamp(index, 0, Math.Max(0, storyEnd - 1)));    // EM_POSFROMCHAR
		return point;
	}
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendPosition(IntPtr hWnd, int msg, ref Point point, IntPtr cp);
	/// Screen pixels per point at the editor's zoom.
	public float PixelsPerPoint => DeviceDpi / 72f * Math.Max(0.1f, ZoomFactor);

	/// Draws the page view's gaps in view: the sheet's end with its footer, a grey band, the next sheet's top with its
	/// header; page 1's header above the first line.
	public void DrawGaps(Graphics g) {
		if (Gaps is not { Applied: true } gaps) return;
		var client = ClientRectangle;
		float ppp = PixelsPerPoint, width = client.Width;
		using var grey = new SolidBrush(PrintLayoutPage.Surround);
		using var edge = new Pen(Color.FromArgb(200, 200, 200));
		using var ink = new SolidBrush(Color.FromArgb(105, 105, 105));
		float zoom = ppp / (DeviceDpi / 72f);
		using var black = new SolidBrush(Color.FromArgb(30, 30, 30));
		// (iteration 55) the header that far below the sheet's top, the footer's last line that far above its bottom, a
		// line below the other; in the page's font (the first page can have its own)
		void Parts(string[]? parts, RectangleF area, bool isHeader, int page) {
			if (parts is not { Length: 3 } || area.Height <= 0) return;
			var (own, distance, line) = BandStyleOf?.Invoke(isHeader, page) ?? (null, PageBands.DefaultDistance, false);
			var band = own ?? BandFont.Plain;
			float from = distance * 0.72f * ppp;
			float height = BandText.Height(parts, band, f => f.GetHeight(g), zoom);
			float top = isHeader ? area.Top + from : area.Bottom - from - height;
			float left = textMargins.Left, textWidth = Math.Max(1, width - textMargins.Left - textMargins.Right);    // (above and below the text, not in the margins)
			BandText.Draw(g, parts, band, left, textWidth, top, zoom, PageBands.IsPlain(own) ? ink : black, black);
			// (iteration 55) the line below the header, above the footer
			if (line && PageBands.LineCount(parts) > 0) {
				float gap = PageBands.RuleGap * 0.72f * ppp, y = isHeader ? top + height + gap : top - gap;
				using var pen = new Pen(black, Math.Max(1, PageBands.RuleWidth * 0.72f * ppp));
				g.DrawLine(pen, left, y, left + textWidth, y);
			}
		}
		// (iteration 51) a sheet's footnotes, at the foot of its text area (above its bottom margin): the lines the printed
		// page has, in its notes' font at the zoom
		Font? noteFont = null;
		void Notes(int page, float bottom) {
			if (NotesOn?.Invoke(page) is not { } notes || notes.Lines.Count == 0) return;
			noteFont ??= Footnotes.NoteFont(notes.Font, notes.Size * Math.Max(0.1f, ZoomFactor));
			float textWidth = Math.Max(1, width - textMargins.Left - textMargins.Right);
			float height = Footnotes.Height(g, notes.Lines, noteFont);
			if (bottom < client.Top || bottom - height > client.Bottom) return;
			Footnotes.Draw(g, notes.Lines, new RectangleF(textMargins.Left, bottom - height, textWidth, height), noteFont, Color.FromArgb(30, 30, 30));
		}
		// (iteration 51) the gap blank first, the page's paper: RichEdit draws a selection over the whole of a line, and a
		// gap is part of one (the next paragraph's space before, or a raised line's foot), so selecting the next page's
		// first line painted the sheet's foot, the band and the next sheet's top blue
		using var paper = new SolidBrush(BackColor);
		try {
		foreach (var gap in gaps.Gaps) {
			var (footer, band, header) = gaps.Areas(gap, width, ppp);
			if (header.Bottom < client.Top || footer.Top > client.Bottom) continue;
			g.FillRectangle(paper, RectangleF.FromLTRB(0, MathF.Ceiling(footer.Top), width, MathF.Floor(header.Bottom)));
			if (gap.Page == 0) { Parts(BandParts?.Invoke(true, 1), header, true, 1); continue; }    // (page 1's top margin: its header)
			Notes(gap.Page, band.Top - gap.Bottom * ppp);
			Parts(BandParts?.Invoke(false, gap.Page), new RectangleF(0, band.Top - gap.Bottom * ppp, width, gap.Bottom * ppp), false, gap.Page);
			g.FillRectangle(grey, band);
			g.DrawLine(edge, band.Left, band.Top, band.Right, band.Top);
			g.DrawLine(edge, band.Left, band.Bottom - 1, band.Right, band.Bottom - 1);
			Parts(BandParts?.Invoke(true, gap.Page + 1), header, true, gap.Page + 1);
		}
		// (iteration 46) the last sheet's end below the text: its footer, grey after it
		if (LastSheetEnd() is { } end && end.Top < client.Bottom + end.Height * 4 && end.Bottom > client.Top - 1) {
			Notes(gaps.Pages, end.Top);
			Parts(BandParts?.Invoke(false, gaps.Pages), end, false, gaps.Pages);
			if (end.Bottom < client.Bottom) {
				g.FillRectangle(grey, RectangleF.FromLTRB(0, end.Bottom, width, client.Bottom));
				g.DrawLine(edge, 0, end.Bottom, width, end.Bottom);
			}
		}
		} finally {
			noteFont?.Dispose();
		}
	}

	/// Iteration 51: a page's footnotes (1: the first page), for its sheet; null or none: none.
	public Func<int, (List<Footnotes.NoteLine> Lines, string Font, float Size)?>? NotesOn;

	/// Iteration 46: how much further the view may scroll down past the text's end (pixels): until the last sheet's
	/// footer, and a little grey after it, are in view (0: no further).
	public int RoomBelow() {
		if (LastSheetEnd() is not { } end) return 0;
		return Math.Max(0, (int)Math.Ceiling(end.Bottom + 18 * PixelsPerPoint - ClientSize.Height));
	}

	/// The last sheet's end (PageGaps.LastSheetEnd), its blank rest no taller than the view can scroll past the text to
	/// show it whole with its footer (null: no gaps).
	public RectangleF? LastSheetEnd() {
		if (Gaps is not { Applied: true } gaps) return null;
		float ppp = PixelsPerPoint;
		return gaps.LastSheetEnd(ClientSize.Width, ppp, ClientSize.Height - (gaps.BottomPoints + 18) * ppp - 40 * ppp);
	}

	/// Draws the boundaries of the printed pages in view: a thin line above each page's first line, and "Page N" at
	/// the right (a page that starts after a page break has that break's line already, and gets only its number).
	/// Pages the page view has a gap before are left out.
	public void DrawPageBoundaries(Graphics g) {
		if (pageStarts is not { Count: > 1 } starts) return;
		var covered = Gaps is { Applied: true } withGaps ? withGaps.Gaps.Where(x => x.Page > 0).Select(x => x.Next).ToHashSet() : new HashSet<int>();
		var client = TextArea;
		int storyEnd = TextDocument.Range2(0, 0).StoryLength;
		string story = "";
		using var pen = new Pen(Color.FromArgb(150, 170, 200)) { DashStyle = DashStyle.Dot };
		using var label = new Font("Segoe UI", 7.5f);
		using var brush = new SolidBrush(Color.FromArgb(110, 130, 160));
		using var back = new SolidBrush(BackColor);
		for (int n = 1; n < starts.Count; n++) {
			int at = starts[n];
			if (at <= 0 || at >= storyEnd - 1 || covered.Contains(at)) continue;
			int y = GetPositionFromCharIndex(at).Y;
			if (y > client.Bottom) break;
			if (y < client.Top - 2) continue;
			if (story == "") story = TextDocument.Range2(0, storyEnd).Text ?? "";
			bool afterBreak = at >= 2 && story[at - 2] == PageBreaks.Break;    // (the page break's own line is there)
			int line = Math.Max(client.Top, y - 1);
			if (!afterBreak) g.DrawLine(pen, client.Left + 4, line, client.Right - 4, line);
			string text = T("ui.pageNumber", "Page {0}", n + 1);
			var size = g.MeasureString(text, label);
			float x = client.Right - size.Width - 8;
			float top = Math.Max(client.Top, line - size.Height / 2);
			g.FillRectangle(back, x - 3, top, size.Width + 6, size.Height);
			g.DrawString(text, label, brush, x, top);
		}
	}

	/// After the editor painted (and so scrolled, changed or zoomed): the line numbers follow.
	public Action? Painted;

	/// Whether the editor paints through a bitmap when something is drawn over its text (false: RichEdit paints on the
	/// screen, then the drawings go over it there).
	public bool Buffered = true;

	/// While true, the editor doesn't paint: a paint asked for is left for later (the page view's gaps going in).
	public bool HoldPaint;
	public bool Typing;    // (RichEdit is handling a key without Ctrl or Alt: what it changes is typing)

	// Iteration 43: RichEdit paints the part of the window to paint into a bitmap (WM_PRINTCLIENT), the gaps, page breaks
	// and page lines are drawn over it there, and the bitmap goes onto the screen at once: drawn over on the screen, the
	// white RichEdit paints there showed for a moment first (a flicker as the caret moved past a sheet's end).
	void PaintBuffered() {
		var paint = new PaintStruct();
		IntPtr screen = BeginPaint(Handle, ref paint);
		// (iteration 51) RichEdit not done laying out a change yet (Tab on a numbered paragraph, a footnote inserted): it
		// would leave those lines out of the bitmap (black on the screen, a moment): the screen kept as it is, and painted
		// again a little later (a few times at most, then as it comes)
		if (pendingPaints < 8 && LayoutPending()) {
			EndPaint(Handle, ref paint);
			pendingPaints++;
			pendingTimer ??= MakePendingTimer();
			pendingTimer.Stop();
			pendingTimer.Start();
			return;
		}
		pendingPaints = 0;
		try {
			var area = Rectangle.FromLTRB(paint.Left, paint.Top, paint.Right, paint.Bottom);
			area.Intersect(ClientRectangle);
			if (area.Width > 0 && area.Height > 0)
				Render(screen, area, memory => BitBlt(screen, area.Left, area.Top, area.Width, area.Height, memory, 0, 0, SRCCOPY));
		}
		finally { EndPaint(Handle, ref paint); }
	}

	int pendingPaints;
	System.Windows.Forms.Timer? pendingTimer;

	System.Windows.Forms.Timer MakePendingTimer() {
		var timer = new System.Windows.Forms.Timer { Interval = 20 };
		timer.Tick += (s, e) => { timer.Stop(); if (IsHandleCreated) Invalidate(); };
		return timer;
	}

	// Whether RichEdit is still laying text out (TOM's E_PENDING) at the caret or the top of the window.
	bool LayoutPending() {
		if (!IsHandleCreated) return false;
		try {
			foreach (int at in new[] { TextDocument.Selection2.Start, GetCharIndexFromPosition(new Point(1, 1)) })
				TextDocument.Range2(at, at).GetPoint(0x20 | 0x200, out _, out _);    // tomStart | tomAllowOffClient
			return false;
		} catch (COMException ex) when (ex.HResult == unchecked((int)0x8000000A)) {
			return true;
		} catch (Exception ex) when (ex is COMException or ArgumentException) {
			return false;
		}
	}

	// The part area of the window as PaintBuffered puts it on the screen, in a bitmap compatible with dc, to done.
	void Render(IntPtr dc, Rectangle area, Action<IntPtr> done) {
		IntPtr memory = CreateCompatibleDC(dc), bitmap = CreateCompatibleBitmap(dc, area.Width, area.Height);
		IntPtr old = SelectObject(memory, bitmap);
		try {
			// (the window's own background first, not the new bitmap's black: lines RichEdit leaves out stay blank)
			using (var g = Graphics.FromHdc(memory)) g.Clear(BackColor);
			SetViewportOrgEx(memory, -area.Left, -area.Top, IntPtr.Zero);
			SendPaint(Handle, WM_PRINTCLIENT, memory, (IntPtr)(PRF_CLIENT | PRF_ERASEBKGND));
			SetViewportOrgEx(memory, 0, 0, IntPtr.Zero);
			// (iteration 87) the paragraphs' shading under the text: what RichEdit painted darkened to the colour (white
			// becomes the colour, the text stays dark), and their lines over it
			var frames = FramePlaces();
			float ppp = PixelsPerPoint;
			foreach (var place in frames)
				ParagraphFrames.Draw(null!, place, ppp, 1, (rect, color) => {
					var r = Rectangle.Round(rect);
					IntPtr brush = CreateSolidBrush(color.R | color.G << 8 | color.B << 16), before = SelectObject(memory, brush);
					PatBlt(memory, r.Left - area.Left, r.Top - area.Top, r.Width, r.Height, 0x00A000C9);    // PATAND
					SelectObject(memory, before);
					DeleteObject(brush);
				}, lines: false);
			// (iteration 92) a block of table cells: shaded as RichEdit shades a selection, the text dark over it
			if (cellBlock is { } block && RtfTables.BlockRectangles(this, block.From, block.To) is { } cells) {
				IntPtr brush = CreateSolidBrush(0x99 | 0xC9 << 8 | 0xEF << 16), before = SelectObject(memory, brush);
				foreach (var r in cells) PatBlt(memory, r.Left - area.Left, r.Top - area.Top, r.Width, r.Height, 0x00A000C9);    // PATAND
				SelectObject(memory, before);
				DeleteObject(brush);
			}
			using (var g = Graphics.FromHdc(memory)) {
				g.TranslateTransform(-area.Left, -area.Top);
				foreach (var place in frames) ParagraphFrames.Draw(g, place, ppp, 1, fill: false);
				DrawDrawings(g);
			}
			done(memory);
		}
		finally {
			SelectObject(memory, old);
			DeleteObject(bitmap);
			DeleteDC(memory);
		}
	}

	/// The part area of the window as it is painted (tests).
	public Bitmap Snapshot(Rectangle area) {
		IntPtr dc = GetDC(Handle);
		try {
			Bitmap? shot = null;
			Render(dc, area, memory => {
				using var copy = new Bitmap(area.Width, area.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);    // (GDI leaves no alpha)
				using (var g = Graphics.FromImage(copy)) {
					IntPtr target = g.GetHdc();
					BitBlt(target, 0, 0, area.Width, area.Height, memory, 0, 0, SRCCOPY);
					g.ReleaseHdc(target);
				}
				shot = new Bitmap(copy);
			});
			return shot!;
		}
		finally { ReleaseDC(Handle, dc); }
	}
	[DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr hWnd);
	[DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hWnd, IntPtr hdc);

	/// What is drawn over the text: the page view's gaps (or the page breaks), the printed pages' boundaries.
	public void DrawDrawings(Graphics g) {
		DrawLeaders(g);
		if (Gaps is { Applied: true }) DrawGaps(g);
		else DrawPageBreaks(g);
		DrawSectionMarks(g);
		DrawPageBoundaries(g);
	}

	/// (iteration 87) The paragraphs with borders or shading (ParagraphFrames), drawn here and on the printed pages.
	public Func<IReadOnlyList<ParagraphFrames.Framed>>? Frames;

	/// (iteration 87) Where the framed paragraphs in view are (their lines' top and bottom, from the left indent, the
	/// first line's if further left, to the right indent), in the window.
	public List<ParagraphFrames.Place> FramePlaces() {
		var places = new List<ParagraphFrames.Place>();
		if (!IsHandleCreated || Frames?.Invoke() is not { Count: > 0 } frames) return places;
		const int Top = 0x20 | 0 | 0x100 | 0x200, Bottom = 0x20 | 8 | 0x100 | 0x200;    // tomStart | TA_TOP or TA_BOTTOM | tomClientCoord | tomAllowOffClient
		float ppp = PixelsPerPoint;
		var text = TextArea;
		int height = ClientSize.Height;
		var found = new List<(int Index, RectangleF Text)>();
		for (int i = 0; i < frames.Count; i++) {
			var f = frames[i];
			try {
				TextDocument.Range2(f.Start, f.Start).GetPoint(Top, out _, out int top);
				TextDocument.Range2(f.End, f.End).GetPoint(Bottom, out _, out int bottom);
				if (bottom < -50 || top > height + 50) continue;
				var para = TextDocument.Range2(f.Start, f.Start).Para;
				float left = text.Left + (para.LeftIndent + Math.Min(0, para.FirstLineIndent)) * ppp, right = text.Right - para.RightIndent * ppp;
				float last = bottom - para.SpaceAfter * ppp;    // (the last line's bottom holds the space after the paragraph)
				found.Add((i, RectangleF.FromLTRB(left, top, Math.Max(left + 1, right), Math.Max(top + 1, last))));
			} catch (Exception ex) when (ex is COMException or ArgumentException) { }
		}
		for (int k = 0; k < found.Count; k++) {
			var (i, rect) = found[k];
			bool above = k > 0 && found[k - 1].Index == i - 1 && ParagraphFrames.Joined(frames[i - 1], frames[i]);
			bool below = k + 1 < found.Count && found[k + 1].Index == i + 1 && ParagraphFrames.Joined(frames[i], frames[i + 1]);
			if (below) rect = RectangleF.FromLTRB(rect.Left, rect.Top, rect.Right, Math.Max(rect.Bottom, found[k + 1].Text.Top));
			places.Add(new ParagraphFrames.Place(frames[i], rect, above, below));
		}
		return places;
	}

	/// Iteration 52: the tables of contents' leaders, drawn here: RichEdit draws a tab's leader in its advanced layout
	/// (which it takes on reading RTF) but not always even then, so the entries' tabs have none (Toc.Format, FitTabs) and
	/// these are drawn instead, on the screen and printed (DocumentPrinter.LeadersOnPages).
	public Func<IReadOnlyList<Toc.Leader>>? Leaders;

	void DrawLeaders(Graphics g) {
		if (Leaders?.Invoke() is not { Count: > 0 } leaders) return;
		int bottom = ClientSize.Height;
		foreach (var leader in leaders) {
			if (LeaderLine(leader) is not { } line || line.Baseline < -line.Size || line.Baseline > bottom + line.Size) continue;
			Toc.DrawLeader(g, leader.Kind, line.X1, line.X2, line.Baseline, line.Size * PixelsPerPoint, ForeColor);
		}
	}

	/// Where a leader goes in the window: from its entry's text's end (its tab's start) to its number's start, on the
	/// number's baseline; the number's size (points). Null: the tab and the number aren't on one line, or not laid out.
	public (float X1, float X2, float Baseline, float Size)? LeaderLine(Toc.Leader leader) {
		const int Flags = 0x20 | 24 | 0x100 | 0x200;    // tomStart | TA_BASELINE | tomClientCoord | tomAllowOffClient
		try {
			var tab = TextDocument.Range2(leader.Tab, leader.Tab);
			tab.GetPoint(Flags, out int x1, out int y1);
			var number = TextDocument.Range2(leader.Number, leader.Number + 1);
			number.GetPoint(Flags, out int x2, out int y2);
			if (Math.Abs(y2 - y1) > 2 || x2 <= x1) return null;
			return (x1, x2, y2, Math.Max(1, number.Font.Size));
		} catch (Exception ex) when (ex is COMException or ArgumentException) {
			return null;
		}
	}

	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendTypography(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	const int EM_SETTYPOGRAPHYOPTIONS = 0x04CA, TO_ADVANCEDTYPOGRAPHY = 1;
	const int WM_PRINTCLIENT = 0x0318, PRF_CLIENT = 0x04, PRF_ERASEBKGND = 0x08, SRCCOPY = 0x00CC0020;
	const int WM_MOUSEWHEEL = 0x020A, WM_VSCROLL = 0x0115, WM_KEYDOWN = 0x0100, WM_CHAR = 0x0102, WM_TIMER = 0x0113, EM_GETSCROLLPOS = 0x04DD, EM_SETSCROLLPOS = 0x04DE;

	// Iteration 46: RichEdit scrolls only until the text's last line is at the window's bottom, but Show Pages draws the
	// last sheet's end (its footer) below it: the wheel and the scroll bar's down arrow and page go on down there (and
	// stay there: RichEdit would pull the view back up to its own end).
	bool ScrollPastEnd(Message m) {
		bool down = m.Msg switch {
			WM_MOUSEWHEEL => (short)((m.WParam.ToInt64() >> 16) & 0xFFFF) < 0 && (ModifierKeys & Keys.Control) == 0,
			WM_VSCROLL => (m.WParam.ToInt64() & 0xFFFF) is 1 or 3,    // SB_LINEDOWN, SB_PAGEDOWN
			_ => false
		};
		if (!down || Gaps is not { Applied: true } || !AtScrollEnd()) return false;
		int room = RoomBelow();
		if (room > 0) {
			int notches = m.Msg == WM_MOUSEWHEEL ? Math.Max(1, -(short)((m.WParam.ToInt64() >> 16) & 0xFFFF) / 120) : 1;
			int step = m.Msg == WM_VSCROLL && (m.WParam.ToInt64() & 0xFFFF) == 3 ? ClientSize.Height * 9 / 10 : notches * Math.Max(20, ClientSize.Height / 6);
			ScrollDown(Math.Min(room, step));
		}
		return true;
	}

	// Iteration 87: RichEdit animates the wheel (smooth scrolling: its timer moves the view on step by step after the
	// notch), and each step here is painted at once with the pages, their gaps, frames and footnotes. A fast spin made
	// a long train of steps: the view went on scrolling after the wheel stopped, and where the paint is slow (page 8 of
	// the fidelity document) the window stopped responding. So the wheel (without Ctrl or Shift, which zoom and scroll
	// sideways in RichEdit) scrolls here, at once as in Word: the notches waiting in the queue (with the same keys
	// held) taken with it, by the lines per notch Windows is set to (a page for "one screen"), one paint.
	void WheelScroll(int delta) {
		int lines = SystemInformation.MouseWheelScrollLines;
		double lineHeight = Font.GetHeight(DeviceDpi) * Math.Max(0.1f, ZoomFactor);
		int by = lines < 0 ? (int)((long)delta * (ClientSize.Height * 9 / 10) / 120)
			: (int)Math.Round(delta / 120.0 * lines * lineHeight);
		if (by == 0) return;
		var position = ScrollPosition();
		position.Y = Math.Max(0, position.Y - by);
		SendScroll(Handle, EM_SETSCROLLPOS, IntPtr.Zero, ref position);
		OnVScroll(EventArgs.Empty);    // (as RichEdit's own scrolling tells: the page overhang is measured again)
		Update();
	}

	// (the notches waiting in the queue, with the same keys held, the same way: one scroll)
	[StructLayout(LayoutKind.Sequential)] struct QueuedMessage { public IntPtr Hwnd; public int Msg; public IntPtr WParam, LParam; public int Time; public Point Point; public int Private; }
	[DllImport("user32.dll", EntryPoint = "PeekMessageW")] static extern bool PeekMessage(out QueuedMessage message, IntPtr hWnd, int first, int last, int remove);
	void CoalesceWheel(ref Message m) {
		const int PM_NOREMOVE = 0, PM_REMOVE = 1;
		long keys = m.WParam.ToInt64() & 0xFFFF;
		int delta = (short)((m.WParam.ToInt64() >> 16) & 0xFFFF);
		while (PeekMessage(out var next, Handle, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_NOREMOVE)) {
			int more = (short)((next.WParam.ToInt64() >> 16) & 0xFFFF);
			if ((next.WParam.ToInt64() & 0xFFFF) != keys || Math.Sign(more) != Math.Sign(delta) || Math.Abs(delta + more) > short.MaxValue) break;
			PeekMessage(out _, Handle, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE);
			delta += more;
		}
		m.WParam = (IntPtr)(((long)(ushort)(short)delta << 16) | keys);
	}

	// Whether the view is at RichEdit's own end of scrolling, or past it.
	bool AtScrollEnd() {
		var info = new ScrollInfo { Size = Marshal.SizeOf<ScrollInfo>(), Mask = 0x17 };    // SIF_ALL
		if (!GetScrollInfo(Handle, 1, ref info)) return false;    // (SB_VERT)
		return info.Page > 0 && ScrollPosition().Y >= info.Max - info.Page;
	}

	void ScrollDown(int by) {
		if (by <= 0) return;
		var position = ScrollPosition();
		position.Y += by;
		SendScroll(Handle, EM_SETSCROLLPOS, IntPtr.Zero, ref position);
		Update();
	}

	[StructLayout(LayoutKind.Sequential)] struct ScrollInfo { public int Size, Mask, Min, Max, Page, Pos, TrackPos; }
	[DllImport("user32.dll")] static extern bool GetScrollInfo(IntPtr hWnd, int bar, ref ScrollInfo info);
	Point ScrollPosition() {
		var point = new Point();
		SendScroll(Handle, EM_GETSCROLLPOS, IntPtr.Zero, ref point);
		return point;
	}
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendScroll(IntPtr hWnd, int msg, IntPtr wParam, ref Point lParam);
	[StructLayout(LayoutKind.Sequential)]
	struct PaintStruct {
		public IntPtr Hdc; public int Erase; public int Left, Top, Right, Bottom; public int Restore, IncUpdate;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] public byte[] Reserved;
	}
	[DllImport("user32.dll")] static extern IntPtr BeginPaint(IntPtr hWnd, ref PaintStruct paint);
	[DllImport("user32.dll")] static extern bool EndPaint(IntPtr hWnd, ref PaintStruct paint);
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendPaint(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
	[DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr hdc);
	[DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int width, int height);
	[DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);
	[DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr obj);
	[DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr hdc);
	[DllImport("gdi32.dll")] static extern bool SetViewportOrgEx(IntPtr hdc, int x, int y, IntPtr previous);
	[DllImport("gdi32.dll")] static extern bool BitBlt(IntPtr dest, int x, int y, int width, int height, IntPtr source, int sourceX, int sourceY, int rop);
	[DllImport("gdi32.dll")] static extern bool PatBlt(IntPtr hdc, int x, int y, int width, int height, int rop);
	[DllImport("gdi32.dll")] static extern IntPtr CreateSolidBrush(int color);

	// Replacing the text (WM_SETTEXT, as New and Open do through RichEditorNET; EM_STREAMIN, EM_SETTEXTEX) resets
	// RichEdit's zoom to 100%: it is put back, so a zoom (the user's, or Page Width's) outlives a new document.
	const int WM_SETTEXT = 0x000C, EM_STREAMIN = 0x0449, EM_SETTEXTEX = 0x0461, EM_GETZOOM = 0x04E0, EM_SETZOOM = 0x04E1;
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendZoom(IntPtr hWnd, int msg, ref int numerator, ref int denominator);
	[DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendZoom(IntPtr hWnd, int msg, IntPtr numerator, IntPtr denominator);

	protected override void WndProc(ref Message m) {
		const int WM_CONTEXTMENU = 0x007B, WM_PAINT = 0x000F, WM_SETCURSOR = 0x0020, HTCLIENT = 1;
		const int WM_LBUTTONDOWN = 0x0201, WM_MOUSEMOVE = 0x0200, WM_LBUTTONUP = 0x0202, MK_LBUTTON = 1;
		// (iteration 92) a drag with the left button from a table cell into another row of its table: those cells as a
		// block (RichEdit selects the whole rows). From the messages (their button state), not WinForms' mouse events,
		// which take the button's state from the mouse as it is now.
		if (m.Msg is WM_LBUTTONDOWN or WM_MOUSEMOVE or WM_LBUTTONUP && IsHandleCreated) {
			long lp = m.LParam.ToInt64();
			var point = PointToScreen(new Point((short)(lp & 0xFFFF), (short)((lp >> 16) & 0xFFFF)));
			try {
				if (m.Msg == WM_LBUTTONDOWN) {
					CellBlock = null;
					dragTo = -1;
					int at = PositionAt(point);
					dragFrom = RtfTables.At(this, at) != null ? at : -1;
				}
				else if (m.Msg == WM_MOUSEMOVE && dragFrom >= 0) {
					if ((m.WParam.ToInt64() & MK_LBUTTON) == 0) dragFrom = dragTo = -1;
					else { int at = PositionAt(point); dragTo = RtfTables.SpansRows(this, dragFrom, at) ? at : -1; }
				}
				else if (m.Msg == WM_LBUTTONUP && dragFrom >= 0) {
					int from = dragFrom, to = dragTo;
					dragFrom = dragTo = -1;
					base.WndProc(ref m);
					if (to >= 0) SelectBlock(from, to);
					return;
				}
			} catch (Exception ex) when (ex is COMException or ArgumentException) { dragFrom = dragTo = -1; }
		}
		if (m.Msg == WM_SETCURSOR && PaintCursor != null && (m.LParam.ToInt64() & 0xFFFF) == HTCLIENT) { Cursor.Current = PaintCursor; m.Result = (IntPtr)1; return; }
		if (m.Msg is WM_SETTEXT or EM_STREAMIN or EM_SETTEXTEX && IsHandleCreated) {
			int numerator = 0, denominator = 0;
			SendZoom(Handle, EM_GETZOOM, ref numerator, ref denominator);
			base.WndProc(ref m);
			int after = 0, afterDenominator = 0;
			SendZoom(Handle, EM_GETZOOM, ref after, ref afterDenominator);
			if (after <= 0 || afterDenominator <= 0) (after, afterDenominator) = (1, 1);    // (0/0: no zoom, 100%)
			if (numerator > 0 && denominator > 0 && (long)after * denominator != (long)numerator * afterDenominator)
				SendZoom(Handle, EM_SETZOOM, (IntPtr)numerator, (IntPtr)denominator);
			return;
		}
		if (m.Msg == WM_PAINT) {
			// (held: not painted now, and the window's need of a paint kept, for the paint after: QuietGaps)
			if (HoldPaint && IsHandleCreated) return;
			if (((pageBreaks == null || pageBreaks.Count > 0) || SectionMarksStale || sectionMarks!.Count > 0 || PageStarts is { Count: > 1 } || Gaps is { Applied: true } || Leaders?.Invoke() is { Count: > 0 } || Frames?.Invoke() is { Count: > 0 } || cellBlock != null) && IsHandleCreated && Buffered)
				PaintBuffered();
			else base.WndProc(ref m);
			Painted?.Invoke();
			return;
		}
		// Iteration 44: a message that scrolled the text (the wheel, the scroll bar, a key, RichEdit's timer while dragging):
		// painted at once, and with it what follows the editor's paint (the margins' bands, the line numbers). RichEdit
		// moves the text's pixels itself at once, but Windows paints only once no input is waiting: while the wheel
		// turned, the rest lagged behind the text by as long as it turned.
		if (m.Msg == WM_MOUSEWHEEL && (ModifierKeys & (Keys.Control | Keys.Shift)) == 0 && IsHandleCreated) {
			CoalesceWheel(ref m);
			if (!ScrollPastEnd(m)) WheelScroll((short)((m.WParam.ToInt64() >> 16) & 0xFFFF));
			m.Result = IntPtr.Zero;
			return;
		}
		if (m.Msg is WM_MOUSEWHEEL or WM_VSCROLL or WM_KEYDOWN or WM_TIMER && IsHandleCreated) {
			if (ScrollPastEnd(m)) return;
			var before = ScrollPosition();
			Typing = m.Msg == WM_KEYDOWN && (ModifierKeys & (Keys.Control | Keys.Alt)) == 0;
			base.WndProc(ref m);
			Typing = false;
			if (IsHandleCreated && ScrollPosition() != before) Update();
			// (iteration 46) Page Down or Ctrl+End at the end: on to the last sheet's footer, the caret still in view
			if (m.Msg == WM_KEYDOWN && ((int)m.WParam == 0x22 || ((int)m.WParam == 0x23 && (ModifierKeys & Keys.Control) != 0)) && AtScrollEnd()) {
				int caretTop = GetPositionFromCharIndex(TextDocument.Selection2.Start).Y;
				ScrollDown(Math.Min(RoomBelow(), Math.Max(0, caretTop)));
			}
			return;
		}
		if (m.Msg == WM_CHAR && (ModifierKeys & (Keys.Control | Keys.Alt)) == 0) {
			Typing = true;
			base.WndProc(ref m);
			Typing = false;
			return;
		}
		if (m.Msg == WM_CONTEXTMENU && ContextMenuHook != null) {
			long lp = m.LParam.ToInt64();
			int x = (short)(lp & 0xFFFF), y = (short)((lp >> 16) & 0xFFFF);
			bool keyboard = x == -1 && y == -1;    // Shift+F10 or the menu key
			var point = keyboard ? PointToScreen(GetPositionFromCharIndex(TextDocument.Selection2.Start)) : new Point(x, y);
			// the character under the mouse as RichEdit's own hit-testing has it (TOM, screen coordinates)
			int position = keyboard ? TextDocument.Selection2.Start : PositionAt(point);
			if (ContextMenuHook(point, position)) return;
		}
		base.WndProc(ref m);
	}
}

// The Windows spell checker (Windows 8 and later): the same one RichEdit uses for its red underlines.
sealed class Spelling
{
	public const int SuggestAction = 1, ReplaceAction = 2, DeleteAction = 3;    // CORRECTIVE_ACTION
	public record Error(int Start, int Length, string Word, int Action, string Replacement);

	readonly ISpellChecker checker;
	public readonly string Language;

	Spelling(ISpellChecker checker, string language) {
		this.checker = checker;
		Language = language;
	}

	// The checker for the current language (or English); null if Windows has no spell checker.
	public static Spelling? Create() {
		try {
			var type = Type.GetTypeFromCLSID(new Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC"));
			if (type == null || Activator.CreateInstance(type) is not ISpellCheckerFactory factory) return null;
			foreach (var language in new[] { CultureInfo.CurrentCulture.Name, CultureInfo.CurrentUICulture.Name, "en-US" })
				if (language != "" && factory.IsSupported(language) != 0) return new Spelling(factory.CreateSpellChecker(language), language);
		} catch (Exception ex) when (ex is COMException or InvalidCastException) { }
		return null;
	}

	// The errors in a paragraph's text; offset is the paragraph's start in the document.
	public List<Error> Check(string text, int offset) {
		var errors = new List<Error>();
		if (string.IsNullOrWhiteSpace(text)) return errors;
		text = CrossRefs.Blanked(text);    // (iteration 50: the cross-references' hidden markers aren't words)
		var found = checker.Check(text);
		while (found.Next(out var error) == 0) {
			int start = (int)error.StartIndex, length = (int)error.Length;
			if (start + length > text.Length) continue;
			errors.Add(new Error(offset + start, length, text.Substring(start, length), error.CorrectiveAction, error.Replacement ?? ""));
		}
		return errors;
	}

	public List<string> Suggest(string word, int max = 8) {
		var suggestions = new List<string>();
		var found = checker.Suggest(word);
		var one = new string[1];
		while (suggestions.Count < max && found.Next(1, one, IntPtr.Zero) == 0) suggestions.Add(one[0]);
		return suggestions;
	}

	public void Add(string word) => checker.Add(word);
}

[ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellCheckerFactory
{
	System.Runtime.InteropServices.ComTypes.IEnumString SupportedLanguages { get; }
	int IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
	ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
}

[ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellChecker
{
	string LanguageTag { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
	IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
	System.Runtime.InteropServices.ComTypes.IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
	void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
}

[ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumSpellingError
{
	[PreserveSig] int Next(out ISpellingError error);
}

[ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpellingError
{
	uint StartIndex { get; }
	uint Length { get; }
	int CorrectiveAction { get; }
	string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
}

// The COM IDropTarget interface (OLE drag and drop).
[ComImport, Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleDropTarget
{
	[PreserveSig] int OleDragEnter([MarshalAs(UnmanagedType.Interface)] object dataObject, int keyState, POINTL point, ref int effect);
	[PreserveSig] int OleDragOver(int keyState, POINTL point, ref int effect);
	[PreserveSig] int OleDragLeave();
	[PreserveSig] int OleDrop([MarshalAs(UnmanagedType.Interface)] object dataObject, int keyState, POINTL point, ref int effect);
}

[StructLayout(LayoutKind.Sequential)] public struct POINTL { public int X, Y; }

// One level of the list scheme (Paragraph > List levels): numbering style, punctuation and start number.
public sealed class ListLevel
{
	public string Style { get; set; } = "Decimal";      // a ListStyle name: Decimal, LowerAlpha, UpperAlpha, LowerRoman, UpperRoman, Bullet
	public string Punctuation { get; set; } = ".";      // ".", ")", "()" or ""
	public int Start { get; set; } = 1;
	// Iteration 38: where the level's number and text start, points from the margin (Word's number position and text
	// indent); null: automatic (the number at 18 points a level, from 18, the text after the level's widest number)
	public float? NumberAt { get; set; }
	public float? TextAt { get; set; }

	public static readonly string[] Punctuations = { ".", ")", "()", "" };

	public ListLevel Copy() => (ListLevel)MemberwiseClone();

	// 1. a. i. 1. a. i. ...
	public static List<ListLevel> Defaults() => Enumerable.Range(0, 9)
		.Select(i => new ListLevel { Style = (i % 3) switch { 0 => "Decimal", 1 => "LowerAlpha", _ => "LowerRoman" } }).ToList();

	// The TOM ListType: the numbering (low word) and the punctuation (high word: 0 is "1)").
	[System.Text.Json.Serialization.JsonIgnore]
	public int ListType {
		get {
			var style = Enum.TryParse<ListStyle>(Style, true, out var s) ? s : ListStyle.Decimal;
			if (style == ListStyle.Bullet) return tomConstants.tomListBullet;
			int punctuation = Punctuation switch { ")" => 0, "()" => tomConstants.tomListParentheses, "" => tomConstants.tomListPlain, _ => tomConstants.tomListPeriod };
			return (int)style | punctuation;
		}
	}

	// How the first number looks, e.g. "a)" or "(i)".
	public string Sample(int number) {
		var style = Enum.TryParse<ListStyle>(Style, true, out var s) ? s : ListStyle.Decimal;
		if (style == ListStyle.Bullet) return "•";
		string n = style switch {
			ListStyle.LowerAlpha => Letters(number).ToLowerInvariant(),
			ListStyle.UpperAlpha => Letters(number),
			ListStyle.LowerRoman => Roman(number).ToLowerInvariant(),
			ListStyle.UpperRoman => Roman(number),
			_ => number.ToString(CultureInfo.CurrentCulture)
		};
		return Punctuation switch { ")" => n + ")", "()" => "(" + n + ")", "" => n, _ => n + "." };
	}

	static string Letters(int n) {
		string result = "";
		for (n = Math.Max(1, n); n > 0; n = (n - 1) / 26) result = (char)('A' + (n - 1) % 26) + result;
		return result;
	}

	static string Roman(int n) {
		if (n <= 0) return "0";
		var sb = new StringBuilder();
		foreach (var (value, digits) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
			while (n >= value) { sb.Append(digits); n -= value; }
		return sb.ToString();
	}
}

// Paragraph > List levels: for each level 1-9 the numbering style, punctuation and start number, like a word
// processor's multilevel list; with a preview.
sealed class ListLevelsDialog : Form
{
	readonly ComboBox[] styles = new ComboBox[9], punctuations = new ComboBox[9];
	readonly NumericUpDown[] starts = new NumericUpDown[9];
	// Iteration 38: where each level's number and text start (points from the margin; empty: automatic)
	readonly TextBox[] numberAts = new TextBox[9], textAts = new TextBox[9];
	readonly CheckBox forNew = new() { Text = T("dlg.listLevels.forNew", "Also for new documents"), AutoSize = true, Checked = true };
	readonly CheckBox autoLevel = new() { Text = T("dlg.listLevels.autoLevel", "Tab and Shift+Tab give an item its new level's numbering style (automatic numbering style per level)"), AutoSize = true };
	public bool AutomaticPerLevel => autoLevel.Checked;
	readonly Label error = new() { AutoSize = true, ForeColor = AppTheme.Warning };
	readonly Label preview = new() { AutoSize = true, Font = new Font("Consolas", 9.5f), Padding = new Padding(8, 4, 8, 4), BackColor = SystemColors.Window };
	readonly CheckBox apply = new() { Text = T("dlg.listLevels.apply", "Apply to the current list"), AutoSize = true };

	static readonly (string Name, string Label)[] Styles = {
		("Decimal", "1, 2, 3"), ("LowerAlpha", "a, b, c"), ("UpperAlpha", "A, B, C"), ("LowerRoman", "i, ii, iii"), ("UpperRoman", "I, II, III"), ("Bullet", "• " + T("dlg.listLevels.bullet", "bullet"))
	};
	static readonly string[] PunctuationLabels = { "1.", "1)", "(1)", "1" };

	public List<ListLevel> Levels => Enumerable.Range(0, 9).Select(i => new ListLevel {
		Style = Styles[styles[i].SelectedIndex].Name, Punctuation = ListLevel.Punctuations[punctuations[i].SelectedIndex], Start = (int)starts[i].Value,
		NumberAt = PointsIn(numberAts[i]), TextAt = PointsIn(textAts[i])
	}).ToList();
	public bool ApplyToCurrentList => apply.Checked;
	public bool ForNewDocuments => forNew.Checked;

	// A box's points (empty or not a number: automatic; Validate says which aren't)
	static float? PointsIn(TextBox box) =>
		float.TryParse(box.Text.Trim().TrimEnd('t', 'p').Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out float v)
		|| float.TryParse(box.Text.Trim().TrimEnd('t', 'p').Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? Math.Clamp(v, 0, 1584) : null;

	bool Validate(out string message) {
		message = "";
		for (int i = 0; i < 9; i++)
			foreach (var (box, what) in new[] { (numberAts[i], T("dlg.listLevels.numberAt", "Number at")), (textAts[i], T("dlg.listLevels.textAt", "Text at")) })
				if (box.Text.Trim() != "" && (PointsIn(box) is not float v || v < 0 || v > 1584)) { message = T("dlg.listLevels.badPoints", "Level {0}, {1}: points from the margin, 0 to 1584 (empty: automatic).", i + 1, what); return false; }
		return true;
	}

	readonly ComboBox legalFormat = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	static readonly (LegalLists.Periods Format, string Label)[] LegalFormats = {
		(LegalLists.Periods.Top, "1.   1.1   1.1.1"), (LegalLists.Periods.All, "1.   1.1.   1.1.1."), (LegalLists.Periods.None, "1   1.1   1.1.1") };
	public LegalLists.Periods LegalFormat => LegalFormats[Math.Max(0, legalFormat.SelectedIndex)].Format;

	public ListLevelsDialog(List<ListLevel> levels, bool inList, LegalLists.Periods legal = LegalLists.Periods.Top, bool automaticPerLevel = true) {
		Dialogs.Setup(this, T("dlg.listLevels.title", "List levels"));
		autoLevel.Checked = automaticPerLevel;

		var table = new TableLayoutPanel { ColumnCount = 6, AutoSize = true };
		foreach (var heading in new[] { T("dlg.listLevels.level", "Level"), T("dlg.listLevels.numbering", "Numbering"), T("dlg.listLevels.punctuation", "Punctuation"), T("dlg.listLevels.startAt", "Start at"), T("dlg.listLevels.numberAtPt", "Number at (pt)"), T("dlg.listLevels.textAtPt", "Text at (pt)") })
			table.Controls.Add(new Label { Text = heading, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 3, 3, 6) });
		for (int i = 0; i < 9; i++) {
			var level = i < levels.Count ? levels[i] : ListLevel.Defaults()[i];
			styles[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = LogicalToDeviceUnits(110) };
			styles[i].Items.AddRange(Styles.Select(st => (object)st.Label).ToArray());
			styles[i].SelectedIndex = Math.Max(0, Array.FindIndex(Styles, st => st.Name.Equals(level.Style, StringComparison.OrdinalIgnoreCase)));
			punctuations[i] = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = LogicalToDeviceUnits(70) };
			punctuations[i].Items.AddRange(PunctuationLabels);
			punctuations[i].SelectedIndex = Math.Max(0, Array.IndexOf(ListLevel.Punctuations, level.Punctuation));
			starts[i] = new NumericUpDown { Minimum = 0, Maximum = 999, Value = Math.Clamp(level.Start, 0, 999), Width = LogicalToDeviceUnits(60) };
			int row = i;
			styles[i].SelectedIndexChanged += (s, e) => { punctuations[row].Enabled = starts[row].Enabled = styles[row].SelectedIndex != Styles.Length - 1; UpdatePreview(); };
			punctuations[i].SelectedIndexChanged += (s, e) => UpdatePreview();
			starts[i].ValueChanged += (s, e) => UpdatePreview();
			table.Controls.Add(new Label { Text = (i + 1).ToString(), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 3, 3) });
			table.Controls.Add(styles[i]);
			table.Controls.Add(punctuations[i]);
			table.Controls.Add(starts[i]);
			punctuations[i].Enabled = starts[i].Enabled = styles[i].SelectedIndex != Styles.Length - 1;
			// (where the number and the text start; empty: automatic, the number at 18 points a level, the text after it)
			numberAts[i] = new TextBox { Width = LogicalToDeviceUnits(70), PlaceholderText = T("dlg.listLevels.autoAt", "auto ({0})", 18 * (i + 1)), Text = Shown(level.NumberAt) };
			textAts[i] = new TextBox { Width = LogicalToDeviceUnits(70), PlaceholderText = T("dlg.listLevels.auto", "auto"), Text = Shown(level.TextAt) };
			numberAts[i].TextChanged += (s, e) => UpdatePreview();
			textAts[i].TextChanged += (s, e) => UpdatePreview();
			table.Controls.Add(numberAts[i]);
			table.Controls.Add(textAts[i]);
		}

		var defaults = Dialogs.Button(this, T("dlg.listLevels.defaults", "Defaults"));
		defaults.Click += (s, e) => {
			var d = ListLevel.Defaults();
			for (int i = 0; i < 9; i++) {
				styles[i].SelectedIndex = Array.FindIndex(Styles, st => st.Name == d[i].Style);
				punctuations[i].SelectedIndex = 0;
				starts[i].Value = 1;
				numberAts[i].Text = textAts[i].Text = "";
			}
		};
		var ok = Dialogs.Ok(this, closes: false);
		ok.Click += (s, e) => { if (Validate(out string message)) { DialogResult = DialogResult.OK; Close(); } else error.Text = message; };
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, defaults, ok, cancel);
		apply.Checked = apply.Enabled = inList;
		if (!inList) apply.Text += " " + T("dlg.listLevels.notInList", "(the caret is not in a list)");

		var right = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(16, 0, 0, 0) };
		right.Controls.Add(new Label { Text = T("dlg.listLevels.preview", "Preview"), AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 3, 3, 6) });
		right.Controls.Add(preview);
		var top = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
		top.Controls.Add(table);
		top.Controls.Add(right);
		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
		layout.Controls.Add(top);
		autoLevel.Margin = new Padding(3, 8, 3, 0);
		layout.Controls.Add(autoLevel);
		layout.Controls.Add(new Label { AutoSize = true, Margin = new Padding(3, 8, 3, 3), MaximumSize = new Size(LogicalToDeviceUnits(640), 0),
			Text = T("dlg.listLevels.note", "New numbered lists use level 1; with the box above ticked, Tab and Shift+Tab give items the numbering of their new level (unticked, they keep theirs). " +
				"Number at and Text at are where a level's number and its text (and the lines after the first) start, in points from the margin (72 points an inch); " +
				"empty: automatic. They are the document's, saved in it, for its lists and legal numbering; numbered styles have their own (Modify Style).") });
		var legalRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 6, 0, 0) };
		legalFormat.Items.AddRange(LegalFormats.Select(f => (object)f.Label).ToArray());
		legalFormat.SelectedIndex = Math.Max(0, Array.FindIndex(LegalFormats, f => f.Format == legal));
		legalFormat.Width = LogicalToDeviceUnits(150);
		legalRow.Controls.Add(new Label { Text = T("dlg.listLevels.legal", "Legal numbering (Home > Paragraph > 1.1):"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) });
		legalRow.Controls.Add(legalFormat);
		layout.Controls.Add(legalRow);
		layout.Controls.Add(apply);
		layout.Controls.Add(forNew);
		layout.Controls.Add(error);
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		UpdatePreview();
	}

	static new string Shown(float? points) => points is float v ? v.ToString("0.##", CultureInfo.CurrentCulture) : "";

	// (the first four levels, each indented as set: the number where it starts, its text where it starts, roughly to
	// scale, a character for 6 points)
	void UpdatePreview() {
		var levels = Levels;
		var sb = new StringBuilder();
		for (int i = 0; i < 4; i++) {
			float number = levels[i].NumberAt ?? 18 * (i + 1), text = levels[i].TextAt ?? number + 18;
			for (int n = 0; n < (i < 3 ? 2 : 1); n++) {
				string sample = levels[i].Sample(levels[i].Start + n);
				int lead = (int)Math.Round(number / 6), gap = Math.Max(1, (int)Math.Round((text - number) / 6) - sample.Length);
				sb.Append(new string(' ', Math.Min(60, lead))).Append(sample).Append(new string(' ', Math.Min(40, gap))).Append(T("dlg.listLevels.previewLevel", "Level {0}", i + 1)).Append('\n');
			}
		}
		preview.Text = sb.ToString().TrimEnd();
	}
}

// Table > Insert table: rows, columns and a bold header row.
// Iteration 18: what the Table Properties dialog changes (null: unchanged). Lengths in twips; Height 0 is automatic,
// negative exactly that high.
sealed class TableChanges
{
	public int? Alignment, Padding, Height, VerticalAlign;
	public List<int>? Widths;
	public string? Borders;
	public RtfTables.Border Border = new(RtfTables.DefaultBorder, null);
	public bool SetShading, WholeTable;
	public Color? Shading;

	static string Pt(int twips) => (twips / 20.0).ToString("0.##", CultureInfo.InvariantCulture);
	static string Hex(Color? c) => c is Color k ? $"#{k.R:X2}{k.G:X2}{k.B:X2}" : "";

	// The automation calls that make the same changes (for macro recording).
	public IEnumerable<string> Calls() {
		if (Widths != null) yield return MacroRecorder.Call("SetColumnWidths", string.Join(", ", Widths.Select(w => Pt(w) + "pt")));
		if (Alignment is int a) yield return MacroRecorder.Call("SetTableAlignment", new[] { "left", "center", "right" }[a]);
		if (Padding is int p) yield return MacroRecorder.Call("SetCellPadding", (float)(p / 20.0));
		if (Height is int h) yield return MacroRecorder.Call("SetRowHeight", (float)(Math.Abs(h) / 20.0), h < 0, WholeTable);
		if (Borders != null) yield return MacroRecorder.Call("SetCellBorders", Borders, (float)(Border.Width / 20.0), Hex(Border.Color), WholeTable);
		if (SetShading) yield return MacroRecorder.Call("SetCellShading", Shading is null ? "none" : Hex(Shading), WholeTable);
		if (VerticalAlign is int v) yield return MacroRecorder.Call("SetCellVerticalAlignment", new[] { "top", "center", "bottom" }[v], WholeTable);
	}
}

// Iteration 18: Table Tools > Properties. Column widths, the table's alignment and cell padding are for the whole table;
// row height, borders, shading and vertical alignment for the selected cells or the whole table.
sealed class TablePropertiesDialog : Form
{
	readonly ComboBox scope = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly ComboBox alignment = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly NumericUpDown padding = new() { DecimalPlaces = 1, Increment = 0.5m, Maximum = 72, Width = 70 };
	readonly ComboBox heightKind = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly NumericUpDown height = new() { DecimalPlaces = 1, Increment = 1, Maximum = 1440, Width = 70 };
	readonly List<NumericUpDown> widths = new();
	readonly Label total = new() { AutoSize = true };
	readonly ComboBox borders = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly ComboBox borderWidth = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80 };
	readonly Button borderColor = new() { Text = T("dlg.tableProps.color", "Color..."), AutoSize = true };
	readonly ComboBox shading = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly Button shadingColor = new() { Text = T("dlg.tableProps.color", "Color..."), AutoSize = true };
	readonly ComboBox verticalAlign = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly List<RtfTables.RowFormat> rows;
	readonly int pageWidth;
	readonly bool metric = RegionInfo.CurrentRegion.IsMetric;
	Color? lineColor;
	Color fillColor = Color.FromArgb(0xDD, 0xEB, 0xF7);
	readonly int initialAlignment, initialPadding, initialHeight;
	readonly List<int> initialWidths;

	static readonly float[] BorderWidths = { 0.25f, 0.5f, 0.75f, 1f, 1.5f, 2.25f, 3f };
	// (the Borders box's items, as TableChanges.Borders names them: the box shows them in the interface's language)
	static readonly string[] BorderKinds = { "", "all", "outside", "inside", "none", "top", "bottom", "left", "right" };

	double Unit => metric ? 1440 / 2.54 : 1440;    // twips per cm or inch

	public TablePropertiesDialog(List<RtfTables.RowFormat> rows, int selectedCells, int pageWidth) {
		this.rows = rows;
		this.pageWidth = pageWidth;
		Dialogs.Setup(this, T("dlg.tableProps.title", "Table Properties"));
		Font = SystemFonts.MessageBoxFont ?? Font;    // (as before the dialog helper)
		var first = rows[0];
		initialAlignment = first.Alignment;
		initialPadding = first.Padding;
		initialHeight = first.Height;
		initialWidths = first.Cells.Select(c => c.Width).ToList();

		var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill };
		Controls.Add(layout);

		// the whole table
		alignment.Items.AddRange(new object[] { T("dlg.tableProps.left", "Left"), T("dlg.tableProps.center", "Center"), T("dlg.tableProps.right", "Right") });
		alignment.SelectedIndex = first.Alignment;
		padding.Value = (decimal)Math.Min(72, first.Padding / 20.0);
		layout.Controls.Add(Group(T("dlg.tableProps.table", "Table"), Row(Label(T("dlg.tableProps.alignment", "&Alignment on the page:")), alignment), Row(Label(T("dlg.tableProps.padding", "Cell &padding (left and right, pt):")), padding)));

		var columns = new FlowLayoutPanel { AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(520), 0), WrapContents = true };
		for (int c = 0; c < first.Cells.Count; c++) {
			var box = new NumericUpDown { DecimalPlaces = 2, Increment = metric ? 0.25m : 0.1m, Minimum = metric ? 0.25m : 0.1m, Maximum = 50, Width = 70,
				Value = (decimal)Math.Clamp(first.Cells[c].Width / Unit, metric ? 0.25 : 0.1, 50) };
			box.ValueChanged += (s, e) => ShowTotal();
			widths.Add(box);
			columns.Controls.Add(Row(Label($"{c + 1}:"), box));
		}
		var equal = new Button { Text = T("dlg.tableProps.equal", "&Equal widths"), AutoSize = true };
		equal.Click += (s, e) => SetWidths(Enumerable.Repeat(widths.Sum(w => (double)w.Value) / widths.Count, widths.Count).ToList());
		var fit = new Button { Text = T("dlg.tableProps.fit", "&Fit to page width"), AutoSize = true };
		fit.Click += (s, e) => { double sum = widths.Sum(w => (double)w.Value); SetWidths(widths.Select(w => (double)w.Value * pageWidth / Unit / sum).ToList()); };
		bool sameCells = rows.All(r => r.Cells.Count == first.Cells.Count);
		layout.Controls.Add(Group((metric ? T("dlg.tableProps.widthsCm", "Column widths (cm)") : T("dlg.tableProps.widthsIn", "Column widths (inches)")) + (sameCells ? "" : T("dlg.tableProps.widthsTheseRows", ", for the rows with this many cells")),
			columns, Row(equal, fit, total)));
		ShowTotal();

		// the selected cells or the whole table
		scope.Items.AddRange(new object[] { TN("dlg.tableProps.selectedCells", "The selected cell ({0})", "The selected cells ({0})", selectedCells), T("dlg.tableProps.wholeTable", "The whole table") });
		scope.SelectedIndex = selectedCells >= rows.Sum(r => r.Cells.Count) ? 1 : 0;
		scope.Width = LogicalToDeviceUnits(200);
		heightKind.Items.AddRange(new object[] { T("dlg.tableProps.automatic", "Automatic"), T("dlg.tableProps.atLeast", "At least"), T("dlg.tableProps.exactly", "Exactly") });
		heightKind.SelectedIndex = first.Height == 0 ? 0 : first.Height > 0 ? 1 : 2;
		height.Value = (decimal)Math.Min(1440, Math.Abs(first.Height) / 20.0);
		height.Enabled = heightKind.SelectedIndex != 0;
		heightKind.SelectedIndexChanged += (s, e) => { height.Enabled = heightKind.SelectedIndex != 0; if (height.Enabled && height.Value == 0) height.Value = 18; };
		borders.Items.AddRange(new object[] { T("dlg.tableProps.noChange", "(no change)"), T("dlg.tableProps.all", "All"), T("dlg.tableProps.outside", "Outside"), T("dlg.tableProps.inside", "Inside"), T("dlg.tableProps.none", "None"),
			T("dlg.tableProps.borderTop", "Top"), T("dlg.tableProps.borderBottom", "Bottom"), T("dlg.tableProps.borderLeft", "Left"), T("dlg.tableProps.borderRight", "Right") });
		borders.SelectedIndex = 0;
		foreach (float w in BorderWidths) borderWidth.Items.Add(w.ToString("0.##", CultureInfo.CurrentCulture) + " pt");
		borderWidth.SelectedIndex = 1;
		borderColor.Click += (s, e) => { if (Pick(lineColor ?? Color.Black) is Color c) { lineColor = c; ShowColors(); if (borders.SelectedIndex == 0) borders.SelectedIndex = 1; } };
		shading.Items.AddRange(new object[] { T("dlg.tableProps.noChange", "(no change)"), T("dlg.tableProps.none", "None"), T("dlg.tableProps.colorItem", "Color") });
		shading.SelectedIndex = 0;
		shadingColor.Click += (s, e) => { if (Pick(fillColor) is Color c) { fillColor = c; shading.SelectedIndex = 2; ShowColors(); } };
		verticalAlign.Items.AddRange(new object[] { T("dlg.tableProps.noChange", "(no change)"), T("dlg.tableProps.top", "Top"), T("dlg.tableProps.center", "Center"), T("dlg.tableProps.bottom", "Bottom") });
		verticalAlign.SelectedIndex = 0;
		ShowColors();
		layout.Controls.Add(Group(T("dlg.tableProps.cells", "Cells"), Row(Label(T("dlg.tableProps.applyTo", "Apply &to:")), scope), Row(Label(T("dlg.tableProps.rowHeight", "Row &height:")), heightKind, height, Label("pt")),
			Row(Label(T("dlg.tableProps.borders", "&Borders:")), borders, borderWidth, borderColor), Row(Label(T("dlg.tableProps.shading", "&Shading:")), shading, shadingColor),
			Row(Label(T("dlg.tableProps.verticalAlignment", "&Vertical alignment:")), verticalAlign)));

		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Row(ok, cancel);
		buttons.Anchor = AnchorStyles.Right;
		layout.Controls.Add(buttons);
		AcceptButton = ok;
		CancelButton = cancel;
	}

	static Label Label(string text) => new() { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) };

	static FlowLayoutPanel Row(params Control[] controls) {
		var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
		row.Controls.AddRange(controls);
		return row;
	}

	static GroupBox Group(string title, params Control[] rows) {
		var group = new GroupBox { Text = title, AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(8) };
		var inner = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill };
		foreach (var r in rows) inner.Controls.Add(r);
		group.Controls.Add(inner);
		return group;
	}

	void SetWidths(List<double> values) {
		for (int i = 0; i < widths.Count; i++) widths[i].Value = (decimal)Math.Clamp(Math.Round(values[i], 2), (double)widths[i].Minimum, (double)widths[i].Maximum);
	}

	void ShowTotal() => total.Text = metric ? T("dlg.tableProps.totalCm", "Total {0:0.##} of {1:0.##} cm", widths.Sum(w => w.Value), pageWidth / Unit) : T("dlg.tableProps.totalIn", "Total {0:0.##} of {1:0.##} in", widths.Sum(w => w.Value), pageWidth / Unit);

	void ShowColors() {
		borderColor.BackColor = lineColor ?? SystemColors.Control;
		borderColor.ForeColor = lineColor is Color l && l.GetBrightness() < 0.5f ? Color.White : SystemColors.ControlText;
		shadingColor.BackColor = fillColor;
	}

	Color? Pick(Color initial) {
		using var dlg = new ColorDialog { Color = initial, FullOpen = true };
		return dlg.ShowDialog(this) == DialogResult.OK ? dlg.Color : null;
	}

	/// What the dialog changes (only what differs from the table as it was, or was chosen).
	public TableChanges Changes {
		get {
			var changes = new TableChanges { WholeTable = scope.SelectedIndex == 1 };
			if (alignment.SelectedIndex != initialAlignment) changes.Alignment = alignment.SelectedIndex;
			int pad = (int)Math.Round(padding.Value * 20);
			if (pad != initialPadding) changes.Padding = pad;
			int h = heightKind.SelectedIndex == 0 ? 0 : (int)Math.Round(height.Value * 20) * (heightKind.SelectedIndex == 2 ? -1 : 1);
			if (h != initialHeight) changes.Height = h;
			var w = widths.Select(box => (int)Math.Round((double)box.Value * Unit)).ToList();
			if (w.Where((value, i) => Math.Abs(value - initialWidths[i]) > 10).Any()) changes.Widths = w;
			if (borders.SelectedIndex > 0) {
				changes.Borders = BorderKinds[borders.SelectedIndex];
				changes.Border = new RtfTables.Border((int)Math.Round(BorderWidths[borderWidth.SelectedIndex] * 20), lineColor);
			}
			if (shading.SelectedIndex > 0) { changes.SetShading = true; changes.Shading = shading.SelectedIndex == 2 ? fillColor : null; }
			if (verticalAlign.SelectedIndex > 0) changes.VerticalAlign = verticalAlign.SelectedIndex - 1;
			return changes;
		}
	}
}

sealed class InsertTableDialog : Form
{
	readonly NumericUpDown rows, columns;
	readonly CheckBox header = new() { Text = T("dlg.insertTable.header", "Header row (bold)"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) };

	public int Rows => (int)rows.Value;
	public int Columns => (int)columns.Value;
	public bool Header => header.Checked;

	public InsertTableDialog(int rowCount, int columnCount, bool headerRow, bool fromFile = false) {
		Dialogs.Setup(this, T("dlg.insertTable.title", "Insert table"));
		rows = new NumericUpDown { Minimum = 1, Maximum = 500, Value = Math.Clamp(rowCount, 1, 500), Width = LogicalToDeviceUnits(70) };
		columns = new NumericUpDown { Minimum = 1, Maximum = 63, Value = Math.Clamp(columnCount, 1, 63), Width = LogicalToDeviceUnits(70) };
		header.Checked = headerRow;
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
		table.Controls.Add(new Label { Text = T("dlg.insertTable.rows", "&Rows:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(rows);
		table.Controls.Add(new Label { Text = T("dlg.insertTable.columns", "&Columns:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(columns);
		table.Controls.Add(header);
		table.SetColumnSpan(header, 2);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		// (iteration 91) From File: the table made from a workbook, CSV or JSON file instead (DialogResult.Yes)
		var buttons = fromFile ? Dialogs.Row(this, 0, Dialogs.Button(this, T("dlg.insertTable.fromFile", "From &File..."), DialogResult.Yes), ok, cancel) : Dialogs.Row(this, 0, ok, cancel);
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

// (iteration 92) Table Tools > Split Cells: the number of columns and rows.
sealed class SplitCellsDialog : Form
{
	readonly NumericUpDown columns, rows;

	public int Columns => (int)columns.Value;
	public int Rows => (int)rows.Value;

	public SplitCellsDialog(int columnCount, int rowCount) {
		Dialogs.Setup(this, T("dlg.splitCells.title", "Split cells"));
		columns = new NumericUpDown { Minimum = 1, Maximum = 63, Value = Math.Clamp(columnCount, 1, 63), Width = LogicalToDeviceUnits(70) };
		rows = new NumericUpDown { Minimum = 1, Maximum = 100, Value = Math.Clamp(rowCount, 1, 100), Width = LogicalToDeviceUnits(70) };
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
		table.Controls.Add(new Label { Text = T("dlg.splitCells.columns", "Number of &columns:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(columns);
		table.Controls.Add(new Label { Text = T("dlg.splitCells.rows", "Number of &rows:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(rows);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

// (iteration 92) Table Tools > Delete Cells: what becomes of the cells around them.
sealed class DeleteCellsDialog : Form
{
	readonly RadioButton left = new() { Text = T("dlg.deleteCells.left", "Shift cells &left"), AutoSize = true, Checked = true },
		up = new() { Text = T("dlg.deleteCells.up", "Shift cells &up"), AutoSize = true },
		row = new() { Text = T("dlg.deleteCells.row", "Delete entire &row"), AutoSize = true },
		column = new() { Text = T("dlg.deleteCells.column", "Delete entire &column"), AutoSize = true };

	/// "left", "up", "row" or "column" (RtfTables.DeleteCells).
	public string How => up.Checked ? "up" : row.Checked ? "row" : column.Checked ? "column" : "left";

	public DeleteCellsDialog() {
		Dialogs.Setup(this, T("dlg.deleteCells.title", "Delete cells"));
		var table = new TableLayoutPanel { ColumnCount = 1, AutoSize = true };
		foreach (var choice in new[] { left, up, row, column }) table.Controls.Add(choice);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		table.Controls.Add(Dialogs.Row(this, 0, ok, cancel));
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

// (iteration 91) Insert table from file, for a workbook: which sheet, a range (or the cells in use), and the header row.
sealed class TableFromFileDialog : Form
{
	readonly ComboBox sheet = new() { DropDownStyle = ComboBoxStyle.DropDownList };
	readonly TextBox range = new();
	readonly CheckBox header = new() { Text = T("dlg.insertTable.header", "Header row (bold)"), AutoSize = true, Margin = new Padding(3, 8, 3, 3) };

	public string Sheet => sheet.SelectedItem as string ?? "";
	public string Range => range.Text.Trim();
	public bool Header => header.Checked;

	public TableFromFileDialog(string file, List<string> sheets, bool headerRow) {
		Dialogs.Setup(this, T("dlg.tableFromFile.title", "Insert table from file"));
		foreach (string name in sheets) sheet.Items.Add(name);
		if (sheet.Items.Count > 0) sheet.SelectedIndex = 0;
		sheet.Enabled = sheets.Count > 1;
		header.Checked = headerRow;
		sheet.Width = range.Width = LogicalToDeviceUnits(220);
		Dialogs.FitWidth(sheet, Font);
		var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
		var title = new Label { Text = file, AutoSize = true, Margin = new Padding(3, 3, 3, 8) };
		table.Controls.Add(title);
		table.SetColumnSpan(title, 2);
		table.Controls.Add(new Label { Text = T("dlg.tableFromFile.sheet", "&Sheet:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(sheet);
		table.Controls.Add(new Label { Text = T("dlg.tableFromFile.range", "&Range:"), AutoSize = true, Anchor = AnchorStyles.Left });
		table.Controls.Add(range);
		var hint = new Label { Text = T("dlg.tableFromFile.rangeHint", "Such as A1:D20; empty for all the cells in use."), AutoSize = true, ForeColor = SystemColors.GrayText };
		table.Controls.Add(new Label { AutoSize = true });
		table.Controls.Add(hint);
		table.Controls.Add(header);
		table.SetColumnSpan(header, 2);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		table.Controls.Add(buttons);
		table.SetColumnSpan(buttons, 2);
		Controls.Add(table);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

// File > Header and Footer: show / hide each, and edit its left, center and right parts.
// The header or footer shown above or below the text: its left, center and right parts in grey, with a dotted line
// between it and the text. Clicking it opens File > Header and Footer.
sealed class BandStrip : Control
{
	readonly bool header;
	string[] parts = { "", "", "" };
	readonly ToolTip tip = new();

	public BandStrip(bool header) {
		this.header = header;
		Dock = header ? DockStyle.Top : DockStyle.Bottom;
		Cursor = Cursors.Hand;
		Visible = false;
		SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
		BackColor = SystemColors.Window;
		Height = 10 + Font.Height;
		tip.SetToolTip(this, (header ? "Header" : "Footer") + " (as on the first page). Click to change it. It is printed on every page, and saved in Word and OpenDocument files.");
	}

	protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); FitHeight(); }
	protected override void OnCreateControl() { base.OnCreateControl(); FitHeight(); }
	// (iteration 55) a line per line of the longest part, the band's font at the strip's size (words in theirs, in
	// proportion)
	new float Scale => Font.SizeInPoints / (band ?? BandFont.Plain).Size;
	void FitHeight() => Height = LogicalToDeviceUnits(10) + (int)Math.Ceiling(Math.Max(Font.Height, BandText.Height(parts, band ?? BandFont.Plain, f => f.GetHeight(DeviceDpi), Scale)));

	/// Whether it is shown (Visible is also false while the window is not).
	public bool Shown { get; private set; }

	public void Display(bool show, string[] parts, BandFont? font = null) {
		this.parts = parts;
		band = font;
		FitHeight();
		Shown = Visible = show;
		Invalidate();
	}
	BandFont? band;

	/// Space at either side, as the editor's text has (Print Layout: the page's margins are the editor's own).
	public (int Left, int Right) Inset {
		get => inset;
		set { if (value == inset) return; inset = value; Invalidate(); }
	}
	(int Left, int Right) inset;

	protected override void OnPaint(PaintEventArgs e) {
		var g = e.Graphics;
		var area = Rectangle.FromLTRB(inset.Left, 0, Math.Max(inset.Left + 1, ClientSize.Width - inset.Right), ClientSize.Height);
		int gap = LogicalToDeviceUnits(4);
		using var pen = new Pen(Color.FromArgb(190, 190, 190)) { DashStyle = DashStyle.Dot };
		int lineY = header ? area.Bottom - 1 : area.Top;
		g.DrawLine(pen, area.Left, lineY, area.Right - LogicalToDeviceUnits(16), lineY);
		var text = new Rectangle(area.Left, header ? area.Top : area.Top + gap, area.Width - LogicalToDeviceUnits(16), area.Height - gap);
		using var plain = new SolidBrush(band != null ? AppTheme.Ink : AppTheme.FaintInk);
		using var formatted = new SolidBrush(AppTheme.Ink);
		g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
		BandText.Draw(g, parts, band ?? BandFont.Plain, text.Left, text.Width, text.Top + LogicalToDeviceUnits(3), Scale, plain, formatted);
	}

	protected override void Dispose(bool disposing) {
		if (disposing) tip.Dispose();
		base.Dispose(disposing);
	}
}

sealed class HeaderFooterDialog : Form
{
	// (iteration 55) a header's or footer's controls: shown or not, its three parts (rich text: words and lines in fonts of
	// their own, BandText's marks), its font (null: Segoe UI 8 pt in grey) and its line
	sealed class BandBox
	{
		public readonly bool IsHeader, First;
		public readonly CheckBox Show = new() { AutoSize = true };
		public readonly RichTextBox[] Parts = Enumerable.Range(0, 3).Select(_ => AppTheme.OnPaper(new RichTextBox {
			Multiline = true, WordWrap = false, ScrollBars = RichTextBoxScrollBars.Vertical, DetectUrls = false, HideSelection = false
		})).ToArray();
		public BandFont? Font;
		public readonly Label FontLabel = new() { AutoSize = true, Anchor = AnchorStyles.Left };
		public readonly CheckBox Line = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(16, 7, 3, 0) };
		public readonly FlowLayoutPanel FontRow = new() { AutoSize = true, WrapContents = false, Margin = new Padding(0, 4, 0, 0) };
		public string Name => (First ? "First page " : "") + (IsHeader ? "header" : "footer");
		public BandBox(bool isHeader, bool first) {
			(IsHeader, First) = (isHeader, first);
			Show.Text = first ? (isHeader ? T("dlg.headerFooter.showFirstHeader", "Show a header on the first page") : T("dlg.headerFooter.showFirstFooter", "Show a footer on the first page")) : isHeader ? T("dlg.headerFooter.showHeader", "Show a header") : T("dlg.headerFooter.showFooter", "Show a footer");
			Line.Text = isHeader ? T("dlg.headerFooter.lineBelow", "Line below the header") : T("dlg.headerFooter.lineAbove", "Line above the footer");
		}
	}

	readonly BandBox header = new(true, false), footer = new(false, false), firstHeader = new(true, true), firstFooter = new(false, true);
	BandBox[] Bands => new[] { header, footer, firstHeader, firstFooter };
	readonly CheckBox makeDefault = new() { Text = T("dlg.headerFooter.makeDefault", "Use for new documents too"), AutoSize = true, Margin = new Padding(3, 6, 3, 0) };
	// (iteration 49) the first page's own; (iteration 55) in every way, on a tab of its own
	readonly CheckBox differentFirst = new() { Text = T("dlg.headerFooter.differentFirst", "Different first page: page 1 has a header and footer of its own (or none), the other pages theirs"), AutoSize = true, Margin = new Padding(3, 4, 3, 4) };
	readonly TabControl tabs = new();
	readonly TabPage otherTab = new(), firstTab = new();
	readonly FlowLayoutPanel firstPanel = new() { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
	// (iteration 55) from the page's edges, in inches or centimetres as the region measures
	readonly NumericUpDown headerDistance = new() { DecimalPlaces = 2, Width = 70, AccessibleName = T("dlg.headerFooter.headerFromTopName", "Header from top") }, footerDistance = new() { DecimalPlaces = 2, Width = 70, AccessibleName = T("dlg.headerFooter.footerFromBottomName", "Footer from bottom") };
	readonly bool metric = System.Globalization.RegionInfo.CurrentRegion.IsMetric;
	decimal ToUnits(int hundredths) => metric ? Math.Round(hundredths * 0.0254m, 2) : hundredths / 100m;
	int FromUnits(decimal value) => (int)Math.Round(metric ? value / 0.0254m : value * 100);
	// the part box the selection buttons format (the last one typed in)
	RichTextBox? lastPart;
	RichFormat.Buttons styleButtons = null!;

	/// Whether the header and footer become the default for new documents.
	public bool MakeDefault => makeDefault.Checked;

	// (iteration 63) a section after a break: the same header and footer as the section before's, or its own; its page
	// numbers starting again (at what number) or going on
	readonly CheckBox sameAsPrevious = new() { Text = T("dlg.headerFooter.sameAsPrevious", "Same as the section before"), AutoSize = true, Margin = new Padding(12, 6, 3, 0) };
	readonly CheckBox restart = new() { Text = T("dlg.headerFooter.restart", "Page numbers start again at"), AutoSize = true, Margin = new Padding(16, 6, 0, 0) };
	readonly NumericUpDown startAt = new() { Minimum = 0, Maximum = 99999, Width = 70, AccessibleName = T("dlg.headerFooter.startAtName", "First page number"), Margin = new Padding(3, 3, 3, 0) };
	readonly SectionSettings? section;
	// a section's own top and bottom margins (the document's until ticked)
	readonly CheckBox ownMargins = new() { Text = T("dlg.headerFooter.ownMargins", "Margins of its own: top"), AutoSize = true, Margin = new Padding(16, 6, 0, 0) };
	readonly NumericUpDown sectionTop = new() { DecimalPlaces = 2, Width = 70, AccessibleName = T("dlg.headerFooter.sectionTopName", "Section's top margin"), Margin = new Padding(3, 3, 3, 0) };
	readonly NumericUpDown sectionBottom = new() { DecimalPlaces = 2, Width = 70, AccessibleName = T("dlg.headerFooter.sectionBottomName", "Section's bottom margin"), Margin = new Padding(3, 3, 3, 0) };

	/// Iteration 63: a section's settings as the dialog has them (null: the document's own header and footer, the first
	/// section's): Same, or its own header and footer (Result), and its numbering.
	public SectionSettings? SectionResult() {
		if (section == null) return null;
		int Twips(NumericUpDown box) => (int)Math.Round(FromUnits(box.Value) * 14.4);
		var page = SectionPage.WithMargins(section.Page, ownMargins.Checked ? Twips(sectionTop) : null, ownMargins.Checked ? Twips(sectionBottom) : null);
		// (a header and footer of its own have the distances the dialog gives them)
		if (page != null && !sameAsPrevious.Checked) page = page with { HeaderY = null, FooterY = null };
		return new SectionSettings { Same = sameAsPrevious.Checked, Bands = sameAsPrevious.Checked ? null : Result(), Restart = restart.Checked, Start = (int)startAt.Value, Page = page == new SectionPage() ? null : page, Columns = section.Columns, Continuous = section.Continuous };
	}

	/// The document's header and footer (section: a section after a break, its number from 1 and its settings; bands are
	/// then its header and footer as they are now, the section before's when it is the same; margins: the document's top
	/// and bottom, 1/100 inch, which a section's own start from).
	public HeaderFooterDialog(PageBands bands, PageBands defaultBands, SectionSettings? section = null, int sectionNumber = 0, int[]? margins = null) {
		defaults_ = defaultBands;
		this.section = section;
		Dialogs.Setup(this, T("dlg.headerFooter.title", "Header and Footer"));

		var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
		if (section != null) {
			Text = T("dlg.headerFooter.titleSection", "Header and Footer: Section {0}", sectionNumber);
			var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(3, 0, 3, 6) };
			row.Controls.Add(new Label { Text = T("dlg.headerFooter.section", "Section {0} (the caret's):", sectionNumber), AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(0, 7, 3, 0) });
			row.Controls.Add(sameAsPrevious);
			row.Controls.Add(restart);
			row.Controls.Add(startAt);
			layout.Controls.Add(row);
			sameAsPrevious.Checked = section.Same;
			restart.Checked = section.Restart;
			startAt.Value = Math.Clamp(section.Start, 0, 99999);
			startAt.Enabled = restart.Checked;
			restart.CheckedChanged += (s, e) => startAt.Enabled = restart.Checked;
			new ToolTip().SetToolTip(sameAsPrevious, T("dlg.headerFooter.sameAsPreviousTip", "Ticked, this section prints the header and footer of the section before (Word's Link to Previous); untick it to give it its own."));
			row.Controls.Add(ownMargins);
			row.Controls.Add(sectionTop);
			row.Controls.Add(new Label { Text = T("dlg.headerFooter.bottom", "bottom"), AutoSize = true, Margin = new Padding(3, 7, 0, 0) });
			row.Controls.Add(sectionBottom);
			row.Controls.Add(new Label { Text = metric ? "cm" : T("dlg.headerFooter.in", "in"), AutoSize = true, Margin = new Padding(3, 7, 3, 0) });
			foreach (var box in new[] { sectionTop, sectionBottom }) box.Maximum = ToUnits(1000);
			ownMargins.Checked = section.Page?.Top != null || section.Page?.Bottom != null;
			decimal Start(int? twips, int side) => Math.Min(sectionTop.Maximum, ToUnits(twips is int t ? (int)Math.Round(Math.Abs(t) / 14.4) : margins is { Length: 2 } ? margins[side] : 100));
			sectionTop.Value = Start(section.Page?.Top, 0);
			sectionBottom.Value = Start(section.Page?.Bottom, 1);
			sectionTop.Enabled = sectionBottom.Enabled = ownMargins.Checked;
			ownMargins.CheckedChanged += (s, e) => sectionTop.Enabled = sectionBottom.Enabled = ownMargins.Checked;
			new ToolTip().SetToolTip(ownMargins, T("dlg.headerFooter.ownMarginsTip", "Ticked, this section's pages have these top and bottom margins; the other sections keep theirs. Page Setup sets the document's."));
		}
		layout.Controls.Add(SelectionRow());
		layout.Controls.Add(differentFirst);
		var otherPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
		otherPanel.Controls.Add(Band(header));
		otherPanel.Controls.Add(Band(footer));
		firstPanel.Controls.Add(Band(firstHeader));
		firstPanel.Controls.Add(Band(firstFooter));
		otherTab.Controls.Add(otherPanel);
		firstTab.Controls.Add(firstPanel);
		tabs.TabPages.Add(otherTab);
		tabs.TabPages.Add(firstTab);
		var inside = new Size(Math.Max(otherPanel.PreferredSize.Width, firstPanel.PreferredSize.Width), Math.Max(otherPanel.PreferredSize.Height, firstPanel.PreferredSize.Height));
		tabs.Size = new Size(inside.Width + LogicalToDeviceUnits(24), inside.Height + tabs.ItemSize.Height + LogicalToDeviceUnits(40));
		layout.Controls.Add(tabs);
		layout.Controls.Add(Position());
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(760), 0), Margin = new Padding(3, 6, 3, 6),
			Text = T("dlg.headerFooter.note", "Fields: {title} document name, {file} full path, {page} page number, {pages} number of pages, {sectionpages} pages in the section, {date}, {time}.\n"
				+ "Example: Page {page} of {pages}. Enter starts a new line in a part; a blank line at the end leaves space below a header.\n"
				+ "They are printed, and saved with the document: Word and OpenDocument files get them as their own header and footer.")
		});
		layout.Controls.Add(makeDefault);

		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var defaults = Dialogs.Button(this, T("dlg.headerFooter.defaults", "Defaults"));
		defaults.Click += (s, e) => LoadFrom(defaults_);
		var buttons = Dialogs.Row(this, 0, defaults, ok, cancel);
		layout.Controls.Add(buttons);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		differentFirst.CheckedChanged += (s, e) => EnableAll();
		foreach (var band in Bands) foreach (var box in band.Parts) Wire(box);
		LoadFrom(bands);
		if (section != null) {
			// (the same as the section before's: shown, not to change; typing in it gives the section its own)
			void Same() { tabs.Enabled = differentFirst.Enabled = !sameAsPrevious.Checked; }
			sameAsPrevious.CheckedChanged += (s, e) => Same();
			Same();
			differentFirst.Text = T("dlg.headerFooter.differentFirstSection", "Different first page: the section's first page has a header and footer of its own (or none), its other pages theirs");
		}
	}

	PageBands defaults_ = new();

	// (iteration 55) B, I, U and Font... for the selected words, in the part box last typed in; with nothing selected, for
	// the whole header or footer
	FlowLayoutPanel SelectionRow() {
		var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(3, 0, 3, 4) };
		row.Controls.Add(new Label { Text = T("dlg.headerFooter.selectedWords", "Selected words:"), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 3, 0) });
		// (iteration 55: buttons that stay pressed while the selection, or the caret's typing, has it)
		styleButtons = new RichFormat.Buttons(row, Font, LogicalToDeviceUnits, Apply);
		var font = new Button { Text = T("dlg.headerFooter.font", "Font..."), AutoSize = true, TabStop = false, AccessibleName = T("dlg.headerFooter.selectionFontName", "Selection font") };
		font.Click += (s, e) => ChooseSelectionFont();
		row.Controls.Add(font);
		row.Controls.Add(new Label { Text = T("dlg.headerFooter.selectionNote", "Select words in a part to give them a font, size, bold, italic or underline of their own; with nothing selected, these format the whole header or footer."),
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(520), 0), Anchor = AnchorStyles.Left, Margin = new Padding(6, 2, 0, 0) });
		return row;
	}

	void Wire(RichTextBox box) {
		box.Enter += (s, e) => { lastPart = box; styleButtons.Show(box); };
		box.SelectionChanged += (s, e) => { if (box == lastPart) styleButtons.Show(box); };
		box.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) e.IsInputKey = true; };    // (a new line, not OK)
		box.KeyDown += (s, e) => {
			FontStyle? flag = e.Control && !e.Alt && !e.Shift ? e.KeyCode switch { Keys.B => FontStyle.Bold, Keys.I => FontStyle.Italic, Keys.U => FontStyle.Underline, _ => null } : null;
			if (flag is { } f) { lastPart = box; Toggle(f); e.SuppressKeyPress = e.Handled = true; }
		};
	}

	BandBox BandOf(RichTextBox box) => Bands.First(b => b.Parts.Contains(box));
	static BandFont FontOrPlain(BandBox band) => band.Font ?? BandFont.Plain;

	// (Ctrl+B, I, U: on unless all of the selection has it already, as Word does)
	void Toggle(FontStyle flag) {
		if (lastPart is { } box) Apply(flag, !RichFormat.Has(box, flag));
	}

	// B, I or U on or off: for the selected words; with nothing selected, for the whole header or footer
	void Apply(FontStyle flag, bool on) {
		if (lastPart is not { } box) { styleButtons.Show(null); return; }
		if (box.SelectionLength == 0) {
			var band = BandOf(box);
			var style = on ? FontOrPlain(band).Style | flag : FontOrPlain(band).Style & ~flag;
			ChangeBand(band, FontOrPlain(band) with { Bold = style.HasFlag(FontStyle.Bold), Italic = style.HasFlag(FontStyle.Italic), Underline = style.HasFlag(FontStyle.Underline) });
		} else RichFormat.Set(box, flag, on);
		box.Focus();
		styleButtons.Show(box);
	}

	void ChooseSelectionFont() {
		if (lastPart is not { } box) return;
		if (box.SelectionLength == 0) { ChooseFont(BandOf(box)); return; }
		var current = box.SelectionFont is { } f ? FontOf(f) : FontOrPlain(BandOf(box));
		if (Pick(current) is not { } chosen) return;
		int start = box.SelectionStart, length = box.SelectionLength;
		box.SelectionFont = chosen.Create(chosen.Size);
		box.Select(start, length);
		box.Focus();
	}

	// Windows' font dialog (null: cancelled)
	BandFont? Pick(BandFont current) {
		using var dlg = new FontDialog { ShowColor = false, ShowEffects = true, FontMustExist = true, AllowVerticalFonts = false, MinSize = 4, MaxSize = 72 };
		using var font = current.Create(current.Size);    // (alive while the dialog is open: it reads the font when it opens)
		dlg.Font = font;
		return dlg.ShowDialog(this) == DialogResult.OK ? FontOf(dlg.Font) : null;
	}

	static BandFont FontOf(Font font) => BandFont.From(font) with { Size = (float)Math.Round(font.SizeInPoints * 2) / 2 };

	void ChooseFont(BandBox band) {
		if (Pick(band.Font is { } f && !PageBands.IsPlain(f) ? f : new BandFont { Name = "Times New Roman", Size = 10 }) is { } chosen) ChangeBand(band, chosen);
	}

	// The whole header's or footer's font (null: plain): its parts shown again in it, words with their own as they were.
	// plain: the words' own fonts dropped too.
	void ChangeBand(BandBox band, BandFont? font, bool plain = false) {
		var old = FontOrPlain(band);
		var parts = band.Parts.Select(b => GetPart(b, old)).ToList();
		band.Font = font;
		for (int i = 0; i < 3; i++) SetPart(band.Parts[i], plain ? BandText.Plain(parts[i]) : parts[i], FontOrPlain(band));
		ShowFonts();
	}

	// A part into its box: the text in the band's font, words with their own in theirs.
	static void SetPart(RichTextBox box, string part, BandFont band) {
		box.Font = band.Create(band.Size);
		var runs = BandText.Runs(part);
		box.Text = string.Concat(runs.Select(r => r.Text));
		int at = 0;
		foreach (var (text, format) in runs) {
			if (format != null) { box.Select(at, text.Length); box.SelectionFont = BandText.Resolve(band, format).Create(format.Size); }
			at += text.Length;
		}
		box.Select(0, 0);
	}

	// A part from its box: a character in the band's font has no mark, others their font's.
	static string GetPart(RichTextBox box, BandFont band) {
		string text = box.Text;
		if (text == "") return "";
		int start = box.SelectionStart, length = box.SelectionLength;
		var runs = new List<(string, RunFormat?)>();
		for (int i = 0; i < text.Length; i++) {
			box.Select(i, 1);
			runs.Add((text[i].ToString(), box.SelectionFont is { } f ? BandText.FormatOf(band, FontOf(f)) : null));
		}
		box.Select(start, length);
		return PageBands.Parts(new[] { BandText.FromRuns(runs) })[0];
	}

	GroupBox Band(BandBox band) {
		string title = band.IsHeader ? T("dlg.headerFooter.headerBox", "Header (above the text)") : T("dlg.headerFooter.footerBox", "Footer (below the text)");
		var box = new GroupBox { Text = title, AutoSize = true, Padding = new Padding(8), Dock = DockStyle.Fill };
		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Fill };
		table.Controls.Add(band.Show, 0, 0);
		table.SetColumnSpan(band.Show, 3);
		string[] labels = { T("dlg.headerFooter.left", "Left"), T("dlg.headerFooter.center", "Center"), T("dlg.headerFooter.right", "Right") };
		string name = band.First ? (band.IsHeader ? T("dlg.headerFooter.firstHeaderName", "First page header") : T("dlg.headerFooter.firstFooterName", "First page footer")) : (band.IsHeader ? T("dlg.headerFooter.headerName", "Header") : T("dlg.headerFooter.footerName", "Footer"));
		for (int i = 0; i < 3; i++) {
			table.Controls.Add(new Label { Text = labels[i], AutoSize = true }, i, 1);
			var part = band.Parts[i];
			part.Height = 3 * Font.Height + LogicalToDeviceUnits(10);    // (three lines in view; Enter a new one)
			part.Width = LogicalToDeviceUnits(260);
			part.AccessibleName = T("dlg.headerFooter.partName", "{0}, {1}", name, labels[i].ToLowerInvariant());
			table.Controls.Add(part, i, 2);
		}
		// Font... for all of it (name, size, bold, italic, underline), Plain (Segoe UI 8 pt in grey, the words' own fonts
		// dropped too), the line
		var fontButton = new Button { Text = T("dlg.headerFooter.font", "Font..."), AutoSize = true, AccessibleName = T("dlg.headerFooter.fontName", "{0} font", name) };
		var plainButton = new Button { Text = T("dlg.headerFooter.plain", "Plain"), AutoSize = true, AccessibleName = T("dlg.headerFooter.plainName", "{0} font plain", name) };
		fontButton.Click += (s, e) => ChooseFont(band);
		plainButton.Click += (s, e) => ChangeBand(band, null, plain: true);
		band.Line.AccessibleName = T("dlg.headerFooter.lineName", "{0} line", name);
		band.FontRow.Controls.AddRange(new Control[] { fontButton, plainButton, band.FontLabel, band.Line });
		table.Controls.Add(band.FontRow, 0, 3);
		table.SetColumnSpan(band.FontRow, 3);
		band.Show.CheckedChanged += (s, e) => EnableAll();
		box.Controls.Add(table);
		return box;
	}

	// Header from top, Footer from bottom (Word's names; the same on every page)
	GroupBox Position() {
		var box = new GroupBox { Text = T("dlg.headerFooter.position", "Position (from the page's edge, on every page)"), AutoSize = true, Padding = new Padding(8), Dock = DockStyle.Fill };
		var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
		string unit = metric ? "cm" : T("dlg.headerFooter.in", "in");
		foreach (var box2 in new[] { headerDistance, footerDistance }) {
			(box2.Maximum, box2.Increment) = metric ? (12.7m, 0.1m) : (5m, 0.05m);
			box2.Width = LogicalToDeviceUnits(70);
		}
		Label L(string text) => new() { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 3, 0) };
		row.Controls.AddRange(new Control[] { L(T("dlg.headerFooter.headerFromTop", "Header from top:")), headerDistance, L(unit + "      " + T("dlg.headerFooter.footerFromBottom", "Footer from bottom:")), footerDistance, L(unit) });
		box.Controls.Add(row);
		return box;
	}

	void ShowFonts() {
		foreach (var band in Bands) band.FontLabel.Text = PageBands.IsPlain(band.Font) ? T("dlg.headerFooter.plainFont", "Plain: Segoe UI 8 pt, grey") : band.Font!.Describe();
	}

	// Each band's parts and font row as its Show box has it; the first page's tab only with a first page of its own.
	void EnableAll() {
		bool first = differentFirst.Checked;
		otherTab.Text = first ? T("dlg.headerFooter.otherPages", "Other pages") : T("dlg.headerFooter.allPages", "All pages");
		firstTab.Text = first ? T("dlg.headerFooter.firstPage", "First page") : T("dlg.headerFooter.firstPageOff", "First page (off: check Different first page)");
		firstPanel.Enabled = first;
		foreach (var band in Bands) {
			foreach (var part in band.Parts) part.Enabled = band.Show.Checked;
			band.FontRow.Enabled = band.Show.Checked;
		}
	}

	void LoadFrom(PageBands bands) {
		differentFirst.Checked = bands.DifferentFirst;
		void Load(BandBox band, bool show, string[]? parts, BandFont? font, bool line) {
			band.Show.Checked = show;
			band.Font = font;
			band.Line.Checked = line;
			var p = PageBands.Parts(parts);
			for (int i = 0; i < 3; i++) SetPart(band.Parts[i], p[i], FontOrPlain(band));
		}
		Load(header, bands.ShowHeader, bands.Header, bands.HeaderFont, bands.HeaderLine);
		Load(footer, bands.ShowFooter, bands.Footer, bands.FooterFont, bands.FooterLine);
		// (the first page's: its own, or until it has them, the other pages' settings and no text)
		Load(firstHeader, bands.DifferentFirst ? bands.ShowHeaderOn(1) : bands.ShowHeader, bands.FirstHeader, bands.DifferentFirst ? bands.HeaderFontOn(1) : bands.HeaderFont, bands.DifferentFirst ? bands.HeaderLineOn(1) : bands.HeaderLine);
		Load(firstFooter, bands.DifferentFirst ? bands.ShowFooterOn(1) : bands.ShowFooter, bands.FirstFooter, bands.DifferentFirst ? bands.FooterFontOn(1) : bands.FooterFont, bands.DifferentFirst ? bands.FooterLineOn(1) : bands.FooterLine);
		ShowFonts();
		headerDistance.Value = Math.Min(headerDistance.Maximum, ToUnits(bands.HeaderFromEdge));
		footerDistance.Value = Math.Min(footerDistance.Maximum, ToUnits(bands.FooterFromEdge));
		EnableAll();
	}

	static string[] Typed(BandBox band) => band.Parts.Select(b => GetPart(b, FontOrPlain(band))).ToArray();
	// the first page's font: null when it is the other pages' (as files from before have it), plain set as such
	BandFont? FirstFont(BandBox first, BandBox other) => Equals(first.Font, other.Font) ? null : first.Font ?? BandFont.Plain;

	public PageBands Result() {
		bool first = differentFirst.Checked;
		return new() {
			ShowHeader = header.Show.Checked, Header = Typed(header),
			ShowFooter = footer.Show.Checked, Footer = Typed(footer),
			HeaderFont = PageBands.IsPlain(header.Font) ? null : header.Font, FooterFont = PageBands.IsPlain(footer.Font) ? null : footer.Font,
			HeaderLine = header.Line.Checked, FooterLine = footer.Line.Checked,
			DifferentFirst = first,
			FirstHeader = first ? Typed(firstHeader) : null,
			FirstFooter = first ? Typed(firstFooter) : null,
			ShowFirstHeader = first ? firstHeader.Show.Checked : null, ShowFirstFooter = first ? firstFooter.Show.Checked : null,
			FirstHeaderFont = first ? FirstFont(firstHeader, header) : null, FirstFooterFont = first ? FirstFont(firstFooter, footer) : null,
			FirstHeaderLine = first ? firstHeader.Line.Checked : null, FirstFooterLine = first ? firstFooter.Line.Checked : null,
			HeaderDistance = FromUnits(headerDistance.Value) is var hd && hd != PageBands.DefaultDistance ? hd : null,
			FooterDistance = FromUnits(footerDistance.Value) is var fd && fd != PageBands.DefaultDistance ? fd : null
		};
	}
}

// Insert > Footnote (iteration 51): a footnote's note, its paragraphs, with their bold, italic and underline.
sealed class FootnoteDialog : Form
{
	readonly RichTextBox text = AppTheme.OnPaper(new RichTextBox { Multiline = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, DetectUrls = false, Dock = DockStyle.Fill });
	readonly NumericUpDown size = new() { Minimum = 4, Maximum = 72, DecimalPlaces = 1, Increment = 0.5m, Width = 70 };
	const FontStyle Formats = FontStyle.Bold | FontStyle.Italic | FontStyle.Underline;

	/// The note, its formatting as Footnotes' marks.
	public string Note {
		get {
			var runs = new List<(string, FontStyle)>();
			int start = text.SelectionStart, length = text.SelectionLength;
			for (int i = 0; i < text.TextLength; i++) {
				text.Select(i, 1);
				char c = text.Text[i];
				runs.Add((c is '\r' or '\n' ? Footnotes.ParagraphMark.ToString() : c.ToString(), (text.SelectionFont?.Style ?? FontStyle.Regular) & Formats));
			}
			text.Select(start, length);
			return Footnotes.FromRuns(runs);
		}
	}

	// The note into the box, each run in its style.
	void SetNote(string note) {
		var runs = Footnotes.Runs(note).Select(r => (Text: r.Text.Replace(Footnotes.ParagraphMark, '\n'), r.Style)).ToList();
		text.Text = string.Concat(runs.Select(r => r.Text));
		int at = 0;
		foreach (var (run, style) in runs) {
			if (style != FontStyle.Regular) { text.Select(at, run.Length); text.SelectionFont = new Font(text.Font, style); }
			at += run.Length;
		}
		text.Select(text.TextLength, 0);
		text.SelectionFont = text.Font;
	}

	// Bold, italic or underline on the selection (off if all of it has it), or for the typing at the caret.
	void Toggle(FontStyle flag) {
		int start = text.SelectionStart, length = text.SelectionLength;
		if (length == 0) {
			var font = text.SelectionFont ?? text.Font;
			text.SelectionFont = new Font(font, font.Style ^ flag);
			return;
		}
		bool all = true;
		for (int i = start; i < start + length && all; i++) { text.Select(i, 1); all = (text.SelectionFont?.Style ?? 0).HasFlag(flag); }
		for (int i = start; i < start + length; i++) {
			text.Select(i, 1);
			var font = text.SelectionFont ?? text.Font;
			text.SelectionFont = new Font(font, all ? font.Style & ~flag : font.Style | flag);
		}
		text.Select(start, length);
		text.Focus();
	}

	/// The size of all the document's notes (its Footnote Text style).
	public float NoteSize => (float)size.Value;

	public FootnoteDialog(string title, string note, float noteSize) {
		Dialogs.Setup(this, title, sizable: true);
		Size = new Size(LogicalToDeviceUnits(520), LogicalToDeviceUnits(360));    // (iteration 81: 300 before the size note had a line of its own)
		MinimumSize = new Size(LogicalToDeviceUnits(360), LogicalToDeviceUnits(290));
		var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(480), 0), Margin = new Padding(3, 0, 3, 6),
			Text = note == "" ? T("dlg.footnote.noteNew", "The footnote's note (printed at the foot of the page). Enter starts a new paragraph; Ctrl+Enter is OK.") : T("dlg.footnote.noteEdit", "The footnote's note (empty: the footnote is taken out). Enter starts a new paragraph; Ctrl+Enter is OK.")
		}, 0, 0);
		var formatRow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0, 0, 0, 3) };
		foreach (var (label, flag, tip) in new[] { (T("dlg.format.b", "B"), FontStyle.Bold, T("dlg.format.bold", "Bold (Ctrl+B)")), (T("dlg.format.i", "I"), FontStyle.Italic, T("dlg.format.italic", "Italic (Ctrl+I)")), (T("dlg.format.u", "U"), FontStyle.Underline, T("dlg.format.underline", "Underline (Ctrl+U)")) }) {
			var button = new Button { Text = label, Font = new Font(Font, flag), Width = LogicalToDeviceUnits(30), Height = LogicalToDeviceUnits(26), TabStop = false, AccessibleName = tip.Split(' ')[0] };
			button.Click += (s, e) => Toggle(flag);
			new ToolTip().SetToolTip(button, tip);
			formatRow.Controls.Add(button);
		}
		text.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) e.IsInputKey = true; };    // (the box's, not the OK button's)
		text.KeyDown += (s, e) => {
			FontStyle? flag = e.Control && !e.Alt && !e.Shift ? e.KeyCode switch { Keys.B => FontStyle.Bold, Keys.I => FontStyle.Italic, Keys.U => FontStyle.Underline, _ => null } : null;
			if (flag is { } f) { Toggle(f); e.SuppressKeyPress = true; e.Handled = true; }
			else if (e.KeyCode == Keys.Enter && e.Control && !e.Alt) { e.SuppressKeyPress = true; DialogResult = DialogResult.OK; }    // (Enter: a new paragraph, as in Word; Ctrl+Enter: OK)
		};
		layout.RowCount = 5;
		layout.RowStyles.Insert(1, new RowStyle(SizeType.AutoSize));
		layout.Controls.Add(formatRow, 0, 1);
		SetNote(note);
		layout.Controls.Add(text, 0, 2);
		size.Value = (decimal)Math.Clamp(noteSize, 4, 72);
		var sizeRow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0, 6, 0, 0) };
		sizeRow.Controls.Add(new Label { Text = T("dlg.footnote.size", "Size of all the notes (pt):"), AutoSize = true, Margin = new Padding(3, 6, 6, 0) });
		sizeRow.Controls.Add(size);
		layout.Controls.Add(sizeRow, 0, 3);
		// (iteration 81: on a line of its own, wrapping: beside the size it ran off the dialog's right edge, in Spanish more)
		layout.RowCount = 6;
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.Controls.Add(new Label { Text = T("dlg.footnote.sizeNote", "(the Footnote Text style; Modify Style changes its font too)"), AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(480), 0), ForeColor = SystemColors.GrayText, Margin = new Padding(3, 2, 3, 6) }, 0, 4);
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		layout.Controls.Add(buttons, 0, 5);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		Shown += (s, e) => { text.Focus(); text.SelectionStart = text.TextLength; };
	}
}

// Insert > Cross-reference (iteration 50): the numbered paragraphs, and whether the reference shows the paragraph's own
// number or its full number.
// Insert > Table of Contents (iteration 52): the headings' levels it takes, its page numbers (shown or not, at the right
// margin after a leader or just after the text), and its title (a new table's only).
sealed class TableOfContentsDialog : Form
{
	readonly NumericUpDown levels = new() { Minimum = 1, Maximum = 6, Width = 50 };
	readonly CheckBox pages = new() { Text = T("dlg.toc.pages", "Show page numbers"), AutoSize = true, Checked = true };
	readonly CheckBox right = new() { Text = T("dlg.toc.right", "Right-align page numbers"), AutoSize = true, Checked = true };
	readonly ComboBox leader = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
	readonly TextBox title = new() { Width = 220 };
	static (string Value, string Shown)[] Leaders => new[] { ("dots", T("dlg.toc.dots", "Dots") + " ......."), ("dashes", T("dlg.toc.dashes", "Dashes") + " -------"), ("underline", T("dlg.toc.underline", "Underline") + " ____"), ("none", T("dlg.toc.none", "None")) };

	public Toc.Options Options => new((int)levels.Value, pages.Checked, right.Checked, Leaders[Math.Max(0, leader.SelectedIndex)].Value);
	public string TitleText => title.Text.Trim();

	public TableOfContentsDialog(Toc.Options options, bool isNew) {
		Dialogs.Setup(this, isNew ? T("dlg.toc.title", "Table of Contents") : T("dlg.toc.titleChange", "Table of Contents: change this one"));
		var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Dock = DockStyle.Fill };
		void Row(string label, Control control) {
			layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 8, 3) });
			layout.Controls.Add(control);
		}
		levels.Value = Math.Clamp(options.Levels, 1, 6);
		Row(T("dlg.toc.levels", "Show levels (headings 1 to):"), levels);
		layout.Controls.Add(pages); layout.SetColumnSpan(pages, 2);
		layout.Controls.Add(right); layout.SetColumnSpan(right, 2);
		foreach (var (_, shown) in Leaders) leader.Items.Add(shown);
		leader.SelectedIndex = Math.Max(0, Array.FindIndex(Leaders, l => l.Value == options.Leader));
		Row(T("dlg.toc.leader", "Tab leader:"), leader);
		pages.Checked = options.PageNumbers;
		right.Checked = options.RightAlign;
		void Enable() { right.Enabled = pages.Checked; leader.Enabled = pages.Checked && right.Checked; }
		pages.CheckedChanged += (s, e) => Enable();
		right.CheckedChanged += (s, e) => Enable();
		Enable();
		if (isNew) {
			title.Text = T("dlg.toc.contents", "Contents");
			Row(T("dlg.toc.titleBox", "Title (empty: none):"), title);
		} else {
			var note = new Label { Text = T("dlg.toc.changeNote", "The table at the caret is made again from the headings with these options."), AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(360), 0), Margin = new Padding(3, 8, 3, 3) };
			layout.Controls.Add(note); layout.SetColumnSpan(note, 2);
		}
		var ok = Dialogs.Ok(this);
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 10, ok, cancel);
		layout.Controls.Add(buttons); layout.SetColumnSpan(buttons, 2);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}

sealed class CrossReferenceDialog : Form
{
	readonly ListBox list = new() { IntegralHeight = false };
	readonly RadioButton number = new() { Text = T("dlg.crossRef.number", "Its number"), AutoSize = true }, full = new() { Text = T("dlg.crossRef.full", "Its full number (with the levels above it)"), AutoSize = true };
	// (a line of text that only repaints itself: a Label, at each selection, laid out and repainted the whole dialog,
	// which flickered)
	readonly TextLine shown = new() { Dock = DockStyle.Fill };
	readonly List<CrossRefs.Target> targets;

	sealed class BufferedTable : TableLayoutPanel
	{
		public BufferedTable() => DoubleBuffered = true;
	}

	sealed class TextLine : Control
	{
		public TextLine() => SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
		protected override void OnTextChanged(EventArgs e) => Invalidate();
		protected override void OnPaint(PaintEventArgs e) {
			e.Graphics.Clear(BackColor);
			TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
		}
	}

	/// The paragraph chosen (its index in the targets), and whether the full number is shown.
	public int? Chosen => list.SelectedIndex >= 0 ? list.SelectedIndex : null;
	public bool Full => full.Checked;

	public CrossReferenceDialog(List<CrossRefs.Target> targets, bool showFull, int selected) {
		this.targets = targets;
		Dialogs.Setup(this, T("dlg.crossRef.title", "Cross-reference"), sizable: true);
		Size = new Size(LogicalToDeviceUnits(620), LogicalToDeviceUnits(460));
		MinimumSize = new Size(LogicalToDeviceUnits(380), LogicalToDeviceUnits(300));

		var layout = new BufferedTable { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6 };
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
		for (int i = 0; i < 4; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.Controls.Add(new Label { Text = T("dlg.crossRef.referTo", "Refer to the numbered paragraph:"), AutoSize = true }, 0, 0);
		list.Dock = DockStyle.Fill;
		foreach (var t in targets) list.Items.Add(t.Full + "\t" + (t.Text.Length > 90 ? t.Text[..90] + "…" : t.Text));
		list.UseTabStops = true;
		layout.Controls.Add(list, 0, 1);
		var how = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 0) };
		how.Controls.Add(new Label { Text = T("dlg.crossRef.show", "Show:"), AutoSize = true });
		how.Controls.Add(number);
		how.Controls.Add(full);
		layout.Controls.Add(how, 0, 2);
		shown.Height = Font.Height + LogicalToDeviceUnits(4);
		shown.Margin = new Padding(3, LogicalToDeviceUnits(6), 3, 3);
		layout.Controls.Add(shown, 0, 3);
		layout.Controls.Add(new Label {
			AutoSize = true, MaximumSize = new Size(LogicalToDeviceUnits(560), 0), Margin = new Padding(3, 4, 3, 6),
			Text = T("dlg.crossRef.note", "The number follows the paragraph's when the paragraphs are renumbered (F9 updates them all). Ctrl+click a cross-reference to go to its paragraph.")
		}, 0, 4);
		var ok = new Button { Text = T("dlg.crossRef.insert", "Insert"), DialogResult = DialogResult.OK, AutoSize = true };
		var cancel = Dialogs.Cancel(this);
		var buttons = Dialogs.Row(this, 0, ok, cancel);
		layout.Controls.Add(buttons, 0, 5);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;

		(showFull ? full : number).Checked = true;
		list.SelectedIndexChanged += (s, e) => ShowWhat();
		number.CheckedChanged += (s, e) => ShowWhat();
		list.DoubleClick += (s, e) => { if (Chosen != null) { DialogResult = DialogResult.OK; Close(); } };
		if (targets.Count > 0) list.SelectedIndex = Math.Clamp(selected, 0, targets.Count - 1);
		ShowWhat();
	}

	void ShowWhat() {
		var target = Chosen is int i ? targets[i] : null;
		shown.Text = target == null ? "" : T("dlg.crossRef.inserts", "Inserts: {0}", CrossRefs.Shown(target, full.Checked ? "w" : "r"));
	}
}

// OLE drop target placed in front of RichEdit's own: dropped files that pass `accepts` go to `drop` (with the screen
// point); everything else (dragged text, other files) goes on to RichEdit's drop target unchanged.
sealed class FileDropTarget : IOleDropTarget
{
	const int DROPEFFECT_NONE = 0, DROPEFFECT_COPY = 1;
	readonly IOleDropTarget? inner;
	readonly Func<string[], bool> accepts;
	readonly Action<string[], Point?> drop;
	string[]? files;    // the files being dragged, while we handle the drag

	FileDropTarget(IOleDropTarget? inner, Func<string[], bool> accepts, Action<string[], Point?> drop) {
		this.inner = inner;
		this.accepts = accepts;
		this.drop = drop;
	}

	[DllImport("ole32")] static extern int OleInitialize(IntPtr reserved);
	[DllImport("ole32")] static extern int RegisterDragDrop(IntPtr hwnd, IOleDropTarget target);
	[DllImport("ole32")] static extern int RevokeDragDrop(IntPtr hwnd);
	[DllImport("user32", CharSet = CharSet.Unicode)] static extern IntPtr GetProp(IntPtr hwnd, string name);

	public static FileDropTarget? Register(IntPtr window, Func<string[], bool> accepts, Action<string[], Point?> drop, out string status) {
		OleInitialize(IntPtr.Zero);
		// OLE keeps the registered drop target of a window in this property.
		IntPtr existing = GetProp(window, "OleDropTargetInterface");
		var inner = existing != IntPtr.Zero ? Marshal.GetObjectForIUnknown(existing) as IOleDropTarget : null;
		RevokeDragDrop(window);
		var target = new FileDropTarget(inner, accepts, drop);
		int hr = RegisterDragDrop(window, target);
		status = $"RegisterDragDrop 0x{hr:X8}, RichEdit's drop target {(inner != null ? "wrapped" : "not found")}";
		return hr == 0 ? target : null;
	}

	public static string[]? GetFiles(object dataObject) {
		try {
			return new DataObject(dataObject).TryGetData<string[]>(DataFormats.FileDrop, out var files) ? files : null;
		} catch (Exception ex) when (ex is COMException or ExternalException or InvalidCastException) {
			return null;
		}
	}

	public int OleDragEnter(object dataObject, int keyState, POINTL point, ref int effect) {
		files = GetFiles(dataObject) is { } dragged && accepts(dragged) ? dragged : null;
		if (files != null) {
			effect = DROPEFFECT_COPY;
			return 0;
		}
		if (inner != null) return inner.OleDragEnter(dataObject, keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}

	public int OleDragOver(int keyState, POINTL point, ref int effect) {
		if (files != null) {
			effect = DROPEFFECT_COPY;
			return 0;
		}
		if (inner != null) return inner.OleDragOver(keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}

	public int OleDragLeave() {
		if (files != null) {
			files = null;
			return 0;
		}
		return inner?.OleDragLeave() ?? 0;
	}

	public int OleDrop(object dataObject, int keyState, POINTL point, ref int effect) {
		if (files != null) {
			var dropped = files;
			files = null;
			effect = DROPEFFECT_COPY;
			drop(dropped, new Point(point.X, point.Y));
			return 0;
		}
		if (inner != null) return inner.OleDrop(dataObject, keyState, point, ref effect);
		effect = DROPEFFECT_NONE;
		return 0;
	}
}

// Hidden top-level window that receives automation requests from other processes as WM_COPYDATA and sends the reply
// back to the sender's window (wParam) as WM_COPYDATA, before returning. Runs on the UI thread.
sealed class AutomationServer : NativeWindow, IDisposable
{
	public const int RequestId = 0x52544531, ReplyId = 0x52544532;    // "RTE1", "RTE2"
	const int WM_COPYDATA = 0x004A;
	readonly Func<string[], string> dispatch;

	[StructLayout(LayoutKind.Sequential)]
	struct COPYDATASTRUCT { public IntPtr dwData; public int cbData; public IntPtr lpData; }

	[DllImport("user32", CharSet = CharSet.Unicode)]
	static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref COPYDATASTRUCT lParam);

	public AutomationServer(Func<string[], string> dispatch) {
		this.dispatch = dispatch;
		CreateHandle(new CreateParams { Caption = "AhkWord automation server " + Environment.ProcessId });
	}

	protected override void WndProc(ref Message m) {
		if (m.Msg == WM_COPYDATA) {
			var data = Marshal.PtrToStructure<COPYDATASTRUCT>(m.LParam);
			if (data.dwData == (IntPtr)RequestId) {
				string request = Marshal.PtrToStringUni(data.lpData, data.cbData / 2).TrimEnd('\0');
				Reply(m.WParam, dispatch(request.Split('\u001F')));
				m.Result = (IntPtr)1;
				return;
			}
		}
		base.WndProc(ref m);
	}

	void Reply(IntPtr window, string text) {
		IntPtr buffer = Marshal.StringToHGlobalUni(text);
		try {
			var data = new COPYDATASTRUCT { dwData = (IntPtr)ReplyId, cbData = (text.Length + 1) * 2, lpData = buffer };
			SendMessage(window, WM_COPYDATA, Handle, ref data);
		} finally {
			Marshal.FreeHGlobal(buffer);
		}
	}

	public void Dispose() => DestroyHandle();
}

// Modeless Find / Replace window; it is hidden rather than closed so it keeps its settings.
sealed class FindReplaceForm : Form
{
	readonly WriterForm owner;
	readonly TextBox findBox = new(), replaceBox = new();
	readonly CheckBox matchCase = new() { Text = T("dlg.find.matchCase", "Match case"), AutoSize = true }, wholeWord = new() { Text = T("dlg.find.wholeWords", "Whole words"), AutoSize = true };
	readonly Label replaceLabel = new() { Text = T("dlg.find.replaceWith", "Replace with:"), AutoSize = true, Anchor = AnchorStyles.Left };
	readonly Button replaceButton = new() { Text = T("dlg.find.replace", "Replace"), AutoSize = true }, replaceAllButton = new() { Text = T("dlg.find.replaceAll", "Replace all"), AutoSize = true };

	public string FindText => findBox.Text;
	public string ReplaceText => replaceBox.Text;
	public bool MatchCase => matchCase.Checked;
	public bool WholeWord => wholeWord.Checked;

	public FindReplaceForm(WriterForm owner) {
		this.owner = owner;
		Text = T("dlg.find.title", "Find and Replace");
		FormBorderStyle = FormBorderStyle.FixedToolWindow;
		ShowInTaskbar = false;
		StartPosition = FormStartPosition.Manual;
		AutoSize = true;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		Padding = new Padding(8);
		findBox.Width = replaceBox.Width = LogicalToDeviceUnits(260);

		var findButton = new Button { Text = T("dlg.find.findNext", "Find next"), AutoSize = true };
		var closeButton = Dialogs.Cancel(this, T("dlg.close", "Close"));
		findButton.Click += (s, e) => owner.FindNext();
		replaceButton.Click += (s, e) => owner.ReplaceNext();
		replaceAllButton.Click += (s, e) => owner.ReplaceAll();
		closeButton.Click += (s, e) => Hide();

		var table = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Fill };
		table.Controls.Add(new Label { Text = T("dlg.find.findWhat", "Find what:"), AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
		table.Controls.Add(findBox, 1, 0);
		table.Controls.Add(findButton, 2, 0);
		table.Controls.Add(replaceLabel, 0, 1);
		table.Controls.Add(replaceBox, 1, 1);
		table.Controls.Add(replaceButton, 2, 1);
		var options = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
		options.Controls.AddRange(new Control[] { matchCase, wholeWord });
		table.Controls.Add(options, 1, 2);
		table.Controls.Add(replaceAllButton, 2, 2);
		table.Controls.Add(closeButton, 2, 3);
		Controls.Add(table);
		AcceptButton = findButton;
		CancelButton = closeButton;
		FormClosing += (s, e) => {
			if (e.CloseReason == CloseReason.UserClosing) {
				e.Cancel = true;
				Hide();
			}
		};
	}

	public void ShowFor(bool replace, string? initial) {
		replaceLabel.Visible = replaceBox.Visible = replaceButton.Visible = replaceAllButton.Visible = replace;
		if (initial != null) findBox.Text = initial;
		if (!Visible) {
			var area = owner.RectangleToScreen(owner.ClientRectangle);
			Show(owner);
			Location = new Point(area.Right - Width - LogicalToDeviceUnits(30), area.Top + LogicalToDeviceUnits(70));
		}
		Activate();
		findBox.SelectAll();
		findBox.Focus();
	}
}
