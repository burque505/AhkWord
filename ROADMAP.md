# Roadmap

AhkWord is built in iterations, one feature or a handful of fixes at a time, each tested before it is committed. This page lists what is done and what is planned next. It has no dates: the order can change as people try AhkWord and say what they need. To suggest something, or to report a problem, open an [issue](https://github.com/burque505/AhkWord/issues).

The [known limitations](docs/reference/limitations.md) are listed in the documentation.

## Done

### Writing and formatting

- [x] Paragraph styles, with the Classic, Plain and Modern style sets, your own styles for new documents, and Copy Styles from Document
- [x] The Paragraph dialog: indents, spacing, line spacing, tab stops with leaders, keep with next (tables too), keep lines together, widow control, page break before
- [x] Lists, legal numbering (1., 1.1, 1.1.1) and numbered styles whose numbers survive cutting and pasting, with sub-levels
- [x] Tables, with column widths, borders, shading and alignment
- [x] Pictures, with alt text
- [x] Spelling as you type, with the Windows spell checker
- [x] Cross-references to numbered paragraphs, kept up to date as they are renumbered ("see paragraph 23(e)")
- [x] Footnotes, laid out at the foot of the page as Word lays them out, with their own size, paragraphs, bold, italic and underline
- [x] A table of contents, with page numbers, dot leaders and links to the headings

### Pages and printing

- [x] Print Layout (lines break as they will be printed), with Page Width, One Page and Multiple Pages
- [x] Printing and PDF export show their progress, page by page, and Esc stops them
- [x] Show Pages: the printed pages as sheets, with their headers, footers and footnotes
- [x] Headers and footers of several lines, with page numbers, a different first page, a font of their own (size, bold, italic, underline) and their distance from the page's edge
- [x] Line numbers, as Word numbers them, printed and on screen
- [x] Print Preview, printing, and PDF export with bookmarks for the headings

### Files

- [x] RTF, Word (.docx), OpenDocument (.odt), HTML, Markdown and plain text, through Word, LibreOffice, the built-in converter or pandoc
- [x] Markdown and HTML views: the document's source, edited side by side with the formatted text
- [x] Templates with fields, parts (signature blocks, captions), and a start page
- [x] Snippets: AutoHotkey hotstrings that turn an abbreviation (`;sig`) into a formatted part, its fields asked for; exported as a plain-text hotstring script for other programs
- [x] Autosave and recovery, and a portable mode that keeps everything in AhkWord's own folder

### Automation

- [x] An automation API for AutoHotkey scripts (`Rte.InsertText(...)`), from other scripts or AhkWord.ahk itself, with events (`Rte.On("Saved", ...)`)
- [x] Macro recording, macro hotkeys, and macros in AhkWord's folder
- [x] A command palette (Ctrl+Shift+P): any command, macro, style, snippet or recent document by typing part of its name
- [x] Event hooks: macros that run by themselves when AhkWord starts, a document is opened, made, saved, printed or closed, or a style changes, without a script of your own kept running

### Getting AhkWord

- [x] A download that needs nothing installed: AhkWord carries its libraries and offers to download .NET for itself; a complete download has .NET inside
- [x] AhkWord's own copy of .NET kept up to date: Microsoft's fixes offered as they come out
- [x] Fast starts: AhkWord is compiled once and kept compiled (about two seconds a start), and a download carries it compiled, so even its first start compiles nothing
- [x] Documentation: [burque505.github.io/AhkWord](https://burque505.github.io/AhkWord/)

## Next

- [ ] **A signed AhkWord:** `AhkWord.exe` and its prebuilt C# host signed, so that Windows' Smart App Control (on by default on some new Windows 11 PCs) lets it run, and SmartScreen doesn't warn on the first start. Unsigned, as AutoHotkey itself is, Smart App Control blocks it with no way round but turning the setting off.
- [ ] **PDF export in the background:** keep working while a long document's PDF is made (it can't be much faster: most of the time is Microsoft Print to PDF's own work)
- [x] **Format Painter:** copy the formatting of the text at the caret and paint it onto other text, as in Word (click once for one use, double-click to keep painting)
- [x] **Ruler:** View > Ruler shows the margins, and the indents and tab stops of the paragraph at the caret, as Word's ruler (for looking)
- [x] **Spanish:** the interface in Spanish (the setting `Language`), its texts in `lang\es.json` (all of AhkWord's own dialogs since iteration 81)
- [x] **An extension's own language file** (`<folder>\lang\<code>.json`, read over AhkWord's; `lang\check_lang.ahk` checks it too)
- [x] **Choosing the language** in AhkWord itself: Review > Language (and the command palette, `Rte.Language`), from the next start
- [ ] **Clicking on the ruler:** add a tab stop where the ruler is clicked; double-click to open Paragraph or Tabs
- [x] **AhkWord's own Page Setup:** margins, orientation and paper, the header's and footer's distances and the different first page in one dialog of AhkWord's, with a picture of the page; the paper saved with the document (RTF, Word); "Make these the default" for new documents (iteration 82)

## Later, perhaps

- [ ] Merged table cells
- [ ] A section's own paper, orientation and left and right margins shown and printed (they are kept in the file now; top and bottom margins are shown): RichEdit wraps the whole document to one width
- [ ] A winget package
- [ ] Windows on ARM
