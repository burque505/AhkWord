# Roadmap

AhkWord is built in iterations, one feature or a handful of fixes at a time, each tested before it is committed. This page lists what is done and what is planned next. It has no dates: the order can change as people try AhkWord and say what they need. To suggest something, or to report a problem, open an [issue](https://github.com/burque505/AhkWord/issues).

The [known limitations](docs/reference/limitations.md) are listed in the documentation.

## Done

### Writing and formatting

- [x] Paragraph styles, with the Classic, Plain and Modern style sets, your own styles for new documents, and Copy Styles from Document
- [x] The Paragraph dialog: indents, spacing, line spacing, tab stops, keep with next, widow control, page break before
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
- [x] Documentation: [burque505.github.io/AhkWord](https://burque505.github.io/AhkWord/)

## Next

- [ ] **A signed AhkWord:** `AhkWord.exe` and its prebuilt C# host signed, so that Windows' Smart App Control (on by default on some new Windows 11 PCs) lets it run, and SmartScreen doesn't warn on the first start. Unsigned, as AutoHotkey itself is, Smart App Control blocks it with no way round but turning the setting off.
- [ ] **PDF export in the background:** keep working while a long document's PDF is made (it can't be much faster: most of the time is Microsoft Print to PDF's own work)

## Later, perhaps

- [ ] Tab leaders set in the Paragraph dialog drawn (RichEdit keeps them, and Word shows them, but AhkWord draws only a table of contents' leaders)
- [ ] Merged table cells
- [ ] A section break on the same page (Word's continuous break), worth it once there are columns or a page setup per section
- [ ] A winget package
- [ ] Windows on ARM
