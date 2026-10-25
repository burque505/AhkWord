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
- [x] Show Pages: the printed pages as sheets, with their headers, footers and footnotes
- [x] Headers and footers, with page numbers and a different first page
- [x] Line numbers, as Word numbers them, printed and on screen
- [x] Print Preview, printing, and PDF export with bookmarks for the headings

### Files

- [x] RTF, Word (.docx), OpenDocument (.odt), HTML, Markdown and plain text, through Word, LibreOffice, the built-in converter or pandoc
- [x] Markdown and HTML views: the document's source, edited side by side with the formatted text
- [x] Templates with fields, parts (signature blocks, captions), and a start page
- [x] Autosave and recovery, and a portable mode that keeps everything in AhkWord's own folder

### Automation

- [x] An automation API for AutoHotkey scripts (`Rte.InsertText(...)`), from other scripts or AhkWord.ahk itself, with events (`Rte.On("Saved", ...)`)
- [x] Macro recording, macro hotkeys, and macros in AhkWord's folder

### Getting AhkWord

- [x] A download that needs nothing installed: AhkWord carries its libraries and offers to download .NET for itself; a complete download has .NET inside
- [x] Documentation: [burque505.github.io/AhkWord](https://burque505.github.io/AhkWord/)

## Next

- [ ] **Snippets as AutoHotkey hotstrings:** type a short abbreviation (`sig`, `cos`) and it becomes a stored passage, a signature block or a certificate of service, drawing on the parts that templates already use
- [ ] **Script event hooks:** macros that run by themselves when a document is opened or saved, or a style changes, without a script of your own kept running
- [ ] **A command palette** (Ctrl+K): any command by typing part of its name
- [ ] **Faster PDF export** for long documents
- [ ] **Updates for AhkWord's own copy of .NET:** when AhkWord downloaded .NET for itself, an offer to update it when Microsoft releases a fix (a .NET installed for all of Windows gets these from Windows Update)

## Later, perhaps

- [ ] Tab leaders set in the Paragraph dialog drawn (RichEdit keeps them, and Word shows them, but AhkWord draws only a table of contents' leaders)
- [ ] Merged table cells
- [ ] A winget package, and a signed AhkWord.exe, so that Windows doesn't warn on the first start
- [ ] Windows on ARM
