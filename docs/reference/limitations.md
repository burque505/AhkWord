# Known limitations

- **Table borders are single lines.** Double, dotted and other border styles from Word documents show and are saved as single lines of the same width and colour: RichEdit draws only single lines.
- **Cells can't be merged.** Tables from Word with merged cells open, but AhkWord can't merge or split cells.
- **Markdown has no table formatting or paragraph spacing.** Column widths, borders, shading, line spacing and space before and after paragraphs are kept in RTF, Word, OpenDocument and HTML files, not in Markdown.
- **Markdown table rows are single lines,** so tables whose cells hold paragraphs are best edited in the formatted view.
- **On-screen line numbers match the printed ones in Print Layout only.** With Print Layout off, lines wrap at the window's edge, so a line's number on screen can differ from its printed number; printing, Print Preview, PDF and `Rte.PrintedLineAt` always number the printed lines.
- **Show Pages follows typing a second behind.** The sheets are worked out again a second after typing stops, so while you type, a page's end can sit a few lines from where it will print until then (and after an undo, a sheet's gap can be missing for that second). A page that ends inside a table is marked with a "Page N" line instead of a sheet of its own.
- **Suppressed line numbers need RTF, or Word or LibreOffice as the converter.** HTML, Markdown, pandoc and the built-in converter keep the line numbering but not which paragraphs are left out. OpenDocument files have no first number: their lines start at 1.
- **Tab leaders from the Paragraph dialog aren't drawn.** A tab stop's leader (dots, dashes) is kept, and Word and LibreOffice show it, but RichEdit draws one only now and then; AhkWord draws a table of contents' leaders itself, but not other tabs'.
- **A table of contents' page numbers sit a little inside the right margin** in AhkWord (RichEdit puts a number on a line of its own when its tab is at the margin); Word and LibreOffice get the tab at the margin.
- **LibreOffice opens AhkWord's PDFs with the text in its own fonts,** as it does any PDF it opens for editing, so lines can run into the margins there. PDF viewers show the pages as printed.
- **PDF export of a long document takes a while** (some 40 seconds for 100 pages): Microsoft Print to PDF prints it page by page, and AhkWord waits for it. The status bar shows the page it is on, and Esc stops it.
- **Find and Replace don't look in footnotes' notes.** A note is kept inside its footnote's marker, which Find and Replace leave alone (as they do the markers of cross-references and tables of contents), so a word only in a note isn't found; change it in the note's own dialog (right-click its number, Edit footnote).
- **x64 Windows only.**
- **AhkWord.exe shows AutoHotkey's icon in Explorer.** It is AutoHotkey unchanged; the running AhkWord shows its own icon in the window, the taskbar and the tray.
