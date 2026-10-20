# Known limitations

- **Table borders are single lines.** Double, dotted and other border styles from Word documents show and are saved as single lines of the same width and colour: RichEdit draws only single lines.
- **Cells can't be merged.** Tables from Word with merged cells open, but AhkWord can't merge or split cells.
- **Markdown has no table formatting or paragraph spacing.** Column widths, borders, shading, line spacing and space before and after paragraphs are kept in RTF, Word, OpenDocument and HTML files, not in Markdown.
- **Markdown table rows are single lines,** so tables whose cells hold paragraphs are best edited in the formatted view.
- **On-screen line numbers match the printed ones in Print Layout only.** With Print Layout off, lines wrap at the window's edge, so a line's number on screen can differ from its printed number; printing, Print Preview, PDF and `Rte.PrintedLineAt` always number the printed lines.
- **Show Pages is a preview, not reliable for typing yet.** The view can shift up and down while you type, most in a new or short document. Leave it off to write, and turn it on to look the pages over; a redesign is planned. A page that ends inside a table is marked with a "Page N" line instead of a sheet of its own.
- **Suppressed line numbers need RTF, or Word or LibreOffice as the converter.** HTML, Markdown, pandoc and the built-in converter keep the line numbering but not which paragraphs are left out. OpenDocument files have no first number: their lines start at 1.
- **AhkWord.exe shows AutoHotkey's icon in Explorer.** It is AutoHotkey unchanged; the running AhkWord shows its own icon in the window, the taskbar and the tray.
