# Pages and printing

## Page breaks, page setup

- **Insert > Page Break** (Ctrl+Enter) starts a new page. In the editor it shows as a dashed line.
- **Layout > Page Setup** sets the paper size, orientation and margins.

## Headers and footers

**Layout > Header & Footer** sets the text at the top and bottom of every page, with fields for the page number, the number of pages, the date and time, and the document's title or file name: `{page}`, `{pages}`, `{date}`, `{time}`, `{title}`, `{file}`. They are shown above and below the text in the editor (View > Header and footer shows or hides them; click one to change it), printed, and saved in the document: as real headers and footers in RTF, Word and OpenDocument files, where Word and LibreOffice keep the page numbers up to date.

**Several lines:** press Enter in a part to start a new line, for a firm's name and address in the footer, say. The band's lines are the parts' lines side by side: the left part's second line beside the right part's second line. Word and LibreOffice get a paragraph for each line, and headers and footers of several paragraphs are read back from them.

**Line below the header** and **Line above the footer** draw a rule between them and the text, as Word's border below a header does. A blank line at the end of a header part (Enter after its last line) leaves that much space below the header. For a line in the document itself, use **Insert > Horizontal Line**. Scripts use `HeaderLine`, `FooterLine` and `InsertHorizontalLine()`.

**Position** sets *Header from top* and *Footer from bottom*, the distance from the page's edge, as in Word (half an inch by default). A header or footer that reaches past its margin pushes the text in, as Word does. Scripts use `HeaderFromTop` and `FooterFromBottom` (inches).

**Words in a font of their own:** select words or lines in a part, then use **B**, **I**, **U** (or Ctrl+B, Ctrl+I, Ctrl+U) or **Font...** in the *Selected words* row at the top of the dialog: a firm's name in bold, say, with its address in the footer's font. With nothing selected, these format the whole header or footer. The words keep their fonts when printed, on the sheets, and in RTF and Word files.

**Font...** under the header's and the footer's parts chooses their font: its name, size, bold, italic and underline, for all three parts (and the first page's). It is printed, shown on Show Pages' sheets, and kept in RTF, Word and OpenDocument files. **Plain** goes back to the small grey Segoe UI of earlier versions, and drops the words' own fonts too. Scripts use `SetHeaderFont` and `SetFooterFont`.

**Different first page** gives page 1 a header and footer of its own in every respect: whether it has one at all, its text, font and line. Check it, and the dialog's tabs become *Other pages* and *First page*; what you set on one tab leaves the other alone. For a letterhead on page 1 only, show the header on the *First page* tab and leave *Show a header* off on *Other pages*. A first page with no footer, and page numbers from page 2 on, works the same way. Page 1's text starts below its own header, the other pages' below theirs. The first page's header and footer are kept in RTF, Word and OpenDocument files, and a first page set up in Word or LibreOffice is read back. Scripts use `DifferentFirstPage`, `SetFirstHeader`, `SetFirstFooter`, `ShowFirstHeader`, `ShowFirstFooter`, `SetFirstHeaderFont`, `SetFirstFooterFont`, `FirstHeaderLine` and `FirstFooterLine` (see the [API reference](../automation/api.md)).

## Print Layout and zoom

**View > Print Layout** (on at first) breaks the lines where they will break on the printed page, as Word and LibreOffice do, and shows the page on grey with its margins as on paper. A dotted line marked "Page 2", "Page 3" ... shows where each printed page begins. With Print Layout off, lines wrap at the window's edge.

Zoom, as in Word: **Page Width** (the default: the page fills the window's width, and follows it when the window is resized), **One Page** (a whole page fits the window), or your own with Zoom In, Zoom Out, 100% and Ctrl+mouse wheel. It is kept when you open another document.

## Show Pages

**View > Show Pages** (on at first) shows each printed page as a sheet of its own, as in Word: its header and footer in its margins, with grey between the sheets. You edit as usual: typing, the caret and the end of the document behave as without it. The sheets are only the view's: saving, copying, printing and the other formats see the document as it is.

The sheets follow your typing a second behind: they are worked out again a second after you stop, so while you type, a page's end can sit a few lines from where it will print until then. A page that ends inside a table is marked with a "Page N" line instead of a sheet of its own.

## Multiple Pages, Print Preview, printing

- **View > Zoom > Multiple Pages** shows the printed pages side by side. They are read-only: double-click a page to edit it there.
- **File > Print > Print Preview** shows the pages as they will print, fitted to the screen, with page buttons and zoom.
- **File > Print** prints; **Layout > Export as PDF** makes a PDF through *Microsoft Print to PDF*. As Word's, the PDF has a bookmark for each heading, under the one of the level above it, each opening its page at the heading, and it opens with its bookmarks shown; a table of contents' entries are links to their headings' pages. (Scripts: `ExportPdf(path)`, or `ExportPdf(path, 0)` for the pages only.)
- While a PDF is made or a document prints, the status bar shows the page it is on (*Exporting as PDF: page 12 of 100*), and **Esc** stops it; a stopped export leaves no file. A long document takes a while, about 0.4 seconds a page, most of it Microsoft Print to PDF's own work.

### Export as PDF, or print to a PDF printer

Other PDF printers, such as CutePDF, PDF24 or PDFCreator, appear in **File > Print** like any printer once installed, and AhkWord prints to them as it prints on paper. They often make smaller files than Microsoft Print to PDF, and some are faster. What they can't have is **bookmarks and links**: Export as PDF adds those to the file Microsoft Print to PDF writes, but other PDF printers ask where to save in a dialog of their own, so AhkWord can't find the file to add them.

- **Export as PDF:** bookmarks for the headings, and the table of contents' entries as links. Best for long documents that will be read on screen.
- **Print to a PDF printer:** often a smaller file, without bookmarks or links.

(Some PDF printers' installers offer other software along the way: read each step before clicking Next.)

Printing follows the paragraphs' line and page break rules (keep with next, keep lines together, widow and orphan control, page break before), as Word does.

See also [Line numbers](line-numbers.md), which are printed too.
