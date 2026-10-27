# Automation API reference

Everything a script can call on AhkWord, as `Rte.<Name>` after `#Include "RteClient2.ahk"` (see [Automation from AutoHotkey](index.md)). Methods are called `Rte.Name(arguments)`; properties are read as `Rte.Name` and set as `Rte.Name := value`.

Types: *text*; *number*; *1/0* for yes or no; enumerations by name (`Heading1`, `LowerAlpha`); colours as `#RRGGBB` or a name. An argument with `= value` can be left out.

!!! note "Made from the program"
    This page is generated from AhkWord's source by `make_docs.py`, so it always matches the program. Tools > Automation API in AhkWord lists the same members, and the editor control's own as well.

## General

### Version

```autohotkey
value := Rte.Version
```

text · read-only

## Documents

### AutosaveSeconds

```autohotkey
value := Rte.AutosaveSeconds
Rte.AutosaveSeconds := value
```

number

Seconds between the recovery copies written while the document has unsaved changes (0: none; at least 10). Remembered.

### InsertTemplate

```autohotkey
Rte.InsertTemplate(template, fields = "", ask = 0)
```

*template*: text · *fields*: text · *ask*: 1/0 · returns: number

Inserts a template's content at the caret (a part, by name, such as "Signature Block", or a document, or a file's path), its fields filled in: fields as for NewFromTemplate. Returns how many fields were filled in.

### New

```autohotkey
Rte.New()
```

Starts a new, empty document (unsaved changes are discarded).

### Portable

```autohotkey
value := Rte.Portable
```

1/0 · read-only

Whether AhkWord runs in portable mode (a file named "portable" or "portable.txt" beside it): its settings, recovery copies and the converters' state are in its data folder.

### RecoveryFolder

```autohotkey
value := Rte.RecoveryFolder
```

text · read-only

The folder of the recovery copies (beside the settings).

### Templates

```autohotkey
Rte.Templates(parts = 0)
```

*parts*: 1/0 · returns: text

The templates' names, one a line: the documents (File &gt; New from Template), or with parts = true the parts (pieces for Insert from Template, in the Parts folders).

## Snippets

### AddSnippet

```autohotkey
Rte.AddSnippet(abbreviation, part, instant = 0, keepEnding = 0, caseSensitive = 0)
```

*abbreviation*: text · *part*: text · *instant*: 1/0 · *keepEnding*: 1/0 · *caseSensitive*: 1/0

Gives a part (by name: Rte.Templates(true) lists them; an RTF template) an abbreviation: typed in the document, then Space, Enter, Tab or punctuation (instant = 1: at once), it becomes the part. keepEnding = 1: the ending character stays after it; caseSensitive = 1: the abbreviation fires only typed with its case. An abbreviation it had already is changed.

### ApplyMyStyles

```autohotkey
Rte.ApplyMyStyles()
```

returns: number

Apply My Styles: the document gets your styles, the ones new documents get (saved with SaveStylesAsDefault, else your style set's): each of its styles that you have takes your definition, and its paragraphs your look (formatting of their own stays); styles of the document's own that you haven't got stay. One undo step. Returns how many paragraphs were restyled.

### ClearFormatting

```autohotkey
Rte.ClearFormatting()
```

Clear Formatting (Ctrl+Space): the selected text back to its paragraphs' style (all formatting of its own goes); with no selection, the caret's whole paragraph, its paragraph formatting (spacing, indents ...) too. One undo step.

### Close

```autohotkey
Rte.Close()
```

Closes the document and starts an empty one, as File &gt; Close does (unsaved changes are discarded, as by New; the window stays open).

### CopyStyles

```autohotkey
Rte.CopyStyles(path, names = "", forNewDocuments = 0)
```

*path*: text · *names*: text · *forNewDocuments*: 1/0 · returns: number

Copy Styles from Document: the styles named (by commas or lines; empty: all of them) come from another document into this one. A style this document has takes the copied definition and its paragraphs the new look (formatting of their own stays); others are added. With forNewDocuments they go into your default styles too (the ones new documents get), the others of those staying as they are. Returns how many styles were copied.

### DeleteStyle

```autohotkey
Rte.DeleteStyle(name)
```

*name*: text

Deletes one of your own styles; its paragraphs become Normal (one undo step).

### ExportSnippets

```autohotkey
Rte.ExportSnippets(path)
```

*path*: text · returns: text

Writes the snippets as an AutoHotkey v2 script of plain-text hotstrings, to use them in any program (their fields as their defaults or [their names]). Returns its full path.

### FileConverter

```autohotkey
value := Rte.FileConverter
```

text · read-only

The converter(s) chosen for the current Word or OpenDocument file when it was opened or saved ("" = automatic); Save uses them.

### FileName

```autohotkey
value := Rte.FileName
```

text · read-only

The full path of the current file ("" for a new document).

### FillFields

```autohotkey
Rte.FillFields(fields = "", ask = 0)
```

*fields*: text · *ask*: 1/0 · returns: number

Fills in the {{fields}} left in the document: fields as for NewFromTemplate. Returns how many were filled in.

### FireSnippet

```autohotkey
Rte.FireSnippet(abbreviation, ending = "")
```

*abbreviation*: text · *ending*: text

For AhkWord.ahk: a snippet's hotstring fired, with the ending character typed after it ("" for one that fires at once). Returns at once; the snippet goes in once the editor has what was typed.

### GetHtml

```autohotkey
Rte.GetHtml()
```

returns: text

The document as HTML, with list numbering that RichEditorNET's HTML leaves out (start numbers, restarts, punctuation such as "a)"; the punctuation needs the &lt;style&gt; element at the start).

### Html

```autohotkey
value := Rte.Html
Rte.Html := value
```

text

The document as HTML (get), or replaced by HTML (set); see GetHtml and SetHtml.

### InsertSnippet

```autohotkey
Rte.InsertSnippet(abbreviation, fields = "", ask = 0)
```

*abbreviation*: text · *fields*: text · *ask*: 1/0 · returns: number

Inserts a snippet's part at the caret (as Insert from Template does): fields as for InsertTemplate. Returns how many fields were filled in.

### ModifyStyle

```autohotkey
Rte.ModifyStyle(name, settings)
```

*name*: text · *settings*: text · returns: number

Changes a style: settings as "Name=Value" lines ("Font=Georgia\nSize=12\nBefore=6"); its paragraphs change too (one undo step). Returns how many paragraphs have it.

### NewFromTemplate

```autohotkey
Rte.NewFromTemplate(template, fields = "", ask = 0)
```

*template*: text · *fields*: text · *ask*: 1/0

Makes an untitled document from a template (by name, or a file's path); unsaved changes are discarded. fields: "Name=Value" lines for its {{Name}} fields; the others get the values typed last time (or their defaults). With ask = true, the fields not given are asked for in the dialog instead.

### NewStyle

```autohotkey
Rte.NewStyle(name, settings = "")
```

*name*: text · *settings*: text

A new style, from the paragraph at the caret (with settings, as ModifyStyle takes them, on top), given to the paragraphs the selection is in.

### NewWindow

```autohotkey
Rte.NewWindow(path = "")
```

*path*: text

Opens another AhkWord window (File &gt; New Window), with a new document or the file given. The new window is a separate editor: scripts reach it once it is the AhkWord used last (or with its RTE_WRITER_HWND).

### Open

```autohotkey
Rte.Open(path, via = "")
```

*path*: text · *via*: text

Opens an .html/.htm, .md/.markdown, .rtf, .docx/.doc, .odt or .txt file (unsaved changes are discarded). Word and OpenDocument files are converted by Word, openxml, LibreOffice or pandoc (see Converters and LastConverter); via names the converter(s) to use, e.g. "openxml" (Save then uses them too).

### RemoveSnippet

```autohotkey
Rte.RemoveSnippet(abbreviation)
```

*abbreviation*: text · returns: 1/0

Takes a snippet's abbreviation away (its part stays). Returns whether there was one.

### Save

```autohotkey
Rte.Save()
```

returns: text

Saves to the current file and returns its path.

### SaveAs

```autohotkey
Rte.SaveAs(path, via = "")
```

*path*: text · *via*: text · returns: text

Saves as HTML, Markdown, RTF, Word, OpenDocument or text, chosen by the extension (.html, .md, .rtf, .docx, .odt, .txt), and returns the full path. .docx and .odt are converted (see Converters and LastConverter); via names the converter(s) to use, e.g. "openxml". HTML and Markdown pictures are saved as files in a &lt;name&gt;_images folder; via = "inside" keeps an HTML file's pictures inside it (as data: URIs), as one file.

### SaveSelectionAsSnippet

```autohotkey
Rte.SaveSelectionAsSnippet(part, abbreviation, instant = 0, keepEnding = 0, caseSensitive = 0)
```

*part*: text · *abbreviation*: text · *instant*: 1/0 · *keepEnding*: 1/0 · *caseSensitive*: 1/0

Saves the selection, formatting and all, as a new part (in your own templates' Parts folder, as part.rtf), and gives it an abbreviation (as AddSnippet).

### SaveStylesAsDefault

```autohotkey
Rte.SaveStylesAsDefault()
```

New documents get this document's styles (your changes included).

### SetHtml

```autohotkey
Rte.SetHtml(html)
```

*html*: text

Replaces the document with HTML (a whole page or a part of one), including the list numbering and tables GetHtml writes.

### SetStyle

```autohotkey
Rte.SetStyle(name)
```

*name*: text

Gives the paragraphs the selection is in a style (by name: Rte.Styles() lists them), in one undo step: their formatting becomes the style's, except formatting given to part of a paragraph (a bold word).

### SnippetHotstrings

```autohotkey
value := Rte.SnippetHotstrings
```

text · read-only

For AhkWord.ahk: the snippets' hotstrings, one a line (their AutoHotkey options, a tab, the abbreviation), and the file they are kept in (read again when it changes).

### Snippets

```autohotkey
value := Rte.Snippets
```

text · read-only

The snippets, one a line: the abbreviation, a tab, the part it inserts, a tab, how it fires ("ending": on Space, Enter, Tab or punctuation; "ending kept": the character stays after it; "instant": at once), a tab, "case" if the abbreviation's case must match (else "any").

### SnippetsFile

```autohotkey
value := Rte.SnippetsFile
```

text · read-only

### StartPage

```autohotkey
value := Rte.StartPage
Rte.StartPage := value
```

1/0

The start page: showing (true) or the document (false).

### Style

```autohotkey
value := Rte.Style
Rte.Style := value
```

text

The style of the paragraphs the selection is in ("" if they differ); setting it gives them the style.

### Styles

```autohotkey
Rte.Styles()
```

returns: text

The document's styles' names, one a line (the Style box's list).

### StyleSet

```autohotkey
value := Rte.StyleSet
Rte.StyleSet := value
```

text

The document's style set ("Classic", "Plain", "Modern"; "" for a document's own styles, from a file made elsewhere: ApplyMyStyles gives it yours); setting it gives the document that set's styles (your own styles stay). New documents keep theirs: SaveStylesAsDefault.

### StylesIn

```autohotkey
Rte.StylesIn(path)
```

*path*: text · returns: text

Iteration 40: the paragraph styles a document has (AhkWord's own RTF and HTML, a Word .docx, other RTF), one name a line, without opening it.

### TemplatesFolder

```autohotkey
value := Rte.TemplatesFolder
```

text · read-only

Your own templates' folder (beside the settings); the shipped ones are in the templates folder beside AhkWord.

### UpdateStyleToMatch

```autohotkey
Rte.UpdateStyleToMatch()
```

returns: text

The caret paragraph's style made to look like the paragraph; its other paragraphs change too. Returns its name.

## Rich Text Format, Word and OpenDocument

### Converters

```autohotkey
Rte.Converters()
```

returns: text

Which converters for Word and OpenDocument files are available.

### ExportDocx

```autohotkey
Rte.ExportDocx(path, via = "")
```

*path*: text · *via*: text · returns: text

Writes a copy of the document as Word (.docx); the current file name does not change. via names the converters to try, e.g. "openxml" or "word,pandoc" (default: Word if it isn't running, openxml, LibreOffice, pandoc). Returns the full path.

### ExportOdt

```autohotkey
Rte.ExportOdt(path, via = "")
```

*path*: text · *via*: text · returns: text

Writes a copy of the document as OpenDocument text (.odt), like ExportDocx (default order: LibreOffice, Word if it isn't running, pandoc).

### LastConverter

```autohotkey
value := Rte.LastConverter
```

text · read-only

The converter used by the last Open, Save, SaveAs or Export of a Word or OpenDocument file ("word", "libreoffice", "openxml" or "pandoc"; "" if none was needed).

### Rtf

```autohotkey
value := Rte.Rtf
```

text · read-only

The document as RTF (get), or replaced by RTF (set). RTF keeps everything the editor has.

## Print Layout

### MultiplePages

```autohotkey
value := Rte.MultiplePages
Rte.MultiplePages := value
```

1/0

Multiple Pages: the printed pages side by side, read-only, in the editor's place (true), or editing (false).

### PageView

```autohotkey
value := Rte.PageView
Rte.PageView := value
```

1/0

Page view (View &gt; Show Pages): in Print Layout, each printed page as a sheet of its own, with its header and footer and a band between the sheets (1); 0: one long page, a line where each page begins. Remembered.

### PageViewReport

```autohotkey
Rte.PageViewReport()
```

returns: text

The page view as it is now, for a bug report: the program's folder, the zoom, where the printed pages start, each gap (positions, line spacing) and any raised text that isn't a gap's. Numbers only: none of the text.

### PrintLayout

```autohotkey
value := Rte.PrintLayout
Rte.PrintLayout := value
```

1/0

Print Layout (1): lines break where they will on the printed page, the page shown on grey; 0: lines wrap at the window's width (WordWrap). Remembered.

### ZoomMode

```autohotkey
value := Rte.ZoomMode
Rte.ZoomMode := value
```

text

Print Layout's zoom: "PageWidth" (the page's width fills the window), "OnePage" (a page's height fits) or "Manual" (ZoomFactor, as Zoom In / Out / 100% and Ctrl+wheel set it). Remembered.

## Line spacing

### LineSpacing

```autohotkey
value := Rte.LineSpacing
```

text · read-only

The line spacing of the selected paragraphs: "1.5" (lines), "exactly 14pt", "at least 12pt", or "mixed".

### SetLineSpacing

```autohotkey
Rte.SetLineSpacing(spacing)
```

*spacing*: text

Sets the line spacing of the selected paragraphs (every paragraph the selection touches): a number of lines ("1", "1.15", "1.5", "2"), or "exactly 14pt" / "at least 12pt". One undo step.

### SpaceAfter

```autohotkey
value := Rte.SpaceAfter
Rte.SpaceAfter := value
```

number

The space after the selected paragraphs, in points (-1 if they differ), as SpaceBefore.

### SpaceBefore

```autohotkey
value := Rte.SpaceBefore
Rte.SpaceBefore := value
```

number

The space before the selected paragraphs, in points (-1 if they differ); setting it sets it for them (one undo step).

## The Paragraph dialog's formats

### DefaultTabStop

```autohotkey
value := Rte.DefaultTabStop
Rte.DefaultTabStop := value
```

number

The spacing of the document's default tab stops (where Tab goes past the paragraph's own), in points.

### ParagraphFormat

```autohotkey
Rte.ParagraphFormat()
```

returns: text

The selected paragraphs' format as "Key=Value" lines, what they share (a key they differ in is left out): Align, LeftIndent, RightIndent, FirstLine (negative: hanging), Before, After (points), LineSpacing, KeepWithNext, KeepTogether, WidowControl, PageBreakBefore, NoLineNumber (1 or 0), Tabs (72/right/dots,...), and their style's Contextual (no space between paragraphs of the same style) and Level (0: body text, 1-6: a heading).

### SetParagraphFormat

```autohotkey
Rte.SetParagraphFormat(settings)
```

*settings*: text

Gives the selected paragraphs a format: "Key=Value" lines as ParagraphFormat gives them (or with ; between them), measures in points or with a unit: Rte.SetParagraphFormat("LeftIndent=0.5in; FirstLine=-0.5in; KeepTogether=1"). What isn't given stays as each paragraph has it. One undo step.

### SetTabs

```autohotkey
Rte.SetTabs(tabs)
```

*tabs*: text

Gives the selected paragraphs these tab stops (and no others): "72/right/dots,3in/decimal" ("" clears them).

### Tabs

```autohotkey
value := Rte.Tabs
```

text · read-only

The selected paragraphs' tab stops: "72/left/none,216/right/dots" (positions in points from the left indent's edge; "" none; "mixed" if the paragraphs differ).

## Line numbers

### LineNumberAt

```autohotkey
Rte.LineNumberAt(position = -1)
```

*position*: number · returns: number

The number of the line with the character at position (-1: the caret's line), whether or not count-by shows it; 0 if the line has none (line numbers off, or a line in a table, a page break or a suppressed paragraph).

### LineNumbers

```autohotkey
value := Rte.LineNumbers
Rte.LineNumbers := value
```

text

The document's line numbering, as in Word's Layout &gt; Line Numbers: "none", "continuous" (counting through the document) or "page" (restarting on each page). Setting it keeps the start, count-by and distance. Saved with the document.

### LineNumberSettings

```autohotkey
Rte.LineNumberSettings()
```

returns: text

The line numbering's settings: "none", or e.g. "continuous, start 1, count by 1, distance auto".

### LineNumbersSuppressed

```autohotkey
Rte.LineNumbersSuppressed()
```

returns: 1/0

Whether the line numbers of the selected paragraphs are suppressed (all of them).

### PrintedLineAt

```autohotkey
Rte.PrintedLineAt(position = -1)
```

*position*: number · returns: text

Where the character at position (-1: the caret) is printed, as "page:line", e.g. "3:12": its page, and its line's printed number (0 if the line has none, or line numbers are off). Laid out for the printer, as Print does (the editor's own lines can wrap elsewhere).

### SetLineNumbers

```autohotkey
Rte.SetLineNumbers(mode = "continuous", start = 1, countBy = 1, distance = 0)
```

*mode*: text · *start*: number · *countBy*: number · *distance*: number

Sets the line numbering: mode as LineNumbers; start: the first line's number; countBy: show only the numbers that are multiples of it (5: 5, 10, 15 ...); distance: from the text in points, for Word and LibreOffice (0 = automatic).

### SuppressLineNumbers

```autohotkey
Rte.SuppressLineNumbers(suppress = 1)
```

*suppress*: 1/0

Leaves the selected paragraphs' lines without numbers, and not counted (suppress = 1), or numbers them again (0), as Word's Suppress for Current Paragraph. One undo step.

## Tables

### DeleteColumn

```autohotkey
Rte.DeleteColumn()
```

Deletes the caret's column (and the table, if it was the only column).

### DeleteRow

```autohotkey
Rte.DeleteRow()
```

Deletes the caret's row (and the table, if it was the only row).

### DeleteSelectedRows

```autohotkey
Rte.DeleteSelectedRows()
```

returns: number

Deletes the table rows the selection covers completely (the whole table if that is all of them), as the Delete key does; returns how many were deleted.

### DeleteTable

```autohotkey
Rte.DeleteTable()
```

Deletes the table the caret is in.

### InsertColumn

```autohotkey
Rte.InsertColumn(right = 1)
```

*right*: 1/0

Inserts an empty column right (right = 1) or left (0) of the caret's column; the table keeps its width.

### InsertRow

```autohotkey
Rte.InsertRow(below = 1)
```

*below*: 1/0

Inserts an empty row below (below = 1) or above (0) the caret's row.

### InsertTable

```autohotkey
Rte.InsertTable(rows = 3, columns = 3, headerRow = 1)
```

*rows*: number · *columns*: number · *headerRow*: 1/0

Inserts a table at the caret (headerRow = 1 makes the first row bold) and puts the caret in its first cell.

### InTable

```autohotkey
Rte.InTable()
```

returns: 1/0

Whether the caret is in a table.

### SelectedRows

```autohotkey
Rte.SelectedRows()
```

returns: number

The number of table rows the selection covers completely (0 if it is not whole rows of one table).

## Pictures' alt text

### InsertPageBreak

```autohotkey
Rte.InsertPageBreak()
```

Inserts a page break at the caret (replacing the selection): the text after it starts a new page when printed, and in Word and LibreOffice. A numbered list it splits goes on with its numbers after it. Ctrl+Enter.

### PageBreakCount

```autohotkey
Rte.PageBreakCount()
```

returns: number

The number of page breaks in the document.

### PictureAltText

```autohotkey
Rte.PictureAltText(position)
```

*position*: number · returns: text

The alt text of the picture at a position ("" if it has none).

### SetPictureAltText

```autohotkey
Rte.SetPictureAltText(position, alt)
```

*position*: number · *alt*: text

Gives the picture at a position a new alt text ("" for none). RichEditorNET can't change a picture's alt text, so the picture is replaced by the same picture (image, size and link) with the new one, in one undo step.

### TableCell

```autohotkey
Rte.TableCell()
```

returns: text

The caret's table position as "row,column" (from 1), or "" outside tables.

### TableCount

```autohotkey
value := Rte.TableCount
```

number · read-only

The number of tables in the document.

## Table properties

### SetCellBorders

```autohotkey
Rte.SetCellBorders(borders = "all", width = 0.5f, color = "", wholeTable = 0)
```

*borders*: text · *width*: number · *color*: text · *wholeTable*: 1/0

Sets the borders of the selected cells (the whole table with wholeTable = 1): "all", "outside", "inside", "none", "top", "bottom", "left" or "right", a line width in points and a color (a name such as "black", or "#RRGGBB").

### SetCellPadding

```autohotkey
Rte.SetCellPadding(points)
```

*points*: number

Sets the space left and right of the text in the cells of the caret's table, in points.

### SetCellShading

```autohotkey
Rte.SetCellShading(color = "", wholeTable = 0)
```

*color*: text · *wholeTable*: 1/0

Sets the background color of the selected cells (the whole table with wholeTable = 1); "" or "none" removes it.

### SetCellVerticalAlignment

```autohotkey
Rte.SetCellVerticalAlignment(alignment = "top", wholeTable = 0)
```

*alignment*: text · *wholeTable*: 1/0

Sets where the text sits in the selected cells (the whole table with wholeTable = 1): "top", "center" or "bottom".

### SetColumnWidths

```autohotkey
Rte.SetColumnWidths(widths)
```

*widths*: text

Sets the widths of the columns of the caret's table, left to right: a list such as "2in, 1.5in, 3cm" (pt, in, cm, mm; points without a unit). Rows with another number of cells keep theirs.

### SetRowHeight

```autohotkey
Rte.SetRowHeight(points, exact = 0, wholeTable = 0)
```

*points*: number · *exact*: 1/0 · *wholeTable*: 1/0

Sets the height of the rows the selection is in (all rows with wholeTable = 1), in points: 0 is automatic (as high as their text), otherwise at least that high, or exactly that high (exact = 1).

### SetTableAlignment

```autohotkey
Rte.SetTableAlignment(alignment)
```

*alignment*: text

Puts the caret's table at the "left", "center" or "right" of the page.

### TableFormat

```autohotkey
Rte.TableFormat()
```

returns: text

The caret's table's format, one line per row: alignment, height and padding, then each cell's width, borders, shading and vertical alignment (lengths in points).

## Font sizes

### DecreaseFontSize

```autohotkey
Rte.DecreaseFontSize(points = 1)
```

*points*: number

Makes each part of the selection points smaller (keeping their differences, down to 1 point).

### IncreaseFontSize

```autohotkey
Rte.IncreaseFontSize(points = 1)
```

*points*: number

Makes each part of the selection points larger (keeping their differences).

### SetFontSize

```autohotkey
Rte.SetFontSize(size)
```

*size*: number

Sets the font size of the selection (free-form mode). Where the selection covers a paragraph's whole text, its paragraph mark gets the size too, so that its list number (and a legal number) follows, as in Word.

## Legal numbering

### AdoptTypedNumbers

```autohotkey
Rte.AdoptTypedNumbers()
```

returns: number

Makes paragraphs that start with typed compound numbers ("1.2&lt;tab&gt;Text") a legal list; returns how many changed. Opening a document does this by itself.

### ConvertTypedNumbers

```autohotkey
Rte.ConvertTypedNumbers()
```

returns: number

Paragraphs that begin with typed list numbers or bullets ("1. ", "a) ", "(iv) ", "• ", or with a tab after them), two or more in a row, become real list items: numbering style, punctuation and start from the numbers, levels from the indents; compound numbers ("1.2") become legal items. Works on the selected paragraphs, or the whole document when nothing is selected. Returns how many paragraphs changed.

### GetMarkdown

```autohotkey
Rte.GetMarkdown(embedImages = 1)
```

*embedImages*: 1/0 · returns: text

The document as Markdown. With embedImages = 1 (the default) images are data: URIs, so the text is complete on its own; with 0 they are written as image1.png, image2.png ... (as when viewing the source).

### InLegalList

```autohotkey
Rte.InLegalList()
```

returns: 1/0

True if the caret is in a legal numbered item.

### LegalBackspace

```autohotkey
Rte.LegalBackspace()
```

returns: 1/0

Backspace in a legal list: at the start of an item's text, its number goes (it becomes a normal paragraph); with a selection, the selection goes and the list is renumbered. Returns 0 when RichEdit's own Backspace is right.

### LegalDelete

```autohotkey
Rte.LegalDelete()
```

returns: 1/0

Delete in a legal list: at the end of an item, the next paragraph joins it (without its number); with a selection, the selection goes and the list is renumbered. Returns 0 when RichEdit's own Delete is right.

### LegalEnter

```autohotkey
Rte.LegalEnter()
```

returns: 1/0

Enter in a legal list: a new item, renumbered; in an empty item, the list ends there. Returns 0 if the caret is not in a legal item.

### LegalFormat

```autohotkey
value := Rte.LegalFormat
Rte.LegalFormat := value
```

text

How legal numbers look: "Top" (1. 1.1 1.1.1), "All" (1. 1.1. 1.1.1.) or "None" (1 1.1 1.1.1). Setting it renumbers the document's legal lists and keeps it in the settings.

### LegalNumbering

```autohotkey
Rte.LegalNumbering()
```

Makes the selected paragraphs a legal numbered list (1., 1.1, 1.1.1 by list level), or, if the caret is in one, ends it for the selected paragraphs. Paragraphs that start with a typed number such as "2.1 " take its level.

### Markdown

```autohotkey
value := Rte.Markdown
Rte.Markdown := value
```

text

The document as Markdown (get), or replaced by Markdown (set); see GetMarkdown and SetMarkdown.

### Renumber

```autohotkey
Rte.Renumber()
```

returns: 1/0

Puts every legal number right (the editor does this itself; for scripts that change the text directly). Returns 1 if anything changed.

### SetMarkdown

```autohotkey
Rte.SetMarkdown(markdown)
```

*markdown*: text

Replaces the document with Markdown. Relative image paths are resolved from the current file's folder.

### SideBySide

```autohotkey
value := Rte.SideBySide
Rte.SideBySide := value
```

1/0

Whether the Markdown or HTML view shows the formatted text beside it (setting it on in the rich text view shows Markdown).

### SourceText

```autohotkey
value := Rte.SourceText
Rte.SourceText := value
```

text

The Markdown or HTML in the source view ("" in the rich text view); setting it replaces that text and loads it.

### TouchesLegalList

```autohotkey
Rte.TouchesLegalList()
```

returns: 1/0

True if the selection includes a legal numbered item.

### View

```autohotkey
value := Rte.View
Rte.View := value
```

text

The document view: "Rich" (the formatted text), "Markdown" or "HTML" (the source view, see SideBySide).

## Text and selection

### ChangeCase

```autohotkey
Rte.ChangeCase(mode)
```

*mode*: text

Changes the case of the selection, keeping its formatting: "upper", "lower", "title", "sentence" or "toggle".

### CharacterCount

```autohotkey
Rte.CharacterCount()
```

returns: number

### Find

```autohotkey
Rte.Find(text, matchCase = 0, wholeWord = 0)
```

*text*: text · *matchCase*: 1/0 · *wholeWord*: 1/0 · returns: number

Finds the next occurrence after the selection (wrapping round), selects it and returns its position, or -1.

### GoTo

```autohotkey
Rte.GoTo(position)
```

*position*: number

Moves the caret; a negative position goes to the end of the document.

### InsertText

```autohotkey
Rte.InsertText(text)
```

*text*: text

Inserts text at the caret, replacing the selection; the caret ends up after the new text.

### Replace

```autohotkey
Rte.Replace(text, replacement, matchCase = 0, wholeWord = 0)
```

*text*: text · *replacement*: text · *matchCase*: 1/0 · *wholeWord*: 1/0 · returns: number

Replaces every occurrence and returns the number of replacements.

### Select

```autohotkey
Rte.Select(start, length = 0)
```

*start*: number · *length*: number

Selects length characters from start (0-based).

### WordCount

```autohotkey
Rte.WordCount()
```

returns: number

## Formatting

### ClearFontColor

```autohotkey
Rte.ClearFontColor()
```

### ClearHighlight

```autohotkey
Rte.ClearHighlight()
```

### SetBlockStyle

```autohotkey
Rte.SetBlockStyle(style)
```

*style*: BlockStyle

Sets the block style of the selected paragraphs: their style (Normal, Heading 1 to 6, Code), as SetStyle.

### ToggleList

```autohotkey
Rte.ToggleList(style)
```

*style*: ListStyle

Toggles a list; numbered lists start at 1 (RichEditorNET 1.0.11 starts new decimal lists at 0).

### ToggleSubscript

```autohotkey
Rte.ToggleSubscript()
```

### ToggleSuperscript

```autohotkey
Rte.ToggleSuperscript()
```

RichEditorNET 1.0.11: ToggleSuperscript sets Superscript and then clears Subscript, and RichEdit clears both, so turning it on has no effect. If that happens, the format is set through TOM (fixed upstream in the PR).

## Character formats

### InList

```autohotkey
Rte.InList()
```

returns: 1/0

True if the caret is in a list item.

### Nest

```autohotkey
Rte.Nest(levels = 1)
```

*levels*: number

Nests (positive) or un-nests (negative) the selected list items by that many levels, like Tab / Shift+Tab. With automatic level styles, numbered items take the style of their new level: 1. -&gt; a. -&gt; i. -&gt; 1. ...

### SetBackgroundColor

```autohotkey
Rte.SetBackgroundColor(color)
```

*color*: color

Sets the highlight (background) colour, as SetFontColor.

### SetFontColor

```autohotkey
Rte.SetFontColor(color)
```

*color*: color

Sets the font colour: of the selected text; with nothing selected, of what is typed next (and of an empty line).

### TabMakesSubLevel

```autohotkey
Rte.TabMakesSubLevel()
```

returns: 1/0

The caret at the start of an unnumbered paragraph of a numbered style (or in a number typed in front of it): Tab makes it a sub-level item there, as on a numbered paragraph, instead of typing a tab. So too a selection over several paragraphs, all of numbered styles (the a.-e. under a "1." item, their numbering off).

### ToggleBold

```autohotkey
Rte.ToggleBold()
```

Toggles bold: of the selected text; with nothing selected, of what is typed next (and of an empty line).

### ToggleItalic

```autohotkey
Rte.ToggleItalic()
```

Toggles italic, as ToggleBold.

### ToggleStrikeThrough

```autohotkey
Rte.ToggleStrikeThrough()
```

Toggles strikethrough, as ToggleBold.

### ToggleUnderline

```autohotkey
Rte.ToggleUnderline(style = UnderlineStyle.Single)
```

*style*: UnderlineStyle

Toggles an underline style (Single, Double, Dotted, Dashed, Wavy), as ToggleBold.

## Images

### InsertImageFile

```autohotkey
Rte.InsertImageFile(path, altText = "", width = 0, height = 0)
```

*path*: text · *altText*: text · *width*: number · *height*: number

Inserts a PNG, JPEG or GIF image at the caret. Without a width and height it keeps the image's own size, scaled down to at most 600 pixels wide; with just one of them, the other follows the image's proportions. (RichEditorNET shows inserted images at DefaultImageWidth x DefaultImageHeight, so those are set for the insert.)

### InsertLinkedThumbnailFile

```autohotkey
Rte.InsertLinkedThumbnailFile(imageUrl, path, altText = "")
```

*imageUrl*: text · *path*: text · *altText*: text

## Events

### EventNames

```autohotkey
value := Rte.EventNames
```

text · read-only

The event names, one per line.

### LastAsyncError

```autohotkey
value := Rte.LastAsyncError
```

text · read-only

The error of the last call made with Rte.Post that failed.

### Subscribe

```autohotkey
Rte.Subscribe(eventName, window)
```

*eventName*: text · *window*: long

Sends an event to an AHK window (A_ScriptHwnd) from now on.

### Unsubscribe

```autohotkey
Rte.Unsubscribe(eventName, window)
```

*eventName*: text · *window*: long

Stops sending an event ("*" for all events) to an AHK window.

## Recent files, settings, printing

### ExportPdf

```autohotkey
Rte.ExportPdf(path, bookmarks = 1)
```

*path*: text · *bookmarks*: 1/0 · returns: text

Exports the document as a PDF file through the "Microsoft Print to PDF" printer, and returns its full path. The page setup (File &gt; Page Setup) applies. This can take a few seconds. bookmarks (1, as at first): a bookmark for each heading, under the one of the level above it, and each table of contents' entry a link to its heading (iteration 53); 0: the pages only.

### Print

```autohotkey
Rte.Print(printerName = "")
```

*printerName*: text

Prints without a dialog, to the default printer or the printer named. The page setup applies.

### RecentFiles

```autohotkey
value := Rte.RecentFiles
```

text · read-only

The recently opened or saved files, newest first, one per line.

### SettingsFile

```autohotkey
value := Rte.SettingsFile
```

text · read-only

The settings file (window, view settings, recent files, page setup).

## List levels

### ApplyListLevels

```autohotkey
Rte.ApplyListLevels()
```

Applies List levels to the whole list the caret is in: each item gets the numbering and indents of its level (legal numbering's items, all of them: one numbering).

### AutomaticNumberingPerLevel

```autohotkey
value := Rte.AutomaticNumberingPerLevel
Rte.AutomaticNumberingPerLevel := value
```

1/0

Whether Tab and Shift+Tab give a list item its new level's numbering style (List levels' "automatic numbering style per level"; 0: items keep theirs). Remembered.

### ContinueNumbering

```autohotkey
Rte.ContinueNumbering()
```

Continues the numbering of the previous list of the same kind, for the list item with the caret and the rest of its list. Empty paragraphs between the two lists are removed, so that they become one list.

### ListLevels

```autohotkey
value := Rte.ListLevels
```

text · read-only

The document's list levels, one per line: level (1-9), style, punctuation, start number, and where the number and the text start (points from the margin, or "auto"), separated by tabs.

### NumberedList

```autohotkey
Rte.NumberedList()
```

Makes the selected paragraphs a numbered list (or ends it), numbered as List levels sets for their levels.

### RemoveList

```autohotkey
Rte.RemoveList()
```

Ends the lists of the selected paragraphs, whatever their numbering (bullets, numbers, legal numbers, or a Word document's numbered style): their numbers go, their text and indents stay. One undo step.

### RestartNumbering

```autohotkey
Rte.RestartNumbering(start = 1)
```

*start*: number

Starts the numbering again at start, from the list item with the caret (that item and the rest of its level). RichEdit numbers on through items that share a start value, so two lists that start at the same number are kept apart by an empty paragraph (at the top level; a nested list cannot be split that way).

### SetListLevel

```autohotkey
Rte.SetListLevel(level, style, punctuation = ".", start = 1, numberAt = -1, textAt = -1, forNewDocuments = 1)
```

*level*: number · *style*: text · *punctuation*: text · *start*: number · *numberAt*: number · *textAt*: number · *forNewDocuments*: 1/0

Sets a list level (1-9) of the document, and of new documents too unless forNewDocuments is 0: style Decimal, LowerAlpha, UpperAlpha, LowerRoman, UpperRoman or Bullet; punctuation ".", ")", "()" or ""; the start number; where the number and the text start (points from the margin; -1: automatic). New items and nested ones get the level; ApplyListLevels gives it to the list with the caret.

## Keys and macros

### Key

```autohotkey
Rte.Key(key, count = 1)
```

*key*: text · *count*: number

Presses a key, as recorded macros do: Left, Right, Up, Down, Home, End, Backspace, Delete, Enter or Tab, optionally with Ctrl+ and/or Shift+ ("Ctrl+Shift+Left"). Shift extends the selection; Ctrl moves by words, paragraphs or the whole document (Home/End). count repeats the key.

### RunMacro

```autohotkey
Rte.RunMacro(name)
```

*name*: text

Runs a macro from the rte_macros folder by name (without .ahk); it runs as a separate script.

## Event hooks

### EventHook

```autohotkey
Rte.EventHook(eventName)
```

*eventName*: text · returns: text

The macro hooked to an event ("" if none).

### EventHookList

```autohotkey
value := Rte.EventHookList
```

text · read-only

The hooks, one per line: "Event=macro" (" (off)" after one switched off).

### EventHookOn

```autohotkey
Rte.EventHookOn(eventName)
```

*eventName*: text · returns: 1/0

Whether an event's hook is on (true also when it has no macro).

### HookEventNames

```autohotkey
value := Rte.HookEventNames
```

text · read-only

The events a hook can run on, one per line.

### IsRecording

```autohotkey
value := Rte.IsRecording
```

1/0 · read-only

Whether a macro is being recorded.

### Macros

```autohotkey
value := Rte.Macros
```

text · read-only

The macros, one per line: file name and hotkey (if any), separated by a tab.

### PaletteItems

```autohotkey
value := Rte.PaletteItems
```

text · read-only

Iteration 57: what the command palette lists (Ctrl+Shift+P), one per line: "name&lt;tab&gt;where&lt;tab&gt;shortcut".

### SetEventHook

```autohotkey
Rte.SetEventHook(eventName, macro)
```

*eventName*: text · *macro*: text

Hooks a macro (a name in the macros folder, or a full path; "" for none) to an event: Started, Opened, NewDocument, BeforeSave, Saved, StyleChanged, Printed or Closing (Tools &gt; Event Hooks).

### SetEventHookOn

```autohotkey
Rte.SetEventHookOn(eventName, on)
```

*eventName*: text · *on*: 1/0

Switches an event's hook on or off (off: its macro stays hooked, and doesn't run).

### SetMacroHotkey

```autohotkey
Rte.SetMacroHotkey(name, keys)
```

*name*: text · *keys*: text

Sets the hotkey of a macro ("Ctrl+Alt+1", "F9", ...), or removes it with "". Works while the editor is active.

### StartRecording

```autohotkey
Rte.StartRecording()
```

Starts recording a macro, as Tools &gt; Record macro does.

### StopRecording

```autohotkey
Rte.StopRecording()
```

returns: text

Stops recording and returns the recorded AHK script ("" if nothing was recorded), without saving it.

## Spelling

### AddToDictionary

```autohotkey
Rte.AddToDictionary(word)
```

*word*: text

Adds a word to the Windows dictionary (shared with other programs), and refreshes the red underlines.

### NextSpellingError

```autohotkey
Rte.NextSpellingError()
```

returns: text

Selects the next spelling error after the caret (wrapping round) and returns the word, or "" if there is none.

### ReplaceRange

```autohotkey
Rte.ReplaceRange(start, length, text)
```

*start*: number · *length*: number · *text*: text

Replaces length characters at start with text, keeping the formatting of the first replaced character.

### SpellingErrors

```autohotkey
Rte.SpellingErrors()
```

returns: text

All spelling errors, one per line: start, length and word, separated by tabs.

### SpellingLanguage

```autohotkey
value := Rte.SpellingLanguage
```

text · read-only

The spell checker's language, e.g. en-US ("" if the Windows spell checker is not available).

### Suggestions

```autohotkey
Rte.Suggestions(word)
```

*word*: text · returns: text

Spelling suggestions for a word, one per line.

## Outline

### GoToHeading

```autohotkey
Rte.GoToHeading(number)
```

*number*: number

Goes to the n-th heading (1 = the first) and scrolls it to the top.

### Outline

```autohotkey
value := Rte.Outline
```

text · read-only

The headings, one per line: level (1-6), start position and text, separated by tabs.

### OutlineVisible

```autohotkey
value := Rte.OutlineVisible
Rte.OutlineVisible := value
```

1/0

Whether the outline panel is shown.

### SetHeadingLevel

```autohotkey
Rte.SetHeadingLevel(position, level)
```

*position*: number · *level*: number

Changes the heading level (1-6, or 0 for normal text) of the paragraph at a position.

## The document's header and footer

### DifferentFirstPage

```autohotkey
value := Rte.DifferentFirstPage
Rte.DifferentFirstPage := value
```

1/0

Whether the first page has a header and footer of its own (SetFirstHeader, SetFirstFooter). Empty parts leave the first page without one, as on a letter's or a pleading's first page.

### FirstFooterLine

```autohotkey
value := Rte.FirstFooterLine
Rte.FirstFooterLine := value
```

1/0

A line above the first page's footer.

### FirstFooterParts

```autohotkey
value := Rte.FirstFooterParts
```

text · read-only

The first page's footer parts, one per line.

### FirstHeaderLine

```autohotkey
value := Rte.FirstHeaderLine
Rte.FirstHeaderLine := value
```

1/0

A line below the first page's header (with DifferentFirstPage).

### FirstHeaderParts

```autohotkey
value := Rte.FirstHeaderParts
```

text · read-only

The first page's header parts, one per line (empty lines if it has none of its own).

### FooterFont

```autohotkey
value := Rte.FooterFont
```

text · read-only

The footer's font, as HeaderFont.

### FooterFromBottom

```autohotkey
value := Rte.FooterFromBottom
Rte.FooterFromBottom := value
```

number

From the page's bottom edge to the footer's last line, in inches (Word's "Footer from bottom"; 0.5 by default).

### FooterLine

```autohotkey
value := Rte.FooterLine
Rte.FooterLine := value
```

1/0

A line above the footer.

### FooterParts

```autohotkey
value := Rte.FooterParts
```

text · read-only

The footer's left, center and right parts, one per line.

### HeaderFont

```autohotkey
value := Rte.HeaderFont
```

text · read-only

The header's font ("Times New Roman 10 pt, bold"), empty without one of its own.

### HeaderFooterVisible

```autohotkey
value := Rte.HeaderFooterVisible
Rte.HeaderFooterVisible := value
```

1/0

Whether the header and footer are shown above and below the text (View &gt; Header and footer).

### HeaderFromTop

```autohotkey
value := Rte.HeaderFromTop
Rte.HeaderFromTop := value
```

number

Iteration 55: from the page's top edge to the header, in inches (Word's "Header from top"; 0.5 by default). A header reaching past the top margin pushes the text down, as in Word.

### HeaderLine

```autohotkey
value := Rte.HeaderLine
Rte.HeaderLine := value
```

1/0

Iteration 55: a line below the header (Word's border below its paragraphs).

### HeaderParts

```autohotkey
value := Rte.HeaderParts
```

text · read-only

The header's left, center and right parts, one per line.

### InsertHorizontalLine

```autohotkey
Rte.InsertHorizontalLine()
```

Iteration 55: a horizontal line across the text, in a paragraph of its own (Insert &gt; Horizontal Line): a picture of a line, as Word's, so that it is printed and kept in every format; Markdown gets a rule (---) and HTML an &lt;hr&gt;.

### MakeHeaderFooterDefault

```autohotkey
Rte.MakeHeaderFooterDefault()
```

Makes the document's header and footer the default for new documents (and documents without their own).

### SetFirstFooter

```autohotkey
Rte.SetFirstFooter(left = "", center = "", right = "")
```

*left*: text · *center*: text · *right*: text

Sets the first page's footer (left, center and right parts), as SetFirstHeader.

### SetFirstFooterFont

```autohotkey
Rte.SetFirstFooterFont(name = "", size = 10, bold = 0, italic = 0, underline = 0)
```

*name*: text · *size*: number · *bold*: 1/0 · *italic*: 1/0 · *underline*: 1/0

The first page's footer font, as SetHeaderFont.

### SetFirstHeader

```autohotkey
Rte.SetFirstHeader(left = "", center = "", right = "")
```

*left*: text · *center*: text · *right*: text

Sets the first page's header (left, center and right parts): the first page gets its own (DifferentFirstPage), shown (ShowFirstHeader). (Iteration 55) The other pages are as they were: ShowHeader off, only the first page has one.

### SetFirstHeaderFont

```autohotkey
Rte.SetFirstHeaderFont(name = "", size = 10, bold = 0, italic = 0, underline = 0)
```

*name*: text · *size*: number · *bold*: 1/0 · *italic*: 1/0 · *underline*: 1/0

The first page's header font, as SetHeaderFont (with DifferentFirstPage; an empty name: plain).

### SetFooter

```autohotkey
Rte.SetFooter(left = "", center = "", right = "")
```

*left*: text · *center*: text · *right*: text

Sets the document's footer (left, center and right parts) and shows it.

### SetFooterFont

```autohotkey
Rte.SetFooterFont(name = "", size = 10, bold = 0, italic = 0, underline = 0)
```

*name*: text · *size*: number · *bold*: 1/0 · *italic*: 1/0 · *underline*: 1/0

The footer's font, as SetHeaderFont.

### SetHeader

```autohotkey
Rte.SetHeader(left = "", center = "", right = "")
```

*left*: text · *center*: text · *right*: text

Sets the document's header (left, center and right parts) and shows it. A part can have several lines ("`n" between them).

### SetHeaderFont

```autohotkey
Rte.SetHeaderFont(name = "", size = 10, bold = 0, italic = 0, underline = 0)
```

*name*: text · *size*: number · *bold*: 1/0 · *italic*: 1/0 · *underline*: 1/0

Iteration 55: the header's font: its name, size in points, bold, italic and underline (all of its parts, the first page's too); printed, on Show Pages' sheets, and in Word and OpenDocument files. An empty name: Segoe UI 8 pt in grey, as before fonts could be chosen.

### ShowFirstFooter

```autohotkey
value := Rte.ShowFirstFooter
Rte.ShowFirstFooter := value
```

1/0

Whether the first page has a footer (with DifferentFirstPage; the other pages: ShowFooter).

### ShowFirstHeader

```autohotkey
value := Rte.ShowFirstHeader
Rte.ShowFirstHeader := value
```

1/0

Iteration 55: whether the first page has a header (with DifferentFirstPage; the other pages: ShowHeader).

### ShowFooter

```autohotkey
value := Rte.ShowFooter
Rte.ShowFooter := value
```

1/0

Whether the document has a footer.

### ShowHeader

```autohotkey
value := Rte.ShowHeader
Rte.ShowHeader := value
```

1/0

Whether the document has a header (printed, and in its files).

## Cross-references to numbered paragraphs

### CrossReferenceTargets

```autohotkey
value := Rte.CrossReferenceTargets
```

text · read-only

The numbered paragraphs a cross-reference can refer to (legal items, numbered styles' paragraphs, numbered lists), one a line: the number, a tab, the full number (with the levels above: "3(b)"), a tab, and the start of the text.

### InsertCrossReference

```autohotkey
Rte.InsertCrossReference(index, show = "number")
```

*index*: number · *show*: text

Inserts a cross-reference at the caret (in place of the selection) to numbered paragraph index (1: the first, as CrossReferenceTargets lists them), showing its number (show "number": "(b)") or its full number ("full": "3(b)"), without a trailing period. Its number follows the paragraph's when the paragraphs are renumbered; Word and LibreOffice files get it as a REF field. One undo step.

### UpdateCrossReferences

```autohotkey
Rte.UpdateCrossReferences()
```

returns: number

Brings every cross-reference's number up to date (done by itself as the paragraphs are renumbered, and before saving, printing and exporting; F9). Returns how many changed.

## Tables of contents

### InsertTableOfContents

```autohotkey
Rte.InsertTableOfContents(levels = 3, title = "Contents", pageNumbers = 1, rightAlign = 1, leader = "dots")
```

*levels*: number · *title*: text · *pageNumbers*: 1/0 · *rightAlign*: 1/0 · *leader*: text

Inserts a table of contents at the caret's paragraph (in place of it if it is empty, else before it): the headings of levels 1 to levels (1-6), each with the page it is printed on, under a title ("": none). pageNumbers 0: no page numbers; rightAlign 0: each just after its heading's text, else at the right margin after a leader (dots, dashes, underline or none). With the caret in a table of contents: that table made again with these options. One undo step.

### TableOfContents

```autohotkey
value := Rte.TableOfContents
```

text · read-only

The (first) table of contents' entries, one a line: the level, a tab, the heading's text, a tab, the page.

### UpdateTableOfContents

```autohotkey
Rte.UpdateTableOfContents(pagesOnly = 0)
```

*pagesOnly*: 1/0 · returns: number

Makes every table of contents again from the headings, with their pages as printed (pagesOnly = 1: only their page numbers brought up to date, as before saving and printing). Returns how many tables were made again (pagesOnly: how many page numbers changed). One undo step.

## Footnotes

### DeleteFootnote

```autohotkey
Rte.DeleteFootnote(number)
```

*number*: number

Takes footnote number (1: the first) out, its number and its note; the others are numbered again. One undo step.

### FootnoteCount

```autohotkey
value := Rte.FootnoteCount
```

number · read-only

The number of footnotes.

### FootnoteList

```autohotkey
value := Rte.FootnoteList
```

text · read-only

The document's footnotes, one a line: the number, a tab and the note (its formatting as &lt;b&gt; &lt;i&gt; &lt;u&gt; tags; its paragraphs apart by "`r", so that each note stays on its line). FootnoteText gives one note.

### FootnoteSize

```autohotkey
value := Rte.FootnoteSize
Rte.FootnoteSize := value
```

number

The size of the footnotes' notes (points): the document's Footnote Text style's (set: made if it has none, as a Word document opened has it), else three quarters of the text's (12 points: 9).

### FootnoteText

```autohotkey
Rte.FootnoteText(number)
```

*number*: number · returns: text

Footnote number's (1: the first) note: its formatting as &lt;b&gt; &lt;i&gt; &lt;u&gt; tags, its paragraphs apart by "`n".

### InsertFootnote

```autohotkey
Rte.InsertFootnote(note)
```

*note*: text

Inserts a footnote at the caret (after the selection) with its note (its paragraphs apart by "`n"; &lt;b&gt; &lt;i&gt; &lt;u&gt; tags for bold, italic and underline): a superscript number in the text, the note at the foot of the page when printed (and on Show Pages' sheets). The footnotes are numbered 1, 2, 3 ... in order. One undo step.

### SetFootnote

```autohotkey
Rte.SetFootnote(number, note)
```

*number*: number · *note*: text

Gives footnote number (1: the first) another note (as InsertFootnote's: paragraphs apart by "`n", &lt;b&gt; &lt;i&gt; &lt;u&gt; tags for its formatting). One undo step.

## Window

### Activate

```autohotkey
Rte.Activate()
```

Brings the editor window to the front.

### Close

```autohotkey
Rte.Close(discardChanges = 0)
```

*discardChanges*: 1/0

Closes the editor after replying. With discardChanges = 0, the user is asked about unsaved changes.

### Help

```autohotkey
Rte.Help()
```

returns: text

Lists everything AHK scripts can call.

### SetStatus

```autohotkey
Rte.SetStatus(text)
```

*text*: text

Shows a message in the status bar.

## Formatting helpers

### SetFontName

```autohotkey
Rte.SetFontName(name)
```

*name*: text

Sets the font of the selected text; with nothing selected, of what is typed next, and on an empty paragraph of its paragraph mark too (the empty line keeps the font), as in Word.
