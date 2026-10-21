# The experimental officeimo converter

This branch adds an "officeimo" converter for Word files (`rte_office7.cs`, version 8). It uses
[OfficeIMO.Word.Rtf](https://github.com/EvotecIT/OfficeIMO) (MIT) to turn the document's RTF into .docx, and a
.docx into RTF, directly: no Word, no LibreOffice, and no detour through HTML. It runs only when named, as the `via`
of `Rte.Open` and `Rte.SaveAs` or in `AHKWORD_CONVERTERS`. `Rte.ConverterReport()` gives what OfficeIMO says it
couldn't carry over.

It is kept off main because, with OfficeIMO 3.5.1 (October 2026), it loses things the built-in openxml converter keeps:

| | openxml | officeimo 3.5.1 |
|---|---|---|
| Heading style, bold, italic, colour, double spacing | kept | kept |
| Table, picture, page break, header, footer, line numbers | kept | kept |
| Centred and right alignment | kept | lost writing and reading |
| An ordinary numbered list (RichEdit's `\pn` numbering) | kept | lost |
| Numbered styles and legal numbering | kept | kept writing, lost reading |
| Style names, reading | kept | without spaces ("ListParagraph") |
| Its loss report | | "nothing lost" in every case |

`where-lost.cs` shows the alignment is lost in OfficeIMO's conversion to Word, not in reading RTF: even
`\pard\qc Centred.\par` is parsed as centred and written to the .docx without its alignment.

## Trying again

Change the four `#:package OfficeIMO.*` versions at the top of `rte_office7.cs`, then run the two tests as AhkWord
self-tests (they are `SelfTest` methods in the form the test runners compile into `WriterForm`):

- `compare-converters.cs`: one made-up document saved and reopened by each converter, and each converter's file read
  by the other, compared with the original.
- `where-lost.cs`: where a paragraph's alignment and RichEdit's list numbering go missing, step by step.

When officeimo matches openxml on these, it can join the default converter order on main.
