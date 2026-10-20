# File formats and converters

## What each format keeps

| | RTF | Word (.docx) | OpenDocument (.odt) | HTML | Markdown |
|---|---|---|---|---|---|
| Styles, with their definitions | yes | yes | yes | yes | headings and lists |
| Numbered styles and list levels | yes | yes, as Word's lists | yes | numbers | numbers |
| Line spacing, space before and after | yes | yes | yes | yes | no |
| Indents and tab stops | yes | yes | yes | yes | no |
| Tables: widths, borders, shading | yes | yes | yes | yes | no |
| Headers and footers | yes | yes | yes | yes | yes |
| Line numbers | yes | yes | yes | yes | yes |
| Page breaks | yes | yes | yes | yes | yes |
| Pictures with alt text | yes | yes | yes | yes | yes |

RTF is AhkWord's own format and keeps everything. Markdown is plain text by design, so it has no table formatting or paragraph spacing.

## Word and OpenDocument converters

Word and OpenDocument files are converted by the first of these that is available:

1. **Microsoft Word**, run invisibly, and only when Word isn't open (AhkWord never touches the Word you are working in).
2. **LibreOffice**, run in the background with a profile of its own.
3. **The built-in converter** (`.docx` only, nothing to install): keeps tables' widths, borders, shading and alignment, and paragraphs' indents, tab stops and spacing.
4. **pandoc**.

`AHKWORD_CONVERTERS` chooses which of them are used, in which order (see [Settings and folders](settings.md#environment-variables)), and scripts can name the converter for one file: `Rte.Open(path, "libreoffice")`, `Rte.ExportDocx(path, "openxml")`.

Whichever converter opens a `.docx`, its paragraphs get the line spacing and space before and after that Word gives them, worked out from the file's defaults, its styles and each paragraph's own format.

## Lists and numbering through Word

Ordinary lists, legal numbering and numbered styles keep their numbers, levels and indents through Word files: AhkWord writes them as Word's own list definitions, and reads a Word document's lists, list levels and numbered styles back. The list levels' indents come from where Word draws the list paragraphs.
