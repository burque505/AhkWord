# Documents and files

## Opening and saving

**File > Open** opens, and **Save** and **Save As** save, documents in these formats:

| Format | Notes |
|---|---|
| **Rich Text Format** (`.rtf`) | AhkWord's own format: keeps everything, styles and numbering included. Word and LibreOffice open it too. |
| **Word** (`.docx`, `.doc`) and **OpenDocument** (`.odt`) | Through a converter: Word, LibreOffice, the built-in one or pandoc (see [File formats and converters](../reference/formats.md)). |
| **HTML** (`.html`) | Readable HTML, with styles, tables and spacing kept. |
| **Markdown** (`.md`) | GitHub's flavour, with tables and pictures. |
| **Plain text** (`.txt`) | Text only. |

You can also drop a file on the window to open it. **File > Export** saves a copy as **PDF** (through *Microsoft Print to PDF*), Word or OpenDocument, leaving the document itself as it is; **Layout > Export as PDF** does the same for PDF.

A document opened from Word keeps its look: its own styles, spacing, indents and tables. See [Styles](styles.md) for giving it yours.

## New documents and templates

**File > New** starts an empty document with your styles. **File > New from Template** starts one from a template (Letter, Memo, Report, Court Paper, Meeting Notes, or your own), asking for its fields, such as the parties' names or a hearing date. See [Templates](templates.md).

**File > Start Page** shows the start page: a blank document, the templates, your recent documents and the macros. It can open by itself when AhkWord starts.

## Several windows

**File > New Window** and **File > Open in New Window** open another AhkWord window, to work on two documents at once. Each window is a complete editor; they share your settings without undoing each other's changes. On the taskbar they are grouped as AhkWord's.

## Autosave and recovery

While a document has unsaved changes, AhkWord writes a recovery copy every minute (never over the document's own file). If it ends without closing (a crash, a power cut, Windows restarting), the next start offers to recover the document. Save replaces the original file; Save As keeps both.

## Recent documents

The **File** menu lists your recent documents, and AhkWord remembers the window's size and place and your page settings.
