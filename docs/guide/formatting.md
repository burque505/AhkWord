# Formatting and paragraphs

## Characters

**Home > Font** has the font and size, bold, italic, underline, colour and highlight; **Home > Text** has change case, the underline styles and the Font dialog. Any font size can be used.

With nothing selected, a font, size or colour you choose applies to what you type next, and at once to an empty line, as in Word.

**Clear Formatting** (Ctrl+Space) puts the selected text back to its paragraph's style; with nothing selected, the whole paragraph. See [Styles](styles.md).

## Alignment, line spacing and spacing

**Home > Paragraph** has the alignment buttons and **Line Spacing**: 1.0 to 3.0 lines, space before and after the paragraph, and *Line Spacing Options* for "at least" or "exactly" so many points.

## The Paragraph dialog

The **Paragraph settings** button in Home > Paragraph opens a Paragraph dialog like Word's:

- **Indents and Spacing:** alignment; left and right indents; a first-line or hanging indent; space before and after; line spacing.
- **Line and Page Breaks:** widow and orphan control, keep with next, keep lines together, page break before, and suppress line numbers.

A preview shows the result. When the selected paragraphs differ in a setting, it is left empty, and each paragraph keeps its own unless you change it. Indents are in inches and spacing in points, but you can type any unit: `1.27 cm`, `36pt`, `0.5in`.

- **Tabs...** sets the paragraphs' tab stops (position, alignment, leader) and the document's default tab stops.
- **Set As Default...** gives the paragraph's style these settings, for this document only or for new documents too.

Printing, Print Preview, PDF and Multiple Pages follow the line and page break rules, as Word does, and Word and LibreOffice get them in their files.

!!! tip "From a script"
    `Rte.ParagraphFormat()`, `Rte.SetParagraphFormat("LeftIndent=0.5in; KeepTogether=1")`, `Rte.Tabs`, `Rte.SetTabs(...)` and `Rte.DefaultTabStop`: see the [API reference](../automation/api.md).
