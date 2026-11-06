# Formatting and paragraphs

## Characters

**Home > Font** has the font and size, bold, italic, underline, colour and highlight; **Home > Text** has change case, the underline styles and the Font dialog. Any font size can be used.

With nothing selected, a font, size or colour you choose applies to what you type next, and at once to an empty line, as in Word.

**Clear Formatting** (Ctrl+Space) puts the selected text back to its paragraph's style; with nothing selected, the whole paragraph. See [Styles](styles.md).

## Format Painter

**Home > Clipboard > Format Painter** (the paint brush) copies formatting from one place to another, as Word's does:

1. Put the caret in the text whose look you want, or select some of it.
2. Click **Format Painter**. The pointer becomes a brush.
3. Click in another paragraph, or select other text: it takes the same formatting.

With the caret in a paragraph (nothing selected), it copies the paragraph's formatting (its style, alignment, spacing, indents and numbering, legal numbering included) and the character formatting at the caret; a click then paints the paragraph clicked and the word under the click. With text selected, it copies the formatting most of the selection has, and the paragraph's formatting only when the selection takes in the paragraph's mark (the end of the paragraph), as in Word.

One click on Format Painter paints once. **Double-click** it to keep painting, one place after another, until you press ++esc++ or click Format Painter again. Each paint is one Undo step.

!!! tip "From a script"
    `Rte.CopyFormatting()` and `Rte.PasteFormatting()`: see the [API reference](../automation/api.md).

## Alignment, line spacing and spacing

**Home > Paragraph** has the alignment buttons and **Line Spacing**: 1.0 to 3.0 lines, space before and after the paragraph, and *Line Spacing Options* for "at least" or "exactly" so many points.

## The Paragraph dialog

The **Paragraph settings** button in Home > Paragraph opens a Paragraph dialog like Word's:

- **Indents and Spacing:** alignment; left and right indents; a first-line or hanging indent; space before and after; line spacing.
- **Line and Page Breaks:** widow and orphan control, keep with next, keep lines together, page break before, and suppress line numbers.

A preview shows the result. When the selected paragraphs differ in a setting, it is left empty, and each paragraph keeps its own unless you change it. Indents are in inches and spacing in points, but you can type any unit: `1.27 cm`, `36pt`, `0.5in`.

- **Tabs...** sets the paragraphs' tab stops (position, alignment, leader) and the document's default tab stops.
- **Set As Default...** gives the paragraph's style these settings, for this document only or for new documents too.

## The ruler

**View > Ruler** (off at first) shows a ruler above the page, as Word's horizontal ruler does. It measures in inches, or in centimetres where Windows measures in them, from the text's left edge. In Print Layout the margins are grey and the text's width white.

On it are the indents and tab stops of the paragraph at the caret (with several paragraphs selected, the first one's):

- the **first line's indent**: the triangle at the top, on the left;
- the **hanging indent**: the triangle at the bottom, with the **left indent**'s box under it;
- the **right indent**: the triangle at the right;
- the paragraph's **tab stops**, and after them the document's default tab stops as small grey ticks.

In a table the ruler shows only the scale, and over Multiple Pages it is hidden. It follows the zoom and the page when it is scrolled sideways.

What you can do on it (each change applies to every selected paragraph and is one Undo step):

| Do | To |
|---|---|
| click the white part | add a left tab stop there |
| drag a tab stop | move it (it keeps its alignment and leader) |
| drag a tab stop down or up off the ruler | clear it |
| press on the white part and drag | add a tab stop where you let go |
| drag the top triangle | set the first line's indent |
| drag the bottom triangle | set the hanging indent (the first line stays put) |
| drag the box under it | move the left indent and the first line together |
| drag the right triangle | set the right indent |
| double-click a tab stop (or the white part) | open **Tabs** |
| double-click an indent marker | open **Paragraph** |
| double-click the grey margins | open **Page Setup** |

Clicks and drags snap to a grid of 1/16 inch (0.25 cm), as in Word; hold ++alt++ while dragging to place a marker freely. While you drag, the ruler shows the marker's new place, a dotted line runs down the page under it, and the status bar shows the value. Indents stay on the paper and leave at least a quarter inch of text.

The Paragraph dialog and its Tabs dialog do the same by numbers.

Printing, Print Preview, PDF and Multiple Pages follow the line and page break rules, as Word does, and Word and LibreOffice get them in their files.

!!! tip "From a script"
    `Rte.ParagraphFormat()`, `Rte.SetParagraphFormat("LeftIndent=0.5in; KeepTogether=1")`, `Rte.Tabs`, `Rte.SetTabs(...)`, `Rte.AddTab`, `Rte.MoveTab`, `Rte.ClearTab`, `Rte.DefaultTabStop` and `Rte.RulerVisible`: see the [API reference](../automation/api.md).
