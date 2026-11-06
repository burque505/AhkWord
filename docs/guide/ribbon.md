# The ribbon and keyboard

AhkWord's commands are on a ribbon, the one built into Windows (as in Paint and WordPad): it follows the screen's scaling, and works with the keyboard and with screen readers.

| Tab | Groups |
|---|---|
| **File** (menu) | New, New from Template, Start Page, Open, Open in New Window, New Window, Save, Save As, Close (Ctrl+W), Print (with Print Preview and Page Setup), Export (PDF, Word, OpenDocument), Exit, and the recent documents |
| **Home** | Clipboard; Font; Text (change case, underline style, the Font dialog); Paragraph (lists, levels, No numbering, alignment, Line Spacing, List Options, the Paragraph dialog); Styles; Editing (find, replace, select all) |
| **Insert** | Page Break, Table of Contents, Table, Picture, Linked Thumbnail, Horizontal Line, Link, Cross-reference, Footnote, Update (F9), Date & Time, Insert from Template, Fill In Fields, Snippets |
| **Layout** | Page Setup, Header & Footer, Line Numbers, Print Preview, Export as PDF |
| **Review** | Next Spelling Error, Add to Dictionary, Check spelling as you type; Language (English or Español, from the next start) |
| **View** | Document View (Rich Text, Markdown, HTML, Side by Side, Wrap source lines, Source line numbers), Print Layout, Show Pages, Wrap to window, Header and footer, Outline, Ruler, Zoom (Page Width, One Page, Multiple Pages, Zoom In, Zoom Out, 100%) |
| **Tools** | Macros, Record Macro, Macro Hotkeys, Event Hooks, Run Script, Macros Folder, Command Palette, Automation API, About |
| **Table Tools** | while the caret is in a table: rows and columns, Properties, Borders, Shading, delete |

The **quick access toolbar** has Save, Undo, Redo and the Command Palette; its arrow adds New, Open, Print, Print Preview or Export as PDF.

The **right-click menu** has, as in Word, the spelling suggestions on a misspelled word (when nothing else is selected), then Cut, Copy and Paste, **Edit footnote** on a footnote's number, **Update table** and **Update page numbers only** on a table of contents, Font... and Paragraph..., and in a table the table commands. Right-clicking outside the selection moves the caret there first; inside it, the selection stays.

## The command palette

Press ++ctrl+shift+p++ (or Tools > Command Palette, or its magnifier on the quick access toolbar), type part of a command's name, and press ++enter++. It finds any ribbon command, shown with the tab and group it is on and its shortcut, so it is also a way to learn where things are; and your macros, the styles, your snippets and the recent documents. ++up++ and ++down++ choose, ++esc++ closes it.

- Every word you type has to match: `insert table`, `page num`, `heading 2`. A command's name counts first, then its place on the ribbon, then its tooltip, so `margins` finds Page Setup.
- A command the ribbon has greyed out (Undo with nothing to undo, the table commands outside a table) is greyed out here too, marked *(not now)*.
- Toggles say whether they are on: *Print Layout (on)*.
- With nothing typed, the ones you used last come first.
- A command runs exactly as its button does: in the Markdown or HTML view it does what the button does there, and a macro being recorded records it.

++ctrl+k++ stays Insert Link, as in Word.

## Keyboard

Press ++alt++ to see the KeyTips: ++h++ Home, ++n++ Insert, ++p++ Layout, ++r++ Review, ++w++ View, ++l++ Tools, ++f++ File. Each button's tooltip shows its shortcut. The main ones:

| Keys | Does |
|---|---|
| ++ctrl+b++, ++ctrl+i++, ++ctrl+u++ | bold, italic, underline |
| ++ctrl+0++ | Normal style |
| ++ctrl+1++ to ++ctrl+6++ | Heading 1 to 6 |
| ++ctrl+7++ | Code style |
| ++ctrl+space++ | Clear Formatting |
| ++ctrl+shift+n++ | numbered list |
| ++ctrl+shift+g++ | legal numbering |
| ++tab++, ++shift+tab++ | list items and numbered paragraphs a level in or out; next and previous table cell |
| ++ctrl+enter++ | page break |
| ++shift+enter++ | line break |
| ++ctrl+k++ | link |
| ++ctrl+shift+p++ | the command palette |
| ++ctrl+w++ | close the document |
| ++f7++ | next spelling error |
| ++ctrl+alt+f++ | footnote |
| ++f9++ | update the cross-references and footnote numbers |
| ++ctrl++ + click | follow a link or a cross-reference |
| ++ctrl+shift+r++ | record a macro |

The outline panel (View > Outline) lists the headings: click one to go there; right-click it to promote or demote it.

The ruler (View > Ruler) shows, above the page, the margins in grey and the text's width in white, measured in inches (or centimetres where Windows measures in them) from the text's left edge. On it are the indents of the paragraph at the caret (the first line's, the triangle at the top; the hanging indent, the triangle at the bottom, with the left indent's box under it; the right indent's, at the right), its tab stops, and the default tab stops after them as small grey ticks. With several paragraphs selected it shows the first one's; in a table, only the scale. A click on its white part adds a left tab stop there to the selected paragraphs (on a grid of 1/16 inch, or 0.25 cm; not on a tab stop or a marker, nor in a table); Undo takes it away. A double-click opens **Tabs** on a tab stop, **Paragraph** on an indent marker, and **Page Setup** on the grey margins. The markers can be dragged, as in Word: a tab stop to move it (drag it down or up off the ruler to clear it); the top triangle for the first line's indent; the bottom triangle for the hanging indent (the first line stays put); the box under it moves the left indent and the first line together; the right triangle for the right indent. A press on the white part dragged puts the new tab stop where it's let go. Dragged markers snap to the grid (hold Alt to place them freely); a dotted line runs down the page under them and the status bar shows where they'd go. Each drag is one Undo step, for the selected paragraphs. **Paragraph...** and its **Tabs...** do the same by numbers (see [Formatting](formatting.md)).
