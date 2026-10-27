# The ribbon and keyboard

AhkWord's commands are on a ribbon, the one built into Windows (as in Paint and WordPad): it follows the screen's scaling, and works with the keyboard and with screen readers.

| Tab | Groups |
|---|---|
| **File** (menu) | New, New from Template, Start Page, Open, Open in New Window, New Window, Save, Save As, Close (Ctrl+W), Print (with Print Preview and Page Setup), Export (PDF, Word, OpenDocument), Exit, and the recent documents |
| **Home** | Clipboard; Font; Text (change case, underline style, the Font dialog); Paragraph (lists, levels, No numbering, alignment, Line Spacing, List Options, the Paragraph dialog); Styles; Editing (find, replace, select all) |
| **Insert** | Page Break, Table of Contents, Table, Picture, Linked Thumbnail, Horizontal Line, Link, Cross-reference, Footnote, Update (F9), Date & Time, Insert from Template, Fill In Fields, Snippets |
| **Layout** | Page Setup, Header & Footer, Line Numbers, Print Preview, Export as PDF |
| **Review** | Next Spelling Error, Add to Dictionary, Check spelling as you type |
| **View** | Document View (Rich Text, Markdown, HTML, Side by Side, Wrap source lines, Source line numbers), Print Layout, Show Pages, Wrap to window, Header and footer, Outline, Zoom (Page Width, One Page, Multiple Pages, Zoom In, Zoom Out, 100%) |
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
