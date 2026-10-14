# AhkWord

A word processor for Windows, written in C# and run inside AutoHotkey v2, that AutoHotkey scripts can drive while you type.

AhkWord is built on [RichEditorNET](https://github.com/ellipticbit/RichEditorNET), a .NET wrapper around the Windows RichEdit control, with the Windows Ribbon for its commands. [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compiles its C# source when it starts and runs it in the AutoHotkey process, so there is no build step and no separate executable.

![AhkWord with the showcase document: the ribbon with Table Tools, the outline, headings, lists, legal numbering, a table, a picture, and the header and footer](images/showcase.png)

*[`examples/Showcase.rtf`](examples/Showcase.rtf), made entirely with the editor's own commands.*

> **Work in progress.** The editor is still named **RTE Writer** in the code, its window and its settings folder, and its files are numbered by iteration (`rte_writer16.*`). Renaming them to AhkWord is planned.

## Features

- **Documents:** HTML, Markdown (GitHub flavor, with tables and images), plain text and RTF. Word (`.docx`, `.doc`) and OpenDocument (`.odt`) files are opened and saved through a converter (see [Optional converters](#optional-converters)). Export as PDF through "Microsoft Print to PDF".
- **Formatting:** fonts, colors, highlight, alignment, change case and underline styles. Structured documents use block styles (normal text, headings, preformatted); free-form documents take any font size.
- **Lists:** numbered and bulleted lists nested with Tab and Shift+Tab, a numbering style per level, restart and continue numbering, and legal numbering (1., 1.1, 1.1.1). Typed numbers in documents from elsewhere can be turned into real lists.
- **Tables:** insert tables, add and delete rows and columns, move between cells with Tab.
- **Pages:** manual page breaks (Ctrl+Enter), headers and footers with page-number and date fields, Page Setup, Print Preview and printing. Headers and footers are shown in the editor and saved in the document.
- **Pictures:** insert from a file or by drag and drop, resize, and give them alt text (shown when the mouse rests on a picture).
- **Writing tools:** spelling suggestions from the Windows spell checker (F7 goes to the next error) and an outline panel of the headings.
- **Comfort:** recent documents and remembered window and page settings. Drop a file on the window to open it.

## The ribbon

The commands are on a ribbon, the one built into Windows (as in Paint and WordPad): it follows the screen's scaling, and works with the keyboard and screen readers.

| Tab | Groups |
|---|---|
| **File** (menu) | New, Open, Save, Save As, Print (with Print Preview and Page Setup), Export (PDF, Word, OpenDocument), Exit, and the recent documents |
| **Home** | Clipboard, Font, Text (change case, underline style, font dialog), Paragraph (lists, levels, alignment, List Options), Styles, Editing (find, replace, select all) |
| **Insert** | Page Break, Table, Picture, Linked Thumbnail, Link, Date & Time |
| **Layout** | Page Setup, Header & Footer, Print Preview |
| **Review** | Next Spelling Error, Add to Dictionary, Check spelling as you type |
| **View** | Structured or Free-form, Word wrap, Header and footer, Outline, Zoom, HTML and Markdown source |
| **Tools** | Macros, Record Macro, Macro Hotkeys, Run Script, Macros Folder, Automation API, About |
| **Table Tools** | Appears while the caret is in a table: insert rows and columns, delete rows, columns or the table |

- **Keyboard:** press **Alt** to see the KeyTips (H Home, N Insert, P Layout, R Review, W View, L Tools, F File). The usual shortcuts work too: Ctrl+B, Ctrl+Shift+N (numbered list), Ctrl+Shift+G (legal numbering), Ctrl+0 to Ctrl+7 (block styles), Ctrl+Enter (page break), F7 (spelling) and more; each button's tooltip shows its shortcut.
- **Numbering:** the arrow next to the Numbering button chooses the numbering style and opens List Levels, where each level's style, punctuation and start number are set. List Options has the rest.
- **Quick access toolbar:** Save, Undo and Redo; its arrow adds New, Open, Print or Print Preview.

## Automation from AutoHotkey

Any AutoHotkey v2 script can control a running editor through `RteClient2.ahk`:

```ahk
#Requires AutoHotkey v2.0
#Include "RteClient2.ahk"

Rte.InsertText("Hello from AutoHotkey")
Rte.InsertTable(3, 4)
MsgBox Rte.WordCount() " words"
Rte.On("Saved", (e) => TrayTip(Rte.FileName, "Saved"))
```

- **Calls** wait for the editor and return its result. `Rte.Post(...)` queues a call and returns at once. **Tools > Automation API** in the editor lists everything that can be called.
- **Events** (`TextChanged`, `SelectionChanged`, `ModifiedChanged`, `Saved`, `Opened`, `Closed`, `AsyncError`) are delivered to callbacks registered with `Rte.On`.
- **Macros:** the **Macros** button on the Tools tab runs the scripts in [`rte_macros`](rte_macros), and **Macro Hotkeys** gives any of them a hotkey. **Record Macro** (Ctrl+Shift+R) records typing and ribbon commands as a new macro script.
- **Hotkeys in the host script:** `rte_writer16.ahk` itself defines a couple of hotkeys (Ctrl+Alt+D inserts the date, Ctrl+Alt+W shows word counts). Add your own there.

The example macros include an auto-save, a word goal tracker, an event monitor, document statistics and Title Case for the selection.

## Requirements

| Component | Notes |
|---|---|
| Windows | 10 or 11 |
| AutoHotkey v2 | **64-bit** (`AutoHotkey64.exe`) |
| .NET runtime | .NET 10, including **Microsoft.WindowsDesktop.App** |
| Roslyn compiler | The `Roslyn` folder containing `csc.exe`, e.g. from Visual Studio or Build Tools |

Csharp-in-AHK finds .NET and Roslyn by itself. If it can't, see [its README](https://github.com/burque505/Csharp-in-AHK#readme) for how to set the paths.

NuGet packages (RichEditorNET, WindowsRibbon, the Open XML libraries, SkiaSharp and others) are downloaded on the first start, which therefore takes longer.

### Optional converters

Word and OpenDocument files are converted by the first of these that is available:

1. **Microsoft Word**, run invisibly, and only when Word isn't already open.
2. **LibreOffice**, run headless with its own profile.
3. **The built-in Open XML converter** (`.docx` only, no extra software needed).
4. **pandoc**.

HTML, Markdown, RTF and plain text need none of them.

### Rebuilding the ribbon

The ribbon comes compiled (`ribbon/RteRibbon.ribbon`), so running AhkWord needs nothing more. Only after changing `ribbon/RteRibbon.xml` or the icons, run `ribbon/build_ribbon.ahk`. It needs:

- the **Windows SDK** (for `uicc.exe` and `rc.exe`), and
- the **C++ desktop tools** of Visual Studio or Build Tools (for `link.exe`).

It downloads the [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) font once, draws the icons from it, and compiles the markup into `RteRibbon.ribbon`.

## Getting started

1. Install the requirements above.
2. Clone this repository.
3. Run `rte_writer16.ahk`. To open a document at start, pass its path: `rte_writer16.ahk "C:\path\document.md"`. Try `examples\Showcase.rtf`.

Settings are kept in `%APPDATA%\RTE Writer\settings.json`. Set the environment variable `RTE_WRITER_SETTINGS` to a file path to keep them elsewhere.

## Files

| File | Purpose |
|---|---|
| `rte_writer16.ahk` | Starts the editor and holds its in-process hotkeys |
| `rte_writer16.cs` | The editor window, the ribbon's commands and the automation interface |
| `rte_markdown5.cs` | Markdown reading and writing |
| `rte_htmllists2.cs` | HTML list numbering (start numbers, restarts, styles) |
| `rte_tables.cs` | Tables |
| `rte_legal.cs` | Legal numbering |
| `rte_pagebreaks.cs` | Page breaks |
| `rte_headerfooter.cs` | Headers and footers |
| `rte_office6.cs` | Word, OpenDocument and RTF conversion |
| `ribbon\RteRibbon.xml` | The ribbon's markup: tabs, groups, commands, KeyTips and tooltips |
| `ribbon\RteRibbon.ribbon` | The compiled ribbon (a resource-only DLL) that the editor loads |
| `ribbon\build_ribbon.ahk`, `ribbon\make_icons.cs` | Rebuild the ribbon and its icons |
| `RteClient2.ahk` | The `Rte` class for scripts that control the editor |
| `rte_macros\` | Example macros |
| `examples\Showcase.rtf` | The showcase document in the picture above |
| `Csharp-in-AHK.ahk` | A copy of [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK), matching its commit `85aea2e` |

## License

AhkWord is under the [MIT license](LICENSE). The components it uses are under their own licenses: see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) and [`licenses/`](licenses).
