# AhkWord

A word processor for Windows, written in C# and run inside AutoHotkey v2, that AutoHotkey scripts can drive while you type.

AhkWord is built on [RichEditorNET](https://github.com/ellipticbit/RichEditorNET), a .NET wrapper around the Windows RichEdit control. [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compiles its C# source when it starts and runs it in the AutoHotkey process, so there is no build step and no separate executable.

> **Work in progress.** The editor is still named **RTE Writer** in the code, its window and its settings folder, and its files are numbered by iteration (`rte_writer15.*`). Renaming them to AhkWord is planned.

## Features

- **Documents:** HTML, Markdown (GitHub flavor, with tables and images), plain text and RTF. Word (`.docx`, `.doc`) and OpenDocument (`.odt`) files are opened and saved through a converter (see [Optional converters](#optional-converters)). Export as PDF through "Microsoft Print to PDF".
- **Formatting:** fonts, colors, highlight, alignment and headings, from the toolbar and menus.
- **Lists:** numbered and bulleted lists nested with Tab and Shift+Tab, a numbering style per level, restart and continue numbering, and legal numbering (1., 1.1, 1.1.1). Typed numbers in documents from elsewhere can be turned into real lists.
- **Tables:** insert tables, add and delete rows and columns, move between cells with Tab, and use the table commands from the right-click menu.
- **Pages:** manual page breaks (Ctrl+Enter), headers and footers with page-number and date fields, Page Setup, Print Preview and printing. Headers and footers are shown in the editor and saved in the document.
- **Pictures:** insert from a file or by drag and drop, resize, and give them alt text (shown when the mouse rests on a picture).
- **Writing tools:** spelling suggestions from the Windows spell checker (F7 goes to the next error) and an outline panel of the headings.
- **Comfort:** recent files and remembered window and page settings. Drop a file on the window to open it.

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

- **Calls** wait for the editor and return its result. `Rte.Post(...)` queues a call and returns at once. **Help > Automation API** in the editor lists everything that can be called.
- **Events** (`TextChanged`, `SelectionChanged`, `ModifiedChanged`, `Saved`, `Opened`, `Closed`, `AsyncError`) are delivered to callbacks registered with `Rte.On`.
- **Macros:** **Tools > Macros** runs the scripts in [`rte_macros`](rte_macros), and **Tools > Macro hotkeys** gives any of them a hotkey. **Tools > Record macro** (Ctrl+Shift+R) records typing and menu commands as a new macro script.
- **Hotkeys in the host script:** `rte_writer15.ahk` itself defines a couple of hotkeys (Ctrl+Alt+D inserts the date, Ctrl+Alt+W shows word counts). Add your own there.

The example macros include an auto-save, a word goal tracker, an event monitor, document statistics and Title Case for the selection.

## Requirements

| Component | Notes |
|---|---|
| Windows | 10 or 11 |
| AutoHotkey v2 | **64-bit** (`AutoHotkey64.exe`) |
| .NET runtime | .NET 10, including **Microsoft.WindowsDesktop.App** |
| Roslyn compiler | The `Roslyn` folder containing `csc.exe`, e.g. from Visual Studio or Build Tools |

Csharp-in-AHK finds .NET and Roslyn by itself. If it can't, see [its README](https://github.com/burque505/Csharp-in-AHK#readme) for how to set the paths.

NuGet packages (RichEditorNET, the Open XML libraries, SkiaSharp and others) are downloaded on the first start, which therefore takes longer.

### Optional converters

Word and OpenDocument files are converted by the first of these that is available:

1. **Microsoft Word**, run invisibly, and only when Word isn't already open.
2. **LibreOffice**, run headless with its own profile.
3. **The built-in Open XML converter** (`.docx` only, no extra software needed).
4. **pandoc**.

HTML, Markdown, RTF and plain text need none of them.

## Getting started

1. Install the requirements above.
2. Clone this repository.
3. Run `rte_writer15.ahk`. To open a document at start, pass its path: `rte_writer15.ahk "C:\path\document.md"`.

Settings are kept in `%APPDATA%\RTE Writer\settings.json`. Set the environment variable `RTE_WRITER_SETTINGS` to a file path to keep them elsewhere.

## Files

| File | Purpose |
|---|---|
| `rte_writer15.ahk` | Starts the editor and holds its in-process hotkeys |
| `rte_writer15.cs` | The editor window, menus and the automation interface |
| `rte_markdown5.cs` | Markdown reading and writing |
| `rte_htmllists2.cs` | HTML list numbering (start numbers, restarts, styles) |
| `rte_tables.cs` | Tables |
| `rte_legal.cs` | Legal numbering |
| `rte_pagebreaks.cs` | Page breaks |
| `rte_headerfooter.cs` | Headers and footers |
| `rte_office6.cs` | Word, OpenDocument and RTF conversion |
| `RteClient2.ahk` | The `Rte` class for scripts that control the editor |
| `rte_macros\` | Example macros |
| `Csharp-in-AHK.ahk` | A copy of [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK), matching its commit `85aea2e` |

## License

[MIT](LICENSE)
