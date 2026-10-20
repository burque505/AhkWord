# For developers

AhkWord is C# compiled when it starts by [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) and run inside AutoHotkey v2: there is no build step. Change a `.cs` file, start `AhkWord.ahk` again, and the change is in.

## The files

| File | Purpose |
|---|---|
| `AhkWord.ahk` | Starts the editor, and holds its in-process hotkeys |
| `AhkWord.cs` | The editor window, the ribbon's commands, the Markdown and HTML views, and the automation interface (`EditorApi`) |
| `rte_styles.cs` | Paragraph styles: the style sets, applying and changing styles, styles in RTF and HTML, Copy Styles, the Modify Style dialog |
| `rte_legal.cs` | Legal numbering, numbered styles' numbers, list indents |
| `rte_office7.cs` | Word, OpenDocument and RTF conversion |
| `rte_docxstyles.cs` | Styles in the built-in converter's `.docx` files |
| `rte_markdown6.cs`, `rte_htmllists2.cs`, `rte_linebreaks.cs` | Markdown and HTML |
| `rte_tables2.cs` | Tables and their properties |
| `rte_paragraph.cs` | The Paragraph and Tabs dialogs, and where printed pages end |
| `rte_linespacing2.cs` | Line and paragraph spacing |
| `rte_printlayout.cs`, `rte_pagegaps.cs`, `rte_pageview.cs` | Print Layout, Show Pages, Multiple Pages and Print Preview |
| `rte_pagebreaks.cs`, `rte_headerfooter.cs`, `rte_linenumbers2.cs` | Page breaks, headers and footers, line numbers |
| `rte_sourcebox.cs` | The Markdown and HTML views' source editor (Scintilla) |
| `rte_templates.cs` | Templates and the start page |
| `rte_recovery.cs` | Recovery copies, and AhkWord's folders |
| `RteClient2.ahk` | The `Rte` class for scripts |
| `ribbon\RteRibbon.xml`, `ribbon\RteRibbon.ribbon` | The ribbon's markup, and its compiled form |
| `images\icon\` | AhkWord's icon, and `make_icon.py`, which draws it |

## Rebuilding the ribbon

The ribbon comes compiled, so running AhkWord needs nothing more. After changing `ribbon\RteRibbon.xml` or the icons, run `ribbon\build_ribbon.ahk`. It needs the **Windows SDK** (`uicc.exe`, `rc.exe`) and the **C++ desktop tools** of Visual Studio or Build Tools (`link.exe`). It downloads the [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) font once, draws the icons from it, and compiles the markup.

## The icon

`images\icon\make_icon.py` (Python 3 with numpy and matplotlib) draws `ahkword.ico` and its PNGs; `python make_icon.py preview` also makes preview sheets of every size on light and dark backgrounds.

## Versions and releases

- `make_version.ps1 41` makes a runnable copy of a committed iteration in `..\rte-versions`, and `-Working` one of the files as they are now. Each copy runs in portable mode, so it never touches your settings.
- `make_release.ps1` makes the release download, `..\AhkWord-releases\AhkWord-NN.zip`, from the last commit: the program files, `AhkWord.exe` (a copy of the installed AutoHotkey64.exe), AutoHotkey's license and a note of its version and source, and `START HERE.txt`.

## This documentation

The pages are Markdown in `docs\`, built with [MkDocs](https://www.mkdocs.org/) and the [Material](https://squidfunk.github.io/mkdocs-material/) theme (`pip install mkdocs-material`). `python make_docs.py` writes the [API reference](automation/api.md) from `AhkWord.cs`'s `EditorApi` (each public member and its `///` comment); `python make_docs.py serve` previews the site at `http://127.0.0.1:8000`, and `python make_docs.py build` builds it into `site\`. On GitHub, the *Documentation* action (`.github/workflows/docs.yml`) builds and publishes it to [burque505.github.io/AhkWord](https://burque505.github.io/AhkWord/) on each push that changes the pages or `AhkWord.cs`.
