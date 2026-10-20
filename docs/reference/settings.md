# Settings and folders

## Where AhkWord keeps its files

| Folder | What |
|---|---|
| `%APPDATA%\AhkWord` | `settings.json` (your settings), `Templates` (your own templates, and their `Parts`), `Recovery` (the recovery copies of unsaved documents) |
| `%LOCALAPPDATA%\AhkWord` | the converters' working files (LibreOffice's profile, the converters' log) |
| `rte_macros`, beside AhkWord | the macros, and their own settings (`My signature.ini`) |

Type `%APPDATA%\AhkWord` into Explorer's address bar to open the folder.

!!! note "From RTE Writer"
    AhkWord was RTE Writer until iteration 38. Its first start copies `%APPDATA%\RTE Writer` and `%LOCALAPPDATA%\RTE Writer` into AhkWord's folders, settings, templates and recovery copies included, and leaves the old folders as they were.

## Portable mode

Put an empty file named `portable` (or `portable.txt`) beside `AhkWord.exe` or `AhkWord.ahk`, and AhkWord keeps its settings, templates, recovery copies and the converters' state in a `data` folder beside itself instead, so the folder can be carried to another computer. The .NET packages and the C# compiler stay in the computer's own caches, so a computer that hasn't run AhkWord before downloads them once.

## Environment variables

| Variable | Does |
|---|---|
| `AHKWORD_DATA` | a folder for all of AhkWord's files (settings, templates, recovery copies, the converters' state), as portable mode's `data` folder |
| `AHKWORD_SETTINGS` | the settings file to use instead of `settings.json` |
| `AHKWORD_CONVERTERS` | the converters to use for Word and OpenDocument files, only those and in that order: `word`, `libreoffice`, `openxml`, `pandoc` (for example `libreoffice,openxml`) |

The old names `RTE_WRITER_SETTINGS` and `RTE_WRITER_CONVERTERS` still work.

## Several windows

The windows of AhkWord share one settings file. Each saves only what it changed, so closing one doesn't undo another's changes: recent documents, macro hotkeys and remembered values are merged item by item.
