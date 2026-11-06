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

## Language

AhkWord's interface (the ribbon, the menus, the status bar, the dialogs and the messages) is in English, or in Spanish: choose it in **Review > Language** (or the command palette's *Language: Español*), and it shows the next time AhkWord starts (Windows' own dialogs, such as Open, Save and Print, are in Windows' language). That is the setting `"Language": "es"` in `settings.json`, which can also be set by hand while AhkWord is closed. The texts are in the `lang` folder beside AhkWord: `en.json` lists every English text with its key, `es.json` the Spanish; a text a language lacks shows in English. A folder beside AhkWord with a `lang` folder of its own (an add-on's) can have its texts there too, as `<folder>\lang\es.json`: they are read after AhkWord's.

## Environment variables

| Variable | Does |
|---|---|
| `AHKWORD_DATA` | a folder for all of AhkWord's files (settings, templates, recovery copies, the converters' state), as portable mode's `data` folder |
| `AHKWORD_LANG` | the interface's language over the setting `Language` (`es`, or `en`), to try one without changing the settings |
| `AHKWORD_SETTINGS` | the settings file to use instead of `settings.json` |
| `AHKWORD_CONVERTERS` | the converters to use for Word and OpenDocument files, only those and in that order: `word`, `libreoffice`, `openxml`, `pandoc` (for example `libreoffice,openxml`) |

The old names `RTE_WRITER_SETTINGS` and `RTE_WRITER_CONVERTERS` still work.

## Several windows

The windows of AhkWord share one settings file. Each saves only what it changed, so closing one doesn't undo another's changes: recent documents, macro hotkeys and remembered values are merged item by item.
