# Scripts, events and macros

Any AutoHotkey v2 script can control a running AhkWord, through `RteClient2.ahk` (beside `AhkWord.ahk`):

```autohotkey
#Requires AutoHotkey v2.0
#Include "RteClient2.ahk"

Rte.InsertText("Hello from AutoHotkey")
Rte.InsertTable(3, 4)
MsgBox Rte.WordCount() " words"
Rte.On("Saved", (e) => TrayTip(Rte.FileName, "Saved"))
```

Everything AhkWord can do is in the [API reference](api.md); in AhkWord, **Tools > Automation API** lists it too.

## Calls

- `Rte.Name(arguments)` calls a method and waits for its result; `Rte.Name` reads a property and `Rte.Name := value` sets one.
- Arguments are text, numbers, 1 or 0 for yes or no, enumerations by name (`Heading1`, `LowerAlpha`) and colours as `#RRGGBB` or a name.
- `Rte.Post(Name, arguments...)` queues a call and returns at once; its errors go to `Rte.LastAsyncError` and the `AsyncError` event.
- A call that fails throws an AutoHotkey error with AhkWord's message.

With several AhkWord windows open, a script talks to the one used last. A macro started from an AhkWord window talks to that window.

## Events

`Rte.On(Event, Callback)` calls `Callback(e)` when something happens in the editor, and `Rte.Off(Event [, Callback])` stops it. Every `e` has `e.Name`; some events say more:

| Event | When | Also in `e` |
|---|---|---|
| `TextChanged` | the text changed | `e.TextLength` |
| `SelectionChanged` | the selection or caret moved | `e.Start`, `e.Length` |
| `ModifiedChanged` | the document became changed, or saved | `e.Modified` |
| `Saved` | the document was saved | |
| `Opened` | a document was opened, or a new one started | |
| `Closed` | the editor closed | |
| `AsyncError` | a call queued with `Rte.Post` failed | |

## Macros

- **Tools > Macros** runs the scripts in the `rte_macros` folder (beside AhkWord); **Macros Folder** opens it.
- **Macro Hotkeys** gives any macro a hotkey.
- **Record Macro** (Ctrl+Shift+R) records your typing and ribbon commands as a new macro script: each line of it is an AhkWord call, so a recorded macro is also a good way to learn the calls.
- **Run Script** runs any AutoHotkey script, connected to this window.

### The example macros

- **Clean up pasted text:** joins lines broken in the middle of sentences, tidies spaces, and makes quotes and dashes typographic.
- **My signature:** inserts your name, title, phone and e-mail (asked for once, then remembered).
- **Collect from anywhere:** select text in any program and press Win+Shift+C to add it, with its source and time, to the end of the document.
- **PDF on save:** each save also makes a PDF beside the document. It is a Saved hook: run it from Tools > Macros, or with a hotkey from Macro Hotkeys, to switch it on or off (it asks before taking the place of another Saved hook).
- **Folder to PDF:** makes a PDF of every document in a folder.
- **Arrange windows:** puts the open AhkWord windows side by side on the screen, each page fitted to its window's width.
- **Open AutoHotkey Dash:** AutoHotkey's Dash, for its help, Window Spy and **Editor settings**, the program that edits macros (see [Getting started](../getting-started.md#which-program-opens-a-macro-to-edit-it)). It needs AutoHotkey installed.
- An auto-save, a word goal tracker, an event monitor, document statistics, and Title Case for the selection.

## Event hooks

**Tools > Event Hooks** gives an event a macro that AhkWord runs by itself when it happens. Unlike a script with `Rte.On`, nothing has to keep running: the macro starts, does its work and ends.

| Event | When | `A_Args[2]` |
|---|---|---|
| Started | AhkWord's window has opened, with its first document | |
| Opened | a document was opened | its file |
| NewDocument | File > New | |
| BeforeSave | a document is about to be saved: AhkWord waits for the macro | the file |
| Saved | a document was saved | its file |
| StyleChanged | a style was changed, made, deleted or copied, or the style set changed | the style |
| Printed | the document was printed or exported as PDF | the printer, or the PDF |
| Closing | the window is about to close: AhkWord waits for the macro | the file |

The macro gets the event's name in `A_Args[1]` (and in the environment variable `AHKWORD_EVENT`), and what it was about in `A_Args[2]` (`AHKWORD_EVENT_DETAIL`). It works on the window the event happened in through `Rte`, as any macro does. **BeforeSave** and **Closing** wait for their macro, for up to 15 seconds, so that it can change the document before it is saved, or save it before the window closes; the other events don't wait. While an event's macro is running, that event doesn't start it again, so a Saved hook that saves doesn't run in a loop.

Untick **On** beside a macro to keep it chosen without it running: a PDF on each save only some of the time, say; tick it again to have it back. In the dialog, **New...** makes a macro for the event (`On Saved.ahk`, say) from a starting script with an example in it, and opens it for editing; **Run Now** tries the macro as the event would run it. Scripts use `SetEventHook("Saved", "On Saved.ahk")`, `SetEventHookOn("Saved", false)` (so a hotkey can switch a hook), `EventHookOn`, `EventHook`, `EventHookList` and `HookEventNames`.

An example, a Saved hook that keeps a dated copy of each save in a `Backups` folder beside the document:

```autohotkey
#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
SavedFile := A_Args.Length >= 2 ? A_Args[2] : ""
if SavedFile != "" {
    SplitPath(SavedFile, , &Dir, &Ext, &Stem)
    DirCreate(Dir "\Backups")
    FileCopy(SavedFile, Dir "\Backups\" Stem " " FormatTime(, "yyyy-MM-dd HHmm") "." Ext, true)
}
```

## Hotkeys in AhkWord.ahk

`AhkWord.ahk` itself defines two hotkeys, active only in AhkWord's windows: Ctrl+Alt+D inserts the date, and Ctrl+Alt+W shows word counts in the status bar. Add your own there; they call the editor directly, in the same process.
