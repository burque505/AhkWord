# Troubleshooting

## AhkWord doesn't start

**Nothing happens for a while on the first start.** That is expected: the first start downloads AhkWord's components and compiles it, with nothing shown until the window opens. Give it a minute or two.

**An error about .NET.** AhkWord needs the .NET 10 Desktop Runtime; an older .NET is not enough. Install it with `winget install Microsoft.DotNet.DesktopRuntime.10` and start AhkWord again.

**An error with `NU1100` or `NU1301`.** The components couldn't be downloaded from nuget.org. Check that the computer can reach nuget.org (a proxy or a firewall can stop it), and the package sources the message lists.

**A message that AhkWord stopped because of an error.** It shows the error's text. Please report it, with that text, on the [issues page](https://github.com/burque505/AhkWord/issues).

## Word and OpenDocument files

**A Word file opens or saves slowly, or through another converter than you expected.** Word is used only while it isn't open: with Word open, AhkWord uses LibreOffice or the built-in converter. Close Word for Word's own conversion.

**A converter fails.** AhkWord tries the next one, and says which it used in the status bar. The converters' log is in `%LOCALAPPDATA%\AhkWord\converters.log`.

## Lists and numbering

**Nested list items all show the first level's numbering (A. A. A.).** In List Levels, tick *Tab and Shift+Tab give an item its new level's numbering style*.

**A paragraph of a numbered style has no number.** Its numbering is turned off: Tab at its start makes it a sub-level item, or the Numbering button numbers it again. See [Numbering off](../guide/lists.md#numbering-off).

## Recovering a document

If AhkWord ended without closing, its next start offers the documents it was editing. Their copies are in `%APPDATA%\AhkWord\Recovery` until you recover or discard them.

## The taskbar shows AutoHotkey's icon

Close every AhkWord window and start it again: Windows sometimes keeps a taskbar button's old icon until the program starts afresh.
