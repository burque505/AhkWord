# Getting started

## 1. Install .NET 10

AhkWord needs the **.NET 10 Desktop Runtime** (64-bit), once per computer. In a terminal (Windows Terminal or a command prompt):

```
winget install Microsoft.DotNet.DesktopRuntime.10
```

or download the *.NET Desktop Runtime* for x64 from [dotnet.microsoft.com](https://dotnet.microsoft.com/download/dotnet/10.0). An older .NET (9 or before) is not enough: without .NET 10, AhkWord stops with an error when it starts.

## 2. Get AhkWord

=== "The download (no AutoHotkey needed)"

    1. On the [Releases page](https://github.com/burque505/AhkWord/releases), download the newest **AhkWord-NN.zip**.
    2. Unzip it anywhere you like (not into Program Files, which Windows keeps read-only).
    3. Double-click **AhkWord.exe**.

    `AhkWord.exe` is AutoHotkey v2, unchanged, under AhkWord's name, so you don't need AutoHotkey installed. (It is under the GNU GPL; its license and source link are in the zip.) Explorer shows AutoHotkey's green **H** on the file itself; the running AhkWord shows its own icon.

=== "From the repository (AutoHotkey installed)"

    1. Install **AutoHotkey v2**, 64-bit, from [autohotkey.com](https://www.autohotkey.com/).
    2. Clone or download [the repository](https://github.com/burque505/AhkWord).
    3. Run **AhkWord.ahk**. To open a document at the start, pass its path: `AhkWord.ahk "C:\path\document.docx"`.

## 3. The first start

The first start takes a minute or two, with nothing shown until the window opens: AhkWord downloads the components it is made of (from nuget.org, into your computer's NuGet folder) and compiles itself. Later starts take a few seconds.

Then try the example documents in the `examples` folder:

- **Showcase.rtf**, the document in the picture on the [home page](index.md), made entirely with AhkWord's own commands;
- **Styles Tutorial.rtf**, a short tutorial on styles, style sets and paragraph settings, which uses them itself.

!!! tip "Pin it to the taskbar"
    Right-click AhkWord's taskbar button and choose **Pin to taskbar**: the pinned button starts AhkWord, with its own icon, apart from any other AutoHotkey scripts you run.

## Where your settings are kept

AhkWord keeps its settings, your own templates and its recovery copies in `%APPDATA%\AhkWord` (type that into Explorer's address bar to go there), and the converters' working files in `%LOCALAPPDATA%\AhkWord`. Nothing is written beside the program, so you can update AhkWord by replacing its folder.

**Portable mode:** put an empty file named `portable` (or `portable.txt`) beside `AhkWord.exe` or `AhkWord.ahk`, and AhkWord keeps all of that in a `data` folder beside itself instead, so you can carry the folder to another computer, on a USB stick for instance. [Settings and folders](reference/settings.md) has the details, and other ways to choose the folders.

## Next

- The [User guide](guide/documents.md), starting with opening and saving documents.
- [For legal drafting](legal.md), if you write pleadings, motions or contracts.
- [Troubleshooting](reference/troubleshooting.md), if the first start doesn't go as described.
