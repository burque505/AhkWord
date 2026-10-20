# Requirements

| Component | Notes |
|---|---|
| **Windows** | 10 or 11 |
| **.NET** | the **.NET 10 Desktop Runtime** (x64). The runtime is enough: no SDK is needed. `winget install Microsoft.DotNet.DesktopRuntime.10` |
| **AutoHotkey v2**, 64-bit | only to run `AhkWord.ahk` from the repository; the release download's `AhkWord.exe` is AutoHotkey |
| **C# compiler** | nothing to install: if none is found, it is downloaded on the first start (Microsoft's compiler package from nuget.org, about 22 MB, kept in `%LOCALAPPDATA%\Csharp-in-AHK\roslyn`). Visual Studio or its Build Tools are used if installed |

Having .NET 9 or older is not enough: without .NET 10, AhkWord stops with an error at the start.

## What the first start downloads

The first start downloads the .NET packages AhkWord is built from (RichEditorNET, WindowsRibbon, the Open XML libraries, SkiaSharp and others) from nuget.org into your computer's NuGet folder (`%USERPROFILE%\.nuget\packages`), as any .NET program does, and the C# compiler if none is installed. It takes a minute or two, with nothing shown until the window opens. Later starts use them from there, without the network.

## Optional

- **Microsoft Word** or **LibreOffice** convert Word and OpenDocument files best; without them, the built-in converter reads and writes `.docx` files. See [File formats and converters](formats.md).
- **pandoc**, as a last resort for formats the others can't do.
