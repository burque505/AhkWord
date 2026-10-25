# Third-party notices

AhkWord itself is under the [MIT license](LICENSE). It builds on the components below, each under its own license; the full texts are in [`licenses/`](licenses).

## NuGet packages

These are not included in this repository: Csharp-in-AHK downloads them from nuget.org when the editor first starts from it. The release download carries them (in `lib\packages`), so it downloads nothing.

| Package | Used for | License | Text |
|---|---|---|---|
| [EllipticBit.RichEditorNET](https://ellipticbit.com/) | The editor control (RichEdit for .NET) | Boost Software License 1.0 | [RichEditorNET-BSL-1.0.txt](licenses/RichEditorNET-BSL-1.0.txt) |
| [WindowsRibbon](https://github.com/harborsiem/WindowsRibbon) | The ribbon (a .NET wrapper of the Windows Ribbon framework) | MS-PL or MIT, at your choice (used here under MIT) | [WindowsRibbon-MIT.txt](licenses/WindowsRibbon-MIT.txt) |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK), DocumentFormat.OpenXml.Framework | Reading and writing Word files (built-in converter) | MIT | [Open-XML-SDK-MIT.txt](licenses/Open-XML-SDK-MIT.txt) |
| [HtmlToOpenXml.dll](https://github.com/onizet/html2openxml) | Writing Word files (built-in converter) | MIT | [HtmlToOpenXml-MIT.txt](licenses/HtmlToOpenXml-MIT.txt) |
| [AngleSharp](https://github.com/AngleSharp/AngleSharp) | HTML parsing for HtmlToOpenXml | MIT | [AngleSharp-MIT.txt](licenses/AngleSharp-MIT.txt) |
| [Codeuctivity.OpenXmlPowerTools](https://github.com/Codeuctivity/OpenXmlPowerTools) | Reading Word files (built-in converter) | MIT | [OpenXmlPowerTools-MIT.txt](licenses/OpenXmlPowerTools-MIT.txt) |
| [Microsoft.Extensions.Logging.Abstractions](https://github.com/dotnet/runtime) | Needed by HtmlToOpenXml | MIT | [Microsoft.Extensions.Logging.Abstractions-MIT.txt](licenses/Microsoft.Extensions.Logging.Abstractions-MIT.txt) |
| [SkiaSharp](https://github.com/mono/SkiaSharp), SkiaSharp.NativeAssets.Win32 (its native library) | Picture conversion for Word files | MIT | [SkiaSharp-MIT.txt](licenses/SkiaSharp-MIT.txt) |
| [Scintilla5.NET](https://github.com/desjarlais/Scintilla.NET) | The editor of the Markdown and HTML views (a .NET wrapper of Scintilla) | MIT | [Scintilla.NET-MIT.txt](licenses/Scintilla.NET-MIT.txt) |
| [Scintilla](https://www.scintilla.org/) and [Lexilla](https://www.scintilla.org/Lexilla.html), in the Scintilla5.NET package | The editing component, and its Markdown and HTML syntax colouring | The Scintilla license (permissive, by Neil Hodgson) | [Scintilla-Lexilla-License.txt](licenses/Scintilla-Lexilla-License.txt) |
| [PDFsharp](https://github.com/empira/PDFsharp) 6.2.4 | PDF bookmarks and links (Export as PDF). Pinned to 6.2.4: later versions are under the AGPL | MIT | [PDFsharp-MIT.txt](licenses/PDFsharp-MIT.txt) |
| [System.Security.Cryptography.Pkcs](https://github.com/dotnet/runtime), [System.IO.Packaging](https://github.com/dotnet/runtime) | Needed by PDFsharp and the Open XML SDK | MIT | [dotnet-MIT.txt](licenses/dotnet-MIT.txt) |

## Included in this repository

| Component | Where | License | Text |
|---|---|---|---|
| [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) (Microsoft) | The ribbon's icons in `ribbon/RteRibbon.ribbon`, drawn from the icon font by `ribbon/make_icons.cs`. The font itself is downloaded by `ribbon/build_ribbon.ahk` when the ribbon is rebuilt, and is not included. | MIT | [FluentUISystemIcons-MIT.txt](licenses/FluentUISystemIcons-MIT.txt) |
| [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) | `Csharp-in-AHK.ahk`, `Csx-in-AHK.ahk`, `Csharp-in-AHK-Paths.ahk`, `Csharp-in-AHK.example.ini` | MIT, by the author of AhkWord | [LICENSE](LICENSE) |

## In the release download only

| Component | Where | License | Text |
|---|---|---|---|
| [AutoHotkey](https://www.autohotkey.com/) v2 (64-bit) | `AhkWord.exe`: AutoHotkey64.exe, unchanged, under AhkWord's name, so that AhkWord runs without AutoHotkey installed. Not in this repository; `make_release.ps1` copies it from an installed AutoHotkey into the download, with its license and `AhkWord.exe.txt`, which names its version and where its source code is (github.com/AutoHotkey/AutoHotkey, that version's tag). | GNU General Public License v2 | `licenses/AutoHotkey-GPL-2.0.txt`, in the download |

| [Roslyn](https://github.com/dotnet/roslyn), the C# compiler 5.9.0 (Microsoft.CodeAnalysis, Microsoft.CodeAnalysis.CSharp) | `lib\roslyn`: AhkWord compiles its C# source when it starts. From the [Microsoft.Net.Compilers.Toolset](https://www.nuget.org/packages/Microsoft.Net.Compilers.Toolset) package, which Csharp-in-AHK downloads when it runs from this repository on a computer without Visual Studio. `lib\host` is Csharp-in-AHK's host, built with it. | MIT | [dotnet-MIT.txt](licenses/dotnet-MIT.txt) |
| [.NET](https://github.com/dotnet/runtime) 10 Desktop Runtime (Microsoft) | `dotnet`, in the complete download only (the other one downloads it from Microsoft when asked to: `AhkWordSetup.ahk`). Microsoft's own zips of the .NET runtime and the Windows Desktop runtime, unchanged. | MIT | `dotnet/LICENSE.txt` and `dotnet/ThirdPartyNotices.txt`, in the download |

AhkWord.exe is only bundled with AhkWord's files; AhkWord itself stays under the MIT license.

## Acknowledgements

| Work | What AhkWord owes it | License | Text |
|---|---|---|---|
| [M-ArkDown](https://github.com/TheArkive/M-ArkDown_ahk2), by TheArkive | The Markdown extensions AhkWord uses for what plain Markdown can't say: a list's numbering style (`[type=a]`) and a picture's size (`(width=200)`), written and read, and colours (`<c=#F00>`), read. AhkWord follows its ideas with its own code (`rte_markdown6.cs`). | MIT | [M-ArkDown-MIT.txt](licenses/M-ArkDown-MIT.txt) |

## Programs used if installed

Microsoft Word, LibreOffice and pandoc convert Word and OpenDocument files when they are installed. They are separate programs under their own licenses, and are neither included nor required. The Windows SDK and the Visual Studio C++ tools are needed only to rebuild the ribbon.
