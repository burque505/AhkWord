# Requisitos

| Componente | Notas |
|---|---|
| **Windows** | 10 u 11 |
| **.NET** | el **.NET 10 Desktop Runtime** (x64). Basta con el runtime: no hace falta el SDK. `winget install Microsoft.DotNet.DesktopRuntime.10` |
| **AutoHotkey v2**, de 64 bits | solo para ejecutar `AhkWord.ahk` desde el repositorio; el `AhkWord.exe` de la versión descargada es AutoHotkey |
| **Compilador de C#** | no hay nada que instalar: si no se encuentra ninguno, se descarga en el primer inicio (el paquete del compilador de Microsoft desde nuget.org, unos 22 MB, que se guarda en `%LOCALAPPDATA%\Csharp-in-AHK\roslyn`). Si están instalados, se usan Visual Studio o sus Build Tools |

Tener .NET 9 o anterior no basta: sin .NET 10, AhkWord se detiene con un error al iniciarse.

## Lo que descarga el primer inicio { #what-the-first-start-downloads }

El primer inicio descarga desde nuget.org los paquetes de .NET con que se construye AhkWord (RichEditorNET, WindowsRibbon, las bibliotecas Open XML, SkiaSharp y otros) en la carpeta NuGet de su equipo (`%USERPROFILE%\.nuget\packages`), como hace cualquier programa .NET, y el compilador de C# si no hay ninguno instalado. Tarda uno o dos minutos, sin que se muestre nada hasta que se abre la ventana. Los inicios siguientes los usan desde allí, sin la red.

## Opcional { #optional }

- **Microsoft Word** o **LibreOffice** convierten mejor los archivos de Word y OpenDocument; sin ellos, el convertidor integrado lee y escribe archivos `.docx`. Vea [Formatos de archivo y convertidores](formats.md).
- **pandoc**, como último recurso para los formatos que los demás no pueden manejar.
