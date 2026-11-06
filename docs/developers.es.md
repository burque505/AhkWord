# Para desarrolladores

AhkWord es C# que [Csharp-in-AHK](https://github.com/burque505/Csharp-in-AHK) compila al iniciarse y que se ejecuta dentro de AutoHotkey v2: no hay paso de compilación. Cambie un archivo `.cs`, inicie de nuevo `AhkWord.ahk`, y el cambio ya está.

## Los archivos { #the-files }

| Archivo | Función |
|---|---|
| `AhkWord.ahk` | Inicia el editor, y contiene sus teclas rápidas en el mismo proceso |
| `AhkWord.cs` | La ventana del editor, los comandos de la cinta, las vistas Markdown y HTML, y la interfaz de automatización (`EditorApi`) |
| `rte_styles.cs` | Estilos de párrafo: los conjuntos de estilos, aplicar y cambiar estilos, los estilos en RTF y HTML, Copiar estilos, el cuadro de diálogo Modificar estilo |
| `rte_legal.cs` | La numeración legal, los números de los estilos numerados, las sangrías de las listas |
| `rte_office7.cs` | La conversión de Word, OpenDocument y RTF |
| `rte_docxstyles.cs` | Los estilos en los archivos `.docx` del convertidor integrado |
| `rte_markdown6.cs`, `rte_htmllists2.cs`, `rte_linebreaks.cs` | Markdown y HTML |
| `rte_tables2.cs` | Las tablas y sus propiedades |
| `rte_paragraph.cs` | Los cuadros de diálogo Párrafo y Tabulaciones, y dónde terminan las páginas impresas |
| `rte_linespacing2.cs` | El interlineado y el espaciado de párrafos |
| `rte_printlayout.cs`, `rte_pagegaps.cs`, `rte_pageview.cs` | Diseño de impresión, Mostrar páginas, Varias páginas y Vista previa de impresión |
| `rte_pagebreaks.cs`, `rte_headerfooter.cs`, `rte_linenumbers2.cs` | Saltos de página, encabezados y pies de página, números de línea |
| `rte_sourcebox.cs` | El editor de código de las vistas Markdown y HTML (Scintilla) |
| `rte_templates.cs` | Las plantillas y la página de inicio |
| `rte_recovery.cs` | Las copias de recuperación, y las carpetas de AhkWord |
| `RteClient2.ahk` | La clase `Rte` para los scripts |
| `ribbon\RteRibbon.xml`, `ribbon\RteRibbon.ribbon` | El marcado de la cinta, y su forma compilada |
| `images\icon\` | El icono de AhkWord, y `make_icon.py`, que lo dibuja |

## Volver a compilar la cinta { #rebuilding-the-ribbon }

La cinta viene compilada, así que para ejecutar AhkWord no hace falta nada más. Después de cambiar `ribbon\RteRibbon.xml` o los iconos, ejecute `ribbon\build_ribbon.ahk`. Necesita el **Windows SDK** (`uicc.exe`, `rc.exe`) y las **herramientas de escritorio de C++** de Visual Studio o de Build Tools (`link.exe`). Descarga una vez la fuente [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons), dibuja los iconos a partir de ella y compila el marcado.

## El icono { #the-icon }

`images\icon\make_icon.py` (Python 3 con numpy y matplotlib) dibuja `ahkword.ico` y sus PNG; `python make_icon.py preview` crea también hojas de vista previa de cada tamaño sobre fondos claros y oscuros.

## Versiones y publicaciones { #versions-and-releases }

- `make_version.ps1 41` crea en `..\rte-versions` una copia ejecutable de una iteración confirmada, y `-Working` una de los archivos tal como están ahora. Cada copia se ejecuta en modo portátil, así que nunca toca su configuración.
- `make_release.ps1` crea la descarga de la versión, `..\AhkWord-releases\AhkWord-NN.zip`, a partir del último commit: los archivos del programa, `AhkWord.exe` (una copia del AutoHotkey64.exe instalado), la licencia de AutoHotkey y una nota de su versión y su origen, y `START HERE.txt`.

## Esta documentación { #this-documentation }

Las páginas son Markdown en `docs\`, compiladas con [MkDocs](https://www.mkdocs.org/) y el tema [Material](https://squidfunk.github.io/mkdocs-material/) (`pip install mkdocs-material`). `python make_docs.py` escribe la [referencia de la API](automation/api.md) (en inglés) a partir de la `EditorApi` de `AhkWord.cs` (cada miembro público y su comentario `///`); `python make_docs.py serve` muestra una vista previa del sitio en `http://127.0.0.1:8000`, y `python make_docs.py build` lo compila en `site\`. En GitHub, la acción *Documentation* (`.github/workflows/docs.yml`) lo compila y lo publica en [burque505.github.io/AhkWord](https://burque505.github.io/AhkWord/) en cada push que cambia las páginas o `AhkWord.cs`.
