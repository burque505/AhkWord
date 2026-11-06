# Documentos y archivos

## Abrir y guardar { #opening-and-saving }

**Archivo > Abrir** abre documentos, y **Guardar** y **Guardar como** los guardan, en estos formatos:

| Formato | Notas |
|---|---|
| **Texto RTF** (`.rtf`) | El formato propio de AhkWord: lo conserva todo, incluidos los estilos y la numeración. Word y LibreOffice también lo abren. |
| **Word** (`.docx`, `.doc`) y **OpenDocument** (`.odt`) | Mediante un convertidor: Word, LibreOffice, el integrado o pandoc (vea [File formats and converters](../reference/formats.md), en inglés). |
| **HTML** (`.html`) | HTML legible, que conserva los estilos, las tablas y el espaciado. |
| **Markdown** (`.md`) | La variante de GitHub, con tablas e imágenes. |
| **Texto sin formato** (`.txt`) | Solo el texto. |

También puede soltar un archivo en la ventana para abrirlo. **Archivo > Exportar** guarda una copia como **PDF** (mediante *Microsoft Print to PDF*), Word u OpenDocument, y deja el documento mismo como está; **Disposición > Exportar como PDF** hace lo mismo para PDF.

Un documento abierto desde Word conserva su aspecto: sus propios estilos, espaciado, sangrías y tablas. Vea [Estilos](styles.md) para darle los suyos.

## Documentos nuevos y plantillas { #new-documents-and-templates }

**Archivo > Nuevo** empieza un documento vacío con sus estilos. **Archivo > Nuevo desde plantilla** empieza uno a partir de una plantilla (Letter, Memo, Report, Court Paper, Meeting Notes o una propia) y pide sus campos, como los nombres de las partes o la fecha de una audiencia. Vea [Templates](templates.md) (en inglés).

**Archivo > Página de inicio** muestra la página de inicio: un documento en blanco, las plantillas, sus documentos recientes y las macros. Puede abrirse sola al iniciar AhkWord.

## Varias ventanas { #several-windows }

**Archivo > Nueva ventana** y **Archivo > Abrir en una ventana nueva** abren otra ventana de AhkWord, para trabajar en dos documentos a la vez. Cada ventana es un editor completo; comparten su configuración sin deshacer los cambios de las demás. En la barra de tareas se agrupan como las de AhkWord.

## Guardado automático y recuperación { #autosave-and-recovery }

Mientras un documento tiene cambios sin guardar, AhkWord escribe una copia de recuperación cada minuto (nunca sobre el archivo del documento). Si se cierra sin terminar bien (un bloqueo, un corte de luz, un reinicio de Windows), el siguiente inicio ofrece recuperar el documento. Guardar reemplaza el archivo original; Guardar como conserva ambos.

## Documentos recientes { #recent-documents }

El menú **Archivo** muestra sus documentos recientes, y AhkWord recuerda el tamaño y la posición de la ventana y su configuración de página.
