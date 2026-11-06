# Formatos de archivo y convertidores

## Lo que conserva cada formato { #what-each-format-keeps }

| | RTF | Word (.docx) | OpenDocument (.odt) | HTML | Markdown |
|---|---|---|---|---|---|
| Estilos, con sus definiciones | sí | sí | sí | sí | títulos y listas |
| Estilos numerados y niveles de lista | sí | sí, como listas de Word | sí | números | números |
| Interlineado, espacio antes y después | sí | sí | sí | sí | no |
| Sangrías y tabulaciones | sí | sí | sí | sí | no |
| Tablas: anchos, bordes, sombreado | sí | sí | sí | sí | no |
| Encabezados y pies de página | sí | sí | sí | sí | sí |
| Números de línea | sí | sí | sí | sí | sí |
| Saltos de página | sí | sí | sí | sí | sí |
| Imágenes con texto alternativo | sí | sí | sí | sí | sí |

RTF es el formato propio de AhkWord y lo conserva todo. Markdown es texto sin formato por diseño, así que no tiene formato de tablas ni espaciado de párrafos. Parte del formato que Markdown simple no puede escribir se conserva con las extensiones de [M-ArkDown](https://github.com/TheArkive/M-ArkDown_ahk2), de TheArkive, cuyas ideas sigue AhkWord con agradecimiento: el estilo de numeración de una lista como `[type=a]` después de su primer elemento, y el tamaño de una imagen como `(width=200)` después de ella; los colores se escriben como `<span style>` de HTML, y también se lee el `<c=#F00>` de M-ArkDown.

## Convertidores de Word y OpenDocument { #word-and-opendocument-converters }

Los archivos de Word y OpenDocument los convierte el primero de estos que esté disponible:

1. **Microsoft Word**, ejecutado de forma invisible, y solo cuando Word no está abierto (AhkWord nunca toca el Word en que usted está trabajando).
2. **LibreOffice**, ejecutado en segundo plano con un perfil propio.
3. **El convertidor integrado** (solo `.docx`, sin nada que instalar): conserva los anchos, bordes, sombreado y alineación de las tablas, y las sangrías, tabulaciones y espaciado de los párrafos.
4. **pandoc**.

`AHKWORD_CONVERTERS` elige cuáles de ellos se usan, y en qué orden (vea [Configuración y carpetas](settings.md#environment-variables)), y los scripts pueden indicar el convertidor para un archivo: `Rte.Open(path, "libreoffice")`, `Rte.ExportDocx(path, "openxml")`.

Sea cual sea el convertidor que abre un `.docx`, sus párrafos reciben el interlineado y el espacio antes y después que les da Word, calculados a partir de los valores predeterminados del archivo, sus estilos y el formato propio de cada párrafo.

## Listas y numeración a través de Word { #lists-and-numbering-through-word }

Las listas corrientes, la numeración legal y los estilos numerados conservan sus números, niveles y sangrías a través de los archivos de Word: AhkWord los escribe como definiciones de lista propias de Word, y lee de vuelta las listas, los niveles de lista y los estilos numerados de un documento de Word. Las sangrías de los niveles de lista salen de donde Word dibuja los párrafos de la lista.
