# Páginas e impresión

## Saltos de página { #page-breaks }

**Insertar > Salto de página** (Ctrl+Entrar) empieza una página nueva. En el editor se muestra como una línea discontinua.

## Configurar página { #page-setup }

**Disposición > Configurar página** (también en Archivo > Imprimir, y con un doble clic en los márgenes grises de la regla) establece, como la Configurar página de Word:

- los **márgenes**: superior, inferior, izquierdo y derecho;
- la **orientación**: vertical u horizontal;
- el **papel**: Letter, Legal, Executive, Tabloid, A3, A4, A5 o B5 (JIS), o un ancho y un alto personalizados;
- la distancia del encabezado desde arriba y la del pie desde abajo, y **Primera página diferente** (vea [Encabezados y pies de página](#headers-and-footers)).

Una imagen de la página muestra el resultado mientras los cambia. Las medidas están en pulgadas, o en centímetros donde Windows mide en ellos.

**Aceptar** los establece solo para este documento: se guardan con él (en los archivos .rtf y de Word, y en las copias de recuperación), y el papel es aquel en el que AhkWord compone las páginas y el que pide a la impresora. Un documento sin papel propio (uno de antes, o uno de otra procedencia que no lo tiene) usa el de la impresora predeterminada, como antes. Marque **Usar también para documentos nuevos** para que los documentos nuevos empiecen también con ellos. Las secciones con márgenes propios los conservan.

!!! tip "Desde un script"
    `Rte.SetPageSetup(...)` (lo graba la grabadora de macros), y `LeftMargin`, `RightMargin`, `TopMargin`, `BottomMargin`, `Landscape`, `PaperWidth`, `PaperHeight`: vea la [referencia de la API](../automation/api.md) (en inglés).

## Encabezados y pies de página { #headers-and-footers }

**Disposición > Encabezado y pie** establece el texto de la parte superior e inferior de cada página, con campos para el número de página, el número de páginas, la fecha y la hora, y el título o el nombre de archivo del documento: `{page}`, `{pages}`, `{date}`, `{time}`, `{title}`, `{file}`. Se muestran encima y debajo del texto en el editor (Vista > Encabezado y pie de página los muestra u oculta; haga clic en uno para cambiarlo), se imprimen y se guardan en el documento: como encabezados y pies de página de verdad en los archivos RTF, Word y OpenDocument, donde Word y LibreOffice mantienen al día los números de página.

**Varias líneas:** presione Entrar en una parte para empezar una línea nueva, por ejemplo para el nombre y la dirección de un despacho en el pie de página. Las líneas de la banda son las líneas de las partes una al lado de otra: la segunda línea de la parte izquierda junto a la segunda línea de la parte derecha. Word y LibreOffice reciben un párrafo por cada línea, y los encabezados y pies de varios párrafos se leen de vuelta desde ellos.

**Línea debajo del encabezado** y **Línea encima del pie de página** trazan una raya entre ellos y el texto, como el borde inferior de un encabezado en Word. Una línea en blanco al final de una parte del encabezado (Entrar después de su última línea) deja ese espacio debajo del encabezado. Para una línea en el propio documento, use **Insertar > Línea horizontal**. Los scripts usan `HeaderLine`, `FooterLine` e `InsertHorizontalLine()`.

**Posición** establece *Encabezado desde arriba* y *Pie de página desde abajo*, la distancia desde el borde de la página, como en Word (media pulgada de manera predeterminada). Un encabezado o pie de página que pasa de su margen empuja el texto hacia dentro, como hace Word. Los scripts usan `HeaderFromTop` y `FooterFromBottom` (en pulgadas).

**Palabras con su propia fuente:** seleccione palabras o líneas en una parte y use **N**, **K**, **S** (o Ctrl+B, Ctrl+I, Ctrl+U) o **Fuente...** en la fila *Palabras seleccionadas* de la parte superior del cuadro de diálogo: por ejemplo, el nombre de un despacho en negrita, con su dirección en la fuente del pie de página. Sin nada seleccionado, estos dan formato a todo el encabezado o pie de página. Las palabras conservan sus fuentes al imprimirse, en las hojas y en los archivos RTF y Word.

**Fuente...** debajo de las partes del encabezado y del pie de página elige su fuente: su nombre, tamaño, negrita, cursiva y subrayado, para las tres partes (y las de la primera página). Se imprime, se muestra en las hojas de Mostrar páginas y se conserva en los archivos RTF, Word y OpenDocument. **Simple** vuelve a la Segoe UI pequeña y gris de las versiones anteriores, y quita también las fuentes propias de las palabras. Los scripts usan `SetHeaderFont` y `SetFooterFont`.

**Primera página diferente** da a la página 1 un encabezado y un pie de página propios en todo: si tiene uno o no, su texto, su fuente y su línea. Márquela, y las pestañas del cuadro de diálogo pasan a ser *Otras páginas* y *Primera página*; lo que establezca en una pestaña no toca la otra. Para un membrete solo en la página 1, muestre el encabezado en la pestaña *Primera página* y deje desmarcada *Mostrar un encabezado* en *Otras páginas*. Una primera página sin pie de página, con números de página desde la página 2, funciona de la misma manera. El texto de la página 1 empieza debajo de su propio encabezado, el de las demás páginas debajo del suyo. El encabezado y el pie de página de la primera página se conservan en los archivos RTF, Word y OpenDocument, y una primera página preparada en Word o LibreOffice se lee de vuelta. Los scripts usan `DifferentFirstPage`, `SetFirstHeader`, `SetFirstFooter`, `ShowFirstHeader`, `ShowFirstFooter`, `SetFirstHeaderFont`, `SetFirstFooterFont`, `FirstHeaderLine` y `FirstFooterLine` (vea la [referencia de la API](../automation/api.md), en inglés).

## Diseño de impresión y zoom { #print-layout-and-zoom }

**Vista > Diseño de impresión** (activado al principio) corta las líneas donde se cortarán en la página impresa, como hacen Word y LibreOffice, y muestra la página sobre gris con sus márgenes como en el papel. Una línea punteada marcada "Página 2", "Página 3" ... muestra dónde empieza cada página impresa. Con Diseño de impresión desactivado, las líneas se ajustan al borde de la ventana.

El zoom, como en Word: **Ancho de página** (el predeterminado: la página llena el ancho de la ventana, y lo sigue cuando se cambia el tamaño de la ventana), **Una página** (una página entera cabe en la ventana), o el suyo propio con Acercar, Alejar, 100% y Ctrl+rueda del mouse. Se mantiene al abrir otro documento.

## Mostrar páginas { #show-pages }

**Vista > Mostrar páginas** (activado al principio) muestra cada página impresa como una hoja propia, como en Word: su encabezado y su pie de página en sus márgenes, con gris entre las hojas. Se edita como siempre: la escritura, el cursor y el final del documento se comportan como sin ella. Las hojas son solo de la vista: guardar, copiar, imprimir y los demás formatos ven el documento tal como es.

Las hojas siguen a la escritura con un segundo de retraso: se calculan de nuevo un segundo después de que usted deja de escribir, así que mientras escribe, el final de una página puede quedar hasta entonces a unas líneas de donde se imprimirá. Una página que termina dentro de una tabla se marca con una línea "Página N" en lugar de una hoja propia.

## Varias páginas, Vista previa de impresión, imprimir { #multiple-pages-print-preview-printing }

- **Vista > Zoom > Varias páginas** muestra las páginas impresas una al lado de otra. Son de solo lectura: haga doble clic en una página para editarla allí.
- **Archivo > Imprimir > Vista previa de impresión** muestra las páginas como se imprimirán, ajustadas a la pantalla, con botones de página y zoom.
- **Archivo > Imprimir** imprime; **Disposición > Exportar como PDF** crea un PDF con *Microsoft Print to PDF*. Como el de Word, el PDF tiene un marcador para cada título, bajo el del nivel superior, y cada uno abre su página en el título; se abre con sus marcadores a la vista, y las entradas de una tabla de contenido son vínculos a las páginas de sus títulos. (Scripts: `ExportPdf(path)`, o `ExportPdf(path, 0)` solo para las páginas.)
- Mientras se crea un PDF o se imprime un documento, la barra de estado muestra la página en que va (*Exportando como PDF: página 12 de 100*), y **Esc** lo detiene; una exportación detenida no deja ningún archivo. Un documento largo tarda un rato, unos 0,4 segundos por página, la mayor parte trabajo del propio Microsoft Print to PDF.

### Exportar como PDF, o imprimir en una impresora PDF { #export-as-pdf-or-print-to-a-pdf-printer }

Otras impresoras PDF, como CutePDF, PDF24 o PDFCreator, aparecen en **Archivo > Imprimir** como cualquier impresora una vez instaladas, y AhkWord imprime en ellas como imprime en papel. A menudo crean archivos más pequeños que Microsoft Print to PDF, y algunas son más rápidas. Lo que no pueden tener son **marcadores y vínculos**: Exportar como PDF los agrega al archivo que escribe Microsoft Print to PDF, pero las demás impresoras PDF preguntan dónde guardar en un cuadro de diálogo propio, así que AhkWord no puede encontrar el archivo para agregárselos.

- **Exportar como PDF:** marcadores para los títulos, y las entradas de la tabla de contenido como vínculos. Lo mejor para documentos largos que se leerán en pantalla.
- **Imprimir en una impresora PDF:** a menudo un archivo más pequeño, sin marcadores ni vínculos.

(Los instaladores de algunas impresoras PDF ofrecen otros programas por el camino: lea cada paso antes de hacer clic en Siguiente.)

La impresión sigue las reglas de saltos de línea y de página de los párrafos (conservar con el siguiente, conservar líneas juntas, control de líneas viudas y huérfanas, salto de página anterior), como hace Word. Conservar con el siguiente funciona también con las tablas: un título se queda con la tabla que le sigue, y las filas marcadas para conservar con el siguiente se quedan en una página, así que seleccionar todas las filas de una tabla menos la última la mantiene entera, y lo mismo en un bloque de firma evita que quede solo en una página.

Vea también [Números de línea](line-numbers.md), que también se imprimen.
