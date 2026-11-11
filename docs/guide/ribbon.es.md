# La cinta de opciones y el teclado

Los comandos de AhkWord están en una cinta de opciones, la que trae Windows (como en Paint y WordPad): sigue la escala de la pantalla, y funciona con el teclado y con los lectores de pantalla.

| Pestaña | Grupos |
|---|---|
| **Archivo** (menú) | Nuevo, Nuevo desde plantilla, Página de inicio, Abrir, Abrir en una ventana nueva, Nueva ventana, Guardar, Guardar como, Cerrar (Ctrl+W), Imprimir (con Vista previa de impresión y Configurar página), Exportar (PDF, Word, OpenDocument), Salir, y los documentos recientes |
| **Inicio** | Portapapeles (con [Copiar formato](formatting.md#format-painter)); Fuente; Texto (mayúsculas y minúsculas, estilo de subrayado, el cuadro de diálogo Fuente); Párrafo (listas, niveles, Sin numeración, alineación, Interlineado, Opciones de lista, el cuadro de diálogo Párrafo); Estilos; Edición (buscar, reemplazar, seleccionar todo) |
| **Insertar** | Salto de página, Tabla de contenido, Tabla, Imagen, Miniatura vinculada, Línea horizontal, Vínculo, Referencia cruzada, Nota a pie, Actualizar (F9), Fecha y hora, Insertar desde plantilla, Rellenar campos, Fragmentos |
| **Disposición** | Configurar página, Encabezado y pie, Números de línea, Vista previa de impresión, Exportar como PDF |
| **Revisar** | Siguiente error ortográfico, Agregar al diccionario, Revisar la ortografía mientras se escribe; Idioma (English o Español, desde el próximo inicio) |
| **Vista** | Vista del documento (Texto enriquecido, Markdown, HTML, En paralelo, Ajustar líneas del código, Números de línea del código), Diseño de impresión, Mostrar páginas, Ajustar a la ventana, Encabezado y pie de página, Esquema, Regla, Zoom (Ancho de página, Una página, Varias páginas, Acercar, Alejar, 100%) |
| **Herramientas** | Macros, Grabar macro, Teclas rápidas de macros, Macros de eventos, Ejecutar script, Carpeta de macros, Paleta de comandos, API de automatización, Acerca de |
| **Herramientas de tabla > Disposición de tabla** | mientras el cursor está en una tabla: filas y columnas, Propiedades, Bordes, Sombreado, eliminar |

La **barra de herramientas de acceso rápido** tiene Guardar, Deshacer, Rehacer y la Paleta de comandos; su flecha agrega Nuevo, Abrir, Imprimir, Vista previa de impresión o Exportar como PDF.

El **menú del botón derecho** tiene, como en Word, las sugerencias ortográficas sobre una palabra mal escrita (cuando no hay nada más seleccionado), luego Cortar, Copiar y Pegar, **Editar nota a pie** sobre el número de una nota a pie, **Actualizar tabla** y **Actualizar solo los números de página** sobre una tabla de contenido, Fuente... y Párrafo..., y en una tabla los comandos de tabla. Un clic derecho fuera de la selección mueve primero el cursor allí; dentro de ella, la selección se mantiene.

## La paleta de comandos { #the-command-palette }

Presione ++ctrl+"Mayús"+p++ (o Herramientas > Paleta de comandos, o su lupa en la barra de herramientas de acceso rápido), escriba parte del nombre de un comando y presione ++"Entrar"++. Encuentra cualquier comando de la cinta, que se muestra con la pestaña y el grupo en que está y su método abreviado, así que también es una manera de aprender dónde están las cosas; y sus macros, los estilos, sus fragmentos y los documentos recientes. ++"Flecha arriba"++ y ++"Flecha abajo"++ eligen, ++esc++ la cierra.

- Cada palabra que escribe tiene que coincidir: `insertar tabla`, `salto pág`, `heading 2`. Cuenta primero el nombre del comando, luego su lugar en la cinta, luego su información sobre herramientas, así que `márgenes` encuentra Configurar página.
- Un comando que la cinta tiene atenuado (Deshacer sin nada que deshacer, los comandos de tabla fuera de una tabla) también aparece atenuado aquí, marcado *(ahora no)*.
- Los botones que se activan y desactivan dicen si están activados: *Diseño de impresión (activado)*.
- Sin nada escrito, los que usó por última vez van primero.
- Un comando se ejecuta exactamente como su botón: en la vista Markdown o HTML hace lo que hace allí el botón, y una macro que se está grabando lo graba.

++ctrl+k++ sigue siendo Insertar vínculo, como en Word.

## Teclado { #keyboard }

Presione ++alt++ para ver las KeyTips (las letras de las pestañas siguen en inglés): ++h++ Inicio, ++n++ Insertar, ++p++ Disposición, ++r++ Revisar, ++w++ Vista, ++l++ Herramientas, ++f++ Archivo. La información sobre herramientas de cada botón muestra su método abreviado. Los principales:

| Teclas | Hace |
|---|---|
| ++ctrl+b++, ++ctrl+i++, ++ctrl+u++ | negrita, cursiva, subrayado |
| ++ctrl+0++ | estilo Normal |
| ++ctrl+1++ a ++ctrl+6++ | Heading 1 a 6 |
| ++ctrl+7++ | estilo Code |
| ++ctrl+"Espacio"++ | Borrar formato |
| ++ctrl+"Mayús"+n++ | lista numerada |
| ++ctrl+"Mayús"+g++ | numeración legal |
| ++tab++, ++"Mayús"+tab++ | elementos de lista y párrafos numerados un nivel hacia dentro o hacia fuera; celda de tabla siguiente y anterior |
| ++ctrl+"Entrar"++ | salto de página |
| ++"Mayús"+"Entrar"++ | salto de línea |
| ++ctrl+k++ | vínculo |
| ++ctrl+"Mayús"+p++ | la paleta de comandos |
| ++ctrl+w++ | cerrar el documento |
| ++f7++ | siguiente error ortográfico |
| ++ctrl+alt+f++ | nota a pie |
| ++f9++ | actualizar las referencias cruzadas y los números de las notas a pie |
| ++ctrl++ + clic | seguir un vínculo o una referencia cruzada |
| ++ctrl+"Mayús"+r++ | grabar una macro |

El panel de esquema (Vista > Esquema) enumera los títulos: haga clic en uno para ir allí; haga clic derecho para subirlo o bajarlo de nivel.

La regla (Vista > Regla) muestra las sangrías y las tabulaciones del párrafo encima de la página; haga clic, arrastre o haga doble clic en ella para cambiarlas: vea [La regla](formatting.md#the-ruler).
