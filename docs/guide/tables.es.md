# Tablas

**Insertar > Tabla** inserta una tabla: elija las filas y las columnas, y si la primera fila es un encabezado en negrita. En un documento vacío queda un párrafo vacío encima de la tabla, listo para un título u otro texto; elimínelo si la tabla debe ir primero.

## Una tabla desde un archivo { #a-table-from-a-file }

**Desde archivo...** en el cuadro Insertar tabla crea la tabla a partir de un libro de Excel (.xlsx, .xlsm), un archivo CSV o TSV, o un archivo JSON. No hace falta tener Excel instalado, y el libro puede estar abierto en Excel mientras tanto (se lee lo último que se guardó). En un libro, elija la hoja y, si lo desea, un rango como `A1:D20`; sin rango se toman las celdas en uso, sin las filas y columnas vacías de alrededor. Los valores aparecen como los muestra Excel: formatos de número, moneda, porcentajes y fechas, en la configuración regional de su equipo; una fórmula da su último resultado. Las columnas se hacen tan anchas como lo pide su texto. Los archivos CSV pueden usar comas, punto y coma (como los guarda Excel donde la coma es el separador decimal) o tabulaciones, en UTF-8 o en la codificación propia de Windows.

Después, la tabla es una tabla como cualquier otra: no queda vinculada al archivo. Un script puede volver a llenarla desde el archivo conservando su formato (vea abajo).

Mientras el cursor está en una tabla, aparece la pestaña **Herramientas de tabla > Disposición de tabla**:

- **Insertar** filas arriba o debajo y columnas a la izquierda o a la derecha; **eliminar** filas, columnas o toda la tabla.
- **Combinar celdas** y **Dividir celdas**, y **Eliminar celdas** (vea abajo).
- **Propiedades** (también en el menú del botón derecho): los anchos de columna, la alineación de la tabla en la página, el margen de celda y el alto de fila para toda la tabla; los bordes, el sombreado y la alineación vertical para las celdas seleccionadas o para toda la tabla.
- **Bordes** (sin borde, todos, externos, internos ..., con grosor de línea y color) y **Sombreado** para las celdas seleccionadas.

**Tab** pasa a la celda siguiente y **Mayús+Tab** a la anterior; Tab en la última celda agrega una fila, como en Word. Esto funciona también en tablas de Word y de otros programas.

El formato de las tablas se conserva en los archivos RTF, Word, OpenDocument y HTML. Las tablas de Markdown no tienen anchos, bordes ni sombreado (vea [Vistas Markdown y HTML](source-views.md)). Un límite: los bordes son líneas simples (los bordes dobles y punteados de Word se muestran como líneas simples del mismo grosor y color). Vea también las [limitaciones conocidas](../reference/limitations.md).

## Combinar, dividir y eliminar celdas { #merging-splitting-and-deleting-cells }

Para trabajar con varias celdas, arrastre el mouse sobre ellas: desde una celda de una fila hasta una celda de otra, AhkWord selecciona las celdas que hay entre ellas (sombreadas como una selección), no las filas enteras. Dentro de una fila, seleccione desde el texto de la primera celda hasta el de la última. Con el teclado, **Mayús+Arriba** y **Mayús+Abajo** en una celda seleccionan como bloque las celdas de arriba o de abajo, y luego **Mayús+Izquierda** y **Mayús+Derecha** agregan o quitan una columna.

- **Combinar celdas** convierte las celdas seleccionadas en una sola. Su texto se conserva, el de cada celda en párrafos propios, como en Word. Las celdas combinadas a lo ancho forman una celda más ancha; las combinadas hacia abajo se ven como una celda alta.
- **Dividir celdas...** pregunta cuántas columnas y filas. Una celda se divide en columnas del mismo ancho. Una celda combinada hacia abajo vuelve a dividirse en sus filas (o en un número de filas que las divida); una celda de una sola fila dividida en filas recibe filas nuevas, y las demás celdas de la fila se combinan hacia abajo sobre ellas para quedar enteras. Si hay varias celdas seleccionadas, primero se combinan.
- **Eliminar celdas...** pregunta qué pasa con las celdas de alrededor: las que les siguen en sus filas se desplazan a la izquierda, las de abajo se desplazan hacia arriba (su texto; abajo quedan celdas vacías), o se eliminan sus filas o columnas enteras. Con celdas seleccionadas, **Eliminar fila** y **Eliminar columna** eliminan todas sus filas o columnas.
- **Supr** (la tecla) sobre celdas seleccionadas las vacía.
- **Copiar** (Ctrl+C) sobre celdas seleccionadas las copia como una tabla propia (péguela en AhkWord o en Word) y como texto separado por tabulaciones (Excel lo pone en celdas). **Cortar** las copia y las vacía.

Las columnas se encuentran por su posición, no contando celdas: eliminar o insertar una columna que abarca una celda combinada hace esa celda más estrecha o más ancha en lugar de eliminarla. Cada una de estas acciones se deshace con un solo **Deshacer**.

!!! tip "Desde un script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: vea la [referencia de la API](../automation/api.md#tables) (en inglés).
    Celdas: `Rte.SelectCells(3, 2, 4, 2)` (de la fila 3, celda 2 a la fila 4, celda 2), y luego `Rte.MergeCells()`, `Rte.SplitCells(2, 1)`, `Rte.DeleteCells("up")` o `Rte.ClearCells()`.

!!! tip "Datos desde un script"
    `Rte.InsertTableFromFile("C:\Datos\clientes.xlsx", 1, "Clientes")` y `Rte.InsertTableFromData(csvOJson)` crean una tabla a partir de datos. `Rte.SetTableData(datos, "", 1, 2)` vuelve a llenar la primera tabla del documento conservando su fila de encabezado y su formato: se agregan o eliminan filas según haga falta, y los objetos JSON van en las columnas que encabezan sus nombres. `Rte.TableData(1, "json")` devuelve las celdas de una tabla como CSV, TSV o JSON, y `Rte.CellText(2, 3)` / `Rte.SetCellText(2, 3, "texto")` leen y cambian una celda. Vea la [referencia de la API](../automation/api.md#tables-from-data-and-data-from-tables) (en inglés).

!!! tip "De vuelta al archivo de origen"
    `Rte.SaveTableData("C:\Datos\clientes.xlsx", 1, "Clientes")` escribe la primera tabla del documento de vuelta en el libro: solo las celdas cuyo valor cambió (las fórmulas y formatos de las demás no se tocan), y se vacían las celdas que la tabla ya no alcanza; el libro debe estar cerrado en Excel. Un archivo .csv, .tsv o .json se escribe entero, como era (su delimitador; los objetos JSON siguen siendo objetos). Las macros **Table to source file** y **Table from source file** (Herramientas > Macros) lo hacen con la tabla donde está el cursor: la primera vez piden el archivo y lo recuerdan para esa tabla de ese documento (mantenga Mayús para elegir otro).
