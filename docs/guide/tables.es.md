# Tablas

**Insertar > Tabla** inserta una tabla: elija las filas y las columnas, y si la primera fila es un encabezado en negrita. En un documento vacío queda un párrafo vacío encima de la tabla, listo para un título u otro texto; elimínelo si la tabla debe ir primero.

## Una tabla desde un archivo { #a-table-from-a-file }

**Desde archivo...** en el cuadro Insertar tabla crea la tabla a partir de un libro de Excel (.xlsx, .xlsm), un archivo CSV o TSV, o un archivo JSON. No hace falta tener Excel instalado, y el libro puede estar abierto en Excel mientras tanto (se lee lo último que se guardó). En un libro, elija la hoja y, si lo desea, un rango como `A1:D20`; sin rango se toman las celdas en uso, sin las filas y columnas vacías de alrededor. Los valores aparecen como los muestra Excel: formatos de número, moneda, porcentajes y fechas, en la configuración regional de su equipo; una fórmula da su último resultado. Las columnas se hacen tan anchas como lo pide su texto. Los archivos CSV pueden usar comas, punto y coma (como los guarda Excel donde la coma es el separador decimal) o tabulaciones, en UTF-8 o en la codificación propia de Windows.

Después, la tabla es una tabla como cualquier otra: no queda vinculada al archivo. Un script puede volver a llenarla desde el archivo conservando su formato (vea abajo).

Mientras el cursor está en una tabla, aparece la pestaña **Herramientas de tabla > Disposición**:

- **Insertar** filas arriba o debajo y columnas a la izquierda o a la derecha; **eliminar** filas, columnas o toda la tabla.
- **Propiedades** (también en el menú del botón derecho): los anchos de columna, la alineación de la tabla en la página, el margen de celda y el alto de fila para toda la tabla; los bordes, el sombreado y la alineación vertical para las celdas seleccionadas o para toda la tabla.
- **Bordes** (sin borde, todos, externos, internos ..., con grosor de línea y color) y **Sombreado** para las celdas seleccionadas.

**Tab** pasa a la celda siguiente y **Mayús+Tab** a la anterior; Tab en la última celda agrega una fila, como en Word. Esto funciona también en tablas de Word y de otros programas.

El formato de las tablas se conserva en los archivos RTF, Word, OpenDocument y HTML. Las tablas de Markdown no tienen anchos, bordes ni sombreado (vea [Vistas Markdown y HTML](source-views.md)). Dos límites: los bordes son líneas simples (los bordes dobles y punteados de Word se muestran como líneas simples del mismo grosor y color), y las celdas no se pueden combinar ni dividir, aunque las tablas de Word con celdas combinadas se abren tal como son.

!!! tip "Desde un script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: vea la [referencia de la API](../automation/api.md#tables) (en inglés).

!!! tip "Datos desde un script"
    `Rte.InsertTableFromFile("C:\Datos\clientes.xlsx", 1, "Clientes")` y `Rte.InsertTableFromData(csvOJson)` crean una tabla a partir de datos. `Rte.SetTableData(datos, "", 1, 2)` vuelve a llenar la primera tabla del documento conservando su fila de encabezado y su formato: se agregan o eliminan filas según haga falta, y los objetos JSON van en las columnas que encabezan sus nombres. `Rte.TableData(1, "json")` devuelve las celdas de una tabla como CSV, TSV o JSON, y `Rte.CellText(2, 3)` / `Rte.SetCellText(2, 3, "texto")` leen y cambian una celda. Vea la [referencia de la API](../automation/api.md#tables-from-data-and-data-from-tables) (en inglés).
