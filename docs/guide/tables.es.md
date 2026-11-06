# Tablas

**Insertar > Tabla** inserta una tabla: elija las filas y las columnas, y si la primera fila es un encabezado en negrita. En un documento vacío queda un párrafo vacío encima de la tabla, listo para un título u otro texto; elimínelo si la tabla debe ir primero.

Mientras el cursor está en una tabla, aparece la pestaña **Herramientas de tabla > Disposición**:

- **Insertar** filas arriba o debajo y columnas a la izquierda o a la derecha; **eliminar** filas, columnas o toda la tabla.
- **Propiedades** (también en el menú del botón derecho): los anchos de columna, la alineación de la tabla en la página, el margen de celda y el alto de fila para toda la tabla; los bordes, el sombreado y la alineación vertical para las celdas seleccionadas o para toda la tabla.
- **Bordes** (sin borde, todos, externos, internos ..., con grosor de línea y color) y **Sombreado** para las celdas seleccionadas.

**Tab** pasa a la celda siguiente y **Mayús+Tab** a la anterior; Tab en la última celda agrega una fila, como en Word. Esto funciona también en tablas de Word y de otros programas.

El formato de las tablas se conserva en los archivos RTF, Word, OpenDocument y HTML. Las tablas de Markdown no tienen anchos, bordes ni sombreado (vea [Vistas Markdown y HTML](source-views.md)). Dos límites: los bordes son líneas simples (los bordes dobles y punteados de Word se muestran como líneas simples del mismo grosor y color), y las celdas no se pueden combinar ni dividir, aunque las tablas de Word con celdas combinadas se abren tal como son.

!!! tip "Desde un script"
    `Rte.InsertTable(3, 4)`, `Rte.SetColumnWidths("2in, 3in")`, `Rte.SetCellBorders("outside", 1, "#C00000")`, `Rte.SetCellShading("LightYellow")`, `Rte.SetTableAlignment("center")`: vea la [referencia de la API](../automation/api.md#tables) (en inglés).
