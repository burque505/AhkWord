# Formato y párrafos

## Caracteres { #characters }

**Inicio > Fuente** tiene la fuente y el tamaño, la negrita, la cursiva, el subrayado, el color y el resaltado; **Inicio > Texto** tiene el cambio de mayúsculas y minúsculas, los estilos de subrayado y el cuadro de diálogo Fuente. Se puede usar cualquier tamaño de fuente.

Sin nada seleccionado, la fuente, el tamaño o el color que elija se aplican a lo que escriba a continuación, y en el acto a una línea vacía, como en Word.

**Borrar formato** (Ctrl+Espacio) devuelve el texto seleccionado al estilo de su párrafo; sin nada seleccionado, el párrafo entero. Vea [Estilos](styles.md).

## Copiar formato { #format-painter }

**Inicio > Portapapeles > Copiar formato** (la brocha) copia el formato de un lugar a otro, como el de Word:

1. Ponga el cursor en el texto cuyo aspecto quiere, o seleccione parte de él.
2. Haga clic en **Copiar formato**. El puntero se convierte en una brocha.
3. Haga clic en otro párrafo, o seleccione otro texto: toma el mismo formato.

Con el cursor en un párrafo (sin nada seleccionado), copia el formato del párrafo (su estilo, alineación, espaciado, sangrías y numeración, incluida la numeración legal) y el formato de caracteres en el cursor; luego un clic aplica el formato al párrafo donde hace clic y a la palabra bajo el clic. Con texto seleccionado, copia el formato que tiene la mayor parte de la selección, y el formato del párrafo solo cuando la selección incluye la marca del párrafo (su final), como en Word.

Un clic en Copiar formato lo aplica una vez. **Haga doble clic** en él para seguir aplicándolo, un lugar tras otro, hasta que presione ++esc++ o vuelva a hacer clic en Copiar formato. Cada aplicación es un paso de Deshacer.

!!! tip "Desde un script"
    `Rte.CopyFormatting()` y `Rte.PasteFormatting()`: vea la [referencia de la API](../automation/api.md) (en inglés).

## Alineación, interlineado y espaciado { #alignment-line-spacing-and-spacing }

**Inicio > Párrafo** tiene los botones de alineación e **Interlineado**: de 1,0 a 3,0 líneas, el espacio antes y después del párrafo, y *Opciones de interlineado* para "mínimo" o "exacto" de tantos puntos.

## El cuadro de diálogo Párrafo { #the-paragraph-dialog }

El botón **Configuración de párrafo** de Inicio > Párrafo abre un cuadro de diálogo Párrafo como el de Word:

- **Sangría y espaciado:** la alineación; las sangrías izquierda y derecha; una sangría de primera línea o francesa; el espacio anterior y posterior; el interlineado.
- **Líneas y saltos de página:** el control de líneas viudas y huérfanas, conservar con el siguiente, conservar líneas juntas, salto de página anterior y suprimir números de línea.

Una vista previa muestra el resultado. Cuando los párrafos seleccionados difieren en un valor, este queda vacío, y cada párrafo conserva el suyo a menos que lo cambie. Las sangrías están en pulgadas y el espaciado en puntos, pero puede escribir cualquier unidad: `1.27 cm`, `36pt`, `0.5in`.

- **Tabulaciones...** fija las tabulaciones de los párrafos (posición, alineación, relleno) y las tabulaciones predeterminadas del documento.
- **Establecer como predeterminado...** da estos valores al estilo del párrafo, solo para este documento o también para los documentos nuevos.

## La regla { #the-ruler }

**Vista > Regla** (desactivada al principio) muestra una regla encima de la página, como la regla horizontal de Word. Mide en pulgadas, o en centímetros donde Windows mide en ellos, desde el borde izquierdo del texto. En Diseño de impresión los márgenes son grises y el ancho del texto, blanco.

En ella están las sangrías y las tabulaciones del párrafo donde está el cursor (con varios párrafos seleccionados, las del primero):

- la **sangría de primera línea**: el triángulo de arriba, a la izquierda;
- la **sangría francesa**: el triángulo de abajo, con el cuadro de la **sangría izquierda** debajo;
- la **sangría derecha**: el triángulo de la derecha;
- las **tabulaciones** del párrafo y, tras ellas, las tabulaciones predeterminadas del documento como pequeñas marcas grises.

En una tabla la regla muestra solo la escala, y en Varias páginas se oculta. Sigue el zoom, y la página cuando se desplaza hacia los lados.

Lo que puede hacer en ella (cada cambio se aplica a todos los párrafos seleccionados y es un paso de Deshacer):

| Haga esto | Para |
|---|---|
| clic en la parte blanca | agregar allí una tabulación izquierda |
| arrastrar una tabulación | moverla (conserva su alineación y su relleno) |
| arrastrar una tabulación hacia abajo o hacia arriba, fuera de la regla | quitarla |
| presionar en la parte blanca y arrastrar | agregar una tabulación donde suelte |
| arrastrar el triángulo de arriba | fijar la sangría de primera línea |
| arrastrar el triángulo de abajo | fijar la sangría francesa (la primera línea no se mueve) |
| arrastrar el cuadro de debajo | mover juntas la sangría izquierda y la primera línea |
| arrastrar el triángulo de la derecha | fijar la sangría derecha |
| doble clic en una tabulación (o en la parte blanca) | abrir **Tabulaciones** |
| doble clic en un marcador de sangría | abrir **Párrafo** |
| doble clic en los márgenes grises | abrir **Configurar página** |

Los clics y los arrastres se ajustan a una cuadrícula de 1/16 de pulgada (0,25 cm), como en Word; mantenga presionada ++alt++ mientras arrastra para colocar un marcador libremente. Mientras arrastra, la regla muestra el nuevo lugar del marcador, una línea de puntos baja por la página bajo él y la barra de estado muestra el valor. Las sangrías se quedan dentro del papel y dejan al menos un cuarto de pulgada de texto.

El cuadro de diálogo Párrafo y su cuadro Tabulaciones hacen lo mismo con números.

La impresión, la Vista previa de impresión, el PDF y Varias páginas siguen las reglas de saltos de línea y de página, como Word, y Word y LibreOffice las reciben en sus archivos.

!!! tip "Desde un script"
    `Rte.ParagraphFormat()`, `Rte.SetParagraphFormat("LeftIndent=0.5in; KeepTogether=1")`, `Rte.Tabs`, `Rte.SetTabs(...)`, `Rte.AddTab`, `Rte.MoveTab`, `Rte.ClearTab`, `Rte.DefaultTabStop` y `Rte.RulerVisible`: vea la [referencia de la API](../automation/api.md) (en inglés).
