# Scripts, eventos y macros

Cualquier script de AutoHotkey v2 puede controlar un AhkWord en ejecución, mediante `RteClient2.ahk` (junto a `AhkWord.ahk`):

```autohotkey
#Requires AutoHotkey v2.0
#Include "RteClient2.ahk"

Rte.InsertText("Hola desde AutoHotkey")
Rte.InsertTable(3, 4)
MsgBox Rte.WordCount() " palabras"
Rte.On("Saved", (e) => TrayTip(Rte.FileName, "Guardado"))
```

Todo lo que AhkWord puede hacer está en la [referencia de la API](api.md) (en inglés); en AhkWord, **Herramientas > API de automatización** también lo enumera.

## Llamadas { #calls }

- `Rte.Nombre(argumentos)` llama a un método y espera su resultado; `Rte.Nombre` lee una propiedad y `Rte.Nombre := valor` la establece.
- Los argumentos son texto, números, 1 o 0 para sí o no, enumeraciones por su nombre (`Heading1`, `LowerAlpha`) y colores como `#RRGGBB` o un nombre.
- `Rte.Post(Nombre, argumentos...)` pone una llamada en cola y vuelve enseguida; sus errores van a `Rte.LastAsyncError` y al evento `AsyncError`.
- Una llamada que falla lanza un error de AutoHotkey con el mensaje de AhkWord.

Con varias ventanas de AhkWord abiertas, un script habla con la que se usó por última vez. Una macro iniciada desde una ventana de AhkWord habla con esa ventana.

## Eventos { #events }

`Rte.On(Evento, Callback)` llama a `Callback(e)` cuando ocurre algo en el editor, y `Rte.Off(Evento [, Callback])` lo detiene. Todo `e` tiene `e.Name`; algunos eventos dicen más:

| Evento | Cuándo | También en `e` |
|---|---|---|
| `TextChanged` | el texto cambió | `e.TextLength` |
| `SelectionChanged` | la selección o el cursor se movieron | `e.Start`, `e.Length` |
| `ModifiedChanged` | el documento pasó a estar modificado, o se guardó | `e.Modified` |
| `Saved` | el documento se guardó | |
| `Opened` | se abrió un documento, o se empezó uno nuevo | |
| `Closed` | el editor se cerró | |
| `AsyncError` | falló una llamada puesta en cola con `Rte.Post` | |

## Macros { #macros }

- **Herramientas > Macros** ejecuta los scripts de la carpeta `rte_macros` (junto a AhkWord); **Carpeta de macros** la abre.
- **Teclas rápidas de macros** da una tecla rápida a cualquier macro.
- **Grabar macro** (Ctrl+Mayús+R) graba lo que escribe y los comandos de la cinta como un script de macro nuevo: cada una de sus líneas es una llamada a AhkWord, así que una macro grabada también es una buena manera de aprender las llamadas.
- **Ejecutar script** ejecuta cualquier script de AutoHotkey, conectado a esta ventana.

### Las macros de ejemplo { #the-example-macros }

Las macros de ejemplo tienen nombres en inglés (son los nombres de sus archivos):

- **Clean up pasted text** (limpiar texto pegado): une las líneas cortadas en medio de las frases, ordena los espacios y convierte las comillas y los guiones en tipográficos.
- **My signature** (mi firma): inserta su nombre, cargo, teléfono y correo electrónico (se piden una vez y luego se recuerdan).
- **Collect from anywhere** (recoger desde cualquier sitio): seleccione texto en cualquier programa y presione Win+Mayús+C para agregarlo, con su procedencia y la hora, al final del documento.
- **PDF on save** (PDF al guardar): cada vez que se guarda, se crea también un PDF junto al documento. Es una macro del evento Saved: ejecútela desde Herramientas > Macros, o con una tecla rápida de Teclas rápidas de macros, para activarla o desactivarla (pregunta antes de ocupar el lugar de otra macro de Saved).
- **Folder to PDF** (carpeta a PDF): crea un PDF de cada documento de una carpeta.
- **Arrange windows** (organizar ventanas): pone las ventanas abiertas de AhkWord una al lado de otra en la pantalla, cada página ajustada al ancho de su ventana.
- **Open AutoHotkey Dash:** el Dash de AutoHotkey, para su ayuda, Window Spy y **Editor settings**, el programa con que se editan las macros (vea [Primeros pasos](../getting-started.md#which-program-opens-a-macro-to-edit-it)). Necesita AutoHotkey instalado.
- Un guardado automático, un seguimiento de objetivo de palabras, un monitor de eventos, estadísticas del documento y Title Case para la selección.

## Macros de eventos { #event-hooks }

**Herramientas > Macros de eventos** da a un evento una macro que AhkWord ejecuta solo cuando ocurre. A diferencia de un script con `Rte.On`, nada tiene que seguir ejecutándose: la macro se inicia, hace su trabajo y termina.

| Evento | Cuándo | `A_Args[2]` |
|---|---|---|
| Started | la ventana de AhkWord se ha abierto, con su primer documento | |
| Opened | se abrió un documento | su archivo |
| NewDocument | Archivo > Nuevo | |
| BeforeSave | un documento está a punto de guardarse: AhkWord espera a la macro | el archivo |
| Saved | se guardó un documento | su archivo |
| StyleChanged | se cambió, creó, eliminó o copió un estilo, o cambió el conjunto de estilos | el estilo |
| Printed | el documento se imprimió o se exportó como PDF | la impresora, o el PDF |
| Closing | la ventana está a punto de cerrarse: AhkWord espera a la macro | el archivo |

La macro recibe el nombre del evento en `A_Args[1]` (y en la variable de entorno `AHKWORD_EVENT`), y aquello de lo que trataba en `A_Args[2]` (`AHKWORD_EVENT_DETAIL`). Trabaja con la ventana en que ocurrió el evento mediante `Rte`, como cualquier macro. **BeforeSave** y **Closing** esperan a su macro, hasta 15 segundos, para que pueda cambiar el documento antes de que se guarde, o guardarlo antes de que la ventana se cierre; los demás eventos no esperan. Mientras se ejecuta la macro de un evento, ese evento no la vuelve a iniciar, así que una macro de Saved que guarda no se ejecuta en bucle.

Desmarque **Activa** junto a una macro para mantenerla elegida sin que se ejecute: un PDF en cada guardado solo algunas veces, por ejemplo; márquela de nuevo para recuperarla. En el cuadro de diálogo, **Nueva...** crea una macro para el evento (`On Saved.ahk`, por ejemplo) a partir de un script inicial con un ejemplo, y la abre para editarla; **Ejecutar ahora** prueba la macro como la ejecutaría el evento. Los scripts usan `SetEventHook("Saved", "On Saved.ahk")`, `SetEventHookOn("Saved", false)` (para que una tecla rápida pueda activar o desactivar una macro de evento), `EventHookOn`, `EventHook`, `EventHookList` y `HookEventNames`.

Un ejemplo, una macro de Saved que guarda una copia fechada de cada guardado en una carpeta `Backups` junto al documento:

```autohotkey
#Requires AutoHotkey v2.0
#Include "..\RteClient2.ahk"
SavedFile := A_Args.Length >= 2 ? A_Args[2] : ""
if SavedFile != "" {
    SplitPath(SavedFile, , &Dir, &Ext, &Stem)
    DirCreate(Dir "\Backups")
    FileCopy(SavedFile, Dir "\Backups\" Stem " " FormatTime(, "yyyy-MM-dd HHmm") "." Ext, true)
}
```

## Teclas rápidas en AhkWord.ahk { #hotkeys-in-ahkwordahk }

El propio `AhkWord.ahk` define dos teclas rápidas, activas solo en las ventanas de AhkWord: Ctrl+Alt+D inserta la fecha, y Ctrl+Alt+W muestra el recuento de palabras en la barra de estado. Agregue allí las suyas; llaman directamente al editor, en el mismo proceso.
