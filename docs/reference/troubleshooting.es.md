# Solución de problemas

## AhkWord no se inicia { #ahkword-doesnt-start }

**No pasa nada durante unos segundos en el primer inicio.** Es lo esperado, sin que se muestre nada hasta que se abre la ventana: el primer inicio de una versión descargada es más lento mientras Windows comprueba sus archivos nuevos. Ejecutado desde el repositorio, el primer inicio compila AhkWord (los inicios siguientes lo cargan ya compilado, hasta que cambia un archivo de código) y descarga sus componentes la primera vez, lo que puede tardar uno o dos minutos.

**"Windows protegió su PC"** (SmartScreen) la primera vez: AhkWord no está firmado. Haga clic en **Más información** y luego en **Ejecutar de todas formas**; Windows no vuelve a preguntar por esa copia.

**El Control inteligente de aplicaciones (Smart App Control) bloquea AhkWord.** El Control inteligente de aplicaciones, que algunos equipos nuevos con Windows 11 tienen activado, bloquea los programas que no están firmados por un editor que conoce, sin la opción de ejecutarlos de todas formas. `AhkWord.exe` es AutoHotkey con el nombre de AhkWord, y AutoHotkey no está firmado, así que queda bloqueado, como lo estaría cualquier script de AutoHotkey. Su configuración está en **Seguridad de Windows**, en el control de aplicaciones y del navegador. Desactivarlo permite que AhkWord se ejecute, pero piénselo antes: en muchas versiones de Windows 11 no se puede volver a activar después sin restablecer Windows. Un AhkWord firmado, que no necesita cambiar esa opción, está en la [hoja de ruta](https://github.com/burque505/AhkWord/blob/main/ROADMAP.md) (en inglés).

**Un error sobre .NET.** AhkWord necesita el .NET 10 Desktop Runtime; un .NET anterior no basta. Instálelo con `winget install Microsoft.DotNet.DesktopRuntime.10` e inicie AhkWord de nuevo.

**Un error con `NU1100` o `NU1301`.** No se pudieron descargar los componentes desde nuget.org. Compruebe que el equipo puede llegar a nuget.org (un proxy o un firewall pueden impedirlo), y los orígenes de paquetes que enumera el mensaje.

**Un mensaje de que AhkWord se detuvo por un error.** Muestra el texto del error. Por favor, notifíquelo, con ese texto, en la [página de incidencias](https://github.com/burque505/AhkWord/issues) (en inglés).

## Archivos de Word y OpenDocument { #word-and-opendocument-files }

**Un archivo de Word se abre o se guarda despacio, o con otro convertidor del que esperaba.** Word se usa solo mientras no está abierto: con Word abierto, AhkWord usa LibreOffice o el convertidor integrado. Cierre Word para que la conversión la haga el propio Word.

**Un convertidor falla.** AhkWord prueba el siguiente, y dice en la barra de estado cuál usó. El registro de los convertidores está en `%LOCALAPPDATA%\AhkWord\converters.log`.

## Listas y numeración { #lists-and-numbering }

**Los elementos de una lista anidada muestran todos la numeración del primer nivel (A. A. A.).** En Niveles de lista, marque *Tab y Mayús+Tab dan a un elemento el estilo de numeración de su nuevo nivel*.

**Un párrafo de un estilo numerado no tiene número.** Su numeración está desactivada: Tab al principio lo convierte en un elemento de subnivel, o el botón Numeración lo vuelve a numerar. Vea [Sin numeración](../guide/lists.md#numbering-off).

## Recuperar un documento { #recovering-a-document }

Si AhkWord terminó sin cerrarse, su siguiente inicio ofrece los documentos que estaba editando. Sus copias están en `%APPDATA%\AhkWord\Recovery` hasta que usted los recupere o los descarte.

## La barra de tareas muestra el icono de AutoHotkey { #the-taskbar-shows-autohotkeys-icon }

Cierre todas las ventanas de AhkWord e inícielo de nuevo: Windows a veces conserva el icono antiguo de un botón de la barra de tareas hasta que el programa se inicia desde cero.
