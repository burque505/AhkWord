# Configuración y carpetas

## Dónde guarda AhkWord sus archivos { #where-ahkword-keeps-its-files }

| Carpeta | Qué |
|---|---|
| `%APPDATA%\AhkWord` | `settings.json` (su configuración), `Templates` (sus propias plantillas, y sus `Parts`), `Recovery` (las copias de recuperación de los documentos sin guardar) |
| `%LOCALAPPDATA%\AhkWord` | los archivos de trabajo de los convertidores (el perfil de LibreOffice, el registro de los convertidores) |
| `rte_macros`, junto a AhkWord | las macros, y su propia configuración (`My signature.ini`) |

Escriba `%APPDATA%\AhkWord` en la barra de direcciones del Explorador para abrir la carpeta.

!!! note "Desde RTE Writer"
    AhkWord se llamaba RTE Writer hasta la iteración 38. Su primer inicio copia `%APPDATA%\RTE Writer` y `%LOCALAPPDATA%\RTE Writer` en las carpetas de AhkWord, con la configuración, las plantillas y las copias de recuperación, y deja las carpetas antiguas como estaban.

## Modo portátil { #portable-mode }

Ponga un archivo vacío llamado `portable` (o `portable.txt`) junto a `AhkWord.exe` o `AhkWord.ahk`, y AhkWord guarda su configuración, sus plantillas, sus copias de recuperación y el estado de los convertidores en una carpeta `data` junto a sí mismo, de modo que la carpeta se puede llevar a otro equipo. Los paquetes de .NET y el compilador de C# siguen en las cachés del propio equipo, así que un equipo que no ha ejecutado AhkWord antes los descarga una vez.

## Idioma { #language }

La interfaz de AhkWord (la cinta de opciones, los menús, la barra de estado, los cuadros de diálogo y los mensajes) está en inglés o en español: elíjalo en **Revisar > Idioma** (o con *Idioma: Español* en la paleta de comandos), y se muestra la próxima vez que se inicie AhkWord (los cuadros de diálogo del propio Windows, como Abrir, Guardar e Imprimir, están en el idioma de Windows). Es la opción `"Language": "es"` de `settings.json`, que también se puede poner a mano mientras AhkWord está cerrado. Los textos están en la carpeta `lang` junto a AhkWord: `en.json` enumera todos los textos en inglés con su clave, `es.json` los españoles; un texto que le falta a un idioma se muestra en inglés. Una carpeta junto a AhkWord con una carpeta `lang` propia (la de un complemento) puede tener allí también sus textos, como `<carpeta>\lang\es.json`: se leen después de los de AhkWord.

## Variables de entorno { #environment-variables }

| Variable | Hace |
|---|---|
| `AHKWORD_DATA` | una carpeta para todos los archivos de AhkWord (configuración, plantillas, copias de recuperación, el estado de los convertidores), como la carpeta `data` del modo portátil |
| `AHKWORD_LANG` | el idioma de la interfaz por encima de la opción `Language` (`es`, o `en`), para probar uno sin cambiar la configuración |
| `AHKWORD_SETTINGS` | el archivo de configuración que se usa en lugar de `settings.json` |
| `AHKWORD_CONVERTERS` | los convertidores que se usan para los archivos de Word y OpenDocument, solo esos y en ese orden: `word`, `libreoffice`, `openxml`, `pandoc` (por ejemplo `libreoffice,openxml`) |

Los nombres antiguos `RTE_WRITER_SETTINGS` y `RTE_WRITER_CONVERTERS` siguen funcionando.

## Varias ventanas { #several-windows }

Las ventanas de AhkWord comparten un archivo de configuración. Cada una guarda solo lo que cambió, así que cerrar una no deshace los cambios de otra: los documentos recientes, las teclas rápidas de macros y los valores recordados se combinan elemento por elemento.
