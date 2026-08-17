# Cómo instalar y usar la macro "Generar Cartas Fail"

Esta solución tiene 2 piezas:

1. **`Template_Cartas_Fail.docx`** — tu misma carta, pero con los campos amarillos reemplazados por marcadores tipo `«AGENCIA_NOMBRE»`. La macro busca estos marcadores y los reemplaza con los datos reales de cada locación.
2. **`Modulo_GenerarCartasFail.bas`** — el código VBA que filtra los "Fail" y genera las cartas.

---

## Paso 1 — Guardar la plantilla de Word

Guarda `Template_Cartas_Fail.docx` en una carpeta fija de tu computadora (por ejemplo `C:\WU\Plantillas\Template_Cartas_Fail.docx`). **No la abras y la edites** — es la plantilla base que la macro usa cada vez.

## Paso 2 — Crear una carpeta de salida

Crea una carpeta donde se guardarán las cartas generadas, por ejemplo `C:\WU\Cartas_Generadas\`. Si no existe, la macro la crea sola.

## Paso 3 — Habilitar el menú Programador (si no lo tienes)

En Excel: `Archivo` → `Opciones` → `Personalizar cinta de opciones` → marca la casilla **Programador** → Aceptar.

## Paso 4 — Abrir el editor de VBA e importar el módulo

1. Abre tu archivo Excel real (con todos los datos).
2. Presiona `Alt + F11` para abrir el editor de Visual Basic.
3. Clic derecho sobre el nombre de tu libro (en el panel izquierdo) → `Insertar` → `Módulo`.
4. Abre `Modulo_GenerarCartasFail.bas` con el Bloc de notas, copia todo el contenido, y pégalo dentro del módulo nuevo en el editor de VBA (reemplaza lo que haya).

   *(Alternativa: en el editor de VBA puedes usar `Archivo` → `Importar archivo...` y seleccionar directamente el `.bas`.)*

## Paso 5 — Editar las líneas de configuración

Al inicio del módulo vas a ver esto:

```vba
Public Const NOMBRE_HOJA_DATOS As String = "Sheet1"
Public Const NOMBRE_HOJA_FAILS As String = "Fails"
Public Const RUTA_PLANTILLA_WORD As String = "C:\Ruta\A\Template_Cartas_Fail.docx"
Public Const CARPETA_SALIDA As String = "C:\Ruta\A\Cartas_Generadas\"
```

Cambia:
- `NOMBRE_HOJA_DATOS` si la hoja con los datos (dentro del archivo que vas a elegir cada vez) no se llama "Sheet1".
- `RUTA_PLANTILLA_WORD` a la ruta real donde guardaste el `.docx` del Paso 1.
- `CARPETA_SALIDA` a la ruta real de la carpeta del Paso 2.

**No hace falta configurar la ruta del archivo de datos** — al ejecutar la macro, se abre un explorador de Windows para que elijas manualmente el Excel con los resultados de esa corrida (puede cambiar de nombre cada vez, por ejemplo un reporte mensual distinto). El archivo se abre solo en modo lectura, se copia lo necesario al tab **"Fails"** (que se crea en el archivo donde vive el macro) y luego se cierra sin guardarle ningún cambio.

**Puedes pegar este módulo en un archivo donde ya tengas otra macro** — solo revisa que no haya un Sub o Function con el mismo nombre que los de este módulo (`GenerarCartasFail`, `GenerarCartaWord`, `ColumnaPorEncabezado`, `TraducirEscenario`, `ConvertirAFecha`, `FormatearFechaLarga`, `FormatearFechaCorta`, `LimpiarNombreArchivo`). Si hay algún choque de nombres, Excel te lo va a marcar como error al compilar — solo renombra uno de los dos.

## Paso 6 — Habilitar macros y permitir el acceso al modelo de objetos de VBA

Como la macro controla Word desde Excel, a veces Windows pide un permiso adicional:
`Archivo` → `Opciones` → `Centro de confianza` → `Configuración del Centro de confianza` → `Configuración de macros` → marca **"Confiar en el acceso al modelo de objetos de proyectos de VBA"**.

## Paso 7 — Guardar como .xlsm

`Archivo` → `Guardar como` → tipo de archivo: **Libro de Excel habilitado para macros (*.xlsm)**.

## Paso 8 — Ejecutar

1. Vuelve a Excel (`Alt + F11` para regresar, o cierra el editor).
2. `Alt + F8` → selecciona `GenerarCartasFail` → `Ejecutar`.
3. Se abrirá un explorador de archivos — selecciona el Excel con los resultados de esa corrida (el que tiene las columnas `Location ID`, `Did this location pass or fail?`, etc.).
4. La macro crea (o actualiza) el tab **"Fails"** en el archivo donde vive el macro, con las locaciones que fallaron, y en la carpeta de salida aparecerá un `.docx` por cada una, nombrado `IDLOCACION_NombreLocacion.docx`.

Al terminar te mostrará un resumen: cuántos "Fail" encontró, cuántas cartas generó, y si hubo errores.

---

## Cosas para revisar / ajustar a tu gusto

- **Traducción del escenario:** dentro de la función `TraducirEscenario` hay una lista inglés → español. Si aparece un escenario que no está en la lista, la carta mostrará el texto en inglés seguido de `[TRADUCIR]` para que lo notes y lo agregues a la lista.
- **Dirección de la carta:** actualmente uso `Ciudad, País` (sin el campo "State") para que coincida con el ejemplo que trae la plantilla. Si prefieres incluir el estado/provincia, es un solo cambio de línea en el módulo (buscar `AGENCIA_CIUDAD_PAIS`).
- **Narrativa:** se usa la columna en **idioma local** (como pediste). Los saltos de línea del texto original se convierten automáticamente a saltos de línea dentro del mismo párrafo de la carta.
- **Fecha:** tanto la fecha del encabezado como "Fecha de envío del informe" usan la fecha del día en que ejecutas la macro (como pediste). "Fecha de la visita" toma la columna `Date shop performed:`.
- **Resaltado amarillo:** se quita automáticamente en la carta final generada (para que no se vea marcada en amarillo cuando se envíe). Si prefieres conservarlo, borra esta línea del código: `doc.Content.HighlightColorIndex = 0`.

## Si algo falla

- **"No se encontró la plantilla de Word"** → revisa que `RUTA_PLANTILLA_WORD` apunte exactamente al archivo, incluyendo la extensión `.docx`.
- **"No se encontraron una o más columnas esperadas"** → algún encabezado de tu hoja de datos no coincide exactamente (revisa espacios extra, mayúsculas, etc.) con las constantes `COL_...` al inicio del módulo.
- **Error de automatización de Word** → asegúrate de tener Microsoft Word instalado en la misma computadora donde corres la macro.
