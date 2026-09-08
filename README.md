# Generador de Cartas de Mystery Shopping (Fails)

Macro de VBA para Excel que arma una lista de resultados "Fail" de mystery shopping y genera una carta de Word por cada uno, usando plantillas distintas por país/región.

## Archivos
- `GenerarCartasFail.bas` — módulo de VBA con todo el código. Impórtalo en tu libro de Excel (macro-habilitado, `.xlsm`).

## Cómo importarlo a tu Excel
1. Abre tu libro `.xlsm` (o crea uno nuevo y guárdalo como macro-habilitado).
2. Presiona `Alt + F11` para abrir el editor de VBA.
3. Clic derecho sobre el proyecto → `Importar archivo...` → selecciona `GenerarCartasFail.bas`.
4. Guarda el libro.

## Configuración inicial
Al inicio del módulo hay algunas constantes que debes ajustar según tu computadora:

| Constante | Qué es |
|---|---|
| `NOMBRE_HOJA_DATOS` | Nombre del tab con los datos dentro del Excel que vas a seleccionar cada vez (por defecto `Sheet1`). |
| `NOMBRE_HOJA_FAILS` | Nombre del tab que el macro crea en este libro (por defecto `Fails`). |
| `CARPETA_SALIDA` | Carpeta donde se guardan las cartas generadas. Se crea automáticamente si no existe. |
| `COL_...` | Nombres exactos de los encabezados (fila 1) que el macro busca en el archivo de datos. Deben coincidir letra por letra. |

Ya **no** hace falta configurar una ruta fija de plantilla de Word: el macro te la pide al momento de generar las cartas (ver más abajo).

## Cómo usarlo

### Paso 1 — `Paso1_ArmarTablaFails`
1. Ejecuta la macro.
2. Selecciona el archivo de Excel con los resultados de las evaluaciones.
3. Indica el rango de fechas (`Evaluation Date`) que quieres incluir — puedes dejarlo en blanco para no poner límite.
4. El macro crea/limpia el tab `Fails` en este libro con todas las filas "Fail" dentro del rango.
5. Revisa la lista y **borra manualmente** cualquier fila que no quieras generar.

### Paso 2 — `Paso2_GenerarCartasDesdeFails`
1. Ejecuta la macro.
2. Confirma que quieres generar las cartas.
3. El macro detecta los países distintos que quedaron en el tab `Fails` y te pregunta:
   - **¿Usar la misma plantilla para todos?** → seleccionas un solo `.docx` y se usa para todas las cartas.
   - **¿Usar una plantilla distinta por país?** → te va pidiendo, uno por uno, el `.docx` correspondiente a cada país (por ejemplo, `Template_Mexico.docx`, `Template_Colombia.docx`, `Template_Peru.docx`, etc.).
4. Si algún país queda sin plantilla asignada, el macro te avisa antes de continuar.
5. Se genera un archivo `.docx` por fila en la carpeta configurada en `CARPETA_SALIDA`, y el nombre del archivo generado (o el error, si algo falla) se registra en la columna "Archivo Generado" del tab `Fails`.

## Tokens que reconoce la plantilla de Word
Tu plantilla `.docx` debe contener estos tokens entre corchetes, que el macro reemplaza automáticamente:

```
[AGENCIA_ID]
[AGENCIA_NOMBRE]
[AGENCIA_DIRECCION]
[AGENCIA_CIUDAD_PAIS]
[AGENCIA_DIRECCION_COMPLETA]
[FECHA_INFORME_LARGA]
[FECHA_INFORME]
[FECHA_VISITA]
[ESCENARIO]
[NARRATIVA]
```

## Notas
- El archivo de datos se abre siempre en modo solo lectura; nunca se le hace ningún cambio.
- El macro usa automatización de Word tardía (`CreateObject("Word.Application")`), por lo que no necesitas agregar ninguna referencia adicional en el editor de VBA.
- Las traducciones de "Scenario" al español se editan directamente en la función `TraducirEscenario` dentro del módulo.
