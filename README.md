# Sistema de Scoring — Oficiales

Migración de las tres aplicaciones de escritorio que había en esta carpeta
(`diccionario_OFICIALES.pyw` en PyQt6, `productividad_general.pyw` y
`productividad_imputados.pyw` en Tkinter) a **una sola aplicación Flet 1.0**
con pestañas.

Las versiones anteriores quedaron guardadas como `*.pyw.bak` y el entorno roto
anterior como `.venv_roto_bak/`.

## Cómo ejecutar

```bash
./scoring-poli          # lanzador
uv run python main.py   # alternativa
```

Para recrear el entorno desde cero:

```bash
uv sync
```

Entorno verificado: **Python 3.14.3**, **flet 1.0.3**, **flet-desktop 1.0.3**,
pandas 3.0.6, openpyxl 3.1.5.

## Estructura

```
main.py                          Aplicación unificada: 4 pestañas
scoring-poli                     Lanzador (bash)
pyproject.toml                   Dependencias (package = false)
variantes_detenidos.xlsx         Tabla de puntajes por tipo de evento y barrio
core/                            Lógica de negocio, sin Flet (pandas puro)
  comunes.py                     Normalización, hojas, rutas, validaciones
  diccionario_oficiales.py       Reconciliación nombres → LP y reversión
  productividad_general.py       Unificación de eventos del CMU
  productividad_imputados.py     Eventos con imputados
  calculo_formulas.py            Cálculo de P.EVENTO / % ZONA / P. TOTAL
ui/                              Interfaz Flet 1.0
  comunes.py                     PanelRegistro, diálogos, selección de hoja
  seleccion.py                   CampoArchivo: selector de archivos en 3 capas
  vista_diccionario.py           Pestaña «Diccionario»
  vista_general.py               Pestaña «Productividad General»
  vista_imputados.py             Pestaña «Productividad Imputados»
  vista_formulas.py              Pestaña «Calculo»
complementos/                    Procesadores pandas reutilizados sin cambios
DATOS/                           Archivos de trabajo y salidas
*.pyw.bak                        Versiones anteriores (PyQt6 / Tkinter)
```

`core/` no importa Flet en ningún punto: se puede probar y reutilizar por
separado. `complementos/` quedó intacto.

## Qué hace cada pestaña

| Pestaña | Entrada | Proceso | Salida |
|---|---|---|---|
| **Diccionario** | Diccionario Maestro (`Sheet1`) + archivo con columna `OFICIAL` | Reemplaza nombres por `LP`; si falta algún nombre, lo agrega al diccionario con backup y se detiene | `<archivo>_PROCESADO.xlsx`, `ANTIGUO_DICCIONARIO.xlsx` |
| **Diccionario** → Revertir | `_PROCESADO` + hoja `validacion` | Reemplaza `LP` por nombres | `<archivo>_REVERTIDO.xlsx` |
| **Productividad General** | Archivo maestro del CMU | Trunca filas vacías, limpia GAP/SAE, extrae columnas por posición y combina los complementos | `eventos_unificados.xlsx` (hoja `eventos`, junto al origen) |
| **Productividad Imputados** | Archivo maestro del CMU | Aplica las cuatro reglas de imputados | `DATOS/productividad_imputados.xlsx` |
| **Calculo** | Listado generado por Imputados (`DATOS/productividad_imputados.xlsx`) | Completa `P.EVENTO`, `% ZONA` y `P. TOTAL` según la tabla de variantes | `DATOS/productividad_imputados_calculado.xlsx` |

## Calculadora de puntajes (pestaña «Calculo»)

Toma el listado generado por **Productividad Imputados** y completa tres
columnas que vienen vacías:

| Columna | Cómo se calcula |
|---|---|
| `P.EVENTO` | Puntaje base del tipo de evento, según `variantes_detenidos.xlsx` |
| `% ZONA` | Porcentaje del barrio (0, 25 %, 50 % o 75 %). Para `CONTRAVENCION` siempre 0 |
| `P. TOTAL` | `P.EVENTO + (P.EVENTO × % ZONA)`, redondeado a 4 decimales |

Puntajes base de la tabla:

| Tipo de evento | `P.EVENTO` |
|---|---|
| DELITO COMPLEJO | 1,50 |
| ROBO/HURTO | 1,25 |
| DELITO | 0,50 |
| COLABORACION CON INTERVENCION DIRECTA | 0,25 |
| CONTRAVENCION | 0,25 |

Equivalencias aplicadas: `INTERVENCION DESTACADA` → DELITO COMPLEJO (1,50) y
`COLABORACION CON IMPUTADO/S POSITIVA` → COLABORACION CON INTERVENCION DIRECTA
(0,25).

- **Salida:** `DATOS/productividad_imputados_calculado.xlsx` con dos hojas,
  `Listado` y `Acumulado por oficial`; el listado original no se modifica.
- Si un tipo de evento o un barrio no está en la tabla, la fila queda en blanco
  y se informa en el registro y en el diálogo final.
- El campo del listado se completa solo si el archivo ya existe.
- Si elegís por error el archivo maestro del CMU, la app lo detecta y pide el
  listado.

### Hoja «Acumulado por oficial»

Una fila por oficial, ordenada de mayor a menor por **Total Suma de P. TOTAL**.
El encabezado ocupa dos filas, como en el reporte:

- **Fila 1:** `oficial`, y por cada tipo de evento un par `cantidad` /
  `puntaje`; al final `Total Cuenta de TIPO EVENTO` y
  `Total Suma de P. TOTAL`.
- **Fila 2:** el nombre del tipo de evento, repetido en las dos columnas del
  par.

Tipos y orden: `DELITO`, `ROBO/HURTO`, `INTERVENCION DESTACADA`,
`CONTRAVENCION`, `COLABORACION CON IMPUTADO/S POSITIVA`. Si aparece un tipo
nuevo, se agrega al final por orden alfabético.

`cantidad` cuenta los eventos de ese tipo y `puntaje` suma su `P. TOTAL`. La
última fila es **Total general**. Los puntajes se muestran con 4 decimales y
los nombres de oficial se comparan sin espacios externos (así «MEZA, GASTON »
y «MEZA, GASTON» son el mismo).

Ejemplo de la primera fila: `MEZA, GASTON` → 16 DELITO (8,375), 5 ROBO/HURTO
(6,25), 0 INTERVENCION DESTACADA, 8 CONTRAVENCION (2), 5 COLABORACION (1,875);
total 34 eventos y 18,5 puntos.

**Cómo actualizar la tabla:** reemplazá `variantes_detenidos.xlsx` en la carpeta
del proyecto. Se leen la sección `DESCRIPCION` (tipo → puntaje) y la sección
`BARRIO` (barrio → porcentaje, escrito como `25/100` o como `0,25`). Las
columnas de monitoreo horario no se usan.

## Bugs corregidos

Cada corrección está comentada en el código y no altera los nombres ni las
rutas de los archivos de salida.

| # | Problema en la versión anterior | Corrección |
|---|---|---|
| 1 | `productividad_imputados.pyw` buscaba el cierre `"FINALIZA CON IMPUTADO/S PROPIOS"`, texto que **no existe** en los datos: la rama de contraventores nunca generaba filas | Se acepta `"FINALIZA CON IMPUTADO/S PROPIO/S"` (y la variante vieja). Recupera **160 filas CONTRAVENCION** |
| 2 | El proceso inverso buscaba `LP` y `APELLIDO Y NOMBRE` dentro de la hoja `validacion`, que usa `LP/DNI` y `NOMBRE Y APELLIDO`: **fallaba siempre** | Se aceptan ambos juegos de encabezados; la reversión funciona (51 entradas) |
| 3 | `ok = procesar_excel(...)` comparaba una tupla de 3 elementos: la interfaz informaba «Proceso finalizado» **aunque hubiera fallado** | El resultado se propaga como objeto con éxito/error reales |
| 4 | La limpieza de GAP/SAE usaba las posiciones 0-based 13 y 14 (`SAE` y `UBICACION`) en lugar de 12 y 13 (`GAP` y `SAE`): GAP nunca se limpiaba y `UBICACION` se convertía a NaN | Se limpian las columnas correctas (sin efecto en los datos actuales, evita corromper `UBICACION`) |
| 5 | Un nombre del diccionario **sin LP** contaba como coincidencia y borraba el nombre en el `_PROCESADO` (pérdida de dato) | Se informa como «sin ID» y bloquea la finalización hasta completarlo |
| 6 | El Diccionario Maestro (hojas `Sheet1` + `validacion`) obligaba a elegir hoja en cada corrida; elegir `validacion` hacía fallar la validación | Se elige automáticamente la hoja con las columnas requeridas; sólo se pregunta si hay más de una candidata |
| 7 | Los reportes indicaban `índice + 1` como fila | Se informa la fila real de Excel (`índice + 2`) |
| 8 | Las filas `DELITO` y `COLABORACION` no incluían las columnas `TURNO`, `P.EVENTO`, `% ZONA` y `P. TOTAL` (quedaban `NaN`) | Todas las filas tienen las mismas 11 columnas |
| 9 | Comparaciones frágiles: igualdad exacta de texto, valores numéricos como texto, `.str.lower()` sobre columnas no textuales | Se normaliza (`strip` + mayúsculas), se convierten cantidades con tolerancia y `BARRIO` no rompe si no es texto |
| 10 | Si un complemento fallaba, el error se descartaba («No procesado») | El error aparece en el registro y en el resumen final |
| 11 | El registro decía «Col 7 → OPERADOR» cuando la extracción usa la columna 6 (`OFICIAL`); typo «guardada» | Textos corregidos |
| 12 | La condición de colaboración de Imputados buscaba `"COLABORACION CON IMPUTADO/S POSITIVA"`, pero el CMU dice `"COLABORACION CON INTERVENCION PROPIA"`: se perdían 128 filas | Se aceptan ambas redacciones (y las variantes con/sin barra) |

## Selección de archivos (importante en macOS 26)

En **macOS 26** el servicio `FilePicker` de Flet puede no abrir ningún diálogo y
no informar error: el botón parece no responder. Es un problema conocido del
cliente, reportado en
[#5700](https://github.com/flet-dev/flet/issues/5700) y
[#5886](https://github.com/flet-dev/flet/issues/5886), y el mantenedor no logra
reproducirlo porque usa macOS 15.

Por eso cada campo de archivo (`ui/seleccion.py`) ofrece **tres caminos** y
todos escriben en el mismo campo:

| Camino | Botón | Cómo funciona |
|---|---|---|
| 1 | **Examinar** | `ft.PickFiles`, la acción de cliente de Flet 1.0: abre el diálogo **dentro del gesto del clic**, sin ida y vuelta a Python. Es el camino recomendado por Flet y el único que funciona en navegadores. |
| 2 | **Diálogo del sistema** | Abre el selector nativo del sistema operativo sin pasar por Flet: `osascript` en macOS, PowerShell en Windows, zenity/kdialog en Linux. |
| 3 | **Campo de ruta** | El campo es editable: la ruta se puede escribir o pegar a mano (⌘V). |

Si a los 6 segundos de presionar «Examinar» no llegó ninguna selección, aparece
un aviso recordando los caminos 2 y 3. El aviso se limpia solo al elegir o
escribir una ruta.

> En macOS el diálogo del camino 2 puede aparecer **detrás** de la ventana de la
> app: buscá la ventana «choose file» en el Dock o con ⌘Tab.

## Notas de implementación (Flet 1.0)

- Punto de entrada `ft.run(main)`; en 1.0 ya no existe `ft.app`.
- Controles de 1.0: `ft.Button`, `ft.Icons`/`ft.Colors` como enums,
  `ft.Event[ft.Button]` en los manejadores.
- Pestañas con `ft.Tabs` + `ft.TabBar` + `ft.TabBarView`.
- `ft.FilePicker` es un servicio que **se auto-registra** con la página;
  `page.services.append(...)` ya no aplica.
- El procesamiento corre en un hilo (`asyncio.to_thread`) con un
  `ProgressRing`, así la ventana no se congela.
- Cada vista escribe su registro en una cola (`queue.Queue`) que se vuelca al
  panel desde el hilo de eventos: **nunca se tocan controles desde el hilo de
  trabajo**.

## Verificación realizada

| Prueba | Resultado |
|---|---|
| Productividad General contra `DATOS/CMU - JUNIO.xlsx` | Salida **idéntica** a la de la versión anterior: 5820 × 18, 0 celdas distintas |
| Productividad Imputados | 499 filas (antes 339): recupera las 160 `CONTRAVENCION`; el resto coincide |
| Diccionario, caso 100 % | Genera `_PROCESADO` con los IDs y habilita el botón de reversión |
| Diccionario, caso < 100 % | Agrega los nombres faltantes, crea `ANTIGUO_DICCIONARIO.xlsx`, conserva la hoja `validacion` y aborta con el mensaje original |
| Proceso inverso | `_REVERTIDO` con nombres (51 entradas en el mapa inverso) |
| `main.py` y las tres vistas | Verificados de punta a punta con una página simulada: pestañas, flujos, diálogos y manejo de error |
| Selector de archivos en capas | Verificado con página simulada: los 4 campos usan `ft.PickFiles`, la selección y la cancelación llegan bien, el comando de `osascript` se arma correctamente (y devuelve `None` si se cancela) y la ruta escrita a mano se respeta |

La **ventana real no se pudo abrir dentro del entorno de la sesión**: el
cliente `Flet.app` se lanza pero no llega a conectar con Python, por lo que la
corrutina de la app nunca arranca. Es una limitación del entorno de ejecución,
no del código. Para la comprobación visual:

```bash
./scoring-poli
```

## Pendientes por decidir

- `src/informe_scoring_pol/` (el stub `main()` de la plantilla de paquete)
  quedó sin uso: se puede eliminar.
- `*.pyw.bak` y `.venv_roto_bak/` se pueden borrar una vez validado todo.
