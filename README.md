# Series horarias de CPF

Lee los índices publicados por el Coordinador Eléctrico de Chile desde un año inicial hasta el año actual. Descubre las carpetas de años y meses, selecciona la **versión de mayor número con documentos publicados en cada mes** y lee exclusivamente la pestaña `CPF Horario` del Excel de esa versión. No promedia las horas: conserva el detalle de cada día.

Requiere Python 3.10 o posterior y acceso HTTPS a `www.coordinador.cl` y `coordinador.cl`.

## Instalación y ejecución

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python Main --unidades datos/unidades.csv --anio 2025
```

En Windows, reemplaza `.venv/bin/python` por `.venv\Scripts\python.exe`.

- `--anio 2025`: incluye 2025 y todos los años publicados hasta el actual. El año actual se calcula en `America/Santiago`.
- `--hasta 2026`: permite fijar el año final.
- `--mes 1 2`: limita opcionalmente la ejecución a enero y febrero de cada año seleccionado.
- `--hoja-unidades CPF`: hoja del archivo de entrada; `CPF` es el valor predeterminado. El anexo también tiene columnas `Unidad` en otros servicios y en el historial; no se mezclan esas hojas.
- `--salida resultados/mi_ejecucion`: carpeta nueva donde guardar la salida. Si ya existe, se detiene para evitar sobrescribir resultados. Sin este argumento se crea una carpeta con fecha y hora.
- `--cache .cache/cpf`: conserva los Excel descargados para siguientes ejecuciones. Las páginas se consultan siempre para detectar nuevas versiones.
- `--actualizar`: vuelve a descargar los Excel aunque ya existan en caché, por ejemplo ante una corrección del archivo sin cambio de URL.

Se incluye `datos/unidades.csv` con las 370 unidades únicas extraídas de la columna `Unidad` de la hoja `CPF` del anexo proporcionado (`2026.11.05 Anexo C. Control de Frecuencia.xlsx`). Puedes actualizar este CSV o pasar directamente otro Excel con `--unidades`.

También se acepta un CSV UTF-8 con una columna `Unidad` (separadores coma, punto y coma o tabulación). En Excel se busca la cabecera en las primeras 50 filas de la hoja elegida. Se omiten celdas vacías y unidades duplicadas.

## Resultados

La carpeta de salida contiene:

- `desempeno.csv`: todas las unidades, con fecha, hora y `Desempeno` (DCPF).
- `factor_desempeno.csv`: todas las unidades, con fecha, hora y `Factor_desempeno` (Fd_CPF).
- `unidades_sin_datos.csv`: unidades del listado sin coincidencias en todo el período procesado.
- `unidades/`: una carpeta por unidad con esas mismas dos series por separado. Las unidades sin registros tienen archivos con cabecera y sin filas.
- `reporte.json`: versión y URL de cada mes, hash SHA-256 del archivo utilizado, cantidad de registros, unidades sin datos por mes y en todo el período, métricas vacías y horas repetidas conservadas. Incluye la correspondencia entre unidad y carpeta.

Cada CSV incluye `Unidad`, `Unidad_fuente`, `Fecha`, `Hora`, la métrica correspondiente, `Version` y `Fuente`. Se utiliza UTF-8 con BOM y coma como separador. La hora se conserva exactamente como figura en el Excel, sin convertirla a UTC ni agregar o quitar horas por cambios de horario. Se conserva el orden de la fuente dentro de cada mes.

Las unidades se comparan ignorando mayúsculas, acentos y espacios repetidos. **No se hacen coincidencias aproximadas ni se descomponen nombres de arreglos o grupos de unidades**: nombres distintos se informan como ausentes para evitar asignar datos de otra unidad. Los valores de desempeño, textos y celdas vacías se conservan; nunca se convierten ausencias a cero ni se interpolan datos.

Solo se procesan meses publicados; los futuros o todavía no publicados no se inventan. Si una carpeta de versión indica explícitamente `Sin documentos.`, se busca la siguiente versión y se registra la carpeta vacía en el reporte. Si todas las versiones están vacías, se registra el mes como no publicado. Una versión con documentos pero sin un Excel único, sin cabeceras requeridas o con fechas fuera de su mes detiene la ejecución con un error: esos errores no provocan el uso de una versión anterior. La salida se publica en su carpeta final únicamente si termina toda la extracción. Los Excel válidamente descargados quedan en caché si un paso posterior falla.

## Pruebas

```bash
python3 -m unittest discover -s tests -v
```

Las pruebas verifican selección de versión, paginación, selección de hoja y unidades, columnas reales de CPF, conservación de valores y horas repetidas, rechazo de datos fuera de mes y generación de las dos series con reporte de unidades ausentes.
