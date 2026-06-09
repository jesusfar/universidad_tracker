# Tracker de universidades con ROR + OpenAlex

Este proyecto actualiza un Excel de universidades agregando identificadores ROR y métricas académicas de OpenAlex.

Versión actual: **1.1**. Incluye soporte para `--all-sheets`, que procesa automáticamente todas las hojas que tengan una tabla reconocible de universidades.

## Qué actualiza

- ROR ID y nombre normalizado de la institución.
- OpenAlex ID.
- Publicaciones totales (`works_count`).
- Citas totales (`cited_by_count`).
- h-index, i10-index y 2yr mean citedness.
- Publicaciones y citas del año objetivo y del año anterior.
- Coordenadas de ROR/OpenAlex, si están disponibles.
- Hoja `Tracker_OpenAlex_ROR`.
- Hoja `Historial_metricas` para comparar ejecuciones.
- Procesamiento de una hoja específica o de todas las hojas con `--all-sheets`.
- Eliminación automática de duplicados entre hojas por nombre+país y, después, por ROR/OpenAlex ID.

## Qué no hace

No actualiza en tiempo real el puesto oficial de QS/THE/ARWU, porque esos rankings se publican por edición. Conserva el ranking base del Excel y agrega métricas dinámicas.

## Instalación

```bash
pip install -r requirements.txt
```

## API key de OpenAlex

OpenAlex pide API key. Se puede pasar de dos maneras:

### Windows CMD

```bat
set OPENALEX_API_KEY=TU_API_KEY
```

### PowerShell

```powershell
$env:OPENALEX_API_KEY="TU_API_KEY"
```

### Linux / macOS

```bash
export OPENALEX_API_KEY="TU_API_KEY"
```

## Uso recomendado con el Excel base

Este modo lee `Top 200 QS 2026`, `Recomendadas AR-LATAM` y cualquier otra hoja que agregues con columnas reconocibles de universidad. Salta automáticamente hojas de notas o fuentes.

```bash
python universidad_ranking_tracker.py ^
  --input data/top_200_universidades_QS_2026_con_ARG_LATAM_recomendadas.xlsx ^
  --output universidades_tracker_actualizado.xlsx ^
  --all-sheets ^
  --target-year 2025
```

En Linux/macOS, reemplazá `^` por `\`.

## Uso con una sola hoja

```bash
python universidad_ranking_tracker.py ^
  --input data/top_200_universidades_QS_2026_con_ARG_LATAM_recomendadas.xlsx ^
  --output universidades_tracker_actualizado.xlsx ^
  --sheet "Top 200 QS 2026" ^
  --header-row 4 ^
  --target-year 2025
```

## Prueba rápida con pocas universidades

Con todas las hojas:

```bash
python universidad_ranking_tracker.py --input data/top_200_universidades_QS_2026_con_ARG_LATAM_recomendadas.xlsx --output prueba_tracker.xlsx --all-sheets --limit 10
```

Con una sola hoja:

```bash
python universidad_ranking_tracker.py --input data/top_200_universidades_QS_2026_con_ARG_LATAM_recomendadas.xlsx --output prueba_tracker.xlsx --sheet "Top 200 QS 2026" --limit 10 --header-row 4
```

## Forzar actualización sin usar caché

```bash
python universidad_ranking_tracker.py --input data/top_200_universidades_QS_2026_con_ARG_LATAM_recomendadas.xlsx --output universidades_tracker_actualizado.xlsx --all-sheets --force
```

## Scraping de rankings oficiales públicos

El script también puede recolectar datos publicados oficialmente cuando la fuente los expone como HTML, JSON, CSV o Excel público.

```bash
python universidad_ranking_tracker.py ^
  --scrape-rankings ^
  --ranking-sources qs,the,arwu ^
  --rankings-output rankings_oficiales_publicos.xlsx
```

El Excel generado incluye:

- `Rankings_Oficiales`: filas extraídas desde datos públicos oficiales.
- `Fuentes_Rankings`: estado por fuente, URL consultada, cantidad de filas y notas de bloqueo o parcialidad.

Si QS, THE u otra fuente muestran una descarga oficial en el navegador, se puede pasar el enlace directamente:

```bash
python universidad_ranking_tracker.py ^
  --scrape-rankings ^
  --ranking-sources qs ^
  --ranking-years 2026 ^
  --ranking-data-urls qs=https://URL_OFICIAL_DEL_ARCHIVO.xlsx
```

El scraper no intenta eludir Cloudflare, paywalls, logins ni restricciones técnicas. Si una fuente oficial no expone datos tabulares públicos para descarga automática, queda registrado en `Fuentes_Rankings`.

## Opciones útiles

- `--all-sheets`: procesa todas las hojas con una columna reconocible de universidad.
- `--no-dedupe`: desactiva la eliminación de duplicados entre hojas.
- `--sheet "Nombre de hoja"`: procesa solo una hoja.
- `--scrape-rankings`: genera un Excel independiente con rankings oficiales públicos disponibles.

## Archivos generados

- `universidades_tracker_actualizado.xlsx`: Excel enriquecido.
- `universidades_tracker_cache.json`: caché local de respuestas para acelerar futuras ejecuciones.
