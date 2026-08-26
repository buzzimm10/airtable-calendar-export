# AI Context - Disney EMEA Calendar Automation

## 1. Project Overview
Este proyecto es una automatización en Python que lee registros de Airtable y genera calendarios de YouTube en formato Excel (`.xlsx`) para Disney Channel (DC) y Disney Junior (DJR) en EMEA y mercados locales.
*   **Librería principal:** `openpyxl`
*   **Archivos clave:**
    *   `src/exporters/full_calendar_djr.py` (Exportador maestro DJR, ID's, Metadata Spain)
    *   `src/exporters/full_calendar.py` (Exportador maestro DC)
    *   `src/services/local_calendars_service.py` (Generador de calendarios locales por país)
    *   `src/services/qa_service.py` (Validador de reglas antes de exportar en Streamlit)

## 2. Openpyxl & Excel Golden Rules (CRITICAL)
*   **NUNCA USAR `ws.delete_rows()` ni `ws.insert_rows()`:** Las plantillas base `.xltx` contienen fórmulas, referencias relativas y celdas combinadas (*merged cells*). Borrar o insertar filas físicamente destruye el layout a partir de la semana 2.
*   **Método de Ocultación Segura ("La Guillotina"):** Para eliminar contenido o "basura" visual, el código debe vaciar las celdas (`cell.value = None`, quitar `fill` y `border`) y colapsar la fila visualmente:
    `ws.row_dimensions[r].hidden = True`
    `ws.row_dimensions[r].height = 0`
*   **Respetar Fórmulas y Merges:** El código NUNCA debe inyectar fechas día por día ni sobreescribir celdas si la plantilla ya las calcula con fórmulas. Por ejemplo, en DJR Metadata Spain, la fecha inicial (`first_monday`) se inyecta UNA SOLA VEZ en la celda combinada de la primera semana (ej. `D7:G7`), dejando que Excel calcule el resto.
*   **Heights (AutoFit):** Para celdas de texto muy variable (Branding, TX IDs, etc.), si se desea que Excel aplique su propio "Autoajuste de altura", **no** se debe forzar la altura desde Python (`ws.row_dimensions[row].height`).

## 3. Business Logic: Disney Junior (DJR) Master
*   **Slots Dinámicos por Día:** La plantilla permite hasta 4 contenidos por día. El código asigna columnas dinámicamente (`get_djr_target_columns`).
    *   Programming tiene prioridad de izquierda a derecha.
    *   Mediaplanning rellena de derecha a izquierda.
    *   El 4º slot está reservado visualmente para `YT TEAM UK`, pero si hay 4 contenidos (ej. 3 de Prog y 1 de Media), se ocupará ese 4º slot para no perder data.
*   **Tab "ID's" (Overon Tracking):** 
    *   **Filtro:** Solo exporta contenido del `Programming Hub` que sea `Compilation` (con TCs llenos), `MEGA Compilation`, o `Live Stream + MEGA`.
    *   **Parser de MEGAs:** Si es un MEGA y el campo de Airtable "Content Details" no tiene el desglose exacto, lee el campo "Notes" para extraer los sub-títulos (ignorando frases como "compile" o "reuse"). 
    *   **Estética:** Cabeceras con relleno `#F8CBAD`. Bordes negros delgados `Side(style='thin')` por cada celda individual y combinada. Columna A unificada con `Activity Name` + `dd/mm`.
*   **Tab "Metadata Programming Spain":**
    *   Extrae `ES Title` y `ES Description`. Hace *fallback* al inglés si están vacíos.
    *   Respeta el mapa de filas estricto (semanas en filas 7, 17, 27, 37, 47).
    *   **Estética:** Letra de títulos/sinopsis en **negrita y verde `#00B050`**. El bloque de cada día se pinta de blanco (`#FFFFFF`) si corresponde al mes en curso, o de gris (`#D6DCE4`) si es de otro mes, desde la fila `Property` hasta `Planning Strategy` (sin tocar `Owner`).

## 4. QA Service (Streamlit App)
*   En `src/services/qa_service.py`, el límite de capacidad diaria para que no salte el `[Blocker]` es de **4 contenidos totales por día** (independientemente del hub), no 2 por hub.