# Registro de Desarrollo - Automatización EPG (Fase 2: Arquitectura IA)

Este documento resume las conversaciones, decisiones y desarrollos realizados durante la **Fase 2** del proyecto, orientada a integrar un Agente IA para procesar formatos desconocidos de EPG.

## 1. El Problema Inicial
Se planteó que la solución original (`app.py` con modos Lineal, Matricial y Bloque) era muy frágil ante el hecho de que llegan nuevos EPGs cada semana/mes con formatos impredecibles. Crear un nuevo script (parser) para cada formato nuevo es insostenible en el NOC.

## 2. La Solución Adoptada: Arquitectura Híbrida
En lugar de usar la IA solo como un "enrutador" de scripts pre-hechos, decidimos usarla como un **Extractor Universal (Parser Dinámico)**. 

### Principios de Ingeniería de Datos (Aprobados):
1. **La IA extrae, Python transforma:** No confiamos en que el LLM formatee fechas u horas en 200 filas (suele fallar). El LLM extrae el texto crudo (`"Lun 15"`, `"7.30pm"`) y Python aplica la lógica determinística (`09-15-2026`).
2. **Structured Outputs (Pydantic):** Forzamos a la API de la IA a devolver un esquema JSON estricto, eliminando problemas de parseo.
3. **Validación y Reintento (Self-Healing):** Si el LLM rompe el esquema, se captura el error y se reenvía al modelo para que se autocorriga.
4. **Fechas como Texto:** Al exportar a Excel (`xlsxwriter`), las fechas y horas se escriben como texto puro forzado (`write_string`) para evitar que la configuración regional del sistema (Windows/Linux) altere el formato `MM-DD-YYYY`.
5. **Human-in-the-Loop:** El operador revisa la extracción en un `st.data_editor` interactivo antes de descargar.
6. **Sanity Checks (Validación de Sanidad):** Si el método rápido (Fast Track) produce más de un 15% de datos nulos, asume que falló y deriva automáticamente el archivo al Agente IA.

## 3. Decisiones Tecnológicas
* **Modelo IA:** **Google Gemini (Familia Flash)**. Elegido por su soporte nativo para esquemas JSON (Structured Outputs), su masiva ventana de contexto y su bajo costo (o nivel gratuito generoso).
* **Gestión de Claves:** Se implementó un archivo `.env` (`GEMINI_API_KEY` y `GEMINI_MODEL_NAME`) para no quemar (hardcodear) la clave ni el modelo en el código, dado que Google actualiza las versiones de Flash periódicamente.

## 4. Refactorización Realizada (`app_v2.py`)
Se duplicó la aplicación para no romper producción, dividiendo la lógica en 3 pilares:

* **`normalization.py`**: El cerebro determinístico compartido. Contiene limpieza de texto (regex para eliminar conductores, metadata), conversión de horas tolerantes y parseo relativo de fechas.
* **`ai_extractor.py`**: El Agente. Toma el Excel crudo, lo limpia de nulos, lo pasa a Markdown y llama a la API de Gemini con Pydantic. Contiene caché por hash para no gastar llamadas repetidas.
* **`app_v2.py`**: La interfaz unificada. Gestiona el flujo, las validaciones de UI (nombre de pestaña de 31 caracteres) y resalta anomalías (horarios duplicados).

## 5. Instrucciones de Despliegue (Migración a otro equipo)
Para continuar el desarrollo o desplegar en otra máquina:
1. Copiar la carpeta completa.
2. Instalar dependencias actualizadas: `pip install -r requirements.txt` (ahora incluye `google-genai`, `pydantic` y `python-dotenv`).
3. Asegurar que el archivo oculto `.env` esté presente con la clave API.
4. Levantar la aplicación: `streamlit run app_v2.py`.

---

## 6. Fase 3: Fix Crítico — Celdas Combinadas de Excel (15-Sep-2026)

### El Bug
Al procesar la grilla de Canal LUZ (`14-09-26 al 20-09-26`) en modo Matricial, el programa "UN MENSAJE A LA CONCIENCIA" (07:30) aparecía solo para el Domingo, pero en la grilla original estaba programado los 7 días de la semana. Otros programas con merges 2D (como "LA CALDERA DE ORACIÓN") también se perdían parcialmente.

### Causa Raíz
El archivo Excel tenía **193 celdas combinadas (merged cells)**. Cuando `pd.read_excel` lee un merge, solo la celda ancla (arriba-izquierda) conserva el valor; el resto queda `NaN`. El `ffill()` vertical que teníamos era una heurística que adivinaba ("si está vacío, asumo que es el mismo programa de arriba"), pero no podía distinguir entre:
- Una celda vacía porque era parte de un merge real.
- Una celda vacía porque genuinamente no hay programación en ese horario.

### La Solución: `excel_preprocessing.py`
Se creó un nuevo módulo (`excel_preprocessing.py`) que usa `openpyxl` para leer los metadatos reales de `ws.merged_cells.ranges` y propagar el valor ancla a todas las celdas del rango. Esto resuelve automáticamente los 3 tipos de merge:
- **Horizontal** → programa repetido en varios días (ej. "UN MENSAJE A LA CONCIENCIA" Lun-Dom).
- **Vertical** → programa que dura 2+ franjas horarias.
- **Bloque 2D** → programa que cubre días Y horas simultáneamente (ej. "LA CALDERA DE ORACIÓN").

### Cambios Arquitectónicos
1. **Paso universal del pipeline:** `expandir_celdas_combinadas()` se ejecuta SIEMPRE como primer paso, antes de decidir qué vía usar (Matricial, Lineal, Bloque, IA). No es un parche del modo Matricial.
2. **Eliminación de `ffill()`:** Se removió el `ffill()` heurístico de `procesar_matricial()`. Ya no es necesario porque los valores reales ya están propagados desde los metadatos del archivo.
3. **Filtrado de columna HORA duplicada:** Canal LUZ tiene una columna "HORA" adicional a la derecha (columna I) solo para legibilidad al imprimir. Ahora se toma solo la primera columna HORA y se ignoran las duplicadas.

### Archivos Modificados
- **`excel_preprocessing.py`** (NUEVO): Módulo de preprocesamiento universal.
- **`app_v2.py`**: Se reemplazó `pd.read_excel()` por `expandir_celdas_combinadas()` como primer paso; se eliminó `ffill()` de `procesar_matricial()`; se filtran columnas HORA duplicadas.

---

## 7. Fase 4: Mejoras en Detección Lineal y Actualización de IA (15-Sep-2026)

### El Problema
Al intentar procesar el EPG de "NEXXOR TV" (un archivo Lineal continuo de Septiembre a Diciembre):
1. **Falsos positivos en detección de cabecera:** El título superior ("EPG CALENDARIO DIARIO...") estaba fusionado en 11 columnas. Al desfusionarlo (gracias al fix de la Fase 3), la palabra "DIARIO" se propagó a todas las columnas. El detector automático buscaba la palabra "DIA" como *subcadena*, por lo que "DIARIO" daba un falso positivo con score altísimo, asumiendo erróneamente que la fila 0 era la cabecera.
2. **Error 404 en Gemini:** El modelo `gemini-2.5-flash` dejó de estar disponible en la API de Google para nuevos usuarios, devolviendo un error 404.
3. **Mapeo estricto de columnas:** Los títulos reales del archivo eran compuestos ("Fecha (AAAA-MM-DD)", "Hora Inicio", "Título del Programa"). La función `encontrar_columnas_lineales` exigía coincidencias exactas (ej. "FECHA" o "PROGRAMA"), por lo que fallaba al mapear estas columnas automáticamente.

### La Solución
1. **Regex con Word Boundaries (`\b`):** Se creó la función `_contiene_palabra()` en `app_v2.py` que usa `re.search(r"\b" + palabra + r"\b")`. Esto asegura que "DIA" solo haga match con la palabra exacta y no con "DIARIO" o "MEDIODIA".
2. **Descarte de Títulos Fusionados:** La heurística `detectar_cabecera_y_modo` ahora ignora las filas donde todos los valores son idénticos (señal inequívoca de que es un título fusionado propagado, no una cabecera de tabla real).
3. **Mapeo flexible:** `encontrar_columnas_lineales` se actualizó para usar `_contiene_palabra()`. Ahora "Hora Inicio" es validado correctamente usando la keyword "HORA".
4. **Actualización de IA:** Se actualizó la variable `GEMINI_MODEL_NAME` en el archivo `.env` para apuntar a `gemini-3.6-flash`.

### Resultado
El sistema ahora procesa el EPG de NEXXOR TV (más de 2,700 filas equivalentes a 113 días continuos) en modo Automático (o Lineal) de forma inmediata, detectando la fila de cabecera correcta (Fila 3) y enrutando cada columna a su lugar sin intervención manual.
