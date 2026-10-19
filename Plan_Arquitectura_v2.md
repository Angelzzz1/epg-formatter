# Plan de Arquitectura y Refactorización: EPG Formatter v2

## 1. Contexto y Problema
En la operación del NOC, se reciben archivos Excel (Guías Electrónicas de Programación o EPGs) de diversos proveedores de TV (ej. Bantel, Canal LUZ, Antena 1, Nexxor TV). Estos archivos vienen en formatos con estructuras variadas y cambiantes.

La solución se estructura en `app_v2.py` con métodos determinísticos optimizados y preprocesamiento avanzado:
1. **Modo Automático:** Detecta si la estructura es lineal o matricial mediante análisis de cabeceras y palabras clave con límites de palabra (`\b`).
2. **Modo Lineal:** 1 fila = 1 programa.
3. **Modo Matricial:** Días como columnas, horas como filas (se aplana con `pd.melt`).
4. **Modo Bloque:** Horarios apilados sin fechas explícitas (títulos como "LUNES A VIERNES", "SABADO", "DOMINGO").
5. **Modo Manual:** Permite al operador mapear directamente las columnas necesarias.

## 2. Requisitos Estrictos (El EPG Uploader)
Independientemente de cómo llegue el archivo crudo, la salida exportada (Excel final) debe cumplir inamoviblemente con estas reglas para no romper el sistema destino:
* **Estructura:** Solo se usan 3 columnas con datos (A: Día, B: Hora, C: Programa). Las columnas D a la P deben existir pero estar completamente vacías.
* **Formato de Fecha:** Obligatorio texto plano en formato `MM-DD-YYYY` (separado por guiones).
* **Formato de Hora:** Obligatorio 24 horas (`HH:MM`).
* **Programa:** Texto limpio en MAYÚSCULAS, sin saltos de línea ni metadata basura (ej. "(R)", nombres de presentadores, directores).
* **Timezone:** Ajuste manual de huso horario para empatar la hora local del proveedor con la emisión final.
* **Nombre de Pestaña:** Debe coincidir exactamente con el canal. Máximo 31 caracteres y sin caracteres especiales (`: \ / ? * [ ]`).

## 3. Investigaciones y Hallazgos Estratégicos
Para garantizar la fiabilidad del procesamiento determinístico en producción:

1. **Preprocesamiento Universal de Celdas Combinadas:**
   Antes de evaluar cualquier estructura, se leen los metadatos de celdas combinadas (`openpyxl`) para desfusionar y propagar valores ancla tanto horizontal como verticalmente y en bloques 2D.
2. **Manejo del Peligro de Excel (Text Formatting):** 
   Si escribimos las fechas como objetos *datetime* nativos, Excel podría auto-convertirlas según la configuración regional de cada PC. Se escribe como texto forzado usando `worksheet.write_string()` de *xlsxwriter*.
3. **Validación Visual (Human-in-the-Loop):** 
   Antes de generar el Excel final, la tabla extraída se muestra en un componente editable (`st.data_editor`), permitiendo al operador corregir cualquier detalle antes de la exportación.
4. **Sanity Checks Estrictos:** 
   Si el procesamiento arroja más del 15% de datos nulos o vacíos, el sistema aborta de inmediato con un mensaje claro al operador ("Formato no reconocido, revisa el archivo o avisa al administrador"), evitando generar archivos incompletos o corruptos.

## 4. Arquitectura del Código

### A. `excel_preprocessing.py` (Módulo de Preprocesamiento)
Paso inicial universal del pipeline. Inspecciona los rangos combinados del archivo Excel con `openpyxl`, desfusiona todas las celdas propagando el valor superior izquierdo y entrega un DataFrame regularizado.

### B. `normalization.py` (Módulo Central de Normalización)
Centraliza toda la lógica de limpieza y estandarización:
- Parseo tolerante de horas ("7.30 pm" -> "19:30", "20: 00" -> "20:00").
- Parseo de fechas y días relativos en español.
- Limpieza semántica del título (uso de regex para eliminar metadata de conductores, repeticiones, sinopsis).
- Sanitización y validación de nombres de hojas según las reglas de Excel.

### C. `app_v2.py` (Interfaz de Usuario en Streamlit)
- Detección automática inteligente con descarte de títulos fusionados y word boundaries (`_contiene_palabra`).
- Sanity Check: si la tasa de nulos supera el 15%, detiene la exportación e informa al usuario.
- Detección de Anomalías: resalta en advertencias fechas vacías, horas inválidas o programas sin nombre.
- Editor interactivo para revisión previa a la descarga.

## 5. Decisiones Técnicas Acordadas
- **Enfoque Determinístico:** Procesamiento 100% predecible, rápido y sin dependencias de servicios externos.
- **Timezone:** Ajuste manual en la interfaz para sumar o restar horas según la diferencia horaria del canal.
- **Fallos Totales:** Si un archivo no se reconoce o falla los controles de sanidad, la aplicación muestra un error claro y detiene la ejecución sin emitir archivos en blanco.
