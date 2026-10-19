"""
excel_preprocessing.py — Preprocesamiento Universal de Excel
=============================================================
Primer paso del pipeline: se ejecuta SIEMPRE sobre cualquier archivo
Excel que entra, ANTES de decidir qué vía de procesamiento usar.

Resuelve el problema clásico de pandas + celdas combinadas:
cuando pd.read_excel lee un merge, solo la celda ancla (arriba-izquierda)
tiene valor; el resto del rango queda NaN.

Este módulo usa openpyxl para leer los metadatos reales de
ws.merged_cells.ranges y propagar el valor ancla a todas las celdas
del rango combinado. Esto resuelve automáticamente los 3 tipos de merge:
  - Horizontal (programa repetido en varios días)
  - Vertical   (programa que dura 2+ franjas horarias)
  - Bloque 2D  (programa que cubre días Y horas simultáneamente)

No adivina ni usa heurísticas (ffill). Lee metadatos reales del archivo.
"""

import openpyxl
import pandas as pd
import io
import logging

logger = logging.getLogger(__name__)


def expandir_celdas_combinadas(fuente, nombre_hoja=None):
    """Lee un Excel con openpyxl, desfusiona TODAS las celdas combinadas
    propagando el valor ancla, y retorna un DataFrame limpio.

    Args:
        fuente: Ruta al archivo Excel (str/Path) o bytes en memoria
                (bytes / io.BytesIO).
        nombre_hoja: Nombre exacto de la pestaña a leer.
                     Si es None, usa la hoja activa.

    Returns:
        pd.DataFrame con todas las celdas desfusionadas y valores
        propagados. Sin cabecera (header=None equivalente).
    """
    # Abrir workbook — aceptar tanto ruta como bytes
    if isinstance(fuente, (bytes, bytearray)):
        fuente = io.BytesIO(fuente)
    elif isinstance(fuente, io.BytesIO):
        pass  # ya es BytesIO
    # else: asumimos que es una ruta string/Path

    wb = openpyxl.load_workbook(fuente, data_only=True)

    if nombre_hoja:
        if nombre_hoja not in wb.sheetnames:
            raise ValueError(
                f"La hoja '{nombre_hoja}' no existe en el archivo. "
                f"Hojas disponibles: {wb.sheetnames}"
            )
        ws = wb[nombre_hoja]
    else:
        ws = wb.active

    # Snapshot de rangos combinados ANTES de modificar.
    # Nunca iterar directo sobre ws.merged_cells.ranges mientras se
    # modifica, porque unmerge_cells() altera la colección en vivo.
    rangos = list(ws.merged_cells.ranges)
    n_merges = len(rangos)

    if n_merges > 0:
        logger.info(
            f"Hoja '{ws.title}': desfusionando {n_merges} rango(s) combinado(s)"
        )

    for rango in rangos:
        min_col, min_row, max_col, max_row = rango.bounds
        valor_ancla = ws.cell(row=min_row, column=min_col).value

        # Desfusionar primero
        ws.unmerge_cells(str(rango))

        # Propagar el valor ancla a todas las celdas del rango
        for fila in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=fila, column=col, value=valor_ancla)

    # Convertir a DataFrame
    data = list(ws.values)
    wb.close()

    if not data:
        return pd.DataFrame()

    df = pd.DataFrame(data)

    # Eliminar filas completamente vacías
    df = df.dropna(how="all").reset_index(drop=True)

    return df, n_merges


def obtener_hojas(fuente):
    """Retorna la lista de nombres de hojas sin cargar todo el workbook.

    Útil para el selector de pestaña en Streamlit sin duplicar lectura.

    Args:
        fuente: Ruta al archivo Excel (str/Path) o bytes en memoria.

    Returns:
        list[str] con los nombres de las hojas.
    """
    if isinstance(fuente, (bytes, bytearray)):
        fuente = io.BytesIO(fuente)
    elif isinstance(fuente, io.BytesIO):
        pass

    wb = openpyxl.load_workbook(fuente, read_only=True)
    nombres = wb.sheetnames
    wb.close()
    return nombres
