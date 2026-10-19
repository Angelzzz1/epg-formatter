"""
normalization.py — Módulo Central de Normalización EPG
======================================================
Centraliza TODA la lógica de limpieza y transformación de datos.
Usado por todos los modos de procesamiento (Lineal, Matricial, Bloque, Manual).

Principio: una sola fuente de verdad para las reglas de parseo.
"""

import re
import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date

# ══════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════

EXCEL_EPOCH = datetime(1899, 12, 30)

DIA_OFFSET = {
    "LUNES": 0,
    "MARTES": 1,
    "MIERCOLES": 2,
    "MIÉRCOLES": 2,
    "JUEVES": 3,
    "VIERNES": 4,
    "SABADO": 5,
    "SÁBADO": 5,
    "DOMINGO": 6,
}

# Mapeo de abreviaturas comunes en español → día de la semana (0=Lun)
DIA_ABREVIADO = {
    "LUN": 0, "L": 0,
    "MAR": 1, "M": 1,
    "MIE": 2, "MI": 2, "MIÉ": 2,
    "JUE": 3, "J": 3,
    "VIE": 4, "V": 4,
    "SAB": 5, "SA": 5, "SÁB": 5,
    "DOM": 6, "D": 6, "DO": 6,
}

# Patrones regex para limpiar metadata basura de títulos de programas
_PATRONES_BASURA = [
    r"\(R\)",                          # (R) = Repetición
    r"\(REP\.?\)",                     # (REP) o (REP.)
    r"\bREPETICION\b",                 # REPETICION suelto
    r"\bREPETICIÓN\b",
    r"\bPRESENTADOR\s*:.*",            # PRESENTADOR: Nombre
    r"\bCONDUCTOR\s*:.*",             # CONDUCTOR: Nombre
    r"\bCONDUCE\s*:.*",               # CONDUCE: Nombre
    r"\bDIRECTOR\s*:.*",              # DIRECTOR: Nombre
    r"\bCLASIFICACION\s*:.*",         # CLASIFICACION: ATP
    r"\bCLASIFICACIÓN\s*:.*",
    r"\bGENERO\s*:.*",                # GENERO: Drama
    r"\bGÉNERO\s*:.*",
    r"\bSINOPSIS\s*:.*",             # SINOPSIS: ...
    r"\bTEMP(?:ORADA)?\s*\d+\s*-?\s*EP(?:ISODIO)?\s*\d+",  # TEMP 1 - EP 5
]
_RE_BASURA = re.compile("|".join(_PATRONES_BASURA), re.IGNORECASE)

# Límite de nombre de hoja en Excel
MAX_LONGITUD_HOJA = 31
_RE_CHARS_HOJA_INVALIDOS = re.compile(r'[:\\/?\*\[\]]')


# ══════════════════════════════════════════════════════════════
# CONVERSIÓN DE FECHAS
# ══════════════════════════════════════════════════════════════


def convertir_fecha(valor):
    """Convierte cualquier representación de fecha a un objeto date de Python.

    Soporta:
      - Números seriales de Excel (ej. 46272)
      - Objetos datetime / date / pd.Timestamp
      - Strings en formatos comunes (MM-DD-YYYY, DD/MM/YYYY, etc.)
    """
    if valor is None:
        return None
    if isinstance(valor, float) and (np.isnan(valor) or np.isinf(valor)):
        return None
    if isinstance(valor, pd.Timestamp):
        return valor.date()
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)):
        n = int(valor)
        if 1 < n < 200000:
            return (EXCEL_EPOCH + timedelta(days=n)).date()
        return None
    if isinstance(valor, str):
        valor = valor.strip()
        for fmt in (
            "%m-%d-%Y",
            "%m/%d/%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%m-%d-%y",
            "%d/%m/%y",
            "%m/%d/%y",
            "%d-%m-%y",
        ):
            try:
                return datetime.strptime(valor, fmt).date()
            except ValueError:
                continue
    return None


# ══════════════════════════════════════════════════════════════
# CONVERSIÓN DE HORAS
# ══════════════════════════════════════════════════════════════


def convertir_hora(valor):
    """Convierte cualquier representación de hora a string HH:MM (24h).

    Soporta:
      - Fracciones de Excel (ej. 0.5 → 12:00, 0.75 → 18:00)
      - Objetos datetime / time
      - Strings: '19:30', '7:30 PM', '7.30pm', '1/08/2024 00:54:27'
    """
    if valor is None:
        return None
    if isinstance(valor, float) and (np.isnan(valor) or np.isinf(valor)):
        return None
    if isinstance(valor, datetime):
        return valor.strftime("%H:%M")
    if hasattr(valor, "hour") and hasattr(valor, "minute"):
        return f"{valor.hour:02d}:{valor.minute:02d}"
    if isinstance(valor, (int, float)):
        frac = float(valor) % 1
        if frac < 0:
            frac += 1
        total_min = round(frac * 1440)
        h = (total_min // 60) % 24
        m = total_min % 60
        return f"{h:02d}:{m:02d}"
    elif isinstance(valor, str):
        valor_limpio = valor.strip()
        # Limpiar espacios alrededor de los dos puntos (ej. "20: 00")
        valor_limpio = re.sub(r'\s*:\s*', ':', valor_limpio)
        # Reemplazar punto como separador (ej. "7.30pm")
        valor_limpio = re.sub(r'(\d{1,2})\.(\d{2})', r'\1:\2', valor_limpio)

        # Detectar AM/PM
        es_pm = bool(re.search(r'[Pp][Mm]', valor_limpio))
        es_am = bool(re.search(r'[Aa][Mm]', valor_limpio))
        valor_limpio = re.sub(r'\s*[AaPp][Mm]\s*', '', valor_limpio)

        match = re.search(r"(\d{1,2}):(\d{2})", valor_limpio)
        if match:
            h = int(match.group(1))
            m = int(match.group(2))
            if es_pm and h < 12:
                h += 12
            elif es_am and h == 12:
                h = 0
            h = h % 24
            return f"{h:02d}:{m:02d}"
    return None


# ══════════════════════════════════════════════════════════════
# LIMPIEZA DE TEXTO DE PROGRAMA
# ══════════════════════════════════════════════════════════════


def limpiar_programa(texto, solo_primera_linea=False):
    """Limpia el nombre del programa: MAYÚSCULAS, sin saltos de línea,
    sin espacios extras, sin metadata basura (conductores, sinopsis, etc.).

    Si solo_primera_linea=True, toma únicamente la primera línea no vacía
    (útil para el formato matricial donde la celda incluye conductor/descripción).
    """
    if pd.isna(texto) or texto is None:
        return ""
    texto = str(texto)
    if solo_primera_linea:
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        texto = lineas[0] if lineas else ""
    else:
        texto = texto.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    texto = texto.replace("\t", " ")
    # Eliminar non-breaking spaces
    texto = texto.replace("\u00a0", " ")
    texto = re.sub(r"\s{2,}", " ", texto).strip()
    texto = texto.upper()
    # Eliminar metadata basura con regex
    texto = _RE_BASURA.sub("", texto).strip()
    texto = re.sub(r"\s{2,}", " ", texto).strip()
    return texto


# ══════════════════════════════════════════════════════════════
# AJUSTE DE ZONA HORARIA
# ══════════════════════════════════════════════════════════════


def ajustar_hora(hhmm, offset):
    """Suma o resta horas a un string HH:MM. Solo mueve la hora, NO la fecha."""
    if not hhmm or offset == 0:
        return hhmm
    try:
        h, m = map(int, hhmm.split(":"))
        h = (h + offset) % 24
        return f"{h:02d}:{m:02d}"
    except Exception:
        return hhmm


# ══════════════════════════════════════════════════════════════
# FORMATO DE SALIDA
# ══════════════════════════════════════════════════════════════


def formatear_fecha_salida(d):
    """Formatea un objeto date a MM-DD-YYYY con guiones (texto puro)."""
    if d is None:
        return None
    try:
        return d.strftime("%m-%d-%Y")
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
# UTILIDADES DE PESTAÑAS / HOJAS
# ══════════════════════════════════════════════════════════════


def fecha_inicio_pestana(nombre):
    """Extrae la fecha de inicio (lunes) de un nombre de pestaña matricial.

    Maneja variantes como:
      '01-12-25 AL 07-12-25'
      'Copia de 25-12-23 al 31-12-23'
      'COPIA 18-12-23 AL 24-12-23'
      ' 28-10-24 al 03-11-24'  (espacio al inicio)
    """
    limpio = nombre.strip()
    limpio = re.sub(r"^[Cc][Oo][Pp][Ii][Aa]\s+[Dd][Ee]\s+", "", limpio)
    limpio = re.sub(r"^[Cc][Oo][Pp][Ii][Aa]\s+", "", limpio)
    limpio = limpio.strip()

    match = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", limpio)
    if match:
        dia, mes, anio = (
            int(match.group(1)),
            int(match.group(2)),
            int(match.group(3)),
        )
        if anio < 100:
            anio += 2000
        try:
            return date(anio, mes, dia)
        except ValueError:
            return None
    return None


def sanitizar_nombre_hoja(nombre):
    """Sanitiza un nombre para usarlo como hoja de Excel (max 31 chars, sin chars inválidos)."""
    nombre = str(nombre).strip()
    nombre = _RE_CHARS_HOJA_INVALIDOS.sub("", nombre)
    return nombre[:MAX_LONGITUD_HOJA] if nombre else "EPG"


def validar_nombre_hoja(nombre):
    """Valida un nombre de hoja y devuelve (es_valido, mensaje_error).

    Reglas de Excel:
      - Máximo 31 caracteres
      - No puede contener : \\ / ? * [ ]
    """
    if not nombre or not nombre.strip():
        return False, "El nombre del canal no puede estar vacío."
    nombre = nombre.strip()
    if len(nombre) > MAX_LONGITUD_HOJA:
        return False, f"El nombre no puede superar {MAX_LONGITUD_HOJA} caracteres (tiene {len(nombre)})."
    chars_invalidos = _RE_CHARS_HOJA_INVALIDOS.findall(nombre)
    if chars_invalidos:
        return False, f"El nombre no puede contener estos caracteres: {' '.join(set(chars_invalidos))}"
    return True, ""


# ══════════════════════════════════════════════════════════════
# RESOLUCIÓN DE DÍAS RELATIVOS A FECHAS ABSOLUTAS
# ══════════════════════════════════════════════════════════════


def resolver_dia_a_fecha(texto_dia, contexto_fecha_inicio=None, contexto_mes=None, contexto_anio=None):
    """Intenta resolver una referencia textual de día a una fecha real.

    Estrategias (en orden):
      1. Si texto_dia es una fecha parseable directamente, la usa.
      2. Si es un nombre de día ("LUNES", "Martes", "Lun") y hay contexto
         de fecha de inicio de semana, calcula el offset.
      3. Si es "DD" o "DD/MM" y hay contexto de mes/año, construye la fecha.

    Returns:
        date o None
    """
    if not texto_dia or not str(texto_dia).strip():
        return None

    texto = str(texto_dia).strip().upper()
    texto = texto.replace("\u00a0", " ")

    # 1. Intento directo como fecha completa
    fecha_directa = convertir_fecha(texto_dia)
    if fecha_directa:
        return fecha_directa

    # 2. Nombre de día de la semana → offset desde fecha de inicio
    if contexto_fecha_inicio:
        # Nombre completo
        if texto in DIA_OFFSET:
            return contexto_fecha_inicio + timedelta(days=DIA_OFFSET[texto])
        # Abreviatura
        for abrev, offset in DIA_ABREVIADO.items():
            if texto == abrev or texto.startswith(abrev):
                return contexto_fecha_inicio + timedelta(days=offset)

    # 3. Solo día numérico (ej. "15") con contexto de mes/año
    match_solo_dia = re.match(r"^(\d{1,2})$", texto)
    if match_solo_dia and contexto_mes and contexto_anio:
        try:
            return date(contexto_anio, contexto_mes, int(match_solo_dia.group(1)))
        except ValueError:
            return None

    # 4. Día/Mes (ej. "15/09") con contexto de año
    match_dia_mes = re.match(r"^(\d{1,2})[/\-](\d{1,2})$", texto)
    if match_dia_mes and contexto_anio:
        try:
            return date(contexto_anio, int(match_dia_mes.group(2)), int(match_dia_mes.group(1)))
        except ValueError:
            return None

    return None


# ══════════════════════════════════════════════════════════════
# VALIDACIONES DE SANIDAD
# ══════════════════════════════════════════════════════════════


def calcular_tasa_nulos(df, columnas=None):
    """Calcula el porcentaje de valores nulos/vacíos en las columnas indicadas.

    Args:
        df: DataFrame a evaluar.
        columnas: Lista de nombres de columnas. Si es None, evalúa todas.

    Returns:
        float entre 0.0 y 1.0
    """
    if df.empty:
        return 1.0
    if columnas:
        subset = df[columnas]
    else:
        subset = df
    total_celdas = subset.shape[0] * subset.shape[1]
    if total_celdas == 0:
        return 1.0
    nulos = subset.isna().sum().sum()
    vacios = (subset.astype(str).str.strip() == "").sum().sum()
    return (nulos + vacios) / total_celdas


def detectar_anomalias(df):
    """Detecta anomalías comunes en un DataFrame de resultado EPG.

    Busca:
      - Horarios de inicio duplicados (mismo Día + misma Hora)
      - Horas fuera de rango (no HH:MM válido)
      - Fechas que no pudieron parsearse (None/vacío)

    Returns:
        Lista de strings con las anomalías encontradas.
    """
    anomalias = []

    if df.empty:
        return ["El resultado está completamente vacío."]

    # Fechas vacías
    fechas_vacias = df["Día"].isna().sum() + (df["Día"].astype(str).str.strip() == "").sum()
    if fechas_vacias > 0:
        anomalias.append(f"⚠️ {fechas_vacias} fila(s) con fecha vacía o inválida.")

    # Horas vacías
    horas_vacias = df["Hora"].isna().sum() + (df["Hora"].astype(str).str.strip() == "").sum()
    if horas_vacias > 0:
        anomalias.append(f"⚠️ {horas_vacias} fila(s) con hora vacía o inválida.")

    # Horas fuera de formato HH:MM
    patron_hora = re.compile(r"^\d{2}:\d{2}$")
    horas_malas = df["Hora"].dropna().apply(lambda x: not patron_hora.match(str(x).strip()))
    n_malas = horas_malas.sum()
    if n_malas > 0:
        anomalias.append(f"⚠️ {n_malas} hora(s) no cumplen el formato HH:MM.")

    # Horarios de inicio duplicados (mismo Día + misma Hora)
    if "Día" in df.columns and "Hora" in df.columns:
        duplicados = df.duplicated(subset=["Día", "Hora"], keep=False)
        n_dup = duplicados.sum()
        if n_dup > 0:
            anomalias.append(
                f"🔴 {n_dup} fila(s) con horario de inicio duplicado "
                f"(mismo Día + misma Hora). Verificar antes de exportar."
            )

    # Programas vacíos
    progs_vacios = df["Programa"].isna().sum() + (df["Programa"].astype(str).str.strip() == "").sum()
    if progs_vacios > 0:
        anomalias.append(f"⚠️ {progs_vacios} fila(s) con nombre de programa vacío.")

    return anomalias


def validar_rango_fechas(df, fecha_inicio_esperada=None, fecha_fin_esperada=None):
    """Valida que todas las fechas extraídas caigan dentro de un rango esperado.

    Returns:
        Lista de strings con las anomalías de rango encontradas.
    """
    avisos = []
    if df.empty or fecha_inicio_esperada is None:
        return avisos

    for i, row in df.iterrows():
        try:
            fecha = datetime.strptime(str(row["Día"]).strip(), "%m-%d-%Y").date()
            if fecha < fecha_inicio_esperada:
                avisos.append(f"Fila {i+1}: Fecha {row['Día']} es anterior al rango esperado ({formatear_fecha_salida(fecha_inicio_esperada)}).")
            elif fecha_fin_esperada and fecha > fecha_fin_esperada:
                avisos.append(f"Fila {i+1}: Fecha {row['Día']} es posterior al rango esperado ({formatear_fecha_salida(fecha_fin_esperada)}).")
        except (ValueError, TypeError):
            continue

    if avisos:
        avisos.insert(0, f"🔴 Se encontraron {len(avisos)} fecha(s) fuera del rango esperado:")

    return avisos
