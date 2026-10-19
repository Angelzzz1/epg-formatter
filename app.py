"""
📺 EPG Formatter Tool
====================
Transforma archivos Excel de programación (EPG) enviados por proveedores
al formato estricto requerido por el EPG Uploader.

Modos de procesamiento:
  📋 Lineal    → Un programa por fila (ej. Suyai, Cine Premium)
  📊 Matricial → Días como columnas, horas como filas (ej. Canal LUZ)
  ✋ Manual     → El operador mapea las columnas manualmente

Ejecutar:
  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import re
import io


# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="📺 EPG Formatter",
    page_icon="📺",
    layout="wide",
    initial_sidebar_state="expanded",
)

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


# ══════════════════════════════════════════════════════════════
# FUNCIONES DE CONVERSIÓN
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


def convertir_hora(valor):
    """Convierte cualquier representación de hora a string HH:MM.

    Soporta:
      - Fracciones de Excel (ej. 0.5 → 12:00, 0.75 → 18:00)
      - Objetos datetime / time
      - Strings con hora embebida (ej. '1/08/2024 00:54:27' → '00:54')
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
        valor = re.sub(r'\s*:\s*', ':', valor)
        match = re.search(r"(\d{1,2}):(\d{2})", valor.strip())
        if match:
            h = int(match.group(1)) % 24
            m = int(match.group(2))
            return f"{h:02d}:{m:02d}"
    return None


def limpiar_programa(texto, solo_primera_linea=False):
    """Limpia el nombre del programa: MAYÚSCULAS, sin saltos de línea, sin espacios extras.

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
    texto = re.sub(r"\s{2,}", " ", texto).strip()
    return texto.upper()


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


def formatear_fecha_salida(d):
    """Formatea un objeto date a MM-DD-YYYY con guiones."""
    if d is None:
        return None
    try:
        return d.strftime("%m-%d-%Y")
    except Exception:
        return None


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
    """Sanitiza un nombre para usarlo como hoja de Excel (max 31 chars)."""
    nombre = str(nombre).strip()
    nombre = re.sub(r'[:\\/?\*\[\]]', "", nombre)
    return nombre[:31] if nombre else "EPG"


# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE FORMATO
# ══════════════════════════════════════════════════════════════


def detectar_cabecera_y_modo(df_crudo, max_filas=5):
    """Escanea las primeras filas buscando la fila de cabecera y el modo.

    Retorna (indice_fila, modo) donde modo puede ser:
      'lineal', 'matricial' o 'desconocido'
    """
    kw_lineal = {"DIA", "DÍA", "PROGRAMA", "INICIO", "FECHA"}
    kw_matrix = {
        "HORA",
        "LUNES",
        "MARTES",
        "MIERCOLES",
        "MIÉRCOLES",
        "JUEVES",
        "VIERNES",
        "SABADO",
        "SÁBADO",
        "DOMINGO",
    }

    mejor_fila = 0
    mejor_score = 0
    mejor_modo = "desconocido"

    for i in range(min(max_filas, len(df_crudo))):
        valores = set()
        for v in df_crudo.iloc[i]:
            if pd.notna(v):
                valores.add(str(v).upper().strip())

        score_m = len(valores & kw_matrix)
        score_l = len(valores & kw_lineal)

        if score_m >= 3 and score_m > mejor_score:
            mejor_fila, mejor_score, mejor_modo = i, score_m, "matricial"
        elif score_l >= 2 and score_l > mejor_score:
            mejor_fila, mejor_score, mejor_modo = i, score_l, "lineal"

    return mejor_fila, mejor_modo


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO
# ══════════════════════════════════════════════════════════════


def procesar_lineal(df, idx_dia, idx_hora, idx_prog, offset_tz):
    """Procesa formato lineal: cada fila es un programa.

    Args:
        df:        DataFrame sin cabecera, columnas numéricas (0, 1, 2, ...)
        idx_dia:   Índice de columna para la fecha
        idx_hora:  Índice de columna para la hora
        idx_prog:  Índice de columna para el programa
        offset_tz: Horas a sumar/restar

    Returns:
        (DataFrame con columnas Día/Hora/Programa, lista de avisos)
    """
    filas = []
    avisos = []

    for i in range(len(df)):
        row = df.iloc[i]
        num = i + 2  # Fila real en Excel (1-indexed + header)

        try:
            d = convertir_fecha(row.iloc[idx_dia])
            if d is None:
                avisos.append(f"Fila {num}: Fecha inválida → `{row.iloc[idx_dia]}`")
                continue

            t = convertir_hora(row.iloc[idx_hora])
            if t is None:
                avisos.append(f"Fila {num}: Hora inválida → `{row.iloc[idx_hora]}`")
                continue

            p = limpiar_programa(row.iloc[idx_prog])
            if not p:
                avisos.append(f"Fila {num}: Programa vacío, omitido")
                continue

            t = ajustar_hora(t, offset_tz)
            filas.append(
                {
                    "_sort_date": d,
                    "Día": formatear_fecha_salida(d),
                    "Hora": t,
                    "Programa": p,
                }
            )

        except Exception as e:
            avisos.append(f"Fila {num}: Error → {e}")

    if not filas:
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    resultado = pd.DataFrame(filas)
    resultado = resultado.sort_values(["_sort_date", "Hora"])
    resultado = resultado.drop(columns=["_sort_date"]).reset_index(drop=True)
    return resultado, avisos


def procesar_matricial(df, nombre_pestana, offset_tz):
    """Procesa formato matricial: columnas = días, filas = horas.

    Pasos:
      1. Encuentra columnas HORA y días de la semana (solo primer set)
      2. Forward-fill en columnas de días (por celdas mergeadas)
      3. Unpivot: cada celda → una fila Día + Hora + Programa
      4. Elimina duplicados consecutivos (producidos por el ffill)
      5. Ordena cronológicamente
    """
    filas = []
    avisos = []

    lunes = fecha_inicio_pestana(nombre_pestana)
    if lunes is None:
        avisos.append(
            f"❌ No se pudo extraer la fecha de la pestaña: '{nombre_pestana}'"
        )
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    cols = df.columns.tolist()

    # Encontrar columna HORA y columnas de días (solo el primer set de 7)
    idx_hora = None
    cols_dia = {}
    dias_vistos = set()

    for i, col in enumerate(cols):
        c = str(col).upper().strip()
        if "HORA" in c and idx_hora is None:
            idx_hora = i
            continue
        for nombre_dia, offset in DIA_OFFSET.items():
            if nombre_dia in c and offset not in dias_vistos:
                cols_dia[i] = offset
                dias_vistos.add(offset)
                break

    if idx_hora is None:
        idx_hora = 0

    if not cols_dia:
        avisos.append("❌ No se encontraron columnas de días (LUNES, MARTES, etc.)")
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    # Forward-fill en cada columna de día (hereda el programa de arriba)
    for ci in cols_dia:
        col_name = cols[ci]
        df[col_name] = df[col_name].ffill()

    # Unpivot
    for i in range(len(df)):
        row = df.iloc[i]

        t = convertir_hora(row.iloc[idx_hora])
        if t is None:
            continue

        t = ajustar_hora(t, offset_tz)

        for ci, offset_dia in cols_dia.items():
            programa_raw = row.iloc[ci]
            if pd.isna(programa_raw) or str(programa_raw).strip() == "":
                continue

            p = limpiar_programa(programa_raw, solo_primera_linea=True)
            if not p:
                continue

            d = lunes + timedelta(days=offset_dia)
            filas.append(
                {
                    "_sort_date": d,
                    "_sort_time": t,
                    "Día": formatear_fecha_salida(d),
                    "Hora": t,
                    "Programa": p,
                }
            )

    if not filas:
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    resultado = pd.DataFrame(filas)
    resultado = resultado.sort_values(["_sort_date", "_sort_time"])

    # Eliminar duplicados consecutivos (mismo día + mismo programa seguido)
    mask = (resultado["Día"] == resultado["Día"].shift(1)) & (
        resultado["Programa"] == resultado["Programa"].shift(1)
    )
    resultado = resultado[~mask]

    resultado = (
        resultado.drop(columns=["_sort_date", "_sort_time"])
        .reset_index(drop=True)
    )
    return resultado, avisos



def procesar_bloque(df, idx_hora, idx_prog, offset_tz, mes, anio):
    import calendar
    from datetime import date
    
    # Generar todas las fechas del mes solicitado
    num_days = calendar.monthrange(anio, mes)[1]
    todas_fechas = [date(anio, mes, day) for day in range(1, num_days + 1)]
    
    filas = []
    avisos = []
    
    # Por defecto asumimos que aplica a toda la semana si no hay títulos
    dias_activos = [0, 1, 2, 3, 4, 5, 6] 
    
    for i in range(len(df)):
        row = df.iloc[i]
        num = i + 2
        
        # 1. Detección Inteligente de Contexto (leer la fila como texto)
        import re
        row_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
        row_str = re.sub(r'\s+', ' ', row_str)
        if "LUNES A VIERNES" in row_str:
            dias_activos = [0, 1, 2, 3, 4]
        elif "SABADO Y DOMINGO" in row_str or "FIN DE SEMANA" in row_str:
            dias_activos = [5, 6]
        elif "SABADO" in row_str or "SÁBADO" in row_str:
            dias_activos = [5]
        elif "DOMINGO" in row_str:
            dias_activos = [6]
        elif "LUNES A DOMINGO" in row_str:
            dias_activos = [0, 1, 2, 3, 4, 5, 6]
            
        # 2. Extracción de Hora y Programa
        try:
            val_hora = row.iloc[idx_hora] if idx_hora < len(row) else None
            val_prog = row.iloc[idx_prog] if idx_prog < len(row) else None
            
            t = convertir_hora(val_hora)
            if t is None:
                continue
                
            p = limpiar_programa(val_prog)
            if not p:
                continue
                
            t = ajustar_hora(t, offset_tz)
            
            # 3. Aplicar solo a los días activos correspondientes a este bloque
            fechas_aplicables = [d for d in todas_fechas if d.weekday() in dias_activos]
            for d in fechas_aplicables:
                filas.append({
                    "_sort_date": d,
                    "Día": formatear_fecha_salida(d),
                    "Hora": t,
                    "Programa": p,
                })
        except Exception as e:
            avisos.append(f"Fila {num}: Error → {e}")
            
    if not filas:
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos
        
    resultado = pd.DataFrame(filas)
    resultado = resultado.sort_values(["_sort_date", "Hora"])
    resultado = resultado.drop(columns=["_sort_date"]).reset_index(drop=True)
    return resultado, avisos

# ══════════════════════════════════════════════════════════════
# GENERACIÓN DE EXCEL
# ══════════════════════════════════════════════════════════════


def generar_excel(df_resultado, nombre_canal):
    """Genera el archivo Excel de salida listo para el EPG Uploader.

    Estructura:
      Col A  → Día       (MM-DD-YYYY)
      Col B  → Hora      (HH:MM)
      Col C  → Programa  (MAYÚSCULAS)
      Col D…P → Vacías   (requeridas por el uploader)
    """
    buf = io.BytesIO()
    nombre_hoja = sanitizar_nombre_hoja(nombre_canal)

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # 3 columnas de datos + 13 vacías = 16 columnas (A-P)
        export = df_resultado[["Día", "Hora", "Programa"]].copy()
        for i in range(13):
            export[f"__vacia_{i}"] = ""

        # Escribir datos sin header (se escribe manualmente abajo)
        export.to_excel(
            writer,
            sheet_name=nombre_hoja,
            index=False,
            header=False,
            startrow=1,
        )

        wb = writer.book
        ws = writer.sheets[nombre_hoja]

        # Formato de cabecera
        fmt_header = wb.add_format(
            {
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_name": "Verdana",
                "font_size": 9,
            }
        )

        # Formato de celdas de datos (texto)
        fmt_texto = wb.add_format({"font_name": "Verdana", "font_size": 9})

        # Escribir cabeceras manualmente
        headers = ["Día", "Hora", "Programa"] + [""] * 13
        for i, h in enumerate(headers):
            ws.write(0, i, h, fmt_header)

        # Anchos de columna y formato texto
        ws.set_column(0, 0, 14, fmt_texto)  # Día
        ws.set_column(1, 1, 8, fmt_texto)  # Hora
        ws.set_column(2, 2, 55, fmt_texto)  # Programa
        ws.set_column(3, 15, 5)  # D-P vacías

        # Congelar primera fila
        ws.freeze_panes(1, 0)

    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
# FUNCIÓN AUXILIAR: PREPARAR DATOS CON CABECERAS
# ══════════════════════════════════════════════════════════════


def preparar_datos(df_crudo, fila_cab, como_nombres=False):
    """Extrae los datos debajo de la fila de cabecera.

    Si como_nombres=True, asigna los valores de la fila cabecera como
    nombres de columna (necesario para modo matricial).
    """
    df_datos = df_crudo.iloc[fila_cab + 1 :].dropna(how="all").reset_index(drop=True)

    if como_nombres:
        cabeceras = []
        for j, v in enumerate(df_crudo.iloc[fila_cab]):
            if pd.notna(v) and str(v).strip():
                cabeceras.append(str(v).strip())
            else:
                cabeceras.append(f"_col_{j}")
        # Ajustar longitud
        n = len(df_datos.columns)
        if len(cabeceras) < n:
            cabeceras += [f"_extra_{i}" for i in range(n - len(cabeceras))]
        df_datos.columns = cabeceras[:n]

    return df_datos


def encontrar_columnas_lineales(df_crudo, fila_cab):
    """Busca automáticamente las columnas Día, Hora y Programa en la fila cabecera."""
    cabeceras = []
    for v in df_crudo.iloc[fila_cab]:
        if pd.notna(v):
            cabeceras.append(str(v).upper().strip())
        else:
            cabeceras.append("")

    idx_dia = idx_hora = idx_prog = None
    for i, c in enumerate(cabeceras):
        if c in ("DÍA", "DIA", "FECHA") and idx_dia is None:
            idx_dia = i
        elif c in ("HORA", "INICIO", "HORA INICIO", "HORA DE INICIO") and idx_hora is None:
            idx_hora = i
        elif c in ("PROGRAMA", "TITULO", "TÍTULO", "SHOW", "NOMBRE") and idx_prog is None:
            idx_prog = i

    return idx_dia, idx_hora, idx_prog, cabeceras


# ══════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════


def main():
    st.title("📺 EPG Formatter Tool")
    st.caption(
        "Transforma archivos EPG de proveedores al formato del EPG Uploader"
    )

    # ── Inicializar session_state ──
    if "resultado" not in st.session_state:
        st.session_state.resultado = None
        st.session_state.avisos_resultado = None
        st.session_state.canal_resultado = None
        st.session_state.archivo_id = None

    # ══════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════
    with st.sidebar:
        st.header("📂 Archivo")
        archivo = st.file_uploader(
            "Sube el Excel del proveedor",
            type=["xlsx"],
            help="Solo archivos .xlsx",
        )

        if not archivo:
            st.info("☝️ Arrastra o selecciona un archivo .xlsx para comenzar")
            st.session_state.resultado = None
            return

        # Limpiar resultados si cambió el archivo
        if st.session_state.archivo_id != archivo.file_id:
            st.session_state.resultado = None
            st.session_state.archivo_id = archivo.file_id

        # Leer archivo
        try:
            file_bytes = archivo.getvalue()
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
            hojas = xls.sheet_names
        except Exception as e:
            st.error(f"❌ Error al leer el archivo: {e}")
            return

        st.success(f"✅ {archivo.name} — {len(hojas)} hoja(s)")

        st.divider()
        st.header("⚙️ Configuración")

        # Modo
        modo = st.radio(
            "Modo de procesamiento",
            ["🔄 Automático", "📋 Lineal", "📊 Matricial", "✋ Manual", "📆 Bloque (Sin Fechas)"],
            index=0,
            help=(
                "**Auto**: detecta el formato automáticamente.\n\n"
                "**Lineal**: cada fila es un programa (Suyai, Cine Premium).\n\n"
                "**Matricial**: días como columnas (Canal LUZ).\n\n"
                "**Manual**: tú eliges qué columnas usar.\n\n"
                "**Bloque**: para horarios que no tienen columna de fecha (ej. Lunes a Viernes)."
            ),
        )

        # Selector de hoja
        if "📊" in modo:
            # Extraer fechas para ordenar las pestañas de más reciente a más antigua
            hojas_con_fecha = []
            hojas_sin_fecha = []
            for h in hojas:
                d = fecha_inicio_pestana(h)
                if d:
                    hojas_con_fecha.append((d, h))
                else:
                    hojas_sin_fecha.append(h)
            
            hojas_con_fecha.sort(key=lambda x: x[0], reverse=True)
            hojas_mostrar = [h[1] for h in hojas_con_fecha] + hojas_sin_fecha
            
            hoja = st.selectbox("Pestaña a procesar (ordenadas por más recientes)", hojas_mostrar)
        else:
            hoja = st.selectbox("Hoja del archivo", hojas)

        # Nombre del canal
        canal = st.text_input(
            "📡 Nombre del canal",
            placeholder="ej: cine Premium",
            help="Será el nombre de la hoja en el Excel de salida. "
            "Debe coincidir con el nombre en el EPG Manager.",
        )

        # Timezone
        tz = st.number_input(
            "🕐 Ajuste timezone (horas)",
            min_value=-12,
            max_value=12,
            value=0,
            step=1,
            help="Suma o resta horas a toda la columna Hora.\n"
            "Ej: -2 para pasar de GMT-3 a GMT-5",
        )

        st.divider()
        acumular = st.checkbox("➕ Acumular con el resultado anterior", help="Útil para unir varios bloques o pestañas en un solo archivo descargable.", value=False)
        procesar = st.button(
            "🚀 PROCESAR", use_container_width=True, type="primary"
        )

    # ══════════════════════════════════════════════════════════
    # MAIN AREA
    # ══════════════════════════════════════════════════════════

    # Leer la hoja seleccionada
    try:
        df_crudo = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=hoja,
            header=None,
            engine="openpyxl",
        )
        df_crudo = df_crudo.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ Error al leer la hoja '{hoja}': {e}")
        return

    if len(df_crudo) < 2:
        st.error("❌ La hoja seleccionada no tiene datos suficientes.")
        return

    # ── MODO MANUAL: UI de mapeo de columnas ──
    if "✋" in modo:
        st.subheader("✋ Modo Manual — Mapeo de columnas")

        with st.expander("📄 Vista previa de datos crudos", expanded=True):
            st.dataframe(df_crudo.head(8), use_container_width=True)

        fila_cab_manual = (
            st.number_input(
                "¿En qué fila están los títulos? (1 = primera fila)",
                min_value=1,
                max_value=min(10, len(df_crudo)),
                value=1,
                step=1,
            )
            - 1
        )

        nombres_cols = []
        for i, v in enumerate(df_crudo.iloc[fila_cab_manual]):
            if pd.notna(v) and str(v).strip():
                nombres_cols.append(f"{i}: {str(v).strip()}")
            else:
                nombres_cols.append(f"{i}: (sin nombre)")

        col1, col2, col3 = st.columns(3)
        with col1:
            sel_dia = st.selectbox("📅 Columna DÍA", nombres_cols, index=0)
        with col2:
            sel_hora = st.selectbox(
                "🕐 Columna HORA",
                nombres_cols,
                index=min(1, len(nombres_cols) - 1),
            )
        with col3:
            sel_prog = st.selectbox(
                "📺 Columna PROGRAMA",
                nombres_cols,
                index=min(3, len(nombres_cols) - 1),
            )

    # ── MODO BLOQUE: UI de mapeo y fechas ──
    if "📆" in modo:
        st.subheader("📆 Modo Bloque Inteligente (Auto-Detección)")
        st.info("💡 **Tip NOC:** El sistema leerá automáticamente si la fila dice 'LUNES A VIERNES', 'SABADO' o 'DOMINGO' y asignará los horarios correctamente para todo el mes.")
        
        with st.expander("📄 Vista previa de datos crudos", expanded=True):
            st.dataframe(df_crudo.head(80), use_container_width=True)
            
        c_mes, c_anio = st.columns(2)
        with c_mes:
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            mes_sel = st.selectbox("📅 Mes a generar", meses, index=datetime.today().month - 1)
            mes_num = meses.index(mes_sel) + 1
        with c_anio:
            anio_sel = st.number_input("📅 Año", min_value=2020, max_value=2050, value=datetime.today().year)

        nombres_cols = []
        for i, v in enumerate(df_crudo.iloc[0]):
            if pd.notna(v) and str(v).strip():
                nombres_cols.append(f"{i}: {str(v).strip()}")
            else:
                nombres_cols.append(f"{i}: (sin nombre)")
                
        c1, c2 = st.columns(2)
        with c1:
            sel_hora_b = st.selectbox("🕐 Columna HORA", nombres_cols, index=min(1, len(nombres_cols)-1) if len(nombres_cols)>1 else 0)
        with c2:
            sel_prog_b = st.selectbox("📺 Columna PROGRAMA", nombres_cols, index=min(2, len(nombres_cols)-1) if len(nombres_cols)>2 else 0)

    # ── Mostrar preview si aún no se ha procesado ──
    if not procesar and st.session_state.resultado is None:
        if "✋" not in modo and "📆" not in modo:
            st.subheader("👁️ Vista previa del archivo")
            st.dataframe(df_crudo.head(10), use_container_width=True)
            st.info("👈 Configura los parámetros y presiona **PROCESAR**")
        return

    # ══════════════════════════════════════════════════════════
    # EJECUTAR PROCESAMIENTO
    # ══════════════════════════════════════════════════════════
    if procesar:
        # Validar nombre del canal
        if not canal or not canal.strip():
            st.error("❌ Ingresa el nombre del canal en la barra lateral.")
            return

        try:
            with st.spinner("⏳ Procesando..."):

                # ── MODO MANUAL ──
                if "✋" in modo:
                    idx_d = int(sel_dia.split(":")[0])
                    idx_h = int(sel_hora.split(":")[0])
                    idx_p = int(sel_prog.split(":")[0])
                    df_datos = preparar_datos(df_crudo, fila_cab_manual)
                    df_result, avisos = procesar_lineal(
                        df_datos, idx_d, idx_h, idx_p, tz
                    )

                # ── MODO BLOQUE ──
                elif "📆" in modo:
                    idx_h = int(sel_hora_b.split(":")[0])
                    idx_p = int(sel_prog_b.split(":")[0])
                    # Usamos df_crudo directamente sin ffill() para no ensuciar los títulos
                    df_result, avisos = procesar_bloque(df_crudo, idx_h, idx_p, tz, mes_num, anio_sel)

                # ── MODO MATRICIAL ──
                elif "📊" in modo:
                    fila_cab, _ = detectar_cabecera_y_modo(df_crudo)
                    df_datos = preparar_datos(df_crudo, fila_cab, como_nombres=True)
                    df_result, avisos = procesar_matricial(df_datos, hoja, tz)

                # ── MODO LINEAL ──
                elif "📋" in modo:
                    fila_cab, _ = detectar_cabecera_y_modo(df_crudo)
                    idx_dia, idx_hora, idx_prog, cabeceras = (
                        encontrar_columnas_lineales(df_crudo, fila_cab)
                    )

                    if None in (idx_dia, idx_hora, idx_prog):
                        st.error(
                            f"❌ No se encontraron las columnas necesarias.\n\n"
                            f"Columnas encontradas: `{cabeceras}`\n\n"
                            f"Se necesitan: DÍA, HORA/INICIO, PROGRAMA\n\n"
                            f"💡 Usa el modo **Manual**."
                        )
                        return

                    st.info(
                        f"📋 Día: `{cabeceras[idx_dia]}` | "
                        f"Hora: `{cabeceras[idx_hora]}` | "
                        f"Programa: `{cabeceras[idx_prog]}`"
                    )

                    df_datos = preparar_datos(df_crudo, fila_cab)
                    df_result, avisos = procesar_lineal(
                        df_datos, idx_dia, idx_hora, idx_prog, tz
                    )

                # ── MODO AUTOMÁTICO ──
                else:
                    fila_cab, modo_det = detectar_cabecera_y_modo(df_crudo)

                    if modo_det == "desconocido":
                        st.error(
                            "❌ No se pudo detectar el formato automáticamente.\n\n"
                            "💡 Usa **Lineal**, **Matricial** o **Manual**."
                        )
                        return

                    st.info(
                        f"🔍 Formato detectado: **{modo_det.upper()}** "
                        f"(cabecera en fila {fila_cab + 1})"
                    )

                    if modo_det == "matricial":
                        df_datos = preparar_datos(
                            df_crudo, fila_cab, como_nombres=True
                        )
                        df_result, avisos = procesar_matricial(
                            df_datos, hoja, tz
                        )
                    else:
                        idx_dia, idx_hora, idx_prog, cabeceras = (
                            encontrar_columnas_lineales(df_crudo, fila_cab)
                        )
                        if None in (idx_dia, idx_hora, idx_prog):
                            st.error(
                                f"❌ Formato lineal detectado pero faltan columnas.\n\n"
                                f"Columnas: `{cabeceras}`\n\n"
                                f"💡 Usa el modo **Manual**."
                            )
                            return
                        df_datos = preparar_datos(df_crudo, fila_cab)
                        df_result, avisos = procesar_lineal(
                            df_datos, idx_dia, idx_hora, idx_prog, tz
                        )

            # Guardar en session_state
            if not df_result.empty:
                if acumular and st.session_state.resultado is not None and not st.session_state.resultado.empty:
                    df_result = pd.concat([st.session_state.resultado, df_result])
                    df_result['_dt'] = pd.to_datetime(df_result['Día'], format='%m-%d-%Y', errors='coerce')
                    df_result = df_result.sort_values(['_dt', 'Hora']).drop(columns=['_dt']).reset_index(drop=True)
                    if st.session_state.avisos_resultado:
                        avisos = st.session_state.avisos_resultado + avisos
                        
            st.session_state.resultado = df_result
            st.session_state.avisos_resultado = avisos
            st.session_state.canal_resultado = canal

        except Exception as e:
            st.error(
                f"❌ Error inesperado durante el procesamiento:\n\n`{e}`\n\n"
                f"Verifica que el archivo tenga el formato esperado."
            )
            return

    # ══════════════════════════════════════════════════════════
    # MOSTRAR RESULTADOS
    # ══════════════════════════════════════════════════════════
    if st.session_state.resultado is not None:
        df_res = st.session_state.resultado
        avisos = st.session_state.avisos_resultado or []
        canal_out = st.session_state.canal_resultado or canal

        if df_res.empty:
            st.error("❌ No se generaron resultados. Revisa los avisos abajo.")
        else:
            # Métricas
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("✅ Filas procesadas", f"{len(df_res):,}")
            c2.metric("⚠️ Avisos", len(avisos))
            c3.metric("📅 Días cubiertos", df_res["Día"].nunique())

            # Vista previa
            st.subheader("👁️ Vista previa del resultado (primeras 20 filas)")
            st.dataframe(df_res.head(20), use_container_width=True, height=400)

            # Botón de descarga
            excel_bytes = generar_excel(df_res, canal_out)
            nombre_archivo = f"{canal_out.strip().replace(' ', '_')}_EPG.xlsx"

            st.download_button(
                label=f"📥 Descargar {nombre_archivo}",
                data=excel_bytes,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        # Avisos
        if avisos:
            with st.expander(
                f"⚠️ Avisos y errores ({len(avisos)})", expanded=False
            ):
                for a in avisos[:200]:
                    st.write(a)
                if len(avisos) > 200:
                    st.caption(f"... y {len(avisos) - 200} avisos más")


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    main()
