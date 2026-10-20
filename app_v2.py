"""
📺 EPG Formatter Tool v2
========================
Transforma archivos Excel de programación (EPG) enviados por proveedores
al formato estricto requerido por el EPG Uploader.

Modos de procesamiento:
  🔄 Automático    → Detecta si el formato es lineal o matricial
  📋 Lineal        → Un programa por fila (ej. Suyai, Cine Premium, Nexxor)
  📊 Matricial     → Días como columnas, horas como filas (ej. Canal LUZ)
  📆 Bloque        → Horarios sin fechas, con títulos incrustados (ej. Antena 1)
  ✋ Manual        → El operador mapea las columnas manualmente

Ejecutar:
  streamlit run app_v2.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import re
import io
import calendar

# Importar módulo compartido de normalización
from normalization import (
    EXCEL_EPOCH,
    DIA_OFFSET,
    convertir_fecha,
    convertir_hora,
    limpiar_programa,
    ajustar_hora,
    formatear_fecha_salida,
    fecha_inicio_pestana,
    sanitizar_nombre_hoja,
    validar_nombre_hoja,
    calcular_tasa_nulos,
    detectar_anomalias,
    validar_rango_fechas,
)

# Preprocesamiento universal de Excel (desfusión de celdas combinadas)
from excel_preprocessing import expandir_celdas_combinadas, obtener_hojas

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="📺 EPG Formatter v2",
    page_icon="📺",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE FORMATO
# ══════════════════════════════════════════════════════════════


def _contiene_palabra(texto, palabra):
    """Verifica si 'palabra' aparece como palabra completa en 'texto'.

    Usa word boundaries (\b) para evitar falsos positivos como
    'DIA' coincidiendo con 'DIARIO' o 'CALENDARIO'.
    """
    return bool(re.search(r"\b" + re.escape(palabra) + r"\b", texto))


def detectar_cabecera_y_modo(df_crudo, max_filas=10):
    """Escanea las primeras filas buscando la fila de cabecera y el modo.

    Retorna (indice_fila, modo) donde modo puede ser:
      'lineal', 'matricial' o 'desconocido'

    Usa coincidencia de palabras completas (word boundary) para evitar
    que títulos fusionados como 'CALENDARIO DIARIO' den falsos positivos
    con la keyword 'DIA'.
    """
    kw_lineal = ["DIA", "DÍA", "PROGRAMA", "INICIO", "FECHA", "TÍTULO", "TITULO"]
    kw_matrix = [
        "HORA", "LUNES", "MARTES", "MIERCOLES", "MIÉRCOLES",
        "JUEVES", "VIERNES", "SABADO", "SÁBADO", "DOMINGO",
    ]

    mejor_fila = 0
    mejor_score = 0
    mejor_modo = "desconocido"

    for i in range(min(max_filas, len(df_crudo))):
        valores = []
        for v in df_crudo.iloc[i]:
            if pd.notna(v):
                valores.append(str(v).upper().strip())

        # Descartar filas donde todos los valores son idénticos
        # (señal de una celda fusionada de título, no una cabecera real)
        if len(set(valores)) <= 1:
            continue

        # Contar cuántas columnas ÚNICAS tienen alguna palabra clave completa
        score_m = sum(1 for v in valores if any(_contiene_palabra(v, kw) for kw in kw_matrix))
        score_l = sum(1 for v in valores if any(_contiene_palabra(v, kw) for kw in kw_lineal))

        if score_m >= 3 and score_m > mejor_score:
            mejor_fila, mejor_score, mejor_modo = i, score_m, "matricial"
        elif score_l >= 2 and score_l > mejor_score:
            mejor_fila, mejor_score, mejor_modo = i, score_l, "lineal"

    return mejor_fila, mejor_modo


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO FAST TRACK
# ══════════════════════════════════════════════════════════════


def procesar_lineal(df, idx_dia, idx_hora, idx_prog, offset_tz):
    """Procesa formato lineal: cada fila es un programa."""
    filas = []
    avisos = []

    for i in range(len(df)):
        row = df.iloc[i]
        num = i + 2

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


def procesar_matricial(df, nombre_pestana, offset_tz):
    """Procesa formato matricial: columnas = días, filas = horas.

    NOTA: Las celdas combinadas ya fueron expandidas por
    excel_preprocessing.expandir_celdas_combinadas() antes de llegar aquí.
    No se usa ffill() — los valores reales ya están propagados desde
    los metadatos del archivo Excel.
    """
    filas = []
    avisos = []

    lunes = fecha_inicio_pestana(nombre_pestana)
    if lunes is None:
        avisos.append(f"❌ No se pudo extraer la fecha de la pestaña: '{nombre_pestana}'")
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    cols = df.columns.tolist()
    idx_hora = None
    cols_dia = {}
    dias_vistos = set()

    for i, col in enumerate(cols):
        c = str(col).upper().strip()
        # Tomar solo la PRIMERA columna HORA; ignorar duplicadas (ej. columna I)
        if "HORA" in c:
            if idx_hora is None:
                idx_hora = i
            # Las columnas HORA adicionales se ignoran silenciosamente
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

    # SIN ffill(): las celdas combinadas ya fueron expandidas por
    # excel_preprocessing. Una celda vacía aquí es un hueco REAL
    # en la programación, no un merge sin propagar.

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
            filas.append({
                "_sort_date": d,
                "_sort_time": t,
                "Día": formatear_fecha_salida(d),
                "Hora": t,
                "Programa": p,
            })

    if not filas:
        return pd.DataFrame(columns=["Día", "Hora", "Programa"]), avisos

    resultado = pd.DataFrame(filas)
    resultado = resultado.sort_values(["_sort_date", "_sort_time"])
    mask = (resultado["Día"] == resultado["Día"].shift(1)) & (
        resultado["Programa"] == resultado["Programa"].shift(1)
    )
    resultado = resultado[~mask]
    resultado = resultado.drop(columns=["_sort_date", "_sort_time"]).reset_index(drop=True)
    return resultado, avisos


def procesar_bloque(df, idx_hora, idx_prog, offset_tz, mes, anio):
    """Procesa formato bloque: horarios sin fecha, con títulos de contexto."""
    num_days = calendar.monthrange(anio, mes)[1]
    todas_fechas = [date(anio, mes, day) for day in range(1, num_days + 1)]

    filas = []
    avisos = []
    dias_activos = [0, 1, 2, 3, 4, 5, 6]

    for i in range(len(df)):
        row = df.iloc[i]
        num = i + 2

        row_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
        row_str = re.sub(r'\s+', ' ', row_str)
        if "LUNES A DOMINGO" in row_str:
            dias_activos = [0, 1, 2, 3, 4, 5, 6]
        elif "LUNES A VIERNES" in row_str:
            dias_activos = [0, 1, 2, 3, 4]
        elif "SABADO Y DOMINGO" in row_str or "FIN DE SEMANA" in row_str:
            dias_activos = [5, 6]
        else:
            match = re.search(r'\b(LUNES|MARTES|MIERCOLES|MIÉRCOLES|JUEVES|VIERNES|SABADO|SÁBADO|DOMINGO)\b', row_str)
            if match:
                dia_str = match.group(1).replace('É', 'E').replace('Á', 'A')
                mapa = {"LUNES": 0, "MARTES": 1, "MIERCOLES": 2, "JUEVES": 3, "VIERNES": 4, "SABADO": 5, "DOMINGO": 6}
                dias_activos = [mapa[dia_str]]

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
# FUNCIONES AUXILIARES
# ══════════════════════════════════════════════════════════════


def preparar_datos(df_crudo, fila_cab, como_nombres=False):
    """Extrae los datos debajo de la fila de cabecera."""
    df_datos = df_crudo.iloc[fila_cab + 1:].dropna(how="all").reset_index(drop=True)

    if como_nombres:
        cabeceras = []
        for j, v in enumerate(df_crudo.iloc[fila_cab]):
            if pd.notna(v) and str(v).strip():
                cabeceras.append(str(v).strip())
            else:
                cabeceras.append(f"_col_{j}")
        n = len(df_datos.columns)
        if len(cabeceras) < n:
            cabeceras += [f"_extra_{i}" for i in range(n - len(cabeceras))]
        df_datos.columns = cabeceras[:n]

    return df_datos


def encontrar_columnas_lineales(df_crudo, fila_cab):
    """Busca automáticamente las columnas Día, Hora y Programa en la fila cabecera.

    Usa coincidencia de palabras completas (word boundary) para evitar
    falsos positivos con subcadenas.
    """
    cabeceras = []
    for v in df_crudo.iloc[fila_cab]:
        if pd.notna(v):
            cabeceras.append(str(v).upper().strip())
        else:
            cabeceras.append("")

    idx_dia = idx_hora = idx_prog = None
    for i, c in enumerate(cabeceras):
        if idx_dia is None and any(_contiene_palabra(c, kw) for kw in ["DÍA", "DIA", "FECHA"]):
            idx_dia = i
        elif idx_hora is None and any(_contiene_palabra(c, kw) for kw in ["HORA", "INICIO"]):
            idx_hora = i
        elif idx_prog is None and any(_contiene_palabra(c, kw) for kw in ["PROGRAMA", "TITULO", "TÍTULO", "SHOW", "NOMBRE"]):
            idx_prog = i

    return idx_dia, idx_hora, idx_prog, cabeceras


# ══════════════════════════════════════════════════════════════
# GENERACIÓN DE EXCEL (TEXTO PURO)
# ══════════════════════════════════════════════════════════════


def generar_excel(df_resultado, nombre_canal):
    """Genera el archivo Excel de salida listo para el EPG Uploader.

    IMPORTANTE: Las celdas de Día y Hora se escriben como TEXTO PURO
    (write_string) para evitar que Excel auto-formatee según la
    configuración regional de cada PC.

    Estructura:
      Col A  → Día       (MM-DD-YYYY, texto)
      Col B  → Hora      (HH:MM, texto)
      Col C  → Programa  (MAYÚSCULAS, texto)
      Col D…P → Vacías   (requeridas por el uploader)
    """
    buf = io.BytesIO()
    nombre_hoja = sanitizar_nombre_hoja(nombre_canal)

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # Crear hoja vacía para escribir manualmente como texto
        pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)

        wb = writer.book
        ws = writer.sheets[nombre_hoja]

        # Formato de cabecera
        fmt_header = wb.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_name": "Verdana",
            "font_size": 9,
        })

        # Formato de texto (forzar texto, no número/fecha)
        fmt_texto = wb.add_format({
            "font_name": "Verdana",
            "font_size": 9,
            "num_format": "@",  # Formato texto explícito
        })

        # Escribir cabeceras (A-P = 16 columnas)
        headers = ["Día", "Hora", "Programa"] + [""] * 13
        for i, h in enumerate(headers):
            ws.write(0, i, h, fmt_header)

        # Escribir datos como texto puro (write_string)
        for row_idx in range(len(df_resultado)):
            ws.write_string(row_idx + 1, 0, str(df_resultado.iloc[row_idx]["Día"]), fmt_texto)
            ws.write_string(row_idx + 1, 1, str(df_resultado.iloc[row_idx]["Hora"]), fmt_texto)
            ws.write_string(row_idx + 1, 2, str(df_resultado.iloc[row_idx]["Programa"]), fmt_texto)

        # Anchos de columna
        ws.set_column(0, 0, 14, fmt_texto)   # Día
        ws.set_column(1, 1, 8, fmt_texto)    # Hora
        ws.set_column(2, 2, 55, fmt_texto)   # Programa
        ws.set_column(3, 15, 5)              # D-P vacías

        # Desactivar advertencia de "número guardado como texto"
        ws.ignore_errors({"number_stored_as_text": "A1:B65536"})

        # Congelar primera fila
        ws.freeze_panes(1, 0)

    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
# UMBRAL DE SANITY CHECK
# ══════════════════════════════════════════════════════════════
UMBRAL_MAX_NULOS = 0.15  # Si más del 15% de datos son nulos, se considera formato no reconocido


# ══════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════


def main():
    st.title("📺 EPG Formatter v2")
    st.caption(
        "Transforma archivos EPG de proveedores al formato del "
        "[EPG Uploader](http://198.71.56.147:8080/epgtool) | "
        "[EPG Manager Canales](http://cdn.oneplaychile.cl:60001/manager/channels)"
    )

    # ── Inicializar session_state ──
    if "resultado" not in st.session_state:
        st.session_state.resultado = None
        st.session_state.avisos_resultado = None
        st.session_state.canal_resultado = None
        st.session_state.archivo_id = None
        st.session_state.anomalias = None
        st.session_state.via_usada = None

    # ══════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════
    with st.sidebar:
        st.subheader("🔗 Enlaces Rápidos NOC")
        col_l1, col_l2 = st.columns(2)
        with col_l1:
            st.link_button(
                "📤 Uploader",
                "http://198.71.56.147:8080/epgtool",
                use_container_width=True,
                help="Subir el Excel generado al EPG Uploader Tool",
            )
        with col_l2:
            st.link_button(
                "📺 Canales",
                "http://cdn.oneplaychile.cl:60001/manager/channels",
                use_container_width=True,
                help="Ver listado de canales en EPG Manager",
            )
        st.divider()
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
            st.session_state.anomalias = None
            st.session_state.via_usada = None
            st.session_state.archivo_id = archivo.file_id

        # Leer archivo (solo nombres de hojas — lectura ligera)
        try:
            file_bytes = archivo.getvalue()
            hojas = obtener_hojas(file_bytes)
        except Exception as e:
            st.error(f"❌ Error al leer el archivo: {e}")
            return

        st.success(f"✅ {archivo.name} — {len(hojas)} hoja(s)")

        st.divider()
        st.header("⚙️ Configuración")

        # Modo
        modo = st.radio(
            "Modo de procesamiento",
            [
                "🔄 Automático",
                "📋 Lineal",
                "📊 Matricial",
                "📆 Bloque (Sin Fechas)",
                "✋ Manual",
            ],
            index=0,
            help=(
                "**Auto**: detecta el formato automáticamente (Lineal o Matricial).\n\n"
                "**Lineal**: cada fila es un programa (Suyai, Cine Premium, Nexxor).\n\n"
                "**Matricial**: días como columnas (Canal LUZ).\n\n"
                "**Bloque**: horarios sin columna de fecha (ej. Antena 1).\n\n"
                "**Manual**: tú eliges qué columnas usar."
            ),
        )

        # Selector de hoja
        if "📊" in modo:
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
            hoja = st.selectbox("Pestaña a procesar (más recientes arriba)", hojas_mostrar)
        else:
            hoja = st.selectbox("Hoja del archivo", hojas)

        # Nombre del canal (con validación)
        canal = st.text_input(
            "📡 Nombre del canal",
            placeholder="ej: cine Premium",
            help="Será el nombre de la hoja en el Excel de salida. "
            "Debe coincidir con el nombre en el EPG Manager. "
            "Máx 31 caracteres, sin : \\ / ? * [ ]",
        )
        st.caption("🔍 Consulta el listado oficial en [EPG Manager Canales](http://cdn.oneplaychile.cl:60001/manager/channels)")

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

        # ── Parámetros extra para Modo Bloque ──
        mostrar_contexto_fecha = "📆" in modo
        if mostrar_contexto_fecha:
            st.divider()
            st.subheader("📅 Contexto de Fechas")
            st.caption("Ayuda a resolver días como 'LUNES' a una fecha real.")

            c_mes, c_anio = st.columns(2)
            with c_mes:
                meses = [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
                ]
                mes_sel = st.selectbox("Mes", meses, index=datetime.today().month - 1)
                mes_num = meses.index(mes_sel) + 1
            with c_anio:
                anio_sel = st.number_input(
                    "Año", min_value=2020, max_value=2050, value=datetime.today().year
                )
        else:
            mes_num = datetime.today().month
            anio_sel = datetime.today().year

        st.divider()
        acumular = st.checkbox(
            "➕ Acumular con resultado anterior",
            help="Útil para unir varios bloques o pestañas en un solo archivo.",
            value=False,
        )
        procesar = st.button("🚀 PROCESAR", use_container_width=True, type="primary")

    # ══════════════════════════════════════════════════════════
    # MAIN AREA
    # ══════════════════════════════════════════════════════════

    # ── PASO UNIVERSAL: Leer hoja con desfusión de celdas combinadas ──
    # Este es el PRIMER paso del pipeline, antes de decidir qué vía usar.
    # Resuelve automáticamente merges horizontales, verticales y 2D.
    try:
        df_crudo, n_merges = expandir_celdas_combinadas(file_bytes, nombre_hoja=hoja)
        if n_merges > 0:
            st.info(
                f"🔗 Se desfusionaron **{n_merges}** rango(s) de celdas combinadas "
                f"en la hoja '{hoja}'. Los valores fueron propagados correctamente."
            )
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
            ) - 1
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
                "🕐 Columna HORA", nombres_cols,
                index=min(1, len(nombres_cols) - 1),
            )
        with col3:
            sel_prog = st.selectbox(
                "📺 Columna PROGRAMA", nombres_cols,
                index=min(3, len(nombres_cols) - 1),
            )

    # ── MODO BLOQUE: UI de mapeo y fechas ──
    if "📆" in modo:
        st.subheader("📆 Modo Bloque Inteligente (Auto-Detección)")
        st.info(
            "💡 **Tip NOC:** El sistema leerá automáticamente si la fila dice "
            "'LUNES A VIERNES', 'SABADO' o 'DOMINGO'."
        )

        with st.expander("📄 Vista previa de datos crudos", expanded=True):
            st.dataframe(df_crudo.head(80), use_container_width=True)

        fila_cab_bloque = (
            st.number_input(
                "¿En qué fila están los títulos? (1 = primera fila)",
                min_value=1,
                max_value=min(10, len(df_crudo)),
                value=2 if pd.isna(df_crudo.iloc[0, 2] if len(df_crudo.columns) > 2 else None) else 1,
                step=1,
                key="fila_cab_bloque"
            ) - 1
        )

        nombres_cols = []
        for i, v in enumerate(df_crudo.iloc[fila_cab_bloque]):
            if pd.notna(v) and str(v).strip():
                nombres_cols.append(f"{i}: {str(v).strip()}")
            else:
                nombres_cols.append(f"{i}: (sin nombre)")

        c1, c2 = st.columns(2)
        with c1:
            sel_hora_b = st.selectbox(
                "🕐 Columna HORA", nombres_cols,
                index=min(1, len(nombres_cols) - 1) if len(nombres_cols) > 1 else 0,
            )
        with c2:
            sel_prog_b = st.selectbox(
                "📺 Columna PROGRAMA", nombres_cols,
                index=min(2, len(nombres_cols) - 1) if len(nombres_cols) > 2 else 0,
            )

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

        valido, msg_error = validar_nombre_hoja(canal)
        if not valido:
            st.error(f"❌ Nombre de canal inválido: {msg_error}")
            return

        try:
            with st.spinner("⏳ Procesando..."):

                # ── MODO MANUAL ──
                if "✋" in modo:
                    via_usada = "Manual"
                    idx_d = int(sel_dia.split(":")[0])
                    idx_h = int(sel_hora.split(":")[0])
                    idx_p = int(sel_prog.split(":")[0])
                    df_datos = preparar_datos(df_crudo, fila_cab_manual)
                    df_result, avisos = procesar_lineal(df_datos, idx_d, idx_h, idx_p, tz)

                # ── MODO BLOQUE ──
                elif "📆" in modo:
                    via_usada = "Bloque"
                    idx_h = int(sel_hora_b.split(":")[0])
                    idx_p = int(sel_prog_b.split(":")[0])
                    df_result, avisos = procesar_bloque(df_crudo, idx_h, idx_p, tz, mes_num, anio_sel)

                # ── MODO MATRICIAL ──
                elif "📊" in modo:
                    via_usada = "Matricial"
                    fila_cab, _ = detectar_cabecera_y_modo(df_crudo)
                    df_datos = preparar_datos(df_crudo, fila_cab, como_nombres=True)
                    df_result, avisos = procesar_matricial(df_datos, hoja, tz)

                # ── MODO LINEAL ──
                elif "📋" in modo:
                    via_usada = "Lineal"
                    fila_cab, _ = detectar_cabecera_y_modo(df_crudo)
                    idx_dia, idx_hora, idx_prog, cabeceras = (
                        encontrar_columnas_lineales(df_crudo, fila_cab)
                    )

                    if None in (idx_dia, idx_hora, idx_prog):
                        st.error(
                            f"❌ No se encontraron las columnas necesarias.\n\n"
                            f"Columnas encontradas: `{cabeceras}`\n\n"
                            f"Se necesitan: DÍA, HORA/INICIO, PROGRAMA\n\n"
                            f"💡 Revisa el archivo, usa el modo **Manual** o avisa al administrador."
                        )
                        return

                    st.info(
                        f"📋 Día: `{cabeceras[idx_dia]}` | "
                        f"Hora: `{cabeceras[idx_hora]}` | "
                        f"Programa: `{cabeceras[idx_prog]}`"
                    )

                    df_datos = preparar_datos(df_crudo, fila_cab)
                    df_result, avisos = procesar_lineal(df_datos, idx_dia, idx_hora, idx_prog, tz)

                # ── MODO AUTOMÁTICO ──
                else:
                    fila_cab, modo_det = detectar_cabecera_y_modo(df_crudo)

                    if modo_det == "desconocido":
                        st.error(
                            "❌ Formato no reconocido, revisa el archivo o avisa al administrador.\n\n"
                            "💡 Si el archivo tiene una estructura conocida, intenta seleccionando directamente "
                            "el modo **Lineal**, **Matricial**, **Bloque** o **Manual** en la barra lateral."
                        )
                        return

                    elif modo_det == "matricial":
                        via_usada = "Automático (Matricial)"
                        st.info(
                            f"🔍 Formato detectado: **MATRICIAL** "
                            f"(cabecera en fila {fila_cab + 1})"
                        )
                        df_datos = preparar_datos(df_crudo, fila_cab, como_nombres=True)
                        df_result, avisos = procesar_matricial(df_datos, hoja, tz)
                    else:
                        via_usada = "Automático (Lineal)"
                        st.info(
                            f"🔍 Formato detectado: **LINEAL** "
                            f"(cabecera en fila {fila_cab + 1})"
                        )
                        idx_dia, idx_hora, idx_prog, cabeceras = (
                            encontrar_columnas_lineales(df_crudo, fila_cab)
                        )
                        if None in (idx_dia, idx_hora, idx_prog):
                            st.error(
                                f"❌ Formato no reconocido, revisa el archivo o avisa al administrador.\n\n"
                                f"Se detectó formato lineal pero faltan columnas requeridas.\n"
                                f"Columnas encontradas: `{cabeceras}`\n\n"
                                f"💡 Intenta usar el modo **Manual** para mapear las columnas manualmente."
                            )
                            return
                        else:
                            df_datos = preparar_datos(df_crudo, fila_cab)
                            df_result, avisos = procesar_lineal(
                                df_datos, idx_dia, idx_hora, idx_prog, tz
                            )

                    # ── SANITY CHECK: Validar calidad del procesamiento ──
                    if not df_result.empty:
                        tasa = calcular_tasa_nulos(df_result, ["Día", "Hora", "Programa"])
                        if tasa > UMBRAL_MAX_NULOS:
                            st.error(
                                f"❌ Formato no reconocido, revisa el archivo o avisa al administrador.\n\n"
                                f"El procesamiento generó un {tasa:.0%} de datos vacíos o nulos, superando el límite permitido (15%)."
                            )
                            return

            # ── Validación final: nunca exportar vacío ──
            if df_result.empty:
                st.error(
                    "❌ No se pudo extraer información válida de este archivo. "
                    "Revisa el formato o contacta soporte."
                )
                return

            # Detectar anomalías
            anomalias = detectar_anomalias(df_result)

            # Acumular si corresponde
            if acumular and st.session_state.resultado is not None and not st.session_state.resultado.empty:
                df_result = pd.concat([st.session_state.resultado, df_result])
                df_result['_dt'] = pd.to_datetime(df_result['Día'], format='%m-%d-%Y', errors='coerce')
                df_result = df_result.sort_values(['_dt', 'Hora']).drop(columns=['_dt']).reset_index(drop=True)
                if st.session_state.avisos_resultado:
                    avisos = st.session_state.avisos_resultado + avisos
                anomalias = detectar_anomalias(df_result)

            st.session_state.resultado = df_result
            st.session_state.avisos_resultado = avisos
            st.session_state.canal_resultado = canal
            st.session_state.anomalias = anomalias
            st.session_state.via_usada = via_usada

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
        anomalias = st.session_state.anomalias or []
        via = st.session_state.via_usada or "Fast Track"

        if df_res.empty:
            st.error(
                "❌ No se pudo extraer información válida de este archivo. "
                "Revisa el formato o contacta soporte."
            )
        else:
            # Indicador de modo usado
            st.divider()
            st.success(f"🚀 Procesado con: **{via}**")

            # Métricas
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("✅ Filas", f"{len(df_res):,}")
            c2.metric("⚠️ Avisos", len(avisos))
            c3.metric("📅 Días", df_res["Día"].nunique())
            c4.metric("🔴 Anomalías", len(anomalias))

            # Anomalías
            if anomalias:
                with st.expander(f"🔴 Anomalías detectadas ({len(anomalias)})", expanded=True):
                    for a in anomalias:
                        st.warning(a)

            # Editor interactivo (Human-in-the-Loop)
            st.subheader("✏️ Revisa y edita antes de exportar")
            st.caption(
                "Puedes hacer doble clic en cualquier celda para corregirla. "
                "Los cambios se reflejarán en el archivo descargado."
            )

            df_editado = st.data_editor(
                df_res,
                use_container_width=True,
                height=400,
                num_rows="dynamic",
                column_config={
                    "Día": st.column_config.TextColumn("Día", help="Formato: MM-DD-YYYY"),
                    "Hora": st.column_config.TextColumn("Hora", help="Formato: HH:MM"),
                    "Programa": st.column_config.TextColumn("Programa", help="Título en MAYÚSCULAS"),
                },
            )

            # Botón de descarga (usa los datos editados)
            excel_bytes = generar_excel(df_editado, canal_out)
            nombre_archivo = f"{canal_out.strip().replace(' ', '_')}_EPG.xlsx"

            col_down, col_up, col_mgr = st.columns([2, 1, 1])
            with col_down:
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo}",
                    data=excel_bytes,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with col_up:
                st.link_button(
                    label="📤 EPG Uploader",
                    url="http://198.71.56.147:8080/epgtool",
                    use_container_width=True,
                    help="Abre el EPG Uploader para subir el archivo descargado",
                )
            with col_mgr:
                st.link_button(
                    label="📺 EPG Manager",
                    url="http://cdn.oneplaychile.cl:60001/manager/channels",
                    use_container_width=True,
                    help="Ver listado de canales y programación en EPG Manager",
                )

        # Avisos
        if avisos:
            with st.expander(f"⚠️ Avisos y errores ({len(avisos)})", expanded=False):
                for a in avisos[:200]:
                    st.write(a)
                if len(avisos) > 200:
                    st.caption(f"... y {len(avisos) - 200} avisos más")


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    main()
