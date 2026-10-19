# EPG Formatter Tool v2

Herramienta ETL (*Extract, Transform, Load*) con interfaz web interactiva en Streamlit para el **NOC (Network Operations Center)**. Automatiza la ingesta, limpieza y estandarización de Guías Electrónicas de Programación (**EPG**) enviadas por proveedores de televisión, exportándolas en el formato estricto requerido por el sistema **EPG Uploader / Manager**.

---

## 🚀 Características Principales

- **Preprocesamiento de Celdas Combinadas:** Desfusiona automáticamente celdas unidas horizontalmente (programas compartidos entre días), verticalmente (programas de varias horas) o en bloques 2D mediante `openpyxl`.
- **Detección Automática Inteligente:** Identifica si el archivo es lineal o matricial usando coincidencia exacta de palabras clave (`\b`) y descartando títulos combinados.
- **Múltiples Modos de Procesamiento:**
  - 🔄 **Automático:** Detecta el modo idóneo entre lineal y matricial.
  - 📋 **Lineal:** Una fila por programa (ej. Suyai, Cine Premium, Nexxor TV).
  - 📊 **Matricial:** Días en columnas y horas en filas (ej. Canal LUZ).
  - 📆 **Bloque:** Horarios apilados con títulos incrustados y sin columna de fecha (ej. Antena 1).
  - ✋ **Manual:** Selección personalizada de columnas por el operador.
- **Normalización Estricta:**
  - Limpieza de metadata basura en nombres de programas (`(R)`, presentadores, directores, sinopsis).
  - Parseo tolerante de formatos de hora (`7.30 pm`, `20: 00` -> `20:00`).
  - Fechas exportadas en formato de texto plano `MM-DD-YYYY`.
- **Sanity Checks y Calidad:** Si un procesamiento genera más del 15% de campos nulos o vacíos, la aplicación detiene el proceso y muestra una alerta clara: *"Formato no reconocido, revisa el archivo o avisa al administrador"*.
- **Human-in-the-Loop:** Tabla interactiva editable (`st.data_editor`) para corregir datos antes de la descarga final.
- **Control de Zona Horaria:** Ajuste dinámico de horas (+/- 12 hrs) según el huso horario de emisión.

---

## 📋 Requisitos del Archivo de Salida (EPG Uploader)

El archivo `.xlsx` generado cumple las siguientes especificaciones:
1. **Columnas A - P:**
   - **Columna A (Día):** Texto plano `MM-DD-YYYY` (separado por guiones).
   - **Columna B (Hora):** Formato 24 horas `HH:MM`.
   - **Columna C (Programa):** Texto en MAYÚSCULAS y limpio de metadatos.
   - **Columnas D a P:** Vacías pero presentes.
2. **Escritura como Texto Forzado:** Uso de `worksheet.write_string()` con `xlsxwriter` para evitar alteraciones por configuración regional del sistema operativo.
3. **Pestaña del Canal:** Máximo 31 caracteres, sin caracteres reservados (`: \ / ? * [ ]`).

---

## 🛠️ Estructura del Proyecto

```
AUTOMATIZACION EPG/
├── app_v2.py                 # Aplicación principal en Streamlit
├── excel_preprocessing.py    # Desfusión y propagación de celdas combinadas (openpyxl)
├── normalization.py          # Reglas centrales de parseo, fechas, horas y regex
├── requirements.txt          # Dependencias de Python
├── README.md                 # Documentación general
└── Plan_Arquitectura_v2.md   # Especificación de arquitectura y diseño técnico
```

---

## 💻 Instalación y Uso

### 1. Requisitos Previos
Tener instalado Python 3.10 o superior.

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Ejecutar la aplicación
```bash
streamlit run app_v2.py
```

La aplicación se abrirá automáticamente en tu navegador predeterminado (por defecto en `http://localhost:8501`).
