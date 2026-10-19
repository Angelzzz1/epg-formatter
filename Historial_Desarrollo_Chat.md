# Registro de Desarrollo - Automatización EPG

Este documento resume las conversaciones, decisiones y desarrollos realizados durante la creación de la aplicación de automatización de EPG con Python y Streamlit.

## 1. Objetivo del Proyecto
Crear una herramienta ETL (Extract, Transform, Load) a prueba de fallos para un entorno NOC. El objetivo es procesar diferentes formatos de archivos Excel (Guías Electrónicas de Programación) enviados por proveedores y estandarizarlos al formato estricto requerido por el "EPG Uploader / Manager" (Columnas de la A a la P, formato de fecha MM-DD-YYYY, etc.).

## 2. Fases de Desarrollo implementadas en `app.py`

### Modo Lineal (Ej. Suyai / Cine Premium)
- **Estructura:** Cada fila representa un programa con su Día, Hora y Nombre en columnas separadas.
- **Solución:** Se implementó mapeo directo de columnas.

### Modo Matricial (Ej. Canal LUZ)
- **Estructura:** Los días de la semana están como columnas (Lunes, Martes, etc.) y las filas representan horas. Los archivos pueden tener múltiples pestañas por semanas.
- **Solución:** Se usa `pd.melt` (unpivot) para aplanar la matriz. Se agregó lógica para ordenar las pestañas leyendo las fechas en sus nombres y mostrar las más recientes arriba en el menú desplegable.

### Modo Bloque Inteligente (Ej. Antena 1)
- **Estructura:** Un archivo apilado verticalmente sin columnas de fechas. Usa títulos incrustados como "PROGRAMACION ANTENA 1 LUNES A VIERNES", "SABADO" y "DOMINGO".
- **Solución:** Se creó un "cerebro" (Auto-Detección). El script escanea las filas; cuando lee "LUNES A VIERNES", asigna todo lo que sigue a los días de semana. Al leer "SABADO" o "DOMINGO", reasigna el contexto. 
- **Correcciones de Errores:** 
  - Se implementó limpieza de *non-breaking spaces* en los títulos (espacios ocultos).
  - Se corrigió el parseo de horas para tolerar espacios accidentales (ej. `20: 00`).
  - Se resolvieron duplicados causados por la función `ffill()` al aislar la lectura inteligente sobre los datos crudos.

### Manejo de Zona Horaria (Timezone)
- Los operadores pueden ajustar horas sumando o restando de forma dinámica, manteniendo la fecha intacta en formato de salida sin importar si el programa cruza la medianoche (según solicitud para EPG Uploader).

## 3. Discusión Arquitectónica futura: Agente IA Supervisor
Se discutió la posibilidad de agregar un Agente de IA para operar el sistema:
- **Concepto:** Un Agente Enrutador (Router Agent) con uso de herramientas (Tool Calling).
- **Funcionamiento:** El operador sube el Excel; el Agente analiza la estructura, detecta anomalías (ej. solapamiento de horarios), escoge automáticamente el script correcto a ejecutar (Lineal, Matricial, Bloque) y notifica irregularidades.
- **Opciones Gratuitas:** Uso de modelos locales (Ollama/Llama 3) si hay hardware en el NOC, o cuotas gratuitas de APIs (Gemini 1.5 Flash, Groq).
- **Integración:** Streamlit ya soporta interfaces de chat, lo cual permite crear una pestaña "🤖 Asistente IA" dentro de la misma app existente.

## 4. Consideraciones de Seguridad (Puerto 8501)
- **Localhost (Seguro):** Si se accede desde la misma máquina usando `http://localhost:8501`, no hay riesgo externo.
- **Red Local / LAN (Bajo Riesgo):** Streamlit abre el puerto `0.0.0.0`. Esto permite que otras computadoras dentro de la red local del NOC (o VPN) accedan a la aplicación (`http://<IP-DEL-SERVIDOR>:8501`). En una red corporativa interna, el riesgo es mínimo y es útil para que los operadores no instalen Python en todas las PCs.
- **Internet Público (Alto Riesgo):** **NO** se debe abrir el puerto 8501 en el router/firewall de la empresa hacia internet. Streamlit no tiene sistema de login (usuario/contraseña) nativo. Si se expone a internet público, cualquier persona podría subir archivos o ejecutar comandos. Si se requiere acceso externo real, debe hacerse a través de una VPN corporativa, o configurando un Reverse Proxy (ej. Nginx) con autenticación básica.

## 5. Pr�ximo Paso Inmediato (Punto de reinicio)
El desarrollo actual tiene listos los 3 modos (Lineal, Matricial, Bloque). El objetivo inmediato al retomar este proyecto es iniciar la implementaci�n del **Agente IA Supervisor** dentro de pp.py, integr�ndolo posiblemente con una pesta�a de chat de Streamlit. El agente debe actuar como Router (Tool Calling) para detectar el tipo de archivo y procesarlo autom�ticamente.
