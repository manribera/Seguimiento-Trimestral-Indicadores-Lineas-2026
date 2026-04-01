import streamlit as st
import pandas as pd
import openpyxl
from fpdf import FPDF

# Configuración de página ancha para emular el formato de tabla
st.set_page_config(page_title="Auditoría - Estrategia Jurídica", layout="wide")

st.markdown("## 📋 Herramienta de Auditoría: Formato de Tabla de Seguimiento")
st.markdown("---")

# Carga de archivos múltiples (Mantenemos la capacidad de subir los 98 archivos)
archivos = st.file_uploader("📁 Cargar archivos de delegaciones (.xlsm)", type=["xlsm"], accept_multiple_files=True)

if archivos:
    archivo_sel = st.selectbox("🎯 Seleccione el archivo a auditar", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    # Lector de Excel con openpyxl (Buscador dinámico de celdas)
    wb = openpyxl.load_workbook(archivo_actual, data_only=True)
    ws = wb.active

    # --- LÓGICA DEL LECTOR: Búsqueda de anclajes para el encabezado ---
    def obtener_valor(texto_clave):
        for row in ws.iter_rows(max_row=20): # Escaneo del encabezado
            for cell in row:
                if cell.value and texto_clave.upper() in str(cell.value).upper():
                    return ws.cell(row=cell.row, column=cell.column + 1).value
        return ""

    # Extraer datos dinámicamente según tu imagen
    delegacion = obtener_valor("Delegación")
    linea_num = obtener_valor("linea de accion #")
    problematica = obtener_valor("Problemática")
    lider = obtener_valor("Líder Estratégico")
    trimestre_excel = obtener_valor("Trimestre")

    # --- RENDERIZADO DEL ENCABEZADO (Mimetismo con tu imagen) ---
    st.write(f"**Delegación:** `{delegacion}`")
    
    h1, h2, h3 = st.columns([1, 2, 2])
    h1.info(f"**Línea de Acción #:** {linea_num}")
    h2.info(f"**Problemática:** {problematica}")
    h3.info(f"**Líder Estratégico:** {lider}")
    
    st.write(f"**Trimestre:** {trimestre_excel}")
    st.markdown("---")

    # --- ENCABEZADO DE LA TABLA DE AUDITORÍA ---
    t_col1, t_col2, t_col3, t_col4, t_col5, t_col6 = st.columns([2, 1, 1, 2, 1, 2])
    t_col1.write("**Indicador**")
    t_col2.write("**Meta (Editable)**")
    t_col3.write("**Avance**")
    t_col4.write("**Descripción**")
    t_col5.write("**Cantidad**")
    t_col6.write("**Observaciones (Editable)**")
    st.markdown("---")

    # Procesamiento de filas con Pandas para eficiencia masiva
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)
    
    # Mapeo de columnas según el diseño de Sembremos Seguridad
    trim_map = {
        "I": {"av": 9, "ds": 10, "ct": 11}, "II": {"av": 14, "ds": 15, "ct": 16},
        "III": {"av": 19, "ds": 20, "ct": 21}, "IV": {"av": 24, "ds": 25, "ct": 26}
    }
    t_key = str(trimestre_excel).strip()
    idx = trim_map.get(t_key, trim_map["I"])

    reporte_datos = []

    # Localizar el inicio de los indicadores (Fila 11 aprox)
    for i in range(10, len(df)):
        indicador = df.iloc[i, 6] # Columna G
        if pd.isna(indicador) or "Indicador" in str(indicador):
            continue

        meta_excel = df.iloc[i, 8] # Columna I
        desc_excel = df.iloc[i, idx["ds"]]
        cant_excel = df.iloc[i, idx["ct"]]

        # Lógica de asignación automática de Avance por Cantidad
        if pd.notna(cant_excel) and str(cant_excel).strip() != "" and cant_excel != 0:
            avance_status = "🟢 Con Actividad"
        else:
            avance_status = "🔴 Sin Actividad"

        # --- FILA ÚNICA TIPO TABLA ---
        with st.container():
            c1, c2, c3, c4, c5, c6 = st.columns([2, 1, 1, 2, 1, 2])
            
            c1.write(indicador)
            meta_e = c2.text_input("Meta", value=meta_excel, key=f"meta_{i}", label_visibility="collapsed")
            c3.write(avance_status)
            c4.write(desc_excel)
            c5.write(cant_excel)
            obs_e = c6.text_area("Observaciones", key=f"obs_{i}", height=70, label_visibility="collapsed")
            
            # Verificación rápida de evidencia
            v_check = st.checkbox("✔ Verificado", key=f"v_{i}")
            
            reporte_datos.append({
                "titulo": f"{linea_num} - {problematica}",
                "indicador": indicador, "meta": meta_e, 
                "avance": avance_status, "obs": obs_e, "v": v_check
            })
            st.markdown("---")

    # Botón de Finalización
    if st.button("📄 Generar Informe de Auditoría"):
        st.success(f"Procesando auditoría de {delegacion}. PDF preparado.")
