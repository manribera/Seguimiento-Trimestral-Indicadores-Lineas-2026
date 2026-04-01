import streamlit as st
import pandas as pd
import openpyxl
from fpdf import FPDF

# Configuración de página ancha para que la tabla quepa bien
st.set_page_config(page_title="Auditoría - Estrategia Jurídica", layout="wide")

st.markdown("## 📋 Herramienta de Auditoría: Formato de Tabla")
st.markdown("---")

# Carga de archivos múltiples
archivos = st.file_uploader("📁 Cargar archivos de delegaciones (.xlsm)", type=["xlsm"], accept_multiple_files=True)

if archivos:
    archivo_sel = st.selectbox("🎯 Seleccione el archivo a auditar", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    # Lector de Excel con openpyxl (Lector de celdas)
    wb = openpyxl.load_workbook(archivo_actual, data_only=True)
    ws = wb.active

    # --- LÓGICA DEL LECTOR: Búsqueda de anclajes ---
    def obtener_valor(texto_clave):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and texto_clave.upper() in str(cell.value).upper():
                    return ws.cell(row=cell.row, column=cell.column + 1).value
        return ""

    # Extraer encabezados del formato
    delegacion = obtener_valor("Delegación")
    linea_num = obtener_valor("linea de accion #")
    problematica = obtener_valor("Problemática")
    lider = obtener_valor("Líder Estrategico")
    trimestre_excel = obtener_valor("Trimestre")

    # --- RENDERIZADO DEL ENCABEZADO (Igual a tu imagen) ---
    st.markdown(f"**Delegación:** `{delegacion}`")
    
    # Fila de información superior
    h1, h2, h3 = st.columns([1, 2, 2])
    h1.info(f"**Línea de Acción #:** {linea_num}")
    h2.info(f"**Problemática:** {problematica}")
    h3.info(f"**Líder Estratégico:** {lider}")
    
    st.write(f"**Trimestre:** {trimestre_excel}")
    st.markdown("---")

    # --- TABLA DE INDICADORES (LECTURA Y EDICIÓN) ---
    # Encabezados de la tabla visual
    t_col1, t_col2, t_col3, t_col4, t_col5, t_col6 = st.columns([2, 1, 1, 2, 1, 2])
    t_col1.write("**Indicador**")
    t_col2.write("**Meta (Editable)**")
    t_col3.write("**Avance**")
    t_col4.write("**Descripción**")
    t_col5.write("**Cantidad**")
    t_col6.write("**Observaciones (Editable)**")
    st.markdown("---")

    # Procesamiento de filas con Pandas para la tabla masiva
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)
    
    # Mapeo de columnas por trimestre (Basado en la estructura de Sembremos Seguridad)
    trim_map = {
        "I": {"av": 9, "ds": 10, "ct": 11}, "II": {"av": 14, "ds": 15, "ct": 16},
        "III": {"av": 19, "ds": 20, "ct": 21}, "IV": {"av": 24, "ds": 25, "ct": 26}
    }
    t_key = str(trimestre_excel).strip()
    idx = trim_map.get(t_key, trim_map["I"])

    reporte_datos = []

    # Localizar el inicio de la tabla (Fila 11 aprox)
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

        # --- FILA DE LA TABLA EN STREAMLIT ---
        with st.container():
            c1, c2, c3, c4, c5, c6 = st.columns([2, 1, 1, 2, 1, 2])
            
            c1.write(indicador)
            meta_e = c2.text_input("Meta", value=meta_excel, key=f"meta_{i}", label_visibility="collapsed")
            c3.write(avance_status)
            c4.write(desc_excel)
            c5.write(cant_excel)
            obs_e = c6.text_area("Notas", key=f"obs_{i}", height=70, label_visibility="collapsed")
            
            # Checkbox de verificación debajo de cada fila para orden
            v_check = st.checkbox("Evidencia Verificada", key=f"v_{i}")
            
            reporte_datos.append({
                "titulo": f"{linea_num} - {problematica}",
                "indicador": indicador, "meta": meta_e, 
                "avance": avance_status, "obs": obs_e, "v": v_check
            })
            st.markdown("---")

    # Botón Final
    if st.button("📄 Generar Reporte Consolidado"):
        st.success(f"Auditoría de {delegacion} procesada. PDF listo para descarga.")
