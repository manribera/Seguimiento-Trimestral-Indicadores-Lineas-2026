import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

st.set_page_config(page_title="Auditoría Consolidada - Estrategia Jurídica", layout="wide")

st.title("📋 Consolidado y Auditoría de Indicadores")
st.write("Cargue los archivos de las delegaciones para iniciar la revisión técnica.")

# 1. CARGA MULTIPLE (Del nuevo código)
archivos = st.file_uploader("📁 Sube archivos .xlsm", type=["xlsm"], accept_multiple_files=True)

def generar_pdf(datos, delegacion, trimestre):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(190, 10, f"REPORTE DE AUDITORIA: {delegacion}", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(190, 10, f"Periodo: {trimestre}", ln=True, align='C')
    pdf.ln(10)
    for d in datos:
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(190, 7, f"SECCION: {d['titulo']}")
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(190, 6, f"Indicador: {d['indicador']}")
        pdf.multi_cell(190, 6, f"Meta: {d['meta']}")
        pdf.set_text_color(0, 0, 255)
        pdf.multi_cell(190, 6, f"OBSERVACIONES: {d['obs']}")
        pdf.set_text_color(0, 0, 0)
        pdf.cell(190, 6, f"Cumplimiento: {'SI' if d['v'] else 'NO'}", ln=True)
        pdf.ln(5)
        pdf.cell(190, 0, '', 'T', ln=True)
    return pdf.output(dest='S').encode('latin-1')

if archivos:
    # Selector para elegir cuál de los archivos cargados auditar ahora
    nombres_archivos = [a.name for a in archivos]
    archivo_sel_nombre = st.selectbox("🎯 Seleccione la delegación a auditar", nombres_archivos)
    
    # Recuperar el objeto de archivo correcto
    archivo_actual = next(a for a in archivos if a.name == archivo_sel_nombre)
    
    # Procesar el archivo seleccionado
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)

    # 🕵️ BUSCADOR DINÁMICO DE DELEGACIÓN (Como el anterior)
    def buscar_delegacion(dataframe):
        for r in range(5):
            for c in range(dataframe.shape[1]):
                celda = str(dataframe.iloc[r, c])
                if "DELEGACION" in celda.upper() or "UNIDAD" in celda.upper():
                    return celda.replace("Delegacion Policial:", "").strip() or str(dataframe.iloc[r, c+1])
        return "No detectada"

    nombre_delegacion = buscar_delegacion(df)
    
    col_header, col_trim = st.columns([3, 1])
    with col_header:
        st.subheader(f"📍 Unidad: {nombre_delegacion}")
    with col_trim:
        trim_map = {
            "I Trimestre": {"av": 9, "ds": 10}, "II Trimestre": {"av": 14, "ds": 15},
            "III Trimestre": {"av": 19, "ds": 20}, "IV Trimestre": {"av": 24, "ds": 25}
        }
        t_sel = st.selectbox("Trimestre", list(trim_map.keys()))

    st.markdown("---")
    reporte_datos = []
    linea_actual, prob_actual = "", ""

    # 🚀 VISTA CONSOLIDADA (APARTADO ÚNICO)
    for i in range(7, len(df)):
        val_c = str(df.iloc[i, 2]) # GL/FP
        val_d = str(df.iloc[i, 3]) # Línea de Acción
        val_f = str(df.iloc[i, 5]) # Problemática (Columna F)
        indicador = df.iloc[i, 6]  # Indicador (Columna G)

        # Capturar y mostrar encabezados de Línea y Problemática
        if "LINEA DE ACCION" in val_d.upper():
            linea_actual = val_d
            prob_actual = val_f if pd.notna(df.iloc[i, 5]) else prob_actual
            st.markdown(f"### 🚩 {linea_actual}")
            st.markdown(f"**Problemática Detectada:** {prob_actual}")
            st.markdown("---")

        if pd.isna(indicador) or "Indicadores" in str(indicador):
            continue

        with st.container():
            c_ind, c_meta, c_res, c_aud = st.columns([2, 1, 2, 2])
            with c_ind:
                st.write(f"**[{val_c}]**\n{indicador}")
            with c_meta:
                m_edit = st.text_input("Meta", value=df.iloc[i, 8], key=f"m_{i}_{archivo_actual.name}")
                st.write(f"**Avance:** {df.iloc[i, trim_map[t_sel]['av']]}")
            with c_res:
                st.text_area("Descripción", value=df.iloc[i, trim_map[t_sel]['ds']], key=f"d_{i}_{archivo_actual.name}", height=120)
            with c_aud:
                obs_e = st.text_area("Observaciones Auditor", key=f"o_{i}_{archivo_actual.name}", height=90)
                ver_e = st.checkbox("Evidencia verificada", key=f"v_{i}_{archivo_actual.name}")

            reporte_datos.append({
                "titulo": f"{linea_actual} - {prob_actual}", "indicador": indicador, 
                "meta": m_edit, "obs": obs_e, "v": ver_e
            })
            st.markdown("<br>", unsafe_allow_html=True)

    if st.button("📄 Generar Informe PDF"):
        if reporte_datos:
            pdf_bytes = generar_pdf(reporte_datos, nombre_delegacion, t_sel)
            st.download_button("📥 Descargar PDF", data=pdf_bytes, 
                               file_name=f"Auditoria_{nombre_delegacion}.pdf", mime="application/pdf")
    st.markdown("---")
    if st.button("📄 Generar y Descargar Reporte Final"):
        if reporte_datos:
            pdf_bytes = generar_pdf(reporte_datos, nombre_delegacion, t_sel)
            st.download_button("📥 Descargar PDF", data=pdf_bytes, 
                               file_name=f"Auditoria_{nombre_delegacion}.pdf", mime="application/pdf")
