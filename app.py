import streamlit as st
import pandas as pd
from fpdf import FPDF

st.set_page_config(page_title="Auditoría - Estrategia Jurídica", layout="wide")

# Estilo para que la tabla se vea limpia
st.markdown("""
    <style>
    .reportview-container .main .block-container{ padding-top: 1rem; }
    .stTextArea textarea { font-size: 14px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📋 Instrumento de Seguimiento y Auditoría")

archivo = st.file_uploader("Cargar libro de delegación (.xlsm)", type=["xlsm"])

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
        pdf.ln(3)
    return pdf.output(dest='S').encode('latin-1')

if archivo:
    df = pd.read_excel(archivo, sheet_name=0, header=None)

    # 🕵️ BUSCADOR DINÁMICO DE DELEGACIÓN (Busca la palabra y toma el valor de la derecha)
    def buscar_delegacion(dataframe):
        for r in range(5): # Solo busca en las primeras 5 filas
            for c in range(dataframe.shape[1]):
                celda = str(dataframe.iloc[r, c])
                if "DELEGACION" in celda.upper() or "UNIDAD" in celda.upper():
                    # Intenta tomar el valor de la misma celda o la de la derecha
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
    linea_actual = ""
    prob_actual = ""

    # Recorrido de filas para emular el Excel
    for i in range(7, len(df)):
        val_c = str(df.iloc[i, 2]) # Categoría GL/FP
        val_d = str(df.iloc[i, 3]) # Línea de Acción
        val_f = str(df.iloc[i, 5]) # Problemática
        indicador = df.iloc[i, 6]  # Indicador

        # Detectar Cambio de Bloque (Línea de Acción)
        if "LINEA DE ACCION" in val_d.upper():
            linea_actual = val_d
            prob_actual = val_f if pd.notna(df.iloc[i, 5]) else prob_actual
            st.markdown(f"### 🚩 {linea_actual}")
            st.markdown(f"**Problemática:** {prob_actual}")
            st.markdown("---")

        # Solo procesar filas que tienen indicadores reales
        if pd.isna(indicador) or "Indicadores" in str(indicador):
            continue

        # FILA ÚNICA CONSOLIDADA (Emulación de Excel)
        with st.container():
            c_ind, c_meta, c_res, c_aud = st.columns([2, 1, 2, 2])
            
            with c_ind:
                st.write(f"**[{val_c}]**\n{indicador}")
            
            with c_meta:
                m_edit = st.text_input("Meta", value=df.iloc[i, 8], key=f"m_{i}")
                st.write(f"**Avance:** {df.iloc[i, trim_map[t_sel]['av']]}")
            
            with c_res:
                st.text_area("Descripción", value=df.iloc[i, trim_map[t_sel]['ds']], key=f"d_{i}", height=120)
            
            with c_aud:
                obs_e = st.text_area("Observaciones Auditor", key=f"o_{i}", height=90)
                ver_e = st.checkbox("Evidencia verificada", key=f"v_{i}")

            reporte_datos.append({
                "titulo": f"{linea_actual} - {prob_actual}", "indicador": indicador, 
                "meta": m_edit, "obs": obs_e, "v": ver_e
            })
            st.markdown("<br>", unsafe_allow_html=True)

    # Botón Final
    st.markdown("---")
    if st.button("📄 Generar y Descargar Reporte Final"):
        if reporte_datos:
            pdf_bytes = generar_pdf(reporte_datos, nombre_delegacion, t_sel)
            st.download_button("📥 Descargar PDF", data=pdf_bytes, 
                               file_name=f"Auditoria_{nombre_delegacion}.pdf", mime="application/pdf")
