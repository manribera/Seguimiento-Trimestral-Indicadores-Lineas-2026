import streamlit as st
import pandas as pd
from fpdf import FPDF
import base64

st.set_page_config(page_title="Auditoría - Estrategia Jurídica", layout="wide")

st.markdown("### 📋 Seguimiento de Líneas de Acción")

archivo = st.file_uploader("Cargar libro de delegación (.xlsm)", type=["xlsm"])

def generar_pdf(datos, delegacion, trimestre):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(190, 10, f"Informe de Auditoria: {delegacion}", ln=True, align='C')
    pdf.set_font("Arial", '', 12)
    pdf.cell(190, 10, f"Periodo: {trimestre}", ln=True, align='C')
    pdf.ln(10)
    
    for d in datos:
        pdf.set_font("Arial", 'B', 11)
        pdf.multi_cell(190, 7, f"{d['titulo']}")
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(190, 6, f"Indicador: {d['indicador']}")
        pdf.multi_cell(190, 6, f"Meta: {d['meta']}")
        pdf.multi_cell(190, 6, f"Observaciones: {d['obs']}")
        pdf.cell(190, 6, f"Evidencia: {'SI' if d['v'] else 'NO'}", ln=True)
        pdf.ln(5)
    return pdf.output(dest='S').encode('latin-1')

if archivo:
    df = pd.read_excel(archivo, sheet_name=0, header=None)
    
    # Identificar Delegación (Celda H2/K2 según combinado)
    delegacion = str(df.iloc[1, 7]).replace("Delegacion Policial:", "").strip()
    st.info(f"📍 Delegación: {delegacion}")

    trimestres = {
        "I Trimestre": {"av": 9, "ds": 10}, "II Trimestre": {"av": 14, "ds": 15},
        "III Trimestre": {"av": 19, "ds": 20}, "IV Trimestre": {"av": 24, "ds": 25}
    }
    t_sel = st.selectbox("Seleccione el Trimestre", list(trimestres.keys()))
    
    reporte_final = []
    
    # Variables para rastrear el título actual
    linea_actual = ""
    problematica_actual = ""

    # Recorremos desde la fila 8 para capturar encabezados de bloques
    for i in range(7, len(df)):
        # 1. Detectar si la fila es un encabezado de Línea de Acción (Columna D)
        val_d = str(df.iloc[i, 3])
        if "LINEA DE ACCION" in val_d.upper():
            linea_actual = val_d
            # La problemática suele estar en la misma fila o la siguiente (Columna F)
            problematica_actual = str(df.iloc[i, 5]) if not pd.isna(df.iloc[i, 5]) else problematica_actual

        # 2. Detectar Indicadores (Columna G)
        indicador = df.iloc[i, 6]
        if pd.isna(indicador) or "Indicadores" in str(indicador):
            continue

        categoria = df.iloc[i, 2] # GL o FP
        meta = df.iloc[i, 8]
        avance = df.iloc[i, trimestres[t_sel]["av"]]
        descripcion = df.iloc[i, trimestres[t_sel]["ds"]]

        # 3. Título del Apartado
        titulo_seccion = f"{linea_actual} - {problematica_actual}"

        with st.expander(f"📂 {titulo_seccion} | {indicador}", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                meta_e = st.text_input("Meta", value=meta, key=f"m_{i}")
                st.write(f"**Avance:** {avance}")
            with c2:
                st.text_area("Descripción", value=descripcion, height=80, key=f"d_{i}")
            
            obs_e = st.text_area("Observaciones del Auditor", key=f"o_{i}")
            verif_e = st.checkbox("Evidencia verificada", key=f"v_{i}")
            
            reporte_final.append({
                "titulo": titulo_seccion, "indicador": indicador, 
                "meta": meta_e, "obs": obs_e, "v": verif_e
            })

    st.markdown("---")
    if st.button("📄 Generar y Descargar PDF"):
        if reporte_final:
            pdf_bytes = generar_pdf(reporte_final, delegacion, t_sel)
            st.download_button("📥 Click para descargar archivo", data=pdf_bytes, 
                               file_name=f"Auditoria_{delegacion}.pdf", mime="application/pdf")
        else:
            st.error("No hay datos para procesar.")
