import streamlit as st
import pandas as pd
from fpdf import FPDF

st.set_page_config(page_title="Estrategia Jurídica - Dashboard", layout="wide")

# Estética minimalista y profesional
st.markdown("## 📋 Panel de Auditoría: Sembremos Seguridad")

archivo = st.file_uploader("Cargar libro de delegación (.xlsm)", type=["xlsm"])

def generar_pdf(datos, delegacion, trimestre):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(190, 10, f"REPORTE DE AUDITORIA - {delegacion}", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(190, 10, f"Trimestre evaluado: {trimestre}", ln=True, align='C')
    pdf.ln(10)
    
    for d in datos:
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(190, 7, f"SECCION: {d['titulo']}")
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(190, 6, f"Indicador: {d['indicador']}")
        pdf.multi_cell(190, 6, f"Meta: {d['meta']}")
        pdf.set_text_color(0, 0, 255) # Azul para observaciones
        pdf.multi_cell(190, 6, f"OBSERVACIONES: {d['obs']}")
        pdf.set_text_color(0, 0, 0)
        pdf.cell(190, 6, f"Cumplimiento: {'APROBADO' if d['v'] else 'RECHAZADO'}", ln=True)
        pdf.ln(4)
        pdf.cell(190, 0, '', 'T', ln=True)
        pdf.ln(4)
    return pdf.output(dest='S').encode('latin-1')

if archivo:
    df = pd.read_excel(archivo, sheet_name=0, header=None)
    
    # --- ENCABEZADO DE LA APP (Emulando el Excel) ---
    # Captura Delegación desde la celda J2/K2
    delegacion = str(df.iloc[1, 10]).strip() if pd.notna(df.iloc[1, 10]) else "No especificada"
    st.subheader(f"📍 Unidad: {delegacion}")

    trim_map = {
        "I Trimestre": {"av": 9, "ds": 10}, "II Trimestre": {"av": 14, "ds": 15},
        "III Trimestre": {"av": 19, "ds": 20}, "IV Trimestre": {"av": 24, "ds": 25}
    }
    t_sel = st.selectbox("Seleccione el Trimestre de Seguimiento", list(trim_map.keys()))
    
    reporte_datos = []

    # --- CUERPO DE LA APP (Iteración por bloques de Líneas de Acción) ---
    linea_actual = ""
    prob_actual = ""

    for i in range(7, len(df)):
        val_d = str(df.iloc[i, 3]) # Columna D: Linea de Accion #X
        val_f = str(df.iloc[i, 5]) # Columna F: Problematica

        # Si detectamos un nuevo bloque de Línea de Acción
        if "LINEA DE ACCION" in val_d.upper():
            linea_actual = val_d
            prob_actual = val_f if pd.notna(df.iloc[i, 5]) else prob_actual

        indicador = df.iloc[i, 6] # Columna G: Indicadores
        # Evitar filas de cabecera o vacías
        if pd.isna(indicador) or "Indicadores" in str(indicador):
            continue

        # Estructura visual similar al Excel
        titulo_seccion = f"{linea_actual} | {prob_actual}"
        
        with st.container():
            st.markdown(f"#### 🏷️ {titulo_seccion}")
            st.write(f"**Indicador:** {indicador}")
            
            c1, c2, c3 = st.columns([1, 2, 1])
            with c1:
                meta_e = st.text_input("Meta Anual", value=df.iloc[i, 8], key=f"m_{i}")
                st.write(f"**Avance:** {df.iloc[i, trim_map[t_sel]['av']]}")
            with c2:
                st.text_area("Descripción de la Delegación", value=df.iloc[i, trim_map[t_sel]['ds']], key=f"d_{i}", height=100)
            with c3:
                obs_e = st.text_area("Observaciones Auditoría", key=f"o_{i}", height=100)
                verif_e = st.checkbox("Verificado", key=f"v_{i}")

            reporte_datos.append({
                "titulo": titulo_seccion, "indicador": indicador, 
                "meta": meta_e, "obs": obs_e, "v": verif_e
            })
            st.markdown("---")

    # --- BOTÓN DE DESCARGA ---
    if st.button("📄 Generar Informe de Auditoría"):
        if reporte_datos:
            pdf_bytes = generar_pdf(reporte_datos, delegacion, t_sel)
            st.download_button("📥 Descargar Reporte Final", data=pdf_bytes, 
                               file_name=f"Auditoria_{delegacion}.pdf", mime="application/pdf")
