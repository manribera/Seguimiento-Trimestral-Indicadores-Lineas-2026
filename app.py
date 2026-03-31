import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

st.set_page_config(page_title="Estrategia Jurídica - Auditoría", layout="wide")

st.title("📋 Seguimiento Técnico: Líneas y Problemáticas")
st.markdown("---")

archivos = st.file_uploader("📁 Cargar archivos de delegaciones", type=["xlsm"], accept_multiple_files=True)

def generar_pdf(datos, delegacion, trimestre):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(190, 10, f"AUDITORÍA: {delegacion}", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(190, 10, f"Trimestre: {trimestre}", ln=True, align='C')
    pdf.ln(10)
    for d in datos:
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(190, 7, f"SECCIÓN: {d['titulo']}")
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(190, 6, f"Indicador: {d['indicador']} | Meta: {d['meta']}")
        pdf.multi_cell(190, 6, f"Estado Final: {d['estado']}")
        pdf.set_text_color(0, 0, 255)
        pdf.multi_cell(190, 6, f"OBSERVACIONES: {d['obs']}")
        pdf.set_text_color(0, 0, 0)
        pdf.cell(190, 6, f"Evidencia Verificada: {'SÍ' if d['v'] else 'NO'}", ln=True)
        pdf.ln(4)
        pdf.cell(190, 0, '', 'T', ln=True)
    return pdf.output(dest='S').encode('latin-1')

if archivos:
    archivo_sel = st.selectbox("🎯 Delegación a revisar", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)

    # Buscador dinámico de nombre de delegación
    nombre_unidad = "No detectada"
    for r in range(5):
        for c in range(df.shape[1]):
            if "DELEGACION" in str(df.iloc[r, c]).upper():
                nombre_unidad = str(df.iloc[r, c+1]) if pd.notna(df.iloc[r, c+1]) else str(df.iloc[r, c])
                break

    st.subheader(f"📍 Unidad: {nombre_unidad}")
    
    # Mapeo de columnas (Avance, Descripción, Cantidad)
    trim_map = {
        "I Trimestre": {"av": 9, "ds": 10, "cant": 11}, 
        "II Trimestre": {"av": 14, "ds": 15, "cant": 16},
        "III Trimestre": {"av": 19, "ds": 20, "cant": 21},
        "IV Trimestre": {"av": 24, "ds": 25, "cant": 26}
    }
    t_sel = st.selectbox("Trimestre de Evaluación", list(trim_map.keys()))
    indices = trim_map[t_sel]

    reporte_final = []
    linea_actual, prob_actual = "", ""

    # Procesamiento por fila (Desde fila 8 para capturar títulos)
    for i in range(7, len(df)):
        val_d = str(df.iloc[i, 3]) # Columna D: Linea de Accion #X
        val_f = str(df.iloc[i, 5]) # Columna F: Problematica

        # Si detectamos un nuevo bloque de Línea de Acción, actualizamos el título
        if "LINEA DE ACCION" in val_d.upper():
            linea_actual = val_d
            prob_actual = val_f if pd.notna(df.iloc[i, 5]) else prob_actual
            st.markdown(f"## 🚩 {linea_actual}")
            st.markdown(f"**Problemática:** {prob_actual}")
            st.markdown("---")

        indicador = df.iloc[i, 6] # Columna G: Indicador
        if pd.isna(indicador) or "Indicadores" in str(indicador):
            continue

        # Lógica de Avance Automático basada en Cantidad
        cantidad_val = df.iloc[i, indices["cant"]]
        if pd.notna(cantidad_val) and str(cantidad_val).strip() != "" and cantidad_val != 0:
            estado_calculado = "Con Actividades / Completado"
            color_estado = "green"
        else:
            estado_calculado = "Sin Actividades"
            color_estado = "red"

        with st.container():
            # Tabla de Datos del Excel
            datos_tabla = {
                "Indicador": [indicador],
                "Cantidad": [cantidad_val],
                "Descripción Reportada": [df.iloc[i, indices["ds"]]]
            }
            st.table(pd.DataFrame(datos_tabla))

            # Panel de Auditoría
            c1, c2, c3 = st.columns([1, 2, 1])
            with c1:
                meta_e = st.text_input("Meta Anual", value=df.iloc[i, 8], key=f"m_{i}")
                st.markdown(f"**Estado Sugerido:** :{color_estado}[{estado_calculado}]")
            with c2:
                obs_e = st.text_area("Observaciones de Auditoría", key=f"o_{i}", height=68)
            with c3:
                st.write("¿Evidencia?")
                ver_e = st.checkbox("Verificado", key=f"v_{i}")

            reporte_final.append({
                "titulo": f"{linea_actual} - {prob_actual}", "indicador": indicador, 
                "meta": meta_e, "estado": estado_calculado, "obs": obs_e, "v": ver_e
            })
            st.markdown("---")

    if st.button("📄 Finalizar y Generar PDF"):
        pdf_bytes = generar_pdf(reporte_final, nombre_unidad, t_sel)
        st.download_button("📥 Descargar Informe Final", data=pdf_bytes, file_name=f"Auditoria_{nombre_unidad}.pdf")
