import streamlit as st
import pandas as pd
from fpdf import FPDF

st.set_page_config(page_title="Auditoría - Estrategia Jurídica", layout="wide")

st.markdown("### 📋 Seguimiento de Líneas de Acción")

archivo = st.file_uploader("Cargar libro de delegación (.xlsm)", type=["xlsm"])

if archivo:
    # Leer la hoja 1 sin encabezados para manejar nosotros las filas
    df = pd.read_excel(archivo, sheet_name=0, header=None)

    # Datos fijos de tu plantilla (según las capturas y el archivo)
    delegacion = df.iloc[1, 10] # Celda K2 
    st.info(f"📍 Delegación: {delegacion}")

    # Mapeo de columnas para los 4 trimestres
    trimestres = {
        "I Trimestre": {"av": 9, "ds": 10},   # Col J, K 
        "II Trimestre": {"av": 14, "ds": 15}, # Col O, P
        "III Trimestre": {"av": 19, "ds": 20},# Col T, U
        "IV Trimestre": {"av": 24, "ds": 25}  # Col Y, Z
    }
    
    t_sel = st.selectbox("Seleccione el Trimestre", list(trimestres.keys()))
    c_av = trimestres[t_sel]["av"]
    c_ds = trimestres[t_sel]["ds"]

    st.markdown("---")

    # Recorremos desde la fila 11 (donde empiezan los datos reales) 
    for i in range(10, len(df)):
        categoria = df.iloc[i, 2]   # Col C (GL/FP) 
        indicador = df.iloc[i, 6]   # Col G (Indicadores) 
        
        # Si no hay indicador, terminamos el bloque
        if pd.isna(indicador): continue

        # La problemática está en la fila 8, columna F 
        problematica = df.iloc[7, 5] 
        meta = df.iloc[i, 8]        # Col I (Meta) 
        avance = df.iloc[i, c_av]
        descripcion = df.iloc[i, c_ds]

        # Visualización limpia por línea
        with st.expander(f"🔹 {categoria} | {indicador}", expanded=True):
            st.caption(f"**Problemática:** {problematica}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.text_input("Meta (Editable)", value=meta, key=f"m_{i}")
                st.markdown(f"**Avance:** `{avance}`")
            with col2:
                st.text_area("Descripción del Resultado", value=descripcion, height=100, key=f"d_{i}")

            st.text_area("Observaciones del Auditor", placeholder="Escriba aquí sus notas...", key=f"obs_{i}")
            st.checkbox("Verificado / Cumple con evidencia", key=f"v_{i}")

    if st.button("Generar Reporte Final"):
        st.success("Reporte listo para descarga.")
