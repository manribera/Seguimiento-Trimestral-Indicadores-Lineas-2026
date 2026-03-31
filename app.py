import streamlit as st
import pandas as pd

# Interfaz Minimalista Profesional
st.set_page_config(page_title="Auditoría Estrategia Jurídica", layout="wide")

st.markdown("### 📋 Detalle de la Auditoría")

archivo = st.file_uploader("Cargar Informe de Avance (.xlsm)", type=["xlsm"])

if archivo:
    # Leer el Excel (el Comparador busca en la hoja de 'Líneas de Acción')
    df = pd.read_excel(archivo, sheet_name=0, header=None) # Ajustar según hoja real

    # 🕵️ El Comparador rastrea la Delegación
    delegacion = "No detectada"
    for r in range(10): 
        for c in range(df.shape[1]):
            if "DELEGACION POLICIAL" in str(df.iloc[r, c]).upper():
                delegacion = df.iloc[r, c+1] 
                break

    st.info(f"📍 Delegación: {delegacion}")

    # --- Filtros ---
    col_t, col_c = st.columns(2)
    with col_t:
        trimestre_sel = st.selectbox("Seleccione el Trimestre", ["I Trimestre", "II Trimestre", "III Trimestre", "IV Trimestre"])
    with col_c:
        cat_sel = st.radio("Categoría", ["GL", "FP"], horizontal=True)

    st.markdown("---")

    # --- Simulación de carga de una línea de acción ---
    # En la versión final, esto será un bucle que recorra cada indicador
    with st.expander("Ver Línea de Acción e Indicadores", expanded=True):
        
        # Datos extraídos del Excel (Ejemplo basado en tu imagen)
        meta_excel = "2 por año, 4 en total." 
        desc_excel = "Charlas de Educación en Seguridad..."

        meta_edit = st.text_input("Meta Anual (Editable)", value=meta_excel)
        
        st.text_area("Avance y Descripción (Lectura)", value=desc_excel, height=100, disabled=True)
        
        # Apartado de observaciones por línea
        st.text_area("Observaciones del Auditor por Línea", placeholder="Escriba aquí su análisis de cumplimiento...")
        
        col_check, col_vacia = st.columns([1, 3])
        with col_check:
            evidencia = st.checkbox("✔ Cumple con la evidencia documental")

    # Botón de acción
    if st.button("Generar Informe Trimestral (PDF)"):
        st.success(f"Informe de {delegacion} - {trimestre_sel} preparado para descarga.")
