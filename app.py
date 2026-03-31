import streamlit as st
import pandas as pd

# Configuración visual minimalista
st.set_page_config(page_title="Estrategia Jurídica", layout="wide")

st.markdown("### 📋 Sistema de Auditoría de Cumplimiento")
st.write("Cargue el archivo .xlsm para iniciar el rastreo de indicadores.")

# 1. Carga de archivo
archivo = st.file_uploader("Seleccionar libro de Excel", type=["xlsm"])

if archivo:
    # 2. El Comparador: Escaneo dinámico de todas las hojas
    hojas = pd.read_excel(archivo, sheet_name=None)
    
    # Intenta detectar la delegación en los nombres de las hojas o contenido
    delegacion = "Sarchí" if "SARCHI" in str(archivo.name).upper() else "No detectada"
    st.info(f"📍 Delegación detectada: {delegacion}")

    # 3. Filtros de Auditoría
    col1, col2 = st.columns(2)
    with col1:
        trimestre = st.selectbox("Trimestre a Evaluar", ["T1", "T2", "T3", "T4"])
    with col2:
        categoria = st.radio("Categoría", ["GL", "FP"], horizontal=True)

    st.markdown("---")
    
    # 4. Panel de Edición y Observaciones por Línea
    st.write("#### 📝 Detalle de la Auditoría")
    
    # Aquí se simula la fila encontrada por el comparador
    with st.expander("Ver Línea de Acción e Indicadores", expanded=True):
        meta = st.text_input("Meta Anual (Editable)", value="Meta detectada en el archivo")
        avance = st.text_area("Avance y Descripción (Lectura)", value="Descripción técnica del trimestre...", height=100, disabled=True)
        
        # Apartado de observaciones solicitado
        observaciones = st.text_area("Observaciones del Auditor por Línea")
        
        evidencia = st.checkbox("✔ Cumple con la evidencia documental")

    # 5. Botón de reporte
    if st.button("Generar Informe Trimestral (PDF)"):
        st.write("Procesando documento formal...")
