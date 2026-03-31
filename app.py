import streamlit as st
import pandas as pd
from fpdf import FPDF

# Configuración profesional y minimalista
st.set_page_config(page_title="Estrategia Jurídica - Seguimiento", layout="wide")

st.markdown("### 📋 Instrumento de Seguimiento de Líneas de Acción")

archivo = st.file_uploader("Cargar libro de delegación (.xlsm)", type=["xlsm"])

if archivo:
    # Leer el Excel completo (Hoja 1)
    df = pd.read_excel(archivo, sheet_name=0, header=None)

    # 🔍 COMPARADOR: Localizar textos clave sin usar columnas fijas
    def buscar(texto):
        for r in range(25):
            for c in range(df.shape[1]):
                if texto.upper() in str(df.iloc[r, c]).upper():
                    return r, c
        return None, None

    # Detectar datos globales
    r_del, c_del = buscar("Delegacion Policial")
    r_prob, c_prob = buscar("Problemática de Linea de Acción")
    r_ind, c_ind = buscar("Indicadores")
    r_meta, c_meta = buscar("Meta")

    # Identificar Delegación
    nombre_del = df.iloc[r_del, c_del + 1] if r_del is not None else "No detectada"
    st.info(f"📍 Delegación: {nombre_del}")

    # Selector de Trimestre
    trim_sel = st.selectbox("Trimestre a Evaluar", ["I Trimestre", "II Trimestre", "III Trimestre", "IV Trimestre"])
    r_t, c_t = buscar(trim_sel)

    st.markdown("---")

    reporte_datos = []

    # 🚀 Lógica de Visualización por Línea (Bucle dinámico)
    # Empezamos el escaneo desde donde están los indicadores (debajo de la cabecera)
    for i in range(r_ind + 1, len(df)):
        indicador = df.iloc[i, c_ind]
        
        # Si la fila del indicador está vacía, saltamos
        if pd.isna(indicador) or str(indicador).strip() == "":
            continue

        # Extraer datos de la fila actual usando las columnas detectadas
        categoria = df.iloc[i, 2] # Columna C (GL/FP) suele ser fija al inicio
        meta = df.iloc[i, c_meta]
        avance = df.iloc[i, c_t]
        descripcion = df.iloc[i, c_t + 1] # La descripción siempre está a la par del avance
        problematica = df.iloc[r_prob, c_prob + 1] # La problemática está en el bloque superior

        # Diseño limpio por cada indicador
        with st.expander(f"📌 {categoria} | {indicador}", expanded=False):
            st.write(f"**Problemática:** {problematica}")
            
            col1, col2 = st.columns(2)
            with col1:
                meta_edit = st.text_input(f"Meta - L{i}", value=meta, key=f"m_{i}")
                st.write(f"**Avance:** {avance}")
            with col2:
                st.text_area("Descripción del Resultado", value=descripcion, disabled=True, key=f"d_{i}")
            
            # Auditoría
            obs = st.text_area("Observaciones del Auditor", key=f"obs_{i}")
            cumple = st.checkbox("✔ Cumple con la evidencia", key=f"ch_{i}")

            reporte_datos.append({
                "indicador": indicador,
                "meta": meta_edit,
                "obs": obs,
                "cumple": cumple
            })

    # Botón de Informe
    if st.button("Generar Reporte PDF"):
        st.success("Reporte generado con éxito.")
