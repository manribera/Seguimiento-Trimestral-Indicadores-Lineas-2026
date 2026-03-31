import streamlit as st
import pandas as pd

st.set_page_config(page_title="Estrategia Jurídica - Auditoría Total", layout="wide")
st.markdown("### 📋 Seguimiento Integral de Líneas de Acción")

archivo = st.file_uploader("Cargar libro de delegación", type=["xlsm"])

if archivo:
    # Leer la primera hoja donde están los datos (índice 0)
    df = pd.read_excel(archivo, sheet_name=0, header=None)

    # 1. Localizar dinámicamente la fila donde empiezan los indicadores
    # Buscamos la celda que contiene "GL" para saber dónde empezar
    inicio_fila = None
    for i, row in df.iterrows():
        if "GL" in row.values:
            inicio_fila = i
            break

    if inicio_fila is not None:
        # Extraemos los datos desde esa fila hacia abajo
        datos = df.iloc[inicio_fila:].copy()
        
        # 2. Definir Columnas según tu imagen (Mapeo dinámico)
        # Col G (Indice 6): Indicadores | Col I (Indice 8): Meta
        # Bloque I Trimestre: Col J (Avance), Col K (Descripción)
        
        trimestre_opciones = {
            "I Trimestre": {"avance": 9, "desc": 10},
            "II Trimestre": {"avance": 14, "desc": 15}, # Ajustado por bloques de 5
            "III Trimestre": {"avance": 19, "desc": 20},
            "IV Trimestre": {"avance": 24, "desc": 25}
        }

        trim_sel = st.selectbox("Seleccione Trimestre para Seguimiento", list(trimestre_opciones.keys()))
        col_idx = trimestre_opciones[trim_sel]

        st.markdown("---")

        # 3. Bucle para visualizar TODAS las líneas encontradas
        for index, fila in datos.iterrows():
            indicador_texto = fila[6] # Columna G
            meta_texto = fila[8]      # Columna I
            avance_valor = fila[col_idx["avance"]]
            desc_valor = fila[col_idx["desc"]]

            # Detener el bucle si llegamos a una fila vacía de indicadores
            if pd.isna(indicador_texto):
                continue

            with st.expander(f"🔍 {indicador_texto}", expanded=False):
                col_a, col_b = st.columns(2)
                
                with col_a:
                    st.text_input(f"Meta - L{index}", value=meta_texto, key=f"meta_{index}")
                    st.write(f"**Avance Reportado:** {avance_valor}")
                
                with col_b:
                    st.text_area(f"Descripción - L{index}", value=desc_valor, disabled=True, key=f"desc_{index}")
                
                st.text_area("Observaciones de Auditoría", key=f"obs_{index}", placeholder="Añada aquí el seguimiento...")
                st.checkbox("Evidencia verificada", key=f"check_{index}")

    else:
        st.error("No se pudo localizar la columna de indicadores (GL/FP). Revise el formato del archivo.")
