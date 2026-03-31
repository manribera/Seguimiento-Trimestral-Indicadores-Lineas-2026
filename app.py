import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Estrategia Jurídica", layout="wide")
st.markdown("### 🔍 Instrumento de Seguimiento Dinámico")

archivo = st.file_uploader("Subir archivo de la delegación", type=["xlsm"])

if archivo:
    # Leer la hoja de forma integral para que el comparador la analice
    df = pd.read_excel(archivo, sheet_name=0, header=None)

    # 🕵️ FUNCION DEL COMPARADOR: Buscar coordenadas por texto
    def localizar_texto(dataframe, texto):
        for r in range(len(dataframe)):
            for c in range(dataframe.shape[1]):
                if texto.upper() in str(dataframe.iloc[r, c]).upper():
                    return r, c
        return None, None

    # 1. Buscar Nombre de Delegación
    r_del, c_del = localizar_texto(df, "Delegacion Policial")
    nombre_delegacion = df.iloc[r_del, c_del + 1] if r_del is not None else "No detectada"
    st.info(f"📍 Delegación: {nombre_delegacion}")

    # 2. Localizar Columnas Clave (Indicadores, Metas, Problemática)
    r_ind, col_indicador = localizar_texto(df, "Indicadores")
    _, col_meta = localizar_texto(df, "Meta")
    _, col_problematica = localizar_texto(df, "Problematica")
    _, col_cat = localizar_texto(df, "GL") # Busca el anclaje de categoría

    # 3. Localizar Trimestres (Búsqueda por bloque)
    trim_opcion = st.selectbox("Trimestre a Evaluar", ["I Trimestre", "II Trimestre", "III Trimestre", "IV Trimestre"])
    r_trim, col_avance = localizar_texto(df, trim_opcion)
    
    # El resultado (Descripción) suele estar a la derecha del Avance
    col_descripcion = col_avance + 1 if col_avance is not None else None

    st.markdown("---")

    # 4. Construcción del Seguimiento por Línea
    if col_indicador is not None:
        # Empezamos a leer desde la fila debajo de la cabecera 'Indicadores'
        for i in range(r_ind + 1, len(df)):
            indicador = df.iloc[i, col_indicador]
            
            # Detener si la fila está vacía o es el final de la tabla
            if pd.isna(indicador) or str(indicador).strip() == "":
                continue

            # Extraer datos usando las columnas localizadas por el comparador
            categoria = df.iloc[i, col_cat] if col_cat is not None else "N/A"
            problematica = df.iloc[i, col_problematica] if col_problematica is not None else "N/A"
            meta = df.iloc[i, col_meta] if col_meta is not None else "N/A"
            resultado_avance = df.iloc[i, col_avance] if col_avance is not None else "Sin dato"
            resultado_desc = df.iloc[i, col_descripcion] if col_descripcion is not None else "Sin descripción"

            # Visualización Profesional
            with st.expander(f"📌 {categoria} | {indicador}", expanded=False):
                st.write(f"**Probleática:** {problematica}")
                
                c1, c2 = st.columns(2)
                with c1:
                    st.text_input(f"Meta - L{i}", value=meta, key=f"m_{i}")
                    st.write(f"**Estado de Avance:** {resultado_avance}")
                with c2:
                    st.text_area(f"Descripción del Resultado", value=resultado_desc, disabled=True, key=f"d_{i}")
                
                st.text_area("Observaciones de Seguimiento Legales/Técnicas", key=f"obs_{i}")
                st.checkbox("Evidencia Verificada", key=f"chk_{i}")

    if st.button("Generar Informe Consolidado"):
        st.write("Generando reporte de auditoría...")
