# app.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Formulario Línea de Acción", layout="wide")

# -----------------------------
# Estilos
# -----------------------------
st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}

.titulo-seccion {
    font-size: 20px;
    font-weight: bold;
    margin-bottom: 10px;
}

.etiqueta {
    font-size: 16px;
    font-weight: 500;
    margin-bottom: 4px;
}

.caja-borde {
    border: 1px solid #000000;
    padding: 10px;
    margin-top: 10px;
    margin-bottom: 10px;
}

.tabla-header {
    font-weight: bold;
    text-align: center;
    border: 1px solid black;
    padding: 8px;
    background-color: #f5f5f5;
}

.tabla-celda {
    border: 1px solid black;
    padding: 6px;
}

.linea {
    border-bottom: 1px solid black;
    height: 20px;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Título
# -----------------------------
st.markdown('<div class="titulo-seccion">Herramienta de Registro</div>', unsafe_allow_html=True)

# -----------------------------
# Encabezado
# -----------------------------
col1, col2 = st.columns([1, 4])
with col1:
    st.markdown('<div class="etiqueta">Delegación :</div>', unsafe_allow_html=True)
with col2:
    delegacion = st.text_input("", key="delegacion", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3, col4, col5, col6 = st.columns([1.2, 1.5, 1.2, 1.5, 1.5, 2])

with col1:
    st.markdown('<div class="etiqueta">línea de acción #:</div>', unsafe_allow_html=True)
with col2:
    linea_accion = st.text_input("", key="linea_accion", label_visibility="collapsed")

with col3:
    st.markdown('<div class="etiqueta">Problemática:</div>', unsafe_allow_html=True)
with col4:
    problematica = st.text_input("", key="problematica", label_visibility="collapsed")

with col5:
    st.markdown('<div class="etiqueta">Líder Estratégico:</div>', unsafe_allow_html=True)
with col6:
    lider = st.text_input("", key="lider", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 1, 6])
with col1:
    st.markdown('<div class="etiqueta">Trimestre:</div>', unsafe_allow_html=True)
with col2:
    trimestre = st.selectbox("", ["I", "II", "III", "IV"], key="trimestre", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

# -----------------------------
# Tabla editable
# -----------------------------
st.markdown("### Detalle")

df_inicial = pd.DataFrame({
    "Indicador": ["", "", ""],
    "Meta (editable)": ["", "", ""],
    "Avance": ["", "", ""],
    "Descripción": ["", "", ""],
    "Cantidad": ["", "", ""],
    "Observaciones (Editable)": ["", "", ""]
})

df_editado = st.data_editor(
    df_inicial,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    key="tabla_principal"
)

st.markdown("<br>", unsafe_allow_html=True)

# -----------------------------
# Botones
# -----------------------------
col1, col2, col3 = st.columns([1, 1, 5])

with col1:
    if st.button("Guardar"):
        st.success("Datos guardados temporalmente en la sesión.")
        st.write("**Delegación:**", delegacion)
        st.write("**Línea de acción #:**", linea_accion)
        st.write("**Problemática:**", problematica)
        st.write("**Líder Estratégico:**", lider)
        st.write("**Trimestre:**", trimestre)
        st.write("**Tabla:**")
        st.dataframe(df_editado, use_container_width=True)

with col2:
    if st.button("Limpiar"):
        st.rerun()
