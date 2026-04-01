import streamlit as st
import pandas as pd
import openpyxl
from fpdf import FPDF

# Configuración de página ancha para que la tabla sea legible
st.set_page_config(page_title="Auditoría Sembremos Seguridad", layout="wide")

st.markdown("## 📋 Herramienta de Auditoría Técnica")
st.markdown("---")

# Carga de archivos (puedes subir los 98 de una vez)
archivos = st.file_uploader("📁 Cargar informes de delegaciones (.xlsm)", type=["xlsm"], accept_multiple_files=True)

if archivos:
    archivo_sel = st.selectbox("🎯 Seleccione la unidad a auditar", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    # Lector de celdas (openpyxl) para encontrar los títulos de tu imagen
    wb = openpyxl.load_workbook(archivo_actual, data_only=True)
    ws = wb.active

    def buscar_dato(texto):
        for row in ws.iter_rows(max_row=30):
            for cell in row:
                if cell.value and texto.upper() in str(cell.value).upper():
                    return ws.cell(row=cell.row, column=cell.column + 1).value
        return ""

    # --- ENCABEZADO (Basado fielmente en tu imagen) ---
    st.write(f"**Delegación:** `{buscar_dato('Delegación')}`")
    
    h1, h2, h3 = st.columns([1, 2, 2])
    h1.write(f"**línea de accion #:** {buscar_dato('linea de accion #')}")
    h2.write(f"**Problemática:** {buscar_dato('Problemática')}")
    h3.write(f"**Líder Estrategico:** {buscar_dato('Líder Estrategico')}")
    
    st.write(f"**Trimestre:** {buscar_dato('Trimestre')}")
    st.markdown("---")

    # --- TABLA DE INDICADORES (Mismo orden que tu imagen) ---
    # Encabezados visuales
    t_col1, t_col2, t_col3, t_col4, t_col5, t_col6 = st.columns([2, 1, 1, 2, 1, 2])
    t_col1.write("**Indicador**")
    t_col2.write("**Meta (editable)**")
    t_col3.write("**Avance**")
    t_col4.write("**Descripción**")
    t_col5.write("**Cantidad**")
    t_col6.write("**Observaciones (Editable)**")
    st.markdown("---")

    # Lectura de datos con Pandas para la tabla continua
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)
    
    # Mapeo de columnas por trimestre (Bloques de 5 según tu Excel original)
    tri_val = str(buscar_dato('Trimestre')).strip()
    mapa = {"I": 9, "II": 14, "III": 19, "IV": 24} # Columna de inicio
    col_inicio = mapa.get(tri_val, 9)

    for i in range(10, len(df)):
        indicador = df.iloc[i, 6] # Columna G
        if pd.isna(indicador) or "Indicador" in str(indicador): continue

        meta_val = df.iloc[i, 8] # Columna I
        desc_val = df.iloc[i, col_inicio + 1] # Descripción
        cant_val = df.iloc[i, col_inicio + 2] # Cantidad

        # Lógica de Avance Automático: Si Cantidad > 0, es positivo
        avance_visual = "🟢 Con Actividad" if pd.notna(cant_val) and cant_val != 0 else "🔴 Sin Actividad"

        # Fila de la Herramienta (Emulación del formato)
        with st.container():
            c1, c2, c3, c4, c5, c6 = st.columns([2, 1, 1, 2, 1, 2])
            c1.write(indicador)
            c2.text_input("Meta", value=meta_val, key=f"m_{i}", label_visibility="collapsed")
            c3.write(avance_visual)
            c4.write(desc_val)
            c5.write(cant_val)
            c6.text_area("Obs", key=f"o_{i}", height=70, label_visibility="collapsed")
            st.markdown("---")

    if st.button("📄 Generar Reporte Final"):
        st.success("Auditoría lista para descarga.")
