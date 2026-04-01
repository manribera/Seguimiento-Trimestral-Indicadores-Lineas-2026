import streamlit as st
import pandas as pd
from fpdf import FPDF
import openpyxl

st.set_page_config(page_title="Auditoría Sembremos Seguridad", layout="wide")

st.title("📋 Lector de Informes de Avance")
st.write("Cargue los archivos para extraer la información según el formato de auditoría.")

archivos = st.file_uploader("📁 Subir archivos .xlsm", type=["xlsm"], accept_multiple_files=True)

def buscar_valor(sheet, texto_buscar):
    """Busca un texto en la hoja y retorna el valor de la celda de la derecha."""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and texto_buscar.upper() in str(cell.value).upper():
                return sheet.cell(row=cell.row, column=cell.column + 1).value
    return ""

if archivos:
    archivo_sel = st.selectbox("🎯 Seleccione archivo", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    # Cargar con openpyxl para lectura celda por celda
    wb = openpyxl.load_workbook(archivo_actual, data_only=True)
    ws = wb.active # Asume la primera hoja

    # --- EXTRACCIÓN DE ENCABEZADOS ---
    delegacion = buscar_valor(ws, "Delegación")
    linea_accion = buscar_valor(ws, "linea de accion #")
    problematica = buscar_valor(ws, "Problemática")
    lider = buscar_valor(ws, "Líder Estratégico")
    trimestre = buscar_valor(ws, "Trimestre")

    # --- INTERFAZ VISUAL (EMULANDO TU IMAGEN) ---
    st.markdown(f"### 🏢 Delegación: {delegacion}")
    
    c1, c2, c3 = st.columns(3)
    with c1: st.write(f"**Línea de Acción #:** {linea_accion}")
    with c2: st.write(f"**Problemática:** {problematica}")
    with c3: st.write(f"**Líder Estratégico:** {lider}")
    
    st.write(f"**Trimestre:** {trimestre}")
    st.markdown("---")

    # --- LECTORA DE TABLA DE INDICADORES ---
    # Usamos pandas para la parte masiva de la tabla
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)
    
    # Mapeo de columnas según tu formato
    # (Ajustado a los índices detectados en tus capturas previas)
    trim_indices = {
        "I": {"av": 9, "ds": 10, "cant": 11},
        "II": {"av": 14, "ds": 15, "cant": 16},
        "III": {"av": 19, "ds": 20, "cant": 21},
        "IV": {"av": 24, "ds": 25, "cant": 26}
    }
    
    idx = trim_indices.get(str(trimestre).strip(), trim_indices["I"])

    datos_auditoria = []
    
    # Buscamos dónde empieza la palabra "Indicador" para leer la tabla
    start_row = 10 
    for i, row in df.iterrows():
        if "INDICADOR" in str(row.values).upper():
            start_row = i + 1
            break

    # Generar filas de la tabla
    for i in range(start_row, len(df)):
        indicador = df.iloc[i, 6] # Col G
        if pd.isna(indicador): continue
        
        meta = df.iloc[i, 8] # Col I
        avance = df.iloc[i, idx["av"]]
        desc = df.iloc[i, idx["ds"]]
        cant = df.iloc[i, idx["cant"]]

        # Lógica de Avance Automático
        if pd.notna(cant) and str(cant).strip() != "" and cant != 0:
            estado = "✅ Con Actividades"
        else:
            estado = "❌ Sin Actividades"

        # Visualización en formato de filas
        with st.container():
            col_ind, col_met, col_av, col_ds, col_cant, col_obs = st.columns([2, 1, 1, 2, 1, 2])
            
            col_ind.write(indicador)
            meta_edit = col_met.text_input("Meta", value=meta, key=f"m_{i}")
            col_av.write(estado)
            col_ds.write(desc)
            col_cant.write(cant)
            obs_edit = col_obs.text_area("Observaciones", key=f"o_{i}")
            
            st.markdown("---")
