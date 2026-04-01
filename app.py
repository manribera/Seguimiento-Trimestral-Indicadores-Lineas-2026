import streamlit as st
import pandas as pd
import openpyxl
from fpdf import FPDF
import io

# Configuración de página ancha para emular el formato de tabla de la imagen
st.set_page_config(page_title="Auditoría Sembremos Seguridad", layout="wide")

st.markdown("## 📋 Herramienta de Auditoría Técnica")
st.markdown("---")

# Carga múltiple para procesar los archivos de las delegaciones
archivos = st.file_uploader("📁 Cargar informes (.xlsm)", type=["xlsm"], accept_multiple_files=True)

if archivos:
    archivo_sel = st.selectbox("🎯 Seleccione la unidad a auditar", [a.name for a in archivos])
    archivo_actual = next(a for a in archivos if a.name == archivo_sel)
    
    # Lector dinámico (openpyxl) para encontrar los títulos exactos de tu imagen
    wb = openpyxl.load_workbook(archivo_actual, data_only=True)
    ws = wb.active

    def buscar_dato_derecha(texto_clave):
        """Busca una palabra y devuelve el valor de la celda inmediata a la derecha."""
        for row in ws.iter_rows(max_row=50, max_col=20):
            for cell in row:
                if cell.value and texto_clave.upper() in str(cell.value).upper():
                    return ws.cell(row=cell.row, column=cell.column + 1).value
        return ""

    # --- ENCABEZADO: Mimetismo con el formato de la imagen ---
    delegacion = buscar_dato_derecha("Delegación")
    linea_n = buscar_dato_derecha("linea de accion #")
    prob = buscar_dato_derecha("Problemática")
    lider = buscar_dato_derecha("Líder Estrategico")
    trimestre = str(buscar_dato_derecha("Trimestre")).replace(".0", "").strip()

    # Visualización del bloque superior
    st.write(f"**Delegación:** `{delegacion}`")
    
    c1, c2, c3 = st.columns([1, 2, 2])
    c1.markdown(f"**línea de accion #:** {linea_n}")
    c2.markdown(f"**Problemática:** {prob}")
    c3.markdown(f"**Líder Estratégico:** {lider}")
    
    st.write(f"**Trimestre:** {trimestre}")
    st.markdown("---")

    # --- TABLA DE AUDITORÍA (Mismo orden que la imagen) ---
    # Encabezados de la cuadrícula
    h_ind, h_met, h_av, h_ds, h_cant, h_obs = st.columns([2, 1, 1, 2, 1, 2])
    h_ind.write("**Indicador**")
    h_met.write("**Meta (editable)**")
    h_av.write("**Avance**")
    h_ds.write("**Descripción**")
    h_cant.write("**Cantidad**")
    h_obs.write("**Observaciones (Editable)**")
    st.markdown("---")

    # Procesamiento de la tabla continua con Pandas
    df = pd.read_excel(archivo_actual, sheet_name=0, header=None)
    
    # Mapa de columnas por trimestre (J, O, T, Y)
    mapa_col = {"I": 9, "II": 14, "III": 19, "IV": 24, "1": 9, "2": 14, "3": 19, "4": 24}
    base = mapa_col.get(trimestre, 9)

    reporte_final = []

    # Buscamos la fila de inicio de indicadores
    fila_inicio = 10
    for i, row in df.iterrows():
        if "INDICADOR" in str(row.values).upper():
            fila_inicio = i + 1
            break

    # Renderizado de filas estilo tabla
    for i in range(fila_inicio, len(df)):
        indicador = df.iloc[i, 6] # Col G
        if pd.isna(indicador) or "Indicador" in str(indicador): continue

        meta_exc = df.iloc[i, 8] # Col I
        desc_exc = df.iloc[i, base + 1] # Descripción
        cant_exc = df.iloc[i, base + 2] # Cantidad

        # Lógica: Si hay datos en Cantidad, el avance es positivo
        if pd.notna(cant_exc) and str(cant_exc).strip() != "" and cant_exc != 0:
            status = "🟢 Con Actividad"
        else:
            status = "🔴 Sin Actividad"

        # Fila de la herramienta consolidada
        with st.container():
            f_ind, f_met, f_av, f_ds, f_cant, f_obs = st.columns([2, 1, 1, 2, 1, 2])
            
            f_ind.write(indicador)
            meta_e = f_met.text_input("Meta", value=meta_exc, key=f"m_{i}", label_visibility="collapsed")
            f_av.write(status)
            f_ds.write(desc_exc)
            f_cant.write(cant_exc)
            obs_e = f_obs.text_area("Notas", key=f"o_{i}", height=70, label_visibility="collapsed")
            
            reporte_final.append({
                "titulo": f"{linea_n} - {prob}",
                "indicador": indicador, "meta": meta_e, "obs": obs_e, "avance": status
            })
            st.markdown("---")

    if st.button("📄 Generar Reporte Final (PDF)"):
        st.success(f"Auditoría de {delegacion} procesada exitosamente.")
