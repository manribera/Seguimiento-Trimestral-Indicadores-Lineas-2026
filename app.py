# app.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Informe de Avance - Lector Dinámico", layout="wide")


# =========================================================
# UTILIDADES
# =========================================================
def norm_text(value) -> str:
    """Normaliza texto para comparar por coincidencia."""
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def contains_any(text: str, patterns) -> bool:
    t = norm_text(text)
    return any(p in t for p in patterns)


def cell_value(ws, row, col):
    return ws.cell(row=row, column=col).value


def row_values(ws, row, max_col):
    return [ws.cell(row=row, column=c).value for c in range(1, max_col + 1)]


def find_best_main_sheet(wb):
    """
    Busca la hoja principal por coincidencias de contenido.
    No usa nombre fijo; evalúa varias posibilidades.
    """
    candidates = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        score = 0

        max_row = min(ws.max_row, 80)
        max_col = min(ws.max_column, 30)

        for r in range(1, max_row + 1):
            vals = row_values(ws, r, max_col)
            row_text = " | ".join("" if v is None else str(v) for v in vals)
            t = norm_text(row_text)

            if "linea de accion" in t or "línea de acción" in t:
                score += 5
            if "problemática" in t or "problematica" in t:
                score += 4
            if "lider estrategico" in t or "líder estratégico" in t or "lider" in t:
                score += 4
            if "indicadores" in t or "indicador" in t:
                score += 3
            if "meta" in t:
                score += 2
            if "avance" in t:
                score += 2
            if "descripción" in t or "descripcion" in t:
                score += 2

        candidates.append((sheet_name, score))

    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0] if candidates else wb.sheetnames[0]


def find_line_action_starts(ws):
    """
    Encuentra el inicio de cada bloque por coincidencia con:
    'Linea de Accion' o 'Línea de Acción'
    """
    starts = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            t = norm_text(v)

            if "linea de accion" in t or "línea de acción" in t:
                starts.append((r, c, v))
                break

    # quitar duplicados por fila
    cleaned = []
    seen_rows = set()
    for item in starts:
        if item[0] not in seen_rows:
            cleaned.append(item)
            seen_rows.add(item[0])

    return cleaned


def detect_header_row(ws, start_row, end_row):
    """
    Busca la fila donde aparecen encabezados como:
    Indicador, Meta, Avance, Descripción
    """
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        vals = row_values(ws, r, ws.max_column)
        row_text = " | ".join("" if v is None else str(v) for v in vals)
        t = norm_text(row_text)

        has_ind = "indicador" in t or "indicadores" in t
        has_meta = "meta" in t
        has_avance = "avance" in t
        has_desc = "descripcion" in t or "descripción" in t

        if has_ind and has_meta and has_avance and has_desc:
            return r

    return None


def get_trimester_from_row_text(text):
    t = norm_text(text)
    if "i trimestre" in t or "1 trimestre" in t or "primer trimestre" in t:
        return "I"
    if "ii trimestre" in t or "2 trimestre" in t or "segundo trimestre" in t:
        return "II"
    if "iii trimestre" in t or "3 trimestre" in t or "tercer trimestre" in t:
        return "III"
    if "iv trimestre" in t or "4 trimestre" in t or "cuarto trimestre" in t:
        return "IV"
    return ""


def search_value_near_keywords(ws, start_row, end_row, keywords):
    """
    Busca una fila que contenga una palabra clave y devuelve
    el valor más probable a la derecha.
    """
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        for c in range(1, ws.max_column):
            v = ws.cell(r, c).value
            if contains_any(v, keywords):
                # buscar hacia la derecha el primer valor útil
                for c2 in range(c + 1, min(c + 8, ws.max_column) + 1):
                    val = ws.cell(r, c2).value
                    if val not in (None, ""):
                        return val

    return ""


def get_delegacion(ws):
    """
    Busca delegación/global dentro de la hoja.
    """
    for r in range(1, min(ws.max_row, 40) + 1):
        for c in range(1, min(ws.max_column, 20) + 1):
            v = ws.cell(r, c).value
            if contains_any(v, ["delegacion policial", "delegación policial", "delegacion", "delegación"]):
                for c2 in range(c + 1, min(c + 6, ws.max_column) + 1):
                    val = ws.cell(r, c2).value
                    if val not in (None, ""):
                        return str(val)
    return ""


def extract_table(ws, header_row, block_end_row):
    """
    Extrae la tabla bajo los encabezados, sin depender de columnas fijas.
    """
    header_cells = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(header_row, c).value
        t = norm_text(val)

        if "indicador" in t:
            header_cells["indicador"] = c
        elif "meta" in t:
            header_cells["meta"] = c
        elif "avance" in t:
            header_cells["avance"] = c
        elif "descripcion" in t or "descripción" in t:
            header_cells["descripcion"] = c
        elif "cantidad" in t:
            header_cells["cantidad"] = c
        elif "observ" in t:
            header_cells["observaciones"] = c

    needed = ["indicador", "meta", "avance", "descripcion"]
    if not all(k in header_cells for k in needed):
        return pd.DataFrame(columns=[
            "Indicador", "Meta (editable)", "Avance",
            "Descripción", "Cantidad", "Observaciones (Editable)"
        ])

    data = []
    for r in range(header_row + 1, block_end_row + 1):
        indicador = ws.cell(r, header_cells["indicador"]).value if "indicador" in header_cells else ""
        meta = ws.cell(r, header_cells["meta"]).value if "meta" in header_cells else ""
        avance = ws.cell(r, header_cells["avance"]).value if "avance" in header_cells else ""
        descripcion = ws.cell(r, header_cells["descripcion"]).value if "descripcion" in header_cells else ""
        cantidad = ws.cell(r, header_cells["cantidad"]).value if "cantidad" in header_cells else ""
        observaciones = ws.cell(r, header_cells["observaciones"]).value if "observaciones" in header_cells else ""

        row_has_content = any(x not in (None, "") for x in [indicador, meta, avance, descripcion, cantidad, observaciones])

        # cortar si vienen muchas filas vacías seguidas
        if not row_has_content:
            # revisa si las próximas 2 también están vacías
            next_empty = True
            for rr in range(r, min(r + 2, block_end_row) + 1):
                vals = []
                for field, col in header_cells.items():
                    vals.append(ws.cell(rr, col).value)
                if any(v not in (None, "") for v in vals):
                    next_empty = False
                    break
            if next_empty:
                break

        if row_has_content:
            data.append({
                "Indicador": indicador if indicador is not None else "",
                "Meta (editable)": meta if meta is not None else "",
                "Avance": avance if avance is not None else "",
                "Descripción": descripcion if descripcion is not None else "",
                "Cantidad": cantidad if cantidad is not None else "",
                "Observaciones (Editable)": observaciones if observaciones is not None else "",
            })

    return pd.DataFrame(data)


def extract_blocks_from_sheet(ws):
    """
    Extrae todos los bloques de líneas de acción encontrados por coincidencia.
    """
    starts = find_line_action_starts(ws)

    if not starts:
        return []

    blocks = []
    delegacion = get_delegacion(ws)

    for i, (start_row, start_col, start_text) in enumerate(starts):
        end_row = starts[i + 1][0] - 1 if i + 1 < len(starts) else ws.max_row

        # Texto de la fila de inicio
        row_txt = " | ".join("" if v is None else str(v) for v in row_values(ws, start_row, ws.max_column))
        trimestre = get_trimester_from_row_text(row_txt)

        # Número / nombre de línea
        linea_accion = str(start_text) if start_text not in (None, "") else f"Línea {i+1}"

        # Problemática y líder por coincidencia
        problematica = search_value_near_keywords(
            ws, start_row, min(start_row + 8, end_row),
            ["problemática", "problematica"]
        )

        lider = search_value_near_keywords(
            ws, start_row, min(start_row + 8, end_row),
            ["lider estrategico", "líder estratégico", "lider", "líder"]
        )

        # Buscar fila de encabezados
        header_row = detect_header_row(ws, start_row, min(start_row + 15, end_row))
        if header_row is None:
            continue

        tabla = extract_table(ws, header_row, end_row)

        blocks.append({
            "delegacion": delegacion,
            "linea_accion": linea_accion,
            "problematica": problematica,
            "lider": lider,
            "trimestre": trimestre,
            "tabla": tabla
        })

    return blocks


# =========================================================
# ESTILOS
# =========================================================
st.markdown("""
<style>
.block-container {
    max-width: 1300px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}

.form-box {
    border: 1px solid #000;
    padding: 14px;
    margin-top: 10px;
    margin-bottom: 14px;
}

.section-title {
    font-size: 24px;
    font-weight: 700;
    margin-bottom: 0.6rem;
}

.label-like {
    font-size: 17px;
    font-weight: 500;
    margin-bottom: 4px;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# INTERFAZ
# =========================================================
st.title("Herramienta de lectura dinámica del instrumento")

st.write(
    "Sube el archivo Excel. La herramienta buscará las líneas de acción por coincidencia "
    "dentro del instrumento, no por filas ni columnas fijas."
)

uploaded_file = st.file_uploader(
    "Arrastra y suelta aquí el archivo .xlsm o .xlsx",
    type=["xlsm", "xlsx"]
)

if uploaded_file is not None:
    try:
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False, keep_vba=True)

        main_sheet_name = find_best_main_sheet(wb)
        ws = wb[main_sheet_name]

        blocks = extract_blocks_from_sheet(ws)

        if not blocks:
            st.warning("No encontré bloques de líneas de acción en el archivo.")
            st.stop()

        opciones = []
        for i, b in enumerate(blocks, start=1):
            etiqueta = f"{i}. {b['linea_accion']} | {b['problematica']}"
            opciones.append(etiqueta)

        idx = st.selectbox(
            "Seleccione la línea de acción encontrada en el archivo",
            options=list(range(len(opciones))),
            format_func=lambda x: opciones[x]
        )

        bloque = blocks[idx]

        st.success(f"Hoja detectada: {main_sheet_name}")

        # -------------------------------------------------
        # FORMATO PARECIDO AL MODELO QUE MOSTRASTE
        # -------------------------------------------------
        st.markdown('<div class="form-box">', unsafe_allow_html=True)

        c1, c2 = st.columns([1.2, 4])
        with c1:
            st.markdown('<div class="label-like">Delegación :</div>', unsafe_allow_html=True)
        with c2:
            st.text_input(
                "",
                value=bloque["delegacion"],
                key="delegacion",
                label_visibility="collapsed"
            )

        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2, c3, c4, c5, c6 = st.columns([1.3, 1.6, 1.2, 2, 1.5, 2])
        with c1:
            st.markdown('<div class="label-like">línea de acción #:</div>', unsafe_allow_html=True)
        with c2:
            st.text_input(
                "",
                value=bloque["linea_accion"],
                key="linea_accion",
                label_visibility="collapsed"
            )

        with c3:
            st.markdown('<div class="label-like">Problemática:</div>', unsafe_allow_html=True)
        with c4:
            st.text_input(
                "",
                value=bloque["problematica"],
                key="problematica",
                label_visibility="collapsed"
            )

        with c5:
            st.markdown('<div class="label-like">Líder Estratégico:</div>', unsafe_allow_html=True)
        with c6:
            st.text_input(
                "",
                value=bloque["lider"],
                key="lider",
                label_visibility="collapsed"
            )

        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2, c3 = st.columns([1, 1, 6])
        with c1:
            st.markdown('<div class="label-like">Trimestre:</div>', unsafe_allow_html=True)
        with c2:
            trim_options = ["", "I", "II", "III", "IV"]
            trim_value = bloque["trimestre"] if bloque["trimestre"] in trim_options else ""
            selected_trim = st.selectbox(
                "",
                trim_options,
                index=trim_options.index(trim_value),
                key="trimestre",
                label_visibility="collapsed"
            )

        st.markdown("<br>", unsafe_allow_html=True)

        df = bloque["tabla"].copy()

        if df.empty:
            df = pd.DataFrame([{
                "Indicador": "",
                "Meta (editable)": "",
                "Avance": "",
                "Descripción": "",
                "Cantidad": "",
                "Observaciones (Editable)": ""
            }])

        st.markdown("### Detalle")
        df_editado = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="detalle_tabla"
        )

        st.markdown("</div>", unsafe_allow_html=True)

        # -------------------------------------------------
        # RESUMEN DE EXTRACCIÓN
        # -------------------------------------------------
        with st.expander("Ver resumen técnico de la extracción"):
            st.write(f"**Hoja detectada:** {main_sheet_name}")
            st.write(f"**Bloques encontrados:** {len(blocks)}")
            st.write(f"**Delegación detectada:** {bloque['delegacion']}")
            st.write(f"**Línea seleccionada:** {bloque['linea_accion']}")
            st.write(f"**Problemática detectada:** {bloque['problematica']}")
            st.write(f"**Líder detectado:** {bloque['lider']}")
            st.write(f"**Trimestre detectado:** {selected_trim}")

        # -------------------------------------------------
        # EXPORTACIÓN SIMPLE
        # -------------------------------------------------
        salida = BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            encabezado = pd.DataFrame({
                "Campo": [
                    "Delegación",
                    "Línea de acción",
                    "Problemática",
                    "Líder Estratégico",
                    "Trimestre"
                ],
                "Valor": [
                    bloque["delegacion"],
                    bloque["linea_accion"],
                    bloque["problematica"],
                    bloque["lider"],
                    selected_trim
                ]
            })
            encabezado.to_excel(writer, index=False, sheet_name="Resumen")
            df_editado.to_excel(writer, index=False, sheet_name="Detalle")

        st.download_button(
            "Descargar extracción en Excel",
            data=salida.getvalue(),
            file_name="extraccion_linea_accion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al leer el archivo: {e}")

else:
    st.info("Sube un archivo para comenzar.")
