# app.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="Lector de Instrumentos", layout="wide")


# =========================================================
# UTILIDADES DE TEXTO
# =========================================================
def norm_text(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def contains_any(text, patterns) -> bool:
    t = norm_text(text)
    return any(p in t for p in patterns)


def is_meaningful_value(v) -> bool:
    if v is None:
        return False
    if str(v).strip() == "":
        return False
    return True


# =========================================================
# UTILIDADES DE EXCEL
# =========================================================
def row_values(ws, row, max_col=None):
    if max_col is None:
        max_col = ws.max_column
    return [ws.cell(row=row, column=c).value for c in range(1, max_col + 1)]


def row_text(ws, row, max_col=None):
    vals = row_values(ws, row, max_col)
    return " | ".join("" if v is None else str(v) for v in vals)


def get_right_value(ws, row, col, max_steps=10):
    for c in range(col + 1, min(ws.max_column, col + max_steps) + 1):
        val = ws.cell(row=row, column=c).value
        if is_meaningful_value(val):
            return val
    return ""


def get_down_value(ws, row, col, max_steps=4):
    for r in range(row + 1, min(ws.max_row, row + max_steps) + 1):
        val = ws.cell(row=r, column=col).value
        if is_meaningful_value(val):
            return val
    return ""


def get_near_value(ws, row, col):
    """
    Busca primero a la derecha y luego abajo.
    """
    val = get_right_value(ws, row, col, max_steps=12)
    if is_meaningful_value(val):
        return val

    val = get_down_value(ws, row, col, max_steps=4)
    if is_meaningful_value(val):
        return val

    return ""


# =========================================================
# DETECCIÓN DE HOJA PRINCIPAL
# =========================================================
def find_best_main_sheet(wb):
    candidates = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        score = 0

        max_row = min(ws.max_row, 150)
        max_col = min(ws.max_column, 40)

        for r in range(1, max_row + 1):
            txt = norm_text(row_text(ws, r, max_col))

            if "línea de acción" in txt or "linea de accion" in txt:
                score += 8
            if "problemática" in txt or "problematica" in txt:
                score += 6
            if "líder estratégico" in txt or "lider estrategico" in txt:
                score += 6
            if "indicador" in txt:
                score += 4
            if "meta" in txt:
                score += 3
            if "avance" in txt:
                score += 3
            if "descripción" in txt or "descripcion" in txt:
                score += 3
            if "cantidad" in txt:
                score += 3
            if "observaciones" in txt or "observación" in txt:
                score += 2
            if "delegación" in txt or "delegacion" in txt:
                score += 3
            if "trimestre" in txt:
                score += 2

        candidates.append((sheet_name, score))

    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0] if candidates else wb.sheetnames[0]


# =========================================================
# DELEGACIÓN
# =========================================================
def get_delegacion(ws):
    for r in range(1, min(ws.max_row, 60) + 1):
        for c in range(1, min(ws.max_column, 20) + 1):
            val = ws.cell(r, c).value
            if contains_any(val, ["delegación", "delegacion"]):
                derecha = get_right_value(ws, r, c, max_steps=8)
                if is_meaningful_value(derecha):
                    return clean_text(derecha)

                abajo = get_down_value(ws, r, c, max_steps=3)
                if is_meaningful_value(abajo):
                    return clean_text(abajo)
    return ""


# =========================================================
# EXTRAER NÚMERO / VALOR DE LÍNEA DE ACCIÓN
# =========================================================
def extract_line_number_from_area(ws, start_row, start_col):
    """
    Busca el número o valor real cerca del texto 'línea de acción'.
    """
    # 1. Mismo renglón, hacia la derecha
    for c in range(start_col + 1, min(ws.max_column, start_col + 10) + 1):
        val = ws.cell(start_row, c).value
        if is_meaningful_value(val):
            txt = clean_text(val)
            if "línea de acción" not in norm_text(txt) and "linea de accion" not in norm_text(txt):
                return txt

    # 2. Renglones cercanos
    for r in range(start_row, min(ws.max_row, start_row + 3) + 1):
        for c in range(start_col, min(ws.max_column, start_col + 10) + 1):
            val = ws.cell(r, c).value
            if is_meaningful_value(val):
                txt = clean_text(val)
                if "línea de acción" not in norm_text(txt) and "linea de accion" not in norm_text(txt):
                    return txt

    return ""


# =========================================================
# DETECTAR TODOS LOS BLOQUES
# =========================================================
def find_line_action_starts(ws):
    starts = []

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            t = norm_text(val)

            if "línea de acción" in t or "linea de accion" in t:
                line_value = extract_line_number_from_area(ws, r, c)
                starts.append({
                    "row": r,
                    "col": c,
                    "line_value": line_value
                })
                break

    # limpiar duplicados cercanos
    cleaned = []
    last_row = -999

    for item in starts:
        if item["row"] - last_row > 2:
            cleaned.append(item)
            last_row = item["row"]

    return cleaned


# =========================================================
# EXTRAER CAMPOS DE CADA BLOQUE
# =========================================================
def search_value_near_keywords(ws, start_row, end_row, keywords):
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if contains_any(val, keywords):
                cerca = get_near_value(ws, r, c)
                if is_meaningful_value(cerca):
                    return clean_text(cerca)
    return ""


def detect_trimester(ws, start_row, end_row):
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        txt = norm_text(row_text(ws, r))

        if "trimestre" in txt:
            # buscar valor directo en las celdas
            for c in range(1, ws.max_column + 1):
                v = clean_text(ws.cell(r, c).value)
                if v in ["I", "II", "III", "IV"]:
                    return v

            # detección por texto
            if " iv" in f" {txt}" or "cuarto" in txt or "4" in txt:
                return "IV"
            if " iii" in f" {txt}" or "tercer" in txt or "3" in txt:
                return "III"
            if " ii" in f" {txt}" or "segundo" in txt or "2" in txt:
                return "II"
            if " i" in f" {txt}" or "primer" in txt or "1" in txt:
                return "I"

    return ""


# =========================================================
# DETECTAR ENCABEZADOS DE TABLA
# =========================================================
def detect_header_row(ws, start_row, end_row):
    best_row = None
    best_score = -1

    for r in range(start_row, min(end_row, ws.max_row) + 1):
        vals = row_values(ws, r)

        found = {
            "indicador": False,
            "meta": False,
            "avance": False,
            "descripcion": False,
            "cantidad": False,
            "observaciones": False,
        }

        for v in vals:
            t = norm_text(v)
            if "indicador" in t:
                found["indicador"] = True
            if "meta" in t:
                found["meta"] = True
            if "avance" in t:
                found["avance"] = True
            if "descripcion" in t or "descripción" in t:
                found["descripcion"] = True
            if "cantidad" in t:
                found["cantidad"] = True
            if "observ" in t:
                found["observaciones"] = True

        score = sum(found.values())

        if found["indicador"] and found["meta"] and score > best_score:
            best_score = score
            best_row = r

    return best_row


def map_headers(ws, header_row):
    header_map = {}

    for c in range(1, ws.max_column + 1):
        val = ws.cell(header_row, c).value
        t = norm_text(val)

        if "indicador" in t:
            header_map["Indicador"] = c
        elif "meta" in t:
            header_map["Meta (editable)"] = c
        elif "avance" in t:
            header_map["Avance"] = c
        elif "descripcion" in t or "descripción" in t:
            header_map["Descripción"] = c
        elif "cantidad" in t:
            header_map["Cantidad"] = c
        elif "observ" in t:
            header_map["Observaciones (Editable)"] = c

    return header_map


# =========================================================
# EXTRAER TABLA
# =========================================================
def extract_table(ws, header_row, block_end_row):
    columns = [
        "Indicador",
        "Meta (editable)",
        "Avance",
        "Descripción",
        "Cantidad",
        "Observaciones (Editable)"
    ]

    header_map = map_headers(ws, header_row)

    if "Indicador" not in header_map:
        return pd.DataFrame(columns=columns)

    data = []
    empty_count = 0

    for r in range(header_row + 1, block_end_row + 1):
        row_data = {}

        for col_name in columns:
            col_idx = header_map.get(col_name)
            if col_idx:
                row_data[col_name] = ws.cell(r, col_idx).value
            else:
                row_data[col_name] = ""

        for k in row_data:
            if row_data[k] is None:
                row_data[k] = ""

        has_content = any(str(v).strip() != "" for v in row_data.values())

        if not has_content:
            empty_count += 1
            if empty_count >= 3:
                break
            continue

        empty_count = 0
        data.append(row_data)

    return pd.DataFrame(data, columns=columns)


# =========================================================
# EXTRAER TODOS LOS BLOQUES DE LA HOJA
# =========================================================
def extract_blocks_from_sheet(ws):
    starts = find_line_action_starts(ws)
    delegacion = get_delegacion(ws)
    blocks = []

    if not starts:
        return blocks

    for i, start in enumerate(starts):
        start_row = start["row"]
        end_row = starts[i + 1]["row"] - 1 if i + 1 < len(starts) else ws.max_row

        linea_numero = start["line_value"]

        problematica = search_value_near_keywords(
            ws,
            start_row,
            min(start_row + 12, end_row),
            ["problemática", "problematica"]
        )

        lider = search_value_near_keywords(
            ws,
            start_row,
            min(start_row + 12, end_row),
            ["líder estratégico", "lider estrategico", "líder", "lider"]
        )

        trimestre = detect_trimester(
            ws,
            start_row,
            min(start_row + 12, end_row)
        )

        header_row = detect_header_row(
            ws,
            start_row,
            min(start_row + 25, end_row)
        )

        if header_row:
            tabla = extract_table(ws, header_row, end_row)
        else:
            tabla = pd.DataFrame(columns=[
                "Indicador",
                "Meta (editable)",
                "Avance",
                "Descripción",
                "Cantidad",
                "Observaciones (Editable)"
            ])

        blocks.append({
            "delegacion": delegacion,
            "linea_accion": linea_numero,
            "problematica": problematica,
            "lider": lider,
            "trimestre": trimestre,
            "tabla": tabla,
            "rango_inicio": start_row,
            "rango_fin": end_row
        })

    return blocks


# =========================================================
# EXPORTAR A EXCEL
# =========================================================
def build_export_file(block, df_editado):
    output = BytesIO()

    resumen = pd.DataFrame({
        "Campo": [
            "Delegación",
            "Línea de acción",
            "Problemática",
            "Líder Estratégico",
            "Trimestre",
            "Fila inicio",
            "Fila fin",
        ],
        "Valor": [
            block.get("delegacion", ""),
            block.get("linea_accion", ""),
            block.get("problematica", ""),
            block.get("lider", ""),
            block.get("trimestre", ""),
            block.get("rango_inicio", ""),
            block.get("rango_fin", ""),
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        df_editado.to_excel(writer, sheet_name="Detalle", index=False)

    output.seek(0)
    return output.getvalue()


# =========================================================
# ESTILOS
# =========================================================
st.markdown("""
<style>
.block-container {
    max-width: 1350px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}

.form-box {
    border: 1px solid #444;
    padding: 16px;
    margin-top: 10px;
    margin-bottom: 14px;
    border-radius: 10px;
}

.small-note {
    font-size: 0.92rem;
    opacity: 0.85;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# INTERFAZ
# =========================================================
st.title("Herramienta de lectura del instrumento")
st.write(
    "Sube un archivo Excel y la app leerá toda la hoja para detectar todas las líneas de acción."
)

uploaded_file = st.file_uploader(
    "Arrastra y suelta el archivo .xlsm o .xlsx",
    type=["xlsm", "xlsx"]
)

if uploaded_file is not None:
    try:
        wb = load_workbook(
            BytesIO(uploaded_file.read()),
            data_only=False,
            keep_vba=True
        )

        main_sheet = find_best_main_sheet(wb)
        ws = wb[main_sheet]

        blocks = extract_blocks_from_sheet(ws)

        if not blocks:
            st.warning("No se encontraron bloques de líneas de acción en la hoja.")
            st.stop()

        st.success(f"Hoja detectada: {main_sheet}")
        st.info(f"Bloques encontrados en toda la hoja: {len(blocks)}")

        options = []
        for i, b in enumerate(blocks, start=1):
            linea = b["linea_accion"] if b["linea_accion"] else f"Bloque {i}"
            prob = b["problematica"] if b["problematica"] else "Sin problemática detectada"
            options.append(f"{i}. Línea {linea} | {prob}")

        selected_idx = st.selectbox(
            "Seleccione la línea de acción detectada",
            options=list(range(len(options))),
            format_func=lambda x: options[x]
        )

        bloque = blocks[selected_idx]

        st.markdown('<div class="form-box">', unsafe_allow_html=True)

        c1, c2 = st.columns([1.2, 4])
        with c1:
            st.markdown("### Delegación :")
        with c2:
            st.text_input(
                "Delegación",
                value=bloque["delegacion"],
                label_visibility="collapsed",
                key="delegacion"
            )

        c1, c2, c3, c4, c5, c6 = st.columns([1.4, 1.5, 1.2, 1.8, 1.4, 1.8])

        with c1:
            st.markdown("### línea de acción #:")
        with c2:
            st.text_input(
                "Línea de acción",
                value=bloque["linea_accion"],
                label_visibility="collapsed",
                key="linea_accion"
            )

        with c3:
            st.markdown("### Problemática:")
        with c4:
            st.text_input(
                "Problemática",
                value=bloque["problematica"],
                label_visibility="collapsed",
                key="problematica"
            )

        with c5:
            st.markdown("### Líder Estratégico:")
        with c6:
            st.text_input(
                "Líder Estratégico",
                value=bloque["lider"],
                label_visibility="collapsed",
                key="lider"
            )

        c1, c2 = st.columns([1.4, 1.2])
        with c1:
            st.markdown("### Trimestre:")
        with c2:
            trim_options = ["", "I", "II", "III", "IV"]
            trim_value = bloque["trimestre"] if bloque["trimestre"] in trim_options else ""
            selected_trim = st.selectbox(
                "Trimestre",
                trim_options,
                index=trim_options.index(trim_value),
                label_visibility="collapsed",
                key="trimestre"
            )

        st.markdown("</div>", unsafe_allow_html=True)

        st.subheader("Detalle")

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

        df_editado = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="detalle_tabla"
        )

        c1, c2 = st.columns([1.5, 4])

        with c1:
            export_bytes = build_export_file(
                {
                    **bloque,
                    "trimestre": selected_trim
                },
                df_editado
            )
            st.download_button(
                "Descargar extracción en Excel",
                data=export_bytes,
                file_name="extraccion_linea_accion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with st.expander("Resumen técnico"):
            st.write(f"**Rango del bloque:** filas {bloque['rango_inicio']} a {bloque['rango_fin']}")
            st.write(f"**Delegación:** {bloque['delegacion']}")
            st.write(f"**Línea de acción detectada:** {bloque['linea_accion']}")
            st.write(f"**Problemática detectada:** {bloque['problematica']}")
            st.write(f"**Líder detectado:** {bloque['lider']}")
            st.write(f"**Trimestre detectado:** {selected_trim}")
            st.write(f"**Filas extraídas en detalle:** {len(df_editado)}")
            st.write("**Columnas detectadas:**")
            st.write(list(df_editado.columns))

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

else:
    st.info("Sube un archivo para comenzar.")
