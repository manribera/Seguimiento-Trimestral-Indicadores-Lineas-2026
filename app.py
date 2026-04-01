# app.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak
)

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
    return str(v).strip() != ""


def safe_str(v):
    if v is None:
        return ""
    return str(v).strip()


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
# EXTRAER VALOR CORRECTO DE LÍNEA DE ACCIÓN
# =========================================================
def looks_like_bad_line_value(text: str) -> bool:
    t = norm_text(text)

    bad_patterns = [
        "problemática",
        "problematica",
        "líder",
        "lider",
        "delegación",
        "delegacion",
        "trimestre",
        "indicador",
        "meta",
        "avance",
        "descripción",
        "descripcion",
        "cantidad",
        "observaciones",
        "linea de accion",
        "línea de acción",
    ]
    return any(p in t for p in bad_patterns)


def extract_line_number_from_area(ws, start_row, start_col):
    """
    Busca el número o valor real de la línea de acción.
    Prioriza valores cortos o numéricos cercanos al rótulo.
    """
    candidates = []

    # misma fila
    for c in range(start_col + 1, min(ws.max_column, start_col + 8) + 1):
        val = ws.cell(start_row, c).value
        if is_meaningful_value(val):
            txt = clean_text(val)
            if not looks_like_bad_line_value(txt):
                candidates.append(txt)

    # filas cercanas
    for r in range(start_row, min(ws.max_row, start_row + 2) + 1):
        for c in range(start_col, min(ws.max_column, start_col + 8) + 1):
            val = ws.cell(r, c).value
            if is_meaningful_value(val):
                txt = clean_text(val)
                if not looks_like_bad_line_value(txt):
                    candidates.append(txt)

    if not candidates:
        return ""

    # preferir valor corto o numérico
    numeric_candidates = []
    for x in candidates:
        m = re.search(r"\d+([.\-]\d+)?", x)
        if m:
            numeric_candidates.append(m.group(0))

    if numeric_candidates:
        return numeric_candidates[0]

    short_candidates = [x for x in candidates if len(x) <= 20]
    if short_candidates:
        return short_candidates[0]

    return candidates[0]


# =========================================================
# DETECTAR BLOQUES
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

    cleaned = []
    last_row = -999

    for item in starts:
        if item["row"] - last_row > 2:
            cleaned.append(item)
            last_row = item["row"]

    return cleaned


# =========================================================
# EXTRAER CAMPOS CERCANOS
# =========================================================
def search_value_near_keywords(ws, start_row, end_row, keywords):
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if contains_any(val, keywords):
                derecha = get_right_value(ws, r, c, max_steps=10)
                if is_meaningful_value(derecha):
                    return clean_text(derecha)

                abajo = get_down_value(ws, r, c, max_steps=3)
                if is_meaningful_value(abajo):
                    return clean_text(abajo)
    return ""


def detect_trimester(ws, start_row, end_row):
    for r in range(start_row, min(end_row, ws.max_row) + 1):
        txt = norm_text(row_text(ws, r))

        if "trimestre" in txt:
            for c in range(1, ws.max_column + 1):
                v = clean_text(ws.cell(r, c).value)
                if v in ["I", "II", "III", "IV"]:
                    return v

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
# ENCABEZADOS DE TABLA
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
            header_map["Meta"] = c
        elif "avance" in t:
            header_map["Avance"] = c
        elif "descripcion" in t or "descripción" in t:
            header_map["Descripción"] = c
        elif "cantidad" in t:
            header_map["Cantidad"] = c
        elif "observ" in t:
            header_map["Observaciones"] = c

    return header_map


# =========================================================
# TABLA DETALLE
# =========================================================
def extract_table(ws, header_row, block_end_row):
    columns = [
        "Indicador",
        "Meta",
        "Avance",
        "Descripción",
        "Cantidad",
        "Observaciones"
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
            row_data[col_name] = ws.cell(r, col_idx).value if col_idx else ""

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
# EXTRAER TODOS LOS BLOQUES
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

        tabla = extract_table(ws, header_row, end_row) if header_row else pd.DataFrame(columns=[
            "Indicador", "Meta", "Avance", "Descripción", "Cantidad", "Observaciones"
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
# RESUMEN PARA TABLERO
# =========================================================
def build_summary_df(blocks):
    rows = []

    for i, b in enumerate(blocks, start=1):
        rows.append({
            "Línea": b["linea_accion"] if b["linea_accion"] else str(i),
            "Problemática": b["problematica"],
            "Líder Estratégico": b["lider"],
            "Trimestre": b["trimestre"],
            "Indicadores": len(b["tabla"]) if not b["tabla"].empty else 0,
        })

    return pd.DataFrame(rows)


# =========================================================
# PDF
# =========================================================
def paragraph(text, style):
    return Paragraph(str(text).replace("\n", "<br/>"), style)


def build_pdf_report(blocks, delegacion, trimestre_general):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=35,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "title_custom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=16,
        leading=20,
        spaceAfter=10
    )

    style_subtitle = ParagraphStyle(
        "subtitle_custom",
        parent=styles["Heading2"],
        alignment=TA_LEFT,
        fontSize=11,
        leading=14,
        spaceAfter=8
    )

    style_normal = ParagraphStyle(
        "normal_custom",
        parent=styles["BodyText"],
        fontSize=9,
        leading=11
    )

    style_small = ParagraphStyle(
        "small_custom",
        parent=styles["BodyText"],
        fontSize=8,
        leading=10
    )

    elements = []

    elements.append(Paragraph("REPORTE TRIMESTRAL DE LÍNEAS DE ACCIÓN", style_title))
    elements.append(Paragraph(f"<b>Delegación:</b> {delegacion or 'No detectada'}", style_subtitle))
    elements.append(Paragraph(f"<b>Trimestre:</b> {trimestre_general or 'No detectado'}", style_subtitle))
    elements.append(Spacer(1, 10))

    # Resumen general
    elements.append(Paragraph("Resumen general", style_subtitle))

    summary_data = [["Línea", "Problemática", "Líder Estratégico", "Trimestre", "Indicadores"]]
    for i, b in enumerate(blocks, start=1):
        summary_data.append([
            safe_str(b["linea_accion"] if b["linea_accion"] else i),
            safe_str(b["problematica"]),
            safe_str(b["lider"]),
            safe_str(b["trimestre"]),
            str(len(b["tabla"]) if not b["tabla"].empty else 0)
        ])

    summary_table = Table(summary_data, repeatRows=1, colWidths=[50, 160, 140, 60, 60])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E2F3")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F7F7")]),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 14))

    # Detalle por línea
    for i, b in enumerate(blocks, start=1):
        linea_txt = safe_str(b["linea_accion"] if b["linea_accion"] else i)

        elements.append(Paragraph(f"Línea de acción #{linea_txt}", style_subtitle))
        elements.append(Paragraph(f"<b>Problemática:</b> {safe_str(b['problematica'])}", style_normal))
        elements.append(Paragraph(f"<b>Líder Estratégico:</b> {safe_str(b['lider'])}", style_normal))
        elements.append(Paragraph(f"<b>Trimestre:</b> {safe_str(b['trimestre'])}", style_normal))
        elements.append(Spacer(1, 6))

        detail_data = [[
            "Indicador", "Meta", "Avance", "Descripción", "Cantidad", "Observaciones"
        ]]

        if not b["tabla"].empty:
            for _, row in b["tabla"].iterrows():
                detail_data.append([
                    paragraph(row.get("Indicador", ""), style_small),
                    paragraph(row.get("Meta", ""), style_small),
                    paragraph(row.get("Avance", ""), style_small),
                    paragraph(row.get("Descripción", ""), style_small),
                    paragraph(row.get("Cantidad", ""), style_small),
                    paragraph(row.get("Observaciones", ""), style_small),
                ])
        else:
            detail_data.append(["Sin datos", "", "", "", "", ""])

        detail_table = Table(
            detail_data,
            repeatRows=1,
            colWidths=[120, 90, 55, 95, 55, 110]
        )

        detail_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#B4C6E7")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
        ]))

        elements.append(detail_table)
        elements.append(Spacer(1, 14))

    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


# =========================================================
# INTERFAZ
# =========================================================
st.markdown("""
<style>
.block-container {
    max-width: 1400px;
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
</style>
""", unsafe_allow_html=True)

st.title("Herramienta de lectura del instrumento")
st.write("Sube un archivo Excel y la app leerá toda la hoja para generar el reporte trimestral en PDF.")

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

        delegacion = blocks[0]["delegacion"] if blocks else ""
        trimestre_general = ""
        for b in blocks:
            if safe_str(b["trimestre"]):
                trimestre_general = b["trimestre"]
                break

        st.success(f"Hoja detectada: {main_sheet}")
        st.info(f"Líneas de acción encontradas en toda la hoja: {len(blocks)}")

        st.markdown('<div class="form-box">', unsafe_allow_html=True)

        c1, c2 = st.columns([1.2, 4])
        with c1:
            st.markdown("### Delegación :")
        with c2:
            st.text_input(
                "Delegación",
                value=delegacion,
                label_visibility="collapsed",
                disabled=True
            )

        c1, c2 = st.columns([1.2, 4])
        with c1:
            st.markdown("### Trimestre:")
        with c2:
            trim_options = ["", "I", "II", "III", "IV"]
            trim_value = trimestre_general if trimestre_general in trim_options else ""
            trimestre_general = st.selectbox(
                "Trimestre",
                trim_options,
                index=trim_options.index(trim_value),
                label_visibility="collapsed"
            )

        st.markdown("</div>", unsafe_allow_html=True)

        st.subheader("Resumen de líneas detectadas")

        summary_df = build_summary_df(blocks)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.caption("El PDF incluirá el detalle completo de todas las líneas de acción detectadas en la hoja.")

        generar_pdf = st.button("Generar reporte trimestral en PDF")

        if generar_pdf:
            pdf_bytes = build_pdf_report(blocks, delegacion, trimestre_general)
            st.success("Reporte PDF generado correctamente.")

            st.download_button(
                label="Descargar reporte trimestral PDF",
                data=pdf_bytes,
                file_name=f"reporte_trimestral_{delegacion or 'delegacion'}.pdf",
                mime="application/pdf"
            )

        with st.expander("Resumen técnico"):
            st.write(f"**Hoja detectada:** {main_sheet}")
            st.write(f"**Delegación:** {delegacion}")
            st.write(f"**Trimestre general:** {trimestre_general}")
            st.write(f"**Cantidad de líneas detectadas:** {len(blocks)}")

            debug_rows = []
            for i, b in enumerate(blocks, start=1):
                debug_rows.append({
                    "Bloque": i,
                    "Línea detectada": b["linea_accion"],
                    "Problemática": b["problematica"],
                    "Líder": b["lider"],
                    "Trimestre": b["trimestre"],
                    "Fila inicio": b["rango_inicio"],
                    "Fila fin": b["rango_fin"],
                    "Filas detalle": len(b["tabla"]) if not b["tabla"].empty else 0
                })
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

else:
    st.info("Sube un archivo para comenzar.")
