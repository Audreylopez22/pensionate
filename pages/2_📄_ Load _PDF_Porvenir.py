import io
import re

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font

from utils.pension_utils import (
    build_gaps_report,
    build_weighted_monthly_report,
    clean_to_num,
)

st.set_page_config(page_title="Extract Tables from PDF - Porvenir", page_icon="📈", layout="wide")

st.markdown("# Extract Tables from PDF to Excel — Porvenir")
st.sidebar.header("Upload PDF")
st.write(
    """
    Upload a Porvenir *Historia Laboral* PDF to automatically extract its contribution tables.
    You will have two download options:
    1. An Excel file with each table in a separate sheet.
    2. A pension report (Consolidated, Final and Missing Months), sorted by period.
    """
)

# Columnas estándar del reporte consolidado
COL_EMPLOYER = "Empleador"
COL_MONTH = "Mes"
COL_DAYS = "Días cotizados"
COL_IBM = "IBM"
COL_AFP = "Administradora de pensiones"

# Palabras clave para reconocer cada columna en el encabezado del PDF (minúsculas)
HEADER_KEYWORDS = {
    COL_EMPLOYER: ["empleador"],
    COL_MONTH: ["mes", "periodo", "período"],
    COL_DAYS: ["días cotizados", "dias cotizados"],
    COL_IBM: ["ibc", "ingreso base"],
    COL_AFP: ["administradora"],
}

DATE_FORMAT = "DD/MM/YYYY"
MONTH_FORMAT = "YYYY/MM"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def normalize_text(value) -> str:
    """Elimina saltos de línea y espacios repetidos."""
    return re.sub(r"\s+", " ", str(value or "")).strip()


def extract_tables_from_pdf(pdf_bytes):
    """
    Extrae todas las tablas del PDF. Devuelve la lista de DataFrames y el texto
    de la primera página (donde Porvenir imprime el número de documento).
    """
    all_tables = []
    first_page_text = ""
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(pdf_document):
            first_page_text = pdf_document.load_page(0).get_text()
        for page in pdf_document:
            for table in page.find_tables():
                table_data = table.extract()
                if not table_data or len(table_data) < 2:
                    continue
                # En algunas páginas la primera fila llega vacía y el encabezado real es la 2ª.
                header_idx = 0
                if not any(normalize_text(h) for h in table_data[0]):
                    header_idx = 1
                headers = [normalize_text(h) for h in table_data[header_idx]]
                rows = table_data[header_idx + 1 :]
                if rows:
                    all_tables.append(pd.DataFrame(rows, columns=headers))
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
    return all_tables, first_page_text


def find_document_number(text: str) -> str:
    """Busca el número de documento (ej. 1.130.669.585) en el encabezado del PDF."""
    after_label = text.split("Número de documento", 1)[-1]
    match = re.search(r"^\s*(\d{1,3}(?:\.\d{3})+|\d{6,12})\s*$", after_label, re.MULTILINE)
    return re.sub(r"\D", "", match.group(1)) if match else ""


def map_contribution_columns(df: pd.DataFrame) -> dict | None:
    """
    Devuelve {columna_original: columna_estándar} si la tabla es de detalle de
    cotizaciones (Empleador, Mes, Días cotizados, IBC, Administradora); si no, None.
    """
    mapping = {}
    for std_col, keywords in HEADER_KEYWORDS.items():
        for col in df.columns:
            col_l = col.lower()
            if col not in mapping and any(k == col_l or k in col_l for k in keywords):
                # "mes" es muy corto: exigir coincidencia exacta para no capturar otras columnas
                if std_col == COL_MONTH and col_l not in keywords:
                    continue
                mapping[col] = std_col
                break
    required = {COL_EMPLOYER, COL_MONTH, COL_DAYS, COL_IBM}
    return mapping if required.issubset(mapping.values()) else None


def build_consolidated_report(tables: list[pd.DataFrame], document: str) -> pd.DataFrame:
    """Une las tablas de detalle, normaliza tipos y ordena cronológicamente por Mes."""
    frames = []
    for df_original in tables:
        mapping = map_contribution_columns(df_original)
        if not mapping:
            continue
        df = df_original[list(mapping)].rename(columns=mapping)
        frames.append(df)

    columns = ["DOCUMENTO", COL_EMPLOYER, COL_MONTH, COL_DAYS, COL_IBM, COL_AFP]
    if not frames:
        return pd.DataFrame(columns=columns)

    df = pd.concat(frames, ignore_index=True)
    if COL_AFP not in df.columns:
        df[COL_AFP] = None

    df[COL_EMPLOYER] = df[COL_EMPLOYER].apply(normalize_text)
    df[COL_AFP] = df[COL_AFP].apply(normalize_text)
    # Mes viene como "YYYY/MM"; se guarda como fecha (día 1) para ordenar y filtrar.
    df[COL_MONTH] = pd.to_datetime(
        df[COL_MONTH].apply(normalize_text), format="%Y/%m", errors="coerce"
    )
    df[COL_DAYS] = pd.to_numeric(df[COL_DAYS].apply(normalize_text), errors="coerce")
    df[COL_IBM] = df[COL_IBM].apply(clean_to_num)

    # Descarta filas vacías o basura (sin mes válido)
    df = df.dropna(subset=[COL_MONTH])
    df[COL_DAYS] = df[COL_DAYS].fillna(0).astype(int)
    df[COL_IBM] = df[COL_IBM].round().astype("int64")
    df.insert(0, "DOCUMENTO", int(document) if document.isdigit() else document)

    # Orden estable: dentro del mismo mes se conserva el orden original del PDF
    df = df.sort_values(COL_MONTH, kind="stable").reset_index(drop=True)
    return df[columns]


def to_excel_multi_sheet(dfs: list[pd.DataFrame]):
    """Exporta cada tabla extraída a una hoja distinta."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)
    return output.getvalue()


def autofit_columns(ws, min_width=10, max_width=60):
    """Ajusta el ancho de cada columna según su contenido."""
    for column_cells in ws.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in column_cells)
        letter = column_cells[0].column_letter
        ws.column_dimensions[letter].width = max(min_width, min(length + 2, max_width))


def to_excel_report(consolidated_df, final_df, gaps_df):
    """
    Genera el Excel con las hojas Consolidated_Report, Final_Report y Missing_Months,
    replicando las fórmulas auxiliares del informe de referencia de Porvenir.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        consolidated_df.to_excel(writer, sheet_name="Consolidated_Report", index=False)
        final_df.to_excel(writer, sheet_name="Final_Report", index=False)
        gaps_df.to_excel(writer, sheet_name="Missing_Months", index=False)

        # --- Consolidated_Report: Mes como YYYY/MM ---
        ws_c = writer.sheets["Consolidated_Report"]
        month_idx = consolidated_df.columns.get_loc(COL_MONTH) + 1
        for row in range(2, ws_c.max_row + 1):
            ws_c.cell(row=row, column=month_idx).number_format = MONTH_FORMAT

        # --- Final_Report: total de semanas (fórmula del informe de referencia) ---
        ws_f = writer.sheets["Final_Report"]
        n_rows = len(final_df)
        if n_rows:
            dias_idx = final_df.columns.get_loc("DIAS") + 1
            dias = ws_f.cell(row=1, column=dias_idx).column_letter
            weeks_col = len(final_df.columns) + 1
            last_row = n_rows + 1

            # Total de semanas cotizadas (Σ DIAS / 7)
            ws_f.cell(row=1, column=weeks_col, value="TOTAL SEMANAS")
            weeks_cell = ws_f.cell(
                row=2, column=weeks_col, value=f"=SUM({dias}2:{dias}{last_row})/7"
            )
            weeks_cell.number_format = "0.00"
            weeks_cell.font = Font(bold=True)

        # --- Missing_Months: semanas con 2 decimales ---
        ws_m = writer.sheets["Missing_Months"]
        if "SEMANAS" in gaps_df.columns:
            w_idx = gaps_df.columns.get_loc("SEMANAS") + 1
            for row in range(2, ws_m.max_row + 1):
                ws_m.cell(row=row, column=w_idx).number_format = "0.00"

        # --- Formato común: fechas DD/MM/YYYY, encabezados centrados, anchos ---
        for ws in writer.sheets.values():
            headers = [cell.value for cell in ws[1]]
            for i, h in enumerate(headers, start=1):
                ws.cell(row=1, column=i).alignment = Alignment(horizontal="center")
                if h and "FECHA" in str(h).upper():
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=i).number_format = DATE_FORMAT
            autofit_columns(ws)

    return output.getvalue()


# --- Streamlit Interface ---
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    base_name = uploaded_file.name.replace(".pdf", "")

    with st.spinner("Finding and extracting tables from the PDF..."):
        extracted_tables, first_page_text = extract_tables_from_pdf(pdf_bytes)

    if not extracted_tables:
        st.warning("No tables were found in the provided PDF file.")
        st.stop()

    st.success(f"Found {len(extracted_tables)} tables in the document!")

    # --- DEBUGGING: SHOW EXTRACTED HEADERS ---
    with st.expander("Show Extracted Headers for Debugging"):
        st.info("Below are the exact headers extracted from each table.")
        for i, df in enumerate(extracted_tables):
            st.write(f"**Table {i + 1} Headers:**")
            st.code(df.columns.tolist())
    st.markdown("---")

    # --- OPTION 1: DOWNLOAD ALL TABLES IN MULTIPLE SHEETS ---
    st.header("Option 1: Download All Tables in Separate Sheets")
    st.download_button(
        label="📥 Download all tables to Excel",
        data=to_excel_multi_sheet(extracted_tables),
        file_name=f"{base_name}_all_tables.xlsx",
        mime=XLSX_MIME,
    )
    st.markdown("---")

    # --- OPTION 2: PENSION REPORT ---
    st.header("Option 2: Pension Report")
    document = find_document_number(first_page_text)

    try:
        consolidated_df = build_consolidated_report(extracted_tables, document)
        if consolidated_df.empty:
            st.warning(
                "No contribution tables (Empleador / Mes / Días cotizados / IBC) were found."
            )
            st.stop()

        st.info(
            f"Extracted {len(consolidated_df)} contributions between "
            f"{consolidated_df[COL_MONTH].min():%Y/%m} and {consolidated_df[COL_MONTH].max():%Y/%m}."
        )

        final_report_df = build_weighted_monthly_report(
            consolidated_df,
            month_col=COL_MONTH,
            employer_col=COL_EMPLOYER,
            days_col=COL_DAYS,
            ibc_col=COL_IBM,
            document=consolidated_df["DOCUMENTO"].iloc[0],
        )
        gaps_report_df = build_gaps_report(final_report_df)

        total_weeks = final_report_df["DIAS"].sum() / 7
        gap_weeks = gaps_report_df["SEMANAS"].sum() if not gaps_report_df.empty else 0

        st.write("#### Consolidated Report Preview")
        st.dataframe(
            consolidated_df.head(15),
            column_config={COL_MONTH: st.column_config.DateColumn(format="YYYY/MM")},
        )

        if not final_report_df.empty:
            st.write("#### Final Weighted Report Preview (IBC)")
            st.dataframe(final_report_df.head(15))



        st.download_button(
            label="📥 Download Report (Multiple Sheets)",
            data=to_excel_report(consolidated_df, final_report_df, gaps_report_df),
            file_name=f"{base_name}_pension_report.xlsx",
            mime=XLSX_MIME,
        )

    except Exception as e:
        st.error(f"Error generating reports: {e}")
        st.exception(e)
