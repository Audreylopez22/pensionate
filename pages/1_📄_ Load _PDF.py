import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re

st.set_page_config(page_title="Extract Tables from PDF", page_icon="📈", layout="wide")

st.markdown("# Extract Tables from PDF to Excel")
st.sidebar.header("Upload PDF")
st.write(
    """
    Upload a PDF file to automatically extract all its tables. You will have two download options:
    1. An Excel file with each table in a separate sheet.
    2. A consolidated Excel report from specific tables, combined and sorted by period.
    """
)

def extract_tables_from_pdf(pdf_bytes):
    """
    Extracts all tables from a PDF, cleaning up header names for consistency.
    """
    all_tables = []
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            tables = page.find_tables()
            if tables:
                for table in tables:
                    table_data = table.extract()
                    if table_data and len(table_data) > 1:
                        # Normalize headers: remove newlines, multiple spaces, and trim whitespace
                        headers = [re.sub(r'\s+', ' ', str(h)).strip() for h in table_data[0]]
                        df = pd.DataFrame(table_data[1:], columns=headers)
                        all_tables.append(df)
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
    return all_tables

def to_excel_multi_sheet(dfs: list[pd.DataFrame]):
    """
    Converts a list of DataFrames to an in-memory Excel file, each in a separate sheet.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
    return output.getvalue()

def to_excel_single_sheet(df):
    """
    Converts a single DataFrame to an in-memory Excel file.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Consolidated_Report', index=False)
    return output.getvalue()

# --- Streamlit Interface ---
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    
    with st.spinner("Finding and extracting tables from the PDF..."):
        extracted_tables = extract_tables_from_pdf(pdf_bytes)

    if extracted_tables:
        st.success(f"Found {len(extracted_tables)} tables in the document!")
        
        # --- DEBUGGING: SHOW EXTRACTED HEADERS ---
        with st.expander("Show Extracted Headers for Debugging"):
            st.info("Below are the exact headers extracted from each table. Use these to verify the keywords for matching.")
            for i, df in enumerate(extracted_tables):
                st.write(f"**Table {i+1} Headers:**")
                st.code(df.columns.tolist())
        st.markdown("---")
        
        # --- OPTION 1: DOWNLOAD ALL TABLES IN MULTIPLE SHEETS ---
        st.header("Option 1: Download All Tables in Separate Sheets")
        excel_data_multi = to_excel_multi_sheet(extracted_tables)
        st.download_button(
            label="📥 Download all tables to Excel",
            data=excel_data_multi,
            file_name=f"{uploaded_file.name.replace('.pdf', '')}_all_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("---")

        # --- OPTION 2: CREATE AND DOWNLOAD CONSOLIDATED REPORT BY MATCHING HEADERS ---
        st.header("Option 2: Create Consolidated Report by Header Content")

        final_tables_to_concat = []
        for df_original in extracted_tables:
            # For each table, check if it contains the three core concepts
            headers_as_string = ' '.join(df_original.columns)
            
            has_id_col = 'Identificación' in headers_as_string
            has_name_col = 'Razón Social' in headers_as_string
            has_period_col = 'Período' in headers_as_string or 'Ciclo' in headers_as_string
            
            # If the table contains all three concepts, process it
            if has_id_col and has_name_col and has_period_col:
                df = df_original.copy()
                
                # Build a rename map to standardize the key columns
                rename_map = {}
                period_col = next((c for c in df.columns if 'Período' in c or 'Ciclo' in c or 'Periodo' in c), None)
                id_col = next((c for c in df.columns if 'Identificación' in c), None)
                name_col = next((c for c in df.columns if 'Razón Social' in c), None)
                
                if period_col: rename_map[period_col] = 'Periodo'
                if id_col: rename_map[id_col] = 'Identificación'
                if name_col: rename_map[name_col] = 'Nombre o Razón Social'
                
                df.rename(columns=rename_map, inplace=True)
                
                # Safely handle potential duplicate columns that arise from renaming
                cols = pd.Series(df.columns)
                for dup in cols[cols.duplicated()].unique():
                    cols[cols[cols == dup].index.values.tolist()] = [f'{dup}.{i}' if i != 0 else dup for i in range(sum(cols == dup))]
                df.columns = cols

                final_tables_to_concat.append(df)

        if final_tables_to_concat:
            st.info(f"Found and standardized {len(final_tables_to_concat)} table(s) for the consolidated report.")
            
            consolidated_df = pd.concat(final_tables_to_concat, ignore_index=True, join='outer')
            
            if 'Periodo' in consolidated_df.columns:
                st.success("Sorting report by the unified 'Periodo' column.")
                sorted_df = consolidated_df.copy()

                # --- Main Sorting Logic ---
                # Create a temporary datetime column for accurate sorting
                sorted_df['Periodo_dt'] = pd.to_datetime(sorted_df['Periodo'].astype(str), format='%Y%m', errors='coerce')
                sorted_df.sort_values(by='Periodo_dt', ascending=True, na_position='last', inplace=True)
                
                # --- Create the second sheet (Summary Report) ---
                st.markdown("---")
                st.header("Generated Summary Report")

                summary_df = pd.DataFrame()
                
                # 1. Map 'DOCUMENTO' from 'Identificación'
                id_col_name = next((c for c in sorted_df.columns if 'Identificación' in c), None)
                if id_col_name:
                    summary_df['DOCUMENTO'] = sorted_df[id_col_name]
                
                # 2. Calculate FECHA INICIAL and FECHA FINAL
                summary_df['FECHA INICIAL'] = sorted_df['Periodo_dt'].dt.strftime('%d/%m/%Y')
                summary_df['FECHA FINAL'] = (sorted_df['Periodo_dt'] + pd.offsets.MonthEnd(0)).dt.strftime('%d/%m/%Y')

                # 3. Robustly find and copy other requested columns
                
                # --- Advanced logic for 'IBC' column (Multi-column search) ---
                # We look for ALL columns that might contain the data, not just the first one
                potential_ibc_cols = [c for c in sorted_df.columns if 'ibc' in c.lower()]
                potential_asig_cols = [c for c in sorted_df.columns if 'asign' in c.lower() and ('básica' in c.lower() or 'basica' in c.lower())]

                def is_really_empty(val):
                    """Checks if a value is effectively empty (NaN, empty string, or zero-like)."""
                    if pd.isna(val): return True
                    s = str(val).strip().lower()
                    return s in ['', '-', '$ 0', '0', '0,00', '0.00', '$ 0,00', '$ 0.00']

                ibc_values = []
                for _, row in sorted_df.iterrows():
                    final_val = None
                    
                    # 1. Try to find a non-empty value in ANY column that looks like 'IBC'
                    for col in potential_ibc_cols:
                        if not is_really_empty(row[col]):
                            final_val = row[col]
                            break
                    
                    # 2. If still empty, try ANY column that looks like 'Asignación Básica'
                    if is_really_empty(final_val):
                        for col in potential_asig_cols:
                            if not is_really_empty(row[col]):
                                final_val = row[col]
                                break
                    
                    # 3. If still empty, just take the first IBC column value found (could be $ 0)
                    if is_really_empty(final_val) and potential_ibc_cols:
                        final_val = row[potential_ibc_cols[0]]
                        
                    ibc_values.append(final_val)

                summary_df['IBC'] = ibc_values

                # --- Map other standard columns ---
                col_map = {
                    'DIAS': 'DIAS', 
                    'TOTAL DIAS': 'TOTAL DIAS', 
                    'SEMANAS': 'SEMANAS'
                }
                for target_col, keyword in col_map.items():
                    # Find the first column in sorted_df that contains the keyword (case-insensitive)
                    source_col = next((c for c in sorted_df.columns if keyword.lower() in c.lower()), None)
                    if source_col:
                        summary_df[target_col] = sorted_df[source_col]
                    else:
                        summary_df[target_col] = None # Or some default value if the column is not found

                st.write("#### Summary Report Preview")
                st.dataframe(summary_df.head())

                # --- Prepare Excel file with TWO sheets ---
                # First, drop the temporary datetime column from the main report
                sorted_df.drop(columns=['Periodo_dt'], inplace=True)
                
                # Use a new function to write multiple sheets
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    sorted_df.to_excel(writer, sheet_name='Consolidated_Report', index=False)
                    summary_df.to_excel(writer, sheet_name='Summary_Report', index=False)
                
                excel_data_with_summary = output.getvalue()
                
                st.download_button(
                    label="📥 Download Report (2 Sheets)",
                    data=excel_data_with_summary,
                    file_name=f"{uploaded_file.name.replace('.pdf', '')}_report_with_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Could not find a unified 'Periodo' column to sort the report.")
                st.dataframe(consolidated_df.head())
        else:
            st.warning("No tables matching the required header concepts were found.")
            
    else:
        st.warning("No tables were found in the provided PDF file.")
