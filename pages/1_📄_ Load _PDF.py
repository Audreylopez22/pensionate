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
            # For each table, check if it contains the core concepts (ID, Name, Period)
            headers_as_string = ' '.join(df_original.columns)
            
            # --- FILTER: EXCLUDE unwanted formats (like the [34], [40] series) ---
            # If the table headers contain markers of the unwanted format, skip it.
            unwanted_markers = ['[34]', '[40]IBC', '[45] Días Cot.']
            if any(marker in headers_as_string for marker in unwanted_markers):
                continue

            # --- DETECTION: Include standard and numbered formats [1], [2], etc. ---
            has_id_col = any(k in headers_as_string for k in ['Identificación', '[1]Identificación'])
            has_name_col = any(k in headers_as_string for k in ['Razón Social', '[2]Nombre'])
            has_period_col = any(k in headers_as_string for k in ['Período', 'Ciclo', 'Periodo', '[3]Desde'])
            
            # If the table contains all three concepts, process it
            if has_id_col and has_name_col and has_period_col:
                df = df_original.copy()
                
                # Build a rename map to standardize the key columns
                rename_map = {}
                # Search for the period column (could be standard or [3]Desde)
                period_col = next((c for c in df.columns if any(k in c for k in ['Período', 'Ciclo', 'Periodo', 'Desde'])), None)
                id_col = next((c for c in df.columns if 'Identificación' in c), None)
                name_col = next((c for c in df.columns if 'Razón Social' in c or 'Nombre' in c), None)
                
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

                # --- Robust Sorting and Date Parsing ---
                # Attempt to parse Periodo. 1st: YYYYMM format, 2nd: General date format (for 'Desde' columns)
                sorted_df['Periodo_dt'] = pd.to_datetime(sorted_df['Periodo'].astype(str), format='%Y%m', errors='coerce')
                
                # If NaT remains, it's likely a full date (DD/MM/YYYY)
                mask_nat_start = sorted_df['Periodo_dt'].isna()
                if mask_nat_start.any():
                    sorted_df.loc[mask_nat_start, 'Periodo_dt'] = pd.to_datetime(
                        sorted_df.loc[mask_nat_start, 'Periodo'].astype(str), 
                        dayfirst=True, 
                        errors='coerce'
                    )

                # Identify and parse the "Hasta" column if it exists for expansion
                hasta_col = next((c for c in sorted_df.columns if 'hasta' in c.lower()), None)
                if hasta_col:
                    sorted_df['Hasta_dt'] = pd.to_datetime(sorted_df[hasta_col].astype(str), dayfirst=True, errors='coerce')
                else:
                    sorted_df['Hasta_dt'] = pd.NA

                sorted_df.sort_values(by='Periodo_dt', ascending=True, na_position='last', inplace=True)
                
                # --- Create the second sheet (Summary Report) with Date Expansion ---
                st.markdown("---")
                st.header("Generated Summary Report")

                expanded_rows = []
                
                # Identify columns for the summary beforehand
                id_col_name = next((c for c in sorted_df.columns if 'Identificación' in c), None)
                potential_ibc_cols = [c for c in sorted_df.columns if 'ibc' in c.lower()]
                potential_asig_cols = [
                    c for c in sorted_df.columns 
                    if ('asign' in c.lower() and ('básica' in c.lower() or 'basica' in c.lower())) or
                       ('último salario' in c.lower() or 'ultimo salario' in c.lower() or '[5]' in c.lower())
                ]

                def is_really_empty(val):
                    """Checks if a value is effectively empty (NaN, empty string, or zero-like)."""
                    if pd.isna(val): return True
                    s = str(val).strip().lower()
                    return s in ['', '-', '$ 0', '0', '0,00', '0.00', '$ 0,00', '$ 0.00']

                # Map standard columns mapping
                col_map = {'DIAS': 'DIAS', 'TOTAL DIAS': 'TOTAL DIAS', 'SEMANAS': 'SEMANAS'}
                source_cols_map = {}
                for target_col, keyword in col_map.items():
                    sc = next((c for c in sorted_df.columns if keyword.lower() in c.lower()), None)
                    if not sc:
                        if target_col == 'TOTAL DIAS':
                            sc = next((c for c in sorted_df.columns if '[9]' in c or 'total' in c.lower()), None)
                        elif target_col == 'SEMANAS':
                            sc = next((c for c in sorted_df.columns if '[6]' in c), None)
                    source_cols_map[target_col] = sc

                # Iterate through each row to expand if the range spans multiple months
                for _, row in sorted_df.iterrows():
                    start_date = row['Periodo_dt']
                    end_date = row['Hasta_dt']
                    
                    if pd.isna(start_date):
                        continue

                    # Determine the segments (one per month) for the expansion
                    month_segments = []
                    if pd.isna(end_date):
                        # No end date provided: assume a single full month from start_date
                        s = start_date.replace(day=1)
                        e = s + pd.offsets.MonthEnd(0)
                        month_segments.append((s, e))
                    else:
                        # Split the range [start_date, end_date] into monthly segments
                        curr_s = start_date
                        while curr_s <= end_date:
                            # End of the current month, but not beyond the overall end_date
                            curr_e = min(curr_s + pd.offsets.MonthEnd(0), end_date)
                            month_segments.append((curr_s, curr_e))
                            # Start of the next month
                            curr_s = (curr_e + pd.Timedelta(days=1)).replace(day=1)

                    # Find IBC for this row using previous robust logic
                    final_ibc = None
                    for col in potential_ibc_cols:
                        if not is_really_empty(row[col]):
                            final_ibc = row[col]
                            break
                    if is_really_empty(final_ibc):
                        for col in potential_asig_cols:
                            if not is_really_empty(row[col]):
                                final_ibc = row[col]
                                break
                    if is_really_empty(final_ibc) and potential_ibc_cols:
                        final_ibc = row[potential_ibc_cols[0]]

                    # Create a summary row for each monthly segment
                    for seg_start, seg_end in month_segments:
                        # Calculate the actual number of days in this segment
                        calculated_days = (seg_end - seg_start).days + 1
                        
                        new_row = {
                            'DOCUMENTO': row[id_col_name] if id_col_name else None,
                            'FECHA INICIAL': seg_start.strftime('%d/%m/%Y'),
                            'FECHA FINAL': seg_end.strftime('%d/%m/%Y'),
                            'IBC': final_ibc,
                            'DIAS': calculated_days,
                            'TOTAL DIAS': row[source_cols_map['TOTAL DIAS']] if source_cols_map['TOTAL DIAS'] else None,
                            'SEMANAS': row[source_cols_map['SEMANAS']] if source_cols_map['SEMANAS'] else None
                        }
                        expanded_rows.append(new_row)

                summary_df = pd.DataFrame(expanded_rows)

                # --- NEW: Create "Final_Report" with Weighted Averages (IBM) ---
                def get_ibm_report(df_summary):
                    if df_summary.empty:
                        return pd.DataFrame()
                    
                    # 1. Deduplicate: Remove exactly identical rows that might come from redundant tables
                    df = df_summary.drop_duplicates(subset=['DOCUMENTO', 'FECHA INICIAL', 'FECHA FINAL', 'IBC']).copy()
                    
                    # 2. Clean IBC to numeric for math
                    def clean_to_num(val):
                        if pd.isna(val): return 0.0
                        s = str(val).replace('$', '').replace('.', '').replace(',', '.').replace(' ', '').strip()
                        try: return float(s)
                        except: return 0.0
                    
                    df['IBC_num'] = df['IBC'].apply(clean_to_num)
                    df['FECHA_DT'] = pd.to_datetime(df['FECHA INICIAL'], format='%d/%m/%Y')
                    df['MONTH_KEY'] = df['FECHA_DT'].dt.to_period('M')
                    
                    # 3. Weighted calculation: (Days * IBC)
                    df['WEIGHTED_VAL'] = df['DIAS'] * df['IBC_num']
                    
                    # 4. Group by Month and Year
                    final_rows = []
                    for period, group in df.groupby('MONTH_KEY'):
                        total_weighted_sum = group['WEIGHTED_VAL'].sum()
                        days_in_month = period.days_in_month
                        
                        # IBC is the weighted average limited by the theoretical month days
                        ibc = total_weighted_sum / days_in_month if days_in_month > 0 else 0
                        
                        # Document from the first valid row
                        doc = group['DOCUMENTO'].iloc[0] if 'DOCUMENTO' in group.columns else None
                        
                        # IMPORTANT: Cap DIAS TOTALES to the actual number of days in that month
                        actual_sum_days = group['DIAS'].sum()
                        capped_days = min(actual_sum_days, days_in_month)
                        
                        final_rows.append({
                            'DOCUMENTO': doc,
                            'FECHA INICIAL': period.start_time.strftime('%d/%m/%Y'),
                            'FECHA FINAL': period.end_time.strftime('%d/%m/%Y'),
                            'IBC(Ponderado)': f"$ {ibc:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                            'DIAS TOTALES': capped_days
                        })
                    
                    return pd.DataFrame(final_rows)

                final_report_df = get_ibm_report(summary_df)

                st.write("#### Summary Report Preview")
                st.dataframe(summary_df.head())

                if not final_report_df.empty:
                    st.write("#### Final Weighted Report Preview (IBM)")
                    st.dataframe(final_report_df.head())

                # --- Prepare Excel file with THREE sheets ---
                # Drop temporary datetime columns from the main report
                cols_to_drop = ['Periodo_dt']
                if 'Hasta_dt' in sorted_df.columns: cols_to_drop.append('Hasta_dt')
                sorted_df.drop(columns=cols_to_drop, inplace=True)
                
                # Use a new function to write multiple sheets
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    sorted_df.to_excel(writer, sheet_name='Consolidated_Report', index=False)
                    summary_df.to_excel(writer, sheet_name='Summary_Report', index=False)
                    if not final_report_df.empty:
                        final_report_df.to_excel(writer, sheet_name='Final_Report', index=False)
                
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
