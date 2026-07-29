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

def is_really_empty(val):
    """Checks if a value is effectively empty (NaN, empty string, or zero-like)."""
    if pd.isna(val):
        return True
    s = str(val).strip().lower()
    return s in ['', '-', '$ 0', '0', '0,00', '0.00', '$ 0,00', '$ 0.00']

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

        # Pre-process: Create a lookup for Days and IBC [40] from ANY table that contains them
        days_lookup_45 = {} 
        ibc_lookup_40 = {} 
        for df_orig in extracted_tables:
            # Identify columns for ID, Period, Days
            id_c = next((c for c in df_orig.columns if any(k in c.lower() for k in ['[34]', '[47]', 'identificación', 'identificacion'])), None)
            per_c = next((c for c in df_orig.columns if any(k in c.lower() for k in ['[37]', '[50]', 'período', 'periodo', 'ciclo'])), None)
            
            if id_c and per_c:
                cot_cols = [c for c in df_orig.columns if any(k in c.lower() for k in ['[45]', '[58]', 'días cot', 'dias cot'])]
                rep_cols = [c for c in df_orig.columns if any(k in c.lower() for k in ['[57]', 'días rep', 'dias rep'])]
                ibc_40_cols = [c for c in df_orig.columns if '[40]ibc' in c.lower()]

                for _, r in df_orig.iterrows():
                    val_id = str(r[id_c]).strip()
                    val_per = str(r[per_c]).strip()
                    clean_id = re.sub(r'\D', '', val_id)
                    
                    # Normalize period to YYYYMM for the lookup
                    norm_per = None
                    if len(val_per) == 6 and val_per.isdigit():
                        norm_per = val_per
                    else:
                        dt_tmp = pd.to_datetime(val_per, dayfirst=True, errors='coerce')
                        if not pd.isna(dt_tmp):
                            norm_per = dt_tmp.strftime('%Y%m')

                    if not clean_id or not norm_per:
                        continue

                    # PRIORITY LOGIC FOR DAYS PER ROW:
                    final_day_val = 0
                    for c in cot_cols + rep_cols:
                        try:
                            v = float(str(r[c]).replace(',', '.'))
                            if v > 0:
                                final_day_val = v
                                break
                        except Exception:
                            continue
                    
                    # Store the best value found for this ID and Period
                    if final_day_val > 0:
                        days_lookup_45[(clean_id, norm_per)] = final_day_val

                    # IBC [40] LOGIC PER ROW:
                    for c in ibc_40_cols:
                        val_ibc40 = r[c]
                        if not is_really_empty(val_ibc40):
                            ibc_lookup_40[(clean_id, norm_per)] = val_ibc40
                            break

        final_tables_to_concat = []
        for df_original in extracted_tables:
            # For each table, check if it contains the core concepts (ID, Name, Period)
            headers_as_string = ' '.join(df_original.columns)
            
            # --- FILTER: EXCLUDE unwanted formats (like the [34], [40] series) ---
            # Now we exclude them from the CONCATENATION, but we already extracted their [45] data above.
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
                # Search for the "Hasta" column (standard or [30]Ciclo Hasta)
                hasta_col_orig = next((c for c in df.columns if any(k in c.lower() for k in ['hasta', '[30]'])), None)
                
                if period_col: rename_map[period_col] = 'Periodo'
                if id_col: rename_map[id_col] = 'Identificación'
                if name_col: rename_map[name_col] = 'Nombre o Razón Social'
                if hasta_col_orig: rename_map[hasta_col_orig] = 'Hasta'
                
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
            
            # --- NEW: Add [40]IBC Reportado column to the Consolidated Report ---
            def get_ibc40_for_row(row):
                raw_id = str(row.get('Identificación', '')).strip()
                clean_id = re.sub(r'\D', '', raw_id)
                val_per = str(row.get('Periodo', '')).strip()
                
                # Normalize period to YYYYMM for matching
                norm_per = val_per
                if not (len(val_per) == 6 and val_per.isdigit()):
                    dt_tmp = pd.to_datetime(val_per, dayfirst=True, errors='coerce')
                    if not pd.isna(dt_tmp):
                        norm_per = dt_tmp.strftime('%Y%m')
                
                return ibc_lookup_40.get((clean_id, norm_per), None)

            consolidated_df['[40]IBC Reportado'] = consolidated_df.apply(get_ibc40_for_row, axis=1)

            # Move [40]IBC Reportado next to 'ultimo salario' if possible
            target_col = next((c for c in consolidated_df.columns if 'ultimo salario' in c.lower() or 'último salario' in c.lower() or any(k in c.lower() for k in ['[5]', '[31]'])), None)
            if target_col:
                cols = list(consolidated_df.columns)
                cols.remove('[40]IBC Reportado')
                idx = cols.index(target_col)
                cols.insert(idx + 1, '[40]IBC Reportado')
                consolidated_df = consolidated_df[cols]

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
                hasta_col = next((c for c in sorted_df.columns if any(k in c.lower() for k in ['hasta'])), None)
                if hasta_col:
                    sorted_df['Hasta_dt'] = pd.to_datetime(sorted_df[hasta_col].astype(str), dayfirst=True, errors='coerce')
                else:
                    sorted_df['Hasta_dt'] = pd.NA

                sorted_df.sort_values(by='Periodo_dt', ascending=True, na_position='last', inplace=True)
                
                # --- Create the second sheet (Summary Report) with Date Expansion ---
                st.markdown("---")
                st.header("Generated Summary Report")

                try:
                    expanded_rows = []
                    
                    # Identify columns for the summary beforehand
                    id_col_name = next((c for c in sorted_df.columns if 'Identificación' in c), None)
                    potential_ibc_cols = [c for c in sorted_df.columns if 'ibc' in c.lower()]
                    potential_asig_cols = [
                        c for c in sorted_df.columns 
                        if ('asign' in c.lower() and ('básica' in c.lower() or 'basica' in c.lower())) or
                        ('último salario' in c.lower() or 'ultimo salario' in c.lower() or any(k in c.lower() for k in ['[5]', '[31]']))
                    ]

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
                        if pd.isna(end_date) or end_date < start_date:
                            # Single month case
                            s = start_date.replace(day=1)
                            e = s + pd.offsets.MonthEnd(0)
                            month_segments.append((s, e))
                        else:
                            # Split the range [start_date, end_date] into monthly segments with safety check
                            curr_s = start_date
                            safety_count = 0
                            while curr_s <= end_date and safety_count < 600: # Max 50 years
                                # End of the current month, but not beyond the overall end_date
                                curr_e = min(curr_s + pd.offsets.MonthEnd(0), end_date)
                                month_segments.append((curr_s, curr_e))
                                # Start of the next month (Robust increment to avoid infinite loops)
                                curr_s = (curr_s + pd.offsets.MonthEnd(0) + pd.offsets.Day(1))
                                safety_count += 1

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
                            # Clean ID for lookup (only digits)
                            raw_id = str(row[id_col_name]).strip() if id_col_name else ""
                            clean_row_id = re.sub(r'\D', '', raw_id)
                            
                            month_key = seg_start.strftime('%Y%m') # Matches the YYYYMM format
                            
                            # Get days from the document lookup first
                            final_days = 0
                            if (clean_row_id, month_key) in days_lookup_45:
                                raw_val = days_lookup_45[(clean_row_id, month_key)]
                                try:
                                    # Clean the value (handle decimal commas)
                                    final_days = float(str(raw_val).replace(',', '.'))
                                except Exception:
                                    final_days = 0
                            
                            # FALLBACK REMOVED: We no longer calculate days based on calendar range.
                            # If final_days is 0 (because it wasn't in lookup or was 0 in PDF), it stays 0.
                            
                            # Format IBC for summary: replace commas with dots as requested
                            formatted_ibc = str(final_ibc).replace(',', '.') if final_ibc is not None else None

                            new_row = {
                                'DOCUMENTO': raw_id,
                                'FECHA INICIAL': seg_start.strftime('%d/%m/%Y'),
                                'FECHA FINAL': seg_end.strftime('%d/%m/%Y'),
                                'IBC': formatted_ibc,
                                'DIAS': final_days,
                                '[40]IBC Reportado': ibc_lookup_40.get((clean_row_id, month_key), None),
                                'TOTAL DIAS': row[source_cols_map['TOTAL DIAS']] if source_cols_map['TOTAL DIAS'] else None,
                                'SEMANAS': row[source_cols_map['SEMANAS']] if source_cols_map['SEMANAS'] else None
                            }
                            expanded_rows.append(new_row)

                    summary_df = pd.DataFrame(expanded_rows)

                    # --- Deduplicate and Sort Summary Report by Date ---
                    if not summary_df.empty:
                        # Create a temp datetime column for robust sorting
                        summary_df['tmp_dt'] = pd.to_datetime(summary_df['FECHA INICIAL'], format='%d/%m/%Y', errors='coerce')
                        # Deduplicate: Keep rows with data if possible (sort by TOTAL DIAS descending first)
                        summary_df.sort_values(by=['tmp_dt', 'TOTAL DIAS'], ascending=[True, False], inplace=True)
                        summary_df.drop_duplicates(subset=['DOCUMENTO', 'FECHA INICIAL', 'FECHA FINAL', 'IBC'], keep='first', inplace=True)
                        # Final sort and cleanup
                        summary_df.sort_values(by='tmp_dt', ascending=True, inplace=True)
                        summary_df.drop(columns=['tmp_dt'], inplace=True)

                    # --- NEW: Create "Final_Report" with Weighted Averages (IBC) ---
                    def get_ibc_report(df_summary):
                        if df_summary.empty:
                            return pd.DataFrame()
                        
                        # 1. Deduplicate: Remove exactly identical rows
                        df = df_summary.drop_duplicates(subset=['DOCUMENTO', 'FECHA INICIAL', 'FECHA FINAL', 'IBC']).copy()
                        
                        # 2. Clean IBC and Days to numeric for math
                        def clean_to_num(val):
                            if pd.isna(val):
                                return 0.0
                            s = str(val).replace('$', '').replace('.', '').replace(',', '.').replace(' ', '').strip()
                            try:
                                return float(s)
                            except Exception:
                                return 0.0
                        
                        df['IBC_num'] = df['IBC'].apply(clean_to_num)
                        df['IBC40_num'] = df['[40]IBC Reportado'].apply(clean_to_num)

                        # NEW logic: choose the smaller IBC if both are non-zero
                        def choose_min_ibc(row):
                            v1 = row['IBC_num']
                            v2 = row['IBC40_num']
                            if v1 > 0 and v2 > 0:
                                return min(v1, v2)
                            return v1 if v1 > 0 else v2

                        df['FINAL_IBC_FOR_WEIGHT'] = df.apply(choose_min_ibc, axis=1)
                        
                        df['DIAS_num'] = pd.to_numeric(df['DIAS'], errors='coerce').fillna(0)
                        
                        df['FECHA_DT'] = pd.to_datetime(df['FECHA INICIAL'], format='%d/%m/%Y')
                        df['MONTH_KEY'] = df['FECHA_DT'].dt.to_period('M')
                        
                        # 3. Weighted calculation: (Days * Final IBC)
                        df['WEIGHTED_VAL'] = df['DIAS_num'] * df['FINAL_IBC_FOR_WEIGHT']
                        
                        # 4. Group by Month and Year
                        final_rows = []
                        for period, group in df.groupby('MONTH_KEY'):
                            # Rule: If multiple rows for the same month, take the maximum days found
                            max_days = group['DIAS_num'].max()
                            
                            # NEW FILTER: Skip months with 0 days reported
                            if max_days <= 0:
                                continue

                            total_weighted_sum = group['WEIGHTED_VAL'].sum()
                            
                            # Logic for final date based on number of days
                            start_date = period.start_time
                            if max_days < 30:
                                # End date corresponds to the number of days reported (e.g., 15 days -> ends on day 15)
                                end_date = start_date + pd.offsets.Day(int(max_days) - 1)
                                fecha_final_str = end_date.strftime('%d/%m/%Y')
                            else:
                                # 30 or more days -> defaults to the full month end
                                fecha_final_str = period.end_time.strftime('%d/%m/%Y')

                            # Calculate weighted IBC
                            ibc_weighted = total_weighted_sum / max_days
                            
                            doc = group['DOCUMENTO'].iloc[0] if 'DOCUMENTO' in group.columns else None
                            
                            final_rows.append({
                                'DOCUMENTO': doc,
                                'FECHA INICIAL': start_date.strftime('%d/%m/%Y'),
                                'FECHA FINAL': fecha_final_str,
                                'IBC(Ponderado)': int(ibc_weighted),
                                'DIAS TOTALES': max_days
                            })
                        
                        return pd.DataFrame(final_rows)

                    final_report_df = get_ibc_report(summary_df)

                    # --- NEW: Create "Gaps_Report" (Comprehensive gaps including partial months) ---
                    def get_gaps_report(df_final):
                        if df_final is None or df_final.empty:
                            return pd.DataFrame()
                        
                        # We copy to avoid modifying the original dataframe
                        df_g = df_final.copy()
                        df_g['FECHA_INI_DT'] = pd.to_datetime(df_g['FECHA INICIAL'], format='%d/%m/%Y')
                        df_g['FECHA_FIN_DT'] = pd.to_datetime(df_g['FECHA FINAL'], format='%d/%m/%Y')
                        
                        start_limit = df_g['FECHA_INI_DT'].min()
                        end_limit = df_g['FECHA_INI_DT'].max()
                        
                        if pd.isna(start_limit) or pd.isna(end_limit):
                            return pd.DataFrame()
                        
                        # 1. Map existing data by month for easy lookup
                        full_range_months = pd.date_range(start=start_limit.replace(day=1), end=end_limit.replace(day=1), freq='MS').to_period('M')
                        existing_data = {row['FECHA_INI_DT'].to_period('M'): row for _, row in df_g.iterrows()}
                        
                        # 2. Identify segments of gaps
                        all_gap_segments = []
                        for period in full_range_months:
                            if period not in existing_data:
                                # Total missing month: assume 30 days gap
                                all_gap_segments.append({
                                    'start': period.start_time,
                                    'end': period.end_time,
                                    'theoretical_days': 30
                                })
                            else:
                                row = existing_data[period]
                                # Partial missing month: if days < 30, gap from end of data to end of month
                                if row['DIAS TOTALES'] < 30:
                                    gap_start = row['FECHA_FIN_DT'] + pd.offsets.Day(1)
                                    # Only add if the gap start is still within the same month
                                    if gap_start.month == row['FECHA_FIN_DT'].month:
                                        all_gap_segments.append({
                                            'start': gap_start,
                                            'end': period.end_time,
                                            'theoretical_days': 30 - row['DIAS TOTALES']
                                        })

                        if not all_gap_segments:
                            return pd.DataFrame()

                        # 3. Group consecutive segments into ranges
                        merged_ranges = []
                        if all_gap_segments:
                            current = all_gap_segments[0].copy()
                            
                            for i in range(1, len(all_gap_segments)):
                                next_seg = all_gap_segments[i]
                                # Check if they are effectively consecutive (less than 2 days apart)
                                if (next_seg['start'] - current['end']).days <= 1:
                                    current['end'] = next_seg['end']
                                    current['theoretical_days'] += next_seg['theoretical_days']
                                else:
                                    merged_ranges.append(current)
                                    current = next_seg.copy()
                        merged_ranges.append(current)

                        # 4. Format the final output
                        gap_rows = []
                        for r in merged_ranges:
                            weeks = r['theoretical_days'] / 7
                            gap_rows.append({
                                'FECHA INICIAL': r['start'].strftime('%d/%m/%Y'),
                                'FECHA FINAL': r['end'].strftime('%d/%m/%Y'),
                                'DÍAS FALTANTES (Base 30)': r['theoretical_days'],
                                'SEMANAS': round(weeks, 2)
                            })
                        
                        return pd.DataFrame(gap_rows)

                    gaps_report_df = get_gaps_report(final_report_df)

                    st.write("#### Summary Report Preview")
                    st.dataframe(summary_df.head())

                    if not final_report_df.empty:
                        st.write("#### Final Weighted Report Preview (IBC)")
                        st.dataframe(final_report_df.head())
                    
                    if not gaps_report_df.empty:
                        st.write("#### Missing Months Report Preview")
                        st.dataframe(gaps_report_df.head())

                    # --- Prepare Excel file with FOUR sheets ---
                    # Drop temporary datetime columns from the main reports
                    cols_to_drop = ['Periodo_dt']
                    if 'Hasta_dt' in sorted_df.columns: cols_to_drop.append('Hasta_dt')
                    sorted_df.drop(columns=cols_to_drop, inplace=True)
                    
                    # Cleanup temporary column in final report if it exists
                    if 'FECHA_DT' in final_report_df.columns:
                        final_report_df.drop(columns=['FECHA_DT'], inplace=True)
                    
                    # Use a new function to write multiple sheets
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        sorted_df.to_excel(writer, sheet_name='Consolidated_Report', index=False)
                        summary_df.to_excel(writer, sheet_name='Summary_Report', index=False)
                        if not final_report_df.empty:
                            final_report_df.to_excel(writer, sheet_name='Final_Report', index=False)
                        if not gaps_report_df.empty:
                            gaps_report_df.to_excel(writer, sheet_name='Missing_Months', index=False)
                    
                    excel_data_with_summary = output.getvalue()
                    
                    st.download_button(
                        label="📥 Download Report (Multiple Sheets)",
                        data=excel_data_with_summary,
                        file_name=f"{uploaded_file.name.replace('.pdf', '')}_pension_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Error generating reports: {e}")
                    st.exception(e)
            else:
                st.warning("Could not find a unified 'Periodo' column to sort the report.")
        else:
            st.warning("No tables matching the required header concepts were found.")
            
    else:
        st.warning("No tables were found in the provided PDF file.")
