import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io

st.set_page_config(page_title="Extract Tables from PDF", page_icon="📈", layout="wide")

st.markdown("# Extract Tables from PDF to Excel")
st.sidebar.header("Upload PDF")
st.write(
    """
    Upload a PDF file to automatically extract all the tables it contains.
    The tables will be displayed as a preview, and you can download them in a single Excel file,
    with each table in a separate sheet.
    """
)

def extract_tables_from_pdf(pdf_bytes):
    """
    Extracts all tables from a PDF using PyMuPDF and returns them as a list of DataFrames.
    """
    all_tables = []
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            tables = page.find_tables()
            if tables:
                for i, table in enumerate(tables):
                    table_data = table.extract()
                    if table_data:
                        # Convert to DataFrame, using the first row as header if possible
                        headers = table_data[0]
                        df = pd.DataFrame(table_data[1:], columns=headers)
                        all_tables.append(df)
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
    return all_tables

def to_excel(dfs: list[pd.DataFrame]):
    """
    Converts a list of DataFrames to an in-memory Excel file.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
    processed_data = output.getvalue()
    return processed_data

# --- Streamlit Interface ---
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    
    with st.spinner("Finding and extracting tables from the PDF..."):
        extracted_tables = extract_tables_from_pdf(pdf_bytes)

    if extracted_tables:
        st.success(f"Found {len(extracted_tables)} tables in the document!")
        
        # Display a preview of each table
        for i, table_df in enumerate(extracted_tables):
            st.write(f"**Table {i+1} Preview:**")
            st.dataframe(table_df.head())
            st.markdown("---")
            
        # Prepare the Excel file for download
        excel_data = to_excel(extracted_tables)
        
        st.download_button(
            label="📥 Download all tables to Excel",
            data=excel_data,
            file_name=f"{uploaded_file.name.replace('.pdf', '')}_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No tables were found in the provided PDF file.")
