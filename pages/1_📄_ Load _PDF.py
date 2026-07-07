import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io

st.set_page_config(page_title="Extraer Tablas de PDF", page_icon="📈", layout="wide")

st.markdown("# Extraer Tablas de PDF a Excel")
st.sidebar.header("Cargar PDF")
st.write(
    """
    Sube un archivo PDF para extraer automáticamente todas las tablas que contenga.
    Las tablas se mostrarán como una vista previa y podrás descargarlas en un solo archivo de Excel,
    con cada tabla en una hoja separada.
    """
)

def extract_tables_from_pdf(pdf_bytes):
    """
    Extrae todas las tablas de un PDF usando PyMuPDF y las devuelve como una lista de DataFrames.
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
                        # Convertir a DataFrame, usando la primera fila como cabecera si es posible
                        headers = table_data[0]
                        df = pd.DataFrame(table_data[1:], columns=headers)
                        all_tables.append(df)
    except Exception as e:
        st.error(f"Error al procesar el PDF: {e}")
    return all_tables

def to_excel(dfs: list[pd.DataFrame]):
    """
    Convierte una lista de DataFrames a un archivo Excel en memoria.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f'Tabla_{i+1}', index=False)
    processed_data = output.getvalue()
    return processed_data

# --- Interfaz de Streamlit ---
uploaded_file = st.file_uploader("Elige un archivo PDF", type="pdf")

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    
    with st.spinner("Buscando y extrayendo tablas del PDF..."):
        extracted_tables = extract_tables_from_pdf(pdf_bytes)

    if extracted_tables:
        st.success(f"¡Se encontraron {len(extracted_tables)} tablas en el documento!")
        
        # Mostrar una vista previa de cada tabla
        for i, table_df in enumerate(extracted_tables):
            st.write(f"**Vista previa de la Tabla {i+1}:**")
            st.dataframe(table_df.head())
            st.markdown("---")
            
        # Preparar el archivo Excel para la descarga
        excel_data = to_excel(extracted_tables)
        
        st.download_button(
            label="📥 Descargar todas las tablas en Excel",
            data=excel_data,
            file_name=f"{uploaded_file.name.replace('.pdf', '')}_tablas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No se encontraron tablas en el archivo PDF proporcionado.")
