import streamlit as st
import zipfile
import os
import pandas as pd
import PyPDF2
from io import BytesIO
import tempfile

def extract_zip(uploaded_file):
    """Extracts uploaded ZIP file to a temporary directory."""
    temp_dir = tempfile.mkdtemp()
    
    with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir

def merge_pdfs(pdf_files):
    """Merges multiple PDFs into a single PDF with a limit to avoid infinite loops."""
    merger = PyPDF2.PdfMerger()
    max_files = 100  # Limiting the number of files to prevent excessive processing
    pdf_files = sorted(pdf_files)[:max_files]  # Process only the first 100 PDFs
    
    for pdf in pdf_files:
        merger.append(pdf)
    
    output_pdf = BytesIO()
    merger.write(output_pdf)
    merger.close()
    output_pdf.seek(0)
    return output_pdf

def process_uploaded_file(uploaded_file):
    """Handles the file upload and processing logic."""
    if not uploaded_file:
        return None
    
    temp_dir = extract_zip(uploaded_file)
    pdf_files = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.endswith('.pdf')]
    excel_files = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.endswith('.xlsx')]
    
    # Load and process Excel file (if needed)
    if excel_files:
        excel_data = pd.ExcelFile(excel_files[0])
        st.write("גיליונות שנמצאו באקסל:", excel_data.sheet_names)
    
    if not pdf_files:
        st.error("לא נמצאו קובצי PDF בארכיון.")
        return None
    
    merged_pdf = merge_pdfs(pdf_files)
    return merged_pdf

def main():
    st.title("איחוד קובצי PDF מתוך ZIP")
    
    if "processed" not in st.session_state:
        st.session_state.processed = False
    
    uploaded_file = st.file_uploader("העלה קובץ ZIP", type=["zip"])
    
    if uploaded_file is not None and not st.session_state.processed:
        st.session_state.processed = True
        st.write("הקובץ נטען בהצלחה! מבצע עיבוד...")
        merged_pdf = process_uploaded_file(uploaded_file)
        
        if merged_pdf:
            st.success("העיבוד הסתיים! ניתן להוריד את הקובץ המאוחד למטה.")
            st.download_button(
                label="הורד את הקובץ המאוחד",
                data=merged_pdf,
                file_name="merged.pdf",
                mime="application/pdf"
            )

if __name__ == "__main__":
    main()
