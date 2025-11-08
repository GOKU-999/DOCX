import streamlit as st
import os
import tempfile
from docx2pdf import convert
import pythoncom
import sys
from pathlib import Path
import time

# Page configuration
st.set_page_config(
    page_title="DOCX to PDF Converter",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-msg {
        padding: 1rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.5rem;
        color: #155724;
    }
    .error-msg {
        padding: 1rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.5rem;
        color: #721c24;
    }
    .upload-section {
        border: 2px dashed #1f77b4;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def convert_docx_to_pdf(docx_file, output_path):
    """
    Convert DOCX file to PDF using docx2pdf
    """
    try:
        # Create a temporary file for the DOCX
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
            temp_docx.write(docx_file.getvalue())
            temp_docx_path = temp_docx.name
        
        # Convert to PDF
        convert(temp_docx_path, output_path)
        
        # Clean up temporary file
        os.unlink(temp_docx_path)
        
        return True, None
    except Exception as e:
        return False, str(e)

def main():
    # Header
    st.markdown('<h1 class="main-header">📄 DOCX to PDF Converter</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    Welcome to the DOCX to PDF Converter! Upload your Microsoft Word document (.docx) 
    and convert it to PDF format instantly.
    """)
    
    # File upload section
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Choose a DOCX file", 
        type=['docx'],
        help="Upload a .docx file to convert to PDF"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # Display file info
        file_details = {
            "Filename": uploaded_file.name,
            "File size": f"{uploaded_file.size / 1024:.2f} KB",
            "File type": uploaded_file.type
        }
        
        st.write("**File Details:**")
        st.json(file_details)
        
        # Preview file content (first few lines)
        try:
            import docx
            doc = docx.Document(uploaded_file)
            text_content = []
            for para in doc.paragraphs[:5]:  # Show first 5 paragraphs
                if para.text.strip():
                    text_content.append(para.text)
            
            if text_content:
                with st.expander("Preview Document Content (First 5 paragraphs)"):
                    for i, text in enumerate(text_content):
                        st.write(f"{i+1}. {text}")
        except Exception as e:
            st.warning(f"Could not preview document content: {str(e)}")
        
        # Convert button
        if st.button("🚀 Convert to PDF", type="primary", use_container_width=True):
            with st.spinner("Converting your document... This may take a few seconds."):
                # Create temporary file for PDF output
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                
                # Perform conversion
                success, error_message = convert_docx_to_pdf(uploaded_file, temp_pdf_path)
                
                if success:
                    st.markdown('<div class="success-msg">✅ Conversion successful!</div>', unsafe_allow_html=True)
                    
                    # Read the converted PDF
                    with open(temp_pdf_path, "rb") as f:
                        pdf_data = f.read()
                    
                    # Download button
                    download_filename = Path(uploaded_file.name).stem + ".pdf"
                    st.download_button(
                        label="📥 Download PDF",
                        data=pdf_data,
                        file_name=download_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                    # Clean up temporary file
                    os.unlink(temp_pdf_path)
                    
                    st.balloons()
                    
                else:
                    st.markdown(f'<div class="error-msg">❌ Conversion failed: {error_message}</div>', unsafe_allow_html=True)
                    
                    # Fallback method suggestion
                    st.info("""
                    **Troubleshooting tips:**
                    - Make sure the DOCX file is not corrupted
                    - Try uploading a different DOCX file
                    - Ensure the file is a valid Word document
                    """)

    # Instructions section
    with st.expander("ℹ️ How to use this converter"):
        st.markdown("""
        1. **Upload**: Click the 'Browse files' button and select your .docx file
        2. **Preview**: Check the file details and content preview
        3. **Convert**: Click the 'Convert to PDF' button
        4. **Download**: Click the 'Download PDF' button to save your converted file
        
        **Supported formats:** .docx files only
        **Maximum file size:** 200MB (Streamlit limit)
        """)
    
    # Features section
    with st.expander("✨ Features"):
        st.markdown("""
        - ✅ Fast and reliable conversion
        - ✅ Preserves formatting and layout
        - ✅ Secure processing (files are temporarily stored)
        - ✅ No registration required
        - ✅ Free to use
        """)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "Made with ❤️ using Streamlit | DOCX to PDF Converter"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()