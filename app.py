import streamlit as st
import os
import tempfile
from pathlib import Path
import sys

# Try to import required packages with error handling
try:
    from docx2pdf import convert
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

try:
    import pythoncom
    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False

try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

try:
    from htmldocx import HtmlToDocx
    HAS_HTMLDOCX = True
except ImportError:
    HAS_HTMLDOCX = False

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
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def convert_docx_to_pdf_simple(docx_file, output_path):
    """
    Simple conversion using available methods
    """
    try:
        # Method 1: Try docx2pdf (works on Windows with Word installed)
        if HAS_DOCX2PDF:
            # Create a temporary file for the DOCX
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx.write(docx_file.getvalue())
                temp_docx_path = temp_docx.name
            
            # Initialize COM if on Windows
            if HAS_PYTHONCOM:
                pythoncom.CoInitialize()
            
            # Convert to PDF
            convert(temp_docx_path, output_path)
            
            # Clean up temporary file
            os.unlink(temp_docx_path)
            
            return True, "Conversion successful using docx2pdf"
        
        # Method 2: Fallback to basic text conversion with FPDF
        elif HAS_FPDF:
            return convert_with_fpdf(docx_file, output_path)
        
        else:
            return False, "No conversion methods available. Please install required packages."
            
    except Exception as e:
        return False, f"Conversion error: {str(e)}"

def convert_with_fpdf(docx_file, pdf_path):
    """
    Basic conversion using FPDF (preserves text content only)
    """
    try:
        import docx
        from fpdf import FPDF
        
        # Read DOCX content
        doc = docx.Document(docx_file)
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        # Add text from DOCX
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                # Simple text extraction
                pdf.multi_cell(0, 10, paragraph.text)
        
        pdf.output(pdf_path)
        return True, "Conversion successful using FPDF (basic text preservation)"
    except Exception as e:
        return False, f"FPDF conversion error: {str(e)}"

def main():
    # Header
    st.markdown('<h1 class="main-header">📄 DOCX to PDF Converter</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    Welcome to the DOCX to PDF Converter! Upload your Microsoft Word document (.docx) 
    and convert it to PDF format instantly.
    """)
    
    # Platform info
    if not HAS_DOCX2PDF:
        st.markdown("""
        <div class="warning-box">
        ⚠️ <strong>Note:</strong> Using basic text conversion. For better formatting preservation, 
        run this app on Windows with Microsoft Word installed.
        </div>
        """, unsafe_allow_html=True)
    
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
                success, message = convert_docx_to_pdf_simple(uploaded_file, temp_pdf_path)
                
                if success:
                    st.markdown(f'<div class="success-msg">✅ {message}</div>', unsafe_allow_html=True)
                    
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
                    
                    # Show file info
                    pdf_size = len(pdf_data) / 1024
                    st.info(f"**Converted PDF size:** {pdf_size:.2f} KB")
                    
                    # Clean up temporary file
                    os.unlink(temp_pdf_path)
                    
                    st.balloons()
                    
                else:
                    st.markdown(f'<div class="error-msg">❌ {message}</div>', unsafe_allow_html=True)
                    
                    # Troubleshooting tips
                    with st.expander("Troubleshooting Tips"):
                        st.markdown("""
                        **If conversion fails:**
                        - Make sure the DOCX file is not corrupted
                        - Try uploading a different DOCX file
                        - For better formatting, run on Windows with Microsoft Word
                        - Large files may take longer to process
                        
                        **Alternative solutions:**
                        - Use Google Docs (File > Download > PDF)
                        - Use Microsoft Word Online
                        - Use LibreOffice (free, cross-platform)
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
    
    # Platform info section
    with st.expander("🔧 Platform Information"):
        st.write("**Available conversion methods:**")
        st.write(f"- docx2pdf (Windows with MS Word): {'✅ Available' if HAS_DOCX2PDF else '❌ Not available'}")
        st.write(f"- FPDF (Basic text conversion): {'✅ Available' if HAS_FPDF else '❌ Not available'}")
        st.write(f"- Python COM (Windows): {'✅ Available' if HAS_PYTHONCOM else '❌ Not available'}")
        
        st.write("**For best results:**")
        st.write("- Run on Windows with Microsoft Word installed")
        "- For complex formatting, use the docx2pdf method"
        "- For simple text documents, FPDF works well"
    
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
