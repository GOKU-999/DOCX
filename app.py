import streamlit as st
import os
import tempfile
from pathlib import Path
import base64

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
    .stButton button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

def convert_docx_to_pdf(docx_file, output_path):
    """
    Convert DOCX to PDF using reportlab (fully cross-platform)
    """
    try:
        import docx
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.units import inch
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        
        # Read the DOCX file
        doc = docx.Document(docx_file)
        
        # Create PDF document
        doc_pdf = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18
        )
        
        # Create styles
        styles = getSampleStyleSheet()
        story = []
        
        # Process each paragraph in the DOCX
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():  # Only add non-empty paragraphs
                # Determine style based on paragraph properties
                if paragraph.style.name.startswith('Heading'):
                    style = styles['Heading2']
                else:
                    style = styles['Normal']
                
                # Create paragraph for PDF
                p = Paragraph(paragraph.text.replace('\n', '<br/>'), style)
                story.append(p)
                story.append(Spacer(1, 12))
        
        # Build PDF
        if story:  # Only build if there's content
            doc_pdf.build(story)
            return True, "✅ Conversion successful! PDF created with text content."
        else:
            return False, "❌ No text content found in the document."
            
    except ImportError as e:
        return False, f"❌ Required package missing: {str(e)}"
    except Exception as e:
        return False, f"❌ Conversion error: {str(e)}"

def convert_with_images(docx_file, output_path):
    """
    Alternative method that handles basic text formatting
    """
    try:
        import docx
        from fpdf import FPDF
        
        # Read DOCX
        doc = docx.Document(docx_file)
        
        # Create PDF
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Set font
        pdf.set_font("Arial", size=12)
        
        # Process content
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text:
                # Basic formatting detection
                is_bold = any(run.bold for run in paragraph.runs)
                is_large = any(run.font.size and run.font.size.pt > 12 for run in paragraph.runs) if paragraph.runs else False
                
                if is_bold or is_large:
                    pdf.set_font("Arial", 'B', 14)
                    pdf.cell(0, 10, text, ln=True)
                    pdf.set_font("Arial", size=12)
                else:
                    pdf.multi_cell(0, 10, text)
                pdf.ln(5)
        
        pdf.output(output_path)
        return True, "✅ Conversion successful! Basic formatting preserved."
        
    except Exception as e:
        return False, f"❌ FPDF conversion error: {str(e)}"

def get_file_preview(docx_file):
    """
    Extract text content for preview
    """
    try:
        import docx
        doc = docx.Document(docx_file)
        content = []
        
        for i, paragraph in enumerate(doc.paragraphs[:10]):  # First 10 paragraphs
            if paragraph.text.strip():
                content.append(f"¶{i+1}: {paragraph.text}")
                if len(content) >= 5:  # Show max 5 paragraphs in preview
                    break
        
        return content
    except:
        return ["Could not preview document content"]

def main():
    # Header
    st.markdown('<h1 class="main-header">📄 DOCX to PDF Converter</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2rem;'>
    Convert your Microsoft Word documents (.docx) to PDF format instantly. 
    <br>Works on any platform - <strong>no Microsoft Word required!</strong>
    </div>
    """, unsafe_allow_html=True)
    
    # Platform info
    st.markdown("""
    <div class="warning-box">
    ⚠️ <strong>Note:</strong> This converter preserves text content and basic formatting. 
    Complex formatting like tables, images, and advanced styling may not be fully preserved.
    </div>
    """, unsafe_allow_html=True)
    
    # File upload section
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Drag and drop or click to upload a DOCX file", 
        type=['docx'],
        help="Upload any .docx file to convert to PDF"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # Display file info
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**File:** {uploaded_file.name}")
        with col2:
            st.info(f"**Size:** {uploaded_file.size / 1024:.2f} KB")
        
        # Preview
        with st.expander("📋 Document Preview (first 5 paragraphs)", expanded=True):
            preview_content = get_file_preview(uploaded_file)
            for item in preview_content:
                st.write(item)
        
        # Conversion options
        st.subheader("Conversion Options")
        conversion_method = st.radio(
            "Choose conversion method:",
            ["Standard (Recommended)", "Basic Text Only"],
            help="Standard preserves some formatting, Basic is faster for simple documents"
        )
        
        # Convert button
        if st.button("🚀 Convert to PDF", type="primary", use_container_width=True):
            with st.spinner("Converting your document... Please wait."):
                # Create temporary file for PDF output
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                
                # Perform conversion based on selected method
                if conversion_method == "Standard (Recommended)":
                    success, message = convert_docx_to_pdf(uploaded_file, temp_pdf_path)
                else:
                    success, message = convert_with_images(uploaded_file, temp_pdf_path)
                
                # Fallback if standard method fails
                if not success and conversion_method == "Standard (Recommended)":
                    st.warning("Standard method failed, trying basic method...")
                    success, message = convert_with_images(uploaded_file, temp_pdf_path)
                
                if success:
                    st.markdown(f'<div class="success-msg">{message}</div>', unsafe_allow_html=True)
                    
                    # Read the converted PDF
                    with open(temp_pdf_path, "rb") as f:
                        pdf_data = f.read()
                    
                    # Download button
                    download_filename = Path(uploaded_file.name).stem + ".pdf"
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Download PDF File",
                            data=pdf_data,
                            file_name=download_filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    with col2:
                        # Show PDF info
                        pdf_size = len(pdf_data) / 1024
                        st.info(f"**PDF Size:** {pdf_size:.2f} KB")
                    
                    # Clean up temporary file
                    os.unlink(temp_pdf_path)
                    
                    st.balloons()
                    
                else:
                    st.markdown(f'<div class="error-msg">{message}</div>', unsafe_allow_html=True)
                    
                    # Installation help
                    with st.expander("🔧 Installation Help"):
                        st.markdown("""
                        **If you see package errors, install required packages:**
                        ```bash
                        pip install python-docx reportlab fpdf
                        ```
                        
                        **For Streamlit Cloud, add to requirements.txt:**
                        ```txt
                        python-docx>=1.1.0
                        reportlab>=4.0.0
                        fpdf2>=2.7.4
                        ```
                        """)

    # Features section
    with st.expander("✨ Features & Limitations"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **✅ What's Preserved:**
            - Text content
            - Paragraph structure  
            - Basic formatting (bold, headings)
            - Page breaks
            """)
        
        with col2:
            st.markdown("""
            **❌ Limitations:**
            - Complex tables
            - Images and charts
            - Advanced styling
            - Font variations
            """)
        
        st.markdown("""
        **💡 For complex documents:** Use Microsoft Word, Google Docs, or LibreOffice for full formatting preservation.
        """)

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "Made with ❤️ using Streamlit | Works on Windows, Mac, Linux & Cloud"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
