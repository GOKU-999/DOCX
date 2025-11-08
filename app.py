import streamlit as st
import os
import tempfile
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="DOCX to PDF Converter",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .success-msg {
        padding: 1rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.5rem;
        color: #155724;
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
        color: #000000 !important; /* Black text color */
    }
    .warning-box strong {
        color: #000000 !important; /* Black text for strong elements */
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">📄 DOCX to PDF Converter</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2rem;'>
    Convert your Microsoft Word documents (.docx) to PDF format instantly. 
    <br>Works on any platform - <strong>no Microsoft Word required!</strong>
    </div>
    """, unsafe_allow_html=True)
    
    # Platform info - NOW WITH BLACK TEXT
    st.markdown("""
    <div class="warning-box">
    ⚠️ <strong>Note:</strong> This converter preserves text content and basic formatting. 
    Complex formatting like tables, images, and advanced styling may not be fully preserved.
    </div>
    """, unsafe_allow_html=True)
    
    # Rest of your code remains the same...
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

if __name__ == "__main__":
    main()
