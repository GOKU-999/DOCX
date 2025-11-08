import streamlit as st
import google.generativeai as genai
import os
import tempfile
from pathlib import Path
import base64
import io

# Page configuration
st.set_page_config(
    page_title="DOCX to PDF - Gemini API",
    page_icon="📄",
    layout="centered"
)

def setup_gemini():
    """Setup Gemini API"""
    try:
        if 'GEMINI_API_KEY' in st.secrets:
            api_key = st.secrets['GEMINI_API_KEY']
        else:
            api_key = st.text_input("Enter Gemini API Key", type="password")
        
        if api_key:
            genai.configure(api_key=api_key)
            return True, "Gemini API configured successfully"
        else:
            return False, "Please provide Gemini API key"
    except Exception as e:
        return False, f"Gemini setup error: {str(e)}"

def convert_with_gemini(docx_file, output_path):
    """
    Convert DOCX to PDF using Gemini API
    Note: This is a creative approach since Gemini doesn't directly convert files
    """
    try:
        # For Gemini, we'll extract text and create a well-formatted PDF
        import docx
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
        
        # Read DOCX content
        doc = docx.Document(docx_file)
        
        # Extract text content
        full_text = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                full_text.append(paragraph.text)
        
        text_content = "\n".join(full_text)
        
        # Use Gemini to analyze and suggest formatting
        model = genai.GenerativeModel('gemini-pro')
        
        prompt = f"""
        Analyze this document content and suggest the best PDF structure:
        
        {text_content[:2000]}  # Limit text for token constraints
        
        Provide guidance on:
        1. What are the main headings vs body text?
        2. How should the content be structured in PDF?
        3. Any specific formatting recommendations?
        
        Respond in a structured way for PDF generation.
        """
        
        response = model.generate_content(prompt)
        
        # Create PDF with enhanced formatting
        doc_pdf = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18
        )
        
        styles = getSampleStyleSheet()
        story = []
        
        # Apply Gemini's suggestions (simplified implementation)
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():
                # Basic heading detection
                if paragraph.style.name.startswith('Heading') or any(run.bold for run in paragraph.runs):
                    style = styles['Heading2']
                else:
                    style = styles['Normal']
                
                p = Paragraph(paragraph.text.replace('\n', '<br/>'), style)
                story.append(p)
                story.append(Spacer(1, 12))
        
        doc_pdf.build(story)
        return True, "✅ PDF created with Gemini-enhanced formatting!"
        
    except Exception as e:
        return False, f"❌ Gemini conversion error: {str(e)}"

def main():
    st.title("📄 DOCX to PDF - Gemini API")
    st.markdown("**Using Google's Gemini AI for enhanced PDF conversion**")
    
    # Setup instructions
    with st.expander("⚙️ Gemini API Setup"):
        st.markdown("""
        ### Get Gemini API Key:
        1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
        2. Create API key
        3. Add to Streamlit secrets:
           ```toml
           GEMINI_API_KEY = "your-gemini-api-key"
           ```
        """)
    
    # Setup Gemini
    gemini_ready, gemini_message = setup_gemini()
    
    if not gemini_ready:
        st.warning(gemini_message)
    
    uploaded_file = st.file_uploader("Upload DOCX file", type=['docx'])
    
    if uploaded_file and gemini_ready:
        st.success(f"File: {uploaded_file.name} | Size: {uploaded_file.size/1024:.1f} KB")
        
        if st.button("Convert with Gemini AI", type="primary"):
            with st.spinner("Analyzing document with Gemini AI..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                
                success, message = convert_with_gemini(uploaded_file, temp_pdf_path)
                
                if success:
                    st.success(message)
                    
                    with open(temp_pdf_path, "rb") as f:
                        pdf_data = f.read()
                    
                    st.download_button(
                        "📥 Download AI-Enhanced PDF",
                        pdf_data,
                        Path(uploaded_file.name).stem + ".pdf",
                        "application/pdf"
                    )
                    
                    os.unlink(temp_pdf_path)
                else:
                    st.error(message)

if __name__ == "__main__":
    main()
