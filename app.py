import streamlit as st
import tempfile
import os
import subprocess
from pathlib import Path
from shutil import which
import platform

st.set_page_config(page_title="DOCX → PDF Converter", page_icon="📄➡️📕", layout="centered")

st.title("📄 DOCX → PDF Converter (Offline, No API Key)")

st.markdown("""
Upload a `.docx` file to convert it to `.pdf` **locally on the server**.

Supported methods:
- 🟢 **LibreOffice (Linux)** — preserves full formatting  
- 🟣 **docx2pdf (Windows only)** — uses Microsoft Word  
- 🟡 **Pandoc (Fallback)** — simple text-based conversion  
""")

uploaded = st.file_uploader("📂 Upload DOCX file", type=["docx"])

def is_libreoffice_available():
    return which("soffice") or which("libreoffice")

def is_docx2pdf_available():
    if platform.system() == "Windows":
        try:
            import docx2pdf
            return True
        except:
            return False
    return False  # disable on Linux/macOS

def is_pandoc_available():
    return which("pandoc") and which("pdflatex")

def convert_with_libreoffice(docx_path, outdir):
    exe = which("soffice") or which("libreoffice")
    subprocess.run([exe, "--headless", "--convert-to", "pdf", docx_path, "--outdir", outdir],
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("LibreOffice failed to create PDF.")
    return pdf_path

def convert_with_pandoc(docx_path, outdir):
    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    subprocess.run(["pandoc", docx_path, "-o", pdf_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if not os.path.exists(pdf_path):
        raise RuntimeError("Pandoc failed to create PDF.")
    return pdf_path

def convert_docx_to_pdf(docx_bytes, filename):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, filename)
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        if is_libreoffice_available():
            backend = "LibreOffice"
            pdf_path = convert_with_libreoffice(docx_path, tmpdir)
        elif is_docx2pdf_available():
            from docx2pdf import convert
            backend = "docx2pdf"
            convert(docx_path, tmpdir)
            pdf_path = os.path.join(tmpdir, Path(docx_path).stem + ".pdf")
        elif is_pandoc_available():
            backend = "Pandoc"
            pdf_path = convert_with_pandoc(docx_path, tmpdir)
        else:
            raise RuntimeError("No supported converter found. Please install LibreOffice or Pandoc.")

        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        return pdf_data, backend

if uploaded:
    if st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            try:
                pdf_bytes, backend = convert_docx_to_pdf(uploaded.getvalue(), uploaded.name)
                st.success(f"✅ Conversion successful using {backend}!")
                st.download_button(
                    label="⬇️ Download PDF",
                    data=pdf_bytes,
                    file_name=Path(uploaded.name).stem + ".pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Upload a DOCX file to start.")

st.caption(f"Running on {platform.system()} | Python {platform.python_version()}")
