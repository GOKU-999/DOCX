import streamlit as st
import tempfile
import os
import subprocess
from pathlib import Path
from shutil import which
import platform

st.set_page_config(page_title="DOCX → PDF BY GOKU", page_icon="📄➡️📕", layout="centered")

st.title("📄 CONVERT YOUR DOCX FILE TO PDF FILE......")

st.markdown("""
Upload a `.docx` file to convert it to `.pdf` **locally on the server**.

 MADE BY :- GOKU(PRATIK)
 
""")

uploaded = st.file_uploader("📂 Upload your DOCX file", type=["docx"])

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
    if not exe:
        raise RuntimeError("LibreOffice not found. Make sure it's installed via packages.txt")

    cmd = [
        exe,
        "--headless",
        "--invisible",
        "--norestore",
        "--convert-to", "pdf",
        docx_path,
        "--outdir", outdir
    ]

    process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    if process.returncode != 0:
        st.error(f"LibreOffice failed:\n{process.stderr}")
        raise RuntimeError("LibreOffice conversion failed. Check the above error output.")

    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("LibreOffice did not produce a PDF file.")
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
    if st.button("Convert your DOCX to PDF"):
        with st.spinner("🕰️ Wait for few seconds ! Converting......"):
            try:
                pdf_bytes, backend = convert_docx_to_pdf(uploaded.getvalue(), uploaded.name)
                st.success(f"✅  SUCCESSFULLY CONVERTED TO PDF USING {backend}!")
                st.download_button(
                    label="⬇️ Download your PDF",
                    data=pdf_bytes,
                    file_name=Path(uploaded.name).stem + ".pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Upload your DOCX file .")

st.caption(" 🤗 THANK YOU FOR USING.VESITE AGAIN..... 😇")


