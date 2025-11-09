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
This app converts `.docx` files to `.pdf` **entirely offline** using one of these local tools:
- **LibreOffice (`soffice`)** → Best for Linux (preserves design & layout)
- **docx2pdf** → Best for Windows (uses Microsoft Word if installed)
- **Pandoc + LaTeX** → Fallback converter (may slightly alter style)

---
""")

uploaded = st.file_uploader("📂 Upload a DOCX file", type=["docx"])

def is_libreoffice_available():
    return which("soffice") is not None or which("libreoffice") is not None

def is_docx2pdf_available():
    try:
        import docx2pdf  # noqa: F401
        return True
    except:
        return False

def is_pandoc_available():
    return which("pandoc") is not None and which("pdflatex") is not None

def convert_with_libreoffice(docx_path, outdir):
    exe = which("soffice") or which("libreoffice")
    cmd = [exe, "--headless", "--convert-to", "pdf", docx_path, "--outdir", outdir]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("LibreOffice conversion failed.")
    return pdf_path

def convert_with_docx2pdf(docx_path, outdir):
    from docx2pdf import convert
    convert(docx_path, outdir)
    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("docx2pdf conversion failed.")
    return pdf_path

def convert_with_pandoc(docx_path, outdir):
    pdf_path = os.path.join(outdir, Path(docx_path).stem + ".pdf")
    subprocess.run(["pandoc", docx_path, "-o", pdf_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("Pandoc conversion failed.")
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
            backend = "docx2pdf"
            pdf_path = convert_with_docx2pdf(docx_path, tmpdir)
        elif is_pandoc_available():
            backend = "Pandoc"
            pdf_path = convert_with_pandoc(docx_path, tmpdir)
        else:
            raise RuntimeError("No suitable converter found! Install LibreOffice or docx2pdf or Pandoc.")

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

st.markdown("---")
st.caption(f"Running on: {platform.system()} | Python {platform.python_version()}")
st.caption("Tip: Install LibreOffice for best results.")
