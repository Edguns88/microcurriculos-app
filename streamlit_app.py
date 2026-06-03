import streamlit as st
from pathlib import Path
import tempfile
import zipfile
import json

from motor_microcurriculos import procesar_pdf


st.set_page_config(
    page_title="Automatizador de Microcurrículos",
    page_icon="📄",
    layout="centered"
)

st.title("Automatizador de Microcurrículos")
st.write(
    "Sube la plantilla Word institucional y uno o varios PDF de microcurrículos. "
    "La aplicación generará los documentos Word diligenciados."
)

st.warning(
    "Versión inicial web. El resultado debe revisarse antes de usarlo oficialmente."
)

plantilla = st.file_uploader(
    "1. Sube la plantilla Word institucional (.docx)",
    type=["docx"]
)

pdfs = st.file_uploader(
    "2. Sube uno o varios PDF de microcurrículos",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Generar microcurrículos"):
    if plantilla is None:
        st.error("Primero debes subir la plantilla Word.")
    elif not pdfs:
        st.error("Debes subir al menos un PDF.")
    else:
        with st.spinner("Procesando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                template_path = tmpdir / plantilla.name
                template_path.write_bytes(plantilla.getbuffer())

                pdf_paths = []
                for pdf in pdfs:
                    pdf_path = tmpdir / pdf.name
                    pdf_path.write_bytes(pdf.getbuffer())
                    pdf_paths.append(pdf_path)

                output_dir = tmpdir / "salidas"
                output_dir.mkdir(exist_ok=True)

                diagnosticos = []
                archivos_generados = []

                for pdf_path in pdf_paths:
                    try:
                        salida, diag = procesar_pdf(template_path, pdf_path, output_dir)
                        archivos_generados.append(salida)
                        diagnosticos.append(diag)
                    except Exception as e:
                        diagnosticos.append({
                            "pdf": pdf_path.name,
                            "error": str(e)
                        })

                diag_path = output_dir / "diagnostico_microcurriculos.json"
                diag_path.write_text(
                    json.dumps(diagnosticos, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )

                zip_path = tmpdir / "microcurriculos_diligenciados.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    for archivo in output_dir.glob("*"):
                        z.write(archivo, arcname=archivo.name)

                st.success("Proceso terminado.")

                st.write("Diagnóstico:")
                st.json(diagnosticos)

                st.download_button(
                    label="Descargar ZIP con documentos generados",
                    data=zip_path.read_bytes(),
                    file_name="microcurriculos_diligenciados.zip",
                    mime="application/zip"
                )
