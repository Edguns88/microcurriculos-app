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

st.caption(
    "Departamento de Matemáticas • Programa de Ciencia de Datos • Universidad Externado de Colombia"
)
col1, col2 = st.columns([1, 3])

with col1:
    st.image("logo_ciencia_datos.png", width=140)

with col2:
    st.title("Automatizador de Microcurrículos")
st.write(
    """
    Esta herramienta ha sido desarrollada para apoyar el proceso de actualización,
    estandarización y consolidación de los microcurrículos del Programa de Ciencia de Datos
    del Departamento de Matemáticas de la Universidad Externado de Colombia, en el marco
    de las actividades asociadas a los procesos de acreditación académica.

    La aplicación toma como insumos la plantilla institucional vigente, la Matriz de
    Alineación Curricular y los microcurrículos históricos en formato PDF, con el fin de
    generar versiones actualizadas en formato Word conservando la estructura oficial del
    documento institucional.
    """
)

st.info(
    "Los archivos PDF que se carguen deben corresponder a microcurrículos existentes "
    "y aprobados previamente al proceso de actualización y acreditación."
)
st.warning("Versión de prueba. Revisa el documento generado antes de usarlo oficialmente.")

plantilla = st.file_uploader(
    "1. Sube la plantilla Word institucional (.docx)",
    type=["docx"]
)

matriz = st.file_uploader(
    "2. Sube la Matriz de Alineación Curricular (.xlsx)",
    type=["xlsx"]
)

pdfs = st.file_uploader(
    "3. Sube uno o varios PDF de microcurrículos",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Generar microcurrículos"):
    if plantilla is None:
        st.error("Primero debes subir la plantilla Word.")
    elif matriz is None:
        st.error("Debes subir la Matriz de Alineación Curricular.")
    elif not pdfs:
        st.error("Debes subir al menos un PDF.")
    else:
        with st.spinner("Procesando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                template_path = tmpdir / plantilla.name
                template_path.write_bytes(plantilla.getbuffer())

                matriz_path = tmpdir / matriz.name
                matriz_path.write_bytes(matriz.getbuffer())

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
                        salida, diag = procesar_pdf(template_path, pdf_path, output_dir, matriz_path=matriz_path)
                        archivos_generados.append(salida)
                        diagnosticos.append(diag)
                    except Exception as e:
                        diagnosticos.append({
                            "pdf": pdf_path.name,
                            "error": str(e)
                        })

                if len(archivos_generados) == 1:
                    archivo = archivos_generados[0]
                    st.download_button(
                        label="Descargar microcurrículo Word",
                        data=archivo.read_bytes(),
                        file_name=archivo.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    zip_path = tmpdir / "microcurriculos_diligenciados.zip"
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                        for archivo in output_dir.glob("*"):
                            z.write(archivo, arcname=archivo.name)

                    st.download_button(
                        label="Descargar ZIP con documentos generados",
                        data=zip_path.read_bytes(),
                        file_name="microcurriculos_diligenciados.zip",
                        mime="application/zip"
                    )
