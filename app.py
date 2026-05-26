import streamlit as st
import pandas as pd
import os
import zipfile
import subprocess
import shutil
import tempfile
from docxtpl import DocxTemplate
from io import BytesIO

# ===============================
# CONFIGURACIÓN
# ===============================
st.set_page_config(
    page_title="DOCFLOW",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===============================
# ESTILOS
# ===============================
st.markdown("""
<style>
.metric-row {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
}

.metric-box {
    background: white;
    border: 1px solid #ddd;
    border-top: 3px solid #2563EB;
    border-radius: 6px;
    padding: 18px;
    flex: 1;
    text-align: center;
}

.metric-num {
    font-size: 32px;
    font-weight: bold;
    color: #2563EB;
}

.metric-label {
    font-size: 12px;
    color: gray;
}

/* Botones */
.stButton > button {
    background-color: #2563EB !important;
    color: white !important;
}

</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown("""
<div style="background:#2563EB;padding:16px;color:white;font-weight:bold;font-size:20px;">
DOCFLOW – Generación masiva de documentos
</div>
""", unsafe_allow_html=True)

# ===============================
# SESSION STATE
# ===============================
if "procesado" not in st.session_state:
    st.session_state.procesado = False

if "zip_buffer" not in st.session_state:
    st.session_state.zip_buffer = None

if "contador" not in st.session_state:
    st.session_state.contador = 0

if "pdf_count" not in st.session_state:
    st.session_state.pdf_count = 0

# ===============================
# SUBIDA ARCHIVOS
# ===============================
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Archivo Excel", type=["xlsx"])

with col2:
    docx_template = st.file_uploader("Plantilla Word", type=["docx"])

# ===============================
# PROCESAMIENTO
# ===============================
if excel_file and docx_template and not st.session_state.procesado:

    df = pd.read_excel(excel_file, dtype=str)
    df = df.dropna(how="all").reset_index(drop=True).fillna("")
    df = df.apply(lambda x: x.astype(str).str.upper())

    total = len(df)
    cols_count = len(df.columns)

    # ✅ METRICS (CORREGIDO)


    # PREVIEW
    st.dataframe(df.head(), use_container_width=True)

    formato = st.radio(
        "Formato",
        ["Word (.docx)", "PDF", "Ambos"]
    )

    if st.button("Procesar documentos", use_container_width=True):

        carpeta = "documentos"
        os.makedirs(carpeta, exist_ok=True)

        template_path = "plantilla.docx"
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        docx_generados = []

        # ===============================
        # GENERAR DOCX
        # ===============================
        for i, fila in enumerate(df.to_dict(orient="records")):

            doc = DocxTemplate(template_path)
            doc.render(fila)

            nombre = f"doc_{i}.docx"
            ruta = os.path.join(carpeta, nombre)

            doc.save(ruta)
            docx_generados.append(ruta)

        # ===============================
        # PDF
        # ===============================
        pdf_generados = []

        if formato in ["PDF", "Ambos"]:

            for docx_path in docx_generados:

                subprocess.run([
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    carpeta,
                    docx_path
                ])

                pdf_path = docx_path.replace(".docx", ".pdf")

                if os.path.exists(pdf_path):
                    pdf_generados.append(pdf_path)

        # ===============================
        # ZIP
        # ===============================
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:

            if formato in ["Word (.docx)", "Ambos"]:
                for f in docx_generados:
                    zipf.write(f, os.path.basename(f))

            if formato in ["PDF", "Ambos"]:
                for f in pdf_generados:
                    zipf.write(f, os.path.basename(f))

        zip_buffer.seek(0)

        st.session_state.procesado = True
        st.session_state.zip_buffer = zip_buffer
        st.session_state.contador = len(docx_generados)
        st.session_state.pdf_count = len(pdf_generados)

        st.rerun()

# ===============================
# RESULTADO
# ===============================
if st.session_state.procesado:

    html_result = f"""


</div>
"""
    st.markdown(html_result, unsafe_allow_html=True)

    st.download_button(
        "Descargar ZIP",
        st.session_state.zip_buffer,
        file_name="documentos.zip",
        use_container_width=True
    )

    if st.button("Limpiar"):

        st.session_state.procesado = False
        st.session_state.zip_buffer = None

        shutil.rmtree("documentos", ignore_errors=True)

        st.rerun()
