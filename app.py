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
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

:root {
    --primary: #2563EB;
    --dark: #1A1A1A;
    --gray: #F5F5F5;
    --mid: #6B6B6B;
    --border: #E0E0E0;
    --white: #FFFFFF;
}

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
    color: var(--dark);
}

.stApp {
    background-color: var(--gray);
}

.header-bar {
    background: var(--primary);
    padding: 18px 36px;
    margin: -1rem -1rem 2rem -1rem;
    display: flex;
    align-items: center;
    gap: 16px;
}

.logo-text {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 28px;
    font-weight: 700;
    color: white;
}

.header-sub {
    font-size: 13px;
    color: rgba(255,255,255,0.8);
    margin-left: auto;
}

.metric-row {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
}

.metric-box {
    background: white;
    border: 1px solid var(--border);
    border-top: 3px solid var(--primary);
    border-radius: 6px;
    padding: 18px;
    flex: 1;
    text-align: center;
}

.metric-num {
    font-size: 34px;
    font-weight: 700;
    color: var(--primary);
}

.metric-label {
    font-size: 12px;
    color: var(--mid);
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown("""
<div class="header-bar">
    <div class="logo-text">DOCFLOW</div>
    <div class="header-sub">Generación masiva de documentos</div>
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
# SUBIDA DE ARCHIVOS
# ===============================
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Archivo Excel", type=["xlsx"])

with col2:
    docx_template = st.file_uploader("Plantilla Word", type=["docx"])

# ===============================
# PROCESAMIENTO
# ===============================
if excel_file and docx_template:

    df = pd.read_excel(excel_file, dtype=str).fillna("")
    df = df.apply(lambda x: x.astype(str).str.upper())

    total = len(df)
    cols_count = len(df.columns)

    # ✅ MÉTRICAS CORREGIDAS
    st.markdown(f"""
    <div class="metric-row">

        <div class="metric-box">
            <div class="metric-num">{total}</div>
            <div class="metric-label">Registros detectados</div>
        </div>

        <div class="metric-box">
            <div class="metric-num">{cols_count}</div>
            <div class="metric-label">Campos encontrados</div>
        </div>

    </div>
    """, unsafe_allow_html=True)

    # preview
    st.dataframe(df.head())

    formato = st.radio(
        "Formato de salida",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"]
    )

    if st.button("Procesar"):

        carpeta = "docs"
        os.makedirs(carpeta, exist_ok=True)

        template_path = "plantilla.docx"
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        archivos = []

        for i, fila in enumerate(df.to_dict(orient="records")):
            doc = DocxTemplate(template_path)
            doc.render(fila)

            nombre = f"doc_{i}.docx"
            ruta = os.path.join(carpeta, nombre)
            doc.save(ruta)
            archivos.append(ruta)

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for a in archivos:
                zipf.write(a, os.path.basename(a))

        zip_buffer.seek(0)

        st.download_button(
            "Descargar ZIP",
            zip_buffer,
            file_name="docs.zip"
        )
