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
    page_title="Generador de Certificados",
    layout="wide"
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
</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown("""
<div style="background:#2563EB;padding:15px;color:white;font-size:20px;font-weight:bold;">
Generador de Certificados
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
# SUBIR ARCHIVOS
# ===============================
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("📊 Subir Excel", type=["xlsx"])

with col2:
    docx_template = st.file_uploader("📄 Subir plantilla Word", type=["docx"])

# ===============================
# PROCESAMIENTO
# ===============================
if excel_file and docx_template and not st.session_state.procesado:

    df = pd.read_excel(excel_file, dtype=str)
    df = df.dropna(how="all").fillna("")
    df.columns = df.columns.str.strip().str.lower()

    total = len(df)
    cols_count = len(df.columns)

    # ✅ MÉTRICAS
    html_metrics = f"""
<div class="metric-row">
    <div class="metric-box">
        <div class="metric-num">{total}</div>
        <div class="metric-label">Registros</div>
    </div>
    <div class="metric-box">
        <div class="metric-num">{cols_count}</div>
        <div class="metric-label">Columnas</div>
    </div>
</div>
"""
    st.markdown(html_metrics, unsafe_allow_html=True)

    # PREVIEW
    st.dataframe(df.head(), use_container_width=True)

    formato = st.radio(
        "Formato de salida",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"],
        horizontal=True
    )

    # ===============================
    # BOTÓN PROCESAR
    # ===============================
    if st.button("⚙️ Procesar documentos", use_container_width=True):

        carpeta = "certificados"
        os.makedirs(carpeta, exist_ok=True)

        template_path = "plantilla_temp.docx"
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        docx_generados = []

        for i, fila in df.iterrows():

            doc = DocxTemplate(template_path)
            doc.render(fila.to_dict())

            nombre = f"{fila.get('nro','')}_{fila.get('contratante','')}_{fila.get('poliza','')}.docx"

            # limpiar nombre
            nombre = nombre.replace("/", "_").replace("\\", "_")

            ruta = os.path.join(carpeta, nombre)

            doc.save(ruta)
            docx_generados.append(ruta)

        # ===============================
        # CONVERTIR A PDF
        # ===============================
        pdf_generados = []

        if formato in ["PDF", "Ambos (Word + PDF)"]:
            output_dir = tempfile.gettempdir()

            for docx_path in docx_generados:
                try:
                    subprocess.run(
                        [
                            "libreoffice",
                            "--headless",
                            "--convert-to",
                            "pdf",
                            "--outdir",
                            output_dir,
                            docx_path
                        ],
                        check=True,
                        capture_output=True
                    )

                    pdf_path = docx_path.replace(".docx", ".pdf")

                    if os.path.exists(pdf_path):
                        pdf_generados.append(pdf_path)

                except:
                    st.warning(f"No se pudo convertir: {docx_path}")

        # ===============================
        # CREAR ZIP
        # ===============================
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:

            if formato in ["Word (.docx)", "Ambos (Word + PDF)"]:
                for f in docx_generados:
                    zipf.write(f, os.path.basename(f))

            if formato in ["PDF", "Ambos (Word + PDF)"]:
                for f in pdf_generados:
                    zipf.write(f, os.path.basename(f))

        zip_buffer.seek(0)

        # guardar estado
        st.session_state.procesado = True
        st.session_state.zip_buffer = zip_buffer
        st.session_state.contador = len(docx_generados)
        st.session_state.pdf_count = len(pdf_generados)
        st.session_state.formato = formato

        st.rerun()

# ===============================
# RESULTADOS
# ===============================
if st.session_state.procesado and st.session_state.zip_buffer:

    fmt = st.session_state.formato

    html_result = f"""
<div class="metric-row">
    <div class="metric-box">
        <div class="metric-num">{st.session_state.contador}</div>
        <div class="metric-label">DOCX generados</div>
    </div>
    <div class="metric-box">
        <div class="metric-num">{st.session_state.pdf_count if fmt != "Word (.docx)" else "—"}</div>
        <div class="metric-label">PDF generados</div>
    </div>
    <div class="metric-box">
        <div class="metric-num">✓</div>
        <div class="metric-label">ZIP listo</div>
    </div>
</div>
"""
    st.markdown(html_result, unsafe_allow_html=True)

    st.download_button(
        label="📦 Descargar ZIP",
        data=st.session_state.zip_buffer,
        file_name="certificados.zip",
        mime="application/zip",
        use_container_width=True
    )

    # limpiar
    if st.button("🧹 Nuevo proceso", use_container_width=True):

        st.session_state.procesado = False
        st.session_state.zip_buffer = None
        st.session_state.contador = 0
        st.session_state.pdf_count = 0

        shutil.rmtree("certificados", ignore_errors=True)

        if os.path.exists("plantilla_temp.docx"):
            os.remove("plantilla_temp.docx")

        st.rerun()
