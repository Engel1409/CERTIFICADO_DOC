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
# CONFIGURACIÓN DE PÁGINA
# ===============================
st.set_page_config(
    page_title="Generador de Certificados · Mapfre",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===============================
# ESTILOS MAPFRE
# ===============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

/* Variables */
:root {
    --mapfre-red:    #E2001A;
    --mapfre-dark:   #1A1A1A;
    --mapfre-gray:   #F5F5F5;
    --mapfre-mid:    #6B6B6B;
    --mapfre-border: #E0E0E0;
    --mapfre-white:  #FFFFFF;
}

/* Reset general */
html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
    color: var(--mapfre-dark);
}

/* Fondo */
.stApp {
    background-color: var(--mapfre-gray);
}

/* Header superior */
.header-bar {
    background: var(--mapfre-red);
    padding: 18px 36px;
    margin: -1rem -1rem 2rem -1rem;
    display: flex;
    align-items: center;
    gap: 16px;
}
.header-bar .logo-text {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 28px;
    font-weight: 700;
    color: white;
    letter-spacing: 1px;
}
.header-bar .header-sub {
    font-size: 13px;
    color: rgba(255,255,255,0.75);
    font-weight: 400;
    margin-left: auto;
}

/* Tarjetas de sección */
.card {
    background: var(--mapfre-white);
    border-radius: 4px;
    border-left: 4px solid var(--mapfre-red);
    padding: 24px 28px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
.card-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 16px;
    font-weight: 700;
    color: var(--mapfre-red);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 12px;
}

/* Métricas */
.metric-row {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
}
.metric-box {
    background: white;
    border: 1px solid var(--mapfre-border);
    border-top: 3px solid var(--mapfre-red);
    border-radius: 4px;
    padding: 16px 24px;
    flex: 1;
    text-align: center;
}
.metric-box .metric-num {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 36px;
    font-weight: 700;
    color: var(--mapfre-red);
    line-height: 1;
}
.metric-box .metric-label {
    font-size: 12px;
    color: var(--mapfre-mid);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 4px;
}

/* Divider rojo */
.red-divider {
    height: 2px;
    background: var(--mapfre-red);
    margin: 24px 0;
    opacity: 0.15;
}

/* Botones */
.stButton > button {
    font-family: 'Barlow', sans-serif !important;
    font-weight: 600 !important;
    border-radius: 3px !important;
    border: none !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"],
.stButton > button:first-child {
    background-color: var(--mapfre-red) !important;
    color: white !important;
}
.stButton > button:hover {
    opacity: 0.88 !important;
    transform: translateY(-1px) !important;
}

/* Download button */
.stDownloadButton > button {
    background-color: #1A1A1A !important;
    color: white !important;
    font-family: 'Barlow', sans-serif !important;
    font-weight: 600 !important;
    border-radius: 3px !important;
    border: none !important;
    width: 100% !important;
    padding: 14px !important;
    font-size: 15px !important;
}
.stDownloadButton > button:hover {
    background-color: var(--mapfre-red) !important;
    transform: translateY(-1px) !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: white;
    border: 1.5px dashed var(--mapfre-border);
    border-radius: 4px;
    padding: 8px;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--mapfre-red);
}

/* Radio buttons */
.stRadio > div {
    gap: 20px;
}
.stRadio label {
    font-weight: 500 !important;
}

/* Progress bar */
.stProgress > div > div {
    background-color: var(--mapfre-red) !important;
}

/* Alerts */
.stSuccess {
    border-left: 4px solid var(--mapfre-red) !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid var(--mapfre-border);
    border-radius: 4px;
}

/* Footer */
.footer-bar {
    text-align: center;
    font-size: 11px;
    color: var(--mapfre-mid);
    margin-top: 40px;
    padding-top: 16px;
    border-top: 1px solid var(--mapfre-border);
}
</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown("""
<div class="header-bar">
    <div class="logo-text">MAPFRE</div>
    <div class="header-sub">Emisión y Renovación Masivos · Generador de Certificados</div>
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
st.markdown('<div class="card"><div class="card-title">📁 Archivos de entrada</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Excel de pólizas (.xlsx)", type=["xlsx"])
with col2:
    docx_template = st.file_uploader("Plantilla Word (.docx)", type=["docx"])
st.markdown('</div>', unsafe_allow_html=True)

if excel_file and docx_template and not st.session_state.procesado:

    # ===============================
    # LEER EXCEL
    # ===============================
    df = pd.read_excel(excel_file, dtype=str)
    df.columns = df.columns.str.strip().str.lower()

    total = len(df)

    # Métricas
    cols_count = len(df.columns)
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-box">
            <div class="metric-num">{total}</div>
            <div class="metric-label">Pólizas detectadas</div>
        </div>
        <div class="metric-box">
            <div class="metric-num">{cols_count}</div>
            <div class="metric-label">Campos en plantilla</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Vista previa
    st.markdown('<div class="card"><div class="card-title">📊 Vista previa · primeras 10 filas</div>', unsafe_allow_html=True)
    st.dataframe(df.head(10), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ===============================
    # CONFIGURACIÓN DE SALIDA
    # ===============================
    st.markdown('<div class="card"><div class="card-title">⚙️ Configuración de salida</div>', unsafe_allow_html=True)
    formato = st.radio(
        "Formato de descarga",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"],
        horizontal=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ===============================
    # BOTÓN PROCESAR
    # ===============================
    if st.button("⚙️ Procesar certificados", use_container_width=True):

        carpeta = "certificados"
        os.makedirs(carpeta, exist_ok=True)

        progreso = st.progress(0)
        status   = st.empty()
        contador = 0

        template_path = "plantilla_temp.docx"
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        docx_generados = []
        for i, fila in df.iterrows():
            doc = DocxTemplate(template_path)
            contexto = fila.to_dict()
            doc.render(contexto)

            nombre = f"{fila.get('poliza', i)}.docx"
            ruta   = os.path.join(carpeta, nombre)
            doc.save(ruta)
            docx_generados.append(ruta)

            contador += 1
            progreso.progress((i + 1) / total)
            status.caption(f"Generando {contador} de {total}...")

        status.empty()

        # Conversión PDF
        pdf_generados = []
        if formato in ["PDF", "Ambos (Word + PDF)"]:
            pdf_status = st.empty()
            pdf_status.info("📄 Convirtiendo a PDF con LibreOffice...")
            output_dir = tempfile.gettempdir()

            for docx_path in docx_generados:
                try:
                    subprocess.run(
                        ["libreoffice", "--headless", "--convert-to", "pdf",
                         "--outdir", output_dir, docx_path],
                        check=True, capture_output=True
                    )
                    pdf_nombre = os.path.basename(docx_path).replace(".docx", ".pdf")
                    pdf_path   = os.path.join(output_dir, pdf_nombre)
                    if os.path.exists(pdf_path):
                        pdf_generados.append(pdf_path)
                except subprocess.CalledProcessError:
                    st.warning(f"⚠️ No se pudo convertir: {os.path.basename(docx_path)}")

            pdf_status.empty()

        # Armar ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            if formato in ["Word (.docx)", "Ambos (Word + PDF)"]:
                for ruta in docx_generados:
                    zipf.write(ruta, os.path.basename(ruta))
            if formato in ["PDF", "Ambos (Word + PDF)"]:
                for ruta in pdf_generados:
                    zipf.write(ruta, os.path.basename(ruta))

        zip_buffer.seek(0)

        st.session_state.procesado  = True
        st.session_state.zip_buffer = zip_buffer
        st.session_state.contador   = contador
        st.session_state.pdf_count  = len(pdf_generados)
        st.session_state.formato    = formato
        st.rerun()

# ===============================
# RESULTADO Y DESCARGA
# ===============================
if st.session_state.procesado and st.session_state.zip_buffer:

    fmt = st.session_state.get("formato", "Word (.docx)")

    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-box">
            <div class="metric-num">{st.session_state.contador}</div>
            <div class="metric-label">Certificados Word</div>
        </div>
        <div class="metric-box">
            <div class="metric-num">{st.session_state.pdf_count if fmt != "Word (.docx)" else "—"}</div>
            <div class="metric-label">Certificados PDF</div>
        </div>
        <div class="metric-box">
            <div class="metric-num">✓</div>
            <div class="metric-label">ZIP listo para descarga</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        label="📦 Descargar certificados (.zip)",
        data=st.session_state.zip_buffer,
        file_name="certificados_mapfre.zip",
        mime="application/zip",
        use_container_width=True
    )

    st.markdown('<div class="red-divider"></div>', unsafe_allow_html=True)

    if st.button("🧹 Limpiar y procesar nuevos certificados", use_container_width=True):
        st.session_state.procesado  = False
        st.session_state.zip_buffer = None
        st.session_state.contador   = 0
        st.session_state.pdf_count  = 0
        shutil.rmtree("certificados", ignore_errors=True)
        if os.path.exists("plantilla_temp.docx"):
            os.remove("plantilla_temp.docx")
        st.rerun()

# ===============================
# FOOTER
# ===============================
st.markdown("""
<div class="footer-bar">
    MAPFRE Seguros Perú · Emisión y Renovación Masivos · Uso interno
</div>
""", unsafe_allow_html=True)
