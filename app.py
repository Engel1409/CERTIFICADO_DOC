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
    letter-spacing: 1px;
}

.header-sub {
    font-size: 13px;
    color: rgba(255,255,255,0.8);
    margin-left: auto;
}

.card {
    background: white;
    border-radius: 6px;
    border-left: 4px solid var(--primary);
    padding: 24px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}

.card-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 16px;
    font-weight: 700;
    color: var(--primary);
    text-transform: uppercase;
    margin-bottom: 14px;
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
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 34px;
    font-weight: 700;
    color: var(--primary);
}

.metric-label {
    font-size: 12px;
    color: var(--mid);
    text-transform: uppercase;
}

.stButton > button,
.stDownloadButton > button {
    border-radius: 6px !important;
    border: none !important;
    font-weight: 600 !important;
}

.stButton > button {
    background-color: var(--primary) !important;
    color: white !important;
}

.stDownloadButton > button {
    background-color: var(--dark) !important;
    color: white !important;
    width: 100%;
    padding: 14px !important;
}

.footer-bar {
    text-align: center;
    font-size: 11px;
    color: var(--mid);
    margin-top: 40px;
    padding-top: 16px;
    border-top: 1px solid var(--border);
}
</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown("""
<div class="header-bar">
    <div class="logo-text">DOCFLOW</div>
    <div class="header-sub">
        Generación masiva de documentos
    </div>
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
st.markdown(
    '<div class="card"><div class="card-title">📁 Archivos de entrada</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader(
        "Archivo Excel (.xlsx)",
        type=["xlsx"]
    )

with col2:
    docx_template = st.file_uploader(
        "Plantilla Word (.docx)",
        type=["docx"]
    )

st.markdown('</div>', unsafe_allow_html=True)

# ===============================
# PROCESAMIENTO
# ===============================
if excel_file and docx_template and not st.session_state.procesado:

    # ===============================
    # LEER EXCEL
    # ===============================
    df = pd.read_excel(excel_file, dtype=str)

    # Limpiar NaN
    df = df.fillna("")

    # Convertir TODO a texto + MAYÚSCULAS
    for col in df.columns:
        df[col] = df[col].astype(str).str.upper()

    total = len(df)
    cols_count = len(df.columns)

    # ===============================
    # MÉTRICAS
    # ===============================
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

    # ===============================
    # PREVIEW
    # ===============================
    st.markdown(
        '<div class="card"><div class="card-title">📊 Vista previa</div>',
        unsafe_allow_html=True
    )

    st.dataframe(df.head(10), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ===============================
    # CONFIGURACIÓN
    # ===============================
    st.markdown(
        '<div class="card"><div class="card-title">⚙️ Configuración</div>',
        unsafe_allow_html=True
    )

    formato = st.radio(
        "Formato de descarga",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"],
        horizontal=True
    )

    st.markdown('</div>', unsafe_allow_html=True)

    # ===============================
    # BOTÓN PROCESAR
    # ===============================
    if st.button("⚙️ Procesar documentos", use_container_width=True):

        carpeta = "documentos_generados"
        os.makedirs(carpeta, exist_ok=True)

        progreso = st.progress(0)
        status = st.empty()

        contador = 0

        # Guardar plantilla temporal
        template_path = "plantilla_temp.docx"

        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        docx_generados = []

        # ===============================
        # GENERAR WORDS
        # ===============================
        for i, fila in enumerate(df.to_dict(orient="records")):

            try:
                doc = DocxTemplate(template_path)

                contexto = {
                    k: str(v).upper()
                    for k, v in fila.items()
                }

                doc.render(contexto)

                nombre_base = str(
                    fila.get("poliza", f"documento_{i}")
                )

                nombre_base = (
                    nombre_base
                    .replace("/", "_")
                    .replace("\\", "_")
                    .replace(":", "_")
                    .replace("*", "_")
                    .replace("?", "_")
                    .replace('"', "_")
                    .replace("<", "_")
                    .replace(">", "_")
                    .replace("|", "_")
                )

                nombre = f"{nombre_base}.docx"

                ruta = os.path.join(carpeta, nombre)

                doc.save(ruta)

                docx_generados.append(ruta)

                contador += 1

                progreso.progress((i + 1) / total)

                status.caption(
                    f"Generando {contador} de {total}..."
                )

            except Exception as e:
                st.warning(f"Error en fila {i + 1}: {e}")

        status.empty()

        # ===============================
        # CONVERTIR PDF
        # ===============================
        pdf_generados = []

        if formato in ["PDF", "Ambos (Word + PDF)"]:

            pdf_status = st.empty()

            pdf_status.info(
                "📄 Convirtiendo documentos a PDF..."
            )

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

                    pdf_nombre = os.path.basename(
                        docx_path
                    ).replace(".docx", ".pdf")

                    pdf_path = os.path.join(
                        output_dir,
                        pdf_nombre
                    )

                    if os.path.exists(pdf_path):
                        pdf_generados.append(pdf_path)

                except subprocess.CalledProcessError:
                    st.warning(
                        f"No se pudo convertir: "
                        f"{os.path.basename(docx_path)}"
                    )

            pdf_status.empty()

        # ===============================
        # CREAR ZIP
        # ===============================
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:

            if formato in [
                "Word (.docx)",
                "Ambos (Word + PDF)"
            ]:

                for ruta in docx_generados:
                    zipf.write(
                        ruta,
                        os.path.basename(ruta)
                    )

            if formato in [
                "PDF",
                "Ambos (Word + PDF)"
            ]:

                for ruta in pdf_generados:
                    zipf.write(
                        ruta,
                        os.path.basename(ruta)
                    )

        zip_buffer.seek(0)

        # ===============================
        # SESSION STATE
        # ===============================
        st.session_state.procesado = True
        st.session_state.zip_buffer = zip_buffer
        st.session_state.contador = contador
        st.session_state.pdf_count = len(pdf_generados)
        st.session_state.formato = formato

        st.rerun()

# ===============================
# RESULTADOS
# ===============================
if (
    st.session_state.procesado
    and st.session_state.zip_buffer
):

    fmt = st.session_state.get(
        "formato",
        "Word (.docx)"
    )

    st.markdown(f"""
    <div class="metric-row">

        <div class="metric-box">
            <div class="metric-num">
                {st.session_state.contador}
            </div>
            <div class="metric-label">
                Documentos Word
            </div>
        </div>

        <div class="metric-box">
            <div class="metric-num">
                {
                    st.session_state.pdf_count
                    if fmt != "Word (.docx)"
                    else "—"
                }
            </div>
            <div class="metric-label">
                Documentos PDF
            </div>
        </div>

        <div class="metric-box">
            <div class="metric-num">✓</div>
            <div class="metric-label">
                ZIP listo
            </div>
        </div>

    </div>
    """, unsafe_allow_html=True)

    # ===============================
    # DESCARGA
    # ===============================
    st.download_button(
        label="📦 Descargar ZIP",
        data=st.session_state.zip_buffer,
        file_name="documentos_generados.zip",
        mime="application/zip",
        use_container_width=True
    )

    # ===============================
    # LIMPIAR
    # ===============================
    if st.button(
        "🧹 Procesar nuevos documentos",
        use_container_width=True
    ):

        st.session_state.procesado = False
        st.session_state.zip_buffer = None
        st.session_state.contador = 0
        st.session_state.pdf_count = 0

        shutil.rmtree(
            "documentos_generados",
            ignore_errors=True
        )

        if os.path.exists("plantilla_temp.docx"):
            os.remove("plantilla_temp.docx")

        st.rerun()

# ===============================
# FOOTER
# ===============================
st.markdown("""
<div class="footer-bar">
    Sistema de automatización documental
</div>
""", unsafe_allow_html=True)
