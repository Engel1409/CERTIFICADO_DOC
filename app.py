import streamlit as st
import pandas as pd
import os
import zipfile
import subprocess
import shutil
import uuid
import time
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime

# ===============================
# CONFIG
# ===============================
st.set_page_config(page_title="Generador de Certificados", layout="wide")
st.title("📄 Generador de Certificados")

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
# UPLOAD
# ===============================
excel_file = st.file_uploader("📊 Subir Excel", type=["xlsx"])
docx_template = st.file_uploader("📄 Subir plantilla Word", type=["docx"])

# ===============================
# PROCESAMIENTO
# ===============================
if excel_file and docx_template and not st.session_state.procesado:

    df = pd.read_excel(excel_file, dtype=str)
    df = df.dropna(how="all").fillna("")
    df.columns = df.columns.str.strip().str.lower()

    st.write(f"Registros: {len(df)}")

    formato = st.radio(
        "Formato de salida",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"],
        horizontal=True
    )

    if st.button("⚙️ Procesar documentos"):

        # ===============================
        # CARPETA ÚNICA POR EJECUCIÓN
        # ===============================
        base_dir = f"work_{uuid.uuid4().hex}"
        docx_dir = os.path.join(base_dir, "docx")
        pdf_dir = os.path.join(base_dir, "pdf")

        os.makedirs(docx_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        # ===============================
        # TEMPLATE
        # ===============================
        template_path = os.path.join(base_dir, "plantilla.docx")
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        # ===============================
        # FECHA
        # ===============================
        hoy = datetime.now()
        fecha_texto = hoy.strftime("%d/%m/%Y")

        # ===============================
        # GENERAR DOCX
        # ===============================
        docx_generados = []

        for fila in df.to_dict("records"):

            fila = {k.strip().lower(): v for k, v in fila.items()}
            fila["fecha"] = fecha_texto

            doc = DocxTemplate(template_path)
            doc.render(fila)

            nombre_base = f"{fila.get('nro','')}_{fila.get('asegurado','')}_{fila.get('poliza','')}"
            nombre_base = nombre_base.replace("/", "_").replace("\\", "_")

            nombre = f"{nombre_base}_{uuid.uuid4().hex[:6]}.docx"
            ruta = os.path.join(docx_dir, nombre)

            doc.save(ruta)

            if os.path.exists(ruta) and os.path.getsize(ruta) > 0:
                docx_generados.append(ruta)
            else:
                st.warning(f"DOCX inválido: {nombre}")

        # ===============================
        # CONVERTIR PDF
        # ===============================
        pdf_generados = []

        if formato in ["PDF", "Ambos (Word + PDF)"]:

            try:
                subprocess.run(
                    [
                        "libreoffice",
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", pdf_dir,
                        *docx_generados
                    ],
                    check=True
                )

                time.sleep(1)  # 🔥 clave estabilidad

            except Exception as e:
                st.error(f"Error PDF: {e}")

            for docx_path in docx_generados:
                nombre_pdf = os.path.basename(docx_path).replace(".docx", ".pdf")
                pdf_path = os.path.join(pdf_dir, nombre_pdf)

                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    pdf_generados.append(pdf_path)
                else:
                    st.warning(f"PDF no generado: {nombre_pdf}")

        st.write("DOCX:", len(docx_generados))
        st.write("PDF:", len(pdf_generados))

        # ===============================
        # CREAR ZIP (FIX DEFINITIVO)
        # ===============================
        zip_buffer = BytesIO()
        files_added = 0

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zipf:

            if formato in ["Word (.docx)", "Ambos (Word + PDF)"]:
                for f in docx_generados:
                    if os.path.exists(f):
                        with open(f, "rb") as file_data:
                            zipf.writestr(os.path.basename(f), file_data.read())
                            files_added += 1

            if formato in ["PDF", "Ambos (Word + PDF)"]:
                for f in pdf_generados:
                    if os.path.exists(f):
                        with open(f, "rb") as file_data:
                            zipf.writestr(os.path.basename(f), file_data.read())
                            files_added += 1

        zip_buffer.seek(0)

        if files_added == 0:
            st.error("❌ No se generaron archivos")
            shutil.rmtree(base_dir, ignore_errors=True)
            st.stop()

        # ✅ 🔥 FIX CRÍTICO (ZIP SIEMPRE BIEN)
        st.session_state.zip_buffer = zip_buffer.getvalue()

        st.session_state.procesado = True
        st.session_state.contador = len(docx_generados)
        st.session_state.pdf_count = len(pdf_generados)

        # ✅ limpieza total
        shutil.rmtree(base_dir, ignore_errors=True)

        st.rerun()

# ===============================
# DESCARGA
# ===============================
if st.session_state.procesado and st.session_state.zip_buffer:

    st.success("✅ Archivos generados correctamente")

    st.write(f"DOCX: {st.session_state.contador}")
    st.write(f"PDF: {st.session_state.pdf_count}")

    st.download_button(
        label="📦 Descargar ZIP",
        data=st.session_state.zip_buffer,
        file_name="certificados.zip",
        mime="application/zip"
    )

    if st.button("🧹 Nuevo proceso"):
        st.session_state.procesado = False
        st.session_state.zip_buffer = None
        st.session_state.contador = 0
        st.session_state.pdf_count = 0
        st.rerun()
