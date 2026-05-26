import streamlit as st
import pandas as pd
import os
import zipfile
import subprocess
import shutil
import uuid
import time
import base64
from docxtpl import DocxTemplate
from datetime import datetime

st.set_page_config(page_title="Generador", layout="wide")
st.title("📄 Generador de Certificados")

excel_file = st.file_uploader("📊 Excel", type=["xlsx"])
docx_template = st.file_uploader("📄 Word", type=["docx"])

if excel_file and docx_template:

    df = pd.read_excel(excel_file, dtype=str).fillna("")
    df.columns = df.columns.str.strip().str.lower()

    formato = st.radio(
        "Formato",
        ["Word (.docx)", "PDF", "Ambos"],
        horizontal=True
    )

    if st.button("⚙️ Procesar"):

        base_dir = f"work_{uuid.uuid4().hex}"
        os.makedirs(base_dir, exist_ok=True)

        docx_dir = os.path.join(base_dir, "docx")
        pdf_dir = os.path.join(base_dir, "pdf")
        os.makedirs(docx_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        # guardar plantilla
        template_path = os.path.join(base_dir, "plantilla.docx")
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        fecha = datetime.now().strftime("%d/%m/%Y")

        docx_generados = []

        # =========================
        # DOCX
        # =========================
        for fila in df.to_dict("records"):

            fila = {k.lower(): v for k, v in fila.items()}
            fila["fecha"] = fecha

            doc = DocxTemplate(template_path)
            doc.render(fila)

            nombre = f"{fila.get('nro','')}_{fila.get('asegurado','')}_{fila.get('poliza','')}"
            nombre = nombre.replace("/", "_")

            ruta = os.path.join(docx_dir, f"{nombre}.docx")
            doc.save(ruta)

            if os.path.exists(ruta):
                docx_generados.append(ruta)

        # =========================
        # PDF
        # =========================
        pdf_generados = []

        if formato in ["PDF", "Ambos"] and docx_generados:

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

            time.sleep(2)

            for f in docx_generados:
                pdf = os.path.join(pdf_dir, os.path.basename(f).replace(".docx", ".pdf"))
                if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
                    pdf_generados.append(pdf)

        # =========================
        # CREAR ZIP
        # =========================
        zip_path = os.path.join(base_dir, "archivos.zip")

        with zipfile.ZipFile(zip_path, "w") as z:

            if formato in ["Word (.docx)", "Ambos"]:
                for f in docx_generados:
                    z.write(f, os.path.basename(f))

            if formato in ["PDF", "Ambos"]:
                for f in pdf_generados:
                    z.write(f, os.path.basename(f))

        # =========================
        # 🔥 FIX FINAL (BASE64)
        # =========================
        with open(zip_path, "rb") as f:
            data = f.read()

        b64 = base64.b64encode(data).decode()

        href = f"""
        <a download="certificados.zip"
           href="data:application/zip;base64,{b64}">
           📦 Descargar ZIP (FUNCIONA SIEMPRE)
        </a>
        """

        st.markdown(href, unsafe_allow_html=True)

        st.success("✅ Listo")

