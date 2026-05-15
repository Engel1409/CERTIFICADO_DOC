import streamlit as st
import pandas as pd
import os
import zipfile
import subprocess
import tempfile
from docxtpl import DocxTemplate
from io import BytesIO

st.set_page_config(page_title="Generador de Certificados", layout="wide")

st.title("📄 Generador de Certificados DOCX")

# ===============================
# SUBIDA DE ARCHIVOS
# ===============================
excel_file = st.file_uploader("📊 Subir Excel de pólizas", type=["xlsx"])
docx_template = st.file_uploader("📄 Subir plantilla DOCX", type=["docx"])

if excel_file and docx_template:

    # ===============================
    # LEER EXCEL
    # ===============================
    df = pd.read_excel(excel_file, dtype=str)
    df.columns = df.columns.str.strip().str.lower()

    st.subheader("📊 Vista previa (primeras 10 pólizas)")
    st.dataframe(df.head(10))

    total = len(df)
    st.success(f"✅ Total de pólizas: {total}")

    # ===============================
    # SELECTOR DE FORMATO
    # ===============================
    formato = st.radio(
        "📁 Formato de descarga",
        ["Word (.docx)", "PDF", "Ambos (Word + PDF)"],
        horizontal=True
    )

    # ===============================
    # BOTÓN PROCESAR
    # ===============================
    if st.button("⚙️ Procesar certificados"):

        carpeta = "certificados"
        os.makedirs(carpeta, exist_ok=True)

        progreso = st.progress(0)
        contador = 0

        # Guardar plantilla temporalmente
        template_path = "plantilla_temp.docx"
        with open(template_path, "wb") as f:
            f.write(docx_template.read())

        # Generar DOCX por cada fila
        docx_generados = []
        for i, fila in df.iterrows():
            doc = DocxTemplate(template_path)
            contexto = fila.to_dict()
            doc.render(contexto)

            nombre = f"{fila.get('poliza', i)}.docx"
            ruta = os.path.join(carpeta, nombre)
            doc.save(ruta)
            docx_generados.append(ruta)

            contador += 1
            progreso.progress((i + 1) / total)

        st.success(f"🎉 Se generaron {contador} certificados")

        # ===============================
        # CONVERSIÓN A PDF (si aplica)
        # ===============================
        pdf_generados = []
        if formato in ["PDF", "Ambos (Word + PDF)"]:
            with st.spinner("📄 Convirtiendo a PDF..."):
                output_dir = tempfile.gettempdir()
                for docx_path in docx_generados:
                    try:
                        subprocess.run(
                            [
                                "libreoffice", "--headless",
                                "--convert-to", "pdf",
                                "--outdir", output_dir,
                                docx_path
                            ],
                            check=True,
                            capture_output=True
                        )
                        pdf_nombre = os.path.basename(docx_path).replace(".docx", ".pdf")
                        pdf_path = os.path.join(output_dir, pdf_nombre)
                        if os.path.exists(pdf_path):
                            pdf_generados.append(pdf_path)
                    except subprocess.CalledProcessError as e:
                        st.warning(f"⚠️ No se pudo convertir: {os.path.basename(docx_path)}")

            st.success(f"✅ {len(pdf_generados)} PDFs generados")

        # ===============================
        # CREAR ZIP EN MEMORIA
        # ===============================
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            if formato in ["Word (.docx)", "Ambos (Word + PDF)"]:
                for ruta in docx_generados:
                    zipf.write(ruta, os.path.basename(ruta))
            if formato in ["PDF", "Ambos (Word + PDF)"]:
                for ruta in pdf_generados:
                    zipf.write(ruta, os.path.basename(ruta))

        zip_buffer.seek(0)

        # ===============================
        # DESCARGA
        # ===============================
        st.download_button(
            label="📦 Descargar ZIP",
            data=zip_buffer,
            file_name="certificados.zip",
            mime="application/zip"
        )
