import streamlit as st
import pandas as pd
import os
import zipfile
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

        for i, fila in df.iterrows():
            doc = DocxTemplate(template_path)

            contexto = fila.to_dict()
            doc.render(contexto)

            nombre = f"{fila.get('poliza', i)}.docx"
            ruta = os.path.join(carpeta, nombre)

            doc.save(ruta)

            contador += 1
            progreso.progress((i + 1) / total)

        st.success(f"🎉 Se generaron {contador} certificados")

        # ===============================
        # CREAR ZIP EN MEMORIA
        # ===============================
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for archivo in os.listdir(carpeta):
                ruta_archivo = os.path.join(carpeta, archivo)
                zipf.write(ruta_archivo, archivo)

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
