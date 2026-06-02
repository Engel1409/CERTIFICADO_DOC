import streamlit as st
import pandas as pd
import os
import zipfile
import subprocess
import re
import uuid
import time
import base64
from docxtpl import DocxTemplate
from docx import Document
from datetime import datetime

st.set_page_config(page_title="Generador", layout="wide")
st.title("📄 Generador de Certificados")

excel_file = st.file_uploader("📊 Excel", type=["xlsx"])
docx_template = st.file_uploader("📄 Word", type=["docx"])

if excel_file and docx_template:

    if not excel_file.name.endswith(".xlsx"):
        st.error("❌ El Excel debe ser .xlsx — convierte el archivo en Excel: Archivo → Guardar como → .xlsx")
        st.stop()

    if not docx_template.name.endswith(".docx"):
        st.error("❌ El Word debe ser .docx — convierte el archivo en Word: Archivo → Guardar como → .docx")
        st.stop()

    try:
        df = pd.read_excel(excel_file, dtype=str).fillna("")
        df.columns = df.columns.str.strip().str.lower()
    except Exception as e:
        st.error(f"❌ No se pudo leer el Excel: {e}")
        st.stop()

    try:
        doc_diag = Document(docx_template)
        patron = re.compile(r"{{(.*?)}}")
        tags_word = set()

        for para in doc_diag.paragraphs:
            for tag in patron.findall(para.text):
                tags_word.add(tag.strip().lower())

        for table in doc_diag.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for tag in patron.findall(para.text):
                            tags_word.add(tag.strip().lower())

    except Exception as e:
        st.error(f"❌ No se pudo leer la plantilla Word: {e}")
        st.stop()

    TAGS_IGNORADOS = {"fecha"}

    columnas_excel = set(df.columns.tolist())

    st.markdown("---")
    st.subheader("🔎 Diagnóstico de coincidencias")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**🏷️ Tags del Word vs Excel**")
        filas_diag = []
        for tag in sorted(tags_word - TAGS_IGNORADOS):
            estado = "✅ OK" if tag in columnas_excel else "❌ No encontrado en Excel"
            filas_diag.append({"Tag en Word": f"{{{{{tag}}}}}", "En Excel": estado})
        st.dataframe(pd.DataFrame(filas_diag), use_container_width=True, hide_index=True)

    with col2:
        st.markdown("**📋 Columnas del Excel sin tag en Word**")
        sobrantes = sorted(columnas_excel - (tags_word - TAGS_IGNORADOS))
        if sobrantes:
            filas_sob = [{"Columna Excel": c, "Estado": "⚠️ Sin tag en Word"} for c in sobrantes]
            st.dataframe(pd.DataFrame(filas_sob), use_container_width=True, hide_index=True)
        else:
            st.success("✅ Todas las columnas del Excel tienen tag en el Word")

    tags_faltantes = [t for t in (tags_word - TAGS_IGNORADOS) if t not in columnas_excel]
    if tags_faltantes:
        st.error(f"❌ {len(tags_faltantes)} tag(s) del Word no encontrados en el Excel — revisa antes de procesar.")
    else:
        st.success("✅ Todos los tags del Word coinciden con el Excel")

    st.markdown("---")

    total = len(df)
    st.info(f"📋 {total} fila(s) detectadas — se generarán {total} certificado(s)")

    formato = st.radio("Formato", ["Word (.docx)", "PDF", "Ambos"], horizontal=True)

    base_dir = f"work_{uuid.uuid4().hex}"
    os.makedirs(base_dir, exist_ok=True)

    template_path = os.path.join(base_dir, "plantilla.docx")
    docx_template.seek(0)
    with open(template_path, "wb") as f:
        f.write(docx_template.read())

    fecha = datetime.now().strftime("%d/%m/%Y")

    st.subheader("🔍 Previsualizar primer certificado")

    if st.button("📄 Generar previsualización"):

        preview_dir = os.path.join(base_dir, "preview")
        os.makedirs(preview_dir, exist_ok=True)

        try:
            fila = df.iloc[0].to_dict()
            fila = {k.lower(): v for k, v in fila.items()}
            fila["fecha"] = fecha

            nro = fila.get('nro', '').strip()
            asegurado = fila.get('asegurado', '').strip()
            poliza = fila.get('poliza', '').strip()
            nombre = f"{nro}_{asegurado}_{poliza}" if nro else f"{asegurado}_{poliza}"
            nombre = nombre.replace("/", "_")

            doc = DocxTemplate(template_path)
            doc.render(fila)

            preview_docx = os.path.join(preview_dir, f"PREVIEW_{nombre}.docx")
            doc.save(preview_docx)

            if formato in ["PDF", "Ambos"]:
                try:
                    subprocess.run(
                        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", preview_dir, preview_docx],
                        check=True,
                        timeout=30
                    )
                    time.sleep(1)

                    preview_pdf = preview_docx.replace(".docx", ".pdf")

                    if os.path.exists(preview_pdf) and os.path.getsize(preview_pdf) > 0:
                        with open(preview_pdf, "rb") as f:
                            st.download_button(
                                "📥 Descargar previsualización PDF",
                                f.read(),
                                file_name=f"PREVIEW_{nombre}.pdf"
                            )
                    else:
                        st.error("❌ No se generó el PDF de previsualización.")

                except subprocess.TimeoutExpired:
                    st.error("❌ LibreOffice tardó demasiado al convertir la previsualización.")
                except subprocess.CalledProcessError as e:
                    st.error(f"❌ Error al convertir a PDF: {e}")

            if formato in ["Word (.docx)", "Ambos"]:
                with open(preview_docx, "rb") as f:
                    st.download_button(
                        "📥 Descargar previsualización Word",
                        f.read(),
                        file_name=f"PREVIEW_{nombre}.docx"
                    )

        except KeyError as e:
            st.error(f"❌ Tag no encontrado en el Excel: {e} — verifica que los {{{{tags}}}} del Word coincidan con las cabeceras del Excel.")
        except Exception as e:
            st.error(f"❌ Error al generar la previsualización: {e}")

    st.markdown("---")

    if st.button("⚙️ Procesar todos"):

        errores_proceso = []
        progress = st.progress(0)
        status = st.empty()
        contador = 0

        docx_dir = os.path.join(base_dir, "docx")
        pdf_dir = os.path.join(base_dir, "pdf")
        os.makedirs(docx_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        docx_generados = []

        for idx, fila in enumerate(df.to_dict("records")):

            fila = {k.lower(): v for k, v in fila.items()}
            fila["fecha"] = fecha

            nro = fila.get('nro', '').strip()
            asegurado = fila.get('asegurado', '').strip()
            poliza = fila.get('poliza', '').strip()
            nombre = f"{nro}_{asegurado}_{poliza}" if nro else f"{asegurado}_{poliza}"
            nombre = nombre.replace("/", "_")

            ruta = os.path.join(docx_dir, f"{nombre}.docx")

            try:
                doc = DocxTemplate(template_path)
                doc.render(fila)
                doc.save(ruta)

                if os.path.exists(ruta):
                    docx_generados.append(ruta)

            except KeyError as e:
                errores_proceso.append(f"⚠️ Fila {idx + 1} — tag no encontrado: {e}")
            except Exception as e:
                errores_proceso.append(f"⚠️ Fila {idx + 1} — error al generar DOCX: {e}")

            contador += 1
            progress.progress(contador / total)
            status.text(f"Generando DOCX {contador} de {total}...")

        pdf_generados = []

        if formato in ["PDF", "Ambos"] and docx_generados:

            status.text("Convirtiendo a PDF... ⏳")

            try:
                subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, *docx_generados],
                    check=True,
                    timeout=300
                )
                time.sleep(2)

                for f in docx_generados:
                    pdf = os.path.join(pdf_dir, os.path.basename(f).replace(".docx", ".pdf"))
                    if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
                        pdf_generados.append(pdf)
                    else:
                        errores_proceso.append(f"⚠️ No se generó PDF para: {os.path.basename(f)}")

            except subprocess.TimeoutExpired:
                errores_proceso.append("❌ LibreOffice tardó demasiado en la conversión a PDF.")
            except subprocess.CalledProcessError as e:
                errores_proceso.append(f"❌ Error en conversión a PDF: {e}")

        status.text("Empaquetando archivos... 📦")

        nombre_zip = f"certificados_{datetime.now().strftime('%d%m%Y_%H%M')}.zip"
        zip_path = os.path.join(base_dir, nombre_zip)

        try:
            with zipfile.ZipFile(zip_path, "w") as z:
                if formato in ["Word (.docx)", "Ambos"]:
                    for f in docx_generados:
                        z.write(f, os.path.basename(f))
                if formato in ["PDF", "Ambos"]:
                    for f in pdf_generados:
                        z.write(f, os.path.basename(f))

            with open(zip_path, "rb") as f:
                data = f.read()

            b64 = base64.b64encode(data).decode()
            href = f'<a download="{nombre_zip}" href="data:application/zip;base64,{b64}">📦 Descargar {nombre_zip}</a>'
            st.markdown(href, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ Error al crear el ZIP: {e}")

        if errores_proceso:
            st.markdown("---")
            st.subheader("⚠️ Advertencias del proceso")
            for err in errores_proceso:
                if err.startswith("❌"):
                    st.error(err)
                else:
                    st.warning(err)

        generados = len(docx_generados)
        fallidos = total - generados

        progress.progress(1.0)
        status.text("✅ Proceso completado")
        st.success(f"✅ {generados} certificado(s) generado(s)" + (f" — ⚠️ {fallidos} con error" if fallidos > 0 else ""))
