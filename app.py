
import streamlit as st
import pandas as pd
import os
import zipfile
import re
import uuid
import time
import base64
import requests
from docxtpl import DocxTemplate
from docx import Document
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN iLovePDF
# Para Streamlit Cloud: usa st.secrets
# ─────────────────────────────────────────────
ILOVEPDF_PUBLIC_KEY = st.secrets["ILOVEPDF_PUBLIC_KEY"]
ILOVEPDF_SECRET_KEY = st.secrets["ILOVEPDF_SECRET_KEY"]


def convertir_multiples_docx_a_pdf_ilovepdf(docx_paths: list, output_dir: str) -> list:
    """
    Convierte una LISTA de archivos DOCX a PDF en una sola petición (Lote).
    Retorna una lista con las rutas de los PDFs generados con éxito.
    """
    if not docx_paths:
        return []

    pdf_generados = []

    try:
        # 1. Autenticar y obtener token
        auth_resp = requests.post(
            "https://api.ilovepdf.com/v1/auth",
            json={"public_key": ILOVEPDF_PUBLIC_KEY}
        )
        auth_resp.raise_for_status()
        token = auth_resp.json()["token"]
        headers = {"Authorization": f"Bearer {token}"}

        # 2. Iniciar una ÚNICA tarea para todo este lote
        start_resp = requests.get(
            "https://api.ilovepdf.com/v1/start/officepdf",
            headers=headers
        )
        start_resp.raise_for_status()
        start_data = start_resp.json()
        server = start_data["server"]
        task = start_data["task"]

        archivos_procesar = []

        # 3. Subir TODOS los archivos del lote actual a la misma tarea
        for docx_path in docx_paths:
            with open(docx_path, "rb") as f:
                upload_resp = requests.post(
                    f"https://{server}/v1/upload",
                    headers=headers,
                    data={"task": task},
                    files={"file": (os.path.basename(docx_path), f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                )
            upload_resp.raise_for_status()
            server_filename = upload_resp.json()["server_filename"]
            
            archivos_procesar.append({
                "server_filename": server_filename, 
                "filename": os.path.basename(docx_path)
            })

        # 4. Procesar la conversión de todo el lote junto
        process_resp = requests.post(
            f"https://{server}/v1/process",
            headers=headers,
            json={
                "task": task,
                "tool": "officepdf",
                "files": archivos_procesar
            }
        )
        process_resp.raise_for_status()

        # 5. Descargar el resultado (iLovePDF devuelve un .zip con los PDFs del lote)
        download_resp = requests.get(
            f"https://{server}/v1/download/{task}",
            headers=headers
        )
        download_resp.raise_for_status()

        # 6. Descomprimir los PDFs en la carpeta de destino
        zip_temporal = os.path.join(output_dir, f"resultado_{uuid.uuid4().hex}.zip")
        with open(zip_temporal, "wb") as f:
            f.write(download_resp.content)

        with zipfile.ZipFile(zip_temporal, "r") as z:
            z.extractall(output_dir)
            for nombre_archivo in z.namelist():
                pdf_path = os.path.join(output_dir, nombre_archivo)
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    pdf_generados.append(pdf_path)

        if os.path.exists(zip_temporal):
            os.remove(zip_temporal)

        return pdf_generados

    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Error en el lote de iLovePDF API: {e}")


def convertir_individual_para_preview(docx_path: str, output_dir: str) -> str | None:
    """
    Función auxiliar rápida para convertir un solo archivo (usado solo en la Previsualización).
    """
    try:
        lista_resultado = convertir_multiples_docx_a_pdf_ilovepdf([docx_path], output_dir)
        return lista_resultado[0] if lista_resultado else None
    except Exception as e:
        raise RuntimeError(str(e))


# ─────────────────────────────────────────────
# APP STREAMLIT
# ─────────────────────────────────────────────

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
                with st.spinner("Convirtiendo a PDF con iLovePDF..."):
                    try:
                        preview_pdf = convertir_individual_para_preview(preview_docx, preview_dir)

                        if preview_pdf and os.path.exists(preview_pdf):
                            with open(preview_pdf, "rb") as f:
                                st.download_button(
                                    "📥 Descargar previsualización PDF",
                                    f.read(),
                                    file_name=f"PREVIEW_{nombre}.pdf"
                                )
                        else:
                            st.error("❌ No se generó el PDF de previsualización.")

                    except RuntimeError as e:
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

        # Paso 1: generar todos los DOCX localmente (Sin tocar API)
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
            # Ajuste de barra de progreso proporcional
            progress.progress(contador / (total * 2 if formato in ["PDF", "Ambos"] else total))
            status.text(f"Generando DOCX {contador} de {total}...")

        # Paso 2: convertir a PDF optimizado en mini-lotes de máximo 10 archivos
        pdf_generados = []

        if formato in ["PDF", "Ambos"] and docx_generados:
            limite_lote = 10
            # Divide la lista completa en una lista de sublistas (de 10 en 10)
            mini_lotes = [docx_generados[i:i + limite_lote] for i in range(0, len(docx_generados), limite_lote)]
            total_lotes = len(mini_lotes)

            for index, lote in enumerate(mini_lotes):
                status.text(f"Convirtiendo lote de PDFs {index + 1} de {total_lotes}... ⏳")
                try:
                    # Se envía el grupo completo a la API
                    pdfs_del_lote = convertir_multiples_docx_a_pdf_ilovepdf(lote, pdf_dir)
                    pdf_generados.extend(pdfs_del_lote)
                except RuntimeError as e:
                    errores_proceso.append(f"❌ Error en el lote {index + 1}: {e}")

                # Avanza la barra dinámicamente en la segunda mitad del proceso
                progress.progress(0.5 + (0.4 * ((index + 1) / total_lotes)))

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
