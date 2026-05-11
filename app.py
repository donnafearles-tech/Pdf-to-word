# streamlit_app.py
import os
import json
import tempfile
import streamlit as st
import docx
from docx.oxml.ns import qn

# Adobe SDK
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.jobs.create_pdf_job import CreatePDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
from adobe.pdfservices.operation.pdfjobs.result.create_pdf_result import CreatePDFResult

# ------------------------------------------------------------
# Función para obtener credenciales (desde st.secrets en Cloud)
# ------------------------------------------------------------
def get_adobe_creds():
    # En local puedes leerlos de un archivo, pero en Cloud leemos de st.secrets
    # Por compatibilidad dejamos también la opción de archivo si existe,
    # pero la forma recomendada es usar los secrets de Streamlit.
    try:
        client_id = st.secrets["PDF_SERVICES_CLIENT_ID"]
        client_secret = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
        return client_id, client_secret
    except Exception:
        # Fallback: si no hay secrets, intenta cargar 'adobesecret.json' (solo desarrollo)
        with open('adobesecret.json', 'r') as f:
            data = json.load(f)
            creds = data["project"]["workspace"]["details"]["credentials"][0]["oauth_server_to_server"]
            return creds["client_id"], creds["client_secrets"][0]

# ----------------------------------------------
# Determinar tipo MIME según extensión
# ----------------------------------------------
def get_media_type(filename):
    ext = filename.lower().split('.')[-1]
    if ext == 'png':
        return PDFServicesMediaType.PNG
    if ext in ['jpg', 'jpeg']:
        return PDFServicesMediaType.JPEG
    return PDFServicesMediaType.PDF

# ----------------------------------------------
# Limpieza de fondos y cajas (tu código original)
# ----------------------------------------------
def clean_word_backgrounds(docx_path):
    try:
        doc = docx.Document(docx_path)

        # 1. Quitar marca de fondo global
        settings = doc.settings.element
        display_bg = settings.find(qn('w:displayBackgroundShape'))
        if display_bg is not None:
            settings.remove(display_bg)

        # 2. Partes a limpiar (cuerpo + encabezados + pies)
        parts = [doc.element]
        for section in doc.sections:
            parts.append(section.header._element)
            parts.append(section.footer._element)

        for part in parts:
            # A. Sombreados
            for shd in part.xpath('.//*[local-name()="shd"]'):
                shd.getparent().remove(shd)
            # B. Bordes de párrafo / página
            for bdr in part.xpath('.//*[local-name()="pBdr"] | .//*[local-name()="pgBdr"]'):
                bdr.getparent().remove(bdr)
            # C. Fondo
            for bg in part.xpath('.//*[local-name()="background"]'):
                bg.getparent().remove(bg)
            # D. Formas vectoriales vacías
            search_query = './/*[local-name()="shape"] | .//*[local-name()="rect"] | .//*[local-name()="roundrect"] | .//*[local-name()="pict"] | .//*[local-name()="drawing"] | .//*[local-name()="wsp"]'
            for shape in part.xpath(search_query):
                has_image = len(shape.xpath('.//*[local-name()="pic"] | .//*[local-name()="imagedata"]')) > 0
                has_text = len(shape.xpath('.//*[local-name()="t"]')) > 0
                if not has_image and not has_text and shape.getparent() is not None:
                    shape.getparent().remove(shape)

        doc.save(docx_path)
    except Exception as e:
        st.warning(f"La limpieza de fondos no se pudo completar: {e}")
        # No detenemos todo el proceso; el archivo se entrega igual

# ----------------------------------------------
# Página principal de Streamlit
# ----------------------------------------------
st.set_page_config(page_title="Conversor a Word con Adobe", page_icon="📄")
st.title("🖼️ Conversor de PDF / Imágenes a Word")
st.markdown("Sube un archivo PDF, JPG o PNG y obtén un documento Word limpio, aprovechando la API de Adobe.")

uploaded_file = st.file_uploader("Selecciona tu archivo", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file is not None:
    # Guardar archivo temporalmente
    with tempfile.NamedTemporaryFile(delete=False, suffix="_" + uploaded_file.name) as tmp:
        tmp.write(uploaded_file.getbuffer())
        input_path = tmp.name

    st.info("Archivo recibido. Presiona el botón para convertir.")
    if st.button("Convertir a Word"):
        try:
            media_type = get_media_type(uploaded_file.name)
            client_id, client_secret = get_adobe_creds()
            credentials = ServicePrincipalCredentials(client_id=client_id, client_secret=client_secret)
            pdf_services = PDFServices(credentials=credentials)

            with st.spinner("Subiendo archivo a Adobe..."):
                with open(input_path, 'rb') as f:
                    input_stream = f.read()
                doc_asset = pdf_services.upload(input_stream=input_stream, mime_type=media_type)

            # Si no es PDF, lo convierte primero a PDF
            if media_type != PDFServicesMediaType.PDF:
                with st.spinner("Creando PDF desde imagen..."):
                    create_pdf_job = CreatePDFJob(doc_asset)
                    location = pdf_services.submit(create_pdf_job)
                    pdf_response = pdf_services.get_job_result(location, CreatePDFResult)
                    doc_asset = pdf_response.get_result().get_asset()

            # Exportar a Word
            with st.spinner("Exportando a Word con Adobe..."):
                export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                export_job = ExportPDFJob(doc_asset, export_params)
                location = pdf_services.submit(export_job)
                word_response = pdf_services.get_job_result(location, ExportPDFResult)

                result_asset = word_response.get_result().get_asset()
                stream_asset = pdf_services.get_content(result_asset)
                word_bytes = stream_asset.get_input_stream()

            # Guardar en un archivo temporal para limpiar
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(word_bytes)
                word_path = tmp_docx.name

            with st.spinner("Eliminando fondos y cajas innecesarias..."):
                clean_word_backgrounds(word_path)

            # Leer el archivo limpiado para ofrecer descarga
            with open(word_path, "rb") as f:
                cleaned_bytes = f.read()

            st.success("¡Conversión completada!")
            output_filename = uploaded_file.name.rsplit('.', 1)[0] + ".docx"
            st.download_button(
                label="📥 Descargar documento Word",
                data=cleaned_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # Limpiar temporales (en Streamlit Cloud son efímeros, pero nos aseguramos)
            try:
                os.unlink(input_path)
                os.unlink(word_path)
            except Exception:
                pass

        except Exception as e:
            st.error(f"Error durante el proceso: {e}")
            st.exception(e)  # Muestra traza completa en la app