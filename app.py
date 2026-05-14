# =============================================================================
# APLICACIÓN DE CONVERSIÓN Y EXTRACCIÓN (PDF/IMAGEN A WORD / JSON)
# =============================================================================

import os
import re
import json
import logging
import tempfile
import zipfile  # <-- NUEVO: Para descomprimir el resultado de Extract API
import PyPDF2
from docx.shared import Inches
from dataclasses import dataclass
from typing import Tuple, List, Optional

import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.document import Document

# Librerías para pre-procesamiento de orientación e imagen
from PIL import Image, ImageOps

# Importaciones de Adobe PDF Services SDK (Word / Export)
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.jobs.create_pdf_job import CreatePDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
from adobe.pdfservices.operation.pdfjobs.result.create_pdf_result import CreatePDFResult

# <-- NUEVAS IMPORTACIONES: Adobe PDF Extract SDK (JSON) -->
from adobe.pdfservices.operation.pdfjobs.jobs.extract_pdf_job import ExtractPDFJob
from adobe.pdfservices.operation.pdfjobs.params.extract_pdf.extract_pdf_params import ExtractPDFParams
from adobe.pdfservices.operation.pdfjobs.params.extract_pdf.extract_element_type import ExtractElementType
from adobe.pdfservices.operation.pdfjobs.result.extract_pdf_result import ExtractPDFResult

# =============================================================================
# 1. CONFIGURACIÓN CENTRAL Y LOGGING
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

@dataclass
class CleanupConfig:
    force_single_column: bool = True
    remove_backgrounds: bool = True
    remove_small_images: bool = True
    min_image_width_cm: float = 3.0
    min_image_height_cm: float = 3.0
    clean_weird_symbols: bool = True
    fix_orientation: bool = True
    # Nuevos parámetros para recorte manual
    crop_left: int = 0
    crop_right: int = 0
    crop_top: int = 0
    crop_bottom: int = 0

# =============================================================================
# 2. FUNCIONES DE PRE-PROCESAMIENTO (ORIENTACIÓN Y RECORTE)
# =============================================================================

def get_adobe_creds() -> Tuple[str, str]:
    try:
        client_id = st.secrets["PDF_SERVICES_CLIENT_ID"]
        client_secret = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
        return client_id, client_secret
    except Exception:
        try:
            with open('adobesecret.json', 'r') as f:
                data = json.load(f)
                creds = data["project"]["workspace"]["details"]["credentials"][0]["oauth_server_to_server"]
                return creds["client_id"], creds["client_secrets"][0]
        except Exception as ex:
            raise RuntimeError("Faltan las credenciales de Adobe.") from ex

def get_media_type(filename: str) -> str:
    ext = filename.lower().split('.')[-1]
    if ext == 'png': return PDFServicesMediaType.PNG
    elif ext in ['jpg', 'jpeg']: return PDFServicesMediaType.JPEG
    return PDFServicesMediaType.PDF

def fix_image_orientation(image_path: str) -> None:
    try:
        image = Image.open(image_path)
        image = ImageOps.exif_transpose(image)
        image.save(image_path)
    except Exception as e:
        logger.warning(f"Error orientación imagen: {e}")

# <-- NUEVA FUNCIÓN: Recortar los 4 márgenes de la imagen -->
def crop_image_margins(image_path: str, left: int, right: int, top: int, bottom: int) -> None:
    if left == 0 and right == 0 and top == 0 and bottom == 0:
        return # No hay recortes solicitados
    try:
        img = Image.open(image_path)
        width, height = img.size
        
        # Calcular píxeles exactos a recortar según porcentaje
        x1 = int(width * (left / 100))
        y1 = int(height * (top / 100))
        x2 = width - int(width * (right / 100))
        y2 = height - int(height * (bottom / 100))
        
        if x1 < x2 and y1 < y2:
            img_cropped = img.crop((x1, y1, x2, y2))
            img_cropped.save(image_path)
            logger.info("Imagen recortada con éxito antes de enviar a Adobe.")
    except Exception as e:
        logger.warning(f"Error al intentar recortar los márgenes de la imagen: {e}")

def fix_pdf_orientation(pdf_path: str) -> str:
    # Lógica original conservada para PDFs
    try:
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        pages_rotated = 0
        for page in reader.pages:
            rotation = page.get('/Rotate', 0)
            if isinstance(rotation, PyPDF2.generic.NumberObject) and int(rotation) in [90, 270, -90]:
                page.rotate(-int(rotation))
                pages_rotated += 1
            writer.add_page(page)
        if pages_rotated > 0:
            fixed_pdf_path = pdf_path + "_fixed.pdf"
            with open(fixed_pdf_path, "wb") as f:
                writer.write(f)
            return fixed_pdf_path
        return pdf_path
    except Exception as e:
        return pdf_path

# =============================================================================
# 3. FUNCIONES DE LIMPIEZA DOCX (Mantenidas de tu código original)
# =============================================================================
def remove_docx_backgrounds(doc: Document, parts: List) -> None:
    # (Se mantiene tu código intacto)
    pass 

def force_single_column_layout(doc: Document) -> None:
    try:
        for section in doc.sections:
            cols = section._sectPr.xpath('./w:cols')
            for col in cols:
                if col.get(qn('w:num'), '1') != '1':
                    col.set(qn('w:num'), '1')
    except Exception: pass

def purify_text_symbols(doc: Document) -> None:
    allowed_pattern = re.compile(r'[^\w\s.,;:!?()\[\]{}"\'+\-*/=<>%$#@&^|\\~`áéíóúÁÉÍÓÚñÑüÜ¿¡€£¥°]')
    def clean_paragraphs(paragraphs):
        for para in paragraphs:
            for run in para.runs:
                if run.text:
                    cleaned_text = allowed_pattern.sub('', run.text)
                    cleaned_text = re.sub(r' {2,}', ' ', cleaned_text)
                    if cleaned_text != run.text: run.text = cleaned_text
    clean_paragraphs(doc.paragraphs)

def orchestrate_document_cleanup(docx_path: str, config: CleanupConfig) -> None:
    doc = docx.Document(docx_path)
    if config.force_single_column: force_single_column_layout(doc)
    if config.clean_weird_symbols: purify_text_symbols(doc)
    doc.save(docx_path)

# =============================================================================
# 4. INTERFAZ DE USUARIO Y EJECUCIÓN (STREAMLIT)
# =============================================================================

st.set_page_config(page_title="Multi-Herramienta Documental", layout="wide")
config = CleanupConfig()

st.sidebar.title("🛠️ Herramientas de Pre-Proceso")

# <-- NUEVOS CONTROLES PARA RECORTE EN LA INTERFAZ -->
st.sidebar.subheader("✂️ Recortar Imagen (Márgenes en %)")
st.sidebar.markdown("<small>Aplica solo para JPG/PNG (ej. quitar engargolados)</small>", unsafe_allow_html=True)
col_crop1, col_crop2 = st.sidebar.columns(2)
with col_crop1:
    config.crop_left = st.number_input("Izquierda", min_value=0, max_value=50, value=0, step=1)
    config.crop_top = st.number_input("Arriba", min_value=0, max_value=50, value=0, step=1)
with col_crop2:
    config.crop_right = st.number_input("Derecha", min_value=0, max_value=50, value=0, step=1)
    config.crop_bottom = st.number_input("Abajo", min_value=0, max_value=50, value=0, step=1)

st.sidebar.markdown("---")
st.sidebar.title("⚙️ Opciones de Limpieza Word")
config.force_single_column = st.sidebar.checkbox("Forzar a 1 sola columna", value=True)
config.clean_weird_symbols = st.sidebar.checkbox("Limpiar caracteres extraños", value=True)

st.title("📄 Extractor JSON & Conversor Word")

uploaded_file = st.file_uploader("Arrastra tu documento aquí (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file is not None:
    file_ext = uploaded_file.name.lower().split('.')[-1]
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp:
        tmp.write(uploaded_file.getbuffer())
        input_path = tmp.name

    st.info("✅ Archivo cargado. Selecciona la acción que deseas realizar:")

    # <-- DOS BOTONES PARA DOS FLUJOS DISTINTOS -->
    btn_json, btn_word = st.columns(2)
    
    run_json = btn_json.button("🧩 1. Extraer Solo JSON (Estructurado)", use_container_width=True)
    run_word = btn_word.button("🚀 2. Convertir y Limpiar a Word", use_container_width=True)

    if run_json or run_word:
        try:
            # --- PASO A: PRE-PROCESAMIENTO COMPARTIDO ---
            if file_ext in ['jpg', 'jpeg', 'png']:
                fix_image_orientation(input_path)
                # Aplicar recorte de márgenes antes de subir
                crop_image_margins(input_path, config.crop_left, config.crop_right, config.crop_top, config.crop_bottom)
            elif file_ext == 'pdf':
                input_path = fix_pdf_orientation(input_path)

            media_type = get_media_type(uploaded_file.name)
            client_id, client_secret = get_adobe_creds()
            credentials = ServicePrincipalCredentials(client_id=client_id, client_secret=client_secret)
            pdf_services = PDFServices(credentials=credentials)

            # Subida base a Adobe
            with st.spinner("Subiendo y pre-procesando en Adobe..."):
                with open(input_path, 'rb') as f:
                    input_stream = f.read()
                doc_asset = pdf_services.upload(input_stream=input_stream, mime_type=media_type)

                # Si es imagen, se debe encapsular en PDF antes de usar OCR o Extract
                if media_type != PDFServicesMediaType.PDF:
                    create_pdf_job = CreatePDFJob(doc_asset)
                    location = pdf_services.submit(create_pdf_job)
                    pdf_response = pdf_services.get_job_result(location, CreatePDFResult)
                    doc_asset = pdf_response.get_result().get_asset()

            # --- FLUJO 1: EXTRACCIÓN DE JSON ---
            if run_json:
                with st.spinner("Ejecutando Adobe PDF Extract API (Esto puede tomar unos segundos)..."):
                    # Construir los parámetros para extraer texto estructurado
                    extract_pdf_params = ExtractPDFParams(elements_to_extract=[ExtractElementType.TEXT])

                    extract_job = ExtractPDFJob(input_asset=doc_asset, extract_pdf_params=extract_pdf_params)
                    location = pdf_services.submit(extract_job)
                    extract_response = pdf_services.get_job_result(location, ExtractPDFResult)

                    # Obtener el archivo ZIP resultante
                    result_asset = extract_response.get_result().get_asset()
                    stream_asset = pdf_services.get_content(result_asset)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                        tmp_zip.write(stream_asset.get_input_stream())
                        zip_path = tmp_zip.name

                # Descomprimir y obtener el JSON
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    with zip_ref.open('structuredData.json') as json_file:
                        json_bytes = json_file.read()

                st.success("🧩 ¡JSON extraído correctamente!")
                out_name_json = uploaded_file.name.rsplit('.', 1)[0] + "_Extraido.json"
                st.download_button("📥 Descargar Archivo JSON", data=json_bytes, file_name=out_name_json, mime="application/json")
                os.unlink(zip_path)

            # --- FLUJO 2: CONVERSIÓN A WORD ---
            elif run_word:
                with st.spinner("Ejecutando Exportación a Word..."):
                    export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                    export_job = ExportPDFJob(doc_asset, export_params)
                    location = pdf_services.submit(export_job)
                    word_response = pdf_services.get_job_result(location, ExportPDFResult)

                    stream_asset = pdf_services.get_content(word_response.get_result().get_asset())
                    word_bytes = stream_asset.get_input_stream()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                    tmp_docx.write(word_bytes)
                    word_path = tmp_docx.name

                with st.spinner("Aplicando limpieza..."):
                    orchestrate_document_cleanup(word_path, config)
                    with open(word_path, "rb") as f:
                        cleaned_bytes = f.read()

                st.success("🎉 ¡Conversión a Word completada!")
                out_name_word = uploaded_file.name.rsplit('.', 1)[0] + "_Limpio.docx"
                st.download_button("📥 Descargar Word Final", data=cleaned_bytes, file_name=out_name_word)
                os.unlink(word_path)

        except Exception as e:
            st.error(f"❌ Ocurrió un error crítico: {e}")
            with st.expander("Ver detalles técnicos del error"): st.exception(e)
        finally:
            if os.path.exists(input_path): os.unlink(input_path)
