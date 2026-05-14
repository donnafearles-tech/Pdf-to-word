# =============================================================================
# APLICACIÓN DE CONVERSIÓN Y EXTRACCIÓN (PDF/IMAGEN A WORD / JSON)
# CON IA GROQ INTEGRADA PARA SEMÁNTICA Y ORTOGRAFÍA
# =============================================================================

import os
import re
import json
import logging
import tempfile
import zipfile
import PyPDF2
from docx.shared import Inches
from dataclasses import dataclass
from typing import Tuple, List, Optional, Dict, Any

import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Librerías para pre-procesamiento de orientación e imagen
from PIL import Image, ImageOps

# Importación de Groq para corrección semántica
from groq import Groq

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

# Importaciones de Adobe PDF Extract SDK (JSON)
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
    remove_small_images: bool = True
    min_image_width_cm: float = 3.0
    min_image_height_cm: float = 3.0
    clean_weird_symbols: bool = True
    crop_left: int = 0
    crop_right: int = 0
    crop_top: int = 0
    crop_bottom: int = 0
    use_groq_correction: bool = False

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

def crop_image_margins(image_path: str, left: int, right: int, top: int, bottom: int) -> None:
    if left == 0 and right == 0 and top == 0 and bottom == 0:
        return
    try:
        img = Image.open(image_path)
        width, height = img.size
        x1 = int(width * (left / 100))
        y1 = int(height * (top / 100))
        x2 = width - int(width * (right / 100))
        y2 = height - int(height * (bottom / 100))
        if x1 < x2 and y1 < y2:
            img_cropped = img.crop((x1, y1, x2, y2))
            img_cropped.save(image_path)
    except Exception as e:
        logger.warning(f"Error recortando imagen: {e}")

def fix_pdf_orientation(pdf_path: str) -> str:
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
    except Exception:
        return pdf_path

# =============================================================================
# 3. FUNCIONES DE LIMPIEZA DOCX Y GROQ IA
# =============================================================================

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
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text:
                cleaned_text = allowed_pattern.sub('', run.text)
                cleaned_text = re.sub(r' {2,}', ' ', cleaned_text)
                if cleaned_text != run.text: run.text = cleaned_text

def remove_garbage_visuals(doc: Document, min_width_cm: float, min_height_cm: float) -> None:
    min_width_emu = min_width_cm * 360000
    min_height_emu = min_height_cm * 360000
    for drawing in doc.element.xpath('//w:drawing'):
        inlines = drawing.xpath('./wp:inline')
        for inline in inlines:
            extents = inline.xpath('./wp:extent')
            if extents:
                cx = int(extents[0].get('cx', 0))
                cy = int(extents[0].get('cy', 0))
                if cx < min_width_emu or cy < min_height_emu:
                    parent = drawing.getparent()
                    if parent is not None: parent.remove(drawing)
        anchors = drawing.xpath('./wp:anchor')
        if anchors:
            parent = drawing.getparent()
            if parent is not None: parent.remove(drawing)

def remove_garbage_paragraphs(doc: Document) -> None:
    word_pattern = re.compile(r'[a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]')
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text and not word_pattern.search(text):
            p_element = paragraph._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)

def correct_text_with_groq(doc: Document) -> None:
    try:
        api_key = st.secrets["GROQ_API_KEY"]
        client = Groq(api_key=api_key)
    except Exception:
        st.warning("⚠️ No se encontró 'GROQ_API_KEY' en secrets. Saltando corrección con IA.")
        return

    progress_bar = st.progress(0)
    total_paras = len([p for p in doc.paragraphs if len(p.text.strip()) > 10])
    procesados = 0

    for paragraph in doc.paragraphs:
        texto_original = paragraph.text.strip()
        if len(texto_original) > 10:
            try:
                chat_completion = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un editor experto. Corrige la ortografía, gramática y semántica del siguiente texto en español. Responde ÚNICAMENTE con el texto corregido. No añadas comillas, comentarios ni explicaciones adicionales."
                        },
                        {
                            "role": "user",
                            "content": texto_original,
                        }
                    ],
                    model="llama3-8b-8192",
                    temperature=0.2,
                )
                texto_corregido = chat_completion.choices[0].message.content.strip()
                paragraph.text = texto_corregido
            except Exception as e:
                logger.warning(f"Error en IA Groq: {e}")
            procesados += 1
            progress_bar.progress(procesados / total_paras)
    progress_bar.empty()

def orchestrate_document_cleanup(docx_path: str, config: CleanupConfig) -> None:
    doc = docx.Document(docx_path)
    
    if config.force_single_column: force_single_column_layout(doc)
    if config.clean_weird_symbols: purify_text_symbols(doc)
    if config.remove_small_images: remove_garbage_visuals(doc, config.min_image_width_cm, config.min_image_height_cm)
    
    remove_garbage_paragraphs(doc)
    
    if config.use_groq_correction:
        st.info("🧠 Aplicando Inteligencia Artificial de Groq (Semántica y Ortografía)...")
        correct_text_with_groq(doc)
        
    doc.save(docx_path)

# =============================================================================
# 4. NUEVA FUNCIÓN: CONVERTIR JSON (ADOBE EXTRACT) A WORD (CORREGIDA)
# =============================================================================

def flatten_elements(elements: List[Any], page_num: int = None, bounds: Dict = None) -> List[Dict]:
    """
    Recorre recursivamente la estructura de elementos de Adobe Extract
    y extrae una lista plana de elementos con texto, página y bounds.
    """
    flat = []
    if not isinstance(elements, list):
        elements = [elements] if elements else []

    for elem in elements:
        if not isinstance(elem, dict):
            continue

        current_page = elem.get('Page', page_num)
        current_bounds = elem.get('Bounds', bounds)

        text = elem.get('Text', '').strip()
        if text and current_page is not None:
            flat.append({
                'Page': current_page,
                'Bounds': current_bounds or {},
                'Text': text
            })

        children = elem.get('Elements', [])
        if children:
            flat.extend(flatten_elements(children, current_page, current_bounds))

    return flat

def group_elements_by_line(elements: List[Dict], page_num: int) -> List[List[Dict]]:
    """Agrupa elementos de una misma página por línea (mismo valor Y dentro de tolerancia)."""
    page_elements = [e for e in elements if e.get('Page') == page_num and e.get('Text', '').strip()]
    if not page_elements:
        return []

    page_elements.sort(key=lambda e: e.get('Bounds', {}).get('y', 0))

    lines = []
    current_line = []
    current_y = None
    y_tolerance = 5

    for elem in page_elements:
        y = elem.get('Bounds', {}).get('y', 0)
        if current_y is None or abs(y - current_y) <= y_tolerance:
            current_line.append(elem)
            if current_y is None:
                current_y = y
        else:
            current_line.sort(key=lambda e: e.get('Bounds', {}).get('x', 0))
            lines.append(current_line)
            current_line = [elem]
            current_y = y

    if current_line:
        current_line.sort(key=lambda e: e.get('Bounds', {}).get('x', 0))
        lines.append(current_line)

    return lines

def convert_json_to_word(json_bytes: bytes, config: CleanupConfig) -> bytes:
    """Convierte un JSON de Adobe Extract (structuredData.json) a Word usando estructura jerárquica."""
    data = json.loads(json_bytes)
    root_elements = data.get('elements', [])
    if not isinstance(root_elements, list):
        root_elements = [root_elements] if root_elements else []

    flat_elements = flatten_elements(root_elements)

    if not flat_elements:
        raise ValueError("No se encontraron elementos con texto en el JSON.")

    pages = sorted(set(e.get('Page', 1) for e in flat_elements))

    doc = docx.Document()

    for page_num in pages:
        lines = group_elements_by_line(flat_elements, page_num)

        for line in lines:
            line_text = ' '.join(elem.get('Text', '') for elem in line).strip()
            if line_text:
                is_heading = (len(line_text) < 100 and not line_text.endswith('.')) or len(line) == 1
                para = doc.add_paragraph()
                run = para.add_run(line_text)
                if is_heading:
                    run.bold = True
                para.paragraph_format.space_after = Inches(0.05)

        if page_num != pages[-1]:
            doc.add_page_break()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        doc.save(tmp_docx.name)
        tmp_path = tmp_docx.name

    orchestrate_document_cleanup(tmp_path, config)

    with open(tmp_path, "rb") as f:
        result_bytes = f.read()
    os.unlink(tmp_path)
    return result_bytes

# =============================================================================
# 5. INTERFAZ DE USUARIO Y EJECUCIÓN (STREAMLIT)
# =============================================================================

st.set_page_config(page_title="Multi-Herramienta Documental", layout="wide")
config = CleanupConfig()

st.sidebar.title("🛠️ Herramientas de Pre-Proceso")

st.sidebar.subheader("✂️ Recortar Imagen (Márgenes en %)")
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
config.use_groq_correction = st.sidebar.checkbox("✨ Corrección Ortográfica y Semántica (IA Groq)", value=False)

st.title("📄 Extractor JSON & Conversor Word")

# --- Opción 1: Convertir PDF/Imagen a Word o JSON ---
st.subheader("📎 Opción 1: Convertir PDF o imagen a Word / JSON")
uploaded_file = st.file_uploader("Arrastra tu documento aquí (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file is not None:
    file_ext = uploaded_file.name.lower().split('.')[-1]
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp:
        tmp.write(uploaded_file.getbuffer())
        input_path = tmp.name

    st.info("✅ Archivo cargado. Selecciona la acción que deseas realizar:")

    btn_json, btn_word = st.columns(2)
    run_json = btn_json.button("🧩 1. Extraer Solo JSON (Estructurado)", use_container_width=True)
    run_word = btn_word.button("🚀 2. Convertir y Limpiar a Word", use_container_width=True)

    if run_json or run_word:
        try:
            if file_ext in ['jpg', 'jpeg', 'png']:
                fix_image_orientation(input_path)
                crop_image_margins(input_path, config.crop_left, config.crop_right, config.crop_top, config.crop_bottom)
            elif file_ext == 'pdf':
                input_path = fix_pdf_orientation(input_path)

            media_type = get_media_type(uploaded_file.name)
            client_id, client_secret = get_adobe_creds()
            credentials = ServicePrincipalCredentials(client_id=client_id, client_secret=client_secret)
            pdf_services = PDFServices(credentials=credentials)

            with st.spinner("Subiendo y pre-procesando en Adobe..."):
                with open(input_path, 'rb') as f:
                    input_stream = f.read()
                doc_asset = pdf_services.upload(input_stream=input_stream, mime_type=media_type)

                if media_type != PDFServicesMediaType.PDF:
                    create_pdf_job = CreatePDFJob(doc_asset)
                    location = pdf_services.submit(create_pdf_job)
                    pdf_response = pdf_services.get_job_result(location, CreatePDFResult)
                    doc_asset = pdf_response.get_result().get_asset()

            # --- FLUJO JSON ---
            if run_json:
                with st.spinner("Ejecutando Adobe PDF Extract API..."):
                    extract_pdf_params = ExtractPDFParams(elements_to_extract=[ExtractElementType.TEXT])
                    extract_job = ExtractPDFJob(input_asset=doc_asset, extract_pdf_params=extract_pdf_params)
                    
                    location = pdf_services.submit(extract_job)
                    extract_response = pdf_services.get_job_result(location, ExtractPDFResult)

                    result_asset = extract_response.get_result().get_resource()
                    stream_asset = pdf_services.get_content(result_asset)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                        tmp_zip.write(stream_asset.get_input_stream())
                        zip_path = tmp_zip.name

                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    with zip_ref.open('structuredData.json') as json_file:
                        json_bytes = json_file.read()

                st.success("🧩 ¡JSON extraído correctamente!")
                out_name_json = uploaded_file.name.rsplit('.', 1)[0] + "_Extraido.json"
                st.download_button("📥 Descargar Archivo JSON", data=json_bytes, file_name=out_name_json, mime="application/json")
                os.unlink(zip_path)

            # --- FLUJO WORD ---
            elif run_word:
                with st.spinner("Ejecutando Exportación a Word..."):
                    export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                    export_job = ExportPDFJob(input_asset=doc_asset, export_pdf_params=export_params)
                    
                    location = pdf_services.submit(export_job)
                    word_response = pdf_services.get_job_result(location, ExportPDFResult)

                    stream_asset = pdf_services.get_content(word_response.get_result().get_asset())
                    word_bytes = stream_asset.get_input_stream()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                    tmp_docx.write(word_bytes)
                    word_path = tmp_docx.name

                with st.spinner("Aplicando limpieza y procesando IA..."):
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

# --- Opción 2: Convertir JSON existente a Word ---
st.markdown("---")
st.subheader("📄 Opción 2: Convertir archivo JSON (Adobe Extract) a Word")
uploaded_json = st.file_uploader("Carga tu archivo structuredData.json", type=["json"])

if uploaded_json is not None:
    if st.button("🔄 Convertir JSON a Word", use_container_width=True):
        try:
            json_bytes = uploaded_json.getvalue()
            with st.spinner("Generando Word a partir del JSON..."):
                word_bytes = convert_json_to_word(json_bytes, config)
            out_name = uploaded_json.name.rsplit('.', 1)[0] + "_Convertido.docx"
            st.success("✅ Conversión JSON → Word completada")
            st.download_button("📥 Descargar Word", data=word_bytes, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error al convertir JSON: {e}")
            with st.expander("Detalles técnicos"): st.exception(e)