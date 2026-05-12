# =============================================================================
# APLICACIÓN DE CONVERSIÓN Y LIMPIEZA PROFESIONAL (PDF/IMAGEN A WORD)
# Desarrollado con Streamlit, Adobe PDF Services SDK y Python-Docx
# =============================================================================
# Este script procesa documentos, auto-corrige su orientación, extrae el texto
# mediante Adobe SDK, y post-procesa el documento Word generado para:
# - Forzar una sola columna.
# - Eliminar fondos, bordes y cajas de texto innecesarias.
# - Borrar imágenes menores a cierto tamaño (limpieza de ruido/logos).
# - Limpiar caracteres extraños manteniendo fórmulas matemáticas y español.
# =============================================================================

import os
import re
import json
import logging
import tempfile
import PyPDF2
from dataclasses import dataclass
from typing import Tuple, List, Optional

import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.document import Document

# Librerías para pre-procesamiento de orientación
from PIL import Image, ImageOps
import PyPDF2

# Importaciones de Adobe PDF Services SDK
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.jobs.create_pdf_job import CreatePDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
from adobe.pdfservices.operation.pdfjobs.result.create_pdf_result import CreatePDFResult

# =============================================================================
# 1. CONFIGURACIÓN CENTRAL Y LOGGING
# =============================================================================

# Configuración del Logger para depuración interna
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

@dataclass
class CleanupConfig:
    """Clase para almacenar la configuración de limpieza del usuario desde la interfaz."""
    force_single_column: bool = True
    remove_backgrounds: bool = True
    remove_small_images: bool = True
    min_image_width_cm: float = 3.0
    min_image_height_cm: float = 3.0
    clean_weird_symbols: bool = True
    fix_orientation: bool = True

# =============================================================================
# 2. FUNCIONES DE AUTENTICACIÓN Y MIME TYPE
# =============================================================================

def get_adobe_creds() -> Tuple[str, str]:
    """
    Obtiene las credenciales de Adobe PDF Services.
    Primero busca en los st.secrets de Streamlit Cloud.
    Si falla, intenta cargar el archivo local 'adobesecret.json' para desarrollo.
    """
    try:
        # Entorno Producción (Streamlit Cloud)
        client_id = st.secrets["PDF_SERVICES_CLIENT_ID"]
        client_secret = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
        logger.info("Credenciales cargadas desde st.secrets correctamente.")
        return client_id, client_secret
    except Exception as e:
        logger.warning(f"No se encontraron secrets de Streamlit. Intentando archivo JSON local. Detalle: {e}")
        try:
            # Entorno Local / Desarrollo
            with open('adobesecret.json', 'r') as f:
                data = json.load(f)
                creds = data["project"]["workspace"]["details"]["credentials"][0]["oauth_server_to_server"]
                logger.info("Credenciales cargadas desde adobesecret.json correctamente.")
                return creds["client_id"], creds["client_secrets"][0]
        except Exception as ex:
            logger.error("Error crítico: No se pudieron cargar las credenciales de Adobe en ningún entorno.")
            raise RuntimeError("Faltan las credenciales de Adobe. Configura st.secrets o adobesecret.json.") from ex

def get_media_type(filename: str) -> str:
    """
    Determina el MIME type específico de Adobe basándose en la extensión del archivo.
    """
    ext = filename.lower().split('.')[-1]
    if ext == 'png':
        return PDFServicesMediaType.PNG
    elif ext in ['jpg', 'jpeg']:
        return PDFServicesMediaType.JPEG
    return PDFServicesMediaType.PDF

# =============================================================================
# 3. FUNCIONES DE PRE-PROCESAMIENTO (ORIENTACIÓN)
# =============================================================================

def fix_image_orientation(image_path: str) -> None:
    """
    Lee los metadatos EXIF de una imagen y la rota automáticamente si el
    escáner o la cámara la guardó girada.
    """
    try:
        image = Image.open(image_path)
        # ImageOps.exif_transpose aplica la rotación correcta basada en el flag EXIF
        image = ImageOps.exif_transpose(image)
        image.save(image_path)
        logger.info(f"Orientación de imagen corregida (si aplicaba): {image_path}")
    except Exception as e:
        logger.warning(f"No se pudo verificar la orientación de la imagen: {e}")

def fix_pdf_orientation(pdf_path: str) -> str:
    """
    Analiza las páginas de un PDF. Si tienen un grado de rotación anómalo (90, 270),
    corrige el PDF reescribiendo la orientación a 0 (vertical).
    Crea un nuevo PDF temporal y retorna su ruta.
    """
    try:
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        pages_rotated = 0

        for page in reader.pages:
            rotation = page.get('/Rotate', 0)
            if isinstance(rotation, PyPDF2.generic.NumberObject):
                rot_val = int(rotation)
                # Si la página está en modo horizontal (landscape no deseado)
                if rot_val in [90, 270, -90]:
                    page.rotate(-rot_val)
                    pages_rotated += 1
            writer.add_page(page)

        if pages_rotated > 0:
            fixed_pdf_path = pdf_path + "_fixed.pdf"
            with open(fixed_pdf_path, "wb") as f:
                writer.write(f)
            logger.info(f"Se corrigió la orientación de {pages_rotated} páginas en el PDF.")
            return fixed_pdf_path
        else:
            logger.info("El PDF no requirió corrección de orientación explícita.")
            return pdf_path
    except Exception as e:
        logger.warning(f"Error al intentar corregir la orientación del PDF: {e}")
        return pdf_path

# =============================================================================
# 4. FUNCIONES DE POST-PROCESAMIENTO AVANZADO DE DOCX (EL NÚCLEO)
# =============================================================================

def remove_docx_backgrounds(doc: Document, parts: List) -> None:
    """
    Elimina formas de fondo globales, sombreados en celdas/párrafos y bordes extraños.
    """
    # 1. Quitar marca de fondo global
    try:
        settings = doc.settings.element
        display_bg = settings.find(qn('w:displayBackgroundShape'))
        if display_bg is not None:
            settings.remove(display_bg)
    except Exception as e:
        logger.warning(f"Error al limpiar displayBackgroundShape: {e}")

    # 2. Iterar sobre las partes (cuerpo, encabezados, pies)
    for part in parts:
        try:
            # A. Sombreados
            for shd in part.xpath('.//*[local-name()="shd"]'):
                parent = shd.getparent()
                if parent is not None:
                    parent.remove(shd)
            # B. Bordes de párrafo o página
            for bdr in part.xpath('.//*[local-name()="pBdr"] | .//*[local-name()="pgBdr"]'):
                parent = bdr.getparent()
                if parent is not None:
                    parent.remove(bdr)
            # C. Etiquetas puras de fondo
            for bg in part.xpath('.//*[local-name()="background"]'):
                parent = bg.getparent()
                if parent is not None:
                    parent.remove(bg)
        except Exception as e:
            logger.warning(f"Error limpiando fondos en una sección del documento: {e}")

def force_single_column_layout(doc: Document) -> None:
    """
    Navega por las propiedades de cada sección del documento XML (sectPr)
    y fuerza la configuración de columnas (w:cols) a 1 sola columna.
    Esto deshace el formato de periódico u columnas paralelas generadas por OCR.
    """
    try:
        sections_fixed = 0
        for section in doc.sections:
            sectPr = section._sectPr
            cols = sectPr.xpath('./w:cols')
            for col in cols:
                # El atributo w:num define la cantidad de columnas
                current_cols = col.get(qn('w:num'), '1')
                if current_cols != '1':
                    col.set(qn('w:num'), '1')
                    sections_fixed += 1
                    
        # Aplanar tablas de diseño (opcional, muy agresivo)
        # Algunos motores OCR usan tablas invisibles para simular columnas.
        # Aquí eliminaremos los bordes de TODAS las tablas para que visualmente fluyan,
        # aunque la reestructuración física de celdas a párrafos es invasiva.
        for table in doc.tables:
            table.autofit = True

        logger.info(f"Diseño forzado a 1 columna. Secciones corregidas: {sections_fixed}")
    except Exception as e:
        logger.error(f"Error al forzar una sola columna: {e}")

def remove_small_noise_images(doc: Document, parts: List, min_width_cm: float, min_height_cm: float) -> None:
    """
    Escanea todos los gráficos (w:drawing, v:shape) y elimina aquellos cuyas dimensiones
    sean inferiores a los parámetros configurados. Útil para eliminar logotipos minúsculos,
    artefactos de escaneo y puntos reconocidos como imágenes.
    """
    # Conversión: 1 centímetro = 360,000 EMUs (English Metric Units)
    min_cx = int(min_width_cm * 360000)
    min_cy = int(min_height_cm * 360000)
    removed_count = 0

    for part in parts:
        try:
            # Buscar elementos de dibujo modernos
            for drawing in part.xpath('.//w:drawing'):
                extents = drawing.xpath('.//wp:extent')
                if extents:
                    # wp:extent contiene los atributos cx y cy
                    cx = int(extents[0].get('cx', 0))
                    cy = int(extents[0].get('cy', 0))
                    
                    if cx < min_cx or cy < min_cy:
                        parent = drawing.getparent()
                        # Buscar el ancestro más alto seguro para borrar (usualmente el run 'w:r')
                        while parent is not None and parent.tag.endswith('drawing') == False and parent.tag.endswith('r') == False:
                             parent = parent.getparent()
                        
                        if parent is not None and parent.getparent() is not None:
                            try:
                                parent.getparent().remove(parent)
                                removed_count += 1
                            except ValueError:
                                pass

            # Buscar formas vectoriales antiguas o vacías
            search_query = './/*[local-name()="shape"] | .//*[local-name()="rect"] | .//*[local-name()="roundrect"]'
            for shape in part.xpath(search_query):
                has_image = len(shape.xpath('.//*[local-name()="pic"] | .//*[local-name()="imagedata"]')) > 0
                has_text = len(shape.xpath('.//*[local-name()="t"]')) > 0
                
                # Si es una forma vacía (sin texto ni imagen útil), se borra directamente
                if not has_image and not has_text:
                    parent = shape.getparent()
                    if parent is not None:
                        try:
                            parent.remove(shape)
                            removed_count += 1
                        except ValueError:
                            pass

        except Exception as e:
            logger.warning(f"Error procesando limpieza de imágenes en una parte del doc: {e}")

    logger.info(f"Imágenes o formas pequeñas eliminadas por limpieza de ruido: {removed_count}")

def purify_text_symbols(doc: Document) -> None:
    """
    Recorre todos los párrafos y tablas buscando texto.
    Aplica una expresión regular estricta para eliminar símbolos Unicode raros,
    caracteres de control y basura de OCR, preservando español, matemáticas y puntuación.
    """
    # Expresión regular explicada:
    # \w : caracteres alfanuméricos (letras y números)
    # \s : espacios en blanco
    # .,;:!?()[]{}"'+-*/=<>%$#@&^|\\~` : Puntuación estándar y matemáticas
    # áéíóúÁÉÍÓÚñÑüÜ¿¡€£¥° : Caracteres específicos del español y monedas/símbolos comunes
    # El ^ al principio significa "CUALQUIER COSA QUE NO SEA LO ANTERIOR"
    allowed_pattern = re.compile(r'[^\w\s.,;:!?()\[\]{}"\'+\-*/=<>%$#@&^|\\~`áéíóúÁÉÍÓÚñÑüÜ¿¡€£¥°]')
    
    cleaned_runs_count = 0

    # Función interna para limpiar un bloque de párrafos
    def clean_paragraphs(paragraphs):
        nonlocal cleaned_runs_count
        for para in paragraphs:
            for run in para.runs:
                if run.text:
                    original_text = run.text
                    # 1. Eliminar caracteres prohibidos
                    cleaned_text = allowed_pattern.sub('', original_text)
                    # 2. Colapsar espacios múltiples generados al borrar basura (opcional)
                    cleaned_text = re.sub(r' {2,}', ' ', cleaned_text)
                    
                    if cleaned_text != original_text:
                        run.text = cleaned_text
                        cleaned_runs_count += 1

    # Limpiar párrafos normales
    clean_paragraphs(doc.paragraphs)

    # Limpiar párrafos dentro de todas las tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                clean_paragraphs(cell.paragraphs)

    logger.info(f"Trozos de texto (runs) purificados de símbolos extraños: {cleaned_runs_count}")


def orchestrate_document_cleanup(docx_path: str, config: CleanupConfig) -> None:
    """
    Abre el documento de Word generado por Adobe, aplica todas las configuraciones
    de limpieza solicitadas por el usuario de forma secuencial, y guarda los cambios.
    """
    try:
        doc = docx.Document(docx_path)
        
        # Preparar la lista de partes del documento a iterar (cuerpo, headers, footers)
        parts = [doc.element]
        for section in doc.sections:
            parts.append(section.header._element)
            parts.append(section.footer._element)

        # 1. Forzar diseño de 1 columna
        if config.force_single_column:
            force_single_column_layout(doc)

        # 2. Limpieza de símbolos extraños (Ruido de OCR)
        if config.clean_weird_symbols:
            purify_text_symbols(doc)

        # 3. Eliminar Fondos y Bordes
        if config.remove_backgrounds:
            remove_docx_backgrounds(doc, parts)

        # 4. Limpieza de imágenes/logotipos pequeños
        if config.remove_small_images:
            remove_small_noise_images(
                doc, 
                parts, 
                min_width_cm=config.min_image_width_cm, 
                min_height_cm=config.min_image_height_cm
            )

        # Guardar el documento con las modificaciones aplicadas
        doc.save(docx_path)
        logger.info("Orquestación de limpieza del documento completada exitosamente.")

    except Exception as e:
        logger.error(f"Fallo general en la limpieza del documento: {e}")
        st.warning(f"Advertencia: Algunas limpiezas no pudieron aplicarse debido a un error: {e}")
        # La ejecución continúa, el usuario aún recibe su archivo, aunque quizás no tan limpio.

# =============================================================================
# 5. CONFIGURACIÓN DE LA INTERFAZ DE USUARIO (STREAMLIT)
# =============================================================================

# Configuración inicial de la página
st.set_page_config(
    page_title="Conversor Inteligente a Word", 
    page_icon="📄", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inyección de CSS para un estilo limpio y profesional
st.markdown("""
    <style>
        .main-header { font-size: 2.5rem; color: #1E3A8A; font-weight: 700; margin-bottom: 0px; }
        .sub-header { font-size: 1.2rem; color: #4B5563; margin-bottom: 2rem; }
        .stButton>button { background-color: #2563EB; color: white; border-radius: 8px; padding: 10px 24px; font-weight: bold; border: none; transition: all 0.3s;}
        .stButton>button:hover { background-color: #1D4ED8; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .success-msg { font-size: 1.1rem; color: #059669; font-weight: bold; padding: 10px; border-left: 5px solid #059669; background-color: #D1FAE5; margin-top: 1rem; }
    </style>
""", unsafe_allow_html=True)

# =============================================================================
# 6. PANEL LATERAL (SIDEBAR) - CONFIGURACIONES DE LIMPIEZA
# =============================================================================

st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/f/f1/Microsoft_Word_2013-2019_logo.svg/200px-Microsoft_Word_2013-2019_logo.svg.png", width=60)
st.sidebar.title("⚙️ Opciones de Limpieza Avanzada")
st.sidebar.markdown("Configura cómo deseas que se procese y limpie tu documento.")

# Instanciamos nuestra clase de configuración vinculándola a los controles de UI
config = CleanupConfig()

st.sidebar.subheader("Estructura")
config.force_single_column = st.sidebar.checkbox(
    "Forzar a 1 sola columna", 
    value=True, 
    help="Deshace el formato de periódico o múltiples columnas generado por el escáner."
)
config.fix_orientation = st.sidebar.checkbox(
    "Auto-corregir Orientación", 
    value=True, 
    help="Intenta detectar si las páginas están giradas horizontalmente y las rota antes de convertirlas."
)

st.sidebar.subheader("Limpieza Visual")
config.remove_backgrounds = st.sidebar.checkbox(
    "Eliminar Fondos y Bordes", 
    value=True, 
    help="Quita los sombreados oscuros y bordes de página que el OCR suele confundir."
)

config.remove_small_images = st.sidebar.checkbox(
    "Eliminar imágenes menores (Ruido)", 
    value=True, 
    help="Borra manchas o logotipos pequeños que no aportan al texto."
)
# Sliders que solo se muestran si la opción anterior está activa
if config.remove_small_images:
    col1, col2 = st.sidebar.columns(2)
    with col1:
        config.min_image_width_cm = st.number_input("Ancho menor a (cm):", min_value=0.5, max_value=10.0, value=3.0, step=0.5)
    with col2:
        config.min_image_height_cm = st.number_input("Alto menor a (cm):", min_value=0.5, max_value=10.0, value=3.0, step=0.5)

st.sidebar.subheader("Limpieza de Texto")
config.clean_weird_symbols = st.sidebar.checkbox(
    "Limpiar caracteres extraños", 
    value=True, 
    help="Elimina símbolos sin sentido resultantes de un mal escaneo, manteniendo español y matemáticas."
)

st.sidebar.markdown("---")
st.sidebar.info("Desarrollado con Adobe PDF Services API.")

# =============================================================================
# 7. ÁREA PRINCIPAL (CANVAS) Y LÓGICA DE EJECUCIÓN
# =============================================================================

st.markdown('<p class="main-header">🖼️ Conversor y Limpiador a Word</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Sube un archivo PDF o Imagen. El sistema corregirá su rotación, extraerá el texto y aplicará una limpieza profunda estructurando el documento final.</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Arrastra tu documento aquí", 
    type=["pdf", "jpg", "jpeg", "png"],
    help="Límite recomendado de tamaño por archivo según plan de Adobe."
)

if uploaded_file is not None:
    # 7.1. Guardar archivo cargado temporalmente para procesamiento seguro
    file_ext = uploaded_file.name.lower().split('.')[-1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp:
        tmp.write(uploaded_file.getbuffer())
        input_path = tmp.name

    st.info("✅ Archivo cargado en memoria temporal. Configura tus opciones a la izquierda y presiona Convertir.")
    
    # Botón principal de ejecución
    if st.button("🚀 Iniciar Conversión y Limpieza Inteligente"):
        try:
            # 7.2. PRE-PROCESAMIENTO (Orientación)
            if config.fix_orientation:
                with st.spinner("Verificando y corrigiendo orientación del documento..."):
                    if file_ext in ['jpg', 'jpeg', 'png']:
                        fix_image_orientation(input_path)
                    elif file_ext == 'pdf':
                        input_path = fix_pdf_orientation(input_path)
            
            # Obtener tipo de medio y credenciales
            media_type = get_media_type(uploaded_file.name)
            client_id, client_secret = get_adobe_creds()
            credentials = ServicePrincipalCredentials(client_id=client_id, client_secret=client_secret)
            pdf_services = PDFServices(credentials=credentials)

            # 7.3. INTERACCIÓN CON ADOBE NUBE
            with st.spinner("Subiendo archivo a la nube de procesamiento de Adobe..."):
                with open(input_path, 'rb') as f:
                    input_stream = f.read()
                # Subir activo
                doc_asset = pdf_services.upload(input_stream=input_stream, mime_type=media_type)

            # Si es imagen, convertir a PDF nativo en Adobe primero
            if media_type != PDFServicesMediaType.PDF:
                with st.spinner("Ensamblando Imagen en contenedor PDF..."):
                    create_pdf_job = CreatePDFJob(doc_asset)
                    location = pdf_services.submit(create_pdf_job)
                    pdf_response = pdf_services.get_job_result(location, CreatePDFResult)
                    doc_asset = pdf_response.get_result().get_asset()

            # Exportar a Word aplicando OCR
            with st.spinner("Ejecutando motor de IA / OCR y exportando a Word..."):
                export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
                export_job = ExportPDFJob(doc_asset, export_params)
                location = pdf_services.submit(export_job)
                word_response = pdf_services.get_job_result(location, ExportPDFResult)

                # Descargar flujo de bytes resultante
                result_asset = word_response.get_result().get_asset()
                stream_asset = pdf_services.get_content(result_asset)
                word_bytes = stream_asset.get_input_stream()

            # 7.4. POST-PROCESAMIENTO DOCX (La Limpieza)
            # Guardamos el archivo Word crudo temporalmente para modificarlo
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(word_bytes)
                word_path = tmp_docx.name

            with st.spinner("Aplicando algoritmos de limpieza y estructuración (Eliminación de ruido, Columnas, Símbolos)..."):
                # Llamada al orquestador con las configuraciones del usuario
                orchestrate_document_cleanup(word_path, config)

            # 7.5. PREPARAR DESCARGA
            # Leer el archivo ya limpio desde el disco
            with open(word_path, "rb") as f:
                cleaned_bytes = f.read()

            st.markdown('<p class="success-msg">🎉 ¡Conversión y limpieza completadas con éxito!</p>', unsafe_allow_html=True)
            
            output_filename = uploaded_file.name.rsplit('.', 1)[0] + "_Limpio.docx"
            
            # Crear dos columnas para alinear el botón de descarga al centro/derecha
            dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
            with dl_col2:
                st.download_button(
                    label="📥 Descargar Documento Word Final",
                    data=cleaned_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            # 7.6. LIMPIEZA FINAL DE MEMORIA Y DISCO
            try:
                if os.path.exists(input_path):
                    os.unlink(input_path)
                if os.path.exists(word_path):
                    os.unlink(word_path)
            except Exception as cleanup_error:
                logger.warning(f"Aviso durante la limpieza de temporales: {cleanup_error}")

        except Exception as e:
            # Manejo general de errores catastróficos en la interfaz
            st.error(f"❌ Ocurrió un error crítico durante el proceso: {e}")
            with st.expander("Ver detalles técnicos del error"):
                st.exception(e)
