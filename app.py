import os
import re
import time
import docx
import streamlit as st
from groq import Groq
from docx.shared import Cm

# =====================================================================
# IMPORTACIONES OFICIALES DEL SDK DE ADOBE (V4)
# =====================================================================
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# =====================================================================
# CONFIGURACIÓN DE LA PÁGINA DE STREAMLIT (Debe ir primero)
# =====================================================================
st.set_page_config(
    page_title="Conversor Editorial PDF", 
    page_icon="📚", 
    layout="centered"
)

# =====================================================================
# 1. MOTOR DE CONVERSIÓN (ADOBE SDK V4) - BLINDADO
# =====================================================================

def limpiar_imagenes_pequenas(doc, min_width_cm=2.0, min_height_cm=2.0):
    """
    Itera sobre las imágenes del documento y elimina las que sean más pequeñas 
    que el umbral especificado (útil para borrar manchas o artefactos de OCR).
    """
    imagenes_eliminadas = 0
    # doc.inline_shapes contiene las imágenes incrustadas en el texto
    for shape in doc.inline_shapes:
        # Convertir el tamaño a centímetros para una validación lógica
        ancho = shape.width.cm
        alto = shape.height.cm
        
        if ancho < min_width_cm or alto < min_height_cm:
            # Eliminar el nodo XML de la imagen del documento
            nodo_imagen = shape._inline
            nodo_imagen.getparent().remove(nodo_imagen)
            imagenes_eliminadas += 1
            
    return imagenes_eliminadas
def convertir_pdf_a_word_adobe(input_pdf_path, output_docx_path, client_id, client_secret):
    """
    Convierte un PDF a DOCX usando la API oficial de Adobe (SDK v4).
    Maneja correctamente la lectura y escritura de bytes puros.
    """
    try:
        # Autenticación con las credenciales del servicio
        credentials = ServicePrincipalCredentials(
            client_id=client_id, 
            client_secret=client_secret
        )
        pdf_services = PDFServices(credentials=credentials)

        # Leer el PDF como bytes antes de enviarlo
        with open(input_pdf_path, 'rb') as f:
            pdf_bytes = f.read()
            
        # Subir el archivo a los servidores de Adobe
        asset = pdf_services.upload(input_stream=pdf_bytes, mime_type=PDFServicesMediaType.PDF)

        # Configurar el trabajo de exportación a formato DOCX
        params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
        job = ExportPDFJob(input_asset=asset, export_pdf_params=params)
        
        # Enviar el trabajo al servidor y esperar
        location = pdf_services.submit(job)
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
        
        # Extraer el resultado
        result_asset = pdf_services_response.get_result().get_asset()
        stream_asset = pdf_services.get_content(result_asset)

        # GUARDADO CORREGIDO: get_input_stream() ya devuelve los bytes en esta versión, no necesita .read()
        with open(output_docx_path, "wb") as f:
            f.write(stream_asset.get_input_stream())
            
        return True

    except Exception as e:
        st.error(f"Error fatal en Adobe PDF Services al convertir: {str(e)}")
        return False

# =====================================================================
# 2. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ) - CON LÍMITE DE TASA
# =====================================================================
def pre_limpiar_ocr(texto):
    """
    Conserva únicamente el alfabeto inglés/español, números y puntuación estándar.
    Elimina artefactos de OCR antes de gastar tokens de IA.
    """
    # 1. Conservar solo letras (inglés/español), números, espacios y puntuación básica
    # Se permiten: a-z, A-Z, vocales acentuadas, ñ, diéresis, números y signos comunes.
    patron_permitido = r'[^a-zA-ZáéíóúÁÉÍÓÚñÑüÜ0-9\s.,;:\-!?¿¡"\'\(\)\[\]/]'
    texto_limpio = re.sub(patron_permitido, '', texto)
    
    # 2. Colapsar espacios múltiples (común en ruido espacial OCR)
    texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()
    
    return texto_limpio

def limpiar_y_traducir_con_groq(texto, groq_api_key):
    """Llama a Groq para limpiar y traducir, blindado contra fallos de red."""
    cliente = Groq(api_key=groq_api_key)
    prompt_sistema = (
        "Eres un editor editorial experto en restauración de textos escaneados.\n"
        "1. Traduce el texto al ESPAÑOL de forma natural.\n"
        "2. Elimina basura de escaneo: símbolos sin sentido o sílabas rotas.\n"
        "3. Corrige la ortografía y puntuación.\n"
        "4. Devuelve ÚNICAMENTE el texto traducido. Sin introducciones."
    )
    
    try:
        respuesta = cliente.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": texto}
            ],
            temperature=0.1,
            max_tokens=1500  # Protege la cuota límite de tokens
        )
        return respuesta.choices[0].message.content.strip()
    except Exception as e:
        # Si un párrafo falla (ej. caída de red de Groq), se devuelve el original
        # para que la aplicación no se detenga por completo.
        st.warning(f"Aviso: Un párrafo no pudo procesarse con IA. Se mantendrá el original. Error: {str(e)}")
        return texto

def procesar_docx_con_groq(docx_path, groq_api_key):
    """Itera sobre el Word preservando formatos, limpia imágenes basura y traduce."""
    doc = docx.Document(docx_path)
    
    # NUEVO: Limpiar imágenes inútiles primero
    texto_estado = st.empty()
    texto_estado.text("Purgando imágenes minúsculas y ruido visual...")
    img_eliminadas = limpiar_imagenes_pequenas(doc, min_width_cm=1.5, min_height_cm=1.5)
    st.info(f"🧹 Se eliminaron {img_eliminadas} artefactos visuales/imágenes pequeñas.")
    
    # Barra de progreso
    barra_progreso = st.progress(0)
    
    total_parrafos = len(doc.paragraphs)
    parrafos_procesados = 0
    
    for i, parrafo in enumerate(doc.paragraphs):
        # ... (Tu código original de iteración de párrafos se mantiene igual) ...
        texto_original = parrafo.text.strip()
        
        progreso = int(((i + 1) / total_parrafos) * 100)
        barra_progreso.progress(progreso)
        texto_estado.text(f"Limpiando y traduciendo párrafo {i + 1} de {total_parrafos}...")
        
        if not texto_original or texto_original.isdigit():
            continue
            
        texto_pre_limpio = pre_limpiar_ocr(texto_original)
        
        if len(texto_pre_limpio) > 3:
            # ... (Llamada a Groq y reconstrucción del run se mantiene igual) ...

# =====================================================================
# 3. INTERFAZ DE USUARIO Y CONTROL DE FLUJO PRINCIPAL
# =====================================================================
st.title("Conversor Editorial: PDF a Word Limpio")
st.markdown("Sube tus archivos **PDF escaneados** para convertirlos a **Word**, traducirlos al español y remover ruido de OCR.")

# Carga de credenciales (Blindado)
try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError as e:
    st.error(f"❌ Error crítico: Falta la credencial {e} en los Secrets de Streamlit.")
    st.stop()

# Uploader
archivo_subido = st.file_uploader("Selecciona el libro o documento en formato PDF", type=["pdf"])

if archivo_subido:
    if st.button("Comenzar Procesamiento Editorial", type="primary"):
        
        # Nombres de archivos temporales únicos (evita choques si lo usas en pestañas)
        id_unico = str(int(time.time()))
        temp_pdf = f"temp_input_{id_unico}.pdf"
        temp_docx = f"temp_output_{id_unico}.docx"
        
        try:
            # 1. Guardar el PDF subido al servidor temporal
            with open(temp_pdf, "wb") as f:
                f.write(archivo_subido.getbuffer())
                
            # 2. Ejecutar Fase 1: Adobe
            with st.spinner("Fase 1/2: Convirtiendo estructura del PDF a Word en servidores de Adobe..."):
                exito_adobe = convertir_pdf_a_word_adobe(
                    temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
                )
                
            # 3. Ejecutar Fase 2: Groq
            if exito_adobe:
                with st.spinner("Fase 2/2: Inicializando Inteligencia Artificial para limpieza..."):
                    procesar_docx_con_groq(temp_docx, GROQ_API_KEY)
                    
                st.success("🎉 ¡El documento ha sido procesado y restaurado con éxito!")
                st.balloons()
                
                # 4. Generar botón de descarga
                with open(temp_docx, "rb") as f:
                    st.download_button(
                        label="📥 Descargar Documento Word Limpio",
                        data=f,
                        file_name="Libro_Procesado_Limpio.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error("❌ El proceso se detuvo porque la conversión de Adobe falló.")
                
        except Exception as e:
            st.error(f"Ha ocurrido un error inesperado en la aplicación: {str(e)}")
            
        finally:
            # BLOQUE BLINDADO: Se ejecuta SIEMPRE, haya éxito o haya error.
            # Limpia los archivos temporales para que tu servidor no se sature.
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            if os.path.exists(temp_docx):
                os.remove(temp_docx)
