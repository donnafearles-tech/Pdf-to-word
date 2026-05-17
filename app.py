import os
import re
import time
import docx
import streamlit as st
from groq import Groq

# --- IMPORTACIONES OFICIALES DEL SDK DE ADOBE (V4) ---
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# =====================================================================
# 1. MOTOR DE CONVERSIÓN (ADOBE SDK V4)
# =====================================================================
def convertir_pdf_a_word_adobe(input_pdf_path, output_docx_path, client_id, client_secret):
    """
    Convierte un PDF a DOCX usando la API oficial de Adobe (SDK v4).
    Maneja correctamente la lectura y escritura de streams nativos.
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
            
        # Subir el archivo temporal a los servidores de Adobe
        asset = pdf_services.upload(input_stream=pdf_bytes, mime_type=PDFServicesMediaType.PDF)

        # Configurar el trabajo de exportación a formato DOCX
        params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
        job = ExportPDFJob(input_asset=asset, export_pdf_params=params)
        
        # Enviar el trabajo al servidor
        location = pdf_services.submit(job)
        
        # Esperar y obtener el resultado del proceso
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
        result_asset = pdf_services_response.get_result().get_asset()
        
        # Descargar el contenido convertido
        stream_asset = pdf_services.get_content(result_asset)

        # Guardar los bytes extraídos del stream interno en el archivo DOCX local
        with open(output_docx_path, "wb") as f:
            f.write(stream_asset.get_input_stream().read())
            
        return True

    except Exception as e:
        st.error(f"Error fatal en Adobe PDF Services: {e}")
        return False

# =====================================================================
# 2. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ)
# =====================================================================
def pre_limpiar_ocr(texto):
    """
    Elimina caracteres residuales comunes del OCR duro antes de enviar a la IA.
    Ahorra tokens y evita errores de contexto.
    """
    texto = re.sub(r'[_<>|~^]', '', texto)
    return re.sub(r'\s+', ' ', texto).strip()

def limpiar_y_traducir_con_groq(texto, groq_api_key):
    """
    Llama a Groq para limpiar el ruido del escáner y traducir al español.
    Límite controlado de tokens para evitar saturación de la API.
    """
    cliente = Groq(api_key=groq_api_key)
    prompt_sistema = (
        "Eres un editor editorial experto en restauración de textos escaneados (OCR).\n"
        "Tus instrucciones estructurales son estrictas:\n"
        "1. Traduce el texto al ESPAÑOL de forma fluida y natural.\n"
        "2. Elimina toda la basura del escaneo: símbolos sin sentido, caracteres rotos o sílabas repetitivas.\n"
        "3. Corrige la ortografía y puntuación para dejar un texto limpio y listo para publicación.\n"
        "4. Devuelve ÚNICAMENTE el texto traducido y corregido. No agregues introducciones, notas ni explicaciones."
    )
    
    try:
        respuesta = cliente.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": texto}
            ],
            temperature=0.1,
            max_tokens=1500  # Ajustado para no exceder los límites por minuto de la cuenta
        )
        return respuesta.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"Error con Groq: {e}")
        return texto  # En caso de error, conserva el texto original para no perder datos

def procesar_docx_con_groq(docx_path, groq_api_key):
    """
    Abre el documento de Word generado, limpia y traduce párrafo por párrafo.
    Preserva las referencias de estilo e introduce pausas para respetar las cuotas TPM.
    """
    doc = docx.Document(docx_path)
    
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text.strip()
        
        # Ignorar líneas vacías o elementos numéricos huérfanos (como números de página)
        if not texto_original or texto_original.isdigit():
            continue
            
        texto_pre_limpio = pre_limpiar_ocr(texto_original)
        
        if len(texto_pre_limpio) > 3:
            texto_final = limpiar_y_traducir_con_groq(texto_pre_limpio, groq_api_key)
            
            # Almacenar estilos tipográficos básicos del fragmento original si existen
            estilo_previo = None
            if parrafo.runs and parrafo.runs[0].style:
                estilo_previo = parrafo.runs[0].style

            # Vaciar los fragmentos de texto anteriores para reescribir limpiamente
            for run in parrafo.runs:
                run.text = ""
                
            # Insertar el texto procesado por la IA
            nuevo_run = parrafo.add_run(texto_final)
            if estilo_previo:
                nuevo_run.style = estilo_previo
            
            # Control de flujo (Rate Limiting): Pausa obligatoria de 2.5 segundos por párrafo
            # Esto evita el error 413 al mantenerse por debajo de los 6,000 tokens por minuto.
            time.sleep(2.5)

    doc.save(docx_path)

# =====================================================================
# 3. INTERFAZ DE USUARIO (STREAMLIT ONLINE)
# =====================================================================
st.set_page_config(page_title="Conversor Editorial PDF", page_icon="📚")
st.title("Conversor Editorial: PDF a Word Limpio")
st.write("Sube tus archivos PDF escaneados para convertirlos a Word, traducirlos al español y remover ruido de OCR.")

# Carga segura de credenciales desde el panel de Secrets de Streamlit Cloud
try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError:
    st.error("Error de configuración: Faltan las credenciales de las APIs en los Secrets de Streamlit.")
    st.stop()

# Componente de subida de archivos
archivo_subido = st.file_uploader("Selecciona el libro o documento en formato PDF", type=["pdf"])

if archivo_subido:
    if st.button("Comenzar Procesamiento Editorial"):
        # Definición de rutas temporales dentro del contenedor de Streamlit
        temp_pdf = "temp_input.pdf"
        temp_docx = "temp_output.docx"
        
        # Escribir el archivo cargado en el almacenamiento temporal
        with open(temp_pdf, "wb") as f:
            f.write(archivo_subido.getbuffer())
            
        # --- FASE 1: Conversor de Adobe ---
        with st.spinner("Fase 1/2: Convirtiendo estructura de PDF a Word con Adobe API..."):
            exito_adobe = convertir_pdf_a_word_adobe(
                temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
            )
            
        # --- FASE 2: Limpieza y Traducción con Groq ---
        if exito_adobe:
            with st.spinner("Fase 2/2: Procesando texto con Groq (Limpieza de ruido y traducción)..."):
                procesar_docx_con_groq(temp_docx, GROQ_API_KEY)
                
            st.success("¡El documento ha sido procesado y restaurado con éxito!")
            
            # Generar el botón de descarga para la usuaria
            with open(temp_docx, "rb") as f:
                st.download_button(
                    label="📥 Descargar Documento Word Limpio",
                    data=f,
                    file_name="Libro_Procesado_E_Inmaculado.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            # Remoción de archivos del servidor temporal para liberar espacio
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            if os.path.exists(temp_docx):
                os.remove(temp_docx)
        else:
            st.error("No se pudo completar la Fase 1 debido a un problema con el servicio de Adobe.")
