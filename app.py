import streamlit as st
import os
import docx
import re
from groq import Groq

# --- IMPORTACIONES DEL SDK DE ADOBE ---
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# ==========================================
# 1. MOTOR DE CONVERSIÓN (ADOBE SDK)
# ==========================================
def convertir_pdf_a_word_adobe(input_pdf_path, output_docx_path, client_id, client_secret):
    """
    Convierte un PDF a DOCX usando la API oficial de Adobe.
    """
    try:
        # Autenticación
        credentials = ServicePrincipalCredentials(
            client_id=client_id, 
            client_secret=client_secret
        )
        pdf_services = PDFServices(credentials=credentials)

        # Subir el archivo a Adobe
        with open(input_pdf_path, 'rb') as f:
            asset = pdf_services.upload(input_stream=f, mime_type=PDFServicesMediaType.PDF.value)

        # Configurar el trabajo de exportación a DOCX
        params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
        job = ExportPDFJob(asset=asset, export_pdf_params=params)

        # Ejecutar y esperar resultado
        location = pdf_services.submit(job)
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
        
        # Obtener el activo resultante
        result_asset = pdf_services_response.get_result().get_asset()
        stream_asset = pdf_services.get_content(result_asset)

        # Guardar el DOCX localmente
        with open(output_docx_path, "wb") as f:
            f.write(stream_asset.read())
            
        return True

    except Exception as e:
        import streamlit as st
        st.error(f"Error fatal en Adobe PDF Services: {e}")
        return False

# ==========================================
# 2. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ)
# ==========================================
def pre_limpiar_ocr(texto):
    # Elimina basura dura antes de gastar tokens
    texto = re.sub(r'[_<>|~^]', '', texto)
    return re.sub(r'\s+', ' ', texto).strip()

def limpiar_y_traducir_con_groq(texto, groq_api_key):
    cliente = Groq(api_key=groq_api_key)
    prompt_sistema = (
        "Eres un editor editorial experto. "
        "1. Traduce al ESPAÑOL. "
        "2. Elimina basura de escaneo y sílabas sin sentido. "
        "3. Devuelve ÚNICAMENTE el texto final."
    )
    
    try:
        respuesta = cliente.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": texto}
            ],
            temperature=0.1,
            max_tokens=8192
        )
        return respuesta.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"Error con Groq: {e}")
        return texto

def procesar_docx_con_groq(docx_path, groq_api_key):
    doc = docx.Document(docx_path)
    
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text.strip()
        
        if not texto_original or texto_original.isdigit():
            continue
            
        texto_pre_limpio = pre_limpiar_ocr(texto_original)
        
        if len(texto_pre_limpio) > 3:
            texto_final = limpiar_y_traducir_con_groq(texto_pre_limpio, groq_api_key)
            
            # Preservar estilos del primer run
            estilo_previo = None
            if parrafo.runs and parrafo.runs[0].style:
                estilo_previo = parrafo.runs[0].style

            # Limpiar contenido anterior
            for run in parrafo.runs:
                run.text = ""
                
            # Insertar texto nuevo
            nuevo_run = parrafo.add_run(texto_final)
            if estilo_previo:
                nuevo_run.style = estilo_previo

    doc.save(docx_path)

# ==========================================
# 3. INTERFAZ DE STREAMLIT
# ==========================================
st.title("Conversor Editorial: PDF a Word Limpio")

# Cargar secretos de Streamlit
try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError:
    st.error("Faltan credenciales en st.secrets. Revisa tu configuración.")
    st.stop()

archivo_subido = st.file_uploader("Sube tu PDF escaneado", type=["pdf"])

if archivo_subido and st.button("Procesar Documento"):
    with st.spinner("Paso 1/2: Adobe está convirtiendo el PDF a Word... (Esto puede tardar)"):
        # Guardar archivo temporalmente
        temp_pdf = "temp_input.pdf"
        temp_docx = "temp_output.docx"
        
        with open(temp_pdf, "wb") as f:
            f.write(archivo_subido.getbuffer())
            
        # Ejecutar Adobe SDK
        exito_adobe = convertir_pdf_a_word_adobe(
            temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
        )
        
    if exito_adobe:
        with st.spinner("Paso 2/2: Groq está limpiando y traduciendo el texto..."):
            # Ejecutar Limpieza Groq
            procesar_docx_con_groq(temp_docx, GROQ_API_KEY)
            
        st.success("¡Documento procesado con éxito!")
        
        # Botón de descarga
        with open(temp_docx, "rb") as f:
            st.download_button(
                label="📥 Descargar Word Limpio",
                data=f,
                file_name="Libro_Limpiado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        # Limpieza de archivos temporales
        os.remove(temp_pdf)
        os.remove(temp_docx)
