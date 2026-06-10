import os
import re
import time
import docx
import streamlit as st
from groq import Groq

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
# CONFIGURACIÓN DE LA PÁGINA DE STREAMLIT
# =====================================================================
st.set_page_config(
    page_title="Conversor Editorial PDF", 
    page_icon="📚", 
    layout="centered"
)

# =====================================================================
# 1. MOTOR DE CONVERSIÓN (ADOBE SDK V4) - BLINDADO
# =====================================================================
def convertir_pdf_a_word_adobe(input_pdf_path, output_docx_path, client_id, client_secret):
    """
    Convierte un PDF a DOCX usando la API oficial de Adobe (SDK v4).
    Maneja correctamente la lectura y escritura de bytes puros.
    """
    try:
        credentials = ServicePrincipalCredentials(
            client_id=client_id, 
            client_secret=client_secret
        )
        pdf_services = PDFServices(credentials=credentials)

        with open(input_pdf_path, 'rb') as f:
            pdf_bytes = f.read()
            
        asset = pdf_services.upload(input_stream=pdf_bytes, mime_type=PDFServicesMediaType.PDF)

        params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
        job = ExportPDFJob(input_asset=asset, export_pdf_params=params)
        
        location = pdf_services.submit(job)
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
        
        result_asset = pdf_services_response.get_result().get_asset()
        stream_asset = pdf_services.get_content(result_asset)

        with open(output_docx_path, "wb") as f:
            f.write(stream_asset.get_input_stream())
            
        return True

    except Exception as e:
        st.error(f"Error fatal en Adobe PDF Services al convertir: {str(e)}")
        return False

# =====================================================================
# 2. HEURÍSTICAS DE FILTRADO Y LIMPIEZA (TEXTO Y GRÁFICOS)
# =====================================================================
def limpiar_imagenes_pequenas(doc, min_width_cm=1.5, min_height_cm=1.5):
    """
    Itera sobre las imágenes incrustadas y elimina las que sean más pequeñas 
    que el umbral especificado para purgar manchas, logos o ruido de escaneo.
    """
    imagenes_eliminadas = 0
    for shape in doc.inline_shapes:
        try:
            ancho = shape.width.cm
            alto = shape.height.cm
            
            if ancho < min_width_cm or alto < min_height_cm:
                nodo_imagen = shape._inline
                nodo_imagen.getparent().remove(nodo_imagen)
                imagenes_eliminadas += 1
        except Exception:
            continue
            
    return imagenes_eliminadas

def pre_limpiar_ocr(texto):
    """
    Conserva únicamente el alfabeto inglés/español, números y puntuación estándar.
    Elimina ráfagas de símbolos basura del OCR antes de procesar con la IA.
    """
    # Expresión regular inclusiva (filtra todo lo que NO sea letra es/en, número o puntuación básica)
    patron_permitido = r'[^a-zA-ZáéíóúÁÉÍÓÚñÑüÜ0-9\s.,;:\-!?¿¡"\'\(\)\[\]/]'
    texto_limpio = re.sub(patron_permitido, '', texto)
    
    # Colapsar espacios múltiples y saltos de línea huérfanos
    return re.sub(r'\s+', ' ', texto_limpio).strip()

# =====================================================================
# 3. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ) - CON BATCH PROCESSING
# =====================================================================
def limpiar_y_traducir_lote(texto_lote, groq_api_key, max_reintentos=3):
    """
    Procesa múltiples párrafos en una sola llamada a Groq.
    texto_lote es una cadena con párrafos separados por '\n|||SEP|||\n'
    Retorna una lista de párrafos traducidos en el mismo orden.
    """
    cliente = Groq(api_key=groq_api_key)
    prompt_sistema = (
        "Eres un editor editorial experto en restauración de textos escaneados.\n"
        "Se te pasarán múltiples bloques de texto separados por la línea '|||SEP|||'.\n"
        "Para CADA bloque:\n"
        "1. Traduce al ESPAÑOL de forma natural.\n"
        "2. Elimina basura de escaneo: símbolos sin sentido o sílabas rotas.\n"
        "3. Corrige la ortografía y puntuación.\n"
        "Devuelve cada bloque traducido separado por exactamente esta línea: '|||SEP|||'\n"
        "IMPORTANTE: Mantén el mismo número de bloques. Sin introducciones ni explicaciones."
    )
    
    for intento in range(max_reintentos):
        try:
            respuesta = cliente.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": prompt_sistema},
                    {"role": "user", "content": texto_lote}
                ],
                temperature=0.1,
                max_tokens=3000  # Incrementado para lotes
            )
            resultado = respuesta.choices[0].message.content.strip()
            bloques = resultado.split('|||SEP|||')
            return [b.strip() for b in bloques]
            
        except Exception as e:
            error_msg = str(e).lower()
            if "rate limit" in error_msg or "429" in error_msg:
                tiempo_espera = 60 * (intento + 1)  # Backoff exponencial
                st.warning(f"⏳ Cuota de Groq saturada. Esperando {tiempo_espera}s... (Intento {intento + 1}/{max_reintentos})")
                time.sleep(tiempo_espera)
            else:
                st.warning(f"Aviso: Error en lote. Se mantendrán los originales. Error: {str(e)}")
                # Retornar los párrafos originales sin procesar
                return [p.strip() for p in texto_lote.split('|||SEP|||')]
    
    # Si agota reintentos, retornar originales
    return [p.strip() for p in texto_lote.split('|||SEP|||')]

def procesar_docx_con_groq(docx_path, groq_api_key, tamaño_lote=5):
    """Itera sobre el Word en lotes, reduciendo drásticamente las llamadas a API."""
    doc = docx.Document(docx_path)
    
    texto_estado = st.empty()
    texto_estado.text("Fase 2a: Purgando imágenes minúsculas y ruido visual de escaneo...")
    
    # 1. Purgar imágenes inútiles primero
    img_eliminadas = limpiar_imagenes_pequenas(doc, min_width_cm=1.5, min_height_cm=1.5)
    st.info(f"🧹 Se eliminaron {img_eliminadas} artefactos visuales/imágenes pequeñas.")
    
    # 2. Pre-filtrar párrafos válidos (eliminar vacíos y solo-dígitos)
    parrafos_validos = []
    indices_validos = []
    
    for i, parrafo in enumerate(doc.paragraphs):
        texto_original = parrafo.text.strip()
        if texto_original and not texto_original.isdigit():
            parrafos_validos.append(parrafo)
            indices_validos.append(i)
    
    if not parrafos_validos:
        st.info("No hay párrafos válidos para procesar.")
        return
    
    # 3. Procesar en lotes
    barra_progreso = st.progress(0)
    total_lotes = (len(parrafos_validos) + tamaño_lote - 1) // tamaño_lote
    parrafos_procesados = 0
    
    for lote_idx in range(0, len(parrafos_validos), tamaño_lote):
        lote_parrafos = parrafos_validos[lote_idx:lote_idx + tamaño_lote]
        lote_numero = (lote_idx // tamaño_lote) + 1
        
        # Construir texto del lote
        textos_limpios = []
        for parrafo in lote_parrafos:
            texto_pre_limpio = pre_limpiar_ocr(parrafo.text.strip())
            if len(texto_pre_limpio) > 3:
                textos_limpios.append(texto_pre_limpio)
            else:
                textos_limpios.append("")  # Preservar orden incluso con párrafos vacíos
        
        texto_lote = '\n|||SEP|||\n'.join(textos_limpios)
        
        # Procesar lote
        texto_estado.text(f"Procesando lote {lote_numero}/{total_lotes}...")
        textos_procesados = limpiar_y_traducir_lote(texto_lote, groq_api_key)
        
        # Aplicar resultados al documento
        for i, parrafo in enumerate(lote_parrafos):
            if i < len(textos_procesados) and textos_procesados[i]:
                estilo_previo = None
                if parrafo.runs and parrafo.runs[0].style:
                    estilo_previo = parrafo.runs[0].style
                
                for run in parrafo.runs:
                    run.text = ""
                
                nuevo_run = parrafo.add_run(textos_procesados[i])
                if estilo_previo:
                    nuevo_run.style = estilo_previo
                
                parrafos_procesados += 1
        
        # Actualizar barra
        progreso = int(((lote_idx + len(lote_parrafos)) / len(parrafos_validos)) * 100)
        barra_progreso.progress(min(progreso, 100))
        
        # Guardado incremental cada 3 lotes
        if lote_idx > 0 and (lote_idx // tamaño_lote) % 3 == 0:
            doc.save(docx_path)
            texto_estado.text(f"💾 Guardado incremental en lote {lote_numero}...")
        
        time.sleep(1)  # Pausa moderada entre lotes
    
    doc.save(docx_path)
    texto_estado.text(f"✅ Completado. {parrafos_procesados} párrafos mejorados.")
    barra_progreso.empty()

# =====================================================================
# 4. INTERFAZ DE USUARIO Y CONTROL DE FLUJO PRINCIPAL
# =====================================================================
st.title("Conversor Editorial: PDF a Word Limpio")
st.markdown("Sube tus archivos **PDF escaneados** para convertirlos a **Word**, traducirlos al español y remover ruido de OCR.")

try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError as e:
    st.error(f"❌ Error crítico: Falta la credencial {e} en los Secrets de Streamlit.")
    st.stop()

archivo_subido = st.file_uploader("Selecciona el libro o documento en formato PDF", type=["pdf"])

if archivo_subido:
    if st.button("Comenzar Procesamiento Editorial", type="primary"):
        
        id_unico = str(int(time.time()))
        temp_pdf = f"temp_input_{id_unico}.pdf"
        temp_docx = f"temp_output_{id_unico}.docx"
        
        try:
            with open(temp_pdf, "wb") as f:
                f.write(archivo_subido.getbuffer())
                
            with st.spinner("Fase 1/2: Convirtiendo estructura del PDF a Word en servidores de Adobe..."):
                exito_adobe = convertir_pdf_a_word_adobe(
                    temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
                )
                
            if exito_adobe:
                with st.spinner("Fase 2/2: Inicializando Inteligencia Artificial para limpieza..."):
                    procesar_docx_con_groq(temp_docx, GROQ_API_KEY)
                    
                st.success("🎉 ¡El documento ha sido procesado y restaurado con éxito!")
                st.balloons()
                
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
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            if os.path.exists(temp_docx):
                os.remove(temp_docx)
