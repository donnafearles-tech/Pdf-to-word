import os
import re
import time
import shutil
import docx
import streamlit as st
from groq import Groq
from datetime import datetime
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import OxmlElement

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
# FUNCIONES AUXILIARES DE GESTIÓN DE ARCHIVOS
# =====================================================================
def crear_carpeta_resultados():
    """Crea la carpeta 'resultados' si no existe."""
    carpeta = "resultados"
    if not os.path.exists(carpeta):
        os.makedirs(carpeta)
    return carpeta

def mover_docx_a_resultados(docx_path):
    """
    Mueve el DOCX procesado a la carpeta de resultados con timestamp.
    Retorna la ruta final del archivo.
    """
    carpeta_resultados = crear_carpeta_resultados()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_original = os.path.basename(docx_path).replace("temp_output_", "")
    nombre_final = f"Libro_Procesado_{timestamp}.docx"
    ruta_final = os.path.join(carpeta_resultados, nombre_final)
    
    try:
        shutil.move(docx_path, ruta_final)
        st.info(f"📁 Documento guardado en: `{ruta_final}`")
        return ruta_final
    except Exception as e:
        st.warning(f"⚠️ No se pudo mover a resultados: {str(e)}")
        return docx_path

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

def preservar_tablas(doc_original):
    """
    Extrae tablas del documento original con su estructura y contenido.
    Retorna lista de (tabla_index, tabla_datos).
    """
    tablas_data = []
    try:
        for tabla_idx, tabla in enumerate(doc_original.tables):
            tabla_contenido = []
            for fila in tabla.rows:
                fila_contenido = []
                for celda in fila.cells:
                    fila_contenido.append(celda.text.strip())
                tabla_contenido.append(fila_contenido)
            tablas_data.append((tabla_idx, tabla_contenido))
    except Exception as e:
        st.warning(f"⚠️ No se pudieron extraer todas las tablas: {str(e)}")
    
    return tablas_data

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
# 3. DETECCIÓN DE IDIOMA
# =====================================================================
def detectar_idioma_muestra(texto_muestra, groq_api_key):
    """
    Detecta el idioma de una muestra de texto usando Groq.
    Retorna el nombre del idioma en español (ej: 'inglés', 'francés', 'portugués').
    """
    try:
        cliente = Groq(api_key=groq_api_key)
        respuesta = cliente.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {
                    "role": "system",
                    "content": "Responde ÚNICAMENTE con el nombre del idioma en español (ej. 'portugués', 'alemán', 'inglés', 'francés') del siguiente texto. Una sola palabra."
                },
                {"role": "user", "content": texto_muestra[:300]}
            ],
            temperature=0,
            max_tokens=15
        )
        idioma_detectado = respuesta.choices[0].message.content.strip().lower()
        return idioma_detectado
    except Exception as e:
        st.warning(f"⚠️ No se pudo detectar idioma: {str(e)}")
        return "desconocido"

# =====================================================================
# 4. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ) - CON BATCH PROCESSING
# =====================================================================
def llamar_groq_con_reintento(texto_lote, groq_api_key, idioma_origen="inglés", max_reintentos=3):
    """
    Llama a Groq con backoff exponencial inteligente.
    - Intento 0: espera 10 segundos
    - Intento 1: espera 20 segundos
    - Intento 2: espera 40 segundos
    Maneja rate limits (429) y otros errores diferenciadamente.
    
    Parámetros:
    - idioma_origen: idioma detectado del PDF (ej: 'inglés', 'francés', 'portugués')
    """
    cliente = Groq(api_key=groq_api_key)
    
    # Construcción dinámicas del prompt según idioma origen
    if idioma_origen.lower() == "español":
        instruccion_traduccion = "Mantén el texto en ESPAÑOL. Solo corrige ortografía, elimina basura de OCR."
    else:
        instruccion_traduccion = f"Traduce del {idioma_origen} al ESPAÑOL de forma natural."
    
    for intento in range(max_reintentos):
        try:
            respuesta = cliente.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "Eres un editor editorial experto en restauración de textos escaneados.\n"
                            f"{instruccion_traduccion}\n"
                            "Se te pasarán múltiples bloques de texto separados por exactamente: '<<BLOCK_SEPARATOR>>'\n"
                            "Para CADA bloque:\n"
                            "1. Realiza la traducción/corrección.\n"
                            "2. Elimina basura de escaneo: símbolos sin sentido o sílabas rotas.\n"
                            "3. Corrige la ortografía y puntuación.\n"
                            "Devuelve cada bloque separado por exactamente: '<<BLOCK_SEPARATOR>>'\n"
                            "IMPORTANTE: Mantén el mismo número de bloques. Sin introducciones ni explicaciones."
                        )
                    },
                    {"role": "user", "content": texto_lote}
                ],
                temperature=0.1,
                max_tokens=3000
            )
            return respuesta.choices[0].message.content.strip()
            
        except Exception as e:
            error_msg = str(e).lower()
            es_rate_limit = "rate limit" in error_msg or "429" in error_msg
            
            if es_rate_limit and intento < max_reintentos - 1:
                # Backoff exponencial: 2^intento * 10 (10, 20, 40 segundos)
                tiempo_espera = (2 ** intento) * 10
                st.warning(f"⏳ Rate limit detectado. Esperando {tiempo_espera}s... (Intento {intento + 1}/{max_reintentos})")
                time.sleep(tiempo_espera)
            else:
                # Error no-rate-limit o último intento agotado
                if es_rate_limit:
                    st.warning(f"❌ Rate limit persistente tras {max_reintentos} intentos. Se mantienen originales.")
                else:
                    st.warning(f"⚠️ Error en API Groq: {str(e)[:100]}. Se mantienen originales.")
                return None
    
    return None

def traducir_lote(textos_lote, groq_api_key, idioma_origen="inglés"):
    """
    Traduce una lista de textos usando Groq.
    Retorna una lista de textos traducidos (o originales si falla).
    """
    if not textos_lote or all(not t.strip() for t in textos_lote):
        return textos_lote
    
    # Unir textos con delimitador único y seguro
    DELIMITER = "\n<<BLOCK_SEPARATOR>>\n"
    texto_combinado = DELIMITER.join(textos_lote)
    
    resultado = llamar_groq_con_reintento(texto_combinado, groq_api_key, idioma_origen=idioma_origen, max_reintentos=3)
    
    if resultado is None:
        # Retornar textos originales sin modificar
        return textos_lote
    
    # Dividir resultado manteniendo el orden
    traducidos = resultado.split(DELIMITER)
    
    # Si la división falla (delimitador no se preservó), retornar originales
    if len(traducidos) != len(textos_lote):
        st.warning(f"⚠️ Integridad de separadores comprometida. Se mantienen originales.")
        return textos_lote
    
    return [t.strip() for t in traducidos]

def obtener_estilos_validos(doc):
    """
    Obtiene lista de estilos válidos en el documento.
    Incluye fallback a 'Normal' si un estilo no existe.
    """
    try:
        estilos_validos = {s.name for s in doc.styles if s.type == 1}  # type=1 es párrafo
        if 'Normal' not in estilos_validos:
            estilos_validos.add('Normal')
        return estilos_validos
    except Exception as e:
        st.warning(f"⚠️ No se pudieron obtener estilos válidos: {str(e)}")
        return {'Normal'}

def procesar_docx_multilingue(docx_path, docx_salida_path, groq_api_key, idioma_origen="inglés", tamano_lote=10):
    """
    Lee el DOCX original, traduce párrafos en lotes al español,
    y guarda un nuevo DOCX limpio.
    
    MEJORAS IMPLEMENTADAS:
    - Validación de estilos para evitar crasheos
    - Preservación de tablas
    - Batch processing con reintentos
    - Guardado incremental y seguro
    - Manejo robusto de errores
    
    Parámetros:
    - idioma_origen: idioma detectado automáticamente del PDF
    - tamano_lote: número de párrafos por lote (default: 10)
    """
    doc_original = docx.Document(docx_path)
    doc_nuevo = docx.Document()  # Documento limpio nuevo
    
    # 1. Purgar imágenes pequeñas del doc original
    texto_estado = st.empty()
    texto_estado.text("Limpiando imágenes y artefactos de escaneo...")
    
    img_eliminadas = limpiar_imagenes_pequenas(doc_original, min_width_cm=1.5, min_height_cm=1.5)
    st.info(f"🧹 Se eliminaron {img_eliminadas} artefactos visuales.")
    
    # 2. Extraer tablas ANTES de procesar (preservación)
    tablas_datos = preservar_tablas(doc_original)
    if tablas_datos:
        st.info(f"📊 Se detectaron {len(tablas_datos)} tabla(s) que serán preservadas.")
    
    # 🔑 NOVEDAD: Obtener los estilos válidos del documento ANTES del bucle
    estilos_validos = obtener_estilos_validos(doc_nuevo)
    
    # 3. Extraer párrafos válidos
    parrafos_datos = []  # Lista de (texto_original, estilo)
    
    for p in doc_original.paragraphs:
        texto = p.text.strip()
        estilo = p.style.name if p.style else 'Normal'
        
        if texto and not texto.isdigit():  # Solo párrafos con contenido
            # Pre-limpiar OCR
            texto_limpio = pre_limpiar_ocr(texto)
            if len(texto_limpio) > 3:
                parrafos_datos.append((texto_limpio, estilo))
        else:
            # Preservar párrafo vacío para mantener espaciado
            parrafos_datos.append(("", estilo))
    
    if not parrafos_datos:
        st.info("No hay párrafos válidos para procesar.")
        doc_nuevo.save(docx_salida_path)
        return
    
    # 4. Procesar en lotes
    total = len(parrafos_datos)
    barra_progreso = st.progress(0)
    
    for i in range(0, total, tamano_lote):
        lote_data = parrafos_datos[i:i+tamano_lote]
        lote_numero = (i // tamano_lote) + 1
        total_lotes = (total + tamano_lote - 1) // tamano_lote
        
        # Extraer solo textos del lote
        textos_lote = [t[0] for t in lote_data]
        
        # Traducir lote (pasando idioma origen)
        texto_estado.text(f"Traduc. lote {lote_numero}/{total_lotes} (desde {idioma_origen})...")
        textos_traducidos = traducir_lote(textos_lote, groq_api_key, idioma_origen=idioma_origen)
        
        # 🔑 VALIDACIÓN CLAVE: Aplicar al documento nuevo con estilos seguros
        for j, (texto_orig, estilo) in enumerate(lote_data):
            texto_final = textos_traducidos[j] if j < len(textos_traducidos) else texto_orig
            
            # VALIDACIÓN: Si el estilo de Adobe no existe, usa 'Normal' para evitar daños
            estilo_seguro = estilo if estilo in estilos_validos else 'Normal'
            
            if texto_final.strip():  # Solo agregar si hay contenido
                doc_nuevo.add_paragraph(texto_final, style=estilo_seguro)
            else:
                # Preservar párrafo vacío
                doc_nuevo.add_paragraph("", style=estilo_seguro)
        
        # Actualizar barra
        progreso = min(i + tamano_lote, total)
        barra_progreso.progress(progreso / total)
        
        # Guardado incremental y seguro (cada 2 lotes)
        if (lote_numero % 2 == 0) or (i + tamano_lote >= total):
            temp_path = docx_salida_path + ".tmp"
            try:
                doc_nuevo.save(temp_path)
                os.replace(temp_path, docx_salida_path)  # Reemplazo atómico
                texto_estado.text(f"💾 Guardado en lote {lote_numero}...")
            except Exception as e:
                st.warning(f"⚠️ Error al guardar lote {lote_numero}: {str(e)}")
        
        time.sleep(0.5)  # Pausa entre lotes
    
    # Guardado final
    try:
        doc_nuevo.save(docx_salida_path)
        texto_estado.text("✅ Traducción completada exitosamente.")
    except Exception as e:
        st.error(f"❌ Error al guardar documento final: {str(e)}")
    
    barra_progreso.empty()

# =====================================================================
# 5. INTERFAZ DE USUARIO Y CONTROL DE FLUJO PRINCIPAL
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
        exito_total = False
        
        try:
            with open(temp_pdf, "wb") as f:
                f.write(archivo_subido.getbuffer())
                
            with st.spinner("Fase 1/2: Convirtiendo estructura del PDF a Word en servidores de Adobe..."):
                exito_adobe = convertir_pdf_a_word_adobe(
                    temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
                )
                
            if exito_adobe:
                # Detectar idioma del documento
                with st.spinner("🔍 Detectando idioma del documento..."):
                    doc_temp = docx.Document(temp_docx)
                    # Extraer muestra de texto (primeros párrafos no vacíos)
                    texto_muestra = ""
                    for p in doc_temp.paragraphs[:10]:
                        if p.text.strip():
                            texto_muestra += p.text.strip() + " "
                            if len(texto_muestra) > 300:
                                break
                    
                    if texto_muestra:
                        idioma_detectado = detectar_idioma_muestra(texto_muestra, GROQ_API_KEY)
                        st.info(f"🌍 Idioma detectado: **{idioma_detectado.capitalize()}**")
                    else:
                        idioma_detectado = "inglés"
                        st.warning("⚠️ No se pudo detectar idioma. Asumiendo inglés.")
                
                with st.spinner("Fase 2/2: Traduciendo a español y limpiando ruido de OCR..."):
                    procesar_docx_multilingue(
                        docx_path=temp_docx,
                        docx_salida_path=temp_docx,  # Sobrescribe con documento limpio
                        groq_api_key=GROQ_API_KEY,
                        idioma_origen=idioma_detectado,
                        tamano_lote=10
                    )
                    
                st.success("🎉 ¡El documento ha sido procesado y restaurado con éxito!")
                st.balloons()
                exito_total = True
                
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
            # SIEMPRE eliminar PDF temporal
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            
            # Gestión inteligente del DOCX
            if exito_total and os.path.exists(temp_docx):
                # ✅ Éxito: Mover a carpeta de resultados
                docx_final = mover_docx_a_resultados(temp_docx)
                st.info(f"📄 Archivo disponible en: `{docx_final}`")
            elif os.path.exists(temp_docx):
                # ⚠️ Fallo: Mantener DOCX para inspección en carpeta debug
                carpeta_debug = "debug_fallos"
                if not os.path.exists(carpeta_debug):
                    os.makedirs(carpeta_debug)
                ruta_debug = os.path.join(carpeta_debug, f"error_{id_unico}.docx")
                try:
                    shutil.copy(temp_docx, ruta_debug)
                    st.warning(f"🔍 Documento de debug guardado en: `{ruta_debug}` para inspección")
                except Exception:
                    pass
                # Eliminar el temp después de copiar
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass
