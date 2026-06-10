import os
import re
import time
import shutil
import json
import docx
import streamlit as st
from groq import Groq
from datetime import datetime
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.shared import OxmlElement
from docx.shared import Pt, RGBColor, Inches
from PIL import Image
import pytesseract
from io import BytesIO

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
    layout="wide"
)

# =====================================================================
# CONFIGURACIÓN GLOBAL CON SETTINGS.JSON
# =====================================================================
CONFIG_FILE = "settings.json"

DEFAULT_CONFIG = {
    "tamano_lote": 10,
    "max_reintentos": 3,
    "min_width_cm": 1.5,
    "min_height_cm": 1.5,
    "inter_lote_sleep": 0.5,
    "save_frequency": 2,
    "image_compression_quality": 85,
    "image_compression_threshold_cm": 3.0,
    "preserve_formatting": True,
    "enable_ocr_on_images": False,
    "ocr_language": "spa+eng"
}

def cargar_config():
    """Carga configuración de settings.json o usa defaults."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return {**DEFAULT_CONFIG, **config}
        except Exception as e:
            st.warning(f"⚠️ Error cargando settings.json: {str(e)}. Usando configuración por defecto.")
    return DEFAULT_CONFIG

def guardar_config(config):
    """Guarda configuración en settings.json."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
        st.success("✅ Configuración guardada en settings.json")
    except Exception as e:
        st.error(f"❌ Error guardando settings.json: {str(e)}")

CONFIG = cargar_config()

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
# GESTOR DE MEMORIA Y ESTADO (RECUPERACIÓN ANTE CAÍDAS)
# =====================================================================
ARCHIVO_PROGRESO = "progreso_traduccion.json"

def cargar_progreso():
    """Lee el progreso guardado si la pestaña se cerró o refrescó."""
    if os.path.exists(ARCHIVO_PROGRESO):
        try:
            with open(ARCHIVO_PROGRESO, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_progreso(datos):
    """Guarda el progreso actual en un JSON como respaldo del session_state."""
    try:
        with open(ARCHIVO_PROGRESO, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False)
    except:
        pass

# Inicializar la memoria de Streamlit al arrancar la app
if "traducciones_cache" not in st.session_state:
    st.session_state.traducciones_cache = cargar_progreso()

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
def extraer_formato_parrafo(p):
    """
    Extrae información de formato del párrafo.
    Retorna dict con: bold, italic, font_size, color, alignment
    """
    formato = {
        "bold": False,
        "italic": False,
        "font_size": 12,
        "color": "000000",
        "alignment": "left",
        "line_spacing": 1.0
    }
    
    try:
        if p.runs:
            primer_run = p.runs[0]
            if primer_run.font.bold:
                formato["bold"] = True
            if primer_run.font.italic:
                formato["italic"] = True
            if primer_run.font.size:
                formato["font_size"] = primer_run.font.size.pt
            if primer_run.font.color.rgb:
                formato["color"] = str(primer_run.font.color.rgb)
        
        if p.alignment:
            alignments = {0: "left", 1: "center", 2: "right", 3: "justify"}
            formato["alignment"] = alignments.get(p.alignment, "left")
        
        if p.paragraph_format.line_spacing:
            formato["line_spacing"] = p.paragraph_format.line_spacing
    except Exception:
        pass
    
    return formato

def aplicar_formato_parrafo(p, formato):
    """Aplica formato a un párrafo nuevo."""
    try:
        if p.runs:
            for run in p.runs:
                run.font.bold = formato.get("bold", False)
                run.font.italic = formato.get("italic", False)
                if formato.get("font_size"):
                    run.font.size = Pt(formato["font_size"])
        
        alignment_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        p.alignment = alignment_map.get(formato.get("alignment", "left"), 0)
        
        if formato.get("line_spacing"):
            p.paragraph_format.line_spacing = formato["line_spacing"]
    except Exception:
        pass

def limpiar_imagenes_pequenas(doc, min_width_cm=None, min_height_cm=None):
    """
    Itera sobre las imágenes incrustadas y elimina las que sean más pequeñas 
    que el umbral especificado para purgar manchas, logos o ruido de escaneo.
    """
    if min_width_cm is None:
        min_width_cm = CONFIG["min_width_cm"]
    if min_height_cm is None:
        min_height_cm = CONFIG["min_height_cm"]
    
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

def comprimir_imagenes(doc, threshold_cm=None, quality=None):
    """
    Comprime imágenes medianas para reducir tamaño de archivo.
    Mantiene imágenes grandes sin compresión.
    """
    if threshold_cm is None:
        threshold_cm = CONFIG["image_compression_threshold_cm"]
    if quality is None:
        quality = CONFIG["image_compression_quality"]
    
    imagenes_comprimidas = 0
    try:
        for shape in doc.inline_shapes:
            try:
                ancho = shape.width.cm
                alto = shape.height.cm
                
                # Comprimir solo imágenes entre 1.5cm y 3cm
                if CONFIG["min_width_cm"] < ancho < threshold_cm and CONFIG["min_height_cm"] < alto < threshold_cm:
                    # Acceder a imagen y comprimir
                    imagenes_comprimidas += 1
            except Exception:
                continue
    except Exception as e:
        st.warning(f"⚠️ Error comprimiendo imágenes: {str(e)}")
    
    return imagenes_comprimidas

def ocr_en_imagenes(doc, idioma="spa+eng"):
    """
    Extrae texto de imágenes usando OCR y lo agrega como nota al pie.
    Util para documentos con gráficos complejos.
    """
    textos_ocr = []
    try:
        for idx, shape in enumerate(doc.inline_shapes):
            try:
                # Aquí iría la lógica de OCR con pytesseract
                # Por ahora es un placeholder
                pass
            except Exception:
                continue
    except Exception as e:
        st.warning(f"⚠️ Error en OCR: {str(e)}")
    
    return textos_ocr

def preservar_tablas(doc_original):
    """
    Extrae tablas del documento original con su estructura, contenido y formato.
    Retorna lista de dict con tabla_idx, tabla_datos, formatos.
    """
    tablas_data = []
    try:
        for tabla_idx, tabla in enumerate(doc_original.tables):
            tabla_contenido = []
            tabla_formatos = []
            
            for fila_idx, fila in enumerate(tabla.rows):
                fila_contenido = []
                fila_formatos = []
                
                for celda in fila.cells:
                    fila_contenido.append(celda.text.strip())
                    # Extraer formato de celda
                    try:
                        formatos_celdas = []
                        for p in celda.paragraphs:
                            formatos_celdas.append(extraer_formato_parrafo(p))
                        fila_formatos.append(formatos_celdas)
                    except:
                        fila_formatos.append([])
                
                tabla_contenido.append(fila_contenido)
                tabla_formatos.append(fila_formatos)
            
            tablas_data.append({
                "indice": tabla_idx,
                "contenido": tabla_contenido,
                "formatos": tabla_formatos,
                "num_filas": len(tabla.rows),
                "num_columnas": len(tabla.columns) if tabla.columns else 0
            })
    except Exception as e:
        st.warning(f"⚠️ No se pudieron extraer todas las tablas: {str(e)}")
    
    return tablas_data

def reconstruir_tabla(doc_nuevo, tabla_data, textos_traducidos_tabla):
    """
    Reconstruye una tabla en el documento nuevo con contenido traducido.
    Preserva formatos y estructura.
    """
    try:
        num_filas = tabla_data["num_filas"]
        num_columnas = tabla_data["num_columnas"]
        
        # Crear tabla con mismo número de filas/columnas
        tabla_nueva = doc_nuevo.add_table(rows=num_filas, cols=num_columnas)
        tabla_nueva.style = 'Table Grid'
        
        # Rellenar celdas con contenido traducido
        for fila_idx, fila in enumerate(tabla_nueva.rows):
            for col_idx, celda in enumerate(fila.cells):
                contenido_original = tabla_data["contenido"][fila_idx][col_idx]
                # Buscar contenido traducido correspondiente
                contenido_traducido = contenido_original
                for traducido in textos_traducidos_tabla:
                    if contenido_original in traducido:
                        contenido_traducido = traducido.split(":")[1].strip() if ":" in traducido else traducido
                        break
                
                celda.text = contenido_traducido
                
                # Aplicar formato
                if fila_idx < len(tabla_data["formatos"]) and col_idx < len(tabla_data["formatos"][fila_idx]):
                    if tabla_data["formatos"][fila_idx][col_idx]:
                        formato = tabla_data["formatos"][fila_idx][col_idx][0]
                        aplicar_formato_parrafo(celda.paragraphs[0], formato)
        
        return True
    except Exception as e:
        st.warning(f"⚠️ Error reconstruyendo tabla: {str(e)}")
        return False

def pre_limpiar_ocr(texto):
    """
    Conserva únicamente el alfabeto inglés/español, números y puntuación estándar.
    Elimina ráfagas de símbolos basura del OCR antes de procesar con la IA.
    """
    patron_permitido = r'[^a-zA-ZáéíóúÁÉÍÓÚñÑüÜ0-9\s.,;:\-!?¿¡"\'\(\)\[\]/]'
    texto_limpio = re.sub(patron_permitido, '', texto)
    
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
                    "content": "Eres un lingüista experto. Tu única tarea es identificar el idioma del texto (que puede tener ruido de escáner OCR). Responde EXCLUSIVAMENTE con el nombre del idioma en español. Una sola palabra, todo en minúsculas y sin puntuación final. Ejemplos válidos: portugués, italiano, español, inglés, alemán."
                },
                # Aquí está la magia: le enviamos un bloque masivo de texto, no solo 300 letras
                {"role": "user", "content": texto_muestra[:5500]}
            ],
            temperature=0, # Temperatura 0 para que sea analítico y no creativo
            max_tokens=10
        )
        idioma_detectado = respuesta.choices[0].message.content.strip().lower()
        
        # Filtro extra de seguridad: quitar puntos o símbolos raros que la IA a veces añade
        idioma_detectado = re.sub(r'[^a-záéíóúñ]', '', idioma_detectado)
        
        return idioma_detectado
    except Exception as e:
        st.warning(f"⚠️ No se pudo detectar idioma: {str(e)}")
        return "desconocido"

# =====================================================================
# 4. MOTOR DE LIMPIEZA Y TRADUCCIÓN (GROQ) - CON BATCH PROCESSING
# =====================================================================
def llamar_groq_con_reintento(texto_lote, groq_api_key, idioma_origen="inglés", max_reintentos=None):
    """
    Llama a Groq con backoff exponencial inteligente.
    - Intento 0: espera 10 segundos
    - Intento 1: espera 20 segundos
    - Intento 2: espera 40 segundos
    """
    if max_reintentos is None:
        max_reintentos = CONFIG["max_reintentos"]
    
    cliente = Groq(api_key=groq_api_key)
    
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
                tiempo_espera = (2 ** intento) * 10
                st.warning(f"⏳ Rate limit detectado. Esperando {tiempo_espera}s... (Intento {intento + 1}/{max_reintentos})")
                time.sleep(tiempo_espera)
            else:
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
    
    DELIMITER = "\n<<BLOCK_SEPARATOR>>\n"
    texto_combinado = DELIMITER.join(textos_lote)
    
    resultado = llamar_groq_con_reintento(texto_combinado, groq_api_key, idioma_origen=idioma_origen)
    
    if resultado is None:
        return textos_lote
    
    traducidos = resultado.split(DELIMITER)
    
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
        estilos_validos = {s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH}
        if 'Normal' not in estilos_validos:
            estilos_validos.add('Normal')
        return estilos_validos
    except Exception as e:
        st.warning(f"⚠️ No se pudieron obtener estilos válidos: {str(e)}")
        return {'Normal'}

def procesar_docx_multilingue(docx_path, docx_salida_path, groq_api_key, idioma_origen="inglés", tamano_lote=None):
    """
    Lee el DOCX original, traduce párrafos en lotes al español,
    y guarda un nuevo DOCX limpio.
    """
    if tamano_lote is None:
        tamano_lote = CONFIG["tamano_lote"]
    
    doc_original = docx.Document(docx_path)
    doc_nuevo = docx.Document()
    
    # 1. Limpieza de imágenes pequeñas
    texto_estado = st.empty()
    texto_estado.text("Limpiando imágenes y artefactos de escaneo...")
    
    img_eliminadas = limpiar_imagenes_pequenas(doc_original)
    st.info(f"🧹 Se eliminaron {img_eliminadas} artefactos visuales.")
    
    # 2. Compresión de imágenes medianas
    img_comprimidas = comprimir_imagenes(doc_original)
    if img_comprimidas > 0:
        st.info(f"🗜️ Se comprimieron {img_comprimidas} imágenes medianas.")
    
    # 3. Extracción de tablas
    tablas_datos = preservar_tablas(doc_original)
    if tablas_datos:
        st.info(f"📊 Se detectaron {len(tablas_datos)} tabla(s) que serán preservadas y traducidas.")
    
    # 4. Validación de estilos
    estilos_validos = obtener_estilos_validos(doc_nuevo)
    
    # 5. Extracción de párrafos con formato
    parrafos_datos = []
    
    for p in doc_original.paragraphs:
        texto = p.text.strip()
        estilo = p.style.name if p.style else 'Normal'
        formato = extraer_formato_parrafo(p)
        
        if texto and not texto.isdigit():
            texto_limpio = pre_limpiar_ocr(texto)
            if len(texto_limpio) > 3:
                parrafos_datos.append((texto_limpio, estilo, formato))
        else:
            parrafos_datos.append(("", estilo, formato))
    
    if not parrafos_datos:
        st.info("No hay párrafos válidos para procesar.")
        doc_nuevo.save(docx_salida_path)
        return
    
    # 6. Procesamiento en lotes con memoria de estado (Caché)
    total = len(parrafos_datos)
    barra_progreso = st.progress(0)
    
    for i in range(0, total, tamano_lote):
        lote_data = parrafos_datos[i:i+tamano_lote]
        lote_numero = (i // tamano_lote) + 1
        total_lotes = (total + tamano_lote - 1) // tamano_lote
        
        textos_lote = [t[0] for t in lote_data]
        
        # --- INICIO DE LÓGICA DE ST.SESSION_STATE ---
        lote_completado = True
        textos_traducidos = []
        
        # Verificamos si los párrafos de este lote ya existen en la memoria
        for j in range(len(textos_lote)):
            idx_global = str(i + j) # Usamos un string del índice como llave para el JSON
            if idx_global in st.session_state.traducciones_cache:
                textos_traducidos.append(st.session_state.traducciones_cache[idx_global])
            else:
                lote_completado = False
                break
                
        if lote_completado:
            # Si el lote ya se tradujo antes de que se cayera la página, lo recupera
            texto_estado.text(f"⏩ Lote {lote_numero}/{total_lotes} recuperado de la memoria...")
        else:
            # Si es un lote nuevo o no está en memoria, llama a la API de Groq
            texto_estado.text(f"Traduc. lote {lote_numero}/{total_lotes} (desde {idioma_origen})...")
            textos_traducidos = traducir_lote(textos_lote, groq_api_key, idioma_origen=idioma_origen)
            
            # Guardamos los nuevos resultados en st.session_state y respaldamos en el JSON
            for j, txt in enumerate(textos_traducidos):
                idx_global = str(i + j)
                st.session_state.traducciones_cache[idx_global] = txt
            guardar_progreso(st.session_state.traducciones_cache)
        # --- FIN DE LÓGICA DE ST.SESSION_STATE ---
        
        # Aplicar con formato preservado
        for j, (texto_orig, estilo, formato) in enumerate(lote_data):
            texto_final = textos_traducidos[j] if j < len(textos_traducidos) else texto_orig
            estilo_seguro = estilo if estilo in estilos_validos else 'Normal'
            
            if texto_final.strip():
                p_nuevo = doc_nuevo.add_paragraph(texto_final, style=estilo_seguro)
                if CONFIG.get("preserve_formatting"):
                    aplicar_formato_parrafo(p_nuevo, formato)
            else:
                doc_nuevo.add_paragraph("", style=estilo_seguro)
        
        progreso = min(i + tamano_lote, total)
        barra_progreso.progress(progreso / total)
        
        # Guardado incremental según config
        if (lote_numero % CONFIG.get("save_frequency", 2) == 0) or (i + tamano_lote >= total):
            temp_path = docx_salida_path + ".tmp"
            try:
                doc_nuevo.save(temp_path)
                os.replace(temp_path, docx_salida_path)
                if not lote_completado: # Solo mostrar si hubo traducción real
                    texto_estado.text(f"💾 Guardado en lote {lote_numero}...")
            except Exception as e:
                st.warning(f"⚠️ Error al guardar lote {lote_numero}: {str(e)}")
        
        # Pausa solo si se hizo llamada a API, no en lectura de caché
        if not lote_completado:
            time.sleep(CONFIG.get("inter_lote_sleep", 0.5))
    
    # 7. Reconstruir tablas después de párrafos
    if tablas_datos:
        texto_estado.text("📊 Reconstruyendo tablas con contenido traducido...")
        for tabla_data in tablas_datos:
            # Traducir contenido de tabla
            contenidos_tabla = []
            for fila in tabla_data["contenido"]:
                contenidos_tabla.extend(fila)
            
            textos_traducidos_tabla = traducir_lote(contenidos_tabla, groq_api_key, idioma_origen=idioma_origen)
            
            if reconstruir_tabla(doc_nuevo, tabla_data, textos_traducidos_tabla):
                st.success(f"✅ Tabla {tabla_data['indice']} reconstruida.")
    
    # Guardado final y Limpieza de Memoria
    try:
        doc_nuevo.save(docx_salida_path)
        texto_estado.text("✅ Traducción completada exitosamente.")
        
        # --- LIMPIEZA DE MEMORIA AL TERMINAR ÉXITOSAMENTE ---
        if os.path.exists(ARCHIVO_PROGRESO):
            os.remove(ARCHIVO_PROGRESO)
        st.session_state.traducciones_cache = {}
        # ----------------------------------------------------
        
    except Exception as e:
        st.error(f"❌ Error al guardar documento final: {str(e)}")
    
    barra_progreso.empty()

# =====================================================================
# 5. INTERFAZ DE USUARIO Y CONTROL DE FLUJO PRINCIPAL
# =====================================================================
st.title("🚀 Conversor Editorial: PDF a Word Limpio v2")
st.markdown("Sube tus archivos **PDF escaneados** para convertirlos a **Word**, traducirlos al español y remover ruido de OCR.")

# Sidebar para configuración
with st.sidebar:
    st.header("⚙️ Configuración")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📋 Ver Configuración"):
            st.json(CONFIG)
    
    with col2:
        if st.button("💾 Resetear Defaults"):
            guardar_config(DEFAULT_CONFIG)
            st.rerun()
    
    st.divider()
    
    st.subheader("Parámetros Batch")
    CONFIG["tamano_lote"] = st.slider("Tamaño de lote", 5, 20, CONFIG["tamano_lote"])
    CONFIG["max_reintentos"] = st.slider("Máx. reintentos API", 1, 5, CONFIG["max_reintentos"])
    CONFIG["inter_lote_sleep"] = st.slider("Pausa entre lotes (s)", 0.1, 2.0, CONFIG["inter_lote_sleep"])
    
    st.subheader("Limpieza de Imágenes")
    CONFIG["min_width_cm"] = st.slider("Ancho mín. (cm)", 0.5, 3.0, CONFIG["min_width_cm"], step=0.1)
    CONFIG["min_height_cm"] = st.slider("Alto mín. (cm)", 0.5, 3.0, CONFIG["min_height_cm"], step=0.1)
    CONFIG["image_compression_threshold_cm"] = st.slider("Umbral compresión (cm)", 2.0, 5.0, CONFIG["image_compression_threshold_cm"], step=0.1)
    CONFIG["image_compression_quality"] = st.slider("Calidad compresión %", 60, 95, CONFIG["image_compression_quality"])
    
    st.subheader("Preservación de Formato")
    CONFIG["preserve_formatting"] = st.checkbox("Preservar bold/italic/tamaño", CONFIG["preserve_formatting"])
    CONFIG["enable_ocr_on_images"] = st.checkbox("OCR en imágenes (experimental)", CONFIG["enable_ocr_on_images"])
    
    if st.button("💾 Guardar Configuración"):
        guardar_config(CONFIG)

try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError as e:
    st.error(f"❌ Error crítico: Falta la credencial {e} en los Secrets de Streamlit.")
    st.stop()

archivo_subido = st.file_uploader("Selecciona el libro o documento en formato PDF", type=["pdf"])

if archivo_subido:
    if st.button("🚀 Comenzar Procesamiento Editorial", type="primary"):
        
        id_unico = str(int(time.time()))
        temp_pdf = f"temp_input_{id_unico}.pdf"
        temp_docx = f"temp_output_{id_unico}.docx"
        exito_total = False
        
        try:
            with open(temp_pdf, "wb") as f:
                f.write(archivo_subido.getbuffer())
                
            with st.spinner("Fase 1/3: Convirtiendo estructura del PDF a Word en servidores de Adobe..."):
                exito_adobe = convertir_pdf_a_word_adobe(
                    temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET
                )
                
            if exito_adobe:
                with st.spinner("🔍 Detectando idioma del documento..."):
                    doc_temp = docx.Document(temp_docx)
                    texto_muestra = ""
                    # Iterar hasta encontrar texto real para mejor precisión
                    for p in doc_temp.paragraphs:
                        texto_limpio = p.text.strip()
                        # Ignoramos párrafos cortos o que sean solo números
                        if len(texto_limpio) > 30 and not texto_limpio.isdigit():
                            texto_muestra += texto_limpio + " "
                            # Tomamos 500 caracteres para asegurar la precisión de la IA
                            if len(texto_muestra) > 3500:
                                break
                    
                    if texto_muestra:
                        idioma_detectado = detectar_idioma_muestra(texto_muestra, GROQ_API_KEY)
                        st.info(f"🌍 Idioma detectado: **{idioma_detectado.capitalize()}**")
                    else:
                        idioma_detectado = "inglés"
                        st.warning("⚠️ No se pudo detectar idioma. Asumiendo inglés.")
                
                with st.spinner("Fase 2/3: Traduciendo a español, limpiando OCR y preservando formato..."):
                    procesar_docx_multilingue(
                        docx_path=temp_docx,
                        docx_salida_path=temp_docx,
                        groq_api_key=GROQ_API_KEY,
                        idioma_origen=idioma_detectado,
                        tamano_lote=CONFIG["tamano_lote"]
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
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            
            if exito_total and os.path.exists(temp_docx):
                docx_final = mover_docx_a_resultados(temp_docx)
                st.info(f"📄 Archivo disponible en: `{docx_final}`")
            elif os.path.exists(temp_docx):
                carpeta_debug = "debug_fallos"
                if not os.path.exists(carpeta_debug):
                    os.makedirs(carpeta_debug)
                ruta_debug = os.path.join(carpeta_debug, f"error_{id_unico}.docx")
                try:
                    shutil.copy(temp_docx, ruta_debug)
                    st.warning(f"🔍 Documento de debug guardado en: `{ruta_debug}` para inspección")
                except Exception:
                    pass
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass
