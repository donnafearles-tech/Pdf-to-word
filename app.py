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
from io import BytesIO

# =====================================================================
# IMPORTACIÓN OPCIONAL DE PYTESSERACT (OCR)
# =====================================================================
try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    pytesseract = None
    PYTESSERACT_AVAILABLE = False

# =====================================================================
# IMPORTACIÓN OPCIONAL DE LANGDETECT (DETECCIÓN DE IDIOMA MÁS PRECISA)
# =====================================================================
try:
    from langdetect import detect
    # Nota: No usamos DetectorFactory.seed porque causa TypeError
    LANGDETECT_AVAILABLE = True
except ImportError:
    LANGDETECT_AVAILABLE = False

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
    "tamano_lote": 5,                # Reducido de 10 a 5 para respetar rate limit
    "max_reintentos": 5,             # Aumentado de 3 a 5
    "min_width_cm": 1.5,
    "min_height_cm": 1.5,
    "inter_lote_sleep": 3.0,         # Aumentado de 0.5 a 3 segundos
    "save_frequency": 2,
    "image_compression_quality": 85,
    "image_compression_threshold_cm": 3.0,
    "preserve_formatting": True,
    "enable_ocr_on_images": False,
    "ocr_language": "spa+eng",
    "groq_model": "llama-3.3-70b-versatile"  # Nuevo modelo con 12,000 TPM
}

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return {**DEFAULT_CONFIG, **config}
        except Exception:
            pass
    return DEFAULT_CONFIG

def guardar_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
        st.success("✅ Configuración guardada")
    except Exception as e:
        st.error(f"❌ Error guardando: {str(e)}")

CONFIG = cargar_config()

# =====================================================================
# GESTIÓN DE RESULTADOS PERSISTENTES Y REANUDACIÓN
# =====================================================================
if "resultado" not in st.session_state:
    st.session_state.resultado = None

if "traducciones_cache" not in st.session_state:
    st.session_state.traducciones_cache = {}

ARCHIVO_PROGRESO = "progreso_traduccion.json"
ARCHIVO_ULTIMO_LOTE = "ultimo_lote_completado.txt"

def cargar_progreso():
    if os.path.exists(ARCHIVO_PROGRESO):
        try:
            with open(ARCHIVO_PROGRESO, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_progreso(datos):
    try:
        with open(ARCHIVO_PROGRESO, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False)
    except:
        pass

def guardar_ultimo_lote(numero):
    try:
        with open(ARCHIVO_ULTIMO_LOTE, 'w') as f:
            f.write(str(numero))
    except:
        pass

def cargar_ultimo_lote():
    if os.path.exists(ARCHIVO_ULTIMO_LOTE):
        try:
            with open(ARCHIVO_ULTIMO_LOTE, 'r') as f:
                return int(f.read().strip())
        except:
            return 0
    return 0

def limpiar_progreso():
    if os.path.exists(ARCHIVO_PROGRESO):
        os.remove(ARCHIVO_PROGRESO)
    if os.path.exists(ARCHIVO_ULTIMO_LOTE):
        os.remove(ARCHIVO_ULTIMO_LOTE)

# Cargar progreso previo al inicio
st.session_state.traducciones_cache = cargar_progreso()

# =====================================================================
# FUNCIONES AUXILIARES DE GESTIÓN DE ARCHIVOS
# =====================================================================
def crear_carpeta_resultados():
    carpeta = "resultados"
    os.makedirs(carpeta, exist_ok=True)
    return carpeta

def mover_docx_a_resultados(docx_path):
    carpeta_resultados = crear_carpeta_resultados()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_final = f"Libro_Procesado_{timestamp}.docx"
    ruta_final = os.path.join(carpeta_resultados, nombre_final)
    try:
        shutil.move(docx_path, ruta_final)
        return ruta_final
    except Exception as e:
        st.warning(f"⚠️ No se pudo mover a resultados: {str(e)}")
        return docx_path

# =====================================================================
# 1. MOTOR DE CONVERSIÓN (ADOBE SDK V4)
# =====================================================================
def convertir_pdf_a_word_adobe(input_pdf_path, output_docx_path, client_id, client_secret):
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
        st.error(f"Error fatal en Adobe PDF Services: {str(e)}")
        return False

# =====================================================================
# 2. HEURÍSTICAS DE FORMATO, TABLAS, IMÁGENES, OCR
# =====================================================================
def extraer_formato_parrafo(p):
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
                shape._inline.getparent().remove(shape._inline)
                imagenes_eliminadas += 1
        except Exception:
            continue
    return imagenes_eliminadas

def comprimir_imagenes(doc, threshold_cm=None, quality=None):
    return 0  # Placeholder

def ocr_en_imagenes(doc, idioma="spa+eng"):
    if not PYTESSERACT_AVAILABLE:
        return []
    return []

def preservar_tablas(doc_original):
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
                    try:
                        formatos_celdas = [extraer_formato_parrafo(p) for p in celda.paragraphs]
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
        st.warning(f"⚠️ Error extrayendo tablas: {str(e)}")
    return tablas_data

def reconstruir_tabla(doc_nuevo, tabla_data, textos_traducidos_tabla):
    try:
        num_filas = tabla_data["num_filas"]
        num_columnas = tabla_data["num_columnas"]
        tabla_nueva = doc_nuevo.add_table(rows=num_filas, cols=num_columnas)
        tabla_nueva.style = 'Table Grid'
        for fila_idx, fila in enumerate(tabla_nueva.rows):
            for col_idx, celda in enumerate(fila.cells):
                contenido_original = tabla_data["contenido"][fila_idx][col_idx]
                contenido_traducido = contenido_original
                for traducido in textos_traducidos_tabla:
                    if contenido_original in traducido:
                        contenido_traducido = traducido.split(":")[1].strip() if ":" in traducido else traducido
                        break
                celda.text = contenido_traducido
                if fila_idx < len(tabla_data["formatos"]) and col_idx < len(tabla_data["formatos"][fila_idx]):
                    if tabla_data["formatos"][fila_idx][col_idx]:
                        formato = tabla_data["formatos"][fila_idx][col_idx][0]
                        aplicar_formato_parrafo(celda.paragraphs[0], formato)
        return True
    except Exception as e:
        st.warning(f"⚠️ Error reconstruyendo tabla: {str(e)}")
        return False

def pre_limpiar_ocr(texto):
    patron = r'[^a-zA-ZáéíóúÁÉÍÓÚñÑüÜ0-9\s.,;:\-!?¿¡"\'\(\)\[\]/]'
    limpio = re.sub(patron, '', texto)
    return re.sub(r'\s+', ' ', limpio).strip()

# =====================================================================
# 3. DETECCIÓN DE IDIOMA MEJORADA
# =====================================================================
def extraer_muestra_representativa(doc_temp):
    parrafos_validos = []
    for p in doc_temp.paragraphs:
        texto = p.text.strip()
        if len(texto) > 30 and not texto.isdigit():
            texto_limpio = pre_limpiar_ocr(texto)
            if len(texto_limpio) > 30:
                parrafos_validos.append(texto_limpio)
    if not parrafos_validos:
        return ""
    inicio = 0
    for i, txt in enumerate(parrafos_validos):
        if re.search(r'\b(LIVRO|CAPÍTULO|CHAPTER|BOOK|PARTE)\b', txt, re.IGNORECASE):
            inicio = i
            break
    if inicio == 0:
        inicio = min(10, len(parrafos_validos)//2)
    return " ".join(parrafos_validos[inicio:inicio+20])

def detectar_idioma_con_groq(texto_muestra, groq_api_key):
    try:
        cliente = Groq(api_key=groq_api_key)
        respuesta = cliente.chat.completions.create(
            model=CONFIG["groq_model"],  # Usa el modelo configurado
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Eres un lingüista experto. El texto puede contener ruido de escáner. "
                        "Identifica el IDIOMA ORIGINAL del texto. Responde EXCLUSIVAMENTE con el nombre "
                        "del idioma en español, una sola palabra. Ejemplos: portugués, español, inglés, francés, alemán, italiano."
                    )
                },
                {"role": "user", "content": texto_muestra[:1500]}
            ],
            temperature=0,
            max_tokens=10
        )
        idioma = respuesta.choices[0].message.content.strip().lower()
        idioma = re.sub(r'[^a-záéíóúñ]', '', idioma)
        return idioma
    except Exception as e:
        st.warning(f"⚠️ Error detectando idioma con Groq: {str(e)}")
        return "desconocido"

def detectar_idioma_muestra(texto_muestra, groq_api_key):
    if not texto_muestra or len(texto_muestra) < 50:
        return "desconocido"
    if LANGDETECT_AVAILABLE:
        try:
            texto_limpio = pre_limpiar_ocr(texto_muestra)[:1000]
            if len(texto_limpio) > 50:
                codigo = detect(texto_limpio)
                mapa = {'pt': 'portugués', 'es': 'español', 'en': 'inglés',
                        'fr': 'francés', 'de': 'alemán', 'it': 'italiano'}
                return mapa.get(codigo, 'desconocido')
        except Exception as e:
            st.warning(f"langdetect falló, usando Groq: {str(e)}")
    return detectar_idioma_con_groq(texto_muestra, groq_api_key)

# =====================================================================
# 4. MOTOR DE TRADUCCIÓN (BATCH) CON NUEVO MODELO
# =====================================================================
def llamar_groq_con_reintento(texto_lote, groq_api_key, idioma_origen="inglés", max_reintentos=None):
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
                model=CONFIG["groq_model"],
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "Eres un editor editorial experto en restauración de textos escaneados.\n"
                            f"{instruccion_traduccion}\n"
                            "Se te pasarán múltiples bloques de texto separados por exactamente: '<<BLOCK_SEPARATOR>>'\n"
                            "Para CADA bloque:\n"
                            "1. Realiza la traducción/corrección.\n"
                            "2. Elimina basura de escaneo.\n"
                            "3. Corrige ortografía y puntuación.\n"
                            "Devuelve cada bloque separado por exactamente: '<<BLOCK_SEPARATOR>>'\n"
                            "Mantén el mismo número de bloques. Sin introducciones."
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
                # Backoff exponencial más agresivo: 10, 20, 40 segundos
                tiempo_espera = min((2 ** intento) * 10, 60)
                st.warning(f"⏳ Rate limit. Esperando {tiempo_espera}s... (Intento {intento+1}/{max_reintentos})")
                time.sleep(tiempo_espera)
            else:
                if es_rate_limit:
                    st.warning(f"❌ Rate limit persistente tras {max_reintentos} intentos.")
                else:
                    st.warning(f"⚠️ Error en API Groq: {str(e)[:200]}")
                return None
    return None

def traducir_lote(textos_lote, groq_api_key, idioma_origen="inglés"):
    if not textos_lote or all(not t.strip() for t in textos_lote):
        return textos_lote
    DELIMITER = "\n<<BLOCK_SEPARATOR>>\n"
    texto_combinado = DELIMITER.join(textos_lote)
    resultado = llamar_groq_con_reintento(texto_combinado, groq_api_key, idioma_origen)
    if resultado is None:
        return textos_lote
    traducidos = resultado.split(DELIMITER)
    if len(traducidos) != len(textos_lote):
        st.warning("⚠️ Integridad de separadores comprometida. Se mantienen originales.")
        return textos_lote
    return [t.strip() for t in traducidos]

def obtener_estilos_validos(doc):
    try:
        estilos_validos = {s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH}
        if 'Normal' not in estilos_validos:
            estilos_validos.add('Normal')
        return estilos_validos
    except:
        return {'Normal'}

def procesar_docx_multilingue(docx_path, docx_salida_path, groq_api_key, idioma_origen="inglés", tamano_lote=None, reanudar=True):
    if tamano_lote is None:
        tamano_lote = CONFIG["tamano_lote"]
    
    doc_original = docx.Document(docx_path)
    doc_nuevo = docx.Document()
    texto_estado = st.empty()
    texto_estado.text("Limpiando imágenes...")
    
    # Limpieza inicial
    img_eliminadas = limpiar_imagenes_pequenas(doc_original)
    st.info(f"🧹 Se eliminaron {img_eliminadas} artefactos visuales.")
    
    tablas_datos = preservar_tablas(doc_original)
    if tablas_datos:
        st.info(f"📊 Se detectaron {len(tablas_datos)} tabla(s).")
    
    estilos_validos = obtener_estilos_validos(doc_nuevo)
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
        st.info("No hay párrafos válidos.")
        doc_nuevo.save(docx_salida_path)
        return
    
    total = len(parrafos_datos)
    barra_progreso = st.progress(0)
    
    # Determinar desde qué lote empezar si se reanuda
    start_lote = 0
    if reanudar:
        last_completed = cargar_ultimo_lote()
        if last_completed > 0:
            start_lote = (last_completed - 1) * tamano_lote
            st.info(f"🔄 Reanudando desde el lote {last_completed} (párrafo {start_lote+1})")
            # Cargar traducciones ya guardadas
            st.session_state.traducciones_cache = cargar_progreso()
    
    for i in range(start_lote, total, tamano_lote):
        lote_data = parrafos_datos[i:i+tamano_lote]
        lote_numero = (i // tamano_lote) + 1
        total_lotes = (total + tamano_lote - 1) // tamano_lote
        textos_lote = [t[0] for t in lote_data]
        
        # Verificar si ya están traducidos en caché
        lote_completado = True
        textos_traducidos = []
        for j in range(len(textos_lote)):
            idx_global = str(i + j)
            if idx_global in st.session_state.traducciones_cache:
                textos_traducidos.append(st.session_state.traducciones_cache[idx_global])
            else:
                lote_completado = False
                break
        
        if lote_completado:
            texto_estado.text(f"⏩ Lote {lote_numero}/{total_lotes} recuperado de caché...")
        else:
            texto_estado.text(f"Traduciendo lote {lote_numero}/{total_lotes} (desde {idioma_origen})...")
            textos_traducidos = traducir_lote(textos_lote, groq_api_key, idioma_origen)
            for j, txt in enumerate(textos_traducidos):
                st.session_state.traducciones_cache[str(i + j)] = txt
            guardar_progreso(st.session_state.traducciones_cache)
        
        # Aplicar texto traducido al documento nuevo
        for j, (_, estilo, formato) in enumerate(lote_data):
            texto_final = textos_traducidos[j] if j < len(textos_traducidos) else ""
            estilo_seguro = estilo if estilo in estilos_validos else 'Normal'
            if texto_final.strip():
                p_nuevo = doc_nuevo.add_paragraph(texto_final, style=estilo_seguro)
                if CONFIG.get("preserve_formatting"):
                    aplicar_formato_parrafo(p_nuevo, formato)
            else:
                doc_nuevo.add_paragraph("", style=estilo_seguro)
        
        # Guardar progreso de lote completado
        guardar_ultimo_lote(lote_numero)
        
        progreso = min(i + tamano_lote, total)
        barra_progreso.progress(progreso / total)
        
        # Guardado incremental del documento
        if (lote_numero % CONFIG.get("save_frequency", 2) == 0) or (i + tamano_lote >= total):
            temp_path = docx_salida_path + ".tmp"
            try:
                doc_nuevo.save(temp_path)
                os.replace(temp_path, docx_salida_path)
                texto_estado.text(f"💾 Guardado en lote {lote_numero}...")
            except Exception as e:
                st.warning(f"⚠️ Error al guardar lote {lote_numero}: {str(e)}")
        
        # Pausa entre lotes para respetar rate limit
        time.sleep(CONFIG.get("inter_lote_sleep", 3.0))
    
    # Reconstruir tablas al final
    if tablas_datos:
        texto_estado.text("Reconstruyendo tablas...")
        for tabla_data in tablas_datos:
            contenidos_tabla = []
            for fila in tabla_data["contenido"]:
                contenidos_tabla.extend(fila)
            textos_traducidos_tabla = traducir_lote(contenidos_tabla, groq_api_key, idioma_origen)
            reconstruir_tabla(doc_nuevo, tabla_data, textos_traducidos_tabla)
    
    # Guardado final
    try:
        doc_nuevo.save(docx_salida_path)
        texto_estado.text("✅ Traducción completada.")
        limpiar_progreso()  # Borra archivos de progreso al terminar exitosamente
        st.session_state.traducciones_cache = {}
    except Exception as e:
        st.error(f"❌ Error al guardar documento final: {str(e)}")
    
    barra_progreso.empty()

# =====================================================================
# 5. INTERFAZ DE USUARIO Y CONTROL DE FLUJO PRINCIPAL
# =====================================================================
st.title("🚀 Conversor Editorial: PDF a Word Limpio v2")
st.markdown("Sube tus archivos **PDF escaneados** para convertirlos a **Word**, traducirlos al español y remover ruido de OCR.")

if not PYTESSERACT_AVAILABLE:
    st.info("ℹ️ pytesseract no instalado. OCR en imágenes deshabilitado.")
if not LANGDETECT_AVAILABLE:
    st.info("ℹ️ langdetect no instalado. La detección de idioma usará Groq (funciona igual).")

with st.sidebar:
    st.header("⚙️ Configuración")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📋 Ver Configuración"):
            st.json(CONFIG)
    with col2:
        if st.button("💾 Resetear Defaults"):
            guardar_config(DEFAULT_CONFIG)
            st.session_state.resultado = None
            limpiar_progreso()
            # No llamamos a st.rerun() para evitar reinicio brusco
    
    st.divider()
    st.subheader("Parámetros Batch")
    CONFIG["tamano_lote"] = st.slider("Tamaño de lote", 3, 10, CONFIG["tamano_lote"])
    CONFIG["max_reintentos"] = st.slider("Máx. reintentos API", 3, 10, CONFIG["max_reintentos"])
    CONFIG["inter_lote_sleep"] = st.slider("Pausa entre lotes (s)", 1.0, 10.0, CONFIG["inter_lote_sleep"], step=0.5)
    
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

# Credenciales
try:
    ADOBE_CLIENT_ID = st.secrets["PDF_SERVICES_CLIENT_ID"]
    ADOBE_CLIENT_SECRET = st.secrets["PDF_SERVICES_CLIENT_SECRET"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except KeyError as e:
    st.error(f"❌ Error crítico: Falta la credencial {e} en los Secrets de Streamlit.")
    st.stop()

# --- Subida y procesamiento con persistencia de resultado (SIN st.rerun()) ---
archivo_subido = st.file_uploader("Selecciona el libro o documento en formato PDF", type=["pdf"])

if archivo_subido is not None:
    if st.session_state.resultado is not None:
        st.session_state.resultado = None

# Opción de reanudación manual si existe progreso previo
progreso_previo = cargar_ultimo_lote()
if progreso_previo > 0 and archivo_subido is None:
    st.warning(f"⚠️ Se encontró un procesamiento interrumpido en el lote {progreso_previo}. Sube el mismo archivo para reanudar automáticamente.")

if archivo_subido and st.button("🚀 Comenzar Procesamiento Editorial", type="primary"):
    id_unico = str(int(time.time()))
    temp_pdf = f"temp_input_{id_unico}.pdf"
    temp_docx = f"temp_output_{id_unico}.docx"
    exito_total = False
    try:
        with open(temp_pdf, "wb") as f:
            f.write(archivo_subido.getbuffer())
            
        with st.spinner("Fase 1/3: Convirtiendo PDF a Word con Adobe..."):
            exito_adobe = convertir_pdf_a_word_adobe(temp_pdf, temp_docx, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET)
            
        if exito_adobe:
            with st.spinner("🔍 Detectando idioma del documento..."):
                doc_temp = docx.Document(temp_docx)
                texto_muestra = extraer_muestra_representativa(doc_temp)
                if texto_muestra and len(texto_muestra) > 100:
                    idioma_detectado = detectar_idioma_muestra(texto_muestra, GROQ_API_KEY)
                    st.info(f"🌍 Idioma detectado: **{idioma_detectado.capitalize()}**")
                else:
                    idioma_detectado = "inglés"
                    st.warning("⚠️ No se pudo detectar idioma. Asumiendo inglés.")
            
            with st.spinner("Fase 2/3: Traduciendo a español y limpiando..."):
                procesar_docx_multilingue(temp_docx, temp_docx, GROQ_API_KEY, idioma_origen=idioma_detectado, tamano_lote=CONFIG["tamano_lote"], reanudar=True)
            
            st.success("🎉 ¡Procesamiento completado!")
            exito_total = True
            
            ruta_final = mover_docx_a_resultados(temp_docx)
            st.session_state.resultado = {
                "exito": True,
                "ruta": ruta_final,
                "mensaje": "Documento procesado correctamente."
            }
        else:
            st.error("❌ La conversión de Adobe falló.")
            st.session_state.resultado = {"exito": False, "ruta": None, "mensaje": "Error en conversión Adobe."}
    except Exception as e:
        st.error(f"Error inesperado: {str(e)}")
        st.session_state.resultado = {"exito": False, "ruta": None, "mensaje": str(e)}
    finally:
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)
        if not exito_total and os.path.exists(temp_docx):
            carpeta_debug = "debug_fallos"
            os.makedirs(carpeta_debug, exist_ok=True)
            shutil.copy(temp_docx, os.path.join(carpeta_debug, f"error_{id_unico}.docx"))
            os.remove(temp_docx)
    # NO llamamos a st.rerun() aquí

# --- Mostrar resultado persistente ---
if st.session_state.resultado:
    res = st.session_state.resultado
    if res["exito"]:
        st.success(res["mensaje"])
        st.balloons()
        with open(res["ruta"], "rb") as f:
            st.download_button(
                label="📥 Descargar Documento Word Limpio",
                data=f,
                file_name=os.path.basename(res["ruta"]),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        st.info(f"📄 Archivo guardado en: `{res['ruta']}`")
    else:
        st.error(f"❌ Error: {res['mensaje']}")