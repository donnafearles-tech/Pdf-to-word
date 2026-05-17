import docx
import re
import streamlit as st
from groq import Groq

def pre_limpiar_ocr(texto):
    """
    Heurística 1: Limpieza dura antes de la IA para ahorrar tokens 
    y quitar basura incomprensible de los escáneres.
    """
    # Eliminar caracteres especiales que no aportan al texto (ruido OCR)
    texto = re.sub(r'[_<>|~^]', '', texto)
    # Eliminar múltiples espacios en blanco o saltos de línea
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def limpiar_y_traducir_con_groq(texto, groq_api_key):
    """
    Heurística 3: IA como correctora contextual y traductora.
    """
    cliente = Groq(api_key=groq_api_key)
    
    prompt_sistema = (
        "Eres un editor editorial experto en restauración de textos escaneados (OCR). "
        "Tus instrucciones son estrictas:\n"
        "1. Traduce el texto al ESPAÑOL si está en otro idioma (inglés, francés, etc.). Si ya está en español, mejóralo.\n"
        "2. Elimina cualquier basura de escaneo: símbolos extraños, marcas de agua leídas por error o sílabas repetitivas sin sentido.\n"
        "3. Corrige ortografía y puntuación manteniendo un tono editorial profesional.\n"
        "4. NO agregues comentarios, saludos, ni explicaciones. Devuelve ÚNICAMENTE el texto final."
    )
    
    try:
        respuesta = cliente.chat.completions.create(
            model="llama-3.1-8b-instant", # Excelente balance velocidad/calidad para este volumen
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": texto}
            ],
            temperature=0.1, # Más bajo para evitar alucinaciones, queremos el texto exacto
            max_tokens=8192
        )
        return respuesta.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"Error al conectar con Groq: {e}")
        return texto # Fallback: devolver original

def mejorar_word_con_groq(docx_path, groq_api_key):
    """
    Abre un .docx, aplica heurísticas de limpieza, traduce/corrige y guarda.
    """
    doc = docx.Document(docx_path)
    
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text.strip()
        
        # Heurística 4: Filtrar párrafos vacíos o que solo son números de página (ej. "  14  ")
        if not texto_original or texto_original.isdigit():
            continue
            
        texto_pre_limpio = pre_limpiar_ocr(texto_original)
        
        # Evitar llamar a la API por fragmentos demasiado pequeños sin valor semántico
        if len(texto_pre_limpio) > 3: 
            texto_final = limpiar_y_traducir_con_groq(texto_pre_limpio, groq_api_key)
            
            # --- MANEJO BÁSICO DE FORMATO ---
            # Borrar el contenido de los runs existentes para evitar duplicados
            for run in parrafo.runs:
                run.text = ""
                
            # Insertar el texto limpio en un nuevo run. 
            # NOTA: Para un manejo avanzado de negritas/cursivas reconstruidas por la IA,
            # aquí se necesitaría un parseador de Markdown a docx.runs.
            nuevo_run = parrafo.add_run(texto_final)
            
            # Si el primer run original tenía una fuente o estilo específico, lo heredamos al nuevo
            if len(parrafo.runs) > 1 and parrafo.runs[0].style:
                nuevo_run.style = parrafo.runs[0].style

    doc.save(docx_path)
