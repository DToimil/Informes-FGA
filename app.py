import streamlit as st
import docx
import re
import os
import ast
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==========================================
# FUNCIONES AUXILIARES COMUNES
# ==========================================

def eliminar_fila(row):
    """Elimina la fila limpiamente desde el código XML de Word."""
    tr = row._tr
    if tr.getparent() is not None:
        tr.getparent().remove(tr)

def reemplazar_manteniendo_formato_estricto(parrafo, datos):
    """Reemplaza los datos copiando TODO el formato original (color, negrita, tamaño)."""
    texto = parrafo.text
    if not texto or "{{" not in texto:
        return

    formato = {}
    if len(parrafo.runs) > 0:
        for run in parrafo.runs:
            if run.text.strip():
                formato['name'] = run.font.name
                formato['size'] = run.font.size
                formato['bold'] = run.font.bold
                formato['italic'] = run.font.italic
                if run.font.color and run.font.color.rgb:
                    formato['color'] = run.font.color.rgb
                break

    for clave, valor in datos.items():
        if clave in texto:
            texto = texto.replace(clave, str(valor))
            
    # Limpia cualquier etiqueta {{ALGO}} que haya quedado sin rellenar
    texto = re.sub(r"\{\{.*?\}\}", "", texto)

    parrafo.clear()
    nuevo_run = parrafo.add_run(texto)
    
    if 'name' in formato and formato['name']: nuevo_run.font.name = formato['name']
    if 'size' in formato and formato['size']: nuevo_run.font.size = formato['size']
    if 'bold' in formato and formato['bold'] is not None: nuevo_run.font.bold = formato['bold']
    if 'italic' in formato and formato['italic'] is not None: nuevo_run.font.italic = formato['italic']
    if 'color' in formato: nuevo_run.font.color.rgb = formato['color']

def activar_checkbox_por_posicion(doc, indice_real, activar=True):
    """Activa o desactiva los cuadraditos grises de Word por su orden de aparición."""
    checkboxes = doc.element.xpath('.//w:checkBox')
    if indice_real < len(checkboxes):
        cb = checkboxes[indice_real]
        valor = "1" if activar else "0"
        
        default_val = cb.find(qn('w:default'))
        if default_val is not None:
            default_val.set(qn('w:val'), valor)
            
        checked_val = cb.find(qn('w:checked'))
        if checked_val is None:
            checked_val = OxmlElement('w:checked')
            cb.append(checked_val)
        checked_val.set(qn('w:val'), valor)


# ==========================================
# GENERADOR: DIRECTOR DE REUNIÓN
# ==========================================

def generar_acta_dr(datos_brutos):
    ruta_base = os.path.dirname(__file__)
    ruta_plantilla = os.path.join(ruta_base, "DR_PISTA_Plantilla_Maestra_Etiquetas.docx")
    doc = docx.Document(ruta_plantilla)
    
    # --- ESCUDO CORRECTOR DE LLAVES ---
    datos = {}
    for clave, valor in datos_brutos.items():
        if clave.startswith("{{") and clave.endswith("}}"):
            datos[clave] = valor
        else:
            datos[f"{{{{{clave}}}}}"] = valor

    # 1. Cabecera y textos sueltos
    for parrafo in doc.paragraphs:
        reemplazar_manteniendo_formato_estricto(parrafo, datos)

    # 2. Tablas (Limpieza de jueces vacíos)
    secciones_conocidas = [
        "CÁMARA DE LLAMADAS", "SALIDAS", "CRONOMETRAJE TRANSP.", 
        "CRONOMETRAJE MANUAL", "LLEGADAS", "CUENTAVUELTAS", 
        "JUECES DE MARCHA", "JUECES DE RECORRIDO", 
        "SECRET. COMPETICIÓN", "OTROS"
    ]

    for tabla in doc.tables:
        filas_a_borrar = []
        cabecera_actual = None
        seccion_con_datos = False
        
        for fila in tabla.rows:
            textos_celdas = [celda.text.strip() for celda in fila.cells]
            texto_fila = "".join(textos_celdas)
            texto_celda_0 = textos_celdas[0].upper() if textos_celdas else ""
            
            if not texto_fila:
                continue

            if texto_celda_0 in secciones_conocidas:
                if cabecera_actual is not None and not seccion_con_datos:
                    filas_a_borrar.append(cabecera_actual)
                cabecera_actual = fila
                seccion_con_datos = False
                continue

            if "{{" in texto_fila:
                if len(fila.cells) > 1:
                    celda_nombre = fila.cells[1].text 
                else:
                    celda_nombre = fila.cells[0].text
                    
                etiqueta_nombre = re.search(r"\{\{.*?_NOMBRE\}\}", celda_nombre)
                
                if etiqueta_nombre:
                    etiqueta = etiqueta_nombre.group()
                    if etiqueta not in datos or not str(datos[etiqueta]).strip():
                        filas_a_borrar.append(fila)
                        continue 
                    else:
                        seccion_con_datos = True
                else:
                    todas_etiquetas = re.findall(r"\{\{.*?\}\}", texto_fila)
                    todas_vacias = True
                    for etiq in todas_etiquetas:
                        if etiq in datos and str(datos[etiq]).strip():
                            todas_vacias = False
                            break
                    if todas_vacias and todas_etiquetas:
                        filas_a_borrar.append(fila)
                        continue
                    else:
                        seccion_con_datos = True

            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_manteniendo_formato_estricto(parrafo, datos)

        if cabecera_actual is not None and not seccion_con_datos:
            filas_a_borrar.append(cabecera_actual)

        for fila in filas_a_borrar:
            eliminar_fila(fila)

        filas_vacias_consecutivas = 0
        filas_rayas_a_borrar = []

        for fila in tabla.rows:
            textos_celdas = [celda.text.strip() for celda in fila.cells]
            if not "".join(textos_celdas):
                filas_vacias_consecutivas += 1
                if filas_vacias_consecutivas > 1: 
                    filas_rayas_a_borrar.append(fila)
            else:
                filas_vacias_consecutivas = 0

        if tabla.rows:
            textos_celdas_ult = [celda.text.strip() for celda in tabla.rows[-1].cells]
            if not "".join(textos_celdas_ult) and tabla.rows[-1] not in filas_rayas_a_borrar:
                filas_rayas_a_borrar.append(tabla.rows[-1])

        for fila in filas_rayas_a_borrar:
            eliminar_fila(fila)

    nombre_competicion = datos.get("{{COMPETICION}}", "Competicion")
    nombre_limpio = nombre_competicion.replace("/", "-").replace("\\", "-")
    nombre_docx = f"DR_{nombre_limpio}.docx"
    doc.save(nombre_docx)
    return nombre_docx


# ==========================================
# GENERADOR: JUEZ JEFE TRANSPONDEDOR
# ==========================================

def generar_acta_jjt(datos_brutos):
    ruta_base = os.path.dirname(__file__)
    ruta_plantilla = os.path.join(ruta_base, "JJT_PLANTILLA_MAESTRA_DEFINITIVA.docx")
    doc = docx.Document(ruta_plantilla)
    
    datos_texto = {}
    estado_cuadraditos = {}
    
    # SEPARADOR: Textos vs Cuadraditos dinámicos
    for clave, valor in datos_brutos.items():
        if clave.startswith("CHECK_"):
            num = int(clave.replace("CHECK_", ""))
            estado_cuadraditos[num] = valor
        else:
            if clave.startswith("{{") and clave.endswith("}}"):
                datos_texto[clave] = valor
            else:
                datos_texto[f"{{{{{clave}}}}}"] = valor

    # 1. Reemplazar texto
    for parrafo in doc.paragraphs:
        reemplazar_manteniendo_formato_estricto(parrafo, datos_texto)
        
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_manteniendo_formato_estricto(parrafo, datos_texto)

    # 2. Marcar cuadraditos
    for posicion, estado in estado_cuadraditos.items():
        activar_checkbox_por_posicion(doc, posicion, activar=estado)

    nombre_competicion = datos_texto.get("{{COMPETICION}}", "Informe_XXT").strip()
    nombre_limpio = nombre_competicion.replace("/", "-").replace("\\", "-")
    nombre_docx = f"JJT_{nombre_limpio}.docx"
    doc.save(nombre_docx)
    return nombre_docx


# ==========================================
# INTERFAZ WEB DE STREAMLIT
# ==========================================

st.set_page_config(page_title="Generador de Actas FGA", page_icon="📝")

st.title("Generador de Actas FGA 📝")
st.write("Sigue los pasos para generar el documento oficial en Word:")

# 1. Menú desplegable para elegir el tipo de informe
tipo_informe = st.selectbox(
    "1️⃣ Selecciona el tipo de informe que quieres generar:",
    ["Director de Reunión (Ruta / Pista)", "Juez Jefe de Transpondedor (JJT)"]
)

st.write("2️⃣ Pega debajo el texto del diccionario que te ha dado la Inteligencia Artificial.")

# 2. Cuadro de texto
texto_pegado = st.text_area("Pega aquí los datos (Diccionario):", height=300)

# 3. Botón de generación
if st.button("3️⃣ Generar Informe"):
    if not texto_pegado.strip():
        st.warning("¡El cuadro de texto está vacío! Pega los datos primero.")
    else:
        with st.spinner(f"Generando informe de {tipo_informe}..."):
            try:
                # Limpiamos espacios raros invisibles
                texto_limpio = texto_pegado.replace('\xa0', ' ')
                
                # Extraemos solo la parte del diccionario {...}
                inicio = texto_limpio.find('{')
                fin = texto_limpio.rfind('}') + 1
                
                if inicio == -1 or fin == 0:
                    st.error("No he encontrado ningún diccionario en el texto. Asegúrate de que empiece por '{' y acabe por '}'.")
                else:
                    texto_diccionario = texto_limpio[inicio:fin]
                    datos_procesados = ast.literal_eval(texto_diccionario)
                    
                    # Decidimos a qué función llamar según el desplegable
                    if tipo_informe == "Director de Reunión (Ruta / Pista)":
                        archivo_generado = generar_acta_dr(datos_procesados)
                    elif tipo_informe == "Juez Jefe de Transpondedor (JJT)":
                        archivo_generado = generar_acta_jjt(datos_procesados)
                    
                    st.success(f"¡{tipo_informe} generado con éxito!")
                    
                    # Creamos el botón de descarga
                    with open(archivo_generado, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Documento en Word",
                            data=file,
                            file_name=archivo_generado,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
            except SyntaxError:
                st.error("Error de formato: El texto que has pegado tiene algún error de sintaxis (falta una coma, unas comillas, etc). Revísalo.")
            except FileNotFoundError:
                st.error("Error: No se encuentra la plantilla en el servidor. Asegúrate de que el archivo .docx está subido a GitHub con el nombre correcto.")
            except Exception as e:
                st.error(f"Error inesperado: {e}")