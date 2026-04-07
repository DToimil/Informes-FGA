import streamlit as st
import docx
import re
import os
import ast

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


def generar_acta_final(datos_brutos):
    # ATENCIÓN: Asegúrate de que el nombre del Word aquí coincide exactamente con el que tienes subido a GitHub
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
    # ----------------------------------

    # 1. Cabecera y textos sueltos
    for parrafo in doc.paragraphs:
        reemplazar_manteniendo_formato_estricto(parrafo, datos)

    # 2. Tablas
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
        
        # PASO 1: Identificar sobrantes (jueces vacíos y cabeceras inútiles) y rellenar datos
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

            # Si la fila es útil y se va a quedar, le escribimos los datos
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_manteniendo_formato_estricto(parrafo, datos)

        if cabecera_actual is not None and not seccion_con_datos:
            filas_a_borrar.append(cabecera_actual)

        # Destruimos todo lo inútil
        for fila in filas_a_borrar:
            eliminar_fila(fila)

        # PASO 2: Limpieza Quirúrgica de "rayas"
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

        # Si el documento termina en raya vacía, también la quitamos
        if tabla.rows:
            textos_celdas_ult = [celda.text.strip() for celda in tabla.rows[-1].cells]
            if not "".join(textos_celdas_ult) and tabla.rows[-1] not in filas_rayas_a_borrar:
                filas_rayas_a_borrar.append(tabla.rows[-1])

        # Destruimos las rayas duplicadas
        for fila in filas_rayas_a_borrar:
            eliminar_fila(fila)

    # 3. Guardado en Word
    nombre_competicion = datos.get("{{COMPETICION}}", "Competicion")
    nombre_limpio = nombre_competicion.replace("/", "-").replace("\\", "-")
    nombre_docx = f"DR {nombre_limpio}.docx"
    
    doc.save(nombre_docx)
    return nombre_docx


# ==========================================
# INTERFAZ WEB DE STREAMLIT (Aquí está el "huequito")
# ==========================================

st.title("Generador de Actas FGA 📝")
st.write("1. Pega debajo el texto del diccionario que te ha dado la Inteligencia Artificial.")
st.write("2. Dale a generar y descarga tu Word.")

# AQUÍ ESTÁ EL CUADRO DE TEXTO
texto_pegado = st.text_area("Pega aquí los datos (Diccionario):", height=300)

if st.button("Generar Acta"):
    if not texto_pegado.strip():
        st.warning("¡Eh! El cuadro de texto está vacío. Pega los datos primero.")
    else:
        with st.spinner("Leyendo datos y generando documento..."):
            try:
                # Limpiamos espacios raros invisibles
                texto_limpio = texto_pegado.replace('\xa0', ' ')
                
                # Extraemos solo la parte del diccionario {...}
                inicio = texto_limpio.find('{')
                fin = texto_limpio.rfind('}') + 1
                
                if inicio == -1 or fin == 0:
                    st.error("No he encontrado ningún diccionario en el texto. Asegúrate de que empiece por '{' y acabe por '}'.")
                else:
                    # Convertimos el texto pegado en un diccionario real de Python
                    texto_diccionario = texto_limpio[inicio:fin]
                    datos_procesados = ast.literal_eval(texto_diccionario)
                    
                    # Llamamos a tu función pasándole los datos
                    archivo_generado = generar_acta_final(datos_procesados)
                    
                    st.success("¡Acta generada con éxito!")
                    
                    # Creamos el botón de descarga
                    with open(archivo_generado, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Acta en Word",
                            data=file,
                            file_name=archivo_generado,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
            except SyntaxError:
                st.error("Error de formato: El texto que has pegado tiene algún error de sintaxis (falta una coma, unas comillas, etc). Revísalo.")
            except Exception as e:
                st.error(f"Error inesperado: {e}")