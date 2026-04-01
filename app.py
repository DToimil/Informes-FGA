import streamlit as st
import ast
import docx
import re
import os

def eliminar_fila(row):
    tr = row._tr
    if tr.getparent() is not None:
        tr.getparent().remove(tr)

def reemplazar_manteniendo_formato_estricto(parrafo, datos):
    texto = parrafo.text
    if not texto or "{{" not in texto: return
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
        if clave in texto: texto = texto.replace(clave, str(valor))
    texto = re.sub(r"\{\{.*?\}\}", "", texto)
    parrafo.clear()
    nuevo_run = parrafo.add_run(texto)
    if 'name' in formato and formato['name']: nuevo_run.font.name = formato['name']
    if 'size' in formato and formato['size']: nuevo_run.font.size = formato['size']
    if 'bold' in formato and formato['bold'] is not None: nuevo_run.font.bold = formato['bold']
    if 'italic' in formato and formato['italic'] is not None: nuevo_run.font.italic = formato['italic']
    if 'color' in formato: nuevo_run.font.color.rgb = formato['color']

def generar_archivos(datos_dict):
    doc = docx.Document("DR_PISTA_Plantilla_Maestra_Etiquetas.docx")
    
    for parrafo in doc.paragraphs:
        reemplazar_manteniendo_formato_estricto(parrafo, datos_dict)

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
            if not texto_fila: continue

            if texto_celda_0 in secciones_conocidas:
                if cabecera_actual is not None and not seccion_con_datos:
                    filas_a_borrar.append(cabecera_actual)
                cabecera_actual = fila
                seccion_con_datos = False
                continue

            if "{{" in texto_fila:
                if len(fila.cells) > 1: celda_nombre = fila.cells[1].text 
                else: celda_nombre = fila.cells[0].text
                etiqueta_nombre = re.search(r"\{\{.*?_NOMBRE\}\}", celda_nombre)
                
                if etiqueta_nombre:
                    etiqueta = etiqueta_nombre.group()
                    if etiqueta not in datos_dict or not str(datos_dict[etiqueta]).strip():
                        filas_a_borrar.append(fila)
                        continue 
                    else: seccion_con_datos = True
                else:
                    todas_etiquetas = re.findall(r"\{\{.*?\}\}", texto_fila)
                    todas_vacias = True
                    for etiq in todas_etiquetas:
                        if etiq in datos_dict and str(datos_dict[etiq]).strip():
                            todas_vacias = False
                            break
                    if todas_vacias and todas_etiquetas:
                        filas_a_borrar.append(fila)
                        continue
                    else: seccion_con_datos = True

            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_manteniendo_formato_estricto(parrafo, datos_dict)

        if cabecera_actual is not None and not seccion_con_datos:
            filas_a_borrar.append(cabecera_actual)
        for fila in filas_a_borrar: eliminar_fila(fila)

        filas_vacias_consecutivas = 0
        filas_rayas_a_borrar = []
        for fila in tabla.rows:
            textos_celdas = [celda.text.strip() for celda in fila.cells]
            if not "".join(textos_celdas):
                filas_vacias_consecutivas += 1
                if filas_vacias_consecutivas > 1: filas_rayas_a_borrar.append(fila)
            else: filas_vacias_consecutivas = 0
        if tabla.rows:
            textos_celdas_ult = [celda.text.strip() for celda in tabla.rows[-1].cells]
            if not "".join(textos_celdas_ult) and tabla.rows[-1] not in filas_rayas_a_borrar:
                filas_rayas_a_borrar.append(tabla.rows[-1])
        for fila in filas_rayas_a_borrar: eliminar_fila(fila)

    nombre_competicion = datos_dict.get("{{COMPETICION}}", "Competicion")
    nombre_limpio = nombre_competicion.replace("/", "-").replace("\\", "-")
    nombre_docx = f"DR {nombre_limpio}.docx"
    
    # Guardamos el Word y terminamos
    doc.save(nombre_docx)
    return nombre_docx


# --- INTERFAZ WEB ---
st.set_page_config(page_title="Actas FGA", page_icon="📝")

st.title("📝 Generador de Actas FGA")
st.write("Pega el diccionario de datos generado por tu Gem y pulsa Generar para obtener el documento de Word.")

codigo_pegado = st.text_area("Pega aquí el código (datos = {...}):", height=250)

if st.button("Generar Acta (Word)", type="primary"):
    if codigo_pegado:
        try:
            texto_limpio = codigo_pegado.replace("```python", "").replace("```", "")
            if "datos =" in texto_limpio:
                texto_limpio = texto_limpio.split("datos =")[1].strip()
            
            datos_diccionario = ast.literal_eval(texto_limpio)
            
            with st.spinner('Ensamblando el Word...'):
                docx_file = generar_archivos(datos_diccionario)
            
            st.success("¡Documento de Word generado con éxito!")
            
            with open(docx_file, "rb") as file:
                st.download_button(
                    label="📥 Descargar Word (.docx)", 
                    data=file, 
                    file_name=docx_file, 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Error al leer los datos. Asegúrate de que copiaste bien el diccionario. Detalle: {e}")
    else:
        st.warning("Pega los datos primero.")