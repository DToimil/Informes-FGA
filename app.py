import streamlit as st
import docx
import re
import os

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


def generar_acta_final():
    doc = docx.Document("DR_PISTA_Plantilla_Maestra_Etiquetas.docx")
    
    # Este es el diccionario que te escupe el Gem (sin llaves, no pasa nada)
    datos_brutos = {
        "COMPETICION": "IV Os 10000 Peregrinos Deputación da Coruña",
        "LUGAR": "SANTIAGO DE COMPOSTELA",
        "DELEGACION": "Santiago",
        "DIA_SEMANA": "Domingo",
        "FECHA_DIA": "22",
        "MES": "Marzo",
        "ANO": "2026",
        "NUM_JORNADA": "1",
        "JORNADA": "Mañana",
        "OBSERVACIONES": "Se destaca la excelente disposición y colaboración de todos los miembros del equipo arbitral durante el desarrollo de la jornada.",
        
        "DIRECTOR_REUNION_NOMBRE": "JULIO RODRÍGUEZ GARCÍA",
        "DIRECTOR_REUNION_CAT": "N1",
        "DIRECTOR_REUNION_DEL": "SA",
        
        "JUEZ_ARBITRO_NOMBRE": "RODRIGO ESPAÑA PETEIRO",
        "JUEZ_ARBITRO_CAT": "N2",
        "JUEZ_ARBITRO_DEL": "SA",
        
        "JUEZ_ARBITRO_2_NOMBRE": "MARÍA JOSÉ BARBANZÁN SUEIRO",
        "JUEZ_ARBITRO_2_CAT": "N2",
        "JUEZ_ARBITRO_2_DEL": "SA",
        
        "DELEGADO_TECNICO_NOMBRE": "ANTÓN NOGUEIRA GONZÁLEZ",
        "DELEGADO_TECNICO_CAT": "N1",
        "DELEGADO_TECNICO_DEL": "SA",
        
        "JUEZ_DE_SALIDAS_NOMBRE": "MOISES IGLESIAS AMENEIRO",
        "JUEZ_DE_SALIDAS_CAT": "N1",
        "JUEZ_DE_SALIDAS_DEL": "SA",
        
        "AYUDANTE_DE_SALIDAS_1_NOMBRE": "MANUEL TREUS PAMPÍN",
        "AYUDANTE_DE_SALIDAS_1_CAT": "N1",
        "AYUDANTE_DE_SALIDAS_1_DEL": "SA",
        
        "JUEZ_JEFE_TRANSPONDEDORES_NOMBRE": "EVA SALVADO PRIETO",
        "JUEZ_JEFE_TRANSPONDEDORES_CAT": "N3",
        "JUEZ_JEFE_TRANSPONDEDORES_DEL": "PO",
        "JUEZ_JEFE_TRANSPONDEDORES_DESP": "X",
        
        "JEFE_CRONOMETRAJE_NOMBRE": "EVA SALVADO PRIETO",
        "JEFE_CRONOMETRAJE_CAT": "N3",
        "JEFE_CRONOMETRAJE_DEL": "PO",
        
        "JUEZ_CRONOMETRAJE_2_NOMBRE": "NEREA IGLESIAS BARBANZÁN",
        "JUEZ_CRONOMETRAJE_2_CAT": "N1",
        "JUEZ_CRONOMETRAJE_2_DEL": "SA",
        
        "JEFE_LLEGADAS_NOMBRE": "ADRIÁN SESAR MÍGUEZ",
        "JEFE_LLEGADAS_CAT": "N1",
        "JEFE_LLEGADAS_DEL": "SA",
        
        "JUEZ_DE_LLEGADAS_1_NOMBRE": "AINOA LAGO FERNÁNDEZ",
        "JUEZ_DE_LLEGADAS_1_CAT": "N1",
        "JUEZ_DE_LLEGADAS_1_DEL": "SA",
        
        "JUEZ_DE_LLEGADAS_2_NOMBRE": "LÍA PEREIRA SEIJO",
        "JUEZ_DE_LLEGADAS_2_CAT": "N1",
        "JUEZ_DE_LLEGADAS_2_DEL": "SA",
        
        "JUEZ_DE_LLEGADAS_3_NOMBRE": "CINTHIA COSTAS GONZÁLEZ",
        "JUEZ_DE_LLEGADAS_3_CAT": "N1",
        "JUEZ_DE_LLEGADAS_3_DEL": "VI",
        
        "JEFE_CUENTAVUELTAS_NOMBRE": "ADRIÁN SESAR MÍGUEZ",
        "JEFE_CUENTAVUELTAS_CAT": "N1",
        "JEFE_CUENTAVUELTAS_DEL": "SA",
        
        "CUENTAVUELTAS_1_NOMBRE": "AINOA LAGO FERNÁNDEZ",
        "CUENTAVUELTAS_1_CAT": "N1",
        "CUENTAVUELTAS_1_DEL": "SA",
        
        "CUENTAVUELTAS_2_NOMBRE": "LÍA PEREIRA SEIJO",
        "CUENTAVUELTAS_2_CAT": "N1",
        "CUENTAVUELTAS_2_DEL": "SA",
        
        "CUENTAVUELTAS_3_NOMBRE": "CINTHIA COSTAS GONZÁLEZ",
        "CUENTAVUELTAS_3_CAT": "N1",
        "CUENTAVUELTAS_3_DEL": "VI",
        
        "JUEZ_DE_RECORRIDO_1_NOMBRE": "VICENTE M. SÁNCHEZ PÉREZ",
        "JUEZ_DE_RECORRIDO_1_CAT": "N1",
        "JUEZ_DE_RECORRIDO_1_DEL": "SA",
        
        "JUEZ_DE_RECORRIDO_2_NOMBRE": "ÓSCAR TOIMIL PLAZA",
        "JUEZ_DE_RECORRIDO_2_CAT": "N1",
        "JUEZ_DE_RECORRIDO_2_DEL": "SA",
        
        "SECRETARIA_1": "EMESPORTS"
    }

    # --- ESCUDO CORRECTOR DE LLAVES ---
    # Transforma automáticamente "COMPETICION" en "{{COMPETICION}}" para que el script no falle
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
            
            # Si es una fila espaciadora (raya vacía), nos la saltamos en este paso para no borrarla aún
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
                if filas_vacias_consecutivas > 1: # Si hay más de 1 raya seguida, marcamos para borrar
                    filas_rayas_a_borrar.append(fila)
            else:
                filas_vacias_consecutivas = 0

        # Si el documento termina en raya vacía, también la quitamos para que el borde se cierre bien
        if tabla.rows:
            textos_celdas_ult = [celda.text.strip() for celda in tabla.rows[-1].cells]
            if not "".join(textos_celdas_ult) and tabla.rows[-1] not in filas_rayas_a_borrar:
                filas_rayas_a_borrar.append(tabla.rows[-1])

        # Destruimos las rayas duplicadas
        for fila in filas_rayas_a_borrar:
            eliminar_fila(fila)

    # 3. Guardado en Word y PDF
    nombre_competicion = datos.get("{{COMPETICION}}", "Competicion")
    nombre_limpio = nombre_competicion.replace("/", "-").replace("\\", "-")
    
    nombre_docx = f"DR_{nombre_limpio}.docx"
    nombre_pdf = f"DR_{nombre_limpio}.pdf"
    
    doc.save(nombre_docx)
    print(f"✅ Archivo Word generado: {nombre_docx}")

if __name__ == "__main__":
    generar_acta_final()




# ==========================================
# INTERFAZ WEB DE STREAMLIT (Lo que tú ves)
# ==========================================

# 1. Título de la página
st.title("Generador de Actas FGA 📝")
st.write("Haz clic en el botón de abajo para procesar los datos y generar el documento.")

# 2. Creamos un botón. Todo lo que esté indentado debajo ocurrirá al pulsarlo.
if st.button("Generar Acta"):
    
    with st.spinner("Generando documento, por favor espera..."):
        try:
            # Llamamos a tu función mágica
            archivo_generado = generar_acta_final()
            
            st.success("¡Acta generada con éxito!")
            
            # 3. Creamos el botón de descarga para que te baje el Word a tu ordenador
            with open(archivo_generado, "rb") as file:
                st.download_button(
                    label="📥 Descargar Acta en Word",
                    data=file,
                    file_name=archivo_generado,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Ocurrió un error: {e}")