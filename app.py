import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

# Configuración de interfaz
st.set_page_config(
    page_title="Generador de Certificaciones",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Generador Automático de Certificaciones Laborales")
st.write("Carga la plantilla de Word y el archivo Excel con los datos actualizados para generar los documentos.")

# Secciones de carga de archivos
col1, col2 = st.columns(2)

with col1:
    archivo_excel = st.file_uploader("1. Cargar Base de Datos (.xlsx)", type=["xlsx"])

with col2:
    archivo_word = st.file_uploader("2. Cargar Plantilla de Word (.docx)", type=["docx"])

if archivo_excel and archivo_word:
    # Lectura y preparación del Excel
    df = pd.read_excel(archivo_excel)
    df['Identificación'] = df['Identificación'].astype(str)
    
    st.success(f" Base de datos cargada correctamente. Se encontraron {len(df)} registros.")
    
    with st.expander("Ver vista previa de los datos"):
        st.dataframe(df.head())
    
    st.divider()
    
    modo = st.radio("Selecciona la modalidad de generación:", ["Generación Individual", "Generación Masiva (ZIP)"])
    
    # ------------------ MODO INDIVIDUAL ------------------
    if modo == "Generación Individual":
        st.subheader(" Búsqueda y Generación Individual")
        
        empleados = df.apply(lambda row: f"{row['Identificación']} - {row['Nombres']} {row['Apellidos']}", axis=1)
        seleccion = st.selectbox("Buscar funcionario por Cédula o Nombre:", empleados)
        
        if seleccion:
            cedula_sel = seleccion.split(" - ")[0]
            registro = df[df['Identificación'] == cedula_sel].iloc[0].to_dict()
            
            # Formato de fechas y moneda
            if 'Fecha Inicio' in registro and pd.notnull(registro['Fecha Inicio']):
                registro['Fecha_Inicio'] = pd.to_datetime(registro['Fecha Inicio']).strftime('%d/%m/%Y')
            
            if 'Sueldo Anterior' in registro and pd.notnull(registro['Sueldo Anterior']):
                registro['Sueldo_Anterior'] = f"{int(registro['Sueldo Anterior']):,}".replace(",", ".")
            
            # Formatear claves eliminando espacios
            datos_contexto = {k.replace(' ', '_'): v for k, v in registro.items()}
            
            st.write("Datos a insertar en el documento:")
            st.json(datos_contexto)
            
            if st.button("Generar Certificado"):
                doc = DocxTemplate(archivo_word)
                doc.render(datos_contexto)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                nombre_salida = f"Certificado_{registro['Identificación']}_{registro['Nombres']}.docx"
                st.download_button(
                    label=" Descargar Documento Word",
                    data=buffer,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    # ------------------ MODO MASIVO ------------------
    elif modo == "Generación Masiva (ZIP)":
        st.subheader("📦 Generación Masiva para Toda la Base de Datos")
        
        if st.button("Procesar Todos los Certificados"):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, fila in df.iterrows():
                    registro = fila.to_dict()
                    
                    if 'Fecha Inicio' in registro and pd.notnull(registro['Fecha Inicio']):
                        registro['Fecha_Inicio'] = pd.to_datetime(registro['Fecha Inicio']).strftime('%d/%m/%Y')
                    
                    if 'Sueldo Anterior' in registro and pd.notnull(registro['Sueldo Anterior']):
                        registro['Sueldo_Anterior'] = f"{int(registro['Sueldo Anterior']):,}".replace(",", ".")
                    
                    datos_contexto = {k.replace(' ', '_'): v for k, v in registro.items()}
                    
                    doc = DocxTemplate(archivo_word)
                    doc.render(datos_contexto)
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    nombre_doc = f"Certificado_{datos_contexto.get('Identificación', idx)}.docx"
                    zip_file.writestr(nombre_doc, doc_buffer.getvalue())
            
            zip_buffer.seek(0)
            st.download_button(
                label=" Descargar Todos los Certificados (.ZIP)",
                data=zip_buffer,
                file_name="Certificados_Masivos.zip",
                mime="application/zip"
            )