import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os

st.set_page_config(
    page_title="Generador de Certificaciones",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Generador Automático de Certificaciones Laborales")

# Rutas locales dentro del proyecto
RUTA_EXCEL = "BD 25 06 2026 (1).xlsx"
RUTA_WORD = "plantilla.docx"

# Verificar si los archivos ya existen en el repositorio
if os.path.exists(RUTA_EXCEL) and os.path.exists(RUTA_WORD):
    df = pd.read_excel(RUTA_EXCEL)
    df['Identificación'] = df['Identificación'].astype(str)
    plantilla_path = RUTA_WORD
else:
    st.info("💡 Sube los archivos base para comenzar.")
    col1, col2 = st.columns(2)
    with col1:
        archivo_excel = st.file_uploader("1. Cargar Base de Datos (.xlsx)", type=["xlsx"])
    with col2:
        archivo_word = st.file_uploader("2. Cargar Plantilla de Word (.docx)", type=["docx"])
    
    if archivo_excel and archivo_word:
        df = pd.read_excel(archivo_excel)
        df['Identificación'] = df['Identificación'].astype(str)
        plantilla_path = archivo_word
    else:
        df = None

if df is not None:
    st.subheader("🔎 Búsqueda Directa por Nombre o Cédula")
    
    # Campo donde solo ingresas el nombre
    nombre_ingresado = st.text_input(
        "Ingresa el nombre o cédula del funcionario para generar la carta:", 
        placeholder="Ej: Natalia Alvarez o 1052393705"
    )
    
    if nombre_ingresado:
        busqueda = nombre_ingresado.strip().lower()
        
        # Filtro por coincidencia en Nombre, Apellido o Cédula
        coincidencias = df[
            df['Identificación'].str.lower().str.contains(busqueda) |
            df['Nombres'].astype(str).str.lower().str.contains(busqueda) |
            df['Apellidos'].astype(str).str.lower().str.contains(busqueda)
        ]
        
        if coincidencias.empty:
            st.warning("⚠️ No se encontró ningún funcionario con ese nombre o cédula.")
        else:
            # Seleccionar la persona encontrada
            opciones = coincidencias.apply(lambda r: f"{r['Identificación']} - {r['Nombres']} {r['Apellidos']}", axis=1)
            persona_seleccionada = st.selectbox("Confirmar persona encontrada:", opciones)
            
            if persona_seleccionada:
                cedula_sel = persona_seleccionada.split(" - ")[0]
                registro = df[df['Identificación'] == cedula_sel].iloc[0].to_dict()
                
                # Formato de variables
                if 'Fecha Inicio' in registro and pd.notnull(registro['Fecha Inicio']):
                    registro['Fecha_Inicio'] = pd.to_datetime(registro['Fecha Inicio']).strftime('%d/%m/%Y')
                
                if 'Sueldo Anterior' in registro and pd.notnull(registro['Sueldo Anterior']):
                    registro['Sueldo_Anterior'] = f"{int(registro['Sueldo Anterior']):,}".replace(",", ".")
                
                datos_contexto = {k.replace(' ', '_'): v for k, v in registro.items()}
                
                # Cargar la plantilla y armar la carta
                doc = DocxTemplate(plantilla_path)
                doc.render(datos_contexto)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                nombre_archivo_descarga = f"Certificado_{registro['Identificación']}_{registro['Nombres']}.docx"
                
                st.success(f" Carta generada para: **{registro['Nombres']} {registro['Apellidos']}**")
                
                # Botón de descarga directa
                st.download_button(
                    label="⬇️ Descargar Carta en Word (.docx)",
                    data=buffer,
                    file_name=nombre_archivo_descarga,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
