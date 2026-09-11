import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
import datetime

st.set_page_config(
    page_title="Certificaciones Laborales - Planta",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Generador Automático de Certificaciones Laborales")

RUTA_EXCEL = "BD 25 06 2026 (1).xlsx"
RUTA_WORD = "plantilla.docx"

# Lista de terminaciones o nombres masculinos comunes
NOMBRES_FEMENINOS = {'maria', 'ana', 'natalia', 'eliana', 'yenny', 'diana', 'sandra', 'martha', 'andrea', 'patricia', 'claudia', 'monica', 'carolina', 'laura', 'paula', 'angela', 'luz', 'gloria', 'esperanza', 'rosa', 'carmen', 'johan', 'julieth'}

def obtener_genero(nombre_completo):
    primer_nombre = str(nombre_completo).strip().split()[0].lower()
    # Si termina en 'a' o está en la lista de nombres femeninos comunes
    if primer_nombre.endswith('a') or primer_nombre in NOMBRES_FEMENINOS:
        if primer_nombre not in {'joshua', 'luca', 'sasha'}:
            return 'f'
    return 'm'

def fecha_a_texto(fecha_val):
    if pd.isnull(fecha_val):
        return ""
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
             'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    fecha_dt = pd.to_datetime(fecha_val)
    return f"{fecha_dt.strftime('%d')} de {meses[fecha_dt.month - 1]} de {fecha_dt.year}"

DICCIONARIO_CENTROS = {
    'CENTRO INDUST. DE MANTENIMIENTO Y MANUF.': 'Centro Industrial de Mantenimiento y Manufactura de la Regional Boyacá',
    'CENTRO DE GES. ADMIN. Y FORTALECIMIENTO EMPR.': 'Centro de Gestión Administrativa y Fortalecimiento Empresarial de la Regional Boyacá',
    'CENTRO MINERO': 'Centro Minero de la Regional Boyacá',
    'CENTRO DE DESAR. AGROP. Y AGROINDUSTRIAL': 'Centro de Desarrollo Agropecuario y Agroindustrial de la Regional Boyacá',
    'Centro de la Innovación Agroindustrial y de Servicios': 'Centro de la Innovación Agroindustrial y de Servicios de la Regional Boyacá',
    'DESPACHO DIRECCION': 'Despacho Dirección Regional Boyacá'
}

if os.path.exists(RUTA_EXCEL) and os.path.exists(RUTA_WORD):
    df = pd.read_excel(RUTA_EXCEL)
    df['Identificación'] = df['Identificación'].astype(str)
    
    st.subheader("🔎 Búsqueda Directa por Nombre o Cédula")
    
    nombre_ingresado = st.text_input(
        "Ingresa el nombre o cédula del funcionario:", 
        placeholder="Ej: Natalia Alvarez o 1052393705"
    )
    
    if nombre_ingresado:
        busqueda = nombre_ingresado.strip().lower()
        
        coincidencias = df[
            df['Identificación'].str.lower().str.contains(busqueda) |
            df['Nombres'].astype(str).str.lower().str.contains(busqueda) |
            df['Apellidos'].astype(str).str.lower().str.contains(busqueda)
        ]
        
        if coincidencias.empty:
            st.warning("⚠️ No se encontró ningún funcionario con ese nombre o cédula.")
        else:
            opciones = coincidencias.apply(lambda r: f"{r['Identificación']} - {r['Nombres']} {r['Apellidos']}", axis=1)
            persona_seleccionada = st.selectbox("Confirmar funcionario encontrado:", opciones)
            
            if persona_seleccionada:
                cedula_sel = persona_seleccionada.split(" - ")[0]
                registro = df[df['Identificación'] == cedula_sel].iloc[0].to_dict()
                
                # 1. Determinar género y adaptar conectores
                genero = obtener_genero(registro.get('Nombres', ''))
                if genero == 'f':
                    registro['Identificado_a'] = 'Identificada'
                    registro['vinculado_a'] = 'vinculada'
                    registro['del_de_la_interesado_a'] = 'de la interesada'
                else:
                    registro['Identificado_a'] = 'Identificado'
                    registro['vinculado_a'] = 'vinculado'
                    registro['del_de_la_interesado_a'] = 'del interesado'
                
                # 2. Formato de fechas y moneda
                registro['Fecha_Inicio_Texto'] = fecha_a_texto(registro.get('Fecha Inicio'))
                if 'Sueldo Anterior' in registro and pd.notnull(registro['Sueldo Anterior']):
                    registro['Sueldo_Anterior'] = f"{int(registro['Sueldo Anterior']):,}".replace(",", ".")
                
                # 3. Centro de costo completo
                centro_original = str(registro.get('Nombre Centro Costo', '')).strip()
                registro['Centro_Costo_Completo'] = DICCIONARIO_CENTROS.get(centro_original, centro_original)
                
                # 4. Fecha de expedición del día de hoy
                registro['Fecha_Expedicion'] = fecha_a_texto(datetime.date.today())
                
                datos_contexto = {k.replace(' ', '_'): v for k, v in registro.items()}
                
                doc = DocxTemplate(RUTA_WORD)
                doc.render(datos_contexto)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                nombre_descarga = f"Certificado_{registro['Identificación']}_{registro['Nombres']}.docx"
                
                st.success(f" Carta generada para: **{registro['Nombres']} {registro['Apellidos']}** ({'Mujer' if genero == 'f' else 'Hombre'})")
                
                st.download_button(
                    label="⬇️ Descargar Carta / Certificado Word",
                    data=buffer,
                    file_name=nombre_descarga,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
else:
    st.error("❌ Verifique que los archivos 'BD 25 06 2026 (1).xlsx' y 'plantilla.docx' estén presentes en la carpeta.")
