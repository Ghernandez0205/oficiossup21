import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os
import zipfile
from io import BytesIO

# Rutas relativas
TEMPLATE_PATH = "plantillas/001 OFICIO ciclo escolar 2024-2025.docx"
EXCEL_PATH = "datos/PLANTILLA.xlsx"
HISTORIAL_PATH = "historial_oficios.xlsx"

# Función para crear carpeta temporal
def crear_carpeta_temporal():
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_folder = os.path.join("/tmp", f"Oficios_{timestamp}")
    os.makedirs(output_folder, exist_ok=True)
    return output_folder

# Función para generar oficios
def generar_oficio(data, num_oficio, sede, ubicacion, fecha_comision, horario, comision):
    archivos_generados = []
    output_folder = crear_carpeta_temporal()

    meses_es = {
        'January': 'enero', 'February': 'febrero', 'March': 'marzo', 'April': 'abril',
        'May': 'mayo', 'June': 'junio', 'July': 'julio', 'August': 'agosto',
        'September': 'septiembre', 'October': 'octubre', 'November': 'noviembre', 'December': 'diciembre'
    }

    for _, fila in data.iterrows():
        nombre = fila['NOMBRE']
        apellido_paterno = fila['APELLIDO PATERNO']
        apellido_materno = fila['APELLIDO MATERNO']
        rfc = fila['RFC']

        doc = Document(TEMPLATE_PATH)

        meses_es = {
    'January': 'enero', 'February': 'febrero', 'March': 'marzo', 'April': 'abril',
    'May': 'mayo', 'June': 'junio', 'July': 'julio', 'August': 'agosto',
    'September': 'septiembre', 'October': 'octubre', 'November': 'noviembre', 'December': 'diciembre'
}

mes_en_espanol = meses_es[fecha_comision.strftime('%B')]
fecha_comision_str = f"{fecha_comision.day} de {mes_en_espanol} del {fecha_comision.year}"

        for en, es in meses_es.items():
            fecha_comision_str = fecha_comision_str.replace(en, es)

        for p in doc.paragraphs:
            p.text = p.text.replace("mes", fecha_comision.strftime('%B').capitalize())
            p.text = p.text.replace("fecha", fecha_comision_str)
            p.text = p.text.replace("numero_oficio", num_oficio)
            p.text = p.text.replace("nombre", nombre)
            p.text = p.text.replace("apellido_paterno", apellido_paterno)
            p.text = p.text.replace("apellido_materno", apellido_materno)
            p.text = p.text.replace("rfc", rfc)
            p.text = p.text.replace("sede", sede)
            p.text = p.text.replace("ubicacion", ubicacion)
            p.text = p.text.replace("horario", horario)
            p.text = p.text.replace("comision", comision)

        nombre_archivo = f"oficio_{apellido_paterno}_{nombre.replace(' ', '_')}.docx"        
        ruta_archivo = os.path.join(output_folder, nombre_archivo)
        doc.save(ruta_archivo)
        archivos_generados.append(ruta_archivo)

    return archivos_generados

# Función para comprimir
def comprimir_archivos(archivos):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for archivo in archivos:
            zipf.write(archivo, os.path.basename(archivo))
    zip_buffer.seek(0)
    return zip_buffer

# Función para actualizar historial
def actualizar_historial(data, num_oficio, comision):
    historial_df = pd.DataFrame()
    if os.path.exists(HISTORIAL_PATH):
        historial_df = pd.read_excel(HISTORIAL_PATH)

    nuevo_historial = pd.DataFrame({
        "Número Consecutivo": [len(historial_df) + i + 1 for i in range(len(data))],
        "Nombre": data['NOMBRE'].values,
        "Apellido Paterno": data['APELLIDO PATERNO'].values,
        "Apellido Materno": data['APELLIDO MATERNO'].values,
        "Número de Oficio": [num_oficio] * len(data),
        "Actividad": [comision] * len(data)
    })

    historial_df = pd.concat([historial_df, nuevo_historial], ignore_index=True)
    historial_df.to_excel(HISTORIAL_PATH, index=False)

# Interfaz Streamlit
st.set_page_config(page_title="Generador de Oficios", page_icon="📄")
st.title("📄 Generador de Oficios en Word")

# Contraseña
password = st.text_input("🔒 Ingrese la contraseña", type="password")
if password != "defvm11":
    st.warning("Ingrese la contraseña correcta para continuar.")
    st.stop()

# Cargar datos
try:
    df = pd.read_excel(EXCEL_PATH)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.upper()
    st.write("🧾 Columnas detectadas en el Excel:", df.columns.tolist())

except FileNotFoundError:
    st.error("El archivo Excel no se encuentra. Sube el archivo correcto en la carpeta datos/")
    st.stop()

selected_rows = st.multiselect(
    "👥 Selecciona los docentes",
    df.index,
    format_func=lambda i: f"{df.loc[i, 'NOMBRE']} {df.loc[i, 'APELLIDO PATERNO']} {df.loc[i, 'APELLIDO MATERNO']}"
)

if selected_rows:
    st.write("✅ Docentes seleccionados:")
    st.write(df.loc[selected_rows])

num_oficio = st.text_input("📄 Número de Oficio")
sede = st.text_input("🏫 Sede")
ubicacion = st.text_input("📍 Ubicación")
fecha_comision = st.date_input("📅 Fecha de Comisión")
horario = st.text_input("🕒 Horario")
comision = st.text_input("🔖 Comisión")

# Generar oficios
if st.button("Generar Oficios"):
    if not selected_rows:
        st.warning("Selecciona al menos un docente.")
    elif not all([num_oficio, sede, ubicacion, fecha_comision, horario, comision]):
        st.warning("Por favor llena todos los campos.")
    else:
        data_to_process = df.loc[selected_rows]
        archivos_generados = generar_oficio(data_to_process, num_oficio, sede, ubicacion, fecha_comision, horario, comision)
        actualizar_historial(data_to_process, num_oficio, comision)
        zip_buffer = comprimir_archivos(archivos_generados)
        st.success("🎉 Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button("📥 Descargar Oficios (ZIP)", data=zip_buffer, file_name="oficios.zip", mime="application/zip")
        if os.path.exists(HISTORIAL_PATH):
            st.download_button("📥 Descargar Historial", data=open(HISTORIAL_PATH, "rb"), file_name="historial_oficios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
