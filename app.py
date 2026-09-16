import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Gestión de Inspectores en la Nube",
    page_icon="☁️",
    layout="wide"
)

# 🔗 URL RAW DE TU LOGO EN GITHUB
URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"

# 🆔 ID DE TU HOJA DE GOOGLE SHEETS
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

# =========================================================
# CONEXIÓN DIRECTA CON GOOGLE SHEETS VIA GSPREAD
# =========================================================
def conectar_google_sheets():
    """Autentica y devuelve la hoja de trabajo activa."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    
    if "private_key" in creds_dict:
        pk = creds_dict["private_key"]
        pk = pk.replace("\\n", "\n").strip()
        lines = [line.strip() for line in pk.split("\n") if line.strip()]
        creds_dict["private_key"] = "\n".join(lines) + "\n"

    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID).sheet1

def cargar_datos_sheets():
    """Lee todas las filas almacenadas en la Hoja de forma directa."""
    try:
        sheet = conectar_google_sheets()
        filas = sheet.get_all_values()
        
        if len(filas) <= 1:
            # Solo está la cabecera o está vacía
            return pd.DataFrame(columns=[
                "fecha", "inspector", "tag_equipo", 
                "actividad_realizada", "observaciones", "estado_liberacion"
            ])
            
        # La primera fila son los encabezados
        encabezados = [c.strip().lower() for c in filas[0]]
        datos = filas[1:]
        
        df = pd.DataFrame(datos, columns=encabezados)
        return df
    except Exception as e:
        st.error(f"Error al leer la hoja de Google Sheets: {e}")
        return pd.DataFrame(columns=[
            "fecha", "inspector", "tag_equipo", 
            "actividad_realizada", "observaciones", "estado_liberacion"
        ])

# =========================================================
# LISTA DE INSPECTORES Y CONFIGURACIÓN DE ESTADOS
# =========================================================
LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Inspector 3",
    "Inspector 4",
    "Inspector 5"
]

ESTADOS_LIBERACION = [
    "Liberado / Conforme (Aprobado)",
    "Pendiente de Reparación",
    "Rechazado",
    "En Proceso de Inspección",
    "En Espera de END / Pruebas"
]

# =========================================================
# ENCABEZADO Y VISUALIZACIÓN DEL LOGO
# =========================================================
try:
    st.sidebar.image(URL_LOGO_GITHUB, use_container_width=True)
except Exception:
    pass

col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
        st.image(URL_LOGO_GITHUB, width=140)
    except Exception:
        st.write("📂 [Logo]")

with col_titulo:
    st.title("Control de Actividades por Inspector")
    st.markdown("##### *Sistema de Gestión QA/QC Sincronizado en la Nube*")

st.markdown("---")

# =========================================================
# MENÚ Y NAVEGACIÓN
# =========================================================
menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    ["📝 Registrar Actividad por Inspector", "📊 Historial en la Nube"]
)

# =========================================================
# MÓDULO 1: REGISTRO DE ACTIVIDADES (ESCRITURA)
# =========================================================
if menu == "📝 Registrar Actividad por Inspector":
    st.subheader("📋 Formulario de Ingreso de Actividades")
    
    with st.form("form_actividades_inspector", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            inspector_seleccionado = st.selectbox(
                "👷‍♂️ Seleccionar Inspector asignado:",
                LISTA_INSPECTORES
            )
            fecha_actividad = st.date_input("📅 Fecha de Inspección:", datetime.now())
            tag_equipo = st.text_input("🏷️ TAG del Equipo / Línea Piping:", placeholder="Ej: C-1302 / E-2101 / PIP-001")

        with col2:
            estado_liberacion = st.selectbox("📌 Estado de la Inspección:", ESTADOS_LIBERACION)
            actividad_realizada = st.text_area(
                "🛠️ Actividades Realizadas por el Inspector:",
                placeholder="Ej: Inspección visual de junta, verificación de Alineación/Torque..."
            )

        observaciones = st.text_area(
            "💬 Observaciones Adicionales / Recomendaciones:",
            placeholder="Escribe comentarios extra sobre la inspección..."
        )
        
        btn_guardar = st.form_submit_button("☁️ Guardar Registro en Google Sheets")
        
        if btn_guardar:
            if tag_equipo and actividad_realizada:
                try:
                    sheet = conectar_google_sheets()
                    
                    # Verificamos si la hoja está completamente vacía para agregar cabecera primero
                    todas_las_filas = sheet.get_all_values()
                    if len(todas_las_filas) == 0:
                        sheet.append_row([
                            "fecha", "inspector", "tag_equipo", 
                            "actividad_realizada", "observaciones", "estado_liberacion"
                        ])
                    
                    nueva_fila = [
                        str(fecha_actividad),
                        inspector_seleccionado,
                        tag_equipo.strip(),
                        actividad_realizada.strip(),
                        observaciones.strip(),
                        estado_liberacion
                    ]
                    
                    sheet.append_row(nueva_fila)
                    
                    st.success(f"✅ ¡Actividad de **{inspector_seleccionado}** guardada con éxito en Google Sheets!")
                except Exception as ex:
                    st.error(f"❌ Ocurrió un error al guardar en la nube: {ex}")
            else:
                st.error("⚠️ Por favor completa los campos obligatorios: **TAG del Equipo** y **Actividades Realizadas**.")

# =========================================================
# MÓDULO 2: HISTORIAL EN LA NUBE (LECTURA Y FILTROS)
# =========================================================
elif menu == "📊 Historial en la Nube":
    st.subheader("🔍 Consulta e Historial de Actividades Registradas")
    
    with st.spinner("Cargando registros desde Google Sheets..."):
        df_historial = cargar_datos_sheets()
    
    if not df_historial.empty:
        col_filtro1, col_filtro2, col_filtro3 = st.columns(3)
        
        with col_filtro1:
            filtro_inspector = st.selectbox(
                "Filtrar por Inspector:",
                ["Todos"] + LISTA_INSPECTORES
            )
        with col_filtro2:
            filtro_tag = st.text_input("Filtrar por TAG de Equipo:")
            
        with col_filtro3:
            filtro_estado = st.selectbox(
                "Filtrar por Estado:",
                ["Todos"] + ESTADOS_LIBERACION
            )

        df_filtrado = df_historial.copy()
        
        if "inspector" in df_filtrado.columns and filtro_inspector != "Todos":
            df_filtrado = df_filtrado[df_filtrado['inspector'] == filtro_inspector]
            
        if "tag_equipo" in df_filtrado.columns and filtro_tag:
            df_filtrado = df_filtrado[df_filtrado['tag_equipo'].astype(str).str.contains(filtro_tag, case=False, na=False)]
            
        if "estado_liberacion" in df_filtrado.columns and filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['estado_liberacion'] == filtro_estado]

        st.markdown(f"**Total de registros encontrados:** `{len(df_filtrado)}`")
        
        st.dataframe(
            df_filtrado, 
            use_container_width=True
        )
        
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar copia local a CSV",
            data=csv_data,
            file_name=f"historial_actividades_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de actividades guardados en la nube.")
