import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Gestión de Inspectores en la Nube",
    page_icon="☁️",
    layout="wide"
)

# 🔗 PEGA AQUÍ LA URL COMPLETA DE TU HOJA DE GOOGLE SHEETS
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I/edit?usp=sharing"

# =========================================================
# CONEXIÓN CON GOOGLE SHEETS
# =========================================================
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_datos_sheets():
    """Lee las filas almacenadas en la Hoja de Google."""
    try:
        # ttl=0 fuerza a leer los datos actualizados de la nube sin usar caché
        return conn.read(spreadsheet=URL_GOOGLE_SHEETS, ttl=0)
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        return pd.DataFrame(columns=[
            "fecha", "inspector", "tag_equipo", 
            "actividad_realizada", "observaciones", "estado_liberacion"
        ])

# =========================================================
# LISTA DE INSPECTORES Y CONFIGURACIÓN
# =========================================================
LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Miguel Chirinos",
    "Harold Castillo",
    "Arlem Sarmiento"
]

ESTADOS_LIBERACION = [
    "Liberado / Conforme (Aprobado)",
    "Pendiente de Reparación",
    "Rechazado",
    "En Proceso de Inspección",
    "En Espera de END / Pruebas"
]

# =========================================================
# INTERFAZ Y NAVEGACIÓN
# =========================================================
st.title("☁️ Control de Actividades por Inspector (Sincronizado en la Nube)")
st.markdown("---")

menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    ["📝 Registrar Actividad por Inspector", "📊 Historial en la Nube"]
)

# =========================================================
# MODULO 1: REGISTRO DE ACTIVIDADES
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
        
        btn_guardar = st.form_submit_button("☁️ Guardar en Google Sheets")
        
        if btn_guardar:
            if tag_equipo and actividad_realizada:
                try:
                    # 1. Cargar el DataFrame actual de la nube
                    df_actual = cargar_datos_sheets()
                    
                    # 2. Crear la nueva fila
                    nueva_fila = pd.DataFrame([{
                        "fecha": str(fecha_actividad),
                        "inspector": inspector_seleccionado,
                        "tag_equipo": tag_equipo.strip(),
                        "actividad_realizada": actividad_realizada.strip(),
                        "observaciones": observaciones.strip(),
                        "estado_liberacion": estado_liberacion
                    }])
                    
                    # 3. Concatenar y actualizar la hoja en Google Sheets
                    df_actualizado = pd.concat([df_actual, nueva_fila], ignore_index=True)
                    conn.update(spreadsheet=URL_GOOGLE_SHEETS, data=df_actualizado)
                    
                    st.success(f"✅ ¡Actividad de **{inspector_seleccionado}** guardada correctamente en la Nube!")
                except Exception as ex:
                    st.error(f"❌ Ocurrió un error al guardar en la nube: {ex}")
            else:
                st.error("⚠️ Por favor completa los campos obligatorios: **TAG del Equipo** y **Actividades Realizadas**.")

# =========================================================
# MODULO 2: HISTORIAL EN LA NUBE
# =========================================================
elif menu == "📊 Historial en la Nube":
    st.subheader("🔍 Consulta e Historial de Actividades Registradas")
    
    with st.spinner("Cargando registros desde Google Sheets..."):
        df_historial = cargar_datos_sheets()
    
    if not df_historial.empty and "inspector" in df_historial.columns:
        # Filtros de búsqueda
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

        # Aplicar filtros
        df_filtrado = df_historial.copy()
        
        if filtro_inspector != "Todos":
            df_filtrado = df_filtrado[df_filtrado['inspector'] == filtro_inspector]
            
        if filtro_tag:
            df_filtrado = df_filtrado[df_filtrado['tag_equipo'].astype(str).str.contains(filtro_tag, case=False, na=False)]
            
        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['estado_liberacion'] == filtro_estado]

        st.markdown(f"**Total de registros en la nube:** `{len(df_filtrado)}`")
        
        # Mostrar tabla interactiva
        st.dataframe(
            df_filtrado, 
            use_container_width=True
        )
        
        # Botón de exportación local
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar copia local a CSV",
            data=csv_data,
            file_name=f"historial_nube_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de actividades guardados en la hoja de Google Sheets.")
