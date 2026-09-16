import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Gestión de Inspectores y Actividades QA/QC",
    page_icon="👷‍♂️",
    layout="wide"
)

DB_NAME = "gestion_inspectores.db"

# =========================================================
# INICIALIZACIÓN DE BASE DE DATOS
# =========================================================
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS registro_actividades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            inspector TEXT,
            tag_equipo TEXT,
            actividad_realizada TEXT,
            observaciones TEXT,
            estado_liberacion TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# =========================================================
# LISTA DE INSPECTORES Y CONFIGURACIÓN
# =========================================================
LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Harold Castillo",
    "Miguel Chirinos",
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
st.title("👷‍♂️ Sistema de Control de Actividades por Inspector")
st.markdown("---")

menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    ["📝 Registrar Actividad por Inspector", "📊 Historial y Registros Consolidados"]
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
                placeholder="Ej: Inspección visual de junta, verificación de Alineación/Torque, liberación de Hito..."
            )

        observaciones = st.text_area(
            "💬 Observaciones Adicionales / Recomendaciones:",
            placeholder="Escribe comentarios extra sobre la inspección..."
        )
        
        btn_guardar = st.form_submit_button("💾 Guardar Actividad en el Historial")
        
        if btn_guardar:
            if tag_equipo and actividad_realizada:
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                c.execute('''
                    INSERT INTO registro_actividades 
                    (fecha, inspector, tag_equipo, actividad_realizada, observaciones, estado_liberacion)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    str(fecha_actividad), 
                    inspector_seleccionado, 
                    tag_equipo.strip(), 
                    actividad_realizada.strip(), 
                    observaciones.strip(), 
                    estado_liberacion
                ))
                conn.commit()
                conn.close()
                st.success(f"✅ ¡Actividad de **{inspector_seleccionado}** registrada correctamente!")
            else:
                st.error("⚠️ Por favor completa los campos obligatorios: **TAG del Equipo** y **Actividades Realizadas**.")

# =========================================================
# MODULO 2: HISTORIAL REGISTRADO
# =========================================================
elif menu == "📊 Historial y Registros Consolidados":
    st.subheader("🔍 Consulta e Historial de Actividades Registradas")
    
    conn = sqlite3.connect(DB_NAME)
    df_historial = pd.read_sql_query("SELECT id, fecha, inspector, tag_equipo, actividad_realizada, observaciones, estado_liberacion FROM registro_actividades ORDER BY id DESC", conn)
    conn.close()
    
    if not df_historial.empty:
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
            df_filtrado = df_filtrado[df_filtrado['tag_equipo'].str.contains(filtro_tag, case=False, na=False)]
            
        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['estado_liberacion'] == filtro_estado]

        st.markdown(f"**Total de registros encontrados:** `{len(df_filtrado)}`")
        
        # Mostrar tabla interactiva
        st.dataframe(
            df_filtrado, 
            use_container_width=True,
            column_config={
                "id": "ID",
                "fecha": "Fecha",
                "inspector": "Inspector",
                "tag_equipo": "TAG Equipo",
                "actividad_realizada": "Actividad Realizada",
                "observaciones": "Observaciones",
                "estado_liberacion": "Estado / Liberación"
            }
        )
        
        # Botón de exportación a Excel / CSV
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Exportar Historial Seleccionado a CSV / Excel",
            data=csv_data,
            file_name=f"historial_actividades_inspectores_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de actividades guardados en la base de datos.")
