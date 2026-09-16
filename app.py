import streamlit as st
import pandas as pd
import datetime
import sqlite3
import os

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Sistema de Inspecciones Diarias",
    page_icon="📋",
    layout="wide"
)

# Mostrar logo si existe sin bloquear la app si falta
if os.path.exists("logo.png"):
    st.image("logo.png", width=150)

# =========================================================
# BASE DE DATOS (PERSISTENCIA DE DATOS)
# =========================================================
def inicializar_db():
    conn = sqlite3.connect("inspecciones_db.db")
    cursor = conn.cursor()
    
    # Tabla Encabezado de Reportes
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reportes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inspector TEXT,
            planta TEXT,
            equipo TEXT,
            semana INTEGER,
            fecha_inicio TEXT,
            fecha_fin TEXT,
            hora_inicio TEXT,
            hora_fin TEXT,
            fecha_registro TEXT
        )
    """)
    
    # Tabla Detalle de Actividades
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS actividades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            reporte_id INTEGER,
            num_actividad INTEGER,
            descripcion TEXT,
            porcentaje_avance INTEGER,
            hallazgo TEXT,
            observaciones TEXT,
            FOREIGN KEY (reporte_id) REFERENCES reportes (id)
        )
    """)
    conn.commit()
    conn.close()

inicializar_db()

# =========================================================
# MENÚ NAVEGACIÓN LATERAL
# =========================================================
st.sidebar.title("📌 Menú Principal")
opcion_menu = st.sidebar.radio("Ir a:", ["📝 Nuevo Reporte Diario", "📊 Historial y Consultas"])

INSPECTORES = [
    "Inspector 1", 
    "Inspector 2", 
    "Inspector 3", 
    "Inspector 4", 
    "Inspector 5"
]

# =========================================================
# MÓDULO 1: INGRESAR REPORTE DIARIO
# =========================================================
if opcion_menu == "📝 Nuevo Reporte Diario":
    st.header("📋 Registro Diario de Inspección Técnica")
    st.info("Ingresa los datos del turno y añade las actividades ejecutadas.")

    # 1. Datos del Inspector y Controles
    col_insp, col_planta, col_equipo = st.columns(3)
    
    with col_insp:
        inspector = st.selectbox("Inspector a cargo:", INSPECTORES)
    with col_planta:
        planta = st.text_input("Planta:", placeholder="Ej: Planta Concentradora")
    with col_equipo:
        equipo = st.text_input("Equipo / TAG:", placeholder="Ej: C705 / Bomba 01")

    # 2. Control Temporal
    col_sem, col_f_ini, col_f_fin, col_h_ini, col_h_fin = st.columns([1, 2, 2, 2, 2])
    
    fecha_actual = datetime.date.today()
    semana_actual = int(fecha_actual.strftime("%V"))

    with col_sem:
        semana = st.number_input("Semana N°:", min_value=1, max_value=53, value=semana_actual)
    with col_f_ini:
        fecha_inicio = st.date_input("Fecha Inicio:", value=fecha_actual)
    with col_f_fin:
        fecha_fin = st.date_input("Fecha Término:", value=fecha_actual)
    with col_h_ini:
        hora_inicio = st.time_input("Hora Inicio:", value=datetime.time(8, 0))
    with col_h_fin:
        hora_fin = st.time_input("Hora Término:", value=datetime.time(18, 0))

    st.divider()
    st.subheader("⚙️ Enumeración de Actividades")

    # Inicializar lista dinámica en sesión
    if "lista_actividades" not in st.session_state:
        st.session_state.lista_actividades = []

    # Botones de control de actividades
    col_b1, col_b2 = st.columns([2, 8])
    with col_b1:
        if st.button("➕ Agregar Actividad", use_container_width=True):
            st.session_state.lista_actividades.append({
                "num": len(st.session_state.lista_actividades) + 1,
                "descripcion": "",
                "avance": 0,
                "hallazgo": "",
                "observaciones": ""
            })

    # Formularios dinámicos para cada actividad
    for idx, act in enumerate(st.session_state.lista_actividades):
        with st.expander(f"🔹 Actividad N° {act['num']}", expanded=True):
            c1, c2 = st.columns([3, 1])
            with c1:
                act["descripcion"] = st.text_input(f"Descripción de la actividad", key=f"desc_{idx}", value=act["descripcion"])
            with c2:
                act["avance"] = st.slider(f"% Avance", 0, 100, key=f"av_{idx}", value=act["avance"])
            
            c3, c4 = st.columns(2)
            with c3:
                act["hallazgo"] = st.text_area(f"Hallazgo", key=f"hal_{idx}", value=act["hallazgo"], height=70)
            with c4:
                act["observaciones"] = st.text_area(f"Observaciones", key=f"obs_{idx}", value=act["observaciones"], height=70)

    st.divider()

    # Botón Final de Guardado
    if st.button("💾 Guardar Reporte Diario en la Web", type="primary", use_container_width=True):
        if not planta or not equipo:
            st.warning("⚠️ Por favor completa los campos de Planta y Equipo.")
        elif len(st.session_state.lista_actividades) == 0:
            st.warning("⚠️ Debes agregar al menos 1 actividad antes de guardar.")
        else:
            # Guardar en Base de Datos
            conn = sqlite3.connect("inspecciones_db.db")
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO reportes (inspector, planta, equipo, semana, fecha_inicio, fecha_fin, hora_inicio, hora_fin, fecha_registro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                inspector, planta, equipo, semana, 
                str(fecha_inicio), str(fecha_fin), 
                str(hora_inicio), str(hora_fin), 
                str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ))
            
            reporte_id = cursor.lastrowid
            
            for act in st.session_state.lista_actividades:
                cursor.execute("""
                    INSERT INTO actividades (reporte_id, num_actividad, descripcion, porcentaje_avance, hallazgo, observaciones)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    reporte_id, act["num"], act["descripcion"], act["avance"], act["hallazgo"], act["observaciones"]
                ))
                
            conn.commit()
            conn.close()
            
            st.success(f"🎉 ¡Reporte N° {reporte_id} de {inspector} guardado correctamente!")
            st.session_state.lista_actividades = [] # Limpiar actividades después de guardar

# =========================================================
# MÓDULO 2: HISTORIAL Y CONSULTAS POR DÍA / SEMANA
# =========================================================
elif opcion_menu == "📊 Historial y Consultas":
    st.header("📂 Historial de Reportes e Inspecciones")
    
    conn = sqlite3.connect("inspecciones_db.db")
    df_reportes = pd.read_sql_query("SELECT * FROM reportes", conn)
    
    if df_reportes.empty:
        st.info("Aún no existen reportes registrados en el sistema.")
    else:
        # Filtros
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_semana = st.multiselect("Filtrar por Semana:", sorted(df_reportes["semana"].unique()))
        with col_f2:
            filtro_inspector = st.multiselect("Filtrar por Inspector:", df_reportes["inspector"].unique())
        with col_f3:
            filtro_planta = st.multiselect("Filtrar por Planta:", df_reportes["planta"].unique())

        # Aplicar Filtros
        df_filtrado = df_reportes.copy()
        if filtro_semana:
            df_filtrado = df_filtrado[df_filtrado["semana"].isin(filtro_semana)]
        if filtro_inspector:
            df_filtrado = df_filtrado[df_filtrado["inspector"].isin(filtro_inspector)]
        if filtro_planta:
            df_filtrado = df_filtrado[df_filtrado["planta"].isin(filtro_planta)]

        st.subheader("Reportes Encontrados")
        st.dataframe(df_filtrado, use_container_width=True)

        # Ver detalle de actividades por reporte seleccionado
        st.divider()
        reporte_sel = st.selectbox("Selecciona un ID de Reporte para ver el detalle de actividades:", df_filtrado["id"].unique())
        
        if reporte_sel:
            df_actividades = pd.read_sql_query(f"SELECT num_actividad as N°, descripcion as 'Descripción', porcentaje_avance as '% Avance', hallazgo as 'Hallazgo', observaciones as 'Observaciones' FROM actividades WHERE reporte_id = {reporte_sel}", conn)
            st.markdown(f"**Detalle de Actividades del Reporte ID #{reporte_sel}:**")
            st.table(df_actividades)

    conn.close()
