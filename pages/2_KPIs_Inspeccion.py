import streamlit as st
import pandas as pd
import datetime

# Configuración de la página
st.set_page_config(page_title="KPIs de Inspección", page_icon="📊", layout="wide")

# ==========================================
# 1. LISTA OFICIAL DE INSPECTORES
# ==========================================
# Copia aquí exactamente los mismos 5 nombres que tienes en tu app.py 
# en la opción "Seleccionar Inspector asignado"
INSPECTORES_OFICIALES = [
    "Juan Navarrete", 
    "Jorge Hernandez", 
    "Arlem Sarmiento", 
    "Harold Castillo", 
    "Miguel Chirinos"
]

# ==========================================
# 2. INICIALIZACIÓN DE DATOS BASE
# ==========================================
# Base de datos temporal en memoria (session_state)
if "inspecciones_data" not in st.session_state:
    st.session_state.inspecciones_data = pd.DataFrame([
        {"Fecha": "2026-09-15", "Semana": 38, "Inspector": INSPECTORES_OFICIALES[0], "Equipo_TAG": "C701", "Horas": 4.0, "Estado": "Conforme", "Hallazgos": 0},
        {"Fecha": "2026-09-15", "Semana": 38, "Inspector": INSPECTORES_OFICIALES[1], "Equipo_TAG": "C702", "Horas": 3.5, "Estado": "Con Hallazgos", "Hallazgos": 2},
        {"Fecha": "2026-09-16", "Semana": 38, "Inspector": INSPECTORES_OFICIALES[2], "Equipo_TAG": "E101", "Horas": 5.0, "Estado": "Conforme", "Hallazgos": 0},
        {"Fecha": "2026-09-16", "Semana": 38, "Inspector": INSPECTORES_OFICIALES[3], "Equipo_TAG": "E102", "Horas": 2.5, "Estado": "Rechazado", "Hallazgos": 3},
        {"Fecha": "2026-09-17", "Semana": 38, "Inspector": INSPECTORES_OFICIALES[4], "Equipo_TAG": "C701", "Horas": 4.5, "Estado": "Conforme", "Hallazgos": 0}
    ])

st.title("📊 Control de Actividades e Indicadores de Inspección (KPIs)")

# ==========================================
# 3. BARRA LATERAL: REGISTRO Y FILTROS
# ==========================================
st.sidebar.header("📝 Registro de Trabajo Semanal")

with st.sidebar.form("form_registro_inspeccion"):
    fecha = st.date_input("Fecha de Inspección", datetime.date.today())
    inspector = st.selectbox("Seleccionar Inspector asignado", INSPECTORES_OFICIALES)
    semana = st.number_input("Semana Nº", min_value=1, max_value=53, value=datetime.date.today().isocalendar()[1])
    equipo = st.text_input("TAG de Equipo", value="C701")
    horas = st.number_input("Horas Invertidas", min_value=0.5, max_value=24.0, value=4.0, step=0.5)
    estado = st.selectbox("Estado de Inspección", ["Conforme", "Con Hallazgos", "Rechazado"])
    hallazgos = st.number_input("Cantidad de Hallazgos", min_value=0, max_value=20, value=0)
    
    btn_guardar = st.form_submit_button("➕ Registrar Actividad")
    
    if btn_guardar:
        nuevo_registro = pd.DataFrame([{
            "Fecha": str(fecha),
            "Semana": int(semana),
            "Inspector": inspector,
            "Equipo_TAG": equipo,
            "Horas": float(horas),
            "Estado": estado,
            "Hallazgos": int(hallazgos)
        }])
        st.session_state.inspecciones_data = pd.concat([st.session_state.inspecciones_data, nuevo_registro], ignore_index=True)
        st.success("Actividad registrada exitosamente.")
        st.rerun()

st.sidebar.divider()
st.sidebar.header("🔍 Filtros del Dashboard")
semanas_avail = sorted(st.session_state.inspecciones_data["Semana"].unique())
filtro_semana = st.sidebar.multiselect("Filtrar por Semana", options=semanas_avail, default=semanas_avail)
filtro_inspector = st.sidebar.multiselect("Filtrar por Inspector", options=INSPECTORES_OFICIALES, default=INSPECTORES_OFICIALES)

# ==========================================
# 4. APLICACIÓN DE FILTROS A LA DATA
# ==========================================
df_filtrado = st.session_state.inspecciones_data[
    (st.session_state.inspecciones_data["Semana"].isin(filtro_semana)) &
    (st.session_state.inspecciones_data["Inspector"].isin(filtro_inspector))
]

# ==========================================
# 5. TARJETAS DE KPIS PRINCIPALES
# ==========================================
col1, col2, col3, col4 = st.columns(4)

total_insp = len(df_filtrado)
total_hrs = df_filtrado["Horas"].sum() if not df_filtrado.empty else 0.0
promedio_hrs = (total_hrs / total_insp) if total_insp > 0 else 0.0
con_hallazgos = len(df_filtrado[df_filtrado["Estado"] != "Conforme"])
tasa_rechazo = (con_hallazgos / total_insp * 100) if total_insp > 0 else 0.0

col1.metric("Inspecciones Totales", f"{total_insp}")
col2.metric("Horas Hombre Totales", f"{total_hrs:.1f} hrs")
col3.metric("Promedio Horas/Equipo", f"{promedio_hrs:.1f} h/eq")
col4.metric("Tasa de No Conformidad", f"{tasa_rechazo:.1f}%", delta_color="inverse")

st.divider()

# ==========================================
# 6. GRÁFICOS COMPARATIVOS ENTRE INSPECTORES
# ==========================================
if not df_filtrado.empty:
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        st.subheader("📌 Inspecciones por Inspector")
        grafico_insp = df_filtrado.groupby("Inspector")["Equipo_TAG"].count().reset_index()
        grafico_insp.columns = ["Inspector", "Inspecciones Realizadas"]
        st.bar_chart(grafico_insp, x="Inspector", y="Inspecciones Realizadas")

    with col_g2:
        st.subheader("⏱️ Horas Invertidas por Inspector")
        grafico_hrs = df_filtrado.groupby("Inspector")["Horas"].sum().reset_index()
        st.bar_chart(grafico_hrs, x="Inspector", y="Horas")

    st.divider()

    # ==========================================
    # 7. TABLA DETALLADA Y EXPORTACIÓN A EXCEL/CSV
    # ==========================================
    st.subheader("📋 Registro Detallado de Actividades")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

    # Botón de descarga
    csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Reporte en CSV",
        data=csv_data,
        file_name="reporte_inspecciones_semana.csv",
        mime="text/csv"
    )
else:
    st.warning("No hay datos disponibles para los filtros seleccionados. Intenta modificar la semana o los inspectores en el panel izquierdo.")
