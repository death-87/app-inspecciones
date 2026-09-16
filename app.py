import streamlit as st

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA (Siempre debe ir primero)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Gestión de Inspecciones",
    page_icon="📋",
    layout="wide"
)

# ---------------------------------------------------------
# 2. CONSTANTES Y DATOS FIJOS
# ---------------------------------------------------------
INSPECTORES = [
    "Juan Navarrete",
    "Arlem Sarmiento",
    "Miguel Chirinos",
    "Harold Castillo",
    "Quinto Inspector"  # Cambiar por el nombre correspondiente
]

# ---------------------------------------------------------
# 3. BARRA LATERAL (Sidebar)
# ---------------------------------------------------------
with st.sidebar:
    # Opción A: Mostrar el logo en la barra lateral
    # st.image("logo.png", use_container_width=True)
    st.title("Panel de Control")
    st.write("Selección de inspector:")
    inspector_sidebar = st.selectbox("Inspector en turno:", INSPECTORES, key="sb_inspector")

# ---------------------------------------------------------
# 4. ENCABEZADO Y LOGO EN EL CUERPO PRINCIPAL
# ---------------------------------------------------------
# Usamos columnas para colocar el logo al lado del título
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    # Opción B: Muestra el logo junto al título principal
    st.image("logo.png", width=120)

with col_titulo:
    st.title("Sistema de Registro de Inspecciones")
    st.caption("Control y seguimiento de calidad")

st.divider()

# ---------------------------------------------------------
# 5. CONTENIDO PRINCIPAL Y FORMULARIO
# ---------------------------------------------------------
st.header("Registrar Nueva Inspección")

inspector_seleccionado = st.selectbox("Seleccione el Inspector:", INSPECTORES)

st.write(f"**Inspector asignado:** {inspector_seleccionado}")