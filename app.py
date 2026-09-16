import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF

# ==========================================
# 1. CONFIGURACIÓN INICIAL Y MEMORIA
# ==========================================
st.set_page_config(page_title="Control de Inspecciones", page_icon="🏗️", layout="wide")

# Inicializar el historial en la memoria temporal de la app si no existe
if 'historial' not in st.session_state:
    st.session_state.historial = []

# Lista de inspectores (puedes editar los nombres predeterminados)
if 'inspectores' not in st.session_state:
    st.session_state.inspectores = ["Seleccione un inspector...", "Juan Pérez", "Ana Silva", "Nuevo Inspector..."]

# ==========================================
# 2. CLASE PARA EL PDF CON LOGO
# ==========================================
class PDFReport(FPDF):
    def footer(self):
        self.set_y(-30)
        # Intenta colocar el logo en el centro inferior
        try:
            self.image("logo.png", x=85, y=self.get_y(), w=40)
        except Exception:
            self.set_font("helvetica", "I", 8)
            self.cell(0, 10, "(Logo no encontrado - Sube 'logo.png')", align="C")
        
        self.set_y(-10)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Reporte oficial generado el {datetime.datetime.now().strftime('%d-%m-%Y')}", align="C")

# ==========================================
# 3. FUNCIONES DE EXPORTACIÓN
# ==========================================
def generar_pdf_diario(inspector, detalles, observaciones, fecha):
    pdf = PDFReport()
    pdf.add_page()
    
    # Título principal
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, f"REPORTE DE INSPECCIÓN - {fecha}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(10)
    
    # Secciones
    secciones = [
        ("1. Información del Inspector:", inspector),
        ("2. Detalle de la Inspección:", detalles),
        ("3. Observaciones:", observaciones)
    ]
    
    for titulo, contenido in secciones:
        pdf.set_font("helvetica", "B", 12)
        pdf.set_text_color(41, 128, 185)
        pdf.cell(0, 10, titulo, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("helvetica", "", 11)
        pdf.set_text_color(0, 0, 0)
        pdf.multi_cell(0, 8, contenido)
        pdf.ln(5)
        
    return bytes(pdf.output())

def generar_excel_historial(historial_datos):
    # Si hay datos los convierte a DataFrame, si no, crea uno vacío con las columnas
    if not historial_datos:
        df = pd.DataFrame(columns=["Fecha", "Inspector", "Detalles", "Observaciones"])
    else:
        df = pd.DataFrame(historial_datos)
        
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Control Histórico')
        worksheet = writer.sheets['Control Histórico']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 25
        worksheet.column_dimensions['C'].width = 50
        worksheet.column_dimensions['D'].width = 40
        
    return buffer.getvalue()

# ==========================================
# 4. INTERFAZ DE LA APLICACIÓN
# ==========================================
st.title("Bitácora de Control de Calidad e Inspección")
fecha_hoy = datetime.datetime.now().strftime("%d-%m-%Y")

# Crear dos pestañas principales
tab1, tab2 = st.tabs(["📝 Registrar Actividad Diaria", "📊 Ver Historial y Exportar"])

# --- PESTAÑA 1: INGRESO DE DATOS ---
with tab1:
    st.subheader("Nueva Inspección en Terreno")
    
    # Selección o ingreso de inspector
    seleccion_inspector = st.selectbox("Inspector / Responsable QA-QC", st.session_state.inspectores)
    if seleccion_inspector == "Nuevo Inspector...":
        inspector_final = st.text_input("Ingresar nombre del nuevo inspector:")
    else:
        inspector_final = seleccion_inspector if seleccion_inspector != "Seleccione un inspector..." else ""
        
    detalles = st.text_area(
        "Detalle de la Inspección", 
        placeholder="Ej: Revisión de isométricos, cuantificación de materiales en sector B4, verificación de estándares de piping..."
    )
    observaciones = st.text_area(
        "Observaciones", 
        placeholder="Ej: Material recibido conforme a especificaciones. Schedule y tipos de bridas cumplen con las tolerancias del plano..."
    )
    
    col_btn1, col_btn2 = st.columns(2)
    
    # Botón para guardar en el historial
    with col_btn1:
        if st.button("💾 Guardar en Historial", type="primary", use_container_width=True):
            if inspector_final and detalles:
                nuevo_registro = {
                    "Fecha": fecha_hoy,
                    "Inspector": inspector_final,
                    "Detalles": detalles,
                    "Observaciones": observaciones
                }
                # Guardar en la memoria
                st.session_state.historial.append(nuevo_registro)
                
                # Si es un inspector nuevo, lo agregamos a la lista desplegable
                if seleccion_inspector == "Nuevo Inspector..." and inspector_final not in st.session_state.inspectores:
                    st.session_state.inspectores.insert(-1, inspector_final)
                    
                st.success("¡Registro guardado exitosamente en el historial!")
            else:
                st.error("Faltan datos. Completa al menos el inspector y los detalles de la inspección.")
                
    # Botón para descargar el PDF del día directamente
    with col_btn2:
        if inspector_final and detalles:
            pdf_data = generar_pdf_diario(inspector_final, detalles, observaciones, fecha_hoy)
            st.download_button(
                label="📄 Exportar este Reporte Diario (PDF)",
                data=pdf_data,
                file_name=f"Reporte_Diario_{fecha_hoy}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# --- PESTAÑA 2: HISTORIAL Y EXCEL ---
with tab2:
    st.subheader("Historial Acumulado de Inspecciones")
    
    if len(st.session_state.historial) > 0:
        # Mostrar la tabla en pantalla
        df_historial = pd.DataFrame(st.session_state.historial)
        st.dataframe(df_historial, use_container_width=True)
        
        st.markdown("---")
        st.write("### 📥 Exportar Base de Datos")
        
        # Descargar el Excel con todo el historial
        excel_data = generar_excel_historial(st.session_state.historial)
        st.download_button(
            label="📊 Descargar Historial Completo en Excel",
            data=excel_data,
            file_name=f"Control_Maestro_{fecha_hoy}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # Opción para limpiar historial si se necesita reiniciar
        if st.button("Limpiar Historial de la Sesión"):
            st.session_state.historial = []
            st.rerun()
    else:
        st.info("El historial está vacío. Ve a la pestaña 'Registrar Actividad Diaria' para agregar inspecciones.")
