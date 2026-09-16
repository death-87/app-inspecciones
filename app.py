import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="App de Inspecciones", page_icon="🏗️", layout="centered")

# ==========================================
# 2. CLASE PARA EL PDF CON LOGO
# ==========================================
class PDFReport(FPDF):
    def footer(self):
        # Posición a 30 mm desde el fondo de la página
        self.set_y(-30)
        
        # Intentar cargar el logo (debe llamarse 'logo.png' y estar en la misma carpeta)
        try:
            # x=85 centra aprox. una imagen de 40mm de ancho en una hoja A4
            self.image("logo.png", x=85, y=self.get_y(), w=40)
        except Exception:
            self.set_font("helvetica", "I", 8)
            self.cell(0, 10, "(Logo no encontrado - Sube 'logo.png' al repositorio)", align="C")
        
        # Texto debajo del logo
        self.set_y(-10)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Reporte oficial de inspección generado el {datetime.datetime.now().strftime('%d-%m-%Y')}", align="C")

# ==========================================
# 3. FUNCIONES DE EXPORTACIÓN
# ==========================================
def generar_pdf(inspector, detalles, observaciones, fecha):
    pdf = PDFReport()
    pdf.add_page()
    
    # Título principal
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, f"REPORTE DE INSPECCIÓN - {fecha}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(10)
    
    # Sección 1: Inspector
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185) # Títulos en azul
    pdf.cell(0, 10, "1. Información del Inspector:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, inspector)
    pdf.ln(5)
    
    # Sección 2: Detalles
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185)
    pdf.cell(0, 10, "2. Detalle de la Inspección:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, detalles)
    pdf.ln(5)
    
    # Sección 3: Observaciones
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185)
    pdf.cell(0, 10, "3. Observaciones:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, observaciones)
    
    return bytes(pdf.output())

def generar_excel(inspector, detalles, observaciones, fecha):
    datos = {
        "Fecha": [fecha],
        "Inspector": [inspector],
        "Detalle de Inspección": [detalles],
        "Observaciones": [observaciones]
    }
    df = pd.DataFrame(datos)
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Inspección Diaria')
        
        # Ajustar ancho de columnas para mejor visualización
        worksheet = writer.sheets['Inspección Diaria']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 60
        worksheet.column_dimensions['D'].width = 60
        
    return buffer.getvalue()

# ==========================================
# 4. INTERFAZ DE USUARIO (UI)
# ==========================================
def main():
    st.title("Control de Calidad e Inspección en Terreno")
    
    fecha_hoy = datetime.datetime.now().strftime("%d-%m-%Y")
    st.markdown(f"**Fecha del reporte:** {fecha_hoy}")

    st.subheader("Ingreso de Datos")
    
    # Cajas de texto para recolectar información
    inspector = st.text_input("Nombre del Inspector / Cargo", placeholder="Ej: Especialista QA/QC")
    detalles = st.text_area(
        "Detalle de la Inspección", 
        placeholder="Ej: Revisión de estándares de tuberías, cuantificación de materiales en sector B4...", 
        height=150
    )
    observaciones = st.text_area(
        "Observaciones", 
        placeholder="Ej: Material recibido conforme a especificaciones técnicas, sin desviaciones de tolerancia...", 
        height=100
    )

    st.markdown("---")
    st.subheader("Exportar Reporte")
    
    # Habilitar descargas solo si hay datos ingresados
    if inspector and detalles:
        col1, col2 = st.columns(2)
        
        # Generar los archivos en memoria
        pdf_bytes = generar_pdf(inspector, detalles, observaciones, fecha_hoy)
        excel_bytes = generar_excel(inspector, detalles, observaciones, fecha_hoy)
        
        with col1:
            st.download_button(
                label="📄 Descargar PDF (Con Logo)",
                data=pdf_bytes,
                file_name=f"Reporte_Inspeccion_{fecha_hoy}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
        with col2:
            st.download_button(
                label="📊 Descargar Control en Excel",
                data=excel_bytes,
                file_name=f"Control_Inspeccion_{fecha_hoy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("💡 Por favor, completa al menos el 'Nombre del Inspector' y el 'Detalle de la Inspección' para habilitar las opciones de descarga.")

# Ejecutar la aplicación
if __name__ == "__main__":
    main()
