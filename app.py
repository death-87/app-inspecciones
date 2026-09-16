import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF

# --- CLASE PARA EL PDF CON LOGO EN EL PIE DE PÁGINA ---
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
            self.cell(0, 10, "(Logo no encontrado - Agrega 'logo.png' a la carpeta)", align="C")
        
        # Texto debajo del logo
        self.set_y(-10)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Reporte oficial generado el {datetime.datetime.now().strftime('%d-%m-%Y')}", align="C")

# --- FUNCIONES GENERADORAS ---
def generar_pdf(inspector, detalles, observaciones, fecha):
    pdf = PDFReport()
    pdf.add_page()
    
    # Título
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, f"REPORTE DE INSPECCIÓN - {fecha}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(10)
    
    # Inspector
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185) # Color azul opcional para subtítulos
    pdf.cell(0, 10, "1. Información del Inspector:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, inspector)
    pdf.ln(5)
    
    # Detalles
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185)
    pdf.cell(0, 10, "2. Detalle de la Inspección:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, detalles)
    pdf.ln(5)
    
    # Observaciones
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(41, 128, 185)
    pdf.cell(0, 10, "3. Observaciones:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, observaciones)
    
    # Retornar los bytes del PDF
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
        df.to_excel(writer, index=False, sheet_name='Control Diario')
        
        # Ajustar ancho de columnas automáticamente para que se lea mejor
        worksheet = writer.sheets['Control Diario']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 25
        worksheet.column_dimensions['C'].width = 50
        worksheet.column_dimensions['D'].width = 50
        
    return buffer.getvalue()


# --- INTERFAZ DE STREAMLIT ---
st.set_page_config(page_title="Control de Inspección", page_icon="📋")
st.title("Generador de Reportes de Inspección")

fecha_hoy = datetime.datetime.now().strftime("%d-%m-%Y")
st.write(f"**Fecha del reporte:** {fecha_hoy}")

# Inputs del usuario
inspector = st.text_input("Nombre del Inspector")
detalles = st.text_area("Detalle de la Inspección (Materiales, estándares, etc.)", height=150)
observaciones = st.text_area("Observaciones", height=100)

st.markdown("---")
col1, col2 = st.columns(2)

if inspector and detalles:
    # Botón PDF
    pdf_bytes = generar_pdf(inspector, detalles, observaciones, fecha_hoy)
    col1.download_button(
        label="📄 Descargar PDF (Con Logo)",
        data=pdf_bytes,
        file_name=f"Reporte_Inspeccion_{fecha_hoy}.pdf",
        mime="application/pdf",
        use_container_width=True
    )
    
    # Botón Excel
    excel_bytes = generar_excel(inspector, detalles, observaciones, fecha_hoy)
    col2.download_button(
        label="📊 Descargar Control en Excel",
        data=excel_bytes,
        file_name=f"Control_Diario_{fecha_hoy}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("Por favor completa el nombre del inspector y los detalles para habilitar las opciones de descarga.")
