import io
import re
import requests
import streamlit as st
from datetime import date
from io import BytesIO

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Generar Informe Visual",
    page_icon="📄",
    layout="wide"
)

URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"

@st.cache_data(ttl=3600)
def obtener_bytes_imagen(url):
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return response.content
    except Exception:
        pass
    return None

def obtener_numero_archivo(nombre_archivo):
    """Extrae el número inicial del nombre del archivo para ordenar las imágenes."""
    match = re.search(r'^(\d+)', nombre_archivo)
    return int(match.group(1)) if match else 9999

# =========================================================
# FUNCIONES GENERADORAS DE ARCHIVOS (DOCX Y PDF)
# =========================================================
def generar_word_plantilla_inspeccion(datos_encabezado, secciones, imagenes_procesadas):
    doc = Document()
    
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    title = doc.add_paragraph("INFORME DE INSPECCIÓN VISUAL")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.bold = True
    title.runs[0].font.size = Pt(16)
    
    subtitle = doc.add_paragraph("CONTROL DE INSPECCIÓN • CALIDAD • TRAZABILIDAD")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(10)
    subtitle.runs[0].font.italic = True

    doc.add_paragraph()

    # Tabla de Datos de Encabezado
    table = doc.add_table(rows=5, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    
    fields = [
        ("N.º DE INFORME", datos_encabezado['num_informe'], "OT", datos_encabezado['ot']),
        ("FECHA", datos_encabezado['fecha'].strftime("%d/%m/%Y"), "UNIDAD", datos_encabezado['unidad']),
        ("TAG", datos_encabezado['tag'], "DESCRIPCIÓN", datos_encabezado['descripcion']),
        ("ACA", datos_encabezado['aca'], "MOTIVO", datos_encabezado['motivo']),
    ]
    
    for row_idx, (k1, v1, k2, v2) in enumerate(fields):
        row = table.rows[row_idx]
        row.cells[0].paragraphs[0].add_run(k1).bold = True
        row.cells[1].paragraphs[0].text = str(v1)
        row.cells[2].paragraphs[0].add_run(k2).bold = True
        row.cells[3].paragraphs[0].text = str(v2)

    row_alcance = table.rows[4]
    row_alcance.cells[0].paragraphs[0].add_run("ALCANCE").bold = True
    cell_span = row_alcance.cells[1]
    for cell in [row_alcance.cells[2], row_alcance.cells[3]]:
        cell_span.merge(cell)
    cell_span.paragraphs[0].text = str(datos_encabezado['alcance'])

    doc.add_paragraph()

    # Secciones dinámicas
    for titulo_sec, subsecciones in secciones.items():
        doc.add_heading(titulo_sec, level=1)
        for sub_tit, contenido in subsecciones.items():
            doc.add_heading(sub_tit, level=2)
            doc.add_paragraph(contenido if contenido else "N/A")

    # Registros fotográficos
    doc.add_heading("5. REGISTROS FOTOGRÁFICOS", level=1)
    
    if imagenes_procesadas:
        img_table = doc.add_table(rows=0, cols=2)
        img_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(imagenes_procesadas), 2):
            row_cells = img_table.add_row().cells
            
            # Imagen 1
            img_data1, label1 = imagenes_procesadas[i]
            p1 = row_cells[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(BytesIO(img_data1), width=Inches(3.0))
            p1_sub = row_cells[0].add_paragraph(str(label1))
            p1_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Imagen 2
            if i + 1 < len(imagenes_procesadas):
                img_data2, label2 = imagenes_procesadas[i+1]
                p2 = row_cells[1].paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(BytesIO(img_data2), width=Inches(3.0))
                p2_sub = row_cells[1].add_paragraph(str(label2))
                p2_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def dibujar_plantilla(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#f8faf6"))
    canvas.rect(0, 0, letter[0], letter[1], fill=1, stroke=0)

    canvas.setFillColor(colors.HexColor("#619b40"))
    canvas.rect(0, letter[1] - (2 * cm), letter[0], 2 * cm, fill=1, stroke=0)

    logo_bytes = obtener_bytes_imagen(URL_LOGO_GITHUB)
    if logo_bytes:
        try:
            img_stream = BytesIO(logo_bytes)
            img = ImageReader(img_stream)
            canvas.drawImage(img, 30, letter[1] - (2 * cm) - 55, width=120, height=45, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#6B7280"))
    canvas.drawCentredString(letter[0] / 2.0, 26, "SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.")
    canvas.drawCentredString(letter[0] / 2.0, 16, "CONTRATO N° AC 31104857")
    canvas.drawRightString(letter[0] - 30, 16, f"Pág. {canvas.getPageNumber()}")
    canvas.restoreState()

def generar_pdf_plantilla_inspeccion(datos_encabezado, secciones, imagenes_procesadas):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=4.5 * cm,
        bottomMargin=2.2 * cm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=15, leading=18, textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=2)
    subtitle_style = ParagraphStyle('DocSubTitle', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=10)
    sec_heading_style = ParagraphStyle('SecHeader', parent=styles['Heading2'], fontSize=11, leading=13, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=4)
    subsec_heading_style = ParagraphStyle('SubSecHeader', parent=styles['Heading3'], fontSize=9.5, leading=11, textColor=colors.HexColor('#619b40'), fontName='Helvetica-Bold', spaceBefore=4, spaceAfter=2)
    text_style = ParagraphStyle('TextStyle', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#1F2937'), spaceAfter=6)
    cell_body = ParagraphStyle('CB', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))

    story = [
        Paragraph("INFORME DE INSPECCIÓN VISUAL", title_style),
        Paragraph("CONTROL DE INSPECCIÓN • CALIDAD • TRAZABILIDAD", subtitle_style),
        Spacer(1, 4)
    ]

    # Tabla Encabezado
    data_enc = [
        [Paragraph("<b>N.º DE INFORME:</b>", cell_body), Paragraph(str(datos_encabezado['num_informe']), cell_body), Paragraph("<b>OT:</b>", cell_body), Paragraph(str(datos_encabezado['ot']), cell_body)],
        [Paragraph("<b>FECHA:</b>", cell_body), Paragraph(datos_encabezado['fecha'].strftime("%d/%m/%Y"), cell_body), Paragraph("<b>UNIDAD:</b>", cell_body), Paragraph(str(datos_encabezado['unidad']), cell_body)],
        [Paragraph("<b>TAG:</b>", cell_body), Paragraph(str(datos_encabezado['tag']), cell_body), Paragraph("<b>DESCRIPCIÓN:</b>", cell_body), Paragraph(str(datos_encabezado['descripcion']), cell_body)],
        [Paragraph("<b>ACA:</b>", cell_body), Paragraph(str(datos_encabezado['aca']), cell_body), Paragraph("<b>MOTIVO:</b>", cell_body), Paragraph(str(datos_encabezado['motivo']), cell_body)],
    ]
    t_enc = Table(data_enc, colWidths=[90, 186, 90, 186])
    t_enc.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F3F4F6')),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F3F4F6')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_enc)

    # Tabla Alcance
    data_alcance = [[Paragraph("<b>ALCANCE:</b>", cell_body), Paragraph(str(datos_encabezado['alcance']), cell_body)]]
    t_alcance = Table(data_alcance, colWidths=[90, 462])
    t_alcance.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F3F4F6')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_alcance)
    story.append(Spacer(1, 10))

    # Secciones dinámicas
    for titulo_sec, subsecciones in secciones.items():
        story.append(Paragraph(titulo_sec, sec_heading_style))
        for sub_tit, contenido in subsecciones.items():
            story.append(Paragraph(sub_tit, subsec_heading_style))
            story.append(Paragraph(contenido if contenido else "N/A", text_style))

    # Registros fotográficos
    story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS", sec_heading_style))
    
    if imagenes_procesadas:
        table_img_data = []
        for i in range(0, len(imagenes_procesadas), 2):
            row_cells = []
            
            # Foto 1
            img_bytes1, label1 = imagenes_procesadas[i]
            img_obj1 = RLImage(BytesIO(img_bytes1), width=6.8*cm, height=5.1*cm)
            cell1 = [img_obj1, Paragraph(f"<font size=7><b>{label1}</b></font>", cell_body)]
            row_cells.append(cell1)
            
            # Foto 2
            if i + 1 < len(imagenes_procesadas):
                img_bytes2, label2 = imagenes_procesadas[i+1]
                img_obj2 = RLImage(BytesIO(img_bytes2), width=6.8*cm, height=5.1*cm)
                cell2 = [img_obj2, Paragraph(f"<font size=7><b>{label2}</b></font>", cell_body)]
                row_cells.append(cell2)
            else:
                row_cells.append("")
                
            table_img_data.append(row_cells)

        t_imgs = Table(table_img_data, colWidths=[276, 276])
        t_imgs.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(t_imgs)

    doc.build(story, onFirstPage=dibujar_plantilla, onLaterPages=dibujar_plantilla)
    buffer.seek(0)
    return buffer

# =========================================================
# INTERFAZ DE USUARIO EN STREAMLIT
# =========================================================
st.title("📋 Generador de Informe de Inspección Visual")
st.markdown("##### Complete los campos de la plantilla e incluya imágenes ordenadas numéricamente.")

# 1. DATOS DE ENCABEZADO
st.markdown("#### 1. Encabezado e Identificación")
col1, col2 = st.columns(2)

with col1:
    num_informe = st.text_input("N.º DE INFORME", placeholder="Ej: IV-2026-001")
    fecha = st.date_input("FECHA", value=date.today())
    tag = st.text_input("TAG", placeholder="Ej: C-1302")
    aca = st.text_input("ACA", placeholder="Ej: ACA-2026")

with col2:
    ot = st.text_input("OT", placeholder="Ej: 45001234")
    unidad = st.text_input("UNIDAD", placeholder="Ej: UNIDAD-100")
    descripcion = st.text_input("DESCRIPCIÓN", placeholder="Ej: Columna de Fraccionamiento")
    motivo = st.text_input("MOTIVO", placeholder="Ej: Inspección Programada")

alcance = st.text_area("ALCANCE", height=70, placeholder="Describa el alcance de la inspección...")

# 2. CONTENIDO TÉCNICO
st.markdown("#### 2. Desarrollo del Informe")
tab1, tab2, tab3, tab4 = st.tabs(["1. Antecedentes", "2. Conclusiones", "3. Resultados", "4. Recomendaciones"])

with tab1:
    ant_gen = st.text_area("1.1 Antecedentes generales")
    doc_ref = st.text_area("1.2 Documentación de referencia")
    cond_prev = st.text_area("1.3 Condiciones previas / información disponible")
    otros_ant = st.text_area("1.4 Otros antecedentes")

with tab2:
    conc_insp = st.text_area("2.1 Conclusión de la inspección")
    cond_obs = st.text_area("2.2 Condición observada")
    cumplimiento = st.text_area("2.3 Cumplimiento / condición respecto de los criterios aplicables")
    conc_adic = st.text_area("2.4 Conclusiones adicionales")

with tab3:
    desc_res = st.text_area("3.1 Descripción de resultados")

with tab4:
    acc_rec = st.text_area("4.1 Acciones recomendadas")
    seguimiento = st.text_area("4.2 Seguimiento / inspecciones complementarias")
    obs_adic = st.text_area("4.3 Observaciones adicionales")

# 3. CARGA DE IMÁGENES
st.markdown("#### 3. Registro Fotográfico")
uploaded_files = st.file_uploader(
    "Seleccione imágenes desde su equipo (Se ordenarán automáticamente según el número del nombre del archivo, ej: 1_vista.jpg, 2_detalle.jpg):", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

imagenes_procesadas = []

if uploaded_files:
    archivos_ordenados = sorted(uploaded_files, key=lambda f: obtener_numero_archivo(f.name))
    st.info(f"📸 Se detectaron {len(archivos_ordenados)} imágenes. Ordenadas numéricamente por nombre de archivo.")
    
    cols = st.columns(3)
    for index, file in enumerate(archivos_ordenados):
        num_extraido = obtener_numero_archivo(file.name)
        with cols[index % 3]:
            st.image(file, caption=f"Archivo: {file.name}", use_container_width=True)
            caption_default = f"{num_extraido}: vista general de equipo" if num_extraido != 9999 else f"{index+1}: detalle de inspección"
            pie_foto = st.text_input(f"Pie de foto {index+1}:", value=caption_default, key=f"img_plantilla_{file.name}_{index}")
            
            file.seek(0)
            imagenes_procesadas.append((file.read(), pie_foto))

st.markdown("---")
st.markdown("#### 4. Descargar Informe Renderizado")
col_btn_w, col_btn_p = st.columns(2)

datos_encabezado = {
    'num_informe': num_informe,
    'ot': ot,
    'fecha': fecha,
    'unidad': unidad,
    'tag': tag,
    'descripcion': descripcion,
    'aca': aca,
    'motivo': motivo,
    'alcance': alcance
}

secciones = {
    "1. ANTECEDENTES": {
        "1.1 Antecedentes generales": ant_gen,
        "1.2 Documentación de referencia": doc_ref,
        "1.3 Condiciones previas / información disponible": cond_prev,
        "1.4 Otros antecedentes": otros_ant
    },
    "2. CONCLUSIONES": {
        "2.1 Conclusión de la inspección": conc_insp,
        "2.2 Condición observada": cond_obs,
        "2.3 Cumplimiento / condición respecto de los criterios aplicables": cumplimiento,
        "2.4 Conclusiones adicionales": conc_adic
    },
    "3. RESULTADOS DE LA INSPECCIÓN": {
        "3.1 Descripción de resultados": desc_res
    },
    "4. RECOMENDACIONES": {
        "4.1 Acciones recomendadas": acc_rec,
        "4.2 Seguimiento / inspecciones complementarias": seguimiento,
        "4.3 Observaciones adicionales": obs_adic
    }
}

with col_btn_w:
    word_buffer = generar_word_plantilla_inspeccion(datos_encabezado, secciones, imagenes_procesadas)
    st.download_button(
        label="📝 Descargar Informe en Word (.docx)",
        data=word_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'NUEVO'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

with col_btn_p:
    pdf_buffer = generar_pdf_plantilla_inspeccion(datos_encabezado, secciones, imagenes_procesadas)
    st.download_button(
        label="📄 Descargar Informe en PDF (.pdf)",
        data=pdf_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'NUEVO'}.pdf",
        mime="application/pdf",
        use_container_width=True
    )