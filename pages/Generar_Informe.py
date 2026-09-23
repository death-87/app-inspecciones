import io
import re
import requests
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
from io import BytesIO

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
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
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Arlem Sarmiento",
    "Harold Castillo",
    "Miguel Chirinos"
]

LISTA_PLANTAS = [
    "A0AEX", "A0ALQ", "A0BUT", "A0CCK", "A0CCR", "A0CKR", "A0HDG", "A0HDT", "A0HCK", 
    "A0ISO", "A0LAB", "A0MHC", "A0NHT", "A0SAR", "A0SHP", "A0SWS", "AACID", "AAMAR", 
    "AAMIN", "AAMPL", "AANTO", "AAREF", "AASER", "AALQU", "ADESO", "ADEV1", "ADEV2", 
    "ADIPE", "AE501", "ALNHT", "ALPG1", "ALPG2", "ALPG3", "AMACO", "AMDEA", "AMRX1", 
    "AMRX2", "AMRX3", "AMRX4", "AMVPR", "AOLEO", "APBMP", "APBTQ", "APCAR", "APFEN", 
    "APRCO", "ARPLU", "AREFO", "AREMO", "ARILE", "ASAIC", "ASOLV", "ASPLI", "ASRCO", 
    "ASUEL", "ASVAQ", "ASVAP", "ASWS2", "ASYBR", "ASEFL", "ASEFQ", "ATOP1", "ATOP2", 
    "ATRAG", "AURA1", "AURA2", "AURA3", "AVAC1", "AVAC2", "ACOKE"
]

# =========================================================
# CONEXIÓN A GOOGLE SHEETS (BASE EQUIPOS)
# =========================================================
def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID)

@st.cache_data(ttl=600)
def cargar_base_equipos():
    """Carga los datos de la hoja BASE EQUIPOS."""
    try:
        client = conectar_google_sheets()
        ws = client.worksheet("BASE EQUIPOS")
        filas = ws.get_all_values()
        if len(filas) <= 1:
            return pd.DataFrame(columns=["UNIDAD", "TAG", "DESCRIPCION", "ACA"])
        
        # Columna C = índice 2 (UNIDAD)
        # Columna D = índice 3 (TAG)
        # Columna E = índice 4 (DESCRIPCION)
        # Columna H = índice 7 (ACA)
        datos = []
        for f in filas[1:]:
            unidad = f[2].strip() if len(f) > 2 else ""
            tag = f[3].strip() if len(f) > 3 else ""
            descripcion = f[4].strip() if len(f) > 4 else ""
            aca_val = f[7].strip() if len(f) > 7 else ""
            
            if tag:
                datos.append({
                    "UNIDAD": unidad, 
                    "TAG": tag, 
                    "DESCRIPCION": descripcion,
                    "ACA": aca_val
                })
        
        return pd.DataFrame(datos)
    except Exception as e:
        st.warning(f"⚠️ No se pudo leer la hoja 'BASE EQUIPOS' ({e}). Se habilitará ingreso manual.")
        return pd.DataFrame(columns=["UNIDAD", "TAG", "DESCRIPCION", "ACA"])

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
    match = re.search(r'^(\d+)', nombre_archivo)
    return int(match.group(1)) if match else 9999

# =========================================================
# GENERACIÓN DE DOCUMENTOS (PDF Y WORD)
# =========================================================
def dibujar_plantilla_pdf(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#f8faf6"))
    canvas.rect(0, 0, letter[0], letter[1], fill=1, stroke=0)

    logo_bytes = obtener_bytes_imagen(URL_LOGO_GITHUB)
    if logo_bytes:
        try:
            img_stream = BytesIO(logo_bytes)
            img = ImageReader(img_stream)
            canvas.drawImage(img, 30, letter[1] - (1.8 * cm) - 35, width=120, height=45, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#6B7280"))
    canvas.drawCentredString(letter[0] / 2.0, 26, "SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.")
    canvas.drawCentredString(letter[0] / 2.0, 16, "CONTRATO N° AC 31104857")
    canvas.drawRightString(letter[0] - 30, 16, f"Pág. {canvas.getPageNumber()}")
    canvas.restoreState()

def generar_pdf_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=2.8 * cm,
        bottomMargin=2.2 * cm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=15, leading=18, textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=2)
    subtitle_style = ParagraphStyle('DocSubTitle', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=10)
    sec_heading_style = ParagraphStyle('SecHeader', parent=styles['Heading2'], fontSize=11, leading=13, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=4)
    subsec_heading_style = ParagraphStyle('SubSecHeader', parent=styles['Heading3'], fontSize=9.5, leading=11, textColor=colors.HexColor('#619b40'), fontName='Helvetica-Bold', spaceBefore=4, spaceAfter=2)
    text_style = ParagraphStyle('TextStyle', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#1F2937'), spaceAfter=6)
    cell_body = ParagraphStyle('CB', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))
    firma_style = ParagraphStyle('FirmaStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', alignment=2)

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

    # Puntos 1, 2, 3, 4
    for sec_num, sec_info in secciones_dinamicas.items():
        subpuntos = sec_info['subpuntos']
        if subpuntos:
            story.append(Paragraph(f"{sec_num}. {sec_info['titulo']}", sec_heading_style))
            for idx, sub in enumerate(subpuntos, start=1):
                num_sub = f"{sec_num}.{idx}"
                titulo_sub = f"{num_sub} {sub['titulo']}" if sub['titulo'] else num_sub
                story.append(Paragraph(titulo_sub, subsec_heading_style))
                story.append(Paragraph(sub['contenido'] if sub['contenido'] else "-", text_style))

    # Salto a Punto 5 (Fotografías)
    story.append(PageBreak())
    story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS", sec_heading_style))
    story.append(Spacer(1, 4))

    if imagenes_procesadas:
        for i in range(0, len(imagenes_procesadas), 2):
            if i > 0 and i % 6 == 0:
                story.append(PageBreak())
                story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS (Continuación)", sec_heading_style))
                story.append(Spacer(1, 4))

            row_cells = []
            img_bytes1, label1 = imagenes_procesadas[i]
            img_obj1 = RLImage(BytesIO(img_bytes1), width=9.5*cm, height=6.8*cm)
            cell1 = [img_obj1, Spacer(1, 2), Paragraph(f"<font size=7.5><b>{label1}</b></font>", cell_body)]
            row_cells.append(cell1)
            
            if i + 1 < len(imagenes_procesadas):
                img_bytes2, label2 = imagenes_procesadas[i+1]
                img_obj2 = RLImage(BytesIO(img_bytes2), width=9.5*cm, height=6.8*cm)
                cell2 = [img_obj2, Spacer(1, 2), Paragraph(f"<font size=7.5><b>{label2}</b></font>", cell_body)]
                row_cells.append(cell2)
            else:
                row_cells.append("")

            t_pair = Table([row_cells], colWidths=[276, 276])
            t_pair.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(t_pair)

    story.append(Spacer(1, 10))
    if inspector_firma:
        story.append(Paragraph(f"<b>Generado por:</b> {inspector_firma}", firma_style))

    doc.build(story, onFirstPage=dibujar_plantilla_pdf, onLaterPages=dibujar_plantilla_pdf)
    buffer.seek(0)
    return buffer

def generar_word_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma):
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

    for sec_num, sec_info in secciones_dinamicas.items():
        subpuntos = sec_info['subpuntos']
        if subpuntos:
            doc.add_heading(f"{sec_num}. {sec_info['titulo']}", level=1)
            for idx, sub in enumerate(subpuntos, start=1):
                num_sub = f"{sec_num}.{idx}"
                titulo_sub = f"{num_sub} {sub['titulo']}" if sub['titulo'] else num_sub
                doc.add_heading(titulo_sub, level=2)
                doc.add_paragraph(sub['contenido'] if sub['contenido'] else "-")

    doc.add_page_break()
    doc.add_heading("5. REGISTROS FOTOGRÁFICOS", level=1)
    
    if imagenes_procesadas:
        img_table = doc.add_table(rows=0, cols=2)
        img_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(imagenes_procesadas), 2):
            row_cells = img_table.add_row().cells
            
            img_data1, label1 = imagenes_procesadas[i]
            p1 = row_cells[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(BytesIO(img_data1), width=Inches(3.2))
            p1_sub = row_cells[0].add_paragraph(str(label1))
            p1_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if i + 1 < len(imagenes_procesadas):
                img_data2, label2 = imagenes_procesadas[i+1]
                p2 = row_cells[1].paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(BytesIO(img_data2), width=Inches(3.2))
                p2_sub = row_cells[1].add_paragraph(str(label2))
                p2_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if inspector_firma:
        p_firma = doc.add_paragraph()
        p_firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_run = p_firma.add_run(f"\nGenerado por: {inspector_firma}")
        p_run.font.bold = True
        p_run.font.size = Pt(10)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# INTERFAZ DE USUARIO STREAMLIT
# =========================================================
st.title("📋 Generador de Informe de Inspección Visual")

# Cargar Base de Equipos desde Google Sheets
df_equipos = cargar_base_equipos()

# 1. ENCABEZADO E IDENTIFICACIÓN
st.markdown("#### 1. Encabezado e Identificación")

# BÚSQUEDA Y SELECCIÓN DESDE BASE EQUIPOS
col_search1, col_search2 = st.columns([2, 1])

with col_search1:
    lista_tags = ["-- Seleccionar de BASE EQUIPOS --"] + df_equipos["TAG"].tolist() if not df_equipos.empty else ["-- Sin datos --"]
    tag_seleccionado = st.selectbox("🔍 Buscar TAG en BASE EQUIPOS:", lista_tags)

unidad_defecto = ""
descripcion_defecto = ""
aca_defecto = ""
tag_defecto = ""

if tag_seleccionado and tag_seleccionado != "-- Seleccionar de BASE EQUIPOS --" and not df_equipos.empty:
    equipo_info = df_equipos[df_equipos["TAG"] == tag_seleccionado].iloc[0]
    unidad_defecto = equipo_info["UNIDAD"]
    descripcion_defecto = equipo_info["DESCRIPCIÓN"] if "DESCRIPCIÓN" in equipo_info else equipo_info["DESCRIPCION"]
    aca_defecto = equipo_info["ACA"]
    tag_defecto = tag_seleccionado

col1, col2 = st.columns(2)

with col1:
    num_informe = st.text_input("N.º DE INFORME", placeholder="Ej: IV-2026-001")
    fecha = st.date_input("FECHA", value=date.today())
    tag = st.text_input("TAG", value=tag_defecto, placeholder="Ej: C-1302")
    aca = st.text_input("ACA", value=aca_defecto, placeholder="Ej: ACA-2026")

with col2:
    ot = st.text_input("OT", placeholder="Ej: 45001234")
    
    col_u1, col_u2 = st.columns([2, 1])
    with col_u1:
        idx_u = LISTA_PLANTAS.index(unidad_defecto) + 1 if unidad_defecto in LISTA_PLANTAS else 0
        unidad_select = st.selectbox("UNIDAD / PLANTA (Seleccionar):", [""] + LISTA_PLANTAS, index=idx_u)
    with col_u2:
        unidad_manual = st.text_input("O escribir Unidad:", value=unidad_defecto if idx_u == 0 else "", placeholder="Manual")
    
    unidad_final = unidad_manual.strip() if unidad_manual.strip() else unidad_select

    descripcion = st.text_input("DESCRIPCIÓN", value=descripcion_defecto, placeholder="Ej: Columna de Fraccionamiento")
    motivo = st.text_input("MOTIVO", placeholder="Ej: Inspección Programada")

alcance = st.text_area("ALCANCE", height=70, placeholder="Describa el alcance de la inspección...")

# 2, 3, 4. DESARROLLO DEL INFORME CON SUBÍNDICES OPCIONALES
st.markdown("---")
st.markdown("#### Desarrollo de Secciones Técnicas (Subíndices Opcionales)")

secciones_base = {
    1: "ANTECEDENTES",
    2: "CONCLUSIONES",
    3: "RESULTADOS DE LA INSPECCIÓN",
    4: "RECOMENDACIONES"
}

secciones_dinamicas = {}

for num_sec, tit_sec in secciones_base.items():
    with st.expander(f"📌 {num_sec}. {tit_sec}", expanded=(num_sec == 1)):
        key_count = f"cant_subpuntos_sec_{num_sec}"
        if key_count not in st.session_state:
            st.session_state[key_count] = 0

        col_b1, col_b2, col_b3 = st.columns([1, 1, 3])
        with col_b1:
            if st.button(f"➕ Agregar subpunto {num_sec}.x", key=f"btn_add_{num_sec}"):
                st.session_state[key_count] += 1
                st.rerun()
        with col_b2:
            if st.session_state[key_count] > 0:
                if st.button(f"➖ Quitar último", key=f"btn_rem_{num_sec}"):
                    st.session_state[key_count] -= 1
                    st.rerun()

        subpuntos_list = []
        cant = st.session_state[key_count]
        
        if cant == 0:
            st.caption("ℹ️ *Sin subpuntos. No se incluirá esta sección en el documento final a menos que agregues subpuntos.*")
        
        for i in range(1, cant + 1):
            sub_num_str = f"{num_sec}.{i}"
            c_t, c_c = st.columns([1.5, 3])
            with c_t:
                tit_sub = st.text_input(f"Título ({sub_num_str}):", placeholder="Ej: Antecedentes generales", key=f"tit_{num_sec}_{i}")
            with c_c:
                cont_sub = st.text_area(f"Contenido ({sub_num_str}):", placeholder="Detalle...", height=70, key=f"cont_{num_sec}_{i}")
            
            subpuntos_list.append({"titulo": tit_sub, "contenido": cont_sub})

        secciones_dinamicas[num_sec] = {
            "titulo": tit_sec,
            "subpuntos": subpuntos_list
        }

# 5. REGISTROS FOTOGRÁFICOS
st.markdown("---")
st.markdown("#### 5. Registros Fotográficos")

uploaded_files = st.file_uploader(
    "Subir imágenes desde su PC (Se ordenarán automáticamente según el número inicial del nombre del archivo, ej: 1_vista.jpg, 2_detalle.jpg):", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

imagenes_procesadas = []

if uploaded_files:
    archivos_ordenados = sorted(uploaded_files, key=lambda f: obtener_numero_archivo(f.name))
    st.info(f"📸 Se detectaron {len(archivos_ordenados)} imágenes. Ordenadas numéricamente.")
    
    cols = st.columns(3)
    for index, file in enumerate(archivos_ordenados):
        num_extraido = obtener_numero_archivo(file.name)
        with cols[index % 3]:
            st.image(file, caption=f"Archivo: {file.name}", use_container_width=True)
            caption_default = f"{num_extraido}: vista general de equipo" if num_extraido != 9999 else f"{index+1}: detalle de inspección"
            pie_foto = st.text_input(f"Pie de foto {index+1}:", value=caption_default, key=f"img_plantilla_{file.name}_{index}")
            
            file.seek(0)
            imagenes_procesadas.append((file.read(), pie_foto))

# FIRMA / GENERADO POR
st.markdown("---")
st.markdown("#### Responsable del Informe")
inspector_firma = st.selectbox("👷‍♂️ Generado por:", [""] + LISTA_INSPECTORES)

# GENERAR Y DESCARGAR ARCHIVOS
st.markdown("---")
st.markdown("#### Exportar Informe Final")

col_btn_p, col_btn_w = st.columns(2)

datos_encabezado = {
    'num_informe': num_informe,
    'ot': ot,
    'fecha': fecha,
    'unidad': unidad_final,
    'tag': tag,
    'descripcion': descripcion,
    'aca': aca,
    'motivo': motivo,
    'alcance': alcance
}

with col_btn_p:
    pdf_buffer = generar_pdf_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma)
    st.download_button(
        label="📄 Descargar Informe en PDF (.pdf)",
        data=pdf_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'INSPECCION'}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

with col_btn_w:
    word_buffer = generar_word_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma)
    st.download_button(
        label="📝 Descargar Informe en Word (.docx)",
        data=word_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'INSPECCION'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
