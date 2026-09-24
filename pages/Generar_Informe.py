import io
import re
import json
import base64
import requests
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import date, datetime
from io import BytesIO

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA Y CONSTANTES
# =========================================================
st.set_page_config(
    page_title="Generar Informe Visual",
    page_icon="📄",
    layout="wide"
)

URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

# ID de la carpeta de Google Drive donde se almacenarán las imágenes
# Puedes definirlo en st.secrets["DRIVE_FOLDER_ID"] o ponerlo directamente aquí
DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "1a2b3c4d5e6f7g8h9i_REEMPLAZAR_POR_TU_ID")

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
# CONEXIÓN A APIS (GOOGLE SHEETS Y GOOGLE DRIVE)
# =========================================================
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

@st.cache_resource
def obtener_credenciales():
    """Obtiene las credenciales globales autorizadas."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(creds_dict, scopes=scopes)

@st.cache_resource
def conectar_google_sheets():
    """Mantiene en caché el cliente de Google Sheets."""
    credentials = obtener_credenciales()
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID)

@st.cache_resource
def conectar_google_drive():
    """Mantiene en caché la conexión con el servicio de Google Drive."""
    credentials = obtener_credenciales()
    return build('drive', 'v3', credentials=credentials)

# =========================================================
# FUNCIONES DE ALMACENAMIENTO EN GOOGLE DRIVE
# =========================================================
def subir_imagen_a_drive(nombre_archivo, img_bytes):
    """Subes bytes de una imagen a Google Drive indicando soporte de unidades compartidas."""
    try:
        drive_service = conectar_google_drive()
        file_metadata = {
            'name': nombre_archivo,
            'parents': [DRIVE_FOLDER_ID]
        }
        media = MediaIoBaseUpload(BytesIO(img_bytes), mimetype='image/jpeg', resumable=True)
        
        # supportsAllDrives=True permite subir archivos a carpetas compartidas sin error de cuota
        archivo_creado = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        
        file_id = archivo_creado.get('id')
        
        # Opcional: Hacer accesible el archivo por enlace de lectura
        drive_service.permissions().create(
            fileId=file_id,
            body={'role': 'reader', 'type': 'anyone'},
            supportsAllDrives=True
        ).execute()

        # Enlace directo para descarga/visualización
        return f"https://lh3.googleusercontent.com/d/{file_id}"
    except Exception as e:
        st.error(f"Error al subir imagen '{nombre_archivo}' a Google Drive: {e}")
        return None

# =========================================================
# LECTURA / ESCRITURA EN HISTORIAL
# =========================================================
@st.cache_data(ttl=300)
def cargar_base_equipos():
    try:
        client = conectar_google_sheets()
        ws = client.worksheet("BASE EQUIPOS")
        filas = ws.get_all_values()
        if len(filas) <= 1:
            return pd.DataFrame(columns=["UNIDAD", "TAG", "DESCRIPCION", "ACA"])
        
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
    except Exception:
        return pd.DataFrame(columns=["UNIDAD", "TAG", "DESCRIPCION", "ACA"])

def obtener_o_crear_hoja_historial():
    client = conectar_google_sheets()
    try:
        ws = client.worksheet("HISTORIAL_INFORMES")
    except Exception:
        ws = client.add_worksheet(title="HISTORIAL_INFORMES", rows="1000", cols="20")
        ws.append_row([
            "num_informe", "ot", "fecha", "unidad", "tag", 
            "descripcion", "aca", "motivo", "alcance", 
            "secciones_json", "inspector", "fotos_json"
        ])
    return ws

def guardar_resguardo_informe(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma):
    """Sube fotos a Google Drive y guarda los enlaces + metadata en Google Sheets."""
    try:
        ws = obtener_o_crear_hoja_historial()
        num_inf = datos_encabezado['num_informe'].strip()
        
        if not num_inf:
            return False, "Debe ingresar un N.º DE INFORME para poder resguardar."

        # Process pictures & upload to Drive
        fotos_guardadas = []
        for idx, (img_bytes, pie) in enumerate(imagenes_procesadas, start=1):
            nombre_foto = f"{num_inf}_foto_{idx}.jpg"
            url_drive = subir_imagen_a_drive(nombre_foto, img_bytes)
            if url_drive:
                fotos_guardadas.append({"url": url_drive, "pie": pie})
            else:
                fotos_guardadas.append({"url": "", "pie": pie})

        filas = ws.get_all_values()
        secciones_serializables = {str(k): v for k, v in secciones_dinamicas.items()}
        secciones_json = json.dumps(secciones_serializables, ensure_ascii=False)
        fotos_json = json.dumps(fotos_guardadas, ensure_ascii=False)
        
        fila_nueva = [
            num_inf,
            str(datos_encabezado['ot']),
            datos_encabezado['fecha'].strftime("%Y-%m-%d"),
            str(datos_encabezado['unidad']),
            str(datos_encabezado['tag']),
            str(datos_encabezado['descripcion']),
            str(datos_encabezado['aca']),
            str(datos_encabezado['motivo']),
            str(datos_encabezado['alcance']),
            secciones_json,
            str(inspector_firma),
            fotos_json
        ]

        fila_idx = None
        for idx, f in enumerate(filas[1:], start=2):
            if len(f) > 0 and f[0].strip().upper() == num_inf.upper():
                fila_idx = idx
                break

        if fila_idx:
            ws.update(f"A{fila_idx}:L{fila_idx}", [fila_nueva])
            mensaje = f"✅ Informe '{num_inf}' actualizado correctamente en Google Sheets con fotos almacenadas en Google Drive."
        else:
            ws.append_row(fila_nueva)
            mensaje = f"✅ Informe '{num_inf}' resguardado exitosamente en Google Sheets con fotos en Google Drive."

        st.cache_data.clear()
        return True, mensaje
    except Exception as e:
        return False, f"Error al guardar el resguardo: {str(e)}"

def obtener_lista_informes_guardados():
    try:
        ws = obtener_o_crear_hoja_historial()
        filas = ws.get_all_values()
        if len(filas) <= 1:
            return []
        return [f[0].strip() for f in filas[1:] if len(f) > 0 and f[0].strip()]
    except Exception:
        return []

def cargar_datos_informe(num_informe_sel):
    try:
        ws = obtener_o_crear_hoja_historial()
        filas = ws.get_all_values()
        for f in filas[1:]:
            if len(f) > 0 and f[0].strip().upper() == num_informe_sel.upper():
                secciones_json = json.loads(f[9]) if len(f) > 9 and f[9] else {}
                secciones_dict = {int(k): v for k, v in secciones_json.items()}
                
                fotos_json_str = f[11] if len(f) > 11 else ""
                imgs_recuperadas = []
                if fotos_json_str:
                    try:
                        fotos_lista = json.loads(fotos_json_str)
                        for item in fotos_lista:
                            url = item.get("url")
                            pie = item.get("pie", "")
                            if url:
                                img_bytes = descargar_bytes_desde_url(url)
                                if img_bytes:
                                    imgs_recuperadas.append((img_bytes, pie))
                    except Exception:
                        pass

                return {
                    "num_informe": f[0],
                    "ot": f[1] if len(f) > 1 else "",
                    "fecha": datetime.strptime(f[2], "%Y-%m-%d").date() if len(f) > 2 and f[2] else date.today(),
                    "unidad": f[3] if len(f) > 3 else "",
                    "tag": f[4] if len(f) > 4 else "",
                    "descripcion": f[5] if len(f) > 5 else "",
                    "aca": f[6] if len(f) > 6 else "",
                    "motivo": f[7] if len(f) > 7 else "",
                    "alcance": f[8] if len(f) > 8 else "",
                    "secciones_dinamicas": secciones_dict,
                    "inspector": f[10] if len(f) > 10 else "",
                    "imagenes_procesadas": imgs_recuperadas
                }
    except Exception as e:
        st.error(f"Error al cargar el informe: {e}")
    return None

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
# GENERACIÓN DE PDF Y WORD
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
        topMargin=2.5 * cm,
        bottomMargin=2.0 * cm
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

    for sec_num, sec_info in secciones_dinamicas.items():
        subpuntos = sec_info['subpuntos']
        if subpuntos:
            story.append(Paragraph(f"{sec_num}. {sec_info['titulo']}", sec_heading_style))
            for idx, sub in enumerate(subpuntos, start=1):
                num_sub = f"{sec_num}.{idx}"
                titulo_sub = f"{num_sub} {sub['titulo']}" if sub['titulo'] else num_sub
                story.append(Paragraph(titulo_sub, subsec_heading_style))
                story.append(Paragraph(sub['contenido'] if sub['contenido'] else "-", text_style))

    story.append(PageBreak())
    story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS", sec_heading_style))
    story.append(Spacer(1, 2))

    if imagenes_procesadas:
        for i in range(0, len(imagenes_procesadas), 2):
            if i > 0 and i % 6 == 0:
                story.append(PageBreak())
                story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS (Continuación)", sec_heading_style))
                story.append(Spacer(1, 2))

            row_cells = []
            img_bytes1, label1 = imagenes_procesadas[i]
            img_obj1 = RLImage(BytesIO(img_bytes1), width=9.4*cm, height=6.2*cm)
            cell1 = [img_obj1, Paragraph(f"<font size=7><b>{label1}</b></font>", cell_body)]
            row_cells.append(cell1)
            
            if i + 1 < len(imagenes_procesadas):
                img_bytes2, label2 = imagenes_procesadas[i+1]
                img_obj2 = RLImage(BytesIO(img_bytes2), width=9.4*cm, height=6.2*cm)
                cell2 = [img_obj2, Paragraph(f"<font size=7><b>{label2}</b></font>", cell_body)]
                row_cells.append(cell2)
            else:
                row_cells.append("")

            t_pair = Table([row_cells], colWidths=[276, 276])
            t_pair.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 1),
                ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('TOPPADDING', (0,0), (-1,-1), 1),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ]))
            story.append(t_pair)

    story.append(Spacer(1, 6))
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

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("INFORME DE INSPECCIÓN VISUAL")
    run_title.font.bold = True
    run_title.font.size = Pt(15)
    run_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("CONTROL DE INSPECCIÓN • CALIDAD • TRAZABILIDAD")
    run_sub.font.size = Pt(9)
    run_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    table = doc.add_table(rows=5, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    fields = [
        ("N.º DE INFORME:", datos_encabezado['num_informe'], "OT:", datos_encabezado['ot']),
        ("FECHA:", datos_encabezado['fecha'].strftime("%d/%m/%Y"), "UNIDAD:", datos_encabezado['unidad']),
        ("TAG:", datos_encabezado['tag'], "DESCRIPCIÓN:", datos_encabezado['descripcion']),
        ("ACA:", datos_encabezado['aca'], "MOTIVO:", datos_encabezado['motivo']),
    ]
    
    for row_idx, (k1, v1, k2, v2) in enumerate(fields):
        row = table.rows[row_idx]
        set_cell_background(row.cells[0], "F3F4F6")
        p = row.cells[0].paragraphs[0]
        r = p.add_run(k1)
        r.font.bold = True
        r.font.size = Pt(8.5)
        
        p = row.cells[1].paragraphs[0]
        r = p.add_run(str(v1))
        r.font.size = Pt(8.5)

        set_cell_background(row.cells[2], "F3F4F6")
        p = row.cells[2].paragraphs[0]
        r = p.add_run(k2)
        r.font.bold = True
        r.font.size = Pt(8.5)

        p = row.cells[3].paragraphs[0]
        r = p.add_run(str(v2))
        r.font.size = Pt(8.5)

    row_alcance = table.rows[4]
    set_cell_background(row_alcance.cells[0], "F3F4F6")
    p0 = row_alcance.cells[0].paragraphs[0]
    r0 = p0.add_run("ALCANCE:")
    r0.font.bold = True
    r0.font.size = Pt(8.5)

    cell_span = row_alcance.cells[1]
    cell_span.merge(row_alcance.cells[2])
    cell_span.merge(row_alcance.cells[3])
    p_alc = cell_span.paragraphs[0]
    r_alc = p_alc.add_run(str(datos_encabezado['alcance']))
    r_alc.font.size = Pt(8.5)

    doc.add_paragraph()

    for sec_num, sec_info in secciones_dinamicas.items():
        subpuntos = sec_info['subpuntos']
        if subpuntos:
            p_sec = doc.add_paragraph()
            r_sec = p_sec.add_run(f"{sec_num}. {sec_info['titulo']}")
            r_sec.font.bold = True
            r_sec.font.size = Pt(11)
            r_sec.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

            for idx, sub in enumerate(subpuntos, start=1):
                num_sub = f"{sec_num}.{idx}"
                titulo_sub = f"{num_sub} {sub['titulo']}" if sub['titulo'] else num_sub
                
                p_subsec = doc.add_paragraph()
                r_subsec = p_subsec.add_run(titulo_sub)
                r_subsec.font.bold = True
                r_subsec.font.size = Pt(9.5)
                r_subsec.font.color.rgb = RGBColor(0x61, 0x9B, 0x40)

                p_cont = doc.add_paragraph()
                r_cont = p_cont.add_run(sub['contenido'] if sub['contenido'] else "-")
                r_cont.font.size = Pt(8.5)
                r_cont.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    doc.add_page_break()
    p_sec5 = doc.add_paragraph()
    r_sec5 = p_sec5.add_run("5. REGISTROS FOTOGRÁFICOS")
    r_sec5.font.bold = True
    r_sec5.font.size = Pt(11)
    r_sec5.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    if imagenes_procesadas:
        img_table = doc.add_table(rows=0, cols=2)
        img_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(imagenes_procesadas), 2):
            row_cells = img_table.add_row().cells
            
            img_data1, label1 = imagenes_procesadas[i]
            p1 = row_cells[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(BytesIO(img_data1), width=Inches(3.5))
            
            p1_sub = row_cells[0].add_paragraph()
            p1_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1_sub = p1_sub.add_run(str(label1))
            r1_sub.font.bold = True
            r1_sub.font.size = Pt(8)

            if i + 1 < len(imagenes_procesadas):
                img_data2, label2 = imagenes_procesadas[i+1]
                p2 = row_cells[1].paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(BytesIO(img_data2), width=Inches(3.5))
                
                p2_sub = row_cells[1].add_paragraph()
                p2_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r2_sub = p2_sub.add_run(str(label2))
                r2_sub.font.bold = True
                r2_sub.font.size = Pt(8)

    if inspector_firma:
        p_firma = doc.add_paragraph()
        p_firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_run = p_firma.add_run(f"\nGenerado por: {inspector_firma}")
        p_run.font.bold = True
        p_run.font.size = Pt(9.5)
        p_run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# INTERFAZ DE USUARIO STREAMLIT
# =========================================================
st.title("📋 Generador de Informe de Inspección Visual")

df_equipos = cargar_base_equipos()

if "imagenes_cargadas_resguardo" not in st.session_state:
    st.session_state["imagenes_cargadas_resguardo"] = []

# SECCIÓN DE CARGAR INFORMES PREVIAMENTE RESGUARDADOS
st.markdown("### 📂 Cargar Informe Resguardado para Modificar")
lista_informes_guardados = ["-- Seleccionar informe resguardado --"] + obtener_lista_informes_guardados()
informe_sel = st.selectbox("Buscar por N.° de Informe Guardado:", lista_informes_guardados)

if st.button("📂 Cargar Datos e Imágenes del Informe Seleccionado", use_container_width=True):
    if informe_sel and informe_sel != "-- Seleccionar informe resguardado --":
        with st.spinner("Descargando imágenes guardadas desde Google Drive..."):
            datos_cargados = cargar_datos_informe(informe_sel)
            if datos_cargados:
                st.session_state['num_informe'] = datos_cargados['num_informe']
                st.session_state['ot'] = datos_cargados['ot']
                st.session_state['fecha'] = datos_cargados['fecha']
                st.session_state['unidad'] = datos_cargados['unidad']
                st.session_state['tag'] = datos_cargados['tag']
                st.session_state['descripcion'] = datos_cargados['descripcion']
                st.session_state['aca'] = datos_cargados['aca']
                st.session_state['motivo'] = datos_cargados['motivo']
                st.session_state['alcance'] = datos_cargados['alcance']
                st.session_state['inspector_firma'] = datos_cargados['inspector']

                # Cargar subpuntos dinámicos
                sec_loaded = datos_cargados['secciones_dinamicas']
                for s_num in [1, 2, 3, 4]:
                    sub_list = sec_loaded.get(s_num, {}).get('subpuntos', [])
                    st.session_state[f"cant_subpuntos_sec_{s_num}"] = len(sub_list)
                    for idx, sub in enumerate(sub_list, start=1):
                        st.session_state[f"tit_{s_num}_{idx}"] = sub.get("titulo", "")
                        st.session_state[f"cont_{s_num}_{idx}"] = sub.get("contenido", "")

                # Cargar imágenes resguardadas desde Google Drive
                st.session_state["imagenes_cargadas_resguardo"] = datos_cargados["imagenes_procesadas"]

                st.success(f"¡Informe '{informe_sel}' cargado correctamente con {len(st.session_state['imagenes_cargadas_resguardo'])} imágenes desde Google Drive!")
                st.rerun()

st.markdown("---")

# 1. ENCABEZADO E IDENTIFICACIÓN
st.markdown("#### 1. Encabezado e Identificación")

col_search1, col_search2 = st.columns([2, 1])

with col_search1:
    lista_tags = ["-- Seleccionar de BASE EQUIPOS --"] + df_equipos["TAG"].tolist() if not df_equipos.empty else ["-- Sin datos --"]
    tag_seleccionado = st.selectbox("🔍 Buscar TAG en BASE EQUIPOS:", lista_tags)

if tag_seleccionado and tag_seleccionado != "-- Seleccionar de BASE EQUIPOS --" and not df_equipos.empty:
    equipo_info = df_equipos[df_equipos["TAG"] == tag_seleccionado].iloc[0]
    st.session_state['unidad'] = equipo_info["UNIDAD"]
    st.session_state['descripcion'] = equipo_info["DESCRIPCIÓN"] if "DESCRIPCIÓN" in equipo_info else equipo_info["DESCRIPCION"]
    st.session_state['aca'] = equipo_info["ACA"]
    st.session_state['tag'] = tag_seleccionado

col1, col2 = st.columns(2)

with col1:
    num_informe = st.text_input("N.º DE INFORME", key='num_informe', placeholder="Ej: IV-2026-001")
    fecha = st.date_input("FECHA", key='fecha', value=date.today())
    tag = st.text_input("TAG", key='tag', placeholder="Ej: C-1302")
    aca = st.text_input("ACA", key='aca', placeholder="Ej: ACA-2026")

with col2:
    ot = st.text_input("OT", key='ot', placeholder="Ej: 45001234")
    
    col_u1, col_u2 = st.columns([2, 1])
    val_u = st.session_state.get('unidad', '')
    idx_u = LISTA_PLANTAS.index(val_u) + 1 if val_u in LISTA_PLANTAS else 0
    with col_u1:
        unidad_select = st.selectbox("UNIDAD / PLANTA (Seleccionar):", [""] + LISTA_PLANTAS, index=idx_u)
    with col_u2:
        unidad_manual = st.text_input("O escribir Unidad:", value=val_u if idx_u == 0 else "", placeholder="Manual")
    
    unidad_final = unidad_manual.strip() if unidad_manual.strip() else unidad_select

    descripcion = st.text_input("DESCRIPCIÓN", key='descripcion', placeholder="Ej: Columna de Fraccionamiento")
    motivo = st.text_input("MOTIVO", key='motivo', placeholder="Ej: Inspección Programada")

alcance = st.text_area("ALCANCE", key='alcance', height=70, placeholder="Describa el alcance de la inspección...")

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
    st.session_state["imagenes_cargadas_resguardo"] = []
    archivos_ordenados = sorted(uploaded_files, key=lambda f: obtener_numero_archivo(f.name))
    st.info(f"📸 Se detectaron {len(archivos_ordenados)} imágenes subidas. Ordenadas numéricamente.")
    
    cols = st.columns(3)
    for index, file in enumerate(archivos_ordenados):
        num_extraido = obtener_numero_archivo(file.name)
        with cols[index % 3]:
            st.image(file, caption=f"Archivo: {file.name}", use_container_width=True)
            caption_default = f"{num_extraido}: vista general de equipo" if num_extraido != 9999 else f"{index+1}: detalle de inspección"
            pie_foto = st.text_input(f"Pie de foto {index+1}:", value=caption_default, key=f"img_plantilla_{file.name}_{index}")
            
            file.seek(0)
            imagenes_procesadas.append((file.read(), pie_foto))

elif st.session_state.get("imagenes_cargadas_resguardo"):
    imgs_res = st.session_state["imagenes_cargadas_resguardo"]
    st.info(f"📸 Se cargaron {len(imgs_res)} imágenes desde Google Drive.")
    
    cols = st.columns(3)
    for index, (img_bytes, pie_orig) in enumerate(imgs_res):
        with cols[index % 3]:
            st.image(img_bytes, caption=f"Foto {index+1}", use_container_width=True)
            pie_foto = st.text_input(f"Pie de foto {index+1}:", value=pie_orig, key=f"img_resguardada_{index}")
            imagenes_procesadas.append((img_bytes, pie_foto))

# FIRMA / GENERADO POR
st.markdown("---")
st.markdown("#### Responsable del Informe")
val_insp = st.session_state.get('inspector_firma', '')
idx_insp = LISTA_INSPECTORES.index(val_insp) + 1 if val_insp in LISTA_INSPECTORES else 0
inspector_firma = st.selectbox("👷‍♂️ Generado por:", [""] + LISTA_INSPECTORES, index=idx_insp, key='inspector_firma')

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

# BOTÓN DE RESGUARDAR / GUARDAR EN GOOGLE SHEETS Y DRIVE
st.markdown("---")
st.markdown("#### 💾 Resguardo del Informe")

if st.button("💾 RESGUARDAR INFORME Y FOTOS EN GOOGLE DRIVE", type="primary", use_container_width=True):
    with st.spinner("Subiendo fotos a Google Drive y registrando informe..."):
        exito, msg = guardar_resguardo_informe(datos_encabezado, secciones_dinamicas, imagenes_procesadas, inspector_firma)
        if exito:
            st.success(msg)
        else:
            st.error(msg)

# GENERAR Y DESCARGAR ARCHIVOS
st.markdown("---")
st.markdown("#### Exportar Informe Final")

col_btn_p, col_btn_w = st.columns(2)

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
