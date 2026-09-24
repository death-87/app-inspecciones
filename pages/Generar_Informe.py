import io
import re
import json
import requests
import streamlit as st
import pandas as pd
import gspread

from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from datetime import date, datetime
from io import BytesIO

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak
)
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


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

URL_LOGO_GITHUB = (
    "https://raw.githubusercontent.com/death-87/"
    "app-inspecciones/main/logo.png"
)

SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Arlem Sarmiento",
    "Harold Castillo",
    "Miguel Chirinos"
]

LISTA_PLANTAS = [
    "A0AEX", "A0ALQ", "A0BUT", "A0CCK", "A0CCR", "A0CKR",
    "A0HDG", "A0HDT", "A0HCK", "A0ISO", "A0LAB", "A0MHC",
    "A0NHT", "A0SAR", "A0SHP", "A0SWS", "AACID", "AAMAR",
    "AAMIN", "AAMPL", "AANTO", "AAREF", "AASER", "AALQU",
    "ADESO", "ADEV1", "ADEV2", "ADIPE", "AE501", "ALNHT",
    "ALPG1", "ALPG2", "ALPG3", "AMACO", "AMDEA", "AMRX1",
    "AMRX2", "AMRX3", "AMRX4", "AMVPR", "AOLEO", "APBMP",
    "APBTQ", "APCAR", "APFEN", "APRCO", "ARPLU", "AREFO",
    "AREMO", "ARILE", "ASAIC", "ASOLV", "ASPLI", "ASRCO",
    "ASUEL", "ASVAQ", "ASVAP", "ASWS2", "ASYBR", "ASEFL",
    "ASEFQ", "ATOP1", "ATOP2", "ATRAG", "AURA1", "AURA2",
    "AURA3", "AVAC1", "AVAC2", "ACOKE"
]


# =========================================================
# FUNCIONES AUXILIARES WORD
# =========================================================

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def agregar_numero_pagina_word(run):
    """Inserta el campo dinámico PAGE en un campo de texto en Word."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)


# =========================================================
# CONEXIÓN MEDIANTE SERVICE ACCOUNT
# =========================================================

@st.cache_resource
def obtener_credenciales_service_account():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return ServiceAccountCredentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )


@st.cache_resource
def conectar_google_sheets():
    credentials = obtener_credenciales_service_account()
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID)


def conectar_google_drive_service_account():
    try:
        credentials = obtener_credenciales_service_account()
        return build("drive", "v3", credentials=credentials)
    except Exception as e:
        st.error(f"Error conectando a Google Drive con Service Account: {e}")
        return None


# =========================================================
# LECTURA DE CARPETAS EN GOOGLE DRIVE (FOTOS Y ESQUEMAS)
# =========================================================

def extraer_id_carpeta(input_text):
    match = re.search(r'folders/([a-zA-Z0-9_-]+)', input_text)
    if match:
        return match.group(1)
    return input_text.strip()


def obtener_imagenes_desde_drive_folder(folder_input):
    folder_id = extraer_id_carpeta(folder_input)
    if not folder_id:
        return [], [], "No se proporcionó un ID o enlace válido de carpeta."

    try:
        drive_service = conectar_google_drive_service_account()
        if not drive_service:
            return [], [], "No se pudo autenticar el servicio de Google Drive."

        query = f"'{folder_id}' in parents and (mimeType contains 'image/' or mimeType = 'application/octet-stream') and trashed = false"
        
        results = drive_service.files().list(
            q=query, 
            fields="files(id, name, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageSize=100
        ).execute()
        
        files = results.get('files', [])

        if not files:
            return [], [], "No se encontraron imágenes en la carpeta de Google Drive."

        fotos = []
        esquemas = []

        files_fotos = [f for f in files if "esquema" not in f['name'].lower()]
        files_esquemas = [f for f in files if "esquema" in f['name'].lower()]

        files_fotos_ordenados = sorted(files_fotos, key=lambda f: obtener_numero_archivo(f['name']))
        files_esquemas_ordenados = sorted(files_esquemas, key=lambda f: obtener_numero_archivo(f['name']))

        for index, file in enumerate(files_fotos_ordenados):
            request = drive_service.files().get_media(fileId=file['id'])
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            
            file_name = file['name']
            num_extraido = obtener_numero_archivo(file_name)
            name_without_ext = re.sub(r'\.[a-zA-Z0-9]+$', '', file_name)
            match_texto = re.search(r'_(.+)$', name_without_ext)
            
            if match_texto:
                caption = match_texto.group(1).strip()
            else:
                caption = f"{num_extraido}: vista general de equipo" if num_extraido != 9999 else f"{index+1}: detalle de inspección"
                
            fotos.append((fh.read(), caption))

        for index, file in enumerate(files_esquemas_ordenados):
            request = drive_service.files().get_media(fileId=file['id'])
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            
            file_name = file['name']
            name_without_ext = re.sub(r'\.[a-zA-Z0-9]+$', '', file_name)
            match_texto = re.search(r'_(.+)$', name_without_ext)
            
            if match_texto:
                caption = match_texto.group(1).strip()
            else:
                caption = f"Esquema {index+1}: Ubicación de hallazgos y sectores afectados"
                
            esquemas.append((fh.read(), caption))

        msg = f"✅ Se cargaron exitosamente {len(fotos)} fotografías y {len(esquemas)} esquemas desde Google Drive."
        return fotos, esquemas, msg

    except Exception as e:
        return [], [], f"⚠️ Error al acceder a la carpeta de Google Drive: {str(e)}"


# =========================================================
# FUNCIONES AUXILIARES DE IMAGEN Y TEXTO
# =========================================================

@st.cache_data(ttl=3600)
def obtener_bytes_imagen(url):
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.content
    except Exception:
        pass
    return None


def obtener_numero_archivo(nombre_archivo):
    match = re.search(r'^(\d+)', nombre_archivo)
    return int(match.group(1)) if match else 9999


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
            "secciones_json", "inspector", "drive_link"
        ])
    return ws


def guardar_resguardo_informe(datos_encabezado, secciones_dinamicas, drive_link, inspector_firma):
    try:
        ws = obtener_o_crear_hoja_historial()
        num_inf = datos_encabezado["num_informe"].strip()

        if not num_inf:
            return False, "Debe ingresar un N.º DE INFORME para poder resguardar."

        filas = ws.get_all_values()
        secciones_serializables = {str(k): v for k, v in secciones_dinamicas.items()}
        secciones_json = json.dumps(secciones_serializables, ensure_ascii=False)

        fila_nueva = [
            num_inf,
            str(datos_encabezado["ot"]),
            datos_encabezado["fecha"].strftime("%Y-%m-%d"),
            str(datos_encabezado["unidad"]),
            str(datos_encabezado["tag"]),
            str(datos_encabezado["descripcion"]),
            str(datos_encabezado["aca"]),
            str(datos_encabezado["motivo"]),
            str(datos_encabezado["alcance"]),
            secciones_json,
            str(inspector_firma),
            str(drive_link)
        ]

        fila_idx = None
        for idx, f in enumerate(filas[1:], start=2):
            if len(f) > 0 and f[0].strip().upper() == num_inf.upper():
                fila_idx = idx
                break

        if fila_idx:
            ws.update(f"A{fila_idx}:L{fila_idx}", [fila_nueva])
            mensaje = f"✅ Informe '{num_inf}' actualizado correctamente en Google Sheets."
        else:
            ws.append_row(fila_nueva, value_input_option="USER_ENTERED")
            mensaje = f"✅ Informe '{num_inf}' resguardado exitosamente en Google Sheets."

        st.cache_data.clear()
        return True, mensaje

    except Exception as e:
        return False, f"Error al guardar el informe: {e}"


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
                drive_link = f[11] if len(f) > 11 else ""

                imgs_recuperadas = []
                esquemas_recuperados = []
                if drive_link:
                    imgs_recuperadas, esquemas_recuperados, _ = obtener_imagenes_desde_drive_folder(drive_link)

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
                    "drive_link": drive_link,
                    "imagenes_procesadas": imgs_recuperadas,
                    "esquemas_procesados": esquemas_recuperados
                }
    except Exception as e:
        st.error(f"Error al cargar el informe: {e}")
    return None


def eliminar_informe_guardado(num_informe_sel):
    try:
        ws = obtener_o_crear_hoja_historial()
        filas = ws.get_all_values()
        
        for idx, f in enumerate(filas[1:], start=2):
            if len(f) > 0 and f[0].strip().upper() == num_informe_sel.strip().upper():
                ws.delete_rows(idx)
                st.cache_data.clear()
                return True, f"🗑️ El informe '{num_informe_sel}' fue eliminado con éxito de Google Sheets."
                
        return False, "No se encontró el informe especificado para eliminar."
    except Exception as e:
        return False, f"Error al intentar eliminar el informe: {e}"


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


def generar_pdf_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, esquemas_procesados, inspector_firma):
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

    # 5. REGISTROS FOTOGRÁFICOS
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

    # 6. ESQUEMA DE EQUIPO
    if esquemas_procesados:
        story.append(PageBreak())
        story.append(Paragraph("6. ESQUEMA DE EQUIPO", sec_heading_style))
        story.append(Spacer(1, 6))

        for idx, (esq_bytes, label_esq) in enumerate(esquemas_procesados, start=1):
            if idx > 1:
                story.append(PageBreak())
                story.append(Paragraph(f"6. ESQUEMA DE EQUIPO (Continuación - Esquema {idx})", sec_heading_style))
                story.append(Spacer(1, 6))

            img_esq = RLImage(BytesIO(esq_bytes), width=18.5*cm, height=13.5*cm)
            t_esq = Table([[img_esq], [Paragraph(f"<b>{label_esq}</b>", cell_body)]], colWidths=[552])
            t_esq.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 2),
            ]))
            story.append(t_esq)
            story.append(Spacer(1, 10))

    story.append(Spacer(1, 6))
    if inspector_firma:
        story.append(Paragraph(f"<b>Generado por:</b> {inspector_firma}", firma_style))

    doc.build(story, onFirstPage=dibujar_plantilla_pdf, onLaterPages=dibujar_plantilla_pdf)
    buffer.seek(0)
    return buffer


def generar_word_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, esquemas_procesados, inspector_firma):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    # =========================================================
    # ENCABEZADO WORD (LOGO PARTE SUPERIOR IZQUIERDA)
    # =========================================================
    header = section.header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    logo_bytes = obtener_bytes_imagen(URL_LOGO_GITHUB)
    if logo_bytes:
        try:
            header_p.add_run().add_picture(BytesIO(logo_bytes), width=Inches(1.8))
        except Exception:
            pass

    # =========================================================
    # PIE DE PÁGINA WORD (TEXTO CENTRADO Y NUMERACIÓN DE PÁGINA)
    # =========================================================
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_ft_1 = footer_p.add_run("SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.\n")
    run_ft_1.font.size = Pt(7)
    run_ft_1.font.name = "Helvetica"
    run_ft_1.font.color.rgb = RGBColor(107, 114, 128)

    run_ft_2 = footer_p.add_run("CONTRATO N° AC 31104857\n")
    run_ft_2.font.size = Pt(7)
    run_ft_2.font.name = "Helvetica"
    run_ft_2.font.color.rgb = RGBColor(107, 114, 128)

    run_ft_3 = footer_p.add_run("Pág. ")
    run_ft_3.font.size = Pt(7)
    run_ft_3.font.name = "Helvetica"
    run_ft_3.font.color.rgb = RGBColor(107, 114, 128)
    agregar_numero_pagina_word(run_ft_3)

    # =========================================================
    # CUERPO DEL INFORME WORD
    # =========================================================
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

    # 5. REGISTROS FOTOGRÁFICOS
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

    # 6. ESQUEMAS EN WORD
    if esquemas_procesados:
        doc.add_page_break()
        p_sec6 = doc.add_paragraph()
        r_sec6 = p_sec6.add_run("6. ESQUEMA DE EQUIPO")
        r_sec6.font.bold = True
        r_sec6.font.size = Pt(11)
        r_sec6.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

        for idx, (esq_data, label_esq) in enumerate(esquemas_procesados, start=1):
            if idx > 1:
                doc.add_page_break()

            p_esq = doc.add_paragraph()
            p_esq.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_esq.add_run().add_picture(BytesIO(esq_data), width=Inches(6.8))

            p_esq_sub = doc.add_paragraph()
            p_esq_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_esq_sub = p_esq_sub.add_run(str(label_esq))
            r_esq_sub.font.bold = True
            r_esq_sub.font.size = Pt(9)

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
if "esquemas_cargados_resguardo" not in st.session_state:
    st.session_state["esquemas_cargados_resguardo"] = []

# CARGAR O ELIMINAR INFORMES PREVIAMENTE RESGUARDADOS
st.markdown("### 📂 Cargar o Eliminar Informe Resguardado")
lista_informes_guardados = ["-- Seleccionar informe resguardado --"] + obtener_lista_informes_guardados()
informe_sel = st.selectbox("Buscar por N.° de Informe Guardado:", lista_informes_guardados)

col_acc1, col_acc2 = st.columns([2, 1])

with col_acc1:
    btn_cargar = st.button("📂 Cargar Datos e Imágenes del Informe Seleccionado", use_container_width=True)

with col_acc2:
    confirmar_eliminar = st.checkbox("⚠️ Confirmar eliminación", key="chk_eliminar")
    btn_eliminar = st.button("🗑️ Eliminar Informe", type="primary", use_container_width=True)

# LÓGICA DE CARGA
if btn_cargar:
    if informe_sel and informe_sel != "-- Seleccionar informe resguardado --":
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
            st.session_state['drive_link'] = datos_cargados['drive_link']

            sec_loaded = datos_cargados['secciones_dinamicas']
            for s_num in [1, 2, 3, 4]:
                sub_list = sec_loaded.get(s_num, {}).get('subpuntos', [])
                st.session_state[f"cant_subpuntos_sec_{s_num}"] = len(sub_list)
                for idx, sub in enumerate(sub_list, start=1):
                    st.session_state[f"tit_{s_num}_{idx}"] = sub.get("titulo", "")
                    st.session_state[f"cont_{s_num}_{idx}"] = sub.get("contenido", "")

            st.session_state["imagenes_cargadas_resguardo"] = datos_cargados["imagenes_procesadas"]
            st.session_state["esquemas_cargados_resguardo"] = datos_cargados["esquemas_procesados"]
            st.success(f"¡Informe '{informe_sel}' cargado correctamente!")
            st.rerun()
    else:
        st.warning("Selecciona un informe válido para cargar.")

# LÓGICA DE ELIMINACIÓN
if btn_eliminar:
    if not informe_sel or informe_sel == "-- Seleccionar informe resguardado --":
        st.warning("Por favor selecciona un informe de la lista antes de intentar eliminar.")
    elif not confirmar_eliminar:
        st.error("Por seguridad, debes marcar la casilla '⚠️ Confirmar eliminación' antes de eliminar.")
    else:
        with st.spinner("Eliminando informe de Google Sheets..."):
            exito, msg = eliminar_informe_guardado(informe_sel)
            if exito:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

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

# SECCIONES TÉCNICAS DINÁMICAS
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
            st.caption("ℹ️ *Sin subpuntos.*")

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

# REGISTROS FOTOGRÁFICOS Y ESQUEMAS
st.markdown("---")
st.markdown("#### 5. Registros Fotográficos y 6. Esquemas")

imagenes_procesadas = []
esquemas_procesados = []

drive_folder_input = st.text_input(
    "Pegar URL o ID de Carpeta en Google Drive con las fotos y esquemas:", 
    key="drive_link",
    placeholder="Ej: https://drive.google.com/drive/folders/1RPXZ8oUmM2eC1U6p3u3FlepKg6hvRLQC"
)

if st.button("📥 CARGAR ARCHIVOS DESDE GOOGLE DRIVE", use_container_width=True):
    if drive_folder_input:
        with st.spinner("Descargando fotografías y esquemas desde Google Drive..."):
            imgs_drive, esq_drive, msg_drive = obtener_imagenes_desde_drive_folder(drive_folder_input)
            if imgs_drive or esq_drive:
                st.session_state["imagenes_cargadas_resguardo"] = imgs_drive
                st.session_state["esquemas_cargados_resguardo"] = esq_drive
                st.success(msg_drive)
                st.rerun()
            else:
                st.error(msg_drive)

# Despliegue de Fotografías Normales
if st.session_state.get("imagenes_cargadas_resguardo"):
    st.markdown("##### 📸 5. Registros Fotográficos")
    imgs_res = st.session_state["imagenes_cargadas_resguardo"]
    
    cols = st.columns(3)
    for index, (img_bytes, pie_orig) in enumerate(imgs_res):
        with cols[index % 3]:
            st.image(img_bytes, caption=f"Foto {index+1}", use_container_width=True)
            pie_foto = st.text_input(f"Pie de foto {index+1}:", value=pie_orig, key=f"img_resguardada_{index}")
            imagenes_procesadas.append((img_bytes, pie_foto))

# Despliegue de Esquemas a Tamaño Completo
if st.session_state.get("esquemas_cargados_resguardo"):
    st.markdown("---")
    st.markdown("##### 📐 6. Esquema de Equipo")
    esq_res = st.session_state["esquemas_cargados_resguardo"]
    
    for index, (esq_bytes, pie_esq_orig) in enumerate(esq_res):
        st.image(esq_bytes, caption=f"Esquema {index+1}", use_container_width=True)
        pie_esq = st.text_input(f"Leyenda / Nombre ({index+1}):", value=pie_esq_orig, key=f"esq_resguardado_{index}")
        esquemas_procesados.append((esq_bytes, pie_esq))

# RESPONSABLE DEL INFORME
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

# RESGUARDO Y DESCARGAS
st.markdown("---")
st.markdown("#### 💾 Resguardo del Informe")

if st.button("💾 RESGUARDAR INFORME EN GOOGLE SHEETS", type="primary", use_container_width=True):
    with st.spinner("Guardando resguardo del informe en Google Sheets..."):
        exito, msg = guardar_resguardo_informe(datos_encabezado, secciones_dinamicas, drive_folder_input, inspector_firma)
        if exito:
            st.success(msg)
        else:
            st.error(msg)

st.markdown("---")
st.markdown("#### Exportar Informe Final")

col_btn_p, col_btn_w = st.columns(2)

with col_btn_p:
    pdf_buffer = generar_pdf_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, esquemas_procesados, inspector_firma)
    st.download_button(
        label="📄 Descargar Informe en PDF (.pdf)",
        data=pdf_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'INSPECCION'}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

with col_btn_w:
    word_buffer = generar_word_plantilla_inspeccion(datos_encabezado, secciones_dinamicas, imagenes_procesadas, esquemas_procesados, inspector_firma)
    st.download_button(
        label="📝 Descargar Informe en Word (.docx)",
        data=word_buffer,
        file_name=f"INFORME_VISUAL_{num_informe if num_informe else 'INSPECCION'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
