import io
import requests
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Gestión de Inspectores en la Nube",
    page_icon="☁️",
    layout="wide"
)

# 🔗 URL RAW DE TU LOGO EN GITHUB
URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"

# 🆔 ID DE TU HOJA DE GOOGLE SHEETS
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

# =========================================================
# DESCARGA Y CACHÉ DEL LOGO EN MEMORIA
# =========================================================
@st.cache_data(ttl=3600)
def obtener_bytes_logo():
    """Descarga el logo una sola vez y lo mantiene en caché."""
    try:
        response = requests.get(URL_LOGO_GITHUB, timeout=5)
        if response.status_code == 200:
            return response.content
    except Exception:
        pass
    return None

# =========================================================
# CONEXIÓN DIRECTA CON GOOGLE SHEETS VIA GSPREAD
# =========================================================
def conectar_google_sheets():
    """Autentica y devuelve el cliente de Google Sheets."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    
    if "private_key" in creds_dict:
        pk = creds_dict["private_key"]
        pk = pk.replace("\\n", "\n").strip()
        lines = [line.strip() for line in pk.split("\n") if line.strip()]
        creds_dict["private_key"] = "\n".join(lines) + "\n"

    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID)

def obtener_hoja_actividades():
    client = conectar_google_sheets()
    return client.sheet1

def obtener_hoja_planificacion():
    client = conectar_google_sheets()
    try:
        return client.worksheet("Planificacion")
    except gspread.exceptions.WorksheetNotFound:
        ws = client.add_worksheet(title="Planificacion", rows=100, cols=25)
        ws.append_row([
            "planta", "tipo_informe", "circuito_equipo", "fecha_inicio", "fecha_fin", 
            "contador_inspeccion", "programa", "inspector_dci", "inspector_ingemars", 
            "dias_ing_enap", "otep", "informe_iv", "status", "fecha_entrega_ope", 
            "contador_liberacion", "liberacion_final", "contador_enap", 
            "contador_cumplimiento_enap", "contador_dias_informe", "archivo_url", "observaciones"
        ])
        return ws

def cargar_datos_sheets():
    """Lee todas las filas almacenadas en la Hoja de Actividades."""
    try:
        sheet = obtener_hoja_actividades()
        filas = sheet.get_all_values()
        
        if len(filas) <= 1:
            return pd.DataFrame(columns=[
                "fecha", "semana", "planta", "inspector", "tag_equipo", 
                "actividad_realizada", "avance", "observaciones", "estado_liberacion"
            ])
            
        encabezados = [c.strip().lower() for c in filas[0]]
        datos = filas[1:]
        
        df = pd.DataFrame(datos, columns=encabezados)
        return df
    except Exception as e:
        st.error(f"Error al leer la hoja de Google Sheets: {e}")
        return pd.DataFrame(columns=[
            "fecha", "semana", "planta", "inspector", "tag_equipo", 
            "actividad_realizada", "avance", "observaciones", "estado_liberacion"
        ])

def cargar_datos_planificacion():
    """Lee todas las filas almacenadas en la Hoja de Planificación."""
    try:
        sheet = obtener_hoja_planificacion()
        filas = sheet.get_all_values()
        
        if len(filas) <= 1:
            return pd.DataFrame(columns=[
                "planta", "tipo_informe", "circuito_equipo", "fecha_inicio", "fecha_fin", 
                "contador_inspeccion", "programa", "inspector_dci", "inspector_ingemars", 
                "dias_ing_enap", "otep", "informe_iv", "status", "fecha_entrega_ope", 
                "contador_liberacion", "liberacion_final", "contador_enap", 
                "contador_cumplimiento_enap", "contador_dias_informe", "archivo_url", "observaciones"
            ])
            
        encabezados = [c.strip().lower() for c in filas[0]]
        datos = filas[1:]
        
        df = pd.DataFrame(datos, columns=encabezados)
        return df
    except Exception as e:
        st.error(f"Error al leer la hoja de Planificación: {e}")
        return pd.DataFrame()

# =========================================================
# DISEÑO DE PLANTILLA (FONDO, CABECERA Y PIE PARA PDF)
# =========================================================
def dibujar_plantilla(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#f8faf6"))
    canvas.rect(0, 0, letter[0], letter[1], fill=1, stroke=0)

    canvas.setFillColor(colors.HexColor("#619b40"))
    canvas.rect(0, letter[1] - (2 * cm), letter[0], 2 * cm, fill=1, stroke=0)

    logo_bytes = obtener_bytes_logo()
    if logo_bytes:
        try:
            img_stream = io.BytesIO(logo_bytes)
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

def generar_pdf_informe(df_filtrado, titulo_doc, subtitulo_doc):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=4.5 * cm,
        bottomMargin=2.2 * cm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=4)
    subtitle_style = ParagraphStyle('DocSubTitle', parent=styles['Normal'], fontSize=11, leading=14, textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=14)
    inspector_heading_style = ParagraphStyle('InspectorHeader', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=6)
    cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8.5, leading=10, textColor=colors.white, fontName='Helvetica-Bold', alignment=1)
    cell_body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))

    story = [
        Paragraph(titulo_doc, title_style),
        Paragraph(f"<b>{subtitulo_doc}</b>", subtitle_style),
        Spacer(1, 6)
    ]

    if not df_filtrado.empty:
        inspectores_grupos = df_filtrado.groupby('inspector', sort=False)
        total_grupos = len(inspectores_grupos)
        
        for idx, (inspector_nom, group_df) in enumerate(inspectores_grupos):
            story.append(Paragraph(f"👷‍♂️ Inspector: <b>{inspector_nom}</b>", inspector_heading_style))
            
            table_data = [[
                Paragraph("Fecha / Planta", cell_header_style),
                Paragraph("Actividad Realizada", cell_header_style),
                Paragraph("Avance", cell_header_style),
                Paragraph("Estado", cell_header_style),
                Paragraph("Observaciones", cell_header_style)
            ]]
            
            for _, row in group_df.iterrows():
                fecha_reg = row.get("fecha", "-")
                planta = row.get("planta", "-")
                tag = row.get("tag_equipo", "-")
                actividad = row.get("actividad_realizada", "-")
                avance = row.get("avance", "-")
                estado = row.get("estado_liberacion", "-")
                obs = row.get("observaciones", "-")
                
                table_data.append([
                    Paragraph(f"<b>Fecha:</b> {fecha_reg}<br/><b>Planta:</b> {planta}<br/><b>TAG:</b> {tag}", cell_body_style),
                    Paragraph(actividad, cell_body_style),
                    Paragraph(avance, cell_body_style),
                    Paragraph(estado, cell_body_style),
                    Paragraph(obs if obs else "-", cell_body_style)
                ])
                
            t = Table(table_data, colWidths=[110, 150, 50, 92, 150])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F3F4F6')]),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))

            if idx < total_grupos - 1:
                divider_table = Table([[""]], colWidths=[552], rowHeights=[4])
                divider_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F97316')),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))
                story.append(divider_table)
                story.append(Spacer(1, 10))

    doc.build(story, onFirstPage=dibujar_plantilla, onLaterPages=dibujar_plantilla)
    buffer.seek(0)
    return buffer

def generar_word_informe(df_filtrado, titulo_doc, subtitulo_doc):
    doc = Document()
    try:
        background = parse_xml(r'<w:background {} w:color="F8FAF6"/>'.format(nsdecls('w')))
        doc.element.insert(0, background)
    except Exception:
        pass

    section = doc.sections[0]
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(1.0)

    header = section.header
    banner_table = header.add_table(rows=1, cols=1, width=Inches(6.5))
    banner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_h = banner_table.rows[0].cells[0]
    cell_h.width = Inches(6.5)
    
    shading_elm = parse_xml(r'<w:shd {} w:fill="619b40"/>'.format(nsdecls('w')))
    cell_h._tc.get_or_add_tcPr().append(shading_elm)
    
    p_banner = cell_h.paragraphs[0]
    p_banner.paragraph_format.space_before = Pt(20)
    p_banner.paragraph_format.space_after = Pt(20)

    logo_bytes = obtener_bytes_logo()
    if logo_bytes:
        try:
            img_stream = io.BytesIO(logo_bytes)
            logo_p = doc.add_paragraph()
            logo_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            logo_p.paragraph_format.space_after = Pt(4)
            logo_run = logo_p.add_run()
            logo_run.add_picture(img_stream, width=Inches(1.5))
        except Exception:
            pass

    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text_run = footer_p.add_run("SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.\nCONTRATO N° AC 31104857")
    text_run.font.size = Pt(7)
    text_run.font.color.rgb = RGBColor(107, 114, 128)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(titulo_doc)
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)
    
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_p.add_run(subtitulo_doc)
    date_run.font.size = Pt(11)
    date_run.font.color.rgb = RGBColor(75, 85, 99)
    
    doc.add_paragraph()

    if not df_filtrado.empty:
        inspectores_grupos = list(df_filtrado.groupby('inspector', sort=False))
        total_grupos = len(inspectores_grupos)

        for idx, (inspector_nom, group_df) in enumerate(inspectores_grupos):
            insp_p = doc.add_paragraph()
            insp_run = insp_p.add_run(f"Inspector: {inspector_nom}")
            insp_run.font.size = Pt(11)
            insp_run.font.bold = True
            insp_run.font.color.rgb = RGBColor(30, 58, 138)
            
            table = doc.add_table(rows=1, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            headers = ["Fecha / Planta", "Actividad Realizada", "Avance", "Estado", "Observaciones"]
            for i, header_text in enumerate(headers):
                hdr_cells[i].text = header_text
                p = hdr_cells[i].paragraphs[0]
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(8.5)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                shd = parse_xml(r'<w:shd {} w:fill="1E3A8A"/>'.format(nsdecls('w')))
                hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

            for _, row in group_df.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = f"Fecha: {row.get('fecha', '-')}\nPlanta: {row.get('planta', '-')}\nTAG: {row.get('tag_equipo', '-')}"
                row_cells[1].text = str(row.get("actividad_realizada", "-"))
                row_cells[2].text = str(row.get("avance", "-"))
                row_cells[3].text = str(row.get("estado_liberacion", "-"))
                row_cells[4].text = str(row.get("observaciones", "-")) if row.get("observaciones") else "-"
                
                for cell in row_cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(8)

            doc.add_paragraph()

            if idx < total_grupos - 1:
                div_table = doc.add_table(rows=1, cols=1)
                div_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                div_cell = div_table.rows[0].cells[0]
                div_cell.width = Inches(6.5)
                shd_orange = parse_xml(r'<w:shd {} w:fill="F97316"/>'.format(nsdecls('w')))
                div_cell._tc.get_or_add_tcPr().append(shd_orange)
                p_div = div_cell.paragraphs[0]
                p_div.paragraph_format.space_before = Pt(2)
                p_div.paragraph_format.space_after = Pt(2)
                doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# LISTAS Y CONFIGURACIONES
# =========================================================
LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Arlem Sarmiento",
    "Harold Castillo",
    "Miguel Chirinos"
]

LISTA_INSPECTORES_DCI = [
    "Eduardo Delgado",
    "Felipe Ponce",
    "José de la Cruz",
    "Pablo Ruiz",
    "Luis Durán"
]

ESTADOS_STATUS = [
    "Finalizado",
    "En curso",
    "Pendiente"
]

TIPOS_INFORME = [
    "IV",
    "Ensayo Dureza",
    "Ensayo LP",
    "IHV",
    "Boroscopio"
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

LISTA_SEMANAS = [f"Semana {i}" for i in range(1, 53)]

ESTADOS_LIBERACION = [
    "Liberado / Conforme (Aprobado)",
    "Pendiente de Reparación",
    "Rechazado",
    "En Proceso de Inspección",
    "En Espera de END / Pruebas"
]

# Cálculo de semana por defecto global
semana_actual_num = datetime.now().isocalendar()[1]
idx_semana_defecto = max(0, min(semana_actual_num - 1, len(LISTA_SEMANAS) - 1))

# =========================================================
# ENCABEZADO Y VISUALIZACIÓN DEL LOGO
# =========================================================
try:
    st.sidebar.image(URL_LOGO_GITHUB, use_container_width=True)
except Exception:
    pass

col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
        st.image(URL_LOGO_GITHUB, width=140)
    except Exception:
        st.write("📂 [Logo]")

with col_titulo:
    st.title("Control de Actividades por Inspector")
    st.markdown("##### *Sistema de Gestión QA/QC Sincronizado en la Nube*")

st.markdown("---")

# =========================================================
# MENÚ Y NAVEGACIÓN
# =========================================================
menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    [
        "📝 Registrar Actividad por Inspector", 
        "📊 Historial e Informes", 
        "📈 Reporte Planificación"
    ]
)

# =========================================================
# MÓDULO 1: REGISTRO DE ACTIVIDADES (ESCRITURA)
# =========================================================
if menu == "📝 Registrar Actividad por Inspector":
    st.subheader("📋 Formulario de Ingreso de Actividades")
    
    with st.form("form_actividades_inspector", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            inspector_seleccionado = st.selectbox("👷‍♂️ Seleccionar Inspector asignado:", LISTA_INSPECTORES)
            fecha_actividad = st.date_input("📅 Fecha de Inspección:", datetime.now())
            semana_seleccionada = st.selectbox("🗓️ Semana Operativa:", LISTA_SEMANAS, index=idx_semana_defecto)
            planta_seleccionada = st.selectbox("🏭 Planta / Unidad:", LISTA_PLANTAS)

        with col2:
            tag_equipo = st.text_input("🏷️ TAG del Equipo / Línea Piping:", placeholder="Ej: C-1302 / E-2101 / PIP-001")
            porcentaje_avance = st.slider("📊 Porcentaje de Avance:", min_value=0, max_value=100, value=0, step=5, format="%d%%")
            estado_liberacion = st.selectbox("📌 Estado de la Inspección:", ESTADOS_LIBERACION)

        actividad_realizada = st.text_area("🛠️ Actividades Realizadas por el Inspector:", placeholder="Ej: Inspección visual de junta...")
        observaciones = st.text_area("💬 Observaciones Adicionales / Recomendaciones:", placeholder="Escribe comentarios extra...")
        
        btn_guardar = st.form_submit_button("☁️ Guardar Registro en Google Sheets")
        
        if btn_guardar:
            if tag_equipo and actividad_realizada:
                try:
                    sheet = obtener_hoja_actividades()
                    todas_las_filas = sheet.get_all_values()
                    if len(todas_las_filas) == 0:
                        sheet.append_row([
                            "fecha", "semana", "planta", "inspector", "tag_equipo", 
                            "actividad_realizada", "avance", "observaciones", "estado_liberacion"
                        ])
                    
                    nueva_fila = [
                        str(fecha_actividad),
                        semana_seleccionada,
                        planta_seleccionada,
                        inspector_seleccionado,
                        tag_equipo.strip(),
                        actividad_realizada.strip(),
                        f"{porcentaje_avance}%",
                        observaciones.strip(),
                        estado_liberacion
                    ]
                    sheet.append_row(nueva_fila)
                    st.success(f"✅ ¡Actividad de **{inspector_seleccionado}** (Planta: {planta_seleccionada} | TAG: {tag_equipo}) guardada con éxito!")
                except Exception as ex:
                    st.error(f"❌ Ocurrió un error al guardar en la nube: {ex}")
            else:
                st.error("⚠️ Por favor completa los campos obligatorios: **TAG del Equipo** y **Actividades Realizadas**.")

# =========================================================
# MÓDULO 2: HISTORIAL E INFORMES EN WORD Y PDF
# =========================================================
elif menu == "📊 Historial e Informes":
    st.subheader("🔍 Consulta de Historial y Generación de Informes por Día o Semana")
    
    with st.spinner("Cargando registros desde Google Sheets..."):
        df_historial = cargar_datos_sheets()
    
    if not df_historial.empty:
        st.markdown("### 🎛️ Filtros Interactivos de Consulta")
        col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
        
        with col_f1:
            filtro_inspector = st.selectbox("Filtrar por Inspector:", ["Todos"] + LISTA_INSPECTORES)
        with col_f2:
            filtro_semana = st.selectbox("Filtrar por Semana:", ["Todas"] + LISTA_SEMANAS)
        with col_f3:
            filtro_planta = st.selectbox("Filtrar por Planta:", ["Todas"] + LISTA_PLANTAS)
        with col_f4:
            filtro_tag = st.text_input("Filtrar por TAG:")
        with col_f5:
            filtro_estado = st.selectbox("Filtrar por Estado:", ["Todos"] + ESTADOS_LIBERACION)

        df_filtrado = df_historial.copy()
        
        if "inspector" in df_filtrado.columns and filtro_inspector != "Todos":
            df_filtrado = df_filtrado[df_filtrado['inspector'] == filtro_inspector]
            
        if "semana" in df_filtrado.columns and filtro_semana != "Todas":
            df_filtrado = df_filtrado[df_filtrado['semana'] == filtro_semana]

        if "planta" in df_filtrado.columns and filtro_planta != "Todas":
            df_filtrado = df_filtrado[df_filtrado['planta'] == filtro_planta]

        if "tag_equipo" in df_filtrado.columns and filtro_tag:
            df_filtrado = df_filtrado[df_filtrado['tag_equipo'].astype(str).str.contains(filtro_tag, case=False, na=False)]
            
        if "estado_liberacion" in df_filtrado.columns and filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['estado_liberacion'] == filtro_estado]

        st.markdown(f"**Total de registros filtrados:** `{len(df_filtrado)}`")
        
        st.dataframe(df_filtrado, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📄 Exportar Informe (Diario o Semanal)")
        
        col_tipo, col_periodo = st.columns([1.5, 2])
        
        with col_tipo:
            tipo_reporte = st.radio("📌 Frecuencia del Reporte:", ["📅 Diario", "🗓️ Semanal"], horizontal=True)
        
        if tipo_reporte == "📅 Diario":
            with col_periodo:
                fecha_informe_dt = st.date_input("📅 Selecciona el Día del Informe:", datetime.now())
                fecha_str_fmt = fecha_informe_dt.strftime('%d/%m/%Y')
                fecha_str_comparar = str(fecha_informe_dt)
                
                titulo_doc = "REPORTE DIARIO DE INSPECCIÓN"
                subtitulo_doc = f"Fecha: {fecha_str_fmt}"
                tag_archivo = f"DIARIO_{fecha_informe_dt.strftime('%Y%m%d')}"
            
            if "fecha" in df_historial.columns:
                df_informe = df_historial[df_historial['fecha'].astype(str) == fecha_str_comparar]
            else:
                df_informe = pd.DataFrame()
                
        else: # 🗓️ Semanal
            with col_periodo:
                semana_informe = st.selectbox("🗓️ Selecciona la Semana para el Informe:", LISTA_SEMANAS, index=idx_semana_defecto)
                
                titulo_doc = "REPORTE SEMANAL DE INSPECCIÓN"
                subtitulo_doc = f"Período: {semana_informe}"
                tag_archivo = f"SEMANAL_{semana_informe.replace(' ', '_')}"
            
            if "semana" in df_historial.columns:
                df_informe = df_historial[df_historial['semana'] == semana_informe]
            else:
                df_informe = pd.DataFrame()

        col_exp1, col_exp2 = st.columns(2)
        
        if not df_informe.empty:
            with col_exp1:
                pdf_bytes = generar_pdf_informe(df_informe, titulo_doc, subtitulo_doc)
                st.download_button(
                    label=f"📄 Descargar {titulo_doc} (PDF)",
                    data=pdf_bytes,
                    file_name=f"INFORME_INSPECCION_{tag_archivo}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col_exp2:
                word_bytes = generar_word_informe(df_informe, titulo_doc, subtitulo_doc)
                st.download_button(
                    label=f"📝 Descargar {titulo_doc} (Word)",
                    data=word_bytes,
                    file_name=f"INFORME_INSPECCION_{tag_archivo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.warning(f"⚠️ No hay registros guardados en la nube para: **{subtitulo_doc}**.")

        st.markdown("---")
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar tabla filtrada actual a CSV",
            data=csv_data,
            file_name=f"historial_actividades_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de actividades guardados en la nube.")

# =========================================================
# MÓDULO 3: REPORTE PLANIFICACIÓN
# =========================================================
elif menu == "📈 Reporte Planificación":
    st.subheader("📅 Módulo de Registro y Control de Planificación")
    st.markdown("##### *Ingrese los datos detallados de planificación para sincronizar con la nube.*")
    
    with st.form("form_planificacion", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            p_planta = st.selectbox("🏭 Planta:", LISTA_PLANTAS)
            p_tipo_informe = st.selectbox("📋 Tipo de Informe:", TIPOS_INFORME)
            p_circuito = st.text_input("🔧 Circuito / Equipo:", placeholder="Ej: C-1302 / Línea 12\"...")
            p_f_inicio = st.date_input("📅 Fecha Inicio Inspección:", datetime.now())
            p_f_fin = st.date_input("📅 Fecha FIN Inspección:", datetime.now())
            p_cont_insp = st.number_input("🔢 Contador Inspección:", min_value=0, value=0)
            p_programa = st.text_input("📌 PROGRAMA:", placeholder="Ej: Programa 2026...")

        with col2:
            p_insp_dci = st.selectbox("👷‍♂️ Inspector DCI:", LISTA_INSPECTORES_DCI)
            p_insp_ingemars = st.selectbox("👷‍♂️ Inspector Ingemars:", ["Todos"] + LISTA_INSPECTORES)
            p_dias_enap = st.number_input("⏱️ Días entregados por ing. ENAP:", min_value=0, value=0)
            p_otep = st.text_input("📄 N° OTEP:", placeholder="Ej: OTEP-9988...")
            p_informe_iv = st.text_input("📑 N° INFORME IV:", placeholder="Ej: IV-2026-01...")
            p_status = st.selectbox("📌 Status:", ESTADOS_STATUS)
            p_f_entrega_ope = st.date_input("📅 Fecha entrega (Ing. Ope.):", datetime.now())

        with col3:
            p_cont_lib = st.number_input("🔢 Contador entrega liberación (INS - ING.OP):", min_value=0, value=0)
            p_lib_final = st.text_input("✅ Liberación informe final:", placeholder="Sí / No / Parcial")
            p_cont_enap = st.number_input("🔢 Contador (INS - ENAP):", min_value=0, value=0)
            p_cont_cump_enap = st.number_input("📊 Contador de cumplimiento Ing. ENAP:", min_value=0, value=0)
            p_cont_dias_inf = st.number_input("📉 Contador días de informe:", min_value=0, value=0)
            p_archivo_url = st.text_input("🔗 Archivo URL:", placeholder="https://...")

        p_observaciones = st.text_area("💬 Observaciones:", placeholder="Notas de planificación...")

        btn_guardar_plan = st.form_submit_button("☁️ Guardar Planificación en Google Sheets")
        
        if btn_guardar_plan:
            try:
                ws_plan = obtener_hoja_planificacion()
                nueva_fila_plan = [
                    p_planta,
                    p_tipo_informe,
                    p_circuito.strip(),
                    str(p_f_inicio),
                    str(p_f_fin),
                    str(p_cont_insp),
                    p_programa.strip(),
                    p_insp_dci,
                    p_insp_ingemars,
                    str(p_dias_enap),
                    p_otep.strip(),
                    p_informe_iv.strip(),
                    p_status,
                    str(p_f_entrega_ope),
                    str(p_cont_lib),
                    p_lib_final.strip(),
                    str(p_cont_enap),
                    str(p_cont_cump_enap),
                    str(p_cont_dias_inf),
                    p_archivo_url.strip(),
                    p_observaciones.strip()
                ]
                ws_plan.append_row(nueva_fila_plan)
                st.success("✅ ¡Datos de planificación guardados con éxito en la pestaña 'Planificacion' de Google Sheets!")
            except Exception as ex:
                st.error(f"❌ Error al guardar planificación: {ex}")

    st.markdown("---")
    st.markdown("### 📊 Historial y Registros de Planificación en la Nube")
    
    with st.spinner("Cargando datos de planificación..."):
        df_plan = cargar_datos_planificacion()
        
    if not df_plan.empty:
        st.dataframe(df_plan, use_container_width=True)
        
        csv_plan = df_plan.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar Planificación a CSV",
            data=csv_plan,
            file_name=f"reporte_planificacion_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de planificación guardados.")
