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
from reportlab.pdfgen import canvas
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
    """Autentica y devuelve la hoja de trabajo activa."""
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
    return client.open_by_key(SPREADSHEET_ID).sheet1

def cargar_datos_sheets():
    """Lee todas las filas almacenadas en la Hoja de forma directa."""
    try:
        sheet = conectar_google_sheets()
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

# =========================================================
# CLASE CANVAS PARA PLANTILLA CORPORATIVA EN PDF
# =========================================================
class PlantillaCorporativaCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.dibujar_elementos_corporativos(num_pages)
            super().showPage()
        super().save()

    def dibujar_elementos_corporativos(self, page_count):
        # 0. Fondo de toda la hoja color #f8faf6
        self.setFillColor(colors.HexColor("#f8faf6"))
        self.rect(0, 0, 612, 792, fill=1, stroke=0)

        # 1. Franja superior verde #619b40 (2 cm exactos)
        self.setFillColor(colors.HexColor("#619b40"))
        self.rect(0, 792 - (2 * cm), 612, 2 * cm, fill=1, stroke=0)

        # 2. Logo en la parte SUPERIOR IZQUIERDA
        logo_bytes = obtener_bytes_logo()
        if logo_bytes:
            try:
                img_stream = io.BytesIO(logo_bytes)
                img = ImageReader(img_stream)
                # Posicionado justo debajo de la franja verde, alineado a la izquierda (X=30)
                self.drawImage(img, 30, 792 - (2 * cm) - 55, width=120, height=45, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                pass

        # 3. Texto Institucional Pie de Página
        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor("#6B7280"))
        self.drawCentredString(306, 26, "SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.")
        self.drawCentredString(306, 16, "CONTRATO N° AC 31104857")
        
        # 4. Numeración de página
        self.drawRightString(582, 16, f"Pág. {self._pageNumber} de {page_count}")

# =========================================================
# FUNCIÓN PARA GENERAR PDF CON TABLAS POR INSPECTOR Y FRANJA NARANJA
# =========================================================
def generar_pdf_informe(df_fecha, fecha_str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=4.5 * cm, # Ampliado para dejar espacio al logo superior
        bottomMargin=2.2 * cm
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontSize=18, leading=22,
        textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle', parent=styles['Normal'], fontSize=11, leading=14,
        textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=14
    )
    inspector_heading_style = ParagraphStyle(
        'InspectorHeader', parent=styles['Heading2'], fontSize=11, leading=14,
        textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=6
    )
    cell_header_style = ParagraphStyle(
        'HeaderStyle', parent=styles['Normal'], fontSize=8.5, leading=10,
        textColor=colors.white, fontName='Helvetica-Bold', alignment=1
    )
    cell_body_style = ParagraphStyle(
        'BodyStyle', parent=styles['Normal'], fontSize=8, leading=10,
        textColor=colors.HexColor('#1F2937')
    )

    story = [
        Paragraph("INFORME INSPECCIÓN", title_style),
        Paragraph(f"<b>Fecha del Reporte:</b> {fecha_str}", subtitle_style),
        Spacer(1, 6)
    ]

    if not df_fecha.empty:
        inspectores_grupos = df_fecha.groupby('inspector', sort=False)
        total_grupos = len(inspectores_grupos)
        
        for idx, (inspector_nom, group_df) in enumerate(inspectores_grupos):
            # Encabezado del Inspector
            story.append(Paragraph(f"👷‍♂️ Inspector: <b>{inspector_nom}</b>", inspector_heading_style))
            
            table_data = [[
                Paragraph("Planta / Tag", cell_header_style),
                Paragraph("Actividad Realizada", cell_header_style),
                Paragraph("Avance", cell_header_style),
                Paragraph("Estado", cell_header_style),
                Paragraph("Observaciones", cell_header_style)
            ]]
            
            for _, row in group_df.iterrows():
                planta = row.get("planta", "-")
                tag = row.get("tag_equipo", "-")
                actividad = row.get("actividad_realizada", "-")
                avance = row.get("avance", "-")
                estado = row.get("estado_liberacion", "-")
                obs = row.get("observaciones", "-")
                
                table_data.append([
                    Paragraph(f"<b>Planta:</b> {planta}<br/><b>TAG:</b> {tag}", cell_body_style),
                    Paragraph(actividad, cell_body_style),
                    Paragraph(avance, cell_body_style),
                    Paragraph(estado, cell_body_style),
                    Paragraph(obs if obs else "-", cell_body_style)
                ])
                
            t = Table(table_data, colWidths=[100, 160, 50, 92, 150])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
                # Fondo intercalado entre blanco y gris claro corporativo
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F3F4F6')]),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))

            # Franja divisoria Naranja Corporativa
            if idx < total_grupos - 1:
                divider_table = Table([[""]], colWidths=[552], rowHeights=[4])
                divider_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F97316')),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))
                story.append(divider_table)
                story.append(Spacer(1, 10))

    doc.build(story, canvasmaker=PlantillaCorporativaCanvas)
    buffer.seek(0)
    return buffer

# =========================================================
# FUNCIÓN PARA GENERAR WORD CON TABLAS POR INSPECTOR Y FRANJA NARANJA
# =========================================================
def generar_word_informe(df_fecha, fecha_str):
    doc = Document()
    
    # Inyectar color de fondo #F8FAF6 en el documento Word
    try:
        background = parse_xml(r'<w:background {} w:color="F8FAF6"/>'.format(nsdecls('w')))
        doc.element.insert(0, background)
    except Exception:
        pass

    section = doc.sections[0]
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(1.0)

    # 1. Franja verde superior (2 cm)
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

    # 2. Logo en parte superior izquierda del cuerpo
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

    # 3. Pie de página fijo en Word
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text_run = footer_p.add_run("SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS DE ENAP REFINERÍAS S.A.\nCONTRATO N° AC 31104857")
    text_run.font.size = Pt(7)
    text_run.font.color.rgb = RGBColor(107, 114, 128)

    # 4. Título y fecha
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("INFORME INSPECCIÓN")
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)
    
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_p.add_run(f"Fecha del Reporte: {fecha_str}")
    date_run.font.size = Pt(11)
    date_run.font.color.rgb = RGBColor(75, 85, 99)
    
    doc.add_paragraph()

    # 5. Tablas agrupadas por Inspector
    if not df_fecha.empty:
        inspectores_grupos = list(df_fecha.groupby('inspector', sort=False))
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
            headers = ["Planta / TAG", "Actividad Realizada", "Avance", "Estado", "Observaciones"]
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
                row_cells[0].text = f"Planta: {row.get('planta', '-')}\nTAG: {row.get('tag_equipo', '-')}"
                row_cells[1].text = str(row.get("actividad_realizada", "-"))
                row_cells[2].text = str(row.get("avance", "-"))
                row_cells[3].text = str(row.get("estado_liberacion", "-"))
                row_cells[4].text = str(row.get("observaciones", "-")) if row.get("observaciones") else "-"
                
                # Fondo condicional opcional para Word si es necesario
                for cell in row_cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(8)

            doc.add_paragraph()

            # Franja Naranja entre distintos inspectores
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
    "Inspector 3",
    "Inspector 4",
    "Inspector 5"
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
    ["📝 Registrar Actividad por Inspector", "📊 Historial e Informes"]
)

# =========================================================
# MÓDULO 1: REGISTRO DE ACTIVIDADES (ESCRITURA)
# =========================================================
if menu == "📝 Registrar Actividad por Inspector":
    st.subheader("📋 Formulario de Ingreso de Actividades")
    
    with st.form("form_actividades_inspector", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        semana_actual_num = datetime.now().isocalendar()[1]
        idx_semana_defecto = min(semana_actual_num - 1, 51)
        
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
                    sheet = conectar_google_sheets()
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
    st.subheader("🔍 Consulta de Historial y Generación de Informes")
    
    with st.spinner("Cargando registros desde Google Sheets..."):
        df_historial = cargar_datos_sheets()
    
    st.markdown("### 📄 Exportar Informe Diarios")
    col_pdf1, col_pdf2, col_pdf3 = st.columns([2, 1.5, 1.5])
    
    with col_pdf1:
        fecha_informe = st.date_input("🗓️ Selecciona la Fecha del Informe:", datetime.now())
        fecha_informe_str = str(fecha_informe)
    
    df_informe = df_historial[df_historial['fecha'] == fecha_informe_str] if not df_historial.empty and 'fecha' in df_historial.columns else pd.DataFrame()
    
    if not df_informe.empty:
        with col_pdf2:
            st.write("&nbsp;")
            pdf_bytes = generar_pdf_informe(df_informe, fecha_informe_str)
            st.download_button(
                label="📄 Descargar Informe (PDF)",
                data=pdf_bytes,
                file_name=f"INFORME_INSPECCION_{fecha_informe_str}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col_pdf3:
            st.write("&nbsp;")
            word_bytes = generar_word_informe(df_informe, fecha_informe_str)
            st.download_button(
                label="📝 Descargar Informe (Word)",
                data=word_bytes,
                file_name=f"INFORME_INSPECCION_{fecha_informe_str}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.info("No hay registros guardados para esta fecha para generar informes.")

    st.markdown("---")
    st.markdown("### 📊 Tabla de Consulta e Historial Completo")

    if not df_historial.empty:
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

        st.markdown(f"**Total de registros encontrados:** `{len(df_filtrado)}`")
        
        st.dataframe(
            df_filtrado, 
            use_container_width=True
        )
        
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar copia local a CSV",
            data=csv_data,
            file_name=f"historial_actividades_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Aún no hay registros de actividades guardados en la nube.")
