import io
import base64
import requests
import streamlit as st
import pandas as pd
import gspread
import plotly.express as px
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
    page_title="Sistema de Gestión de Activos Físicos",
    page_icon="☁️",
    layout="wide"
)

# =========================================================
# CONTROL DE PERMISOS (MODO PÚBLICO DE LECTURA + ADMIN OPCIONAL)
# =========================================================
if "admin_logueado" not in st.session_state:
    st.session_state["admin_logueado"] = False

# =========================================================
# A PARTIR DE AQUÍ FUNCIONA LA APLICACIÓN PARA TODOS
# =========================================================

URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"
URL_FRANJA_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/franja.png"
URL_PERSONAJE_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/personaje.png"

SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

@st.cache_data(ttl=3600)
def obtener_bytes_imagen(url):
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return response.content
    except Exception:
        pass
    return None

def obtener_bytes_logo():
    return obtener_bytes_imagen(URL_LOGO_GITHUB)

def obtener_bytes_franja():
    return obtener_bytes_imagen(URL_FRANJA_GITHUB)

def obtener_bytes_personaje():
    return obtener_bytes_imagen(URL_PERSONAJE_GITHUB)

def conectar_google_sheets():
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
    try:
        sheet = obtener_hoja_actividades()
        filas = sheet.get_all_values()
        if len(filas) <= 1:
            return pd.DataFrame(columns=[
                "fecha", "semana", "planta", "inspector", "tag_equipo", 
                "actividad_realizada", "avance", "observaciones", "estado_liberacion"
            ])
        encabezados = [c.strip().lower() for c in filas[0]]
        return pd.DataFrame(filas[1:], columns=encabezados)
    except Exception as e:
        st.error(f"Error al leer la hoja de Google Sheets: {e}")
        return pd.DataFrame()

def cargar_datos_planificacion():
    try:
        sheet = obtener_hoja_planificacion()
        filas = sheet.get_all_values()
        if len(filas) <= 1:
            return pd.DataFrame()
        encabezados = [c.strip().lower() for c in filas[0]]
        return pd.DataFrame(filas[1:], columns=encabezados)
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
        buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=4.5 * cm, bottomMargin=2.2 * cm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=4)
    subtitle_style = ParagraphStyle('DocSubTitle', parent=styles['Normal'], fontSize=11, leading=14, textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=14)
    inspector_heading_style = ParagraphStyle('InspectorHeader', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=6)
    cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8.5, leading=10, textColor=colors.white, fontName='Helvetica-Bold', alignment=1)
    cell_body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))

    story = [Paragraph(titulo_doc, title_style), Paragraph(f"<b>{subtitulo_doc}</b>", subtitle_style), Spacer(1, 6)]

    if not df_filtrado.empty:
        inspectores_grupos = df_filtrado.groupby('inspector', sort=False)
        for inspector_nom, group_df in inspectores_grupos:
            story.append(Paragraph(f"👷‍♂️ Inspector: <b>{inspector_nom}</b>", inspector_heading_style))
            table_data = [[Paragraph("Fecha / Planta", cell_header_style), Paragraph("Actividad Realizada", cell_header_style), Paragraph("Avance", cell_header_style), Paragraph("Estado", cell_header_style), Paragraph("Observaciones", cell_header_style)]]
            for _, row in group_df.iterrows():
                table_data.append([
                    Paragraph(f"<b>Fecha:</b> {row.get('fecha', '-')}<br/><b>Planta:</b> {row.get('planta', '-')}<br/><b>TAG:</b> {row.get('tag_equipo', '-')}", cell_body_style),
                    Paragraph(str(row.get("actividad_realizada", "-")), cell_body_style),
                    Paragraph(str(row.get("avance", "-")), cell_body_style),
                    Paragraph(str(row.get("estado_liberacion", "-")), cell_body_style),
                    Paragraph(str(row.get("observaciones", "-")) if row.get("observaciones") else "-", cell_body_style)
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

    doc.build(story, onFirstPage=dibujar_plantilla, onLaterPages=dibujar_plantilla)
    buffer.seek(0)
    return buffer

def generar_word_informe(df_filtrado, titulo_doc, subtitulo_doc):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(1.0)
    
    header = section.header
    banner_table = header.add_table(rows=1, cols=1, width=Inches(6.5))
    banner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_h = banner_table.rows[0].cells[0]
    cell_h.width = Inches(6.5)
    shd = parse_xml(r'<w:shd {} w:fill="619b40"/>'.format(nsdecls('w')))
    cell_h._tc.get_or_add_tcPr().append(shd)

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
        for inspector_nom, group_df in df_filtrado.groupby('inspector', sort=False):
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
            for i, h_text in enumerate(headers):
                hdr_cells[i].text = h_text
                p = hdr_cells[i].paragraphs[0]
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(8.5)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                shd_hdr = parse_xml(r'<w:shd {} w:fill="1E3A8A"/>'.format(nsdecls('w')))
                hdr_cells[i]._tc.get_or_add_tcPr().append(shd_hdr)

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

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# LISTAS Y CONFIGURACIONES
# =========================================================
LISTA_INSPECTORES = ["Juan Navarrete", "Jorge Hernandez", "Arlem Sarmiento", "Harold Castillo", "Miguel Chirinos"]
LISTA_INSPECTORES_DCI = ["Eduardo Delgado", "Felipe Ponce", "José de la Cruz", "Pablo Ruiz", "Luis Durán"]
ESTADOS_STATUS = ["Finalizado", "En curso", "Pendiente"]
TIPOS_INFORME = ["IV", "Ensayo Dureza", "Ensayo LP", "IHV", "Boroscopio"]
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
    "Liberado / Conforme (Aprobado)", "Pendiente de Reparación", "Rechazado", "En Proceso de Inspección", "En Espera de END / Pruebas"
]

semana_actual_num = datetime.now().isocalendar()[1]
idx_semana_defecto = max(0, min(semana_actual_num - 1, len(LISTA_SEMANAS) - 1))

# =========================================================
# BARRA LATERAL Y NAVEGACIÓN
# =========================================================
logo_bytes_sidebar = obtener_bytes_logo()
if logo_bytes_sidebar:
    st.sidebar.image(logo_bytes_sidebar, use_container_width=True)

# Panel de control de acceso de administrador discreto en la barra lateral
if not st.session_state["admin_logueado"]:
    with st.sidebar.expander("🔐 Acceso Administrador"):
        pass_ingresada = st.text_input("Contraseña de Admin:", type="password")
        if st.button("Ingresar como Admin"):
            if pass_ingresada == st.secrets.get("password", "Mechanix123"):
                st.session_state["admin_logueado"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta")
else:
    st.sidebar.success("👑 Modo Administrador Activo")
    if st.sidebar.button("🔓 Cerrar Sesión Admin", use_container_width=True):
        st.session_state["admin_logueado"] = False
        st.rerun()

st.sidebar.markdown("---")

st.markdown("<h1 style='color: #619b40; margin-bottom: 0px;'>Sistema de Gestión de Activos Físicos - QA/QC</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='color: #F97316; margin-top: 5px;'><i>Control Operativo de Inspectores e Histórico de Informes</i></h4>", unsafe_allow_html=True)

franja_bytes = obtener_bytes_franja()
if franja_bytes:
    st.image(franja_bytes, use_container_width=True)
else:
    st.markdown("<hr style='border: 2px solid #619b40;'/>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    ["📝 Registrar Actividad por Inspector", "📊 Historial e Informes", "📈 Reporte Planificación"]
)

personaje_bytes = obtener_bytes_personaje()
if personaje_bytes:
    b64_img = base64.b64encode(personaje_bytes).decode("utf-8")
    st.sidebar.markdown(
        f"""
        <style>
            .personaje-flotante {{
                margin-top: 16cm;
                margin-left: 0.8cm;
                width: 150px;
                display: block;
            }}
        </style>
        <img src="data:image/png;base64,{b64_img}" class="personaje-flotante" />
        """,
        unsafe_allow_html=True
    )

# =========================================================
# MÓDULOS DE LA APLICACIÓN
# =========================================================
if menu == "📝 Registrar Actividad por Inspector":
    st.subheader("📋 Formulario de Ingreso de Actividades")
    
    if not st.session_state["admin_logueado"]:
        st.info("👁️ **Modo Visualización:** Cualquier visitante puede ver esta sección, pero el formulario de ingreso está protegido. Si eres administrador, abre la sección '🔐 Acceso Administrador' en la barra lateral.")
    
    # Si no es admin, los campos se muestran desactivados (solo lectura)
    editable = st.session_state["admin_logueado"]
    
    with st.form("form_actividades_inspector", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            inspector_seleccionado = st.selectbox("👷‍♂️ Seleccionar Inspector asignado:", LISTA_INSPECTORES, disabled=not editable)
            fecha_actividad = st.date_input("📅 Fecha de Inspección:", datetime.now(), disabled=not editable)
            semana_seleccionada = st.selectbox("🗓️ Semana Operativa:", LISTA_SEMANAS, index=idx_semana_defecto, disabled=not editable)
            planta_seleccionada = st.selectbox("🏭 Planta / Unidad:", LISTA_PLANTAS, disabled=not editable)
        with col2:
            tag_equipo = st.text_input("🏷️ TAG del Equipo / Línea Piping:", placeholder="Ej: C-1302 / E-2101 / PIP-001", disabled=not editable)
            porcentaje_avance = st.slider("📊 Porcentaje de Avance:", min_value=0, max_value=100, value=0, step=5, format="%d%%", disabled=not editable)
            estado_liberacion = st.selectbox("📌 Estado de la Inspección:", ESTADOS_LIBERACION, disabled=not editable)

        actividad_realizada = st.text_area("🛠️ Actividades Realizadas por el Inspector:", placeholder="Ej: Inspección visual de junta...", disabled=not editable)
        observaciones = st.text_area("💬 Observaciones Adicionales / Recomendaciones:", placeholder="Escribe comentarios extra...", disabled=not editable)
        
        btn_guardar = st.form_submit_button("☁️ Guardar Registro en Google Sheets", disabled=not editable)
        if btn_guardar:
            if editable:
                if tag_equipo and actividad_realizada:
                    try:
                        sheet = obtener_hoja_actividades()
                        todas_las_filas = sheet.get_all_values()
                        if len(todas_las_filas) == 0:
                            sheet.append_row(["fecha", "semana", "planta", "inspector", "tag_equipo", "actividad_realizada", "avance", "observaciones", "estado_liberacion"])
                        sheet.append_row([str(fecha_actividad), semana_seleccionada, planta_seleccionada, inspector_seleccionado, tag_equipo.strip(), actividad_realizada.strip(), f"{porcentaje_avance}%", observaciones.strip(), estado_liberacion])
                        st.success(f"✅ ¡Actividad de **{inspector_seleccionado}** guardada con éxito!")
                    except Exception as ex:
                        st.error(f"❌ Error al guardar en la nube: {ex}")
                else:
                    st.error("⚠️ Completa los campos obligatorios: **TAG del Equipo** y **Actividades Realizadas**.")

elif menu == "📊 Historial e Informes":
    st.subheader("🔍 Consulta de Historial y Generación de Informes por Día o Semana")
    df_historial = cargar_datos_sheets()
    if not df_historial.empty:
        col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
        with col_f1: filtro_inspector = st.selectbox("Filtrar por Inspector:", ["Todos"] + LISTA_INSPECTORES)
        with col_f2: filtro_semana = st.selectbox("Filtrar por Semana:", ["Todas"] + LISTA_SEMANAS)
        with col_f3: filtro_planta = st.selectbox("Filtrar por Planta:", ["Todas"] + LISTA_PLANTAS)
        with col_f4: filtro_tag = st.text_input("Filtrar por TAG:")
        with col_f5: filtro_estado = st.selectbox("Filtrar por Estado:", ["Todos"] + ESTADOS_LIBERACION)

        df_filtrado = df_historial.copy()
        if filtro_inspector != "Todos": df_filtrado = df_filtrado[df_filtrado['inspector'] == filtro_inspector]
        if filtro_semana != "Todas": df_filtrado = df_filtrado[df_filtrado['semana'] == filtro_semana]
        if filtro_planta != "Todas": df_filtrado = df_filtrado[df_filtrado['planta'] == filtro_planta]
        if filtro_tag: df_filtrado = df_filtrado[df_filtrado['tag_equipo'].astype(str).str.contains(filtro_tag, case=False, na=False)]
        if filtro_estado != "Todos": df_filtrado = df_filtrado[df_filtrado['estado_liberacion'] == filtro_estado]

        st.markdown(f"**Total de registros filtrados:** `{len(df_filtrado)}`")
        st.dataframe(df_filtrado, use_container_width=True)

        st.markdown("---")
        tipo_reporte = st.radio("📌 Frecuencia del Reporte:", ["📅 Diario", "🗓️ Semanal"], horizontal=True)
        if tipo_reporte == "📅 Diario":
            fecha_informe_dt = st.date_input("📅 Selecciona el Día:", datetime.now())
            titulo_doc, subtitulo_doc, tag_archivo = "REPORTE DIARIO DE INSPECCIÓN", f"Fecha: {fecha_informe_dt.strftime('%d/%m/%Y')}", f"DIARIO_{fecha_informe_dt.strftime('%Y%m%d')}"
            df_informe = df_historial[df_historial['fecha'].astype(str) == str(fecha_informe_dt)] if "fecha" in df_historial.columns else pd.DataFrame()
        else:
            semana_informe = st.selectbox("🗓️ Selecciona la Semana:", LISTA_SEMANAS, index=idx_semana_defecto)
            titulo_doc, subtitulo_doc, tag_archivo = "REPORTE SEMANAL DE INSPECCIÓN", f"Período: {semana_informe}", f"SEMANAL_{semana_informe.replace(' ', '_')}"
            df_informe = df_historial[df_historial['semana'] == semana_informe] if "semana" in df_historial.columns else pd.DataFrame()

        col_exp1, col_exp2 = st.columns(2)
        if not df_informe.empty:
            with col_exp1:
                st.download_button(f"📄 Descargar {titulo_doc} (PDF)", generar_pdf_informe(df_informe, titulo_doc, subtitulo_doc), f"INFORME_{tag_archivo}.pdf", "application/pdf", use_container_width=True)
            with col_exp2:
                st.download_button(f"📝 Descargar {titulo_doc} (Word)", generar_word_informe(df_informe, titulo_doc, subtitulo_doc), f"INFORME_{tag_archivo}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        else:
            st.warning("⚠️ No hay registros para este período.")
    else:
        st.info("ℹ️ Aún no hay registros guardados en la nube.")

elif menu == "📈 Reporte Planificación":
    st.subheader("📅 Módulo de Registro y Control de Planificación")
    
    editable_plan = st.session_state["admin_logueado"]
    if not editable_plan:
        st.info("👁️ **Modo Visualización:** Puedes ver y descargar los datos de planificación, pero las opciones de edición están restringidas al administrador.")

    with st.form("form_planificacion", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            p_planta = st.selectbox("🏭 Planta:", LISTA_PLANTAS, disabled=not editable_plan)
            p_tipo_informe = st.selectbox("📋 Tipo de Informe:", TIPOS_INFORME, disabled=not editable_plan)
            p_circuito = st.text_input("🔧 Circuito / Equipo:", disabled=not editable_plan)
            p_f_inicio = st.date_input("📅 Fecha Inicio:", datetime.now(), disabled=not editable_plan)
            p_f_fin = st.date_input("📅 Fecha Fin:", datetime.now(), disabled=not editable_plan)
            p_cont_insp = st.number_input("🔢 Contribución Inspección:", min_value=0, value=0, disabled=not editable_plan)
            p_programa = st.text_input("📌 Programa:", disabled=not editable_plan)
        with col2:
            p_insp_dci = st.selectbox("👷‍♂️ Inspector DCI:", LISTA_INSPECTORES_DCI, disabled=not editable_plan)
            p_insp_ingemars = st.selectbox("👷‍♂️ Inspector Ingemars:", ["Todos"] + LISTA_INSPECTORES, disabled=not editable_plan)
            p_dias_enap = st.number_input("⏱️ Días ENAP:", min_value=0, value=0, disabled=not editable_plan)
            p_otep = st.text_input("📄 N° OTEP:", disabled=not editable_plan)
            p_informe_iv = st.text_input("📑 N° Informe IV:", disabled=not editable_plan)
            p_status = st.selectbox("📌 Status:", ESTADOS_STATUS, disabled=not editable_plan)
            p_f_entrega_ope = st.date_input("📅 Fecha entrega Ope.:", datetime.now(), disabled=not editable_plan)
        with col3:
            p_cont_lib = st.number_input("🔢 Contador liberación:", min_value=0, value=0, disabled=not editable_plan)
            p_lib_final = st.text_input("✅ Liberación final:", disabled=not editable_plan)
            p_cont_enap = st.number_input("🔢 Contador ENAP:", min_value=0, value=0, disabled=not editable_plan)
            p_cont_cump_enap = st.number_input("📊 Cumplimiento ENAP:", min_value=0, value=0, disabled=not editable_plan)
            p_cont_dias_inf = st.number_input("📉 Días informe:", min_value=0, value=0, disabled=not editable_plan)
            p_archivo_url = st.text_input("🔗 URL Archivo:", disabled=not editable_plan)
        p_observaciones = st.text_area("💬 Observaciones:", disabled=not editable_plan)
        
        if st.form_submit_button("☁️ Guardar Planificación", disabled=not editable_plan):
            if editable_plan:
                try:
                    ws_plan = obtener_hoja_planificacion()
                    ws_plan.append_row([p_planta, p_tipo_informe, p_circuito.strip(), str(p_f_inicio), str(p_f_fin), str(p_cont_insp), p_programa.strip(), p_insp_dci, p_insp_ingemars, str(p_dias_enap), p_otep.strip(), p_informe_iv.strip(), p_status, str(p_f_entrega_ope), str(p_cont_lib), p_lib_final.strip(), str(p_cont_enap), str(p_cont_cump_enap), str(p_cont_dias_inf), p_archivo_url.strip(), p_observaciones.strip()])
                    st.success("✅ ¡Planificación guardada con éxito!")
                except Exception as ex:
                    st.error(f"❌ Error: {ex}")

    df_plan = cargar_datos_planificacion()
    if not df_plan.empty:
        st.dataframe(df_plan, use_container_width=True)

# =========================================================
# PIE DE PÁGINA
# =========================================================
st.markdown("<br/><br/>", unsafe_allow_html=True)
col_foot1, col_foot2, col_foot3 = st.columns([2, 1, 2])
with col_foot2:
    logo_footer_bytes = obtener_bytes_logo()
    if logo_footer_bytes:
        st.image(logo_footer_bytes, width=150)

franja_footer_bytes = obtener_bytes_franja()
if franja_footer_bytes:
    st.image(franja_footer_bytes, use_container_width=True)
else:
    st.markdown("<hr style='border: 2px solid #619b40;'/>", unsafe_allow_html=True)
