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

# 🔗 URLs RAW DE LOGO, FRANJA Y PERSONAJE EN GITHUB
URL_LOGO_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/logo.png"
URL_FRANJA_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/franja.png"
URL_PERSONAJE_GITHUB = "https://raw.githubusercontent.com/death-87/app-inspecciones/main/personaje.png"

# 🆔 ID DE TU HOJA DE GOOGLE SHEETS
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"

# =========================================================
# GESTIÓN DE AUTENTICACIÓN Y ROLES DE USUARIO
# =========================================================
USUARIOS_SISTEMA = {
    "invitado": {"password": "123", "rol": "invitado", "nombre": "Visitante / Solo Lectura"},
    "jnavarrete": {"password": "Mechanix123", "rol": "operador", "nombre": "jnavarrete (Agregar Datos)"},
    "jhernandez": {"password": "jorge2026", "rol": "operador", "nombre": "jhernandez (Agregar Datos)"},
    "admin": {"password": "Mechanix123", "rol": "admin", "nombre": "Administrador General"}
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None
if "rol_actual" not in st.session_state:
    st.session_state.rol_actual = "invitado"
if "nombre_usuario" not in st.session_state:
    st.session_state.nombre_usuario = "Visitante"

# =========================================================
# FUNCIONES PARA DESCARGA Y CACHÉ DE IMÁGENES EN MEMORIA
# =========================================================
@st.cache_data(ttl=3600)
def obtener_bytes_imagen(url):
    """Descarga una imagen de internet una sola vez y la mantiene en caché."""
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

    # Corregir saltos de línea de la clave privada
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    credentials = Credentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )

    client = gspread.authorize(credentials)

    return client.open_by_key(SPREADSHEET_ID)

def obtener_hoja_actividades():
    client = conectar_google_sheets()
    return client.sheet1


def plantas_por_tag(filas):
    """BASE EQUIPOS: columna C = planta, columna D = TAG."""
    resultado = {}
    for fila in filas[1:]:
        if len(fila) > 3 and str(fila[3]).strip():
            tag = str(fila[3]).strip().upper()
            planta = str(fila[2]).strip()
            if planta:
                resultado.setdefault(tag, set()).add(planta)
    return {tag: sorted(plantas) for tag, plantas in resultado.items()}

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

def preparar_indicadores_actividad(datos):
    requeridas = ['fecha', 'planta', 'inspector', 'tag_equipo', 'avance', 'estado_liberacion']
    if not set(requeridas).issubset(datos.columns):
        raise ValueError('Faltan columnas del historial de actividades.')
    df = datos.copy()
    for campo in requeridas:
        df[campo] = df[campo].fillna('').astype(str).str.strip()
    df['tag_equipo'] = df['tag_equipo'].str.upper()
    df = df[df['tag_equipo'].ne('')].copy()
    df['_planta'] = df['planta'].str.upper()
    df['Fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
    iso = df['Fecha'].dt.isocalendar()
    df['Año ISO'] = iso.year
    df['Semana ISO'] = iso.week
    df['Avance (%)'] = pd.to_numeric(df['avance'].str.replace('%', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
    df.loc[~df['Avance (%)'].between(0, 100), 'Avance (%)'] = float('nan')
    estados = {'pendiente de inspección': 'Proceso de inspección',
        'pendiente de informe': 'Proceso de informe',
        'proceso de inspección': 'Proceso de inspección',
               'en proceso de inspección': 'Proceso de inspección',
               'proceso de informe': 'Proceso de informe', 'finalizada': 'Finalizada'}
    df['Situación'] = df['estado_liberacion'].str.casefold().map(estados).fillna('Estado por revisar')
    # El historial se agrega por filas: coincide con la función de continuar un TAG.
    actuales = df.drop_duplicates(['_planta', 'tag_equipo'], keep='last').copy()
    return df, actuales


def mostrar_indicadores_actividad():
    st.subheader('📈 Indicadores de actividades')
    st.caption('Datos del mismo historial que Registrar actividad e Historial e Informes.')
    st.button('Actualizar indicadores', key='principal_ind_actualizar')
    datos = cargar_datos_sheets()
    if datos.empty:
        st.info('Todavía no hay actividades disponibles para calcular indicadores.')
        return
    try:
        df, actuales = preparar_indicadores_actividad(datos)
    except ValueError as exc:
        st.error(str(exc))
        return
    a, b = st.columns(2)
    plantas = a.multiselect('Plantas / unidades', sorted(df['planta'].unique()), default=sorted(df['planta'].unique()), key='principal_ind_plantas')
    inspectores = b.multiselect('Inspectores', sorted(df['inspector'].unique()), default=sorted(df['inspector'].unique()), key='principal_ind_inspectores')
    # Resolver el último estado ANTES de filtrar por inspector: evita revivir pendientes antiguos.
    actuales = actuales[actuales['planta'].isin(plantas) & actuales['inspector'].isin(inspectores)].copy()
    actividad = df[df['planta'].isin(plantas) & df['inspector'].isin(inspectores)].copy()
    st.markdown('### Pendientes actuales')
    st.caption('Acumulados de todas las semanas. El inspector corresponde al último registro de cada equipo; no necesariamente a una asignación formal.')
    cols = st.columns(4)
    for col, estado in zip(cols, ['Proceso de inspección', 'Proceso de informe', 'Finalizada', 'Estado por revisar']):
        col.metric(estado, int(actuales['Situación'].eq(estado).sum()))
    pendientes = actuales[actuales['Situación'].isin(['Proceso de inspección', 'Proceso de informe'])]
    if not pendientes.empty:
        resumen = pendientes.groupby(['inspector', 'Situación']).size().reset_index(name='Equipos')
        fig = px.bar(resumen, x='Equipos', y='inspector', color='Situación', orientation='h', barmode='stack',
                     color_discrete_map={'Proceso de inspección': '#355C83', 'Proceso de informe': '#C48A25'}, text='Equipos')
        fig.update_layout(template='plotly_white', height=330, margin=dict(t=10,b=10), yaxis_title='Inspector del último registro')
        st.plotly_chart(fig, use_container_width=True)
    with st.expander('Detalle de pendientes y estados por revisar', expanded=True):
        detalle = actuales[actuales['Situación'].ne('Finalizada')].copy()
        detalle['Días desde última actividad'] = (pd.Timestamp(datetime.now().date()) - detalle['Fecha'].dt.normalize()).dt.days
        columnas = ['planta', 'tag_equipo', 'inspector', 'fecha', 'Situación', 'estado_liberacion', 'Avance (%)', 'Días desde última actividad']
        if detalle.empty:
            st.info('Sin pendientes ni estados por revisar para esta selección.')
        else:
            st.dataframe(detalle[columnas], hide_index=True, use_container_width=True)
        st.caption('Los días indican antigüedad de la última actividad, no atraso contractual. Los estados antiguos sin equivalencia segura requieren revisión.')
    st.divider()
    st.markdown('### Actividad por período')
    a,b = st.columns(2)
    anios = sorted(set(df['Año ISO'].dropna().astype(int)) | {datetime.now().date().isocalendar().year}, reverse=True)
    anio = a.selectbox('Año ISO', anios, key='principal_ind_anio')
    semanas = list(range(1, datetime(int(anio),12,28).date().isocalendar().week + 1))
    elegidas = b.multiselect('Semanas ISO', semanas, default=semanas, key=f'principal_ind_semanas_{anio}')
    periodo = actividad[(actividad['Año ISO'] == anio) & actividad['Semana ISO'].isin(elegidas)].copy()
    if actividad['Fecha'].isna().any():
        st.warning(f'{int(actividad["Fecha"].isna().sum())} registros tienen fecha inválida y no se incluyen en los gráficos semanales.')
    if periodo.empty:
        st.info('No hay actividades en las semanas seleccionadas. Los pendientes de arriba siguen mostrando el acumulado actual.')
        return
    ultimos_periodo = periodo.drop_duplicates(['_planta', 'tag_equipo'], keep='last')
    cols = st.columns(3)
    cols[0].metric('Actividades registradas', len(periodo))
    cols[1].metric('Equipos atendidos', len(ultimos_periodo))
    promedio = ultimos_periodo['Avance (%)'].mean()
    cols[2].metric('Avance medio registrado', 'Sin datos' if pd.isna(promedio) else f'{promedio:.1f}%')
    st.caption('Avance: último valor por equipo dentro del período seleccionado, sin sumar porcentajes de jornadas. Actividades y horas no equivalen a productividad.')
    a,b = st.columns(2)
    semanal = periodo.groupby('Semana ISO').size().reindex(sorted(elegidas), fill_value=0).rename('Actividades').reset_index()
    por_inspector = periodo.groupby('inspector').size().rename('Actividades').reset_index()
    for columna, tabla, eje, titulo in [(a,semanal,'Semana ISO','Actividad semanal'), (b,por_inspector,'inspector','Actividad por inspector')]:
        fig = px.bar(tabla, x=eje, y='Actividades', text='Actividades', title=titulo, color_discrete_sequence=['#087F73'])
        fig.update_layout(template='plotly_white', height=330, margin=dict(t=45,b=10))
        fig.update_yaxes(dtick=1)
        columna.plotly_chart(fig, use_container_width=True)
    st.caption('Cada actividad cuenta como un registro. No se calculan horas porque este historial no las registra. Equipos sin actividad y trabajos simultáneos del mismo TAG requieren una base de planificación o un identificador de trabajo para un backlog completo.')


def obtener_ultimo_registro_tag(tag_busqueda):
    """Busca el registro más reciente de un TAG en el historial de actividades."""
    df_historial = cargar_datos_sheets()
    if df_historial.empty or "tag_equipo" not in df_historial.columns:
        return None
    
    # Filtrar coincidentes sin importar mayúsculas/minúsculas
    df_tag = df_historial[df_historial['tag_equipo'].astype(str).str.strip().str.upper() == tag_busqueda.strip().upper()]
    
    if not df_tag.empty:
        # Retorna el último registro ingresado (última fila)
        return df_tag.iloc[-1].to_dict()
    return None

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
    "Proceso de inspección",
    "Proceso de informe",
    "Finalizada"
]

# Cálculo de semana por defecto global
semana_actual_num = datetime.now().isocalendar()[1]
idx_semana_defecto = max(0, min(semana_actual_num - 1, len(LISTA_SEMANAS) - 1))

# =========================================================
# BARRA LATERAL (LOGIN Y CONTROL DE ACCESO)
# =========================================================
logo_bytes_sidebar = obtener_bytes_logo()
if logo_bytes_sidebar:
    st.sidebar.image(logo_bytes_sidebar, use_container_width=True)

st.sidebar.markdown("### 🔐 Control de Acceso")

if not st.session_state.autenticado:
    with st.sidebar.form("form_login"):
        user_input = st.text_input("Usuario:")
        pass_input = st.text_input("Contraseña:", type="password")
        btn_login = st.form_submit_button("🔑 Iniciar Sesión")
        
        if btn_login:
            if user_input in USUARIOS_SISTEMA and USUARIOS_SISTEMA[user_input]["password"] == pass_input:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = user_input
                st.session_state.rol_actual = USUARIOS_SISTEMA[user_input]["rol"]
                st.session_state.nombre_usuario = USUARIOS_SISTEMA[user_input]["nombre"]
                st.rerun()
            else:
                st.sidebar.error("❌ Usuario o contraseña incorrectos")
    
    st.sidebar.info("ℹ️ Entrando como **Visitante** por defecto (Solo Lectura).")
else:
    st.sidebar.success(f"👤 Conectado:\n**{st.session_state.nombre_usuario}**")
    if st.sidebar.button("🚪 Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.usuario_actual = None
        st.session_state.rol_actual = "invitado"
        st.session_state.nombre_usuario = "Visitante"
        st.rerun()

st.sidebar.markdown("---")

# =========================================================
# ENCABEZADO PRINCIPAL CON FRANJA Y ESTILOS
# =========================================================
st.markdown("<h1 style='color: #619b40; margin-bottom: 0px;'>Sistema de Gestión de Activos Físicos - QA/QC</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='color: #F97316; margin-top: 5px;'><i>Control Operativo de Inspectores e Histórico de Informes</i></h4>", unsafe_allow_html=True)

# 📸 Carga segura de la Franja Decorativa Superior
franja_bytes = obtener_bytes_franja()
if franja_bytes:
    st.image(franja_bytes, use_container_width=True)
else:
    st.markdown("<hr style='border: 2px solid #619b40;'/>", unsafe_allow_html=True)

# =========================================================
# MENÚ Y NAVEGACIÓN (SIDEBAR)
# =========================================================
menu = st.sidebar.radio(
    "📌 Selecciona una Opción:",
    [
        "📝 Registrar Actividad por Inspector", 
        "📊 Historial e Informes", 
        "📈 Indicadores de actividades",
        "📈 Reporte Planificación"
    ]
)

# 🟢 PERSONAJE EN BASE64 CON POSICIONAMIENTO EN CENTÍMETROS
personaje_bytes = obtener_bytes_personaje()
if personaje_bytes:
    b64_img = base64.b64encode(personaje_bytes).decode("utf-8")
    st.sidebar.markdown(
        f"""
        <style>
            .personaje-flotante {{
                margin-top: 14cm;   /* ⬇️ AJUSTE VERTICAL */
                margin-left: 0.8cm; /* ➡️ AJUSTE HORIZONTAL */
                width: 150px;       /* 📐 Ancho fijado en 150px */
                display: block;
            }}
        </style>
        <img src="data:image/png;base64,{b64_img}" class="personaje-flotante" />
        """,
        unsafe_allow_html=True
    )

rol_usuario = st.session_state.rol_actual

# =========================================================
# MÓDULO 1: REGISTRO DE ACTIVIDADES (ESCRITURA CON REUTILIZACIÓN)
# =========================================================
if menu == "📝 Registrar Actividad por Inspector":
    st.subheader("📋 Formulario de Ingreso de Actividades")
    
    if rol_usuario == "invitado":
        st.warning("⚠️ Tu cuenta actual es de **Visitante (Solo Lectura)**. No tienes permisos para registrar actividades. Por favor inicia sesión con un usuario autorizado en la barra lateral.")
    
    # 🔄 SECCIÓN PARA CONTINUAR TRABAJOS PENDIENTES
    st.markdown("##### 🔄 CONTINUAR UNA ACTIVIDAD PENDIENTE")
    col_retoma1, col_retoma2 = st.columns([3, 1])
    
    with col_retoma1:
        tag_para_retomar = st.text_input("Buscar TAG para continuar trabajo pendiente:", placeholder="Ej: C-1302")
    
    with col_retoma2:
        st.markdown("<br/>", unsafe_allow_html=True)
        btn_cargar_tag = st.button("🔎 Cargar Datos Últimos")
    
    # Valores por defecto para el formulario
    def_inspector = LISTA_INSPECTORES[0]
    def_planta = LISTA_PLANTAS[0]
    def_tag = ""
    def_avance = 0
    def_actividad = ""
    def_obs = ""
    def_estado = ESTADOS_LIBERACION[0]

    if btn_cargar_tag and tag_para_retomar:
        ultimo_reg = obtener_ultimo_registro_tag(tag_para_retomar)
        if ultimo_reg:
            # Obtener el entero del porcentaje guardado (ej: "45%" -> 45)
            avance_str = str(ultimo_reg.get("avance", "0")).replace("%", "").strip()
            avance_val = int(float(avance_str)) if avance_str.replace('.', '', 1).isdigit() else 0
            
            if avance_val >= 100:
                st.info(f"ℹ️ El TAG **{tag_para_retomar}** ya figura completado al 100%. Se cargarán sus datos base como referencia.")
            else:
                st.success(f"✅ Último registro cargado para **{tag_para_retomar}** (Avance previo: {avance_val}%). Puedes actualizar el progreso y guardar la nueva jornada.")

            def_tag = str(ultimo_reg.get("tag_equipo", tag_para_retomar))
            def_actividad = str(ultimo_reg.get("actividad_realizada", ""))
            def_obs = f"Continuación de inspección anterior. Obs previas: {ultimo_reg.get('observaciones', '')}"
            def_avance = avance_val
            
            if ultimo_reg.get("inspector") in LISTA_INSPECTORES:
                def_inspector = ultimo_reg.get("inspector")
            if ultimo_reg.get("planta") in LISTA_PLANTAS:
                def_planta = ultimo_reg.get("planta")
            if ultimo_reg.get("estado_liberacion") in ESTADOS_LIBERACION:
                def_estado = ultimo_reg.get("estado_liberacion")
            elif ultimo_reg.get("estado_liberacion"):
                st.info(f"Estado anterior: {ultimo_reg.get('estado_liberacion')}. Selecciona el estado actual de esta jornada; el registro anterior se conserva.")
            st.info(f"Inspector del registro anterior: {ultimo_reg.get('inspector', '-')}. Si continúa otra persona, selecciónala como inspector asignado antes de guardar.")
        else:
            st.warning(f"⚠️ No se encontraron registros anteriores para el TAG '{tag_para_retomar}'. Se creará uno desde cero.")
            def_tag = tag_para_retomar

    st.markdown("---")

    # 📝 FORMULARIO DE INGRESO
    valores_actividad = {
        'act_inspector': def_inspector, 'act_planta': def_planta,
        'act_tag': def_tag, 'act_avance': max(0, min(100, def_avance)),
        'act_estado': def_estado, 'act_actividad': def_actividad,
        'act_observaciones': def_obs,
    }
    for clave, valor in valores_actividad.items():
        if (btn_cargar_tag and tag_para_retomar) or clave not in st.session_state:
            st.session_state[clave] = valor
    if st.session_state['act_estado'] not in ESTADOS_LIBERACION:
        st.session_state['act_estado'] = ESTADOS_LIBERACION[0]

    if 'act_base_equipos' not in st.session_state:
        try:
            filas_base = conectar_google_sheets().worksheet('BASE EQUIPOS').get_all_values()
            st.session_state['act_base_equipos'] = plantas_por_tag(filas_base)
        except Exception:
            st.warning('No se pudo leer BASE EQUIPOS. Puedes ingresar el TAG y seleccionar la planta manualmente.')
    if st.button('Actualizar lista de equipos', key='act_actualizar_base'):
        st.session_state.pop('act_base_equipos', None)
        st.rerun()
    base_equipos = st.session_state.get('act_base_equipos', {})
    tag_actual = st.session_state['act_tag'].strip().upper()
    opciones_tag = [''] + sorted(set(base_equipos) | ({tag_actual} if tag_actual else set()))
    st.session_state['act_tag'] = tag_actual
    if st.checkbox('Ingresar TAG manualmente', key='act_tag_manual'):
        tag_equipo = st.text_input('🏷️ TAG del Equipo / Línea Piping:', key='act_tag')
    else:
        tag_equipo = st.selectbox('🏷️ TAG del Equipo / Línea Piping:', opciones_tag, key='act_tag')
    tag_normalizado = tag_equipo.strip().upper()
    plantas_tag = base_equipos.get(tag_normalizado, [])
    if len(plantas_tag) == 1:
        st.session_state['act_planta'] = plantas_tag[0]
    elif len(plantas_tag) > 1:
        st.warning('Este TAG figura en varias plantas. Selecciona la que corresponde a esta actividad.')
    elif tag_normalizado:
        st.caption('TAG sin planta asociada en BASE EQUIPOS. Verifica la planta manualmente.')
    opciones_planta = list(dict.fromkeys(LISTA_PLANTAS + plantas_tag + [st.session_state['act_planta']]))

    # Widgets fuera de st.form: el TAG actualiza la planta inmediatamente.
    with st.container():
        col1, col2 = st.columns(2)
        
        with col1:
            idx_insp = LISTA_INSPECTORES.index(def_inspector) if def_inspector in LISTA_INSPECTORES else 0
            idx_plan = LISTA_PLANTAS.index(def_planta) if def_planta in LISTA_PLANTAS else 0
            
            inspector_seleccionado = st.selectbox("👷‍♂️ Seleccionar Inspector asignado:", LISTA_INSPECTORES, key='act_inspector')
            fecha_actividad = st.date_input("📅 Fecha de Inspección:", datetime.now().date(), key='act_fecha')
            semana_seleccionada = st.selectbox("🗓️ Semana Operativa:", LISTA_SEMANAS, index=idx_semana_defecto, key='act_semana')
            planta_seleccionada = st.selectbox("🏭 Planta / Unidad:", opciones_planta, key='act_planta', disabled=len(plantas_tag) == 1)

        with col2:
            porcentaje_avance = st.slider("📊 Porcentaje de Avance Acumulado:", min_value=0, max_value=100, key='act_avance', step=5, format="%d%%")
            
            idx_est = ESTADOS_LIBERACION.index(def_estado) if def_estado in ESTADOS_LIBERACION else 0
            estado_liberacion = st.selectbox("📌 Estado de la Inspección:", ESTADOS_LIBERACION, key='act_estado')

        actividad_realizada = st.text_area("🛠️ Actividades Realizadas en esta jornada:", key='act_actividad', placeholder="Ej: Inspección visual de junta...")
        observaciones = st.text_area("💬 Observaciones Adicionales / Recomendaciones:", key='act_observaciones', placeholder="Escribe comentarios extra...")
        
        btn_guardar = st.button("☁️ Guardar Nuevo Registro en Google Sheets", key='act_guardar')
        
        if btn_guardar:
            if rol_usuario == "invitado":
                st.error("❌ Acción no permitida para el rol de Visitante.")
            elif tag_equipo.strip() and actividad_realizada.strip():
                try:
                    sheet = obtener_hoja_actividades()
                    todas_las_filas = sheet.get_all_values()
                    if len(todas_las_filas) == 0:
                        sheet.append_row([
                            "fecha", "semana", "planta", "inspector", "tag_equipo", 
                            "actividad_realizada", "avance", "observaciones", "estado_liberacion"
                        ])
                    
                    # 💡 SE CREA UNA NUEVA FILA (INDEPENDIENTE Y AUDITABLE)
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
                    if any([str(v) for v in fila[:9]] == [str(v) for v in nueva_fila] for fila in todas_las_filas[1:]):
                        st.info("Este registro exacto ya está guardado. Modifica los datos para registrar una nueva jornada.")
                    else:
                        sheet.append_row(nueva_fila, value_input_option="RAW")
                        st.success(f"✅ ¡Nuevo registro de **{inspector_seleccionado}** (TAG: {tag_equipo} | Avance: {porcentaje_avance}%) guardado con éxito!")
                except Exception as ex:
                    st.error("No se pudo confirmar el guardado. Los campos se conservan en esta sesión. Reintenta para comprobar si el registro ya existe.")
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
            estados_anteriores = sorted(set(df_historial['estado_liberacion'].dropna().astype(str)) - set(ESTADOS_LIBERACION) - {''}) if 'estado_liberacion' in df_historial.columns else []
            filtro_estado = st.selectbox("Filtrar por Estado:", ["Todos"] + ESTADOS_LIBERACION + estados_anteriores)

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

        # 📊 GRÁFICO AUTOMÁTICO CON ESCALA GRADUAL DE COLOR
        if tipo_reporte == "🗓️ Semanal" and not df_informe.empty:
            st.markdown("---")
            st.markdown(f"#### 📊 Porcentaje de Avance Promedio por Inspector ({semana_informe})")
            
            df_grafico = df_informe.copy()
            if 'avance' in df_grafico.columns and 'inspector' in df_grafico.columns:
                df_grafico['avance_num'] = pd.to_numeric(
                    df_grafico['avance'].astype(str).str.replace('%', '').str.strip(),
                    errors='coerce'
                ).fillna(0)
                
                df_resumen_avance = df_grafico.groupby('inspector')['avance_num'].mean().reset_index()
                df_resumen_avance.columns = ['Inspector', 'Avance Promedio (%)']
                
                escala_colores_custom = [
                    [0.0, "#9CA3AF"],
                    [0.2, "#64748B"],
                    [0.4, "#2563EB"],
                    [0.6, "#0D9488"],
                    [0.8, "#16A34A"],
                    [1.0, "#619b40"]
                ]

                fig = px.bar(
                    df_resumen_avance,
                    x='Inspector',
                    y='Avance Promedio (%)',
                    color='Avance Promedio (%)',
                    color_continuous_scale=escala_colores_custom,
                    range_color=[0, 100],
                    text_auto='.1f',
                    title=f"Avance Semanal - {semana_informe}"
                )

                fig.update_layout(
                    yaxis=dict(range=[0, 105], title="Porcentaje (%)"),
                    xaxis_title="Inspector",
                    coloraxis_showscale=True,
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)"
                )
                
                fig.update_traces(texttemplate='%{y:.1f}%', textposition='outside')
                st.plotly_chart(fig, use_container_width=True)

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
elif menu == "📈 Indicadores de actividades":
    mostrar_indicadores_actividad()

elif menu == "📈 Reporte Planificación":
    st.subheader("📅 Módulo de Registro y Control de Planificación")
    st.markdown("##### *Ingrese los datos detallados de planificación para sincronizar con la nube.*")
    
    if rol_usuario == "invitado":
        st.warning("⚠️ Tu cuenta actual es de **Visitante (Solo Lectura)**. Puedes visualizar la planificación pero no registrar nuevos datos.")

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
            if rol_usuario == "invitado":
                st.error("❌ Acción no permitida para el rol de Visitante.")
            else:
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

# =========================================================
# PIE DE PÁGINA (LOGO CENTRADO SIN FRANJA INFERIOR)
# =========================================================
st.markdown("<br/><br/>", unsafe_allow_html=True)

col_foot1, col_foot2, col_foot3 = st.columns([2, 1, 2])
with col_foot2:
    logo_footer_bytes = obtener_bytes_logo()
    if logo_footer_bytes:
        st.image(logo_footer_bytes, width=150)
