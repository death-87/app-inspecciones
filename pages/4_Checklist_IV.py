"""Página independiente. Instalar en pages/ junto a app.py.
Dependencias: streamlit, pandas, reportlab, Pillow, gspread, google-auth, google-api-python-client.
El respaldo JSON incluye los campos, respuestas y fotografías.
"""
import base64
import hashlib
import io
import json
import re
from datetime import date
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as PDFImage

PREFIJO = "checklist_iv_"
VERSION = 1
INSPECTORES = ['Juan Navarrete', 'Jorge Hernández', 'Harold Castillo', 'Miguel Chirinos', 'Arlem Sarmiento']
SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"
GRUPOS = {
    "Inspección externa general": [
        "Presenta placa de identificación", "Acceso al sustrato metálico, estado",
        "Deformación en virola", "Corrosión externa", "Pitting", "C.U.I.",
        "Abolladuras / Ampollamientos", "Grietas", "Fugas / Filtraciones",
        "Decoloración", "Aislamiento térmico", "Recubrimiento por pintura"],
    "Conexiones / Boquillas": ["Bridas", "Cuello boquillas", "Uniones soldadas",
        "Abrazadera o zapata deformada, fracturada", "Pernos / espárragos", "Conexiones roscadas", "Instrumentos"],
    "Soportación": ["Faldón", "Sillín", "Pernos de anclaje", "Soldadura de soportes",
        "Fundación / placa base", "Estructura asociada"],
    "Accesorios": ["Dispositivos de alivio", "Conexión a tierra eléctrica", "Instrumentación"],
}
CAMPOS = {
    "numero": "Informe N.º", "contrato": "Contrato", "cliente": "Cliente", "ot": "OT / PPTO",
    "procedimiento": "Procedimiento", "planta": "Planta", "equipo": "Equipo / circuito",
    "codigo": "Código de evaluación / inspección", "ensayo": "Ensayo complementario",
    "complementario": "N.º informe complementario", "solicitante": "Solicitante",
    "inspector": "Inspector visual", "ingeniero": "Ingeniero de operaciones",
    "descripcion": "Descripción del equipo", "aca": "ACA",
}


def credenciales_google():
    from google.oauth2.service_account import Credentials
    config = dict(st.secrets['connections']['gsheets'])
    config['private_key'] = config['private_key'].replace('\\n', '\n')
    return Credentials.from_service_account_info(config, scopes=[
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/drive.readonly'])


def leer_base_equipos():
    import gspread
    filas = gspread.authorize(credenciales_google()).open_by_key(SPREADSHEET_ID).worksheet('BASE EQUIPOS').get_all_values()
    # Misma distribución de columnas que la página de informe visual.
    return [{'planta': f[2].strip(), 'equipo': f[3].strip(),
             'descripcion': f[4].strip() if len(f)>4 else '', 'aca': f[7].strip() if len(f)>7 else ''}
            for f in filas[1:] if len(f)>3 and f[3].strip()]


def aplicar_equipo_formulario(estado, campos, equipo, revision):
    for campo in ('planta', 'equipo', 'descripcion', 'aca'):
        valor = str(equipo.get(campo, '') or '')
        campos[campo] = valor
        estado[f'{PREFIJO}{revision}_{campo}'] = valor
    estado.pop(PREFIJO+'pdf', None)


def preparar_imagen(raw, nombre):
    if len(raw)>10*1024*1024:
        raise ValueError(f'{nombre}: máximo 10 MB por imagen.')
    with Image.open(io.BytesIO(raw)) as im:
        im = ImageOps.exif_transpose(im).convert('RGB')
        im.thumbnail((1800,1800))
        salida = io.BytesIO()
        im.save(salida, format='JPEG', quality=88)
    return {'id': hashlib.sha256(raw).hexdigest(),
            'datos': base64.b64encode(salida.getvalue()).decode(),
            'leyenda': re.sub(r'^\d+[_ -]*', '', nombre.rsplit('.',1)[0]).replace('_',' ')}


def leer_carpeta_drive(enlace):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    match = re.search(r'folders/([A-Za-z0-9_-]+)', enlace)
    folder = match.group(1) if match else enlace.strip()
    if not re.fullmatch(r'[A-Za-z0-9_-]+', folder):
        raise ValueError('Introduce un enlace de carpeta de Drive o su ID.')
    servicio = build('drive','v3',credentials=credenciales_google())
    archivos, token = [], None
    while True:
        resultado = servicio.files().list(q=f"'{folder}' in parents and trashed = false and (mimeType contains 'image/' or mimeType = 'application/octet-stream')",
            fields='nextPageToken,files(id,name,size)', pageSize=100, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        archivos.extend(resultado.get('files',[]))
        token = resultado.get('nextPageToken')
        if not token:
            break
    def orden(f):
        n = re.match(r'^(\d+)',f['name'])
        return (int(n.group(1)) if n else 999999, f['name'].lower())
    fotos, esquemas, avisos = [], [], []
    for f in sorted(archivos,key=orden):
        destino = esquemas if 'esquema' in f['name'].lower() else fotos
        if len(destino)>=20:
            avisos.append('Se omitió '+f['name']+': máximo 20 por categoría.')
            continue
        try:
            if int(f.get('size',0))>10*1024*1024:
                raise ValueError('supera 10 MB')
            buffer = io.BytesIO()
            descarga = MediaIoBaseDownload(buffer,servicio.files().get_media(fileId=f['id'],supportsAllDrives=True))
            fin=False
            while not fin:
                _,fin=descarga.next_chunk()
                if buffer.tell()>10*1024*1024:
                    raise ValueError('supera 10 MB')
            destino.append(preparar_imagen(buffer.getvalue(),f['name']))
        except Exception:
            avisos.append('No se pudo cargar '+f['name']+'. Verifica formato, tamaño y permisos.')
    return fotos,esquemas,avisos


def nuevo():
    campos = {k: "" for k in CAMPOS}
    campos.update(contrato="AC31104857", cliente="Enap Refinerías Aconcagua S.A.",
                  codigo="API 510 / 572 / ASME VIII", ingeniero="Gabriel Allendes V.", fecha_inspeccion=date.today().isoformat(),
                  fecha_emision=date.today().isoformat(), alcance="", conclusiones="", recomendaciones="")
    return {"version": VERSION, "tipo": "checklist_iv", "campos": campos,
            "grupos": {g: {"observaciones": "", "filas": [{"Punto": p, "Respuesta": "Sin evaluar", "Evaluación": "Sin evaluar", "Observaciones": ""} for p in puntos]} for g, puntos in GRUPOS.items()},
            "fotos": [], "esquemas": []}


def validar(datos):
    if not isinstance(datos, dict) or datos.get("version") != VERSION or datos.get("tipo") != "checklist_iv":
        raise ValueError("El archivo no corresponde a esta versión del checklist.")
    base = nuevo()
    datos['campos'].setdefault('descripcion', '')
    datos['campos'].setdefault('aca', '')
    datos.setdefault('esquemas', [])
    if set(datos.get("campos", {})) != set(base["campos"]):
        raise ValueError("El respaldo tiene campos incompletos.")
    if not all(isinstance(v, str) and len(v) <= 20000 for v in datos["campos"].values()):
        raise ValueError("Hay campos no válidos o demasiado largos.")
    for campo in ("fecha_inspeccion", "fecha_emision"):
        date.fromisoformat(datos["campos"][campo])
    if set(datos.get("grupos", {})) != set(GRUPOS):
        raise ValueError("Los grupos del checklist no coinciden.")
    for grupo, puntos in GRUPOS.items():
        contenido = datos["grupos"][grupo]
        if not isinstance(contenido.get("observaciones"), str):
            raise ValueError("Observaciones de grupo no válidas.")
        filas = contenido["filas"]
        if [r.get("Punto") for r in filas] != puntos:
            raise ValueError("Los puntos del checklist no coinciden.")
        for fila in filas:
            if fila.get("Respuesta") not in ("Sin evaluar", "Sí", "No", "N/A") or fila.get("Evaluación") not in ("Sin evaluar", "C", "NC", "N/A"):
                raise ValueError("Respuesta de checklist no válida.")
            if not isinstance(fila.get("Observaciones"), str):
                raise ValueError("Observación no válida.")
    if not isinstance(datos.get("fotos"), list) or len(datos["fotos"]) > 20:
        raise ValueError("Máximo 20 fotografías.")
    if not isinstance(datos['esquemas'], list) or len(datos['esquemas'])>20:
        raise ValueError('Máximo 20 esquemas.')
    ids = set()
    for foto in datos["fotos"] + datos['esquemas']:
        if not isinstance(foto.get("id"), str) or not foto["id"] or foto["id"] in ids:
            raise ValueError("Identificador de fotografía no válido o repetido.")
        ids.add(foto["id"])
        if not isinstance(foto.get("leyenda"), str):
            raise ValueError("Leyenda no válida.")
        raw = base64.b64decode(foto["datos"], validate=True)
        with Image.open(io.BytesIO(raw)) as imagen:
            imagen.verify()
    return datos


def inconsistencias(datos):
    problemas = []
    for grupo, contenido in datos["grupos"].items():
        for fila in contenido["filas"]:
            if (fila["Respuesta"] == "N/A") != (fila["Evaluación"] == "N/A"):
                problemas.append(f'{grupo}: {fila["Punto"]} - marca N/A en ambas columnas.')
            if fila["Evaluación"] == "NC" and not fila["Observaciones"].strip():
                problemas.append(f'{grupo}: {fila["Punto"]} - añade observación para NC.')
    return problemas


def revisar_informe(datos, solo_completados=False):
    errores = []
    for campo in ('numero', 'equipo', 'inspector'):
        if not datos['campos'][campo].strip():
            errores.append({'seccion': 'Identificación', 'campo': campo,
                            'mensaje': f'Completa {CAMPOS[campo]}.'})
    for grupo, contenido in datos['grupos'].items():
        for fila in contenido['filas']:
            if solo_completados and fila['Respuesta'] != 'Sí':
                continue
            mensajes = []
            if (fila['Respuesta'] == 'N/A') != (fila['Evaluación'] == 'N/A'):
                mensajes.append('Marca N/A en ambas columnas o corrige la que no corresponde.')
            if fila['Evaluación'] == 'NC' and not fila['Observaciones'].strip():
                mensajes.append('Añade una observación que explique la no conformidad.')
            for mensaje in mensajes:
                errores.append({'seccion': grupo, 'campo': fila['Punto'],
                                'mensaje': f'{fila["Punto"]}: {mensaje}'})
    return errores


TONOS_GRUPOS = [('#EFF6FF', '#35689B'), ('#EFF8F3', '#397353'),
                 ('#FFF8EC', '#8C652B'), ('#F6F1FC', '#765092')]


def punto_con_datos(fila):
    return (fila['Respuesta'] != 'Sin evaluar'
            or fila['Evaluación'] != 'Sin evaluar'
            or bool(fila['Observaciones'].strip()))


def grupos_para_exportar(datos, solo_completados=False):
    resultado = {}
    for grupo, contenido in datos['grupos'].items():
        filas = [r for r in contenido['filas'] if not solo_completados or r['Respuesta'] == 'Sí']
        if filas or not solo_completados:
            resultado[grupo] = {'filas': filas, 'observaciones': contenido['observaciones']}
    return resultado


def generar_pdf(datos, solo_completados=False):
    validar(datos)
    campos = datos["campos"]
    buffer = io.BytesIO()
    doc = BaseDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=40)
    styles = getSampleStyleSheet()
    body = ParagraphStyle("texto", parent=styles["Normal"], fontSize=9, leading=12, spaceAfter=5)
    small = ParagraphStyle("celda", parent=body, fontSize=7, leading=9, spaceAfter=0)
    heading = ParagraphStyle("seccion", parent=styles["Heading2"], fontSize=11, leading=14, textColor=colors.HexColor("#1E3A8A"), keepWithNext=True)
    titulo = ParagraphStyle('titulo_informe', parent=heading, fontSize=15, leading=18, alignment=1, spaceAfter=10)
    def p(texto, style=small):
        return Paragraph(escape(str(texto)).replace("\n", "<br/>"), style)
    centrado = ParagraphStyle('responsable', parent=small, alignment=1)
    recuadro = Table([
        [p('Solicitante: '+campos['solicitante']), '', ''],
        [p('Destino\nOriginal: Enap S.A. - DCEET\nCopia 1: Enap S.A. - DCEET\nCopia 2: Ingemars Ingeniería Ltda.'), p(campos['inspector'],centrado), p(campos['ingeniero'],centrado)],
        ['', p('Inspector visual',centrado), p('Ingeniero de operaciones',centrado)],
    ], colWidths=[180,180,180])
    recuadro.setStyle(TableStyle([
        ('SPAN',(0,0),(2,0)),('SPAN',(0,1),(0,2)),
        ('GRID',(0,0),(-1,-1),.5,colors.HexColor('#94A3B8')),
        ('VALIGN',(0,0),(-1,-1),'BOTTOM'),
        ('BACKGROUND',(1,2),(2,2),colors.HexColor('#F3F4F6')),
        ('TOPPADDING',(1,1),(2,1),26),('BOTTOMPADDING',(0,0),(-1,-1),6),
    ]))
    _,alto_recuadro = recuadro.wrap(540,700)
    if alto_recuadro>200:
        raise ValueError('Los nombres o el solicitante son demasiado extensos para el recuadro de responsables.')
    def marco(canvas, documento):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#526579"))
        canvas.drawString(36, 23, "Este informe no debe ser reproducido salvo en su totalidad.")
        canvas.drawRightString(576, 23, f"Página {documento.page}")
        if documento.page == 1:
            recuadro.drawOn(canvas,36,40)
        canvas.restoreState()
    inicio_cuerpo = 40 + alto_recuadro + 14
    doc.addPageTemplates([
        PageTemplate(id='primera', frames=[Frame(36,inicio_cuerpo,540,756-inicio_cuerpo,id='cuerpo_primera')],onPage=marco,autoNextPageTemplate='desarrollo'),
        PageTemplate(id='desarrollo',frames=[Frame(36,40,540,716,id='cuerpo_desarrollo')],onPage=marco),
    ])
    # Identificación como contenido inicial, nunca como encabezado repetido.
    story = []
    ruta_logo = Path(__file__).resolve().parent.parent / 'logo.png'
    if ruta_logo.is_file():
        logo = PDFImage(str(ruta_logo))
        escala = min(120/logo.imageWidth, 45/logo.imageHeight)
        logo.drawWidth, logo.drawHeight = logo.imageWidth*escala, logo.imageHeight*escala
        logo.hAlign = 'RIGHT'
        story.append(logo)
    story.append(p('INFORME DE INSPECCIÓN VISUAL - CHECKLIST', titulo))
    pares = [('numero','ot'),('contrato','cliente'),('fecha_inspeccion','fecha_emision'),
             ('planta','equipo'),('descripcion','aca'),('procedimiento','codigo'),('ensayo','complementario')]
    etiquetas = dict(CAMPOS, fecha_inspeccion='Fecha de inspección', fecha_emision='Fecha de emisión')
    cab = [[p(etiquetas[a].upper()+':'), p(campos[a]), p(etiquetas[b].upper()+':'), p(campos[b])] for a,b in pares]
    cab.append([p('ALCANCE:'), p(campos['alcance'] or '-'), '', ''])
    tabla = Table(cab, colWidths=[90,180,90,180])
    tabla.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.4,colors.HexColor('#CBD5E1')),
        ('BACKGROUND',(0,0),(0,-1),colors.HexColor('#F3F4F6')),
        ('BACKGROUND',(2,0),(2,-2),colors.HexColor('#F3F4F6')),
        ('SPAN',(1,-1),(3,-1)),('VALIGN',(0,0),(-1,-1),'TOP'),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    story.extend([tabla, Spacer(1,12),
             p("1. CONCLUSIONES", heading), p(campos["conclusiones"] or "Sin información registrada.", body),
             p("2. RECOMENDACIONES", heading), p(campos["recomendaciones"] or "Sin información registrada.", body)])
    story += [PageBreak(),
              p('3. DESARROLLO CHECK LIST', heading), p('N/A: No aplica · C: Conforme · NC: No conforme. Sin evaluar: celdas vacías.', small)]
    grupos_exportados = grupos_para_exportar(datos, solo_completados)
    if not grupos_exportados:
        story.append(p('No se registraron puntos ni observaciones en el checklist.', body))
    for letra, (grupo, contenido) in enumerate(grupos_exportados.items()):
        filas = [[p(f'{chr(97+letra)}) {grupo}'), p('SÍ'), p('NO'), p('N/A'), p('C'), p('NC'), p('Observaciones')]]
        for fila in contenido['filas']:
            respuesta, evaluacion = fila['Respuesta'], fila['Evaluación']
            filas.append([p(fila['Punto']), 'X' if respuesta=='Sí' else '', 'X' if respuesta=='No' else '',
                          'X' if respuesta=='N/A' or evaluacion=='N/A' else '', 'X' if evaluacion=='C' else '', 'X' if evaluacion=='NC' else '', p(fila['Observaciones'])])
        tabla = Table(filas, colWidths=[208,24,24,24,24,24,212], repeatRows=1)
        tabla.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E4EEDD')), ('GRID',(0,0),(-1,-1),.4,colors.HexColor('#AAB8B1')),
            ('VALIGN',(0,0),(-1,-1),'TOP'),('ALIGN',(1,1),(5,-1),'CENTER'),('FONTSIZE',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),1.5),('BOTTOMPADDING',(0,0),(-1,-1),1.5)]))
        story.extend([tabla, p('Observaciones: '+(contenido['observaciones'] or '-'), small), Spacer(1,5)])
    story.extend([PageBreak(), p('4. SET FOTOGRÁFICO', heading)])
    if not datos['fotos']:
        story.append(p('Sin fotografías adjuntas.', body))
    for inicio in range(0, len(datos['fotos']), 4):
        if inicio:
            story.extend([PageBreak(), p('4. SET FOTOGRÁFICO (continuación)', heading)])
        celdas = []
        for n, foto in enumerate(datos['fotos'][inicio:inicio+4], start=inicio+1):
            raw = base64.b64decode(foto['datos'])
            img = PDFImage(io.BytesIO(raw))
            factor = min(248/img.imageWidth, 160/img.imageHeight)
            img.drawWidth, img.drawHeight = img.imageWidth*factor, img.imageHeight*factor
            celdas.append([img, p(f'Imagen N.º {n}: {foto["leyenda"]}', small)])
        for pos in range(0,len(celdas),2):
            tabla = Table([[celdas[pos], celdas[pos+1] if pos+1<len(celdas) else '']], colWidths=[270,270])
            tabla.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('BOX',(0,0),(-1,-1),.4,colors.lightgrey),('BOTTOMPADDING',(0,0),(-1,-1),12)]))
            story.append(tabla)
    for n,esquema in enumerate(datos['esquemas'],start=1):
        story.extend([PageBreak(),p('5. ESQUEMAS',heading)])
        img=PDFImage(io.BytesIO(base64.b64decode(esquema['datos'])))
        factor=min(520/img.imageWidth,440/img.imageHeight)
        img.drawWidth,img.drawHeight=img.imageWidth*factor,img.imageHeight*factor
        story.extend([img,p(f'Esquema {n}: {esquema["leyenda"]}',body)])
    doc.build(story)
    return buffer.getvalue()


def main():
    st.set_page_config(page_title="Checklist IV · En desarrollo", page_icon="☑", layout="wide")
    if not st.session_state.get('autenticado') or st.session_state.get('usuario_actual') != 'jnavarrete':
        st.info('Inicia sesión con jnavarrete en la página principal para abrir este formulario.')
        st.stop()
    def key(nombre):
        return PREFIJO+nombre
    if key('datos') not in st.session_state:
        st.session_state[key('datos')] = nuevo()
    if key('revision') not in st.session_state:
        st.session_state[key('revision')] = 0
    revision = st.session_state[key('revision')]
    def widget(nombre):
        return key(f'{revision}_{nombre}')
    def cargar(datos):
        st.session_state[key('datos')] = datos
        st.session_state[key('revision')] += 1
        st.session_state.pop(key('pdf'), None)
        st.session_state[key('revision_activa')] = False
        st.rerun()
    datos = st.session_state[key('datos')]
    datos.setdefault('esquemas', [])
    datos['campos'].setdefault('descripcion','')
    datos['campos'].setdefault('aca','')
    st.title('Informe visual tipo checklist')
    st.caption('Página independiente · En desarrollo · Basada en el formato de piping y equipos')
    st.info('Esta versión guarda durante la sesión. Descarga un respaldo editable para conservar campos y fotos; no escribe en Google Sheets.')
    with st.expander('Nuevo informe / recuperar respaldo'):
        confirmar = st.checkbox('Descartar la edición actual', key=widget('confirmar'))
        if st.button('Nuevo checklist', disabled=not confirmar, key=widget('nuevo')):
            cargar(nuevo())
        archivo = st.file_uploader('Respaldo editable JSON', type=['json'], key=widget('respaldo'))
        if st.button('Recuperar respaldo', disabled=archivo is None or not confirmar, key=widget('recuperar')):
            try:
                if archivo.size > 40*1024*1024:
                    raise ValueError('El respaldo supera 40 MB.')
                recuperado = validar(json.loads(archivo.getvalue()))
            except Exception as exc:
                st.error(f'No se pudo recuperar: {exc}')
            else:
                cargar(recuperado)
    if st.button('Revisar informe', key=widget('revisar')):
        st.session_state[key('revision_activa')] = True
    resumen_revision = st.empty()
    st.caption('La revisión comprueba los datos obligatorios y la coherencia de las respuestas; no sustituye la evaluación técnica del inspector. Los puntos vacíos no se consideran errores.')
    revision_activa = st.session_state.get(key('revision_activa'), False)
    avisos_campos = {}
    avisos_grupos = {}
    tabs = st.tabs(['Identificación', '1. Conclusiones / 2. Recomendaciones', '3. Checklist', '4. Fotografías', 'Exportar'])
    campos = datos['campos']
    with tabs[0]:
        if st.button('Cargar / actualizar BASE EQUIPOS',key=widget('base')):
            try:
                st.session_state[key('equipos')] = leer_base_equipos()
            except Exception:
                st.error('No se pudo leer BASE EQUIPOS. Revisa las credenciales y el acceso al documento de Google Sheets.')
        equipos=st.session_state.get(key('equipos'),[])
        if equipos:
            indice = st.selectbox('Elegir TAG de BASE EQUIPOS',range(len(equipos)),index=None,
                format_func=lambda i: f'{equipos[i]["equipo"]} · {equipos[i]["planta"]} · {equipos[i]["descripcion"]}',
                key=widget('tag_selector'))
            reaplicar = st.button('Aplicar datos del equipo', disabled=indice is None, key=widget('aplicar_tag'))
            if indice is not None:
                seleccionado = equipos[indice]
                firma = json.dumps(seleccionado, sort_keys=True, ensure_ascii=False)
                if firma != st.session_state.get(widget('ultimo_equipo_aplicado')) or reaplicar:
                    # Aplicar antes de construir los campos en esta ejecución.
                    aplicar_equipo_formulario(st.session_state, campos, seleccionado, revision)
                    st.session_state[widget('ultimo_equipo_aplicado')] = firma
                    st.success('Datos del equipo aplicados: TAG, planta, descripción y ACA.')
                    vacios = [CAMPOS[c] for c in ('planta','descripcion','aca') if not campos[c].strip()]
                    if vacios:
                        st.info('Sin dato en BASE EQUIPOS: '+', '.join(vacios)+'. Puedes completarlo manualmente.')
            else:
                st.session_state.pop(widget('ultimo_equipo_aplicado'), None)
        columnas = st.columns(3)
        campos_identificacion = [(c,e) for c,e in CAMPOS.items() if c not in ('inspector','ingeniero')]
        for n,(campo, etiqueta) in enumerate(campos_identificacion):
            if widget(campo) not in st.session_state:
                st.session_state[widget(campo)] = campos[campo]
            campos[campo] = columnas[n%3].text_input(etiqueta, key=widget(campo))
            avisos_campos[campo] = columnas[n%3].empty()
        st.caption('El código proviene del modelo adjunto: confírmalo según el equipo y alcance del informe.')
        a,b = st.columns(2)
        campos['fecha_inspeccion'] = a.date_input('Fecha de inspección', date.fromisoformat(campos['fecha_inspeccion']), key=widget('fecha_inspeccion')).isoformat()
        campos['fecha_emision'] = b.date_input('Fecha de emisión', date.fromisoformat(campos['fecha_emision']), key=widget('fecha_emision')).isoformat()
        campos['alcance'] = st.text_area('Alcance', value=campos['alcance'], key=widget('alcance'))
    with tabs[1]:
        campos['conclusiones'] = st.text_area('1. Conclusiones', value=campos['conclusiones'], height='content', key=widget('conclusiones'))
        campos['recomendaciones'] = st.text_area('2. Recomendaciones', value=campos['recomendaciones'], height='content', key=widget('recomendaciones'))
        st.subheader('Responsables del informe')
        st.caption('Estos nombres aparecerán en el recuadro inferior de la primera página del PDF.')
        nombre_actual = campos['inspector']
        if nombre_actual == 'Jorge Hernandez':
            nombre_actual = 'Jorge Hernández'
        opciones = ['']+INSPECTORES
        if nombre_actual and nombre_actual not in opciones:
            opciones.append(nombre_actual)
        campos['inspector'] = st.selectbox('Inspector visual', opciones, index=opciones.index(nombre_actual), key=widget('responsable_inspector'))
        avisos_campos['inspector'] = st.empty()
        campos['ingeniero'] = st.text_input('Ingeniero de operaciones', value=campos['ingeniero'], placeholder='Gabriel Allendes V.', key=widget('responsable_ingeniero'))
    with tabs[2]:
        st.caption('Respuesta: Sí / No / N/A. Evaluación: C / NC / N/A. Sí no significa Conforme: describe la presencia del aspecto consultado. Nada se marca automáticamente.')
        for n,(grupo, contenido) in enumerate(datos['grupos'].items()):
            fondo, acento = TONOS_GRUPOS[n]
            panel = key(f'grupo_color_{n}')
            st.markdown(f'''<style>
            .st-key-{panel} {{background:{fondo}; border-left:5px solid {acento}; border-radius:10px; padding:14px; margin-bottom:18px;}}
            .st-key-{panel} [data-testid="stExpander"] summary {{background:{fondo}; color:{acento} !important;}}
            .st-key-{panel} [data-testid="stExpander"] summary p {{color:{acento} !important; font-weight:600;}}
            </style>''', unsafe_allow_html=True)
            with st.container(key=panel):
                with st.expander(grupo, expanded=True):
                    editado = st.data_editor(pd.DataFrame(contenido['filas']), hide_index=True, disabled=['Punto'], num_rows='fixed', use_container_width=True,
                        column_config={'Respuesta': st.column_config.SelectboxColumn(options=['Sin evaluar','Sí','No','N/A'], required=True),
                                       'Evaluación': st.column_config.SelectboxColumn(options=['Sin evaluar','C','NC','N/A'], required=True),
                                       'Observaciones': st.column_config.TextColumn(width='large')}, key=widget(f'editor_{n}'))
                    contenido['filas'] = editado.fillna('').to_dict('records')
                    contenido['observaciones'] = st.text_area('Observaciones del grupo', value=contenido['observaciones'], key=widget(f'obs_{n}'))
                    avisos_grupos[grupo] = st.empty()
    with tabs[3]:
        st.caption('Drive: los archivos cuyo nombre contiene “esquema” se muestran aparte. Las fotos existentes se conservan y no se duplican.')
        enlace=st.text_input('Enlace o ID de carpeta de Google Drive',key=widget('drive'))
        if st.button('Cargar fotografías y esquemas de Drive',key=widget('cargar_drive')):
            try:
                with st.spinner('Descargando archivos de Google Drive...'):
                    fotos_drive,esquemas_drive,avisos=leer_carpeta_drive(enlace)
                conocidos={f['id'] for f in datos['fotos']+datos['esquemas']}
                for categoria,nuevas in [('fotos',fotos_drive),('esquemas',esquemas_drive)]:
                    for imagen in nuevas:
                        if imagen['id'] in conocidos:
                            continue
                        if len(datos[categoria])>=20:
                            avisos.append(f'Máximo 20 {categoria}; no se agregó {imagen["leyenda"]}.')
                            continue
                        datos[categoria].append(imagen)
                        conocidos.add(imagen['id'])
                for aviso in avisos:
                    st.warning(aviso)
                st.success(f'Carga finalizada. Total: {len(datos["fotos"])} fotos y {len(datos["esquemas"])} esquemas.')
            except Exception:
                st.error('No se pudo cargar la carpeta. Revisa el enlace y que esté compartida con la cuenta de servicio de la aplicación.')
        archivos = st.file_uploader('Fotografías JPG o PNG (máximo 20)', type=['jpg','jpeg','png'], accept_multiple_files=True, key=widget('fotos'))
        if st.button('Agregar fotografías', key=widget('agregar')):
            try:
                nuevas = []
                hashes = {f['id'] for f in datos['fotos']}
                for archivo in archivos:
                    if archivo.size > 10*1024*1024:
                        raise ValueError(f'{archivo.name}: máximo 10 MB por imagen.')
                    identificador = hashlib.sha256(archivo.getvalue()).hexdigest()
                    if identificador in hashes:
                        continue
                    with Image.open(archivo) as im:
                        im = ImageOps.exif_transpose(im).convert('RGB')
                        im.thumbnail((1800,1800))
                        salida = io.BytesIO()
                        im.save(salida, format='JPEG', quality=88)
                    nuevas.append({'id': identificador, 'datos': base64.b64encode(salida.getvalue()).decode(), 'leyenda': archivo.name})
                    hashes.add(identificador)
                if len(datos['fotos'])+len(nuevas)>20:
                    raise ValueError('Máximo 20 fotos por informe.')
                datos['fotos'].extend(nuevas)
            except Exception as exc:
                st.error(f'No se agregaron fotografías: {exc}')
        columnas_fotos=st.columns(3)
        for n,foto in enumerate(datos['fotos']):
            with columnas_fotos[n%3]:
                st.image(base64.b64decode(foto['datos']), caption=f'Foto {n+1}', use_container_width=True)
                foto['leyenda'] = st.text_input('Leyenda', value=foto['leyenda'], key=widget('leyenda_'+foto['id']))
                if st.button('Quitar fotografía', key=widget('quitar_'+foto['id'])):
                    datos['fotos'].pop(n)
                    st.rerun()
        st.subheader('Esquemas')
        for n,esquema in enumerate(datos['esquemas']):
            st.image(base64.b64decode(esquema['datos']),use_container_width=True)
            esquema['leyenda']=st.text_input(f'Leyenda del esquema {n+1}',value=esquema['leyenda'],key=widget('esquema_'+esquema['id']))
            if st.button('Quitar esquema',key=widget('quitar_esquema_'+esquema['id'])):
                datos['esquemas'].pop(n)
                st.rerun()
    with tabs[4]:
        respaldo = json.dumps(datos, ensure_ascii=False, indent=2).encode('utf-8')
        huella = hashlib.sha256(respaldo).hexdigest()
        st.download_button('Descargar respaldo editable', respaldo, 'checklist_respaldo.json', 'application/json', key=widget('json'))
        modo = st.radio('Contenido del PDF', ['Checklist completo', 'Solo puntos completados'], key=widget('modo_exportacion'))
        solo_completados = modo == 'Solo puntos completados'
        llenos = sum(punto_con_datos(r) for g in datos['grupos'].values() for r in g['filas'])
        st.caption('Solo puntos completados incluye exclusivamente respuestas Sí. No, N/A y Sin evaluar se omiten, aunque tengan observaciones. Checklist completo muestra todos los puntos. Las observaciones de grupo aparecen solo en grupos incluidos.')
        if st.button('Preparar PDF', type='primary', key=widget('preparar')):
            errores = revisar_informe(datos, solo_completados)
            st.session_state[key('revision_activa')] = True
            revision_activa = True
            if errores:
                st.error('No se puede exportar todavía. Revisa el resumen superior y los avisos rojos en Identificación y 3. Checklist.')
            else:
                try:
                    copia = json.loads(respaldo)
                    pdf = generar_pdf(copia, solo_completados=solo_completados)
                    st.session_state[key('pdf')] = (huella, modo, pdf)
                except Exception as exc:
                    st.error(f'No se pudo generar el PDF: {exc}')
        preparado = st.session_state.get(key('pdf'))
        if preparado and preparado[:2] == (huella,modo):
            st.download_button('Descargar PDF', preparado[2], 'informe_checklist.pdf', 'application/pdf', key=widget('descargar'))
        elif preparado:
            st.info('Hay cambios: prepara nuevamente el PDF.')

    # Evaluar después de leer todos los controles para mostrar el estado actual.
    if revision_activa:
        errores = revisar_informe(datos, solo_completados)
        if errores:
            with resumen_revision.container():
                st.error(f'{len(errores)} errores que impiden exportar. Abre la pestaña indicada y corrige los puntos señalados.')
                for error in errores:
                    ubicacion = 'Conclusiones / Recomendaciones → Responsables' if error['campo']=='inspector' else error['seccion']
                    st.write(f'• {ubicacion} → {error["mensaje"]}')
        else:
            resumen_revision.success('Revisión correcta: no hay errores de validación que impidan preparar el PDF.')
        estilos_errores = []
        for campo, aviso in avisos_campos.items():
            encontrados = [e for e in errores if e['seccion'] == 'Identificación' and e['campo'] == campo]
            if encontrados:
                aviso.error(encontrados[0]['mensaje'])
                estilos_errores.append(f'.st-key-{widget(campo)} input {{border:2px solid #B42318 !important; background:#FFF1F0 !important;}}')
        for grupo, aviso in avisos_grupos.items():
            encontrados = [e for e in errores if e['seccion'] == grupo]
            if encontrados:
                with aviso.container():
                    for error in encontrados:
                        st.error(error['mensaje'])
            else:
                aviso.caption('Sin errores de validación en este grupo.')
        if estilos_errores:
            st.markdown('<style>'+''.join(estilos_errores)+'</style>', unsafe_allow_html=True)


if __name__ == '__main__':
    main()
