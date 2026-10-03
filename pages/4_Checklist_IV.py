"""Página independiente. Instalar en pages/ junto a app.py.
Dependencias: streamlit, pandas, reportlab, Pillow.
El respaldo JSON incluye los campos, respuestas y fotografías.
"""
import base64
import hashlib
import io
import json
from datetime import date
from html import escape

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as PDFImage

PREFIJO = "checklist_iv_"
VERSION = 1
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
}


def nuevo():
    campos = {k: "" for k in CAMPOS}
    campos.update(contrato="AC31104857", cliente="Enap Refinerías Aconcagua S.A.",
                  codigo="API 510 / 572 / ASME VIII", fecha_inspeccion=date.today().isoformat(),
                  fecha_emision=date.today().isoformat(), alcance="", conclusiones="", recomendaciones="")
    return {"version": VERSION, "tipo": "checklist_iv", "campos": campos,
            "grupos": {g: {"observaciones": "", "filas": [{"Punto": p, "Respuesta": "Sin evaluar", "Evaluación": "Sin evaluar", "Observaciones": ""} for p in puntos]} for g, puntos in GRUPOS.items()},
            "fotos": []}


def validar(datos):
    if not isinstance(datos, dict) or datos.get("version") != VERSION or datos.get("tipo") != "checklist_iv":
        raise ValueError("El archivo no corresponde a esta versión del checklist.")
    base = nuevo()
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
    ids = set()
    for foto in datos["fotos"]:
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


def revisar_informe(datos):
    errores = []
    for campo in ('numero', 'equipo', 'inspector'):
        if not datos['campos'][campo].strip():
            errores.append({'seccion': 'Identificación', 'campo': campo,
                            'mensaje': f'Completa {CAMPOS[campo]}.'})
    for grupo, contenido in datos['grupos'].items():
        for fila in contenido['filas']:
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
        filas = [r for r in contenido['filas'] if not solo_completados or punto_con_datos(r)]
        if filas or contenido['observaciones'].strip() or not solo_completados:
            resultado[grupo] = {'filas': filas, 'observaciones': contenido['observaciones']}
    return resultado


def generar_pdf(datos, solo_completados=False):
    validar(datos)
    campos = datos["campos"]
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=169, bottomMargin=40)
    styles = getSampleStyleSheet()
    body = ParagraphStyle("texto", parent=styles["Normal"], fontSize=9, leading=12, spaceAfter=5)
    small = ParagraphStyle("celda", parent=body, fontSize=7, leading=9, spaceAfter=0)
    heading = ParagraphStyle("seccion", parent=styles["Heading2"], fontSize=11, leading=14, textColor=colors.HexColor("#17324D"))
    def p(texto, style=small):
        return Paragraph(escape(str(texto)).replace("\n", "<br/>"), style)
    def marco(canvas, documento):
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.HexColor("#17324D"))
        canvas.drawCentredString(306, 768, "SERVICIO DE INSPECCIÓN Y EVALUACIÓN DE ACTIVOS FÍSICOS")
        canvas.drawCentredString(306, 755, "DE ENAP REFINERÍA ACONCAGUA")
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawCentredString(306, 741, "INFORME DE INSPECCIÓN VISUAL DE CIRCUITOS Y COMPONENTES DE CAÑERÍA")
        def corto(nombre):
            valor = campos[nombre]
            return valor if len(valor) <= 85 else valor[:82] + "..."
        cab = [[p(f'Informe: {corto("numero")}'), p(f'Contrato: {corto("contrato")}'), p(f'OT/PPTO: {corto("ot")}')],
               [p(f'Cliente: {corto("cliente")}'), p(f'Planta: {corto("planta")}'), p(f'Equipo: {corto("equipo")}')],
               [p(f'Inspección: {campos["fecha_inspeccion"]}'), p(f'Emisión: {campos["fecha_emision"]}'), p(f'Procedimiento: {corto("procedimiento")}')],
               [p(f'Código: {corto("codigo")}'), p(f'Ensayo: {corto("ensayo")}'), p(f'Complementario: {corto("complementario")}')]]
        tabla = Table(cab, colWidths=[180]*3)
        tabla.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.4,colors.HexColor('#CBD5E1')),('VALIGN',(0,0),(-1,-1),'TOP'),('FONTSIZE',(0,0),(-1,-1),7)]))
        _, alto = tabla.wrap(540, 120)
        tabla.drawOn(canvas, 36, 730-alto)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#526579"))
        canvas.drawString(36, 23, "Este informe no debe ser reproducido salvo en su totalidad.")
        canvas.drawRightString(576, 23, f"Página {documento.page}")
        canvas.restoreState()
    story = [p("ALCANCE", heading), p(campos["alcance"] or "Sin información registrada.", body),
             p("1. CONCLUSIONES", heading), p(campos["conclusiones"] or "Sin información registrada.", body),
             p("2. RECOMENDACIONES", heading), p(campos["recomendaciones"] or "Sin información registrada.", body)]
    # Incluir íntegros los datos que se abrevíen en el encabezado repetido.
    for nombre in CAMPOS:
        if len(campos[nombre]) > 85:
            story.append(p(f"{CAMPOS[nombre]}: {campos[nombre]}", body))
    story += [Spacer(1,18), p(f'Solicitante: {campos["solicitante"]}', body),
              p(f'Inspector visual: {campos["inspector"]}', body), p(f'Ingeniero de operaciones: {campos["ingeniero"]}', body),
              p('Destino: Original y copia 1: Enap S.A. - DCEET. Copia 2: Ingemars Ingeniería Ltda.', small), PageBreak(),
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
    doc.build(story, onFirstPage=marco, onLaterPages=marco)
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
        columnas = st.columns(3)
        for n,(campo, etiqueta) in enumerate(CAMPOS.items()):
            campos[campo] = columnas[n%3].text_input(etiqueta, value=campos[campo], key=widget(campo))
            avisos_campos[campo] = columnas[n%3].empty()
        st.caption('El código proviene del modelo adjunto: confírmalo según el equipo y alcance del informe.')
        a,b = st.columns(2)
        campos['fecha_inspeccion'] = a.date_input('Fecha de inspección', date.fromisoformat(campos['fecha_inspeccion']), key=widget('fecha_inspeccion')).isoformat()
        campos['fecha_emision'] = b.date_input('Fecha de emisión', date.fromisoformat(campos['fecha_emision']), key=widget('fecha_emision')).isoformat()
        campos['alcance'] = st.text_area('Alcance', value=campos['alcance'], key=widget('alcance'))
    with tabs[1]:
        campos['conclusiones'] = st.text_area('1. Conclusiones', value=campos['conclusiones'], height=200, key=widget('conclusiones'))
        campos['recomendaciones'] = st.text_area('2. Recomendaciones', value=campos['recomendaciones'], height=200, key=widget('recomendaciones'))
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
        for n,foto in enumerate(datos['fotos']):
            with st.expander(f'Imagen {n+1}', expanded=False):
                st.image(base64.b64decode(foto['datos']), width=300)
                foto['leyenda'] = st.text_input('Leyenda', value=foto['leyenda'], key=widget('leyenda_'+foto['id']))
                if st.button('Quitar fotografía', key=widget('quitar_'+foto['id'])):
                    datos['fotos'].pop(n)
                    st.rerun()
    with tabs[4]:
        respaldo = json.dumps(datos, ensure_ascii=False, indent=2).encode('utf-8')
        huella = hashlib.sha256(respaldo).hexdigest()
        st.download_button('Descargar respaldo editable', respaldo, 'checklist_respaldo.json', 'application/json', key=widget('json'))
        modo = st.radio('Contenido del PDF', ['Checklist completo', 'Solo puntos completados'], key=widget('modo_exportacion'))
        solo_completados = modo == 'Solo puntos completados'
        llenos = sum(punto_con_datos(r) for g in datos['grupos'].values() for r in g['filas'])
        st.caption(f'{llenos} de 28 puntos tienen datos. Solo puntos completados incluye cualquier punto con una respuesta, evaluación u observación. Los puntos vacíos se omiten; las observaciones de grupo se conservan.')
        if st.button('Preparar PDF', type='primary', key=widget('preparar')):
            errores = revisar_informe(datos)
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
        errores = revisar_informe(datos)
        if errores:
            with resumen_revision.container():
                st.error(f'{len(errores)} errores que impiden exportar. Abre la pestaña indicada y corrige los puntos señalados.')
                for error in errores:
                    st.write(f'• {error["seccion"]} → {error["mensaje"]}')
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
