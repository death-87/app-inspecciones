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

    # Puntos 1, 2, 3, 4 con subíndices opcionales
    for sec_num, sec_info in secciones_dinamicas.items():
        subpuntos = sec_info['subpuntos']
        if subpuntos:
            story.append(Paragraph(f"{sec_num}. {sec_info['titulo']}", sec_heading_style))
            for idx, sub in enumerate(subpuntos, start=1):
                num_sub = f"{sec_num}.{idx}"
                titulo_sub = f"{num_sub} {sub['titulo']}" if sub['titulo'] else num_sub
                story.append(Paragraph(titulo_sub, subsec_heading_style))
                story.append(Paragraph(sub['contenido'] if sub['contenido'] else "-", text_style))

    # SALTO DE PÁGINA OBLIGATORIO PARA PUNTO 5
    story.append(PageBreak())
    story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS", sec_heading_style))
    story.append(Spacer(1, 4))

    # 📸 REGISTROS FOTOGRÁFICOS: Maximizados a 9.5 cm x 6.8 cm (Ancho total de columnas)
    if imagenes_procesadas:
        for i in range(0, len(imagenes_procesadas), 2):
            # Salto de página cada 6 fotos (3 filas)
            if i > 0 and i % 6 == 0:
                story.append(PageBreak())
                story.append(Paragraph("5. REGISTROS FOTOGRÁFICOS (Continuación)", sec_heading_style))
                story.append(Spacer(1, 4))

            row_cells = []
            
            # Foto Izquierda (Aprovechamiento máximo de columna)
            img_bytes1, label1 = imagenes_procesadas[i]
            img_obj1 = RLImage(BytesIO(img_bytes1), width=9.5*cm, height=6.8*cm)
            cell1 = [img_obj1, Spacer(1, 2), Paragraph(f"<font size=7.5><b>{label1}</b></font>", cell_body)]
            row_cells.append(cell1)
            
            # Foto Derecha
            if i + 1 < len(imagenes_procesadas):
                img_bytes2, label2 = imagenes_procesadas[i+1]
                img_obj2 = RLImage(BytesIO(img_bytes2), width=9.5*cm, height=6.8*cm)
                cell2 = [img_obj2, Spacer(1, 2), Paragraph(f"<font size=7.5><b>{label2}</b></font>", cell_body)]
                row_cells.append(cell2)
            else:
                row_cells.append("")

            # Tabla con 2 columnas de 9.6 cm cada una (552 pt total = ancho útil exacto)
            t_pair = Table([row_cells], colWidths=[276, 276])
            t_pair.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(t_pair)

    # Pie final con el inspector
    story.append(Spacer(1, 10))
    if inspector_firma:
        story.append(Paragraph(f"<b>Generado por:</b> {inspector_firma}", firma_style))

    doc.build(story, onFirstPage=dibujar_plantilla_pdf, onLaterPages=dibujar_plantilla_pdf)
    buffer.seek(0)
    return buffer
