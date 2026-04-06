from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import os
from datetime import datetime

def generate_estimate_pdf(estimate_text: str, client_name: str, job_type: str) -> str:
    
    # Имя файла
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"estimate_{client_name.replace(' ', '_')}_{date_str}.pdf"
    filepath = f"/tmp/{filename}"
    
    doc = SimpleDocTemplate(
        filepath,
        pagesize=letter,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )
    
    styles = getSampleStyleSheet()
    
    # Кастомные стили
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Normal'],
        fontSize=24,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=6,
        fontName='Helvetica-Bold',
        alignment=TA_LEFT
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#666666'),
        spaceAfter=4,
        fontName='Helvetica',
        alignment=TA_LEFT
    )

    section_style = ParagraphStyle(
        'Section',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#1a1a2e'),
        spaceBefore=12,
        spaceAfter=6,
        fontName='Helvetica-Bold',
    )
    
    body_style = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#333333'),
        spaceAfter=4,
        fontName='Helvetica',
        leading=16
    )

    total_style = ParagraphStyle(
        'Total',
        parent=styles['Normal'],
        fontSize=14,
        textColor=colors.HexColor('#ffffff'),
        fontName='Helvetica-Bold',
        alignment=TA_CENTER
    )

    story = []
    
    # Хедер
    header_data = [[
        Paragraph("ESTIMATE", title_style),
        Paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}\nEstimate #: EST-{datetime.now().strftime('%Y%m%d%H%M')}", subtitle_style)
    ]]
    
    header_table = Table(header_data, colWidths=[4*inch, 3*inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(header_table)
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1a1a2e')))
    story.append(Spacer(1, 0.2*inch))
    
    # Парсим текст эстимейта
    lines = estimate_text.split('\n')
    
    total_line = ""
    timeline_line = ""
    breakdown_lines = []
    notes_lines = []
    in_breakdown = False
    in_notes = False
    client_info = {}
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith('CLIENT:'):
            client_info['name'] = line.replace('CLIENT:', '').strip()
        elif line.startswith('JOB:'):
            client_info['job'] = line.replace('JOB:', '').strip()
        elif line.startswith('LOCATION:'):
            client_info['location'] = line.replace('LOCATION:', '').strip()
        elif 'BREAKDOWN' in line:
            in_breakdown = True
            in_notes = False
        elif 'NOTES' in line:
            in_notes = True
            in_breakdown = False
        elif 'TOTAL ESTIMATE' in line or 'TOTAL:' in line:
            total_line = line
            in_breakdown = False
        elif 'TIMELINE' in line:
            timeline_line = line
        elif in_breakdown and line.startswith('-'):
            breakdown_lines.append(line)
        elif in_notes and line.startswith('-'):
            notes_lines.append(line)

    # Инфо о клиенте
    story.append(Paragraph("PREPARED FOR", section_style))
    
    client_data = [
        ["Client:", client_info.get('name', client_name)],
        ["Job Type:", client_info.get('job', job_type)],
        ["Location:", client_info.get('location', '')],
    ]
    
    client_table = Table(client_data, colWidths=[1.5*inch, 5.5*inch])
    client_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME', (1,0), (1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#333333')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8f9fa')),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [colors.HexColor('#f8f9fa'), colors.white]),
    ]))
    story.append(client_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Breakdown таблица
    if breakdown_lines:
        story.append(Paragraph("SCOPE OF WORK & PRICING", section_style))
        
        table_data = [["Description", "Cost"]]
        
        current_section = ""
        for line in breakdown_lines:
            line = line.lstrip('- ').strip()
            line = line.replace('**', '')
            
            if ':' in line:
                parts = line.rsplit(':', 1)
                desc = parts[0].strip()
                cost = parts[1].strip() if len(parts) > 1 else ""
                
                # Проверяем если это секция (без $)
                if '$' not in cost and cost == "":
                    table_data.append([Paragraph(f"<b>{desc}</b>", body_style), ""])
                else:
                    table_data.append([desc, cost])
            else:
                table_data.append([line, ""])
        
        col_widths = [5*inch, 2*inch]
        breakdown_table = Table(table_data, colWidths=col_widths)
        breakdown_table.setStyle(TableStyle([
            # Хедер
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a1a2e')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 10),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dddddd')),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(breakdown_table)
    
    story.append(Spacer(1, 0.2*inch))
    
    # Тотал
    if total_line:
        total_amount = total_line.split(':', 1)[-1].strip().replace('**', '')
        total_data = [[Paragraph(f"TOTAL ESTIMATE: {total_amount}", total_style)]]
        total_table = Table(total_data, colWidths=[7*inch])
        total_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#1a1a2e')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('TOPPADDING', (0,0), (-1,-1), 12),
            ('BOTTOMPADDING', (0,0), (-1,-1), 12),
        ]))
        story.append(total_table)
    
    # Timeline
    if timeline_line:
        story.append(Spacer(1, 0.1*inch))
        timeline = timeline_line.split(':', 1)[-1].strip().replace('**', '')
        story.append(Paragraph(f"⏱ Estimated Timeline: {timeline}", body_style))
    
    # Notes
    if notes_lines:
        story.append(Spacer(1, 0.2*inch))
        story.append(Paragraph("NOTES & CONDITIONS", section_style))
        for note in notes_lines:
            note = note.lstrip('- ').strip().replace('**', '')
            story.append(Paragraph(f"• {note}", body_style))
    
    # Футер
    story.append(Spacer(1, 0.3*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#cccccc')))
    story.append(Spacer(1, 0.1*inch))
    
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#999999'),
        alignment=TA_CENTER
    )
    story.append(Paragraph(
        "This estimate is valid for 30 days. Prices may vary based on final scope of work.",
        footer_style
    ))
    
    doc.build(story)
    print(f"✅ PDF создан: {filepath}")
    return filepath

if __name__ == "__main__":
    test_estimate = """
CLIENT: John Smith
JOB: Full Bathroom Remodel with Tile Work
LOCATION: Brooklyn, NY

BREAKDOWN:
- **DEMOLITION & PREP**
- Existing tile removal (floor + shower walls): $800
- Debris hauling (3rd floor walkup surcharge): $450
- Subfloor inspection & repair: $300
- **TILE WORK**
- Floor tile 60 sq ft porcelain: $900
- Shower wall tile 85 sq ft: $1,530
- Waterproofing membrane: $500
- **FIXTURES**
- Vanity supply & install: $550
- Toilet supply & install: $275
- Shower valve & fixtures: $450

TOTAL ESTIMATE: $7,755
TIMELINE: 10-14 business days

NOTES:
- Brooklyn 3rd floor walkup adds surcharge
- Price assumes no major plumbing relocation
- Permits not included
"""
    
    path = generate_estimate_pdf(test_estimate, "John Smith", "Bathroom Remodel")
    print(f"PDF сохранён: {path}")
