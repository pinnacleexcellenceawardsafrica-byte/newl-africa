# core/certificate.py - Updated with certified page sizing and accurate Site ID naming
#
# NOTE: generate_certificate_pdf() in this file is NOT currently called
# anywhere in views.py — the certificate shown to users (with the
# NOKIA/airtel banner, colours, borders) is produced by
# generate_certificate_excel() in excel_handler.py, filling the uploaded
# Excel template. If the visual layout (logo sizing, Site ID & Site Name
# cell position) needs fixing, that fix belongs in excel_handler.py.
#
# This file's PO Date / signature date formatting has been updated below
# so that if/when this generator is used, it matches the same
# "7 Aug 2026"-style format that excel_handler.py should also use.

import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import inch, mm, cm


# ============================================================
# CERTIFICATE PAGE SIZING
# Sourced from the certificate sizing reference:
#   Whole page: 21.59cm x 27.94cm (US Letter)
#   Margins:    top 1.91cm, bottom 1.91cm, left 1.87cm, right 1.87cm,
#               header 0.6cm, footer 0.6cm
# (ReportLab has no separate header/footer margin concept — those two
# values are respected implicitly by keeping content within top/bottom
# margins; adjust here if you add a running header/footer flowable.)
# ============================================================
CERTIFICATE_PAGE_WIDTH = 21.59 * cm
CERTIFICATE_PAGE_HEIGHT = 27.94 * cm
CERTIFICATE_MARGIN_TOP = 1.91 * cm
CERTIFICATE_MARGIN_BOTTOM = 1.91 * cm
CERTIFICATE_MARGIN_LEFT = 1.87 * cm
CERTIFICATE_MARGIN_RIGHT = 1.87 * cm


# ============================================================
# DATE FORMATTING
# Displays dates as "7 Aug 2026" instead of raw "2026-08-07" / whatever
# format the source data happens to be in. Tries a set of likely input
# formats (ISO, slash-separated, already-abbreviated) and falls back to
# returning the original string untouched if none match, so a weird
# value never crashes certificate generation.
# ============================================================
_DATE_INPUT_FORMATS = (
    '%Y-%m-%d',
    '%Y/%m/%d',
    '%d-%m-%Y',
    '%d/%m/%Y',
    '%m-%d-%Y',
    '%m/%d/%Y',
    '%d %b %Y',
    '%d %B %Y',
    '%Y-%m-%d %H:%M:%S',
)


def format_display_date(value):
    """Format a date value as 'D Mon YYYY', e.g. '7 Aug 2026'."""
    if not value:
        return ''

    if isinstance(value, datetime):
        dt = value
    else:
        raw = str(value).strip()
        if not raw:
            return ''
        dt = None
        for fmt in _DATE_INPUT_FORMATS:
            try:
                dt = datetime.strptime(raw, fmt)
                break
            except ValueError:
                continue
        if dt is None:
            # Unparseable - return the original value rather than guessing.
            return raw

    return f"{dt.day} {dt.strftime('%b')} {dt.year}"


def generate_certificate_pdf(site, po_items, output_path):
    """Generate PDF certificate using ReportLab"""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=(CERTIFICATE_PAGE_WIDTH, CERTIFICATE_PAGE_HEIGHT),
        rightMargin=CERTIFICATE_MARGIN_RIGHT,
        leftMargin=CERTIFICATE_MARGIN_LEFT,
        topMargin=CERTIFICATE_MARGIN_TOP,
        bottomMargin=CERTIFICATE_MARGIN_BOTTOM
    )
    
    styles = getSampleStyleSheet()
    story = []
    
    # Title style
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#1e6b3a'),
        spaceAfter=12,
        fontName='Helvetica-Bold'
    )
    
    # Subtitle style
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=16,
        textColor=colors.HexColor('#333333')
    )
    
    # Body style
    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_LEFT,
        spaceAfter=6
    )
    
    # Bold body
    bold_body = ParagraphStyle(
        'BoldBody',
        parent=body_style,
        fontName='Helvetica-Bold'
    )
    
    # Label style
    label_style = ParagraphStyle(
        'LabelStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_LEFT,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1e6b3a')
    )
    
    # Value style
    value_style = ParagraphStyle(
        'ValueStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_LEFT
    )
    
    # ============================================================
    # TITLE
    # ============================================================
    story.append(Paragraph('WORK COMPLETION CERTIFICATE', title_style))
    story.append(Spacer(1, 6))
    
    story.append(Paragraph(
        'This is to certify that the Services mentioned for below site have been done as per the following PO',
        subtitle_style
    ))
    story.append(Spacer(1, 12))
    
    # ============================================================
    # SITE DETAILS - Two column layout
    # Site ID & Site Name uses "{Site ID}-{Site Name}" as captured in
    # Book2/Purchase Orders — NOT "{SMP}-{Site Name}".
    # PO Date is displayed as "7 Aug 2026" style via format_display_date().
    # ============================================================
    details_data = [
        ['SMP ID:', site.smp or ''],
        ['Site ID & Site Name:', f"{site.site_id}-{site.site_name}"],
        ['Region:', site.region or 'Unknown'],
        ['Subcontractor:', 'NEWL LTD'],
        ['Site Type:', 'Green Field' if site.sub_category != 'Modernization' else 'Modernization'],
        ['PO Number:', site.po_number],
        ['PO Date:', format_display_date(site.po_date)],
    ]
    
    details_table = Table(details_data, colWidths=[80, 300])
    details_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#1e6b3a')),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 12))
    
    # ============================================================
    # SERVICES STATEMENT
    # ============================================================
    story.append(Paragraph(
        'Services for the above mentioned site have been performed as per NSN requirement',
        body_style
    ))
    story.append(Spacer(1, 8))
    
    # ============================================================
    # ITEMS TABLE
    # ============================================================
    if po_items.exists():
        items_data = [['Item No', 'SID (Description)', 'Qty', 'Remarks']]
        for po in po_items:
            items_data.append([
                str(po.item_no or ''),
                po.description or '',
                str(po.quantity or ''),
                ''
            ])
        
        items_table = Table(items_data, colWidths=[40, 280, 40, 60])
        items_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e6b3a')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(items_table)
        story.append(Spacer(1, 12))
    
    # ============================================================
    # CHECK POINTS
    # ============================================================
    story.append(Paragraph('Check Points:', bold_body))
    story.append(Spacer(1, 4))
    
    checkpoints = [
        ('NDPd', '✓'),
        ('Report Received', '✓'),
        ('Report Accepted', '✓')
    ]
    
    checkpoint_data = [[f"{label}: {status}" for label, status in checkpoints]]
    checkpoint_table = Table(checkpoint_data, colWidths=[90, 90, 90])
    checkpoint_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(checkpoint_table)
    story.append(Spacer(1, 16))
    
    # ============================================================
    # SIGNATURES
    # Dates displayed in "7 Aug 2026" style via format_display_date().
    # ============================================================
    current_date = format_display_date(datetime.now())
    
    sig_data = [
        ['', ''],
        ['TI Partners Representative', 'Nokia Project Manager/ROM'],
        ['', ''],
        [f'Name: Onesmus Nyaga', f'Name: {site.nokia_pm or "Julius Kamemba"}'],
        ['', ''],
        ['Signature: ____________________', 'Signature: ____________________'],
        ['', ''],
        [f'Date: {current_date}', f'Date: {current_date}'],
    ]
    
    sig_table = Table(sig_data, colWidths=[200, 200])
    sig_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('FONTNAME', (0, 1), (1, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (1, 3), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 1), (1, 1), colors.HexColor('#1e6b3a')),
        ('FONTNAME', (0, 7), (1, 7), 'Helvetica'),
    ]))
    story.append(sig_table)
    
    # ============================================================
    # BUILD PDF
    # ============================================================
    doc.build(story)
    return output_path