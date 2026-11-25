import os
import shutil
import zipfile
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

os.makedirs("LAB3", exist_ok=True)
os.makedirs("LAB3/diagrams", exist_ok=True)

COLOR_PRIMARY = RGBColor(26, 35, 126)     # Deep Navy #1A237E
COLOR_SECONDARY = RGBColor(63, 81, 181)   # Indigo #3F51B5
COLOR_DARK = RGBColor(33, 33, 33)         # Near Black #212121
COLOR_MUTED = RGBColor(102, 102, 102)     # Gray #666666
COLOR_GREEN = RGBColor(46, 125, 50)       # Green #2E7D32

HEX_PRIMARY = "1A237E"
HEX_BG_LIGHT = "F4F6F9"
HEX_BORDER = "D0D7DE"

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D0D7DE", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

# ==============================================================================
# 1. GENERATE STRICT 1-PAGE WRITTEN JUSTIFICATION DOCX
# ==============================================================================
def create_1page_justification_docx(filename):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    # University & Department Header
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(1)
    r1 = p_inst.add_run("PES UNIVERSITY — DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING")
    r1.font.name = "Arial"
    r1.font.size = Pt(9.5)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_course.paragraph_format.space_before = Pt(0)
    p_course.paragraph_format.space_after = Pt(2)
    r2 = p_course.add_run("UE23CS301: Software Engineering Lab | Lab 3: Architectural Pattern Selection & Component Modelling")
    r2.font.name = "Calibri"
    r2.font.size = Pt(8.5)
    r2.font.bold = True
    r2.font.color.rgb = COLOR_SECONDARY

    # Metadata Table
    meta = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Assigned Scenario", "Self-Service Coffee Kiosk System"), ("Date", "November 2026")]
    ]
    tbl = doc.add_table(rows=len(meta), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="D0D7DE")
    for r_idx, (col1, col2) in enumerate(meta):
        c1, c2 = tbl.cell(r_idx, 0), tbl.cell(r_idx, 1)
        c1.width = Inches(3.6)
        c2.width = Inches(3.6)
        set_cell_background(c1, HEX_BG_LIGHT)
        set_cell_background(c2, HEX_BG_LIGHT)
        set_cell_margins(c1, top=30, bottom=30, left=80, right=80)
        set_cell_margins(c2, top=30, bottom=30, left=80, right=80)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r = p1.add_run(col1[0] + ": ")
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = COLOR_PRIMARY
        p1.add_run(col1[1]).font.size = Pt(8.5)

        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        r = p2.add_run(col2[0] + ": ")
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = COLOR_PRIMARY
        p2.add_run(col2[1]).font.size = Pt(8.5)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(2)
    p_sp.paragraph_format.space_after = Pt(4)

    # Core Required Statement
    p_stmt = doc.add_paragraph()
    p_stmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_stmt.paragraph_format.space_before = Pt(2)
    p_stmt.paragraph_format.space_after = Pt(6)
    r_stmt = p_stmt.add_run('Architecture Selection: "We chose Layered Architecture for the Self-Service Coffee Kiosk System."')
    r_stmt.font.name = "Arial"
    r_stmt.font.size = Pt(11)
    r_stmt.font.bold = True
    r_stmt.font.color.rgb = COLOR_PRIMARY

    def add_section(title, body_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        r_t = p.add_run(title)
        r_t.font.name = "Arial"
        r_t.font.size = Pt(9.5)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_SECONDARY

        p_b = doc.add_paragraph()
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(4)
        p_b.paragraph_format.line_spacing = 1.08
        r_b = p_b.add_run(body_text)
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(8.8)
        r_b.font.color.rgb = COLOR_DARK

    add_section(
        "1. Architectural Choice",
        "We selected a 3-Tier Layered Architecture structured into: (a) Presentation Layer containing the Touchscreen User Interface Component; "
        "(b) Business/Domain Layer comprising the Order Manager Component, Payment Service Component, and Receipt Printer Component; and "
        "(c) Data Access Layer hosting the Database Component (beverage menu catalog, pricing rules, and order transaction store)."
    )

    add_section(
        "2. Two Scenario-Related Reasons for Selection",
        "• Reason 1 (Hardware & Concern Decoupling): The kiosk operates distinct hardware peripherals: a customer touchscreen display, a credit card reader, "
        "and a physical thermal receipt printer. A layered architecture encapsulates hardware-specific drivers within dedicated components behind standard interfaces "
        "(IPrinterService and IPaymentService). Changing receipt printer hardware, updating touchscreen layouts, or adding seasonal beverages requires modifications "
        "strictly within isolated components without rippling across domain ordering logic or database schemas.\n"
        "• Reason 2 (Local Embedded Reliability & Resource Efficiency): The self-service kiosk is a dedicated, single-device terminal in a busy café. "
        "Unlike Microservices, which introduce distributed network latency, IPC serialization, and operational deployment overhead, Layered Architecture executes "
        "efficiently in-memory on local kiosk hardware. This guarantees robust fault containment, zero network-induced failure modes, and seamless offline reliability."
    )

    add_section(
        "3. Security Advantage (PCI-DSS Cardholder Data Isolation)",
        "The kiosk accepts credit card payments exclusively. Layered architecture enforces the Principle of Least Privilege and strict tier isolation: "
        "the Payment Service Component is encapsulated within the Business Layer behind an encrypted, tokenized IPaymentService interface. "
        "Neither the customer-facing User Interface Component nor the Receipt Printer Component has access to raw cardholder data, memory buffers, or database credentials. "
        "Sensitive financial transactions are compartmentalized from presentation logic, drastically reducing the kiosk's security attack surface and fulfilling PCI-DSS compliance."
    )

    add_section(
        "4. Performance Benefit (Zero-Latency In-Memory Execution)",
        "During peak café rush hours, transaction speed is vital. Communication across the kiosk's primary interfaces (IOrderService, IPaymentService, IPrinterService, "
        "and IDatabaseService) executes via high-speed in-memory method invocations with zero inter-service network overhead. Frequently accessed menu data "
        "(Espresso, Americano, Latte) and size pricing (Small, Large) are loaded into memory cache from the Database Component, ensuring immediate touchscreen feedback, "
        "instantaneous sub-second card validation, and rapid receipt printing to maximize kiosk customer throughput."
    )

    doc.save(filename)
    print(f"Saved 1-page docx: {filename}")

create_1page_justification_docx("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Justification.docx")
create_1page_justification_docx("LAB3/LAB3_Justification.docx")

# ==============================================================================
# 2. GENERATE STRICT 1-PAGE WRITTEN JUSTIFICATION PDF (ReportLab)
# ==============================================================================
def create_1page_justification_pdf(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=32,
        bottomMargin=32
    )

    styles = getSampleStyleSheet()

    inst_style = ParagraphStyle(
        'InstHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.HexColor('#1A237E'),
        spaceAfter=2
    )

    sub_style = ParagraphStyle(
        'SubHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.HexColor('#3F51B5'),
        spaceAfter=4
    )

    stmt_style = ParagraphStyle(
        'StmtHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.HexColor('#0D47A1'),
        spaceBefore=4,
        spaceAfter=6
    )

    sec_title = ParagraphStyle(
        'SecTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor('#1565C0'),
        spaceBefore=4,
        spaceAfter=2
    )

    body_text = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#212121'),
        spaceAfter=3
    )

    meta_text = ParagraphStyle(
        'MetaText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#333333')
    )

    story = []

    # Header
    story.append(Paragraph("PES UNIVERSITY — DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING", inst_style))
    story.append(Paragraph("UE23CS301: Software Engineering Lab | Lab 3: Architectural Pattern Selection & Component Modelling", sub_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1A237E'), spaceAfter=4))

    # Meta Table
    m_data = [
        [Paragraph("<b>Student Name:</b> Navyakishore Dash", meta_text), Paragraph("<b>SRN:</b> PES2UG24CS307", meta_text)],
        [Paragraph("<b>Assigned Scenario:</b> Self-Service Coffee Kiosk System", meta_text), Paragraph("<b>Date:</b> November 2026", meta_text)]
    ]
    m_table = Table(m_data, colWidths=[270, 270])
    m_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F4F6F9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(m_table)
    story.append(Spacer(1, 4))

    # Required Quote
    story.append(Paragraph('Architecture Selection: "We chose Layered Architecture for the Self-Service Coffee Kiosk System."', stmt_style))

    # Section 1
    story.append(Paragraph("1. Architectural Choice", sec_title))
    story.append(Paragraph(
        "We selected a <b>3-Tier Layered Architecture</b> organized into: (a) <i>Presentation Layer</i> containing the User Interface Component (touchscreen UI); "
        "(b) <i>Business/Domain Layer</i> comprising the Order Manager Component, Payment Service Component, and Receipt Printer Component; and "
        "(c) <i>Data Access Layer</i> hosting the Database Component (beverage menu catalog, size pricing, and transaction logs).",
        body_text
    ))

    # Section 2
    story.append(Paragraph("2. Two Scenario-Related Reasons for Selection", sec_title))
    story.append(Paragraph(
        "• <b>Reason 1 (Hardware & Concern Decoupling):</b> The kiosk coordinates three physical interactions: touchscreen menu display, credit card reader, "
        "and physical thermal receipt printer. Layered architecture encapsulates hardware-specific drivers and device protocols inside dedicated components behind clean "
        "interfaces (<code>IPrinterService</code>, <code>IPaymentService</code>). Swapping receipt printer hardware, updating touchscreen styling, or adding seasonal beverages "
        "requires changes strictly within isolated components without cascading into order orchestration or database logic.<br/>"
        "• <b>Reason 2 (Local Embedded Reliability & Resource Efficiency):</b> The kiosk is a standalone on-premise terminal in a busy café. "
        "Unlike Microservices, which introduce distributed network latency, IPC serialization, and multi-service failure points, Layered Architecture executes in-memory "
        "on local hardware. This provides deterministic execution, simple deployment, and uninterrupted offline reliability even if internet connectivity fluctuates.",
        body_text
    ))

    # Section 3
    story.append(Paragraph("3. Security Advantage (PCI-DSS Cardholder Data Isolation)", sec_title))
    story.append(Paragraph(
        "The kiosk operates exclusively via credit card transactions, making PCI-DSS regulatory compliance essential. Layered architecture enforces the Principle "
        "of Least Privilege: the <code>Payment Service Component</code> is strictly isolated within the Business Layer behind an encrypted, tokenized <code>IPaymentService</code> "
        "interface. The customer touchscreen UI and Receipt Printer components have zero access to cardholder data, card reader memory buffers, or database credentials. "
        "Sensitive financial data is fully contained, minimizing the kiosk's security attack surface.",
        body_text
    ))

    # Section 4
    story.append(Paragraph("4. Performance Benefit (Zero-Latency In-Memory Execution)", sec_title))
    story.append(Paragraph(
        "In a busy café environment, transaction cycle speed is critical to minimize customer queuing. In Layered Architecture, inter-component communication across "
        "the 4 primary interfaces (<code>IOrderService</code>, <code>IPaymentService</code>, <code>IPrinterService</code>, <code>IDatabaseService</code>) occurs via direct, in-memory method invocations "
        "with zero network serialization overhead. Coffee menu data (Espresso, Americano, Latte) and size pricing (Small, Large) are loaded into memory cache from the Database Component, "
        "providing instantaneous touchscreen transitions, sub-second card authorization, and rapid receipt printing to maximize kiosk customer throughput.",
        body_text
    ))

    doc.build(story)
    print(f"Saved 1-page pdf: {pdf_path}")

create_1page_justification_pdf("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Justification.pdf")
create_1page_justification_pdf("LAB3/LAB3_Justification.pdf")

# ==============================================================================
# 3. GENERATE FULL COMPREHENSIVE LAB 3 REPORT (DOCX & PDF)
# ==============================================================================
def create_comprehensive_lab3_docx(filename):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # University Header
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_inst.add_run("PES UNIVERSITY\nDEPARTMENT OF COMPUTER SCIENCE & ENGINEERING")
    r1.font.name = "Arial"
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run("LAB 3: COMPONENT MODELLING & ARCHITECTURAL PATTERN SELECTION")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(16)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_s = p_sub.add_run("Architectural Evaluation, UML 2.0 Component Diagram & Design Justification for Self-Service Coffee Kiosk System")
    r_s.font.name = "Calibri"
    r_s.font.size = Pt(11)
    r_s.font.color.rgb = COLOR_SECONDARY

    # Metadata Table
    meta = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course Code", "UE23CS301"), ("Course Title", "Software Engineering")],
        [("Assigned Scenario", "Self-Service Coffee Kiosk System"), ("Architectural Pattern", "Layered Architecture (3-Tier)")],
        [("Diagram Format", "UML 2.0 Component Diagram"), ("Date", "November 2026")]
    ]
    tbl = doc.add_table(rows=len(meta), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)
    for r_idx, (col1, col2) in enumerate(meta):
        c1, c2 = tbl.cell(r_idx, 0), tbl.cell(r_idx, 1)
        c1.width = Inches(3.25)
        c2.width = Inches(3.25)
        set_cell_background(c1, HEX_BG_LIGHT)
        set_cell_background(c2, HEX_BG_LIGHT)
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c2, top=60, bottom=60, left=100, right=100)
        p1 = c1.paragraphs[0]
        r = p1.add_run(col1[0] + ": ")
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = COLOR_PRIMARY
        p1.add_run(col1[1]).font.size = Pt(9)

        p2 = c2.paragraphs[0]
        r = p2.add_run(col2[0] + ": ")
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = COLOR_PRIMARY
        p2.add_run(col2[1]).font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY

    def add_body(text, prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15
        if prefix:
            r_pre = p.add_run(prefix)
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(10)
            r_pre.font.bold = True
            r_pre.font.color.rgb = COLOR_DARK
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.color.rgb = COLOR_DARK

    def add_bullet(text, prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        if prefix:
            r_pre = p.add_run(prefix)
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(9.5)
            r_pre.font.bold = True
            r_pre.font.color.rgb = COLOR_DARK
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_DARK

    # Section 1: Objective & Scenario Overview
    add_h1("1. Objective & Scenario Overview")
    add_body(
        "The objective of Lab 3 is to evaluate competing architectural styles (Layered, Microservices, Client-Server), "
        "select the optimal architecture for an assigned system, and construct a formal UML 2.0 Component Diagram with "
        "proper provided/required interfaces, ports, and assembly connectors."
    )
    add_body("Assigned Scenario: Self-Service Coffee Kiosk System", prefix="Scenario: ")
    add_bullet("Drink Customization: Customers select from 3 coffee types (Espresso, Americano, Latte) and 2 drink sizes (Small, Large).", "Core Functionality: ")
    add_bullet("Payment: Customers pay exclusively via Credit Card.", "Payment Mechanism: ")
    add_bullet("Dispensation: Physical receipt printer dispenses itemized receipt with order details.", "Receipt Output: ")
    add_bullet("Hardware Constraints: Kiosk operates touch screen interaction, thermal printer hardware, and local menu/pricing store.", "Technical Constraints: ")

    # Section 2: Architectural Style Evaluation
    add_h1("2. Architectural Style Evaluation & Pattern Selection")
    add_body(
        "To select the best architectural pattern, we systematically evaluated three candidate styles against the kiosk requirements:"
    )

    # Comparison Table
    tbl_cmp = doc.add_table(rows=4, cols=4)
    tbl_cmp.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_cmp)
    c_heads = ["Architectural Pattern", "Key Advantages", "Disadvantages / Risks", "Suitability for Kiosk"]
    c_widths = [Inches(1.6), Inches(2.2), Inches(2.0), Inches(1.2)]
    for c_idx, h in enumerate(c_heads):
        cell = tbl_cmp.cell(0, c_idx)
        cell.width = c_widths[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    cmp_data = [
        ["Layered Architecture (3-Tier)", "Clear separation of concerns, zero network overhead, deterministic in-memory execution, local hardware driver encapsulation, robust offline reliability.", "Coupling between adjacent layers; scaling applies to the whole application.", "HIGHLY SUITABLE (Selected)"],
        ["Microservices Architecture", "Independent service deployability, polyglot technology stack, horizontal scaling of individual business functions.", "Severe operational complexity, inter-service network serialization, high latency, IPC failure modes on local embedded kiosk.", "NOT SUITABLE (Excessive Overhead)"],
        ["Client-Server (Thin Client)", "Centralized menu and order management, simple multi-kiosk telemetry.", "Single point of failure, reliance on continuous internet connectivity; network outage stalls café kiosk ordering.", "PARTIALLY SUITABLE (Cloud Sync Only)"],
    ]
    for r_idx, row in enumerate(cmp_data):
        for c_idx, val in enumerate(row):
            cell = tbl_cmp.cell(r_idx + 1, c_idx)
            cell.width = c_widths[c_idx]
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            p = cell.paragraphs[0]
            if c_idx in (0, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 3 and "Selected" in val:
                r.bold = True
                r.font.color.rgb = COLOR_GREEN

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Core Required Statement
    p_sel = doc.add_paragraph()
    p_sel.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sel.paragraph_format.space_before = Pt(4)
    p_sel.paragraph_format.space_after = Pt(8)
    r_sel = p_sel.add_run('Architecture Selection: "We chose Layered Architecture for the Self-Service Coffee Kiosk System."')
    r_sel.font.name = "Arial"
    r_sel.font.size = Pt(11)
    r_sel.font.bold = True
    r_sel.font.color.rgb = COLOR_PRIMARY

    # Section 3: UML Component Diagram
    add_h1("3. UML 2.0 Component Diagram")
    add_body(
        "The complete UML 2.0 Component Diagram illustrates the 5 identified components partitioned across the 3 layers, "
        "interconnected via provided (ball/lollipop) and required (socket) interfaces:"
    )

    img_path = "LAB3/diagrams/coffee_kiosk_component_diagram.png"
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(img_path, width=Inches(6.4))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        r_c = p_cap.add_run("Figure 1: Coffee Kiosk System UML Component Diagram (Layered Architecture with Ball & Socket Notation).")
        r_c.font.italic = True
        r_c.font.size = Pt(9)
        r_c.font.color.rgb = COLOR_MUTED

    # Section 4: Component Dictionary
    add_h1("4. Component Dictionary (5 Components)")
    comp_records = [
        ("User Interface Component", "Presentation", "Touchscreen UI handler rendering drink selections (Espresso, Americano, Latte), size selectors (Small, Large), payment prompt screens, and live status displays. Captures customer touch events and forwards requests to Order Manager."),
        ("Order Manager Component", "Business / Domain", "Central coordination engine orchestrating order lifecycle state machine. Calculates itemized prices based on drink type and size, invokes payment authorization, commands receipt printing, and persists order records to Database."),
        ("Payment Service Component", "Business / Domain", "Dedicated payment processing component communicating with card reader hardware and financial gateway. Handles card data tokenization, authorization, and fraud prevention in strict PCI-DSS isolation."),
        ("Receipt Printer Component", "Business / Domain", "Hardware abstraction component interfacing with physical thermal printer. Formats receipt templates (order ID, timestamp, items, sizes, card last-4 digits, tax, total) and issues low-level ESC/POS device commands."),
        ("Database Component", "Data Access", "Persistent storage layer encapsulating relational data access (SQLite / SQL store). Houses beverage catalog, drink size multipliers, tax configurations, and transaction audit trails."),
    ]
    tbl_comp = doc.add_table(rows=len(comp_records)+1, cols=3)
    tbl_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_comp)
    c_w = [Inches(1.8), Inches(1.2), Inches(3.5)]
    for c_idx, h in enumerate(["Component Name", "Layer", "Responsibilities & Functionality"]):
        cell = tbl_comp.cell(0, c_idx)
        cell.width = c_w[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, (cname, clayer, cdesc) in enumerate(comp_records):
        row_c = tbl_comp.cell(r_idx + 1, 0), tbl_comp.cell(r_idx + 1, 1), tbl_comp.cell(r_idx + 1, 2)
        for c_idx, cell in enumerate(row_c):
            cell.width = c_w[c_idx]
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            p = cell.paragraphs[0]
            if c_idx == 0:
                p.add_run(cname).bold = True
            elif c_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(clayer)
            else:
                p.add_run(cdesc)
            p.runs[0].font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 5: Interface Specifications
    add_h1("5. Interface Contract Specifications (4 Primary Interfaces)")
    if_records = [
        ("IOrderService\n(Order Interface)", "Order Manager", "User Interface", "In-Memory API", "createOrder(), selectBeverage(type), selectSize(size), submitPayment(), cancelOrder()"),
        ("IPaymentService\n(Payment Interface)", "Payment Service", "Order Manager", "Encrypted Card API", "processCreditCard(amount, cardToken), verifyAuth(), voidTransaction()"),
        ("IPrinterService\n(Receipt Interface)", "Receipt Printer", "Order Manager", "Hardware ESC/POS", "printReceipt(receiptPayload), getPrinterStatus(), paperFeed()"),
        ("IDatabaseService\n(Database Interface)", "Database Component", "Order Manager & UI", "DAO / SQL Query", "getMenuCatalog(), getPriceTable(), saveOrderRecord(transaction)"),
    ]
    tbl_if = doc.add_table(rows=len(if_records)+1, cols=5)
    tbl_if.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_if)
    if_w = [Inches(1.5), Inches(1.2), Inches(1.2), Inches(1.1), Inches(2.0)]
    for c_idx, h in enumerate(["Interface", "Provided By (Ball)", "Required By (Socket)", "Protocol", "Key Operations"]):
        cell = tbl_if.cell(0, c_idx)
        cell.width = if_w[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, row in enumerate(if_records):
        for c_idx, val in enumerate(row):
            cell = tbl_if.cell(r_idx + 1, c_idx)
            cell.width = if_w[c_idx]
            set_cell_margins(cell, top=45, bottom=45, left=60, right=60)
            p = cell.paragraphs[0]
            if c_idx in (1, 2, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)
            if c_idx == 0:
                r.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 6: Written Justification
    add_h1("6. Detailed Architectural Justification")
    add_h2("6.1 Two Scenario-Related Reasons for Selection")
    add_bullet("Decoupling of Hardware Dependencies & Maintainability: The kiosk operates distinct physical peripherals (touchscreen monitor, card reader, thermal printer). Layered architecture isolates hardware-specific drivers within dedicated components behind well-defined interfaces (IPrinterService, IPaymentService). Upgrading receipt printer hardware or updating touchscreen layout requires changes strictly within isolated components without affecting core ordering logic.", "Reason 1: ")
    add_bullet("Local Embedded Reliability & Zero Network Overhead: Self-service coffee kiosks are standalone edge terminals. Unlike Microservices, which introduce distributed network latency, IPC serialization, and operational deployment overhead, Layered Architecture executes in-memory on local kiosk hardware, guaranteeing deterministic response times, fault isolation, and uninterrupted offline ordering.", "Reason 2: ")

    add_h2("6.2 Security Advantage: PCI-DSS Compliance & Cardholder Isolation")
    add_body(
        "Because customer payments are conducted exclusively via credit card, PCI-DSS compliance is mandatory. "
        "Layered Architecture establishes rigid layer boundaries enforcing the Principle of Least Privilege. "
        "The Payment Service Component is isolated within the Business Layer behind an encrypted IPaymentService interface. "
        "Neither the User Interface Component nor the Receipt Printer Component has access to raw card numbers, card reader buffers, "
        "or database credentials. Financial transactions are securely compartmentalized, minimizing the attack surface."
    )

    add_h2("6.3 Performance Benefit: In-Memory Latency & Deterministic Rush-Hour Throughput")
    add_body(
        "In a busy café environment, customer wait time at the kiosk directly impacts store revenue. Communication across "
        "the 4 primary interfaces (IOrderService, IPaymentService, IPrinterService, IDatabaseService) executes via high-speed, "
        "in-memory method invocations with zero network latency. Beverage catalog data (Espresso, Americano, Latte) and size pricing "
        "(Small, Large) are loaded into memory cache from the Database Component, ensuring immediate touchscreen feedback, "
        "instantaneous sub-second card validation, and rapid receipt printing to maximize customer ordering throughput."
    )

    doc.save(filename)
    print(f"Saved comprehensive docx: {filename}")

create_comprehensive_lab3_docx("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Report.docx")
create_comprehensive_lab3_docx("LAB3/LAB3_Report.docx")
create_comprehensive_lab3_docx("LAB3_Report_PES2UG24CS307.docx")

# ==============================================================================
# 4. GENERATE FULL COMPREHENSIVE LAB 3 REPORT (PDF via ReportLab)
# ==============================================================================
class NumberedCanvas(canvas.Canvas):
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
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#666666"))
        self.setStrokeColor(colors.HexColor("#D0D0D0"))
        self.setLineWidth(0.5)
        self.line(40, 40, letter[0] - 40, 40)
        self.drawString(40, 26, "PES University | UE23CS301 SE Lab 3 Report — Component Modelling")
        self.drawRightString(letter[0] - 40, 26, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def create_comprehensive_lab3_pdf(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=55
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#1A237E'),
        alignment=1,
        spaceAfter=4
    )

    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor('#3F51B5'),
        alignment=1,
        spaceAfter=12
    )

    h1_style = ParagraphStyle(
        'Heading1_Custom',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12.5,
        leading=16,
        textColor=colors.HexColor('#0D47A1'),
        spaceBefore=12,
        spaceAfter=5
    )

    h2_style = ParagraphStyle(
        'Heading2_Custom',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=colors.HexColor('#1565C0'),
        spaceBefore=8,
        spaceAfter=3
    )

    body_style = ParagraphStyle(
        'Body_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#212121'),
        spaceAfter=5
    )

    bullet_style = ParagraphStyle(
        'Bullet_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#212121'),
        leftIndent=12,
        spaceAfter=2.5
    )

    caption_style = ParagraphStyle(
        'Caption',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=1,
        spaceBefore=3,
        spaceAfter=8
    )

    tbl_hdr = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    tbl_cell = ParagraphStyle('TC', fontName='Helvetica', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#212121'))
    tbl_cell_center = ParagraphStyle('TCC', fontName='Helvetica', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#212121'), alignment=1)

    story = []

    # Title & Subtitle
    story.append(Paragraph("PES UNIVERSITY — DEPARTMENT OF CSE", subtitle_style))
    story.append(Paragraph("LAB 3: COMPONENT MODELLING & ARCHITECTURAL PATTERN SELECTION", title_style))
    story.append(Paragraph("Architectural Evaluation & UML Component Design for Self-Service Coffee Kiosk System", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor('#1A237E'), spaceAfter=8))

    # Meta Table
    m_data = [
        [Paragraph("<b>Student Name:</b> Navyakishore Dash", tbl_cell), Paragraph("<b>SRN:</b> PES2UG24CS307", tbl_cell)],
        [Paragraph("<b>Course:</b> UE23CS301 Software Engineering", tbl_cell), Paragraph("<b>Pattern:</b> Layered Architecture (3-Tier)", tbl_cell)]
    ]
    m_table = Table(m_data, colWidths=[265, 265])
    m_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F4F6F9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(m_table)
    story.append(Spacer(1, 6))

    # Section 1
    story.append(Paragraph("1. Objective & Scenario Description", h1_style))
    story.append(Paragraph(
        "The objective of Lab 3 is to evaluate architectural styles, select the most appropriate pattern for an assigned kiosk scenario, "
        "and produce a comprehensive UML 2.0 Component Diagram with at least 5 components and 4 interfaces using proper ball-and-socket notation.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Scenario (Self-Service Coffee Kiosk):</b> A customer-facing kiosk in a busy café enabling selection of 3 coffee types "
        "(Espresso, Americano, Latte), 2 drink sizes (Small, Large), exclusive Credit Card payment, and automated thermal receipt printing.",
        body_style
    ))

    # Section 2
    story.append(Paragraph("2. Architectural Selection Justification", h1_style))
    story.append(Paragraph('<b>Architecture Selection:</b> <i>"We chose Layered Architecture for the Self-Service Coffee Kiosk System."</i>', body_style))
    story.append(Paragraph("<b>Two Scenario-Related Reasons:</b>", h2_style))
    story.append(Paragraph("• <b>Hardware Abstraction & Modular Maintenance:</b> The kiosk relies on physical touchscreen, card reader, and thermal printer hardware. Layered Architecture encapsulates low-level device drivers inside dedicated components behind clean interfaces (<code>IPrinterService</code>, <code>IPaymentService</code>). Swapping receipt printer models or updating touchscreen UI layouts does not ripple into core order orchestration.", bullet_style))
    story.append(Paragraph("• <b>Local Resource Efficiency & Offline Resilience:</b> Because the kiosk operates as an autonomous on-premise edge terminal, a layered monolith executes directly in-memory without the network latency, serialization cost, or multi-service failure risks of microservices. It continues operating deterministically even during transient café Wi-Fi disruptions.", bullet_style))

    story.append(Paragraph("<b>Security Advantage (PCI-DSS Compliance):</b>", h2_style))
    story.append(Paragraph("By enforcing strict tier boundaries, the <code>Payment Service Component</code> is isolated within the Business Layer. Customer-facing touchscreen components and receipt printing components never touch raw card numbers or card reader buffers, fulfilling PCI-DSS compliance and minimizing the attack surface.", body_style))

    story.append(Paragraph("<b>Performance Benefit (Sub-Second Interaction):</b>", h2_style))
    story.append(Paragraph("All component communication executes via zero-latency in-memory method invocations. Beverage catalog data and size pricing are cached in memory from the Database Component, ensuring instant touchscreen transitions and rapid receipt dispensation during busy rush hours.", body_style))

    story.append(PageBreak())

    # Section 3: UML Component Diagram
    story.append(Paragraph("3. UML 2.0 Component Diagram", h1_style))
    story.append(Paragraph("The diagram below illustrates the 5 components organized across Presentation, Business, and Data layers with ball-and-socket interfaces:", body_style))

    diag_img = "LAB3/diagrams/coffee_kiosk_component_diagram.png"
    if os.path.exists(diag_img):
        story.append(RLImage(diag_img, width=530, height=395))
        story.append(Paragraph("Figure 1: Coffee Kiosk System UML Component Diagram with 5 components and 4 interfaces.", caption_style))

    story.append(Spacer(1, 8))

    # Section 4: Component & Interface Tables
    story.append(Paragraph("4. Component & Interface Specifications", h1_style))

    c_table_data = [
        [Paragraph("Component Name", tbl_hdr), Paragraph("Layer", tbl_hdr), Paragraph("Responsibilities & Interfaces", tbl_hdr)],
        [Paragraph("User Interface Component", tbl_cell), Paragraph("Presentation", tbl_cell_center), Paragraph("Renders touchscreen menu, beverage choices, size selection, card payment prompts. Requires: <code>IOrderService</code>.", tbl_cell)],
        [Paragraph("Order Manager Component", tbl_cell), Paragraph("Business", tbl_cell_center), Paragraph("Manages order state machine, computes size pricing, triggers payment & printing. Provides: <code>IOrderService</code>; Requires: <code>IPaymentService</code>, <code>IPrinterService</code>, <code>IDatabaseService</code>.", tbl_cell)],
        [Paragraph("Payment Service Component", tbl_cell), Paragraph("Business", tbl_cell_center), Paragraph("Encapsulates card reader hardware and gateway authorization in PCI-DSS isolation. Provides: <code>IPaymentService</code>.", tbl_cell)],
        [Paragraph("Receipt Printer Component", tbl_cell), Paragraph("Business", tbl_cell_center), Paragraph("Controls physical thermal printer driver (ESC/POS) to print itemized receipt tickets. Provides: <code>IPrinterService</code>.", tbl_cell)],
        [Paragraph("Database Component", tbl_cell), Paragraph("Data Access", tbl_cell_center), Paragraph("Stores coffee menu catalog, size pricing rules, tax rates, and transaction audit logs. Provides: <code>IDatabaseService</code>.", tbl_cell)],
    ]
    c_table = Table(c_table_data, colWidths=[140, 80, 310])
    c_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A237E')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(c_table)

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Saved comprehensive pdf: {pdf_path}")

create_comprehensive_lab3_pdf("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf")
create_comprehensive_lab3_pdf("LAB3/LAB3_Report.pdf")

# ==============================================================================
# 5. PACKAGE LAB 3 SUBMISSION ZIP
# ==============================================================================
staging_dir = "LAB3_Submission_Staging"
if os.path.exists(staging_dir):
    shutil.rmtree(staging_dir)
os.makedirs(staging_dir, exist_ok=True)
os.makedirs(os.path.join(staging_dir, "diagrams"), exist_ok=True)

# Copy diagrams
shutil.copy2("LAB3/diagrams/coffee_kiosk_component_diagram.png", os.path.join(staging_dir, "diagrams/coffee_kiosk_component_diagram.png"))
shutil.copy2("LAB3/diagrams/coffee_kiosk_component_diagram.pdf", os.path.join(staging_dir, "diagrams/coffee_kiosk_component_diagram.pdf"))

# Copy 1-page justification documents (exact match for page 5 & 7 requirement)
shutil.copy2("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Justification.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB3_Justification.docx"))
shutil.copy2("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Justification.pdf", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB3_Justification.pdf"))

# Copy full reports
shutil.copy2("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Report.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB3_Report.docx"))
shutil.copy2("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf"))

# Create github_link.txt
github_txt = """========================================================================
LAB 3: COMPONENT MODELLING & ARCHITECTURAL PATTERN SELECTION
========================================================================

STUDENT DETAILS:
Name: Navyakishore Dash
SRN: PES2UG24CS307
Course: UE23CS301 - Software Engineering
Department: Department of Computer Science and Engineering
Institution: PES University

ASSIGNED SCENARIO:
Self-Service Coffee Kiosk System (3 Coffee Types, 2 Sizes, Credit Card Only, Receipt Printer)

SELECTED ARCHITECTURE:
Layered Architecture (3-Tier: Presentation, Business, Data)

GITHUB REPOSITORY LINK:
https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill
(Directory: LAB3)

========================================================================
LAB 3 DELIVERABLES SUMMARY
========================================================================
1. Component Diagram:
   - diagrams/coffee_kiosk_component_diagram.png (High-Res 300 DPI)
   - diagrams/coffee_kiosk_component_diagram.pdf (Vector PDF)
   - Features: 5 components, 4 interfaces (Ball & Socket notation), ports, and layer packages.

2. Written Justification (Strict 1-Page Document):
   - PES2UG24CS307_NavyakishoreDash_LAB3_Justification.docx (Word Document)
   - PES2UG24CS307_NavyakishoreDash_LAB3_Justification.pdf (PDF Document)
   - Includes: Architectural Choice, Two Reasons, Security Advantage, Performance Benefit.

3. Comprehensive Full Lab Report:
   - PES2UG24CS307_NavyakishoreDash_LAB3_Report.docx
   - PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf
   - Includes: Full background, component dictionary, interface contract tables, and embedded diagrams.
"""

with open(os.path.join(staging_dir, "README.txt"), "w", encoding="utf-8") as f:
    f.write(github_txt)

with open(os.path.join(staging_dir, "github_link.txt"), "w", encoding="utf-8") as f:
    f.write(github_txt)

# Zip files
zips = [
    "PES2UG24CS307_Lab3.zip",
    "PES2UG24CS307_NavyakishoreDash_LAB3.zip"
]
for z in zips:
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(staging_dir):
            for file in files:
                fp = os.path.join(root, file)
                rp = os.path.relpath(fp, staging_dir)
                zipf.write(fp, rp)
    print(f"Created {z} ({os.path.getsize(z)/1024:.1f} KB)")

shutil.rmtree(staging_dir)
print("Lab 3 build and packaging completed successfully!")
