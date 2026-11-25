import os
import shutil
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

os.makedirs("LAB3", exist_ok=True)

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
        self.setFont("Helvetica", 8.5)
        self.setFillColor(colors.HexColor("#555555"))
        self.setStrokeColor(colors.HexColor("#D0D0D0"))
        self.setLineWidth(0.5)
        self.line(36, 36, letter[0] - 36, 36)
        
        self.drawString(36, 24, "PES University | Department of CSE | UE23CS301 SE Lab 3 Report")
        self.drawRightString(letter[0] - 36, 24, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def build_pdf_report(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=48
    )

    styles = getSampleStyleSheet()

    inst_style = ParagraphStyle(
        'InstHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.HexColor('#1A237E'),
        spaceAfter=2
    )

    sub_style = ParagraphStyle(
        'SubHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.HexColor('#3F51B5'),
        spaceAfter=6
    )

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#1A237E'),
        alignment=1,
        spaceAfter=4
    )

    doc_subtitle = ParagraphStyle(
        'DocSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#3F51B5'),
        alignment=1,
        spaceAfter=10
    )

    h1_style = ParagraphStyle(
        'Heading1_Custom',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=11.5,
        leading=15,
        textColor=colors.HexColor('#0D47A1'),
        spaceBefore=10,
        spaceAfter=4
    )

    h2_style = ParagraphStyle(
        'Heading2_Custom',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#1565C0'),
        spaceBefore=6,
        spaceAfter=2
    )

    body_style = ParagraphStyle(
        'Body_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#212121'),
        spaceAfter=4
    )

    bullet_style = ParagraphStyle(
        'Bullet_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.2,
        leading=11.5,
        textColor=colors.HexColor('#212121'),
        leftIndent=12,
        spaceAfter=2
    )

    caption_style = ParagraphStyle(
        'Caption',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=1,
        spaceBefore=4,
        spaceAfter=8
    )

    tbl_hdr = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    tbl_cell = ParagraphStyle('TC', fontName='Helvetica', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#212121'))
    tbl_cell_bold = ParagraphStyle('TCB', fontName='Helvetica-Bold', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#1A237E'))
    tbl_cell_center = ParagraphStyle('TCC', fontName='Helvetica', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#212121'), alignment=1)
    tbl_cell_success = ParagraphStyle('TCS', fontName='Helvetica-Bold', fontSize=7.5, leading=9.5, textColor=colors.HexColor('#2E7D32'), alignment=1)

    story = []

    # =========================================================================
    # PAGE 1: Header, Metadata, Objectives, Scenario & Architectural Comparison
    # =========================================================================
    story.append(Paragraph("PES UNIVERSITY", inst_style))
    story.append(Paragraph("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING", sub_style))
    story.append(Paragraph("LAB 3: COMPONENT MODELLING & ARCHITECTURAL PATTERN SELECTION", title_style))
    story.append(Paragraph("Architectural Evaluation, UML Component Modelling & Technical Justification for Coffee Kiosk System", doc_subtitle))
    story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor('#1A237E'), spaceAfter=6))

    # Student metadata table
    m_data = [
        [Paragraph("<b>Student Name:</b> Navyakishore Dash", tbl_cell), Paragraph("<b>SRN:</b> PES2UG24CS307", tbl_cell)],
        [Paragraph("<b>Course Code:</b> UE23CS301", tbl_cell), Paragraph("<b>Course:</b> Software Engineering Laboratory", tbl_cell)],
        [Paragraph("<b>Assigned Scenario:</b> Self-Service Coffee Kiosk System", tbl_cell), Paragraph("<b>Selected Pattern:</b> Layered Architecture (3-Tier)", tbl_cell)],
        [Paragraph("<b>GitHub Repository:</b> NavyakishoreDash/L6_L7_58_restaurant_bill", tbl_cell), Paragraph("<b>Date of Submission:</b> November 2026", tbl_cell)],
    ]
    m_table = Table(m_data, colWidths=[270, 270])
    m_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F4F6F9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(m_table)
    story.append(Spacer(1, 4))

    # Section 1
    story.append(Paragraph("1. Objective & Scenario Description", h1_style))
    story.append(Paragraph(
        "<b>Objective:</b> To evaluate candidate software architectural styles, select the most suitable architecture for an assigned "
        "interactive hardware system, and create a comprehensive UML 2.0 Component Diagram with proper interface notation (provided/required ball-and-socket), "
        "ports, and dependencies, accompanied by technical architectural justifications.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Assigned Scenario (Self-Service Coffee Kiosk System):</b> Designing the software architecture for an autonomous touchscreen coffee "
        "ordering kiosk located in a busy café. The system must fulfill the following operational requirements and constraints:",
        body_style
    ))
    story.append(Paragraph("• <b>Beverage Selection:</b> Customers choose from 3 coffee types (Espresso, Americano, Latte) and 2 drink sizes (Small, Large).", bullet_style))
    story.append(Paragraph("• <b>Payment Processing:</b> Customers pay exclusively via Credit Card with secure cardholder data handling.", bullet_style))
    story.append(Paragraph("• <b>Order Output:</b> System interfaces with dedicated hardware to print itemized paper receipts upon order completion.", bullet_style))
    story.append(Paragraph("• <b>Technical Constraints:</b> Handles touchscreen input, integrates physical receipt printer hardware, and persists local menu catalog and pricing data.", bullet_style))

    # Section 2
    story.append(Paragraph("2. Architectural Style Evaluation & Selection", h1_style))
    story.append(Paragraph(
        "Three architectural styles were evaluated against the kiosk's hardware integration and embedded execution requirements:",
        body_style
    ))

    cmp_data = [
        [Paragraph("Architectural Pattern", tbl_hdr), Paragraph("Key Strengths", tbl_hdr), Paragraph("Weaknesses / Risks", tbl_hdr), Paragraph("Kiosk Suitability", tbl_hdr)],
        [
            Paragraph("<b>Layered Architecture (3-Tier)</b>", tbl_cell_bold),
            Paragraph("Clean separation of concerns, zero network overhead, deterministic in-memory execution, hardware driver encapsulation, robust offline reliability.", tbl_cell),
            Paragraph("Coupling between adjacent layers; scaling applies to the whole terminal package.", tbl_cell),
            Paragraph("<b>HIGHLY SUITABLE<br/>(Selected)</b>", tbl_cell_success)
        ],
        [
            Paragraph("<b>Microservices Architecture</b>", tbl_cell_bold),
            Paragraph("Independent service deployment, polyglot technology stack, horizontal scaling of distinct functions.", tbl_cell),
            Paragraph("Severe operational complexity, inter-process network serialization, latency, and IPC failure modes on local embedded kiosk.", tbl_cell),
            Paragraph("NOT SUITABLE<br/>(Excessive Overhead)", tbl_cell_center)
        ],
        [
            Paragraph("<b>Client-Server (Thin Client)</b>", tbl_cell_bold),
            Paragraph("Centralized menu and order management across multiple store branches.", tbl_cell),
            Paragraph("Single point of failure; reliance on continuous internet connectivity stalls the kiosk during café Wi-Fi outages.", tbl_cell),
            Paragraph("PARTIALLY SUITABLE<br/>(Sync Only)", tbl_cell_center)
        ],
    ]
    t_cmp = Table(cmp_data, colWidths=[120, 180, 160, 80])
    t_cmp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A237E')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_cmp)
    story.append(Spacer(1, 4))

    # Architecture Selection Callout
    story.append(Paragraph(
        '<b>Architecture Selection Statement:</b> <i>"We chose Layered Architecture for the Self-Service Coffee Kiosk System."</i>',
        ParagraphStyle('Callout', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=colors.HexColor('#0D47A1'), alignment=1)
    ))

    story.append(PageBreak())

    # =========================================================================
    # PAGE 2: UML Component Diagram & Structural Visual Architecture
    # =========================================================================
    story.append(Paragraph("3. UML 2.0 Component Diagram", h1_style))
    story.append(Paragraph(
        "The diagram below defines the structural decomposition of the Self-Service Coffee Kiosk System using formal UML 2.0 notation. "
        "It features <b>5 components</b> organized into three cohesive horizontal layers (Presentation, Business, Data Access), "
        "with <b>4 primary assembly connectors</b> using Provided (Ball / Lollipop) and Required (Socket) interfaces, boundary ports, and data flows:",
        body_style
    ))

    diag_img = "LAB3/diagrams/coffee_kiosk_component_diagram.png"
    if os.path.exists(diag_img):
        story.append(RLImage(diag_img, width=540, height=405))
        story.append(Paragraph("Figure 1: Coffee Kiosk System UML Component Diagram with 5 components, 4 interfaces, and 3-Tier Layering.", caption_style))

    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "<b>Diagram Elements & UML 2.0 Compliance:</b><br/>"
        "• <b>Layer Containers:</b> <i>Presentation Layer</i>, <i>Business / Domain Layer</i>, and <i>Data Access Layer</i> are explicitly partitioned.<br/>"
        "• <b>Component Symbols:</b> Each component is depicted with the <code>&lt;&lt;component&gt;&gt;</code> stereotype, name, and standard component symbol.<br/>"
        "• <b>Interface Connectors:</b> Services offered are shown as Provided Interfaces (Ball notation <code>—(o</code>) and services consumed are shown as Required Interfaces (Socket notation <code>—(</code>) meeting at assembly connectors.",
        body_style
    ))

    story.append(PageBreak())

    # =========================================================================
    # PAGE 3: Component Dictionary, Interface Contracts & Data Flow
    # =========================================================================
    story.append(Paragraph("4. Component Dictionary & Interface Contracts", h1_style))

    # Component Table
    story.append(Paragraph("<b>4.1 Component Dictionary (5 Identified Components)</b>", h2_style))
    comp_data = [
        [Paragraph("Component Name", tbl_hdr), Paragraph("Layer", tbl_hdr), Paragraph("Responsibilities & Hardware Interaction", tbl_hdr)],
        [
            Paragraph("<b>User Interface Component</b>", tbl_cell),
            Paragraph("Presentation", tbl_cell_center),
            Paragraph("Controls the touchscreen GUI; renders coffee types (Espresso, Americano, Latte), drink size toggles (Small, Large), order summaries, and payment status. Captures customer touch events and forwards sanitized requests to the Order Manager via <code>IOrderService</code>.", tbl_cell)
        ],
        [
            Paragraph("<b>Order Manager Component</b>", tbl_cell),
            Paragraph("Business / Domain", tbl_cell_center),
            Paragraph("Central domain engine governing the order lifecycle state machine. Validates drink selections, calculates pricing based on size multipliers, coordinates credit card authorization via <code>IPaymentService</code>, dispatches formatted print commands to <code>IPrinterService</code>, and logs completed orders to <code>IDatabaseService</code>.", tbl_cell)
        ],
        [
            Paragraph("<b>Payment Service Component</b>", tbl_cell),
            Paragraph("Business / Domain", tbl_cell_center),
            Paragraph("Dedicated financial subsystem encapsulating credit card reader communication and payment gateway tokenization. Operates under strict PCI-DSS isolation to authorize transactions and return verified payment status.", tbl_cell)
        ],
        [
            Paragraph("<b>Receipt Printer Component</b>", tbl_cell),
            Paragraph("Business / Domain", tbl_cell_center),
            Paragraph("Hardware controller component communicating with physical thermal receipt printer hardware using low-level device drivers (ESC/POS / USB). Renders itemized receipt layouts (drink type, size, subtotal, tax, payment token, timestamp).", tbl_cell)
        ],
        [
            Paragraph("<b>Database Component</b>", tbl_cell),
            Paragraph("Data Access", tbl_cell_center),
            Paragraph("Persistent storage layer managing relational access (SQLite/SQL store). Maintains beverage catalog records, size pricing matrices, tax configurations, and transaction audit trails.", tbl_cell)
        ],
    ]
    t_comp = Table(comp_data, colWidths=[120, 80, 340])
    t_comp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A237E')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_comp)
    story.append(Spacer(1, 6))

    # Interface Table
    story.append(Paragraph("<b>4.2 Interface Contract Specifications (4 Primary Interfaces)</b>", h2_style))
    if_data = [
        [Paragraph("Interface Name", tbl_hdr), Paragraph("Provided By (Ball)", tbl_hdr), Paragraph("Required By (Socket)", tbl_hdr), Paragraph("Protocol / Mechanism", tbl_hdr), Paragraph("Key Operations & Data Flow", tbl_hdr)],
        [
            Paragraph("<b>IOrderService<br/>(Order Interface)</b>", tbl_cell),
            Paragraph("Order Manager", tbl_cell_center),
            Paragraph("User Interface", tbl_cell_center),
            Paragraph("In-Memory Synchronous API", tbl_cell_center),
            Paragraph("<code>createOrder()</code>, <code>selectBeverage(type)</code>, <code>selectSize(size)</code>, <code>submitPayment()</code>, <code>cancelOrder()</code>", tbl_cell)
        ],
        [
            Paragraph("<b>IPaymentService<br/>(Payment Interface)</b>", tbl_cell),
            Paragraph("Payment Service", tbl_cell_center),
            Paragraph("Order Manager", tbl_cell_center),
            Paragraph("Encrypted Peripheral API", tbl_cell_center),
            Paragraph("<code>processCreditCard(amount, cardToken)</code>, <code>verifyAuthorization()</code>, <code>refundPayment()</code>", tbl_cell)
        ],
        [
            Paragraph("<b>IPrinterService<br/>(Receipt Interface)</b>", tbl_cell),
            Paragraph("Receipt Printer", tbl_cell_center),
            Paragraph("Order Manager", tbl_cell_center),
            Paragraph("Hardware ESC/POS Driver", tbl_cell_center),
            Paragraph("<code>printReceipt(orderSummary)</code>, <code>getPrinterStatus()</code>, <code>paperFeed()</code>, <code>paperLowAlert()</code>", tbl_cell)
        ],
        [
            Paragraph("<b>IDatabaseService<br/>(Database Interface)</b>", tbl_cell),
            Paragraph("Database Component", tbl_cell_center),
            Paragraph("Order Manager & UI", tbl_cell_center),
            Paragraph("DAO / SQL Query Engine", tbl_cell_center),
            Paragraph("<code>getMenuCatalog()</code>, <code>getPricingInfo(drink, size)</code>, <code>saveOrderRecord(orderData)</code>", tbl_cell)
        ],
    ]
    t_if = Table(if_data, colWidths=[100, 85, 85, 95, 175])
    t_if.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A237E')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_if)

    story.append(PageBreak())

    # =========================================================================
    # PAGE 4: Detailed Written Justification, Security, Performance & Conclusion
    # =========================================================================
    story.append(Paragraph("5. Detailed Architectural Justification", h1_style))
    story.append(Paragraph(
        '<b>Selected Architecture:</b> <i>"We chose Layered Architecture for the Self-Service Coffee Kiosk System."</i>',
        h2_style
    ))

    story.append(Paragraph("<b>5.1 Two Specific Scenario-Related Reasons for Selection</b>", h2_style))
    story.append(Paragraph(
        "• <b>Reason 1 (Decoupling of Hardware Dependencies & High Maintainability):</b><br/>"
        "The kiosk interacts with heterogeneous physical hardware: a capacitive touchscreen monitor, a secure credit card reader, "
        "and an electro-mechanical thermal receipt printer. Layered Architecture encapsulates these hardware drivers inside dedicated components "
        "(<code>Receipt Printer Component</code> and <code>Payment Service Component</code>) behind well-defined abstract interfaces. "
        "If café management switches to a different receipt printer model, adopts a new touchscreen display, or updates the menu UI to add "
        "seasonal drink specials, modifications remain strictly isolated to the affected component. The central <code>Order Manager Component</code> "
        "and the <code>Database Component</code> remain entirely untouched, minimizing software regression and development costs.",
        body_style
    ))
    story.append(Paragraph(
        "• <b>Reason 2 (Local Resource Efficiency & Uninterrupted Offline Reliability):</b><br/>"
        "A self-service kiosk operates as a standalone on-premise edge terminal inside a busy retail café. Unlike Microservices, which introduce "
        "distributed network latency, remote service discovery, container orchestrators, and inter-process communication serialization, "
        "a Layered Architecture executes in-memory on the kiosk's single physical terminal. This ensures low computational overhead, zero "
        "network-induced failure points, and robust fault containment. The kiosk can continue taking orders and printing receipts deterministically "
        "even during intermittent or absent café Wi-Fi connectivity.",
        body_style
    ))

    story.append(Paragraph("<b>5.2 Security Advantage: PCI-DSS Compliance & Cardholder Isolation</b>", h2_style))
    story.append(Paragraph(
        "Because customer payments are conducted strictly via credit card, compliance with the Payment Card Industry Data Security Standard "
        "(PCI-DSS) is a mandatory requirement. The Layered Architecture enforces the Principle of Least Privilege through rigid tier boundaries: "
        "the <code>Payment Service Component</code> is completely encapsulated within the Business Layer behind a secure, tokenized <code>IPaymentService</code> "
        "interface. Neither the customer-facing <code>User Interface Component</code> nor the <code>Receipt Printer Component</code> has access to raw "
        "Primary Account Numbers (PAN), cardholder data, card reader memory buffers, or database credentials. Sensitive card data is processed exclusively "
        "within the isolated payment service module, drastically reducing the kiosk's attack surface and preventing UI injection or memory-scraping vulnerabilities.",
        body_style
    ))

    story.append(Paragraph("<b>5.3 Performance Benefit: In-Memory Latency & Deterministic Rush-Hour Throughput</b>", h2_style))
    story.append(Paragraph(
        "In a busy café environment, customer ordering speed at the kiosk directly affects store queue lengths and revenue. In Layered Architecture, "
        "inter-component communication across the 4 primary interfaces (<code>IOrderService</code>, <code>IPaymentService</code>, <code>IPrinterService</code>, "
        "and <code>IDatabaseService</code>) executes via instantaneous, in-memory procedure calls rather than REST/HTTP network requests. "
        "Beverage catalog data (Espresso, Americano, Latte) and size pricing tables (Small, Large) are loaded from the <code>Database Component</code> "
        "into memory cache during kiosk startup. This ensures zero-latency screen transitions on touchscreen touch events, instantaneous order calculation, "
        "sub-second card validation, and immediate receipt printing, maximizing patron ordering throughput during morning rush hours.",
        body_style
    ))

    story.append(Paragraph("6. Conclusion & Learning Outcomes", h1_style))
    story.append(Paragraph(
        "Through this laboratory exercise, we successfully evaluated competing architectural styles and established that Layered Architecture is "
        "the optimal pattern for the Self-Service Coffee Kiosk System. The resulting UML 2.0 Component Diagram demonstrates clear separation of concerns, "
        "robust interface abstraction via ball-and-socket notation, strict security boundaries for PCI-DSS payment compliance, and high in-memory performance "
        "tailored for standalone edge kiosk operations.",
        body_style
    ))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Generated comprehensive 4-page PDF: {pdf_path}")

# Build in both LAB3 and root directory
build_pdf_report("LAB3/PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf")
build_pdf_report("LAB3/LAB3_Report.pdf")
build_pdf_report("LAB3_Report_PES2UG24CS307.pdf")
build_pdf_report("PES2UG24CS307_NavyakishoreDash_LAB3_Report.pdf")
