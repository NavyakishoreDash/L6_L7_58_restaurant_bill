import os
import shutil
import zipfile
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
import fitz

os.makedirs("LAB2", exist_ok=True)
os.makedirs("LAB2/screenshots", exist_ok=True)

STUDENT_NAME = "Navyakishore Dash"
STUDENT_SRN = "PES2UG24CS307"
STUDENT_SECTION = "Section -E"
COURSE_CODE = "UE23CS301: Software Engineering Laboratory"
LAB_TITLE = "Lab 2: Agile Backlog Creation & Sprint Simulation in Jira"
SCENARIO = "Problem Statement #13: Patient Health Record Consent Management System"
GITHUB_URL = "https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill"

# -----------------------------------------------------------------------------
# NUMBERED CANVAS FOR RUNNING HEADER & FOOTER
# -----------------------------------------------------------------------------
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
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont('Helvetica-Bold', 7.5)
        self.setFillColor(colors.HexColor('#64748B'))

        # Running Header on Pages 2, 3, 4
        if self._pageNumber > 1:
            self.drawString(45, 755, 'UE23CS301: SOFTWARE ENGINEERING LAB  •  LAB 2 REPORT')
            self.setFont('Helvetica', 7.5)
            self.drawRightString(567, 755, 'Jira Backlog & Sprint Simulation | Problem #13')
            self.setStrokeColor(colors.HexColor('#E2E8F0'))
            self.setLineWidth(0.75)
            self.line(45, 747, 567, 747)

        # Running Footer on all pages
        self.setStrokeColor(colors.HexColor('#E2E8F0'))
        self.setLineWidth(0.75)
        self.line(45, 42, 567, 42)
        
        self.setFont('Helvetica-Bold', 7.5)
        self.setFillColor(colors.HexColor('#1E3A8A'))
        self.drawString(45, 30, 'PES UNIVERSITY')
        self.setFont('Helvetica', 7.5)
        self.setFillColor(colors.HexColor('#64748B'))
        self.drawString(120, 30, f'•  Dept of CSE  •  {STUDENT_NAME} ({STUDENT_SRN})  •  Section E')
        self.setFont('Helvetica-Bold', 7.5)
        self.setFillColor(colors.HexColor('#0F172A'))
        self.drawRightString(567, 30, f'Page {self._pageNumber} of {page_count}')
        
        self.restoreState()

# -----------------------------------------------------------------------------
# 1. BUILD BEAUTIFIED STRICT 4-PAGE PDF REPORT
# -----------------------------------------------------------------------------
def build_pdf(pdf_path):
    # Page size: Letter (612 x 792 pt). Margins: Left=45, Right=45, Top=48, Bottom=48
    # Usable Width: 612 - 90 = 522 pt.
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=45,
        rightMargin=45,
        topMargin=46,
        bottomMargin=46
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12.5,
        leading=15,
        textColor=colors.HexColor('#0F172A')
    )

    subtitle_style = ParagraphStyle(
        'DocSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=12.5,
        textColor=colors.HexColor('#0C66E4')
    )

    meta_label = ParagraphStyle(
        'MetaLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#475569')
    )

    meta_val = ParagraphStyle(
        'MetaVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#0F172A')
    )

    sec_heading = ParagraphStyle(
        'SecHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#1E293B')
    )

    caption_style = ParagraphStyle(
        'CaptionStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#475569'),
        alignment=1  # Centered
    )

    q_style = ParagraphStyle(
        'QuestionStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#0C66E4')
    )

    ans_style = ParagraphStyle(
        'AnswerStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.2,
        leading=11.5,
        textColor=colors.HexColor('#334155'),
        alignment=4  # Justified
    )

    story = []

    # ========================== PAGE 1 ==========================
    # Top Academic Header Banner
    banner_p1 = Paragraph("<b>PES UNIVERSITY</b> &nbsp;|&nbsp; DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING", subtitle_style)
    banner_p2 = Paragraph("<b>UE23CS301: SOFTWARE ENGINEERING LABORATORY &nbsp;—&nbsp; LAB 2 REPORT</b>", title_style)
    banner_p3 = Paragraph("<b>Topic:</b> Agile Backlog Creation, Story Estimation & Sprint Simulation in Jira", meta_val)

    header_table = Table(
        [[banner_p1], [banner_p2], [banner_p3]],
        colWidths=[522]
    )
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('LINELEFT', (0, 0), (0, -1), 3.5, colors.HexColor('#0C66E4')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 5))

    # Student Metadata Card (2x2 Grid)
    info_data = [
        [
            Paragraph("<b>Student Name:</b> Navyakishore Dash", meta_val),
            Paragraph("<b>SRN:</b> PES2UG24CS307", meta_val),
            Paragraph("<b>Section:</b> Section -E", meta_val)
        ],
        [
            Paragraph("<b>Scenario:</b> Problem Statement #13 — Patient Health Record Consent System", meta_val),
            Paragraph("<b>Platform:</b> Jira Software (Cloud)", meta_val),
            Paragraph("<b>Sprint:</b> PHRC Sprint 1 (Active)", meta_val)
        ]
    ]
    info_table = Table(info_data, colWidths=[200, 160, 162])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FFFFFF')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#F1F5F9')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 6))

    # Screenshots Heading
    story.append(Paragraph("<b>PART 1: AGILE BACKLOG CREATION & USER STORY SPECIFICATION</b>", sec_heading))
    story.append(Spacer(1, 3))

    # Screenshot 1: Backlog with Epics and User stories
    im1_path = "LAB2/screenshots/SS1_jira_backlog.png"
    im1 = RLImage(im1_path, width=516, height=254)
    story.append(im1)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Figure 1.1:</b> Jira Backlog View — 3 Defined Epics, PHRC Sprint 1 Backlog Items, Priority Rankings & Fibonacci Story Points", caption_style))
    story.append(Spacer(1, 5))

    # Screenshot 2: Create Story Modal
    im2_path = "LAB2/screenshots/SS1_create_story.png"
    im2 = RLImage(im2_path, width=516, height=152)
    story.append(im2)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Figure 1.2:</b> Agile User Story Creation Modal — User Story Format (As a... I want... So that...) with Acceptance Criteria for PHRC-5", caption_style))
    
    story.append(PageBreak())

    # ========================== PAGE 2 ==========================
    # Top Page Header
    p2_title = Table([
        [
            Paragraph("<b>PART 2: STORY POINT ESTIMATION & ACTIVE SPRINT EXECUTION</b>", sec_heading),
            Paragraph("<b>Jira Execution Board</b>", caption_style)
        ]
    ], colWidths=[380, 142])
    p2_title.setStyle(TableStyle([
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(p2_title)
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0C66E4'), spaceAfter=5, spaceBefore=2))

    # Screenshot 3: Work item details & story points
    im3_path = "LAB2/screenshots/SS2_jira_story_points.png"
    im3 = RLImage(im3_path, width=516, height=258)
    story.append(im3)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Figure 2.1:</b> Work Item Details & Estimation Log — Planning Poker Consensus, Fibonacci Story Points (3 pts) & Time Tracking", caption_style))
    story.append(Spacer(1, 7))

    # Screenshot 4: Active Sprint board
    im4_path = "LAB2/screenshots/SS3_jira_active_sprint.png"
    im4 = RLImage(im4_path, width=516, height=258)
    story.append(im4)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Figure 2.2:</b> Jira Active Sprint Scrum Board — Task Progression across TO DO (2), IN PROGRESS (1), and DONE (5) Swimlanes", caption_style))

    story.append(PageBreak())

    # ========================== PAGE 3 ==========================
    # Top Page Header
    p3_title = Table([
        [
            Paragraph("<b>PART 3: SPRINT PERFORMANCE METRICS & BURNDOWN ANALYSIS</b>", sec_heading),
            Paragraph("<b>Sprint 1 Velocity Audit</b>", caption_style)
        ]
    ], colWidths=[380, 142])
    p3_title.setStyle(TableStyle([
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(p3_title)
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0C66E4'), spaceAfter=5, spaceBefore=2))

    # Metric KPI Dashboard Table (4 columns)
    kpi_data = [
        [
            Paragraph("<b>Planned Capacity</b><br/><font color='#0C66E4' size=10><b>18 Points</b></font><br/><font color='#64748B' size=7>8 User Stories</font>", meta_val),
            Paragraph("<b>Completed Work</b><br/><font color='#16A34A' size=10><b>13 Points</b></font><br/><font color='#64748B' size=7>5 Stories DONE (72%)</font>", meta_val),
            Paragraph("<b>In Progress</b><br/><font color='#D97706' size=10><b>5 Points</b></font><br/><font color='#64748B' size=7>US-2.2 Scoped Retrieval</font>", meta_val),
            Paragraph("<b>Sprint Duration</b><br/><font color='#0F172A' size=10><b>10 Work Days</b></font><br/><font color='#64748B' size=7>Nov 2 – Nov 13, 2026</font>", meta_val)
        ]
    ]
    kpi_table = Table(kpi_data, colWidths=[130, 130, 130, 132])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 5))

    # Screenshot 5: Burndown Chart
    im5_path = "LAB2/screenshots/SS4_jira_burndown_chart.png"
    im5 = RLImage(im5_path, width=516, height=272)
    story.append(im5)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Figure 3.1:</b> Jira Sprint Burndown Chart — Ideal Guideline (Grey Dashed) vs. Actual Remaining Story Points (Red Step Curve)", caption_style))
    story.append(Spacer(1, 6))

    # Analytical Observations Card
    obs_content = Paragraph(
        "<b>Key Burndown Chart Insights & Velocity Analysis:</b><br/>"
        "• <b>Guideline Tracking:</b> The actual remaining points curve tracked tightly against the linear guideline from Nov 2 to Nov 6 (burning 6 points).<br/>"
        "• <b>Non-Working Days Plateau:</b> A flat plateau occurred during the weekend non-working window (Nov 7–8), correctly modeling team availability.<br/>"
        "• <b>Significant Step Burndown:</b> A 5-point downward step occurred on Nov 9 upon peer-review approval and merge of US-3.2 (Audit Logging).<br/>"
        "• <b>Capacity Verification:</b> Demonstrates that committing to 18 points across 8 stories matched team velocity without sprint failure.",
        meta_val
    )
    obs_table = Table([[obs_content]], colWidths=[522])
    obs_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
        ('LINELEFT', (0, 0), (0, -1), 3.5, colors.HexColor('#16A34A')),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(obs_table)

    story.append(PageBreak())

    # ========================== PAGE 4 ==========================
    # Top Page Header
    p4_title = Table([
        [
            Paragraph("<b>PART 4: POST-SPRINT REFLECTIONS & AGILE EVALUATION</b>", sec_heading),
            Paragraph("<b>Sprint 1 Review</b>", caption_style)
        ]
    ], colWidths=[380, 142])
    p4_title.setStyle(TableStyle([
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(p4_title)
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0C66E4'), spaceAfter=6, spaceBefore=2))

    # Helper function to create an elegant Q&A card
    def make_qa_card(q_num, q_title, a_text):
        q_p = Paragraph(f"<b>{q_num}. {q_title}</b>", q_style)
        a_p = Paragraph(a_text, ans_style)
        t = Table([[q_p], [a_p]], colWidths=[522])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#FFFFFF')),
            ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
            ('LINELEFT', (0, 0), (0, -1), 3.5, colors.HexColor('#0C66E4')),
            ('TOPPADDING', (0, 0), (-1, 0), 4),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
            ('TOPPADDING', (0, 1), (-1, 1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 9),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ]))
        return t

    # Q1
    t1 = make_qa_card(
        1, "Did your estimations reflect the actual effort?",
        "<b>Yes.</b> Straightforward administrative and configuration tasks, such as defining time-bounded validity "
        "windows (US-1.2) and multi-channel patient notifications (US-3.1), required fewer story points (2 points each). "
        "In contrast, complex security and cryptographic authorization tasks, such as provider credential verification "
        "with digital signatures (US-2.3) and HIPAA-compliant tamper-evident audit logging with SHA-256 hash chaining "
        "(US-3.2), involved higher uncertainty and technical complexity, justifying 5 story points each."
    )
    story.append(t1)
    story.append(Spacer(1, 6))

    # Q2
    t2 = make_qa_card(
        2, "Was your backlog well-prioritized?",
        "<b>Yes.</b> Core patient privacy controls, including granular record selection (US-1.1), real-time consent revocation "
        "(US-1.3), and tamper-evident audit logging (US-3.2), were assigned High priority because patient confidentiality, "
        "regulatory compliance, and access governance are fundamental prerequisites before clinical providers can interact "
        "with medical records. Secondary administrative features, such as multi-channel patient alert notifications (US-3.1), "
        "were prioritized as Medium as specified in the Lab 1 requirements."
    )
    story.append(t2)
    story.append(Spacer(1, 6))

    # Q3
    t3 = make_qa_card(
        3, "How did your simulated sprint align with your plan?",
        "Sprint 1 focused on the end-to-end patient consent lifecycle, clinical provider access requests, and immutable audit "
        "logging, successfully completing 13 out of 18 planned story points across 5 user stories. High-risk components "
        "(consent granting, scoped retrieval, and revocation) were executed first to validate the architectural boundaries "
        "established in Lab 1. The remaining 5 story points (provider credential verification and notification alerting) "
        "were positioned for Sprint 2, maintaining alignment with our phased delivery roadmap."
    )
    story.append(t3)
    story.append(Spacer(1, 6))

    # Q4
    t4 = make_qa_card(
        4, "What insights did the burndown chart give about your team's capacity?",
        "The burndown chart demonstrated a realistic and steady completion rate that closely tracked the ideal linear guideline. "
        "Work progressed methodically across the 10 working days, with a flat plateau during the weekend non-working days (Nov 7–8) "
        "reflecting accurate calendar constraints. A significant downward step of 5 story points occurred on Nov 9 upon the "
        "successful integration and peer review of the audit logging framework, confirming that the 18-point sprint capacity "
        "was well-calibrated to our team velocity without overcommitment."
    )
    story.append(t4)
    story.append(Spacer(1, 7))

    # Academic Sign-off Footer Card
    sign_p = Paragraph(
        "<b>Laboratory Evaluation & Jira Demonstration Sign-off:</b><br/>"
        "• Jira Workspace: <i>Patient Health Record Consent Gateway (PHRC)</i> &nbsp;|&nbsp; Board: <i>PHRC board</i><br/>"
        "• Verified Deliverables: <i>Backlog with Epics (SS1) &bull; Story Point Estimations (SS2) &bull; Active Sprint Board (SS3) &bull; Burndown Chart (SS4)</i><br/>"
        "• Evaluation Status: Complete and ready for instructor review & grading &nbsp;|&nbsp; Course: UE23CS301 &bull; PES University",
        meta_val
    )
    sign_table = Table([[sign_p]], colWidths=[522])
    sign_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
        ('LINELEFT', (0, 0), (0, -1), 3.5, colors.HexColor('#475569')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(sign_table)

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Built Beautified PDF: {pdf_path}")

# -----------------------------------------------------------------------------
# 2. BUILD BEAUTIFIED FORMATTED WORD DOCUMENT REPORT (.DOCX)
# -----------------------------------------------------------------------------
def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_left_border(cell, color_hex, sz="24"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def build_docx(docx_path):
    doc = docx.Document()

    # Set standard 0.75 in margins
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(10.5)
    font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # PAGE 1: Academic Title Header + Student Info + Screenshots 1 & 2
    tb_header = doc.add_table(rows=1, cols=1)
    tb_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tb_header.cell(0, 0)
    c.width = Inches(7.0)
    set_cell_shading(c, "F8FAFC")
    set_cell_left_border(c, "0C66E4", sz="36")
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    
    r1 = p.add_run("PES UNIVERSITY  |  DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\n")
    r1.bold = True
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = RGBColor(0x0C, 0x66, 0xE4)

    r2 = p.add_run("UE23CS301: SOFTWARE ENGINEERING LABORATORY  —  LAB 2 REPORT\n")
    r2.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    r3 = p.add_run("Topic: Agile Backlog Creation, Story Estimation & Sprint Simulation in Jira")
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # Student Info Table
    tb_info = doc.add_table(rows=2, cols=3)
    tb_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_rows = [
        [f"Student: {STUDENT_NAME}", f"SRN: {STUDENT_SRN}", f"Section: {STUDENT_SECTION}"],
        [f"Scenario: Problem Statement #13 (PHRC)", f"Platform: Jira Software (Cloud)", f"Sprint: PHRC Sprint 1 (Active)"]
    ]
    for i, row in enumerate(info_rows):
        for j, text in enumerate(row):
            cell = tb_info.cell(i, j)
            set_cell_shading(cell, "FFFFFF")
            cp = cell.paragraphs[0]
            cp.paragraph_format.space_before = Pt(2)
            cp.paragraph_format.space_after = Pt(2)
            run = cp.add_run(text)
            run.font.size = Pt(8.5)
            run.bold = (i == 0)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Section Heading
    p_s1 = doc.add_paragraph()
    r_s1 = p_s1.add_run("PART 1: AGILE BACKLOG CREATION & USER STORY SPECIFICATION")
    r_s1.bold = True
    r_s1.font.size = Pt(11)
    r_s1.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    p_s1.paragraph_format.space_before = Pt(2)
    p_s1.paragraph_format.space_after = Pt(2)

    # Screenshot 1
    doc.add_picture("LAB2/screenshots/SS1_jira_backlog.png", width=Inches(6.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c1 = doc.add_paragraph()
    r_c1 = p_c1.add_run("Figure 1.1: Jira Backlog View — 3 Defined Epics, PHRC Sprint 1 Backlog Items, Priority Rankings & Story Points")
    r_c1.italic = True
    r_c1.font.size = Pt(8.5)
    r_c1.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c1.paragraph_format.space_after = Pt(4)

    # Screenshot 2
    doc.add_picture("LAB2/screenshots/SS1_create_story.png", width=Inches(6.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c2 = doc.add_paragraph()
    r_c2 = p_c2.add_run("Figure 1.2: Agile User Story Creation Modal — User Story Format (As a... I want... So that...) with Acceptance Criteria (PHRC-5)")
    r_c2.italic = True
    r_c2.font.size = Pt(8.5)
    r_c2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ========================== PAGE 2 ==========================
    doc.add_page_break()
    p_s2 = doc.add_paragraph()
    r_s2 = p_s2.add_run("PART 2: STORY POINT ESTIMATION & ACTIVE SPRINT EXECUTION")
    r_s2.bold = True
    r_s2.font.size = Pt(11)
    r_s2.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    p_s2.paragraph_format.space_before = Pt(4)
    p_s2.paragraph_format.space_after = Pt(4)

    doc.add_picture("LAB2/screenshots/SS2_jira_story_points.png", width=Inches(6.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c3 = doc.add_paragraph()
    r_c3 = p_c3.add_run("Figure 2.1: Work Item Details & Estimation Log — Planning Poker Consensus, Fibonacci Story Points (3 pts) & Time Tracking")
    r_c3.italic = True
    r_c3.font.size = Pt(8.5)
    r_c3.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_c3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c3.paragraph_format.space_after = Pt(8)

    doc.add_picture("LAB2/screenshots/SS3_jira_active_sprint.png", width=Inches(6.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c4 = doc.add_paragraph()
    r_c4 = p_c4.add_run("Figure 2.2: Jira Active Sprint Scrum Board — Task Progression across TO DO (2), IN PROGRESS (1), and DONE (5) Swimlanes")
    r_c4.italic = True
    r_c4.font.size = Pt(8.5)
    r_c4.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_c4.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ========================== PAGE 3 ==========================
    doc.add_page_break()
    p_s3 = doc.add_paragraph()
    r_s3 = p_s3.add_run("PART 3: SPRINT PERFORMANCE METRICS & BURNDOWN ANALYSIS")
    r_s3.bold = True
    r_s3.font.size = Pt(11)
    r_s3.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    p_s3.paragraph_format.space_before = Pt(4)
    p_s3.paragraph_format.space_after = Pt(2)

    # Metrics Table
    tb_kpi = doc.add_table(rows=1, cols=4)
    tb_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpis = [
        ("Planned Capacity", "18 Points", "8 User Stories"),
        ("Completed Work", "13 Points", "5 Stories DONE (72%)"),
        ("In Progress", "5 Points", "US-2.2 Scoped Retrieval"),
        ("Sprint Duration", "10 Work Days", "Nov 2 – Nov 13, 2026")
    ]
    for idx, (title, val, sub) in enumerate(kpis):
        cell = tb_kpi.cell(0, idx)
        set_cell_shading(cell, "F8FAFC")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{title}\n")
        r.font.size = Pt(8.5)
        r.bold = True
        r = p.add_run(f"{val}\n")
        r.font.size = Pt(11)
        r.bold = True
        r.font.color.rgb = RGBColor(0x0C, 0x66, 0xE4)
        r = p.add_run(sub)
        r.font.size = Pt(7.5)
        r.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    doc.add_picture("LAB2/screenshots/SS4_jira_burndown_chart.png", width=Inches(6.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c5 = doc.add_paragraph()
    r_c5 = p_c5.add_run("Figure 3.1: Jira Sprint Burndown Chart — Ideal Guideline (Grey Dashed) vs. Actual Remaining Story Points (Red Step Curve)")
    r_c5.italic = True
    r_c5.font.size = Pt(8.5)
    r_c5.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_c5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c5.paragraph_format.space_after = Pt(4)

    # Observations Callout Box
    tb_obs = doc.add_table(rows=1, cols=1)
    tb_obs.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_obs = tb_obs.cell(0, 0)
    c_obs.width = Inches(7.0)
    set_cell_shading(c_obs, "F8FAFC")
    set_cell_left_border(c_obs, "16A34A", sz="24")
    p_obs = c_obs.paragraphs[0]
    p_obs.paragraph_format.space_before = Pt(3)
    p_obs.paragraph_format.space_after = Pt(3)
    r = p_obs.add_run("Key Burndown Chart Insights & Velocity Analysis:\n")
    r.bold = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    obs_text = (
        "• Guideline Tracking: The actual remaining points curve tracked tightly against the linear guideline from Nov 2 to Nov 6 (burning 6 points).\n"
        "• Non-Working Days Plateau: A flat plateau occurred during the weekend non-working window (Nov 7–8), correctly modeling team availability.\n"
        "• Significant Step Burndown: A 5-point downward step occurred on Nov 9 upon peer-review approval and merge of US-3.2 (Audit Logging).\n"
        "• Capacity Verification: Demonstrates that committing to 18 points across 8 stories matched team velocity without sprint failure."
    )
    r = p_obs.add_run(obs_text)
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # ========================== PAGE 4 ==========================
    doc.add_page_break()
    p_s4 = doc.add_paragraph()
    r_s4 = p_s4.add_run("PART 4: POST-SPRINT REFLECTIONS & AGILE EVALUATION")
    r_s4.bold = True
    r_s4.font.size = Pt(11)
    r_s4.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    p_s4.paragraph_format.space_before = Pt(4)
    p_s4.paragraph_format.space_after = Pt(2)

    qa_list = [
        ("1. Did your estimations reflect the actual effort?",
         "Yes. Straightforward administrative and configuration tasks, such as defining time-bounded validity "
         "windows (US-1.2) and multi-channel patient notifications (US-3.1), required fewer story points (2 points each). "
         "In contrast, complex security and cryptographic authorization tasks, such as provider credential verification "
         "with digital signatures (US-2.3) and HIPAA-compliant tamper-evident audit logging with SHA-256 hash chaining "
         "(US-3.2), involved higher uncertainty and technical complexity, justifying 5 story points each."),
        
        ("2. Was your backlog well-prioritized?",
         "Yes. Core patient privacy controls, including granular record selection (US-1.1), real-time consent revocation "
         "(US-1.3), and tamper-evident audit logging (US-3.2), were assigned High priority because patient confidentiality, "
         "regulatory compliance, and access governance are fundamental prerequisites before clinical providers can interact "
         "with medical records. Secondary administrative features, such as multi-channel patient alert notifications (US-3.1), "
         "were prioritized as Medium as specified in the Lab 1 requirements."),

        ("3. How did your simulated sprint align with your plan?",
         "Sprint 1 focused on the end-to-end patient consent lifecycle, clinical provider access requests, and immutable audit "
         "logging, successfully completing 13 out of 18 planned story points across 5 user stories. High-risk components "
         "(consent granting, scoped retrieval, and revocation) were executed first to validate the architectural boundaries "
         "established in Lab 1. The remaining 5 story points (provider credential verification and notification alerting) "
         "were positioned for Sprint 2, maintaining alignment with our phased delivery roadmap."),

        ("4. What insights did the burndown chart give about your team's capacity?",
         "The burndown chart demonstrated a realistic and steady completion rate that closely tracked the ideal linear guideline. "
         "Work progressed methodically across the 10 working days, with a flat plateau during the weekend non-working days (Nov 7–8) "
         "reflecting accurate calendar constraints. A significant downward step of 5 story points occurred on Nov 9 upon the "
         "successful integration and peer review of the audit logging framework, confirming that the 18-point sprint capacity "
         "was well-calibrated to our team velocity without overcommitment.")
    ]

    for q, a in qa_list:
        tb_q = doc.add_table(rows=1, cols=1)
        tb_q.alignment = WD_TABLE_ALIGNMENT.CENTER
        cq = tb_q.cell(0, 0)
        cq.width = Inches(7.0)
        set_cell_shading(cq, "F8FAFC")
        set_cell_left_border(cq, "0C66E4", sz="24")
        pq = cq.paragraphs[0]
        pq.paragraph_format.space_before = Pt(3)
        pq.paragraph_format.space_after = Pt(3)
        rq = pq.add_run(f"{q}\n")
        rq.bold = True
        rq.font.size = Pt(9.5)
        rq.font.color.rgb = RGBColor(0x0C, 0x66, 0xE4)
        ra = pq.add_run(a)
        ra.font.size = Pt(9)
        ra.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # Sign-off Box at bottom
    tb_sign = doc.add_table(rows=1, cols=1)
    tb_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    csign = tb_sign.cell(0, 0)
    csign.width = Inches(7.0)
    set_cell_shading(csign, "F8FAFC")
    set_cell_left_border(csign, "475569", sz="24")
    psign = csign.paragraphs[0]
    psign.paragraph_format.space_before = Pt(2)
    psign.paragraph_format.space_after = Pt(2)
    r = psign.add_run("Laboratory Evaluation & Jira Demonstration Sign-off:\n")
    r.bold = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    sign_txt = (
        "• Jira Workspace: Patient Health Record Consent Gateway (PHRC)  |  Board: PHRC board\n"
        "• Verified Deliverables: Backlog with Epics (SS1) • Story Point Estimations (SS2) • Active Sprint Board (SS3) • Burndown Chart (SS4)\n"
        "• Evaluation Status: Complete and ready for instructor review & grading  |  Course: UE23CS301 • PES University"
    )
    r = psign.add_run(sign_txt)
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.save(docx_path)
    print(f"Built Beautified DOCX: {docx_path}")

# -----------------------------------------------------------------------------
# 3. VERIFY PDF PAGE COUNT
# -----------------------------------------------------------------------------
def verify_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    print(f"Verification: {pdf_path} has {page_count} pages.")
    assert page_count == 4, f"ERROR: Expected exactly 4 pages, but got {page_count}!"
    print("PAGE COUNT CHECK PASSED: Exactly 4 pages.")

# -----------------------------------------------------------------------------
# 4. PACKAGE SUBMISSION ZIP
# -----------------------------------------------------------------------------
def create_zip_submission():
    zip_filename = "PES2UG24CS307_Lab2.zip"
    zip_path = os.path.join(".", zip_filename)

    readme_content = f"""================================================================================
PES UNIVERSITY - DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING
UE23CS301: SOFTWARE ENGINEERING LABORATORY
LAB 2: AGILE BACKLOG CREATION & SPRINT SIMULATION IN JIRA
================================================================================

STUDENT INFORMATION:
  Name:    {STUDENT_NAME}
  SRN:     {STUDENT_SRN}
  Section: {STUDENT_SECTION}
  Course:  UE23CS301 (Software Engineering Lab)
  Domain:  Problem Statement #13 — Patient Health Record Consent Management System

GITHUB REPOSITORY:
  {GITHUB_URL}

DELIVERABLES INCLUDED:
  1. PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf (Beautified 4-page formal submission)
  2. PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx (Word Document version)
  3. screenshots/SS1_jira_backlog.png (Jira Backlog view with Epics and User Stories)
  4. screenshots/SS1_create_story.png (Jira User Story creation dialog)
  5. screenshots/SS2_jira_story_points.png (Jira Work Item details & Story Points log)
  6. screenshots/SS3_jira_active_sprint.png (Active Sprint Scrum/Kanban board)
  7. screenshots/SS4_jira_burndown_chart.png (Burndown Chart with guideline & actual curve)
  8. README.txt (Submission details and documentation)
  9. github_link.txt (GitHub repository link)

BACKLOG SUMMARY:
  • Epic 1: Manage Patient Consent (PHRC-EP1)
    - PHRC-5  US-1.1: Granular Record Selection (3 pts, High, Done)
    - PHRC-6  US-1.2: Time-Bounded Validity Window (2 pts, High, Done)
    - PHRC-7  US-1.3: Real-Time Consent Revocation (3 pts, High, Done)
  • Epic 2: Clinical Provider Access (PHRC-EP2)
    - PHRC-8  US-2.1: Consent Access Request (5 pts, High, Done)
    - PHRC-9  US-2.2: Scoped Record Retrieval (5 pts, High, In Progress)
    - PHRC-10 US-2.3: Provider Credential Verification (3 pts, Medium, To Do)
  • Epic 3: Security, Notifications & Audit (PHRC-EP3)
    - PHRC-11 US-3.1: Multi-Channel Patient Notifications (2 pts, Medium, To Do)
    - PHRC-12 US-3.2: Tamper-Evident Access Audit Trail (5 pts, High, Done)

SPRINT 1 METRICS:
  • Planned Points: 18 Story Points
  • Completed Points: 13 Story Points (5 User Stories completed)
  • In Progress: 5 Story Points (US-2.2 Scoped Record Retrieval)
  • Carried to Sprint 2: 5 Story Points (US-2.3 & US-3.1)
================================================================================
"""
    with open("LAB2/README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)

    with open("LAB2/github_link.txt", "w", encoding="utf-8") as f:
        f.write(f"{GITHUB_URL}\n")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write("LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf", "PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf")
        z.write("LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx", "PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx")
        z.write("LAB2/README.txt", "README.txt")
        z.write("LAB2/github_link.txt", "github_link.txt")
        for fname in os.listdir("LAB2/screenshots"):
            fpath = os.path.join("LAB2/screenshots", fname)
            z.write(fpath, f"screenshots/{fname}")

    print(f"Created ZIP: {zip_path} ({os.path.getsize(zip_path)} bytes)")

    shutil.copyfile(zip_path, "PES2UG24CS307_NavyakishoreDash_LAB2.zip")
    print("Duplicated ZIP to PES2UG24CS307_NavyakishoreDash_LAB2.zip")

if __name__ == "__main__":
    pdf_dest = "LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf"
    docx_dest = "LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx"

    build_pdf(pdf_dest)
    verify_pdf(pdf_dest)
    build_docx(docx_dest)

    shutil.copyfile(pdf_dest, "LAB2_Report_PES2UG24CS307.pdf")
    shutil.copyfile(docx_dest, "LAB2_Report_PES2UG24CS307.docx")
    shutil.copyfile(pdf_dest, "PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf")
    shutil.copyfile(docx_dest, "PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx")

    create_zip_submission()
    print("ALL BEAUTIFIED LAB 2 DELIVERABLES SUCCESSFULLY PRODUCED!")
