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

os.makedirs("LAB1", exist_ok=True)
os.makedirs("LAB1/diagrams", exist_ok=True)

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

def set_cell_margins(cell, top=50, bottom=50, left=70, right=70):
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
# 1. BUILD COMPREHENSIVE LAB 1 WORD DOCUMENT
# ==============================================================================
def create_lab1_docx(filename):
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
    r_t = p_title.add_run("LAB 1: REQUIREMENTS ENGINEERING & UML USE-CASE MODELLING")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(16)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_s = p_sub.add_run("Problem Statement #13: Patient Health Record Consent Management System")
    r_s.font.name = "Calibri"
    r_s.font.size = Pt(11)
    r_s.font.color.rgb = COLOR_SECONDARY

    # Metadata Table
    meta = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course Code", "UE23CS301"), ("Course Title", "Software Engineering")],
        [("Domain", "Healthcare & Telemedicine"), ("Problem Statement", "#13 Patient Consent Gateway")],
        [("Deliverables", "Requirements Table, Use-Case Diagram, Flow Spec"), ("Date", "November 2026")]
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
        set_cell_margins(c1, top=50, bottom=50, left=80, right=80)
        set_cell_margins(c2, top=50, bottom=50, left=80, right=80)
        
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
        p.paragraph_format.space_after = Pt(4)
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

    # Section 1: Problem Context & Overview
    add_h1("1. Problem Context & System Overview")
    add_body(
        "Modern healthcare requires dynamic, secure, and privacy-preserving data sharing across disparate clinics, diagnostic laboratories, "
        "and medical specialists. The Patient Health Record Consent Management System is a patient-centric electronic health data gateway "
        "empowering citizens to explicitly govern granular, time-bound access permissions to their electronic health history."
    )
    add_body("Target Stakeholders and Actors:", prefix="Key Stakeholders: ")
    add_bullet("Patient (Primary Actor): The owner of the health records who grants, modifies, monitors, and revokes access permissions.", "Patient: ")
    add_bullet("Clinic Administrator (Institutional Actor): Manages institutional onboarding, doctor licensing verification, and administrative settings.", "Clinic Administrator: ")
    add_bullet("Doctor / Healthcare Provider (Clinical Actor): Submits consent requests to review relevant patient medical records for consultations.", "Care Provider: ")
    add_bullet("Audit & Logging System (Secondary System): Records tamper-evident, append-only logs for statutory compliance (HIPAA, GDPR, ABDM).", "Audit Service: ")

    # Section 2: Complete Requirements Table
    add_h1("2. Complete Requirements Table (5 FRs & 2 NFRs)")
    add_body(
        "Below is the complete requirements specification detailing exactly 5 Functional Requirements (FR-001 to FR-005) "
        "and 2 Non-Functional Requirements (NFR-001 & NFR-002) with ID, Type, Description, Priority, Acceptance Criteria, and Rationale:"
    )

    req_data = [
        ("FR-001", "Functional (Granting)", "The system shall allow authenticated patients to grant granular, time-bound consent (e.g., 24-hour window) for specific diagnostic records (lab reports, prescriptions, scans) to verified clinic doctors.", "High", "Pass: Access grant issues token with expiry timestamp and scope; Fail: Clinic accesses records after consent expiration or accesses unselected record categories.", "Enforces patient data autonomy and statutory privacy mandates by replacing blanket access with granular, time-bound permissions."),
        ("FR-002", "Functional (Revocation)", "The system shall enable patients to revoke or modify active consent permissions in real time prior to scheduled expiration timestamp.", "High", "Pass: Immediately upon revocation confirmation, subsequent provider access requests return HTTP 403 Forbidden; Fail: Provider retrieves records after revocation event.", "Guarantees immediate patient control in case of clinic changes, dispute, or accidental over-permissioning."),
        ("FR-003", "Functional (Verification)", "The system shall allow clinic administrators to register institutional clinic credentials and verify practitioners' professional medical licenses before enabling them to request patient consent.", "Medium", "Pass: Only verified doctors associated with licensed clinics can be selected as consent recipients; Fail: Unverified practitioner receives patient medical data.", "Prevents unauthorized actors or impersonators from obtaining sensitive health records by enforcing institutional verification."),
        ("FR-004", "Functional (Notification)", "The system shall permit verified clinics to submit episode-based consent requests and deliver multi-channel push, SMS, and email notifications to the patient within 3 seconds.", "Medium", "Pass: Patient receives instant notification with requesting doctor identity, clinical purpose, and record scope; Fail: Patient not notified or request lacks scope details.", "Streamlines outpatient consultations and emergency intake workflows by allowing doctors to request necessary records in advance."),
        ("FR-005", "Functional (Audit Trail)", "The system shall automatically record an immutable, tamper-evident log entry for every consent event (grant, modify, revoke, expire) and data access transaction, viewable by the patient.", "High", "Pass: Audit entry captures timestamp, actor ID, action type, IP address, and resource accessed within 500ms; Fail: Data access occurs without a corresponding audit trail record.", "Provides end-to-end transparency, non-repudiation, and auditability for statutory compliance and patient verification."),
        ("NFR-001", "Security & Privacy", "All patient health records and consent tokens shall be encrypted in transit using TLS 1.3 and at rest using AES-256 with role-based access control (RBAC) and zero-knowledge consent verification.", "High", "Pass: Automated vulnerability scans confirm zero plaintext leaks in network traces or storage dumps; Fail: Plaintext data or weak cryptographic ciphers detected.", "Protects Protected Health Information (PHI) against interception, data breaches, and insider threats in accordance with national health privacy laws."),
        ("NFR-002", "Performance & Availability", "The consent verification gateway shall evaluate access tokens with a p99 latency < 200 ms and maintain 99.9% uptime under simulated peak loads of up to 5,000 concurrent access requests.", "High", "Pass: Benchmark load tests verify p99 response time <= 200ms with zero failed transactions over 60 minutes; Fail: Latency exceeds 200ms or error rate > 0.1%.", "Ensures rapid emergency and outpatient clinical access without creating operational bottlenecks during active medical consultations."),
    ]

    tbl_req = doc.add_table(rows=len(req_data)+1, cols=6)
    tbl_req.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_req)

    req_widths = [Inches(0.7), Inches(1.0), Inches(1.8), Inches(0.6), Inches(1.3), Inches(1.1)]
    req_headers = ["ID", "Type", "Description", "Priority", "Acceptance Criteria", "Rationale"]
    for c_idx, h in enumerate(req_headers):
        cell = tbl_req.cell(0, c_idx)
        cell.width = req_widths[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, (r_id, r_type, r_desc, r_prio, r_crit, r_rat) in enumerate(req_data):
        row_cells = tbl_req.cell(r_idx+1, 0), tbl_req.cell(r_idx+1, 1), tbl_req.cell(r_idx+1, 2), tbl_req.cell(r_idx+1, 3), tbl_req.cell(r_idx+1, 4), tbl_req.cell(r_idx+1, 5)
        for c_idx, cell in enumerate(row_cells):
            cell.width = req_widths[c_idx]
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
            p = cell.paragraphs[0]
            val = [r_id, r_type, r_desc, r_prio, r_crit, r_rat][c_idx]
            if c_idx in (0, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                r.bold = True
            else:
                r = p.add_run(val)
            r.font.size = Pt(7.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 3: UML Use-Case Diagram
    add_h1("3. UML Use-Case Diagram")
    add_body(
        "The UML Use-Case Diagram models all primary stakeholders, core use cases, and includes explicit <<include>> and <<extend>> relationships:"
    )

    diag_path = "LAB1/diagrams/patient_consent_use_case_diagram.png"
    if os.path.exists(diag_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(diag_path, width=Inches(6.4))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r_c = p_cap.add_run("Figure 1: UML Use-Case Diagram — Patient Health Record Consent Management System.")
        r_c.font.italic = True
        r_c.font.size = Pt(8.5)
        r_c.font.color.rgb = COLOR_MUTED

    add_body(
        "Key Stereotyped Relationships Shown in Diagram:\n"
        "• <<include>> Relationships:\n"
        "  - 'Grant Granular Time-Bound Consent' <<include>> 'Authenticate User Identity': Enforces multi-factor authentication before generating any legally binding consent grant.\n"
        "  - 'Revoke / Modify Active Consent' <<include>> 'Authenticate User Identity': Prevents unauthorized session tampering during permission updates.\n"
        "  - 'Access Patient Diagnostic Records' <<include>> 'Authenticate User Identity': Clinical practitioners must verify digital credentials prior to record retrieval.\n"
        "• <<extend>> Relationships:\n"
        "  - 'Set Emergency Access Override' <<extend>> 'Grant Granular Time-Bound Consent': Triggered conditionally when a patient is incapacitated during critical trauma episodes.\n"
        "  - 'Send Access Notification' <<extend>> 'Access Patient Diagnostic Records': Triggered automatically upon data retrieval to notify the patient across registered channels."
    )

    # Section 4: Use-Case Flow Specification
    add_h1("4. Use-Case Flow Specification (Core Use Case: UC-01)")
    add_body(
        "Below is the complete 1-page Use-Case Flow Specification for the core use case 'Grant Granular Time-Bound Consent':"
    )

    uc_meta = [
        ("Use Case ID", "UC-01"),
        ("Use Case Name", "Grant Granular Time-Bound Consent"),
        ("Primary Actor", "Patient"),
        ("Secondary Actors", "Doctor / Healthcare Provider, Audit & Logging System"),
        ("Description", "Enables an authenticated patient to select a verified doctor/clinic, choose specific health record categories, define a finite expiration window, and issue a cryptographic consent grant."),
        ("Preconditions", "1. Patient is registered, logged in, and identity-verified.\n2. Target doctor/clinic is registered and credential-verified in the institutional registry.\n3. Patient has existing diagnostic records linked to their identifier."),
        ("Postconditions", "1. A secure, time-stamped consent token with expiration timestamp is issued.\n2. Target doctor gains scoped read access to approved records until expiry.\n3. Immutable audit ledger entry is appended recording the grant transaction.")
    ]

    tbl_uc = doc.add_table(rows=len(uc_meta), cols=2)
    tbl_uc.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_uc)
    for r_idx, (k, v) in enumerate(uc_meta):
        c1, c2 = tbl_uc.cell(r_idx, 0), tbl_uc.cell(r_idx, 1)
        c1.width = Inches(1.8)
        c2.width = Inches(4.7)
        set_cell_background(c1, HEX_BG_LIGHT)
        set_cell_margins(c1, top=40, bottom=40, left=60, right=60)
        set_cell_margins(c2, top=40, bottom=40, left=60, right=60)
        p1 = c1.paragraphs[0]
        r = p1.add_run(k)
        r.bold = True
        r.font.size = Pt(8.5)
        p2 = c2.paragraphs[0]
        p2.add_run(v).font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_h2("Main Success Scenario (Basic Flow):")
    mss_steps = [
        "Patient navigates to the 'Consent Management' dashboard on the mobile/web health portal.",
        "System displays active consent permissions and an action button to 'Grant New Consent'.",
        "Patient clicks 'Grant New Consent'.",
        "System prompts patient to search for and select the target clinic or consulting doctor.",
        "Patient selects a verified healthcare provider (e.g., Dr. Sharma, Apollo Telehealth Clinic).",
        "System queries the patient's record store and displays available diagnostic categories (Lab Tests, Cardiology Reports, Prescriptions, Radiology Scans).",
        "Patient selects desired record categories (e.g., Lab Tests & Cardiology Reports only).",
        "Patient defines the time-bounded validity period (e.g., 24 hours starting immediately).",
        "Patient reviews summary (doctor identity, document scope, expiry timestamp) and clicks 'Confirm & Authorize'.",
        "System triggers secondary authentication: <<include>> Authenticate User Identity (biometric or OTP prompt).",
        "Patient completes biometric / OTP verification.",
        "System generates a digitally signed consent token, updates access control lists, and appends an immutable transaction entry to the audit log.",
        "System displays a confirmation receipt with expiration countdown timer and dispatches an SMS/push notification to both patient and clinic."
    ]
    for idx, s in enumerate(mss_steps, 1):
        add_body(f"{idx}. {s}")

    add_h2("Alternate Flows:")
    add_body(
        "• Alternate Flow A1 — Granular Document Exclusion:\n"
        "  - At Step 7, patient expands a record category and unchecks specific sensitive documents (e.g., excluding past psychiatric evaluations or genetic screenings).\n"
        "  - System recalculates permission mask to exclude unselected document IDs.\n"
        "  - Flow resumes at Step 8 with the restricted document scope.\n\n"
        "• Alternate Flow A2 — Pre-Authorized Emergency Trauma Override (<<extend>>):\n"
        "  - At Step 8, patient designates the grant as 'Emergency Trauma Override' with a strict auto-revoking 4-hour window.\n"
        "  - System tags the consent policy with high-priority clinical emergency flags.\n"
        "  - Flow resumes at Step 10."
    )

    add_h2("Exception Flows:")
    add_body(
        "• Exception Flow E1 — Provider Credential Invalidation:\n"
        "  - At Step 5, if the target doctor's institutional license has lapsed or is suspended.\n"
        "  - System halts the flow, displays: 'Selected provider license cannot be verified. Consent cannot be granted.', and prompts patient to select another verified clinic.\n\n"
        "• Exception Flow E2 — Multi-Factor Authentication Failure:\n"
        "  - At Step 11, if biometric or OTP verification fails after 3 attempts.\n"
        "  - System cancels the transaction, discards the draft policy, logs a security alert, and returns the patient to the main dashboard."
    )

    doc.save(filename)
    print(f"Saved Lab 1 docx: {filename}")

create_lab1_docx("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_Report.docx")
create_lab1_docx("LAB1/LAB1_Report.docx")
create_lab1_docx("LAB1_Report_PES2UG24CS307.docx")

# ==============================================================================
# 2. BUILD COMPREHENSIVE LAB 1 PDF REPORT (ReportLab)
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
        self.setFont("Helvetica", 8.5)
        self.setFillColor(colors.HexColor("#555555"))
        self.setStrokeColor(colors.HexColor("#D0D0D0"))
        self.setLineWidth(0.5)
        self.line(36, 36, letter[0] - 36, 36)
        
        self.drawString(36, 24, "PES University | Dept. of CSE | UE23CS301 SE Lab 1 Report — Requirements & Use Cases")
        self.drawRightString(letter[0] - 36, 24, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def build_lab1_pdf(pdf_path):
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

    tbl_hdr = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=7.5, leading=9.5, textColor=colors.white, alignment=1)
    tbl_cell = ParagraphStyle('TC', fontName='Helvetica', fontSize=7, leading=8.5, textColor=colors.HexColor('#212121'))
    tbl_cell_bold = ParagraphStyle('TCB', fontName='Helvetica-Bold', fontSize=7, leading=8.5, textColor=colors.HexColor('#1A237E'))
    tbl_cell_center = ParagraphStyle('TCC', fontName='Helvetica', fontSize=7, leading=8.5, textColor=colors.HexColor('#212121'), alignment=1)

    story = []

    # =========================================================================
    # PAGE 1: Header, Metadata, Context, Requirements Table
    # =========================================================================
    story.append(Paragraph("PES UNIVERSITY", inst_style))
    story.append(Paragraph("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING", sub_style))
    story.append(Paragraph("LAB 1: REQUIREMENTS ENGINEERING & UML USE-CASE MODELLING", title_style))
    story.append(Paragraph("Problem Statement #13: Patient Health Record Consent Management System", doc_subtitle))
    story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor('#1A237E'), spaceAfter=6))

    # Meta Table
    m_data = [
        [Paragraph("<b>Student Name:</b> Navyakishore Dash", tbl_cell), Paragraph("<b>SRN:</b> PES2UG24CS307", tbl_cell)],
        [Paragraph("<b>Course Code:</b> UE23CS301", tbl_cell), Paragraph("<b>Course:</b> Software Engineering Laboratory", tbl_cell)],
        [Paragraph("<b>Problem Domain:</b> Healthcare & Telemedicine", tbl_cell), Paragraph("<b>Target Actors:</b> Patient, Clinic Administrator, Doctor", tbl_cell)],
        [Paragraph("<b>Repository:</b> NavyakishoreDash/L6_L7_58_restaurant_bill", tbl_cell), Paragraph("<b>Submission Date:</b> November 2026", tbl_cell)],
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
    story.append(Paragraph("1. Problem Context & Overview", h1_style))
    story.append(Paragraph(
        "A patient-centric electronic health data gateway where patients explicitly manage granular, time-bound consent permissions "
        "for clinics, diagnostic labs, and consulting doctors to access their medical history. The system enforces strict compliance with "
        "healthcare privacy standards (HIPAA, GDPR, ABDM) by guaranteeing patient data autonomy, ephemeral token validity, and an immutable audit trail.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Target Stakeholders & Actors:</b> <i>Patient</i> (data owner), <i>Clinic Administrator</i> (institutional credentials verifier), "
        "<i>Doctor / Healthcare Provider</i> (clinical consumer), and <i>Audit Logging System</i> (statutory ledger).",
        body_style
    ))

    # Section 2: Requirements Table
    story.append(Paragraph("2. Complete Requirements Table (Exactly 5 FRs & 2 NFRs)", h1_style))

    req_data = [
        [Paragraph("ID", tbl_hdr), Paragraph("Type", tbl_hdr), Paragraph("Description", tbl_hdr), Paragraph("Pri.", tbl_hdr), Paragraph("Acceptance Criteria", tbl_hdr), Paragraph("Rationale", tbl_hdr)],
        [
            Paragraph("<b>FR-001</b>", tbl_cell_bold),
            Paragraph("Functional (Granting)", tbl_cell),
            Paragraph("The system shall allow patients to grant granular, time-bounded (e.g., 24-hour) access permissions for specific diagnostic records (lab reports, prescriptions, scans) to verified clinic doctors.", tbl_cell),
            Paragraph("High", tbl_cell_center),
            Paragraph("<b>Pass:</b> Access grant generates token with expiry timestamp and approved scope; <b>Fail:</b> Clinic accesses records after expiration or unselected categories.", tbl_cell),
            Paragraph("Enforces patient data autonomy and regulatory compliance by eliminating blanket access.", tbl_cell)
        ],
        [
            Paragraph("<b>FR-002</b>", tbl_cell_bold),
            Paragraph("Functional (Revocation)", tbl_cell),
            Paragraph("The system shall allow patients to revoke or modify active consent permissions in real-time prior to the scheduled expiration timestamp.", tbl_cell),
            Paragraph("High", tbl_cell_center),
            Paragraph("<b>Pass:</b> Subsequent access attempts return HTTP 403 Forbidden immediately upon revocation; <b>Fail:</b> Provider reads records post-revocation.", tbl_cell),
            Paragraph("Empowers patients to immediately terminate access in case of dispute or clinic changes.", tbl_cell)
        ],
        [
            Paragraph("<b>FR-003</b>", tbl_cell_bold),
            Paragraph("Functional (Verification)", tbl_cell),
            Paragraph("The system shall enable clinic administrators to register institutional clinic credentials and verify doctors' medical licenses before allowing them to request patient consent.", tbl_cell),
            Paragraph("Med.", tbl_cell_center),
            Paragraph("<b>Pass:</b> Only verified doctors associated with licensed clinics can receive consent; <b>Fail:</b> Unverified practitioner receives patient data.", tbl_cell),
            Paragraph("Prevents unauthorized actors or impersonators from accessing protected health information.", tbl_cell)
        ],
        [
            Paragraph("<b>FR-004</b>", tbl_cell_bold),
            Paragraph("Functional (Notification)", tbl_cell),
            Paragraph("The system shall permit verified clinics to submit episode-based consent requests and deliver multi-channel push, SMS, and email notifications to the patient within 3 seconds.", tbl_cell),
            Paragraph("Med.", tbl_cell_center),
            Paragraph("<b>Pass:</b> Patient receives notification with requesting clinic details, record scope, and purpose within 3s; <b>Fail:</b> Notification delayed or missing scope.", tbl_cell),
            Paragraph("Streamlines outpatient consultations and telemedicine workflows through advance consent requests.", tbl_cell)
        ],
        [
            Paragraph("<b>FR-005</b>", tbl_cell_bold),
            Paragraph("Functional (Audit Trail)", tbl_cell),
            Paragraph("The system shall automatically record an immutable, tamper-evident log entry for every consent event (grant, modify, revoke, expire) and data access transaction, viewable by the patient.", tbl_cell),
            Paragraph("High", tbl_cell_center),
            Paragraph("<b>Pass:</b> Audit entry captures timestamp, actor ID, action type, IP address, and record ID within 500ms; <b>Fail:</b> Data access occurs without audit log entry.", tbl_cell),
            Paragraph("Guarantees non-repudiation, transparency, and traceability for patient peace of mind and statutory audits.", tbl_cell)
        ],
        [
            Paragraph("<b>NFR-001</b>", tbl_cell_bold),
            Paragraph("Security & Privacy", tbl_cell),
            Paragraph("All patient health records and consent tokens shall be encrypted in transit using TLS 1.3 and at rest using AES-256 with role-based access control (RBAC) and zero-knowledge consent verification.", tbl_cell),
            Paragraph("High", tbl_cell_center),
            Paragraph("<b>Pass:</b> Penetration tests and automated scans confirm zero plaintext leaks in transit or at rest; <b>Fail:</b> Plaintext data or weak ciphers detected.", tbl_cell),
            Paragraph("Protects Protected Health Information (PHI) against interception, data breaches, and insider threats.", tbl_cell)
        ],
        [
            Paragraph("<b>NFR-002</b>", tbl_cell_bold),
            Paragraph("Performance & Reliability", tbl_cell),
            Paragraph("The consent verification gateway shall evaluate access tokens with a p99 latency < 200 ms and maintain 99.9% uptime under simulated peak loads of up to 5,000 concurrent requests.", tbl_cell),
            Paragraph("High", tbl_cell_center),
            Paragraph("<b>Pass:</b> Load tests confirm p99 response time <= 200ms with zero failed transactions over 60 mins; <b>Fail:</b> Latency exceeds 200ms or error rate > 0.1%.", tbl_cell),
            Paragraph("Ensures emergency and clinical data retrieval operates without latency bottlenecks during consultations.", tbl_cell)
        ],
    ]
    t_req = Table(req_data, colWidths=[42, 68, 140, 28, 138, 124])
    t_req.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A237E')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_req)

    story.append(PageBreak())

    # =========================================================================
    # PAGE 2: UML Use-Case Diagram
    # =========================================================================
    story.append(Paragraph("3. UML Use-Case Diagram", h1_style))
    story.append(Paragraph(
        "The diagram models all primary and secondary stakeholders, primary use cases, system boundaries, and includes explicit "
        "<code>&lt;&lt;include&gt;&gt;</code> and <code>&lt;&lt;extend&gt;&gt;</code> relationships:",
        body_style
    ))

    diag_path = "LAB1/diagrams/patient_consent_use_case_diagram.png"
    if os.path.exists(diag_path):
        story.append(RLImage(diag_path, width=540, height=395))
        story.append(Paragraph("Figure 1: UML Use-Case Diagram for Patient Health Record Consent Management System.", caption_style))

    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "<b>Modelled Relationships:</b><br/>"
        "• <b>&lt;&lt;include&gt;&gt; Relationships:</b><br/>"
        "  - <i>'Grant Granular Time-Bound Consent'</i> &lt;&lt;include&gt;&gt; <i>'Authenticate User Identity'</i>: Enforces mandatory biometric / MFA authentication before issuing any legally binding consent token.<br/>"
        "  - <i>'Revoke / Modify Active Consent'</i> &lt;&lt;include&gt;&gt; <i>'Authenticate User Identity'</i>: Ensures only the authenticated patient can revoke or alter permissions.<br/>"
        "  - <i>'Access Patient Diagnostic Records'</i> &lt;&lt;include&gt;&gt; <i>'Authenticate User Identity'</i>: Medical practitioners must be authenticated before record retrieval.<br/>"
        "• <b>&lt;&lt;extend&gt;&gt; Relationships:</b><br/>"
        "  - <i>'Set Emergency Access Override'</i> &lt;&lt;extend&gt;&gt; <i>'Grant Granular Time-Bound Consent'</i>: Extends the base grant flow conditionally during critical trauma/emergency episodes.<br/>"
        "  - <i>'Send Access Notification'</i> &lt;&lt;extend&gt;&gt; <i>'Access Patient Diagnostic Records'</i>: Extends the data access flow by dispatching multi-channel alerts to the patient.",
        body_style
    ))

    story.append(PageBreak())

    # =========================================================================
    # PAGE 3: Use-Case Flow Specification (Core Use Case: UC-01)
    # =========================================================================
    story.append(Paragraph("4. Use-Case Flow Specification", h1_style))
    story.append(Paragraph("<b>Core Use Case: UC-01 — Grant Granular Time-Bound Consent</b>", h2_style))

    uc_meta = [
        [Paragraph("<b>Use Case ID:</b>", tbl_cell_bold), Paragraph("UC-01", tbl_cell), Paragraph("<b>Use Case Name:</b>", tbl_cell_bold), Paragraph("Grant Granular Time-Bound Consent", tbl_cell)],
        [Paragraph("<b>Primary Actor:</b>", tbl_cell_bold), Paragraph("Patient", tbl_cell), Paragraph("<b>Secondary Actors:</b>", tbl_cell_bold), Paragraph("Doctor / Healthcare Provider, Audit System", tbl_cell)],
        [Paragraph("<b>Description:</b>", tbl_cell_bold), Paragraph("Allows authenticated patients to select a verified doctor/clinic, pick specific medical document categories (lab reports, prescriptions, scans), define an expiration timeframe (e.g. 24h), and issue an authorized consent grant.", tbl_cell), Paragraph("", tbl_cell), Paragraph("", tbl_cell)],
        [Paragraph("<b>Preconditions:</b>", tbl_cell_bold), Paragraph("1. Patient is registered and authenticated with MFA.<br/>2. Target doctor/clinic is registered and institutional credentials are verified.<br/>3. Patient has existing diagnostic records linked to their identifier.", tbl_cell), Paragraph("", tbl_cell), Paragraph("", tbl_cell)],
        [Paragraph("<b>Postconditions:</b>", tbl_cell_bold), Paragraph("1. A cryptographic, time-stamped consent token is generated.<br/>2. Target doctor receives scoped read access to approved records until expiration.<br/>3. An immutable entry is appended to the audit ledger.", tbl_cell), Paragraph("", tbl_cell), Paragraph("", tbl_cell)],
    ]
    t_uc = Table(uc_meta, colWidths=[90, 180, 100, 170])
    t_uc.setStyle(TableStyle([
        ('SPAN', (1,2), (3,2)),
        ('SPAN', (1,3), (3,3)),
        ('SPAN', (1,4), (3,4)),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F4F6F9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_uc)
    story.append(Spacer(1, 4))

    story.append(Paragraph("<b>Main Success Scenario (Basic Flow):</b>", h2_style))
    steps = [
        "Patient navigates to the 'Consent Management' dashboard on the mobile/web health gateway.",
        "System displays active consent permissions and an option to 'Grant New Consent'.",
        "Patient clicks 'Grant New Consent'.",
        "System prompts the patient to search for and select the target clinic or consulting doctor.",
        "Patient selects a verified healthcare provider (e.g., Dr. Sharma, Apollo Telehealth Clinic).",
        "System queries the patient's record catalog and presents record categories (Lab Reports, Cardiology Reports, Prescriptions, Radiology Scans).",
        "Patient selects desired record categories (e.g., Lab Reports & Cardiology Reports only).",
        "Patient specifies the time-bounded validity window (e.g., 24 hours, starting immediately).",
        "Patient reviews summary (doctor identity, document scope, expiry timestamp) and clicks 'Authorize & Issue'.",
        "System prompts patient for identity confirmation: <code>&lt;&lt;include&gt;&gt; Authenticate User Identity</code> (biometric / OTP prompt).",
        "Patient completes biometric / OTP verification successfully.",
        "System generates a digitally signed consent token, sets expiry timestamp, and appends an immutable transaction record to the audit log.",
        "System displays a confirmation receipt with expiration countdown timer and dispatches an instant notification to both patient and clinic."
    ]
    for idx, s in enumerate(steps, 1):
        story.append(Paragraph(f"<b>{idx}.</b> {s}", bullet_style))

    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>Alternate Flows:</b>", h2_style))
    story.append(Paragraph(
        "• <b>Alternate Flow A1 — Granular Document Exclusion:</b><br/>"
        "  - At Step 7, patient expands a record category and unchecks specific sensitive documents (e.g., excluding past psychiatric evaluations or genetic screenings).<br/>"
        "  - System updates the permission mask to exclude the unselected documents.<br/>"
        "  - Flow resumes at Step 8 with the restricted document scope.<br/>"
        "• <b>Alternate Flow A2 — Pre-Authorized Emergency Trauma Override (<code>&lt;&lt;extend&gt;&gt;</code>):</b><br/>"
        "  - At Step 8, patient designates the grant as 'Emergency Trauma Override' with an auto-revoking 4-hour window.<br/>"
        "  - System tags the consent policy with high-priority clinical emergency flags.<br/>"
        "  - Flow resumes at Step 10.",
        body_style
    ))

    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Exception Flows:</b>", h2_style))
    story.append(Paragraph(
        "• <b>Exception Flow E1 — Provider Credential Invalidation:</b><br/>"
        "  - At Step 5, if the target doctor's institutional license has lapsed or is suspended, system halts the flow, displays: <i>'Selected provider license cannot be verified. Consent cannot be granted.'</i>, and prompts patient to select another verified clinic.<br/>"
        "• <b>Exception Flow E2 — Multi-Factor Authentication Failure:</b><br/>"
        "  - At Step 11, if biometric or OTP verification fails after 3 consecutive attempts, system aborts the transaction, discards the draft policy, logs a security alert, and returns the patient to the main dashboard.",
        body_style
    ))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Generated comprehensive Lab 1 PDF: {pdf_path}")

# Build in both LAB1 and root directory
build_lab1_pdf("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_Report.pdf")
build_lab1_pdf("LAB1/LAB1_Report.pdf")
build_lab1_pdf("LAB1_Report_PES2UG24CS307.pdf")
build_lab1_pdf("PES2UG24CS307_NavyakishoreDash_LAB1_Report.pdf")

# ==============================================================================
# 3. BUILD STANDALONE 1-PAGE USE-CASE SPECIFICATION (PDF & DOCX)
# ==============================================================================
def create_standalone_usecase_spec_docx(filename):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_inst.add_run("PES UNIVERSITY — DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING\nUE23CS301 SE Lab 1 | Use-Case Flow Specification: UC-01")
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    r_t = p_title.add_run("Use-Case Specification: Grant Granular Time-Bound Consent")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(12)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_SECONDARY

    uc_meta = [
        ("Use Case ID & Name", "UC-01: Grant Granular Time-Bound Consent"),
        ("Primary Actor", "Patient (Data Owner)"),
        ("Secondary Actors", "Doctor / Healthcare Provider, Audit & Logging System"),
        ("Preconditions", "1. Patient is logged in with multi-factor authentication (MFA).\n2. Target doctor/clinic is registered and verified in institutional directory.\n3. Patient has existing diagnostic records linked to their identifier."),
        ("Postconditions", "1. Time-bounded cryptographic consent token is generated.\n2. Target doctor receives scoped read access to approved records until expiration.\n3. Immutable audit trail entry is appended recording the grant event.")
    ]
    tbl = doc.add_table(rows=len(uc_meta), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)
    for r_idx, (k, v) in enumerate(uc_meta):
        c1, c2 = tbl.cell(r_idx, 0), tbl.cell(r_idx, 1)
        c1.width = Inches(1.8)
        c2.width = Inches(5.4)
        set_cell_background(c1, HEX_BG_LIGHT)
        set_cell_margins(c1, top=30, bottom=30, left=60, right=60)
        set_cell_margins(c2, top=30, bottom=30, left=60, right=60)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r = p1.add_run(k)
        r.bold = True
        r.font.size = Pt(8.5)
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        p2.add_run(v).font.size = Pt(8.5)

    def add_sub(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY

    def add_step(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(8.5)

    add_sub("Main Success Scenario (Basic Flow):")
    steps = [
        "1. Patient navigates to the 'Consent Management' dashboard on the mobile/web health portal.",
        "2. System displays active consent permissions and an action button to 'Grant New Consent'.",
        "3. Patient clicks 'Grant New Consent'.",
        "4. System prompts patient to search for and select the target clinic or consulting doctor.",
        "5. Patient selects a verified healthcare provider (e.g., Dr. Sharma, Apollo Telehealth Clinic).",
        "6. System queries the patient's record catalog and presents record categories (Lab Reports, Cardiology, Prescriptions, Radiology).",
        "7. Patient selects desired record categories (e.g., Lab Reports & Cardiology Reports only).",
        "8. Patient specifies the time-bounded validity window (e.g., 24 hours starting immediately).",
        "9. Patient reviews summary (doctor identity, document scope, expiry timestamp) and clicks 'Authorize & Issue'.",
        "10. System prompts patient for secondary authentication: <<include>> Authenticate User Identity (biometric/OTP).",
        "11. Patient completes biometric / OTP verification successfully.",
        "12. System generates a digitally signed consent token, sets expiry timestamp, and appends an immutable entry to the audit log.",
        "13. System displays confirmation receipt with expiration countdown and dispatches real-time SMS/push alerts to patient and clinic."
    ]
    for s in steps:
        add_step(s)

    add_sub("Alternate Flows:")
    add_step("• A1 (Granular Document Exclusion): At Step 7, patient expands a category and unchecks specific sensitive documents (e.g. psychiatric evaluations). System updates permission mask to exclude those IDs. Flow resumes at Step 8.")
    add_step("• A2 (Pre-Authorized Emergency Override <<extend>>): At Step 8, patient designates the grant as 'Emergency Trauma Override' with a strict auto-revoking 4-hour window. Flow resumes at Step 10.")

    add_sub("Exception Flows:")
    add_step("• E1 (Provider Credential Invalidation): At Step 5, if doctor's medical license is suspended/unverified, system halts flow, displays warning: 'Selected provider license cannot be verified.', and prompts patient to select another verified clinic.")
    add_step("• E2 (MFA Authentication Failure): At Step 11, if biometric/OTP verification fails after 3 attempts, system cancels the transaction, discards draft policy, logs a security alert, and returns to dashboard.")

    doc.save(filename)
    print(f"Saved standalone usecase spec docx: {filename}")

create_standalone_usecase_spec_docx("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_UseCase_Spec.docx")
create_standalone_usecase_spec_docx("LAB1/LAB1_UseCase_Spec.docx")

# ==============================================================================
# 4. PACKAGE LAB 1 SUBMISSION ZIP
# ==============================================================================
staging_dir = "LAB1_Submission_Staging"
if os.path.exists(staging_dir):
    shutil.rmtree(staging_dir)
os.makedirs(staging_dir, exist_ok=True)
os.makedirs(os.path.join(staging_dir, "diagrams"), exist_ok=True)

# Copy diagrams
shutil.copy2("LAB1/diagrams/patient_consent_use_case_diagram.png", os.path.join(staging_dir, "diagrams/patient_consent_use_case_diagram.png"))
shutil.copy2("LAB1/diagrams/patient_consent_use_case_diagram.pdf", os.path.join(staging_dir, "diagrams/patient_consent_use_case_diagram.pdf"))

# Copy reports
shutil.copy2("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_Report.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB1_Report.docx"))
shutil.copy2("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_Report.pdf", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB1_Report.pdf"))
shutil.copy2("LAB1/PES2UG24CS307_NavyakishoreDash_LAB1_UseCase_Spec.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB1_UseCase_Spec.docx"))

# Create github_link.txt
github_txt = """========================================================================
LAB 1: REQUIREMENTS ENGINEERING & UML USE-CASE MODELLING
========================================================================

STUDENT DETAILS:
Name: Navyakishore Dash
SRN: PES2UG24CS307
Course: UE23CS301 - Software Engineering
Department: Department of Computer Science and Engineering
Institution: PES University

PROBLEM STATEMENT #13:
Healthcare & Telemedicine: Patient Health Record Consent Management System

GITHUB REPOSITORY LINK:
https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill
(Directory: LAB1)

========================================================================
LAB 1 DELIVERABLES SUMMARY
========================================================================
1. Complete Requirements Table:
   - Exactly 5 Functional Requirements (FR-001 to FR-005)
   - 2 Non-Functional Requirements (NFR-001 & NFR-002)
   - Formatted with ID, Type, Description, Priority, Acceptance Criteria, Rationale.

2. UML Use-Case Diagram:
   - diagrams/patient_consent_use_case_diagram.png (High-Res 300 DPI)
   - diagrams/patient_consent_use_case_diagram.pdf (Vector PDF)
   - Features: All primary actors (Patient, Clinic Admin, Doctor, Audit System), core use cases,
     explicit <<include>> relationships, and <<extend>> relationships.

3. Use-Case Flow Specification (Core Use Case UC-01):
   - Preconditions, Postconditions, Main Success Scenario (13 steps),
     Alternate Flows (A1 Granular Exclusion, A2 Emergency Trauma Override),
     Exception Flows (E1 Invalid License, E2 Authentication Timeout).

4. Lab 1 Reports:
   - PES2UG24CS307_NavyakishoreDash_LAB1_Report.pdf (Full Academic PDF Report)
   - PES2UG24CS307_NavyakishoreDash_LAB1_Report.docx (Full Academic Word Document)
   - PES2UG24CS307_NavyakishoreDash_LAB1_UseCase_Spec.docx (Dedicated Use-Case Specification)
"""

with open(os.path.join(staging_dir, "README.txt"), "w", encoding="utf-8") as f:
    f.write(github_txt)

with open(os.path.join(staging_dir, "github_link.txt"), "w", encoding="utf-8") as f:
    f.write(github_txt)

# Zip files
zips = [
    "PES2UG24CS307_Lab1.zip",
    "PES2UG24CS307_NavyakishoreDash_LAB1.zip"
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
print("Lab 1 build and packaging completed successfully!")
