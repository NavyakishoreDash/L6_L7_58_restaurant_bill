import os
import shutil
import zipfile
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
import fitz

os.makedirs("LAB2", exist_ok=True)
os.makedirs("LAB2/screenshots", exist_ok=True)

STUDENT_NAME = "Navyakishore Dash"
STUDENT_SRN = "PES2UG24CS307"
STUDENT_SECTION = "Section -E"
SCENARIO = "Problem Statement #13: Patient Health Record Consent Management System"
GITHUB_URL = "https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill"

# -----------------------------------------------------------------------------
# 1. BUILD STRICT 4-PAGE PDF REPORT (REPORTLAB)
# -----------------------------------------------------------------------------
def build_pdf(pdf_path):
    # Letter size: 612 x 792 pt, matching reference PDF
    # Reference bboxes: left=73.5 pt, right=541.5 pt (margin = 72 pt / 1 inch)
    # Available width = 612 - 72*2 = 468 pt
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=72,
        rightMargin=72,
        topMargin=54,
        bottomMargin=54
    )

    styles = getSampleStyleSheet()

    # Define custom styles
    h_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#111827')
    )

    lbl_style = ParagraphStyle(
        'LabelStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#1F2937'),
        spaceAfter=8
    )

    ref_title_style = ParagraphStyle(
        'RefTitleStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=colors.HexColor('#111827'),
        spaceAfter=14
    )

    q_style = ParagraphStyle(
        'QuestionStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14.5,
        textColor=colors.HexColor('#1E293B'),
        spaceBefore=10,
        spaceAfter=4
    )

    ans_style = ParagraphStyle(
        'AnswerStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        textColor=colors.HexColor('#334155'),
        alignment=4,  # Justified
        spaceAfter=10
    )

    story = []

    # ========================== PAGE 1 ==========================
    # Student Header
    story.append(Paragraph(f"Name: {STUDENT_NAME}", h_style))
    story.append(Paragraph(f"SRN: {STUDENT_SRN}", h_style))
    story.append(Paragraph(f"{STUDENT_SECTION}", h_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Screenshots:", lbl_style))
    story.append(Spacer(1, 4))

    # Image 1: Backlog view (width 468 pt, height 260 pt)
    im1_path = "LAB2/screenshots/SS1_jira_backlog.png"
    story.append(RLImage(im1_path, width=468, height=255))
    story.append(Spacer(1, 10))

    # Image 2: Create Story Modal (width 468 pt, height 160 pt)
    im2_path = "LAB2/screenshots/SS1_create_story.png"
    story.append(RLImage(im2_path, width=468, height=155))
    story.append(PageBreak())

    # ========================== PAGE 2 ==========================
    story.append(Spacer(1, 10))
    # Image 3: Story Points Details & Poker (width 468 pt, height 260 pt)
    im3_path = "LAB2/screenshots/SS2_jira_story_points.png"
    story.append(RLImage(im3_path, width=468, height=260))
    story.append(Spacer(1, 14))

    # Image 4: Active Sprint Board (width 468 pt, height 260 pt)
    im4_path = "LAB2/screenshots/SS3_jira_active_sprint.png"
    story.append(RLImage(im4_path, width=468, height=260))
    story.append(PageBreak())

    # ========================== PAGE 3 ==========================
    story.append(Spacer(1, 20))
    # Image 5: Jira Burndown Chart (width 468 pt, height 280 pt)
    im5_path = "LAB2/screenshots/SS4_jira_burndown_chart.png"
    story.append(RLImage(im5_path, width=468, height=275))
    story.append(PageBreak())

    # ========================== PAGE 4 ==========================
    story.append(Paragraph("Reflections:", ref_title_style))
    story.append(Spacer(1, 6))

    # Q1
    q1_text = "1. Did your estimations reflect the actual effort?"
    a1_text = (
        "Yes. Straightforward administrative and configuration tasks, such as defining time-bounded validity "
        "windows (US-1.2) and multi-channel patient notifications (US-3.1), required fewer story points (2 points each). "
        "In contrast, complex security and cryptographic authorization tasks, such as provider credential verification "
        "with digital signatures (US-2.3) and HIPAA-compliant tamper-evident audit logging with SHA-256 hash chaining "
        "(US-3.2), involved higher uncertainty and technical complexity, justifying 5 story points each."
    )
    story.append(Paragraph(q1_text, q_style))
    story.append(Paragraph(a1_text, ans_style))

    # Q2
    q2_text = "2. Was your backlog well-prioritized?"
    a2_text = (
        "Yes. Core patient privacy controls, including granular record selection (US-1.1), real-time consent revocation "
        "(US-1.3), and tamper-evident audit logging (US-3.2), were assigned High priority because patient confidentiality, "
        "regulatory compliance, and access governance are fundamental prerequisites before clinical providers can interact "
        "with medical records. Secondary administrative features, such as multi-channel patient alert notifications (US-3.1), "
        "were prioritized as Medium as specified in the Lab 1 requirements."
    )
    story.append(Paragraph(q2_text, q_style))
    story.append(Paragraph(a2_text, ans_style))

    # Q3
    q3_text = "3. How did your simulated sprint align with your plan?"
    a3_text = (
        "Sprint 1 focused on the end-to-end patient consent lifecycle, clinical provider access requests, and immutable audit "
        "logging, successfully completing 13 out of 18 planned story points across 5 user stories. High-risk components "
        "(consent granting, scoped retrieval, and revocation) were executed first to validate the architectural boundaries "
        "established in Lab 1. The remaining 5 story points (provider credential verification and notification alerting) "
        "were positioned for Sprint 2, maintaining alignment with our phased delivery roadmap."
    )
    story.append(Paragraph(q3_text, q_style))
    story.append(Paragraph(a3_text, ans_style))

    # Q4
    q4_text = "4. What insights did the burndown chart give about your team's capacity?"
    a4_text = (
        "The burndown chart demonstrated a realistic and steady completion rate that closely tracked the ideal linear guideline. "
        "Work progressed methodically across the 10 working days, with a flat plateau during the weekend non-working days (Nov 7–8) "
        "reflecting accurate calendar constraints. A significant downward step of 5 story points occurred on Nov 9 upon the "
        "successful integration and peer review of the audit logging framework, confirming that the 18-point sprint capacity "
        "was well-calibrated to our team velocity without overcommitment."
    )
    story.append(Paragraph(q4_text, q_style))
    story.append(Paragraph(a4_text, ans_style))

    doc.build(story)
    print(f"Built PDF: {pdf_path}")

# -----------------------------------------------------------------------------
# 2. BUILD FORMATTED WORD DOCUMENT REPORT (.DOCX)
# -----------------------------------------------------------------------------
def build_docx(docx_path):
    doc = docx.Document()
    
    # Set standard margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    # Style definitions
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # PAGE 1: Header + Screenshots 1 & 2
    p_header = doc.add_paragraph()
    r = p_header.add_run(f"Name: {STUDENT_NAME}\n")
    r.bold = True
    r.font.size = Pt(12)
    r = p_header.add_run(f"SRN: {STUDENT_SRN}\n")
    r.bold = True
    r.font.size = Pt(12)
    r = p_header.add_run(f"{STUDENT_SECTION}\n\n")
    r.bold = True
    r.font.size = Pt(12)
    r = p_header.add_run("Screenshots:\n")
    r.bold = True
    r.font.size = Pt(12)
    p_header.paragraph_format.space_after = Pt(4)

    # Image 1 (Backlog)
    doc.add_picture("LAB2/screenshots/SS1_jira_backlog.png", width=Inches(6.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.space_after = Pt(8)

    # Image 2 (Create Story Modal)
    doc.add_picture("LAB2/screenshots/SS1_create_story.png", width=Inches(6.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # PAGE 2: Story Points & Active Sprint
    doc.add_page_break()
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(8)
    doc.add_picture("LAB2/screenshots/SS2_jira_story_points.png", width=Inches(6.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.space_after = Pt(12)

    doc.add_picture("LAB2/screenshots/SS3_jira_active_sprint.png", width=Inches(6.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # PAGE 3: Burndown Chart
    doc.add_page_break()
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(16)
    doc.add_picture("LAB2/screenshots/SS4_jira_burndown_chart.png", width=Inches(6.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # PAGE 4: Reflections
    doc.add_page_break()
    p_ref = doc.add_paragraph()
    r_ref = p_ref.add_run("Reflections:\n")
    r_ref.bold = True
    r_ref.font.size = Pt(14)
    r_ref.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
    p_ref.paragraph_format.space_after = Pt(10)

    qa_pairs = [
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

    for q, a in qa_pairs:
        pq = doc.add_paragraph()
        rq = pq.add_run(q)
        rq.bold = True
        rq.font.size = Pt(11)
        rq.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        pq.paragraph_format.space_before = Pt(8)
        pq.paragraph_format.space_after = Pt(2)

        pa = doc.add_paragraph()
        ra = pa.add_run(a)
        ra.font.size = Pt(10.5)
        ra.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
        pa.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pa.paragraph_format.space_after = Pt(8)

    doc.save(docx_path)
    print(f"Built DOCX: {docx_path}")

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

    # Also create README.txt and github_link.txt
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
  1. PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf (Strict 4-page formal submission)
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
        # Add PDF & DOCX
        z.write("LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf", "PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf")
        z.write("LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx", "PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx")
        z.write("LAB2/README.txt", "README.txt")
        z.write("LAB2/github_link.txt", "github_link.txt")
        
        # Add screenshots
        for fname in os.listdir("LAB2/screenshots"):
            fpath = os.path.join("LAB2/screenshots", fname)
            z.write(fpath, f"screenshots/{fname}")

    print(f"Created ZIP: {zip_path} ({os.path.getsize(zip_path)} bytes)")

    # Also duplicate to PES2UG24CS307_NavyakishoreDash_LAB2.zip
    shutil.copyfile(zip_path, "PES2UG24CS307_NavyakishoreDash_LAB2.zip")
    print("Duplicated ZIP to PES2UG24CS307_NavyakishoreDash_LAB2.zip")

if __name__ == "__main__":
    pdf_dest = "LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf"
    docx_dest = "LAB2/PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx"

    build_pdf(pdf_dest)
    verify_pdf(pdf_dest)
    build_docx(docx_dest)

    # Copy to root directory for easy access
    shutil.copyfile(pdf_dest, "LAB2_Report_PES2UG24CS307.pdf")
    shutil.copyfile(docx_dest, "LAB2_Report_PES2UG24CS307.docx")
    shutil.copyfile(pdf_dest, "PES2UG24CS307_NavyakishoreDash_LAB2_Report.pdf")
    shutil.copyfile(docx_dest, "PES2UG24CS307_NavyakishoreDash_LAB2_Report.docx")

    # Create submission zip
    create_zip_submission()
    print("ALL LAB 2 ARTIFACTS AND DELIVERABLES SUCCESSFULLY PRODUCED!")
