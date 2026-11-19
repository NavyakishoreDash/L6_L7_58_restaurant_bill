import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

COLOR_PRIMARY = RGBColor(26, 35, 126)     # Deep Navy #1A237E
COLOR_SECONDARY = RGBColor(63, 81, 181)   # Indigo #3F51B5
COLOR_DARK = RGBColor(33, 33, 33)         # Near Black #212121
COLOR_MUTED = RGBColor(102, 102, 102)     # Gray #666666
COLOR_GREEN = RGBColor(46, 125, 50)       # Green #2E7D32

HEX_PRIMARY = "1A237E"
HEX_BG_LIGHT = "F4F6F9"
HEX_CODE_BG = "F8F9FA"
HEX_BORDER = "D0D7DE"
HEX_ACCENT_BG = "E8EAF6"

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_header(doc):
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("PES UNIVERSITY")
    r_inst.font.name = "Arial"
    r_inst.font.size = Pt(13)
    r_inst.font.bold = True
    r_inst.font.color.rgb = COLOR_PRIMARY

    p_dept = doc.add_paragraph()
    p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dept.paragraph_format.space_before = Pt(0)
    p_dept.paragraph_format.space_after = Pt(4)
    r_dept = p_dept.add_run("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
    r_dept.font.name = "Calibri"
    r_dept.font.size = Pt(10.5)
    r_dept.font.bold = True
    r_dept.font.color.rgb = COLOR_MUTED

    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_course.paragraph_format.space_before = Pt(0)
    p_course.paragraph_format.space_after = Pt(12)
    r_course = p_course.add_run("UE23CS301 — SOFTWARE ENGINEERING LABORATORY")
    r_course.font.name = "Calibri"
    r_course.font.size = Pt(11)
    r_course.font.bold = True
    r_course.font.color.rgb = COLOR_SECONDARY

    # Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("LAB 6: CODE AND BRANCH COVERAGE")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("White-Box Structural Testing of Restaurant Bill Module (100% Statement & Branch Coverage)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = COLOR_SECONDARY

def add_student_table(doc):
    metadata = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course Code", "UE23CS301"), ("Course Title", "Software Engineering")],
        [("Lab Title", "Lab 6: Code & Branch Coverage"), ("Repository", "SETAPESU26/L6_L7_58_restaurant_bill")],
        [("Tools Used", "Python 3.12, Pytest 9.2, Pytest-Cov"), ("Date", "November 2026")]
    ]
    table = doc.add_table(rows=len(metadata), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="D0D7DE")
    
    for r_idx, (col1, col2) in enumerate(metadata):
        cell1 = table.cell(r_idx, 0)
        cell2 = table.cell(r_idx, 1)
        cell1.width = Inches(3.25)
        cell2.width = Inches(3.25)
        set_cell_background(cell1, HEX_BG_LIGHT)
        set_cell_background(cell2, HEX_BG_LIGHT)
        set_cell_margins(cell1, top=70, bottom=70, left=120, right=120)
        set_cell_margins(cell2, top=70, bottom=70, left=120, right=120)

        p1 = cell1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1_lbl = p1.add_run(col1[0] + ": ")
        r1_lbl.bold = True
        r1_lbl.font.size = Pt(9.5)
        r1_lbl.font.color.rgb = COLOR_PRIMARY
        r1_val = p1.add_run(col1[1])
        r1_val.font.size = Pt(9.5)

        p2 = cell2.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        r2_lbl = p2.add_run(col2[0] + ": ")
        r2_lbl.bold = True
        r2_lbl.font.size = Pt(9.5)
        r2_lbl.font.color.rgb = COLOR_PRIMARY
        r2_val = p2.add_run(col2[1])
        r2_val.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(13.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = COLOR_SECONDARY
    return p

def add_body(doc, text, bold_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_DARK
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_DARK
    return p

def add_bullet(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(9.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_DARK
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.color.rgb = COLOR_DARK
    return p

def add_callout(doc, text, title="NOTE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=80, bottom=80, left=140, right=140)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{HEX_PRIMARY}"/>'
        f'  <w:top w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r_t = p.add_run(f"[{title}] ")
    r_t.bold = True
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(9.5)
    r_t.font.color.rgb = COLOR_PRIMARY
    r_b = p.add_run(text)
    r_b.font.name = "Calibri"
    r_b.font.size = Pt(9.5)
    r_b.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, HEX_CODE_BG)
    set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(30, 30, 30)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_image_with_caption(doc, img_path, caption, width_in=6.2):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(3)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(width_in))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(10)
        run_cap = p_cap.add_run(caption)
        run_cap.font.name = "Calibri"
        run_cap.font.size = Pt(9)
        run_cap.font.italic = True
        run_cap.font.color.rgb = COLOR_MUTED

def generate_lab6_document(filename):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    add_header(doc)
    add_student_table(doc)

    # 1. Aim & Objectives
    add_h1(doc, "1. Aim & Objectives")
    add_body(
        doc,
        "The aim of Lab 6 is to perform white-box structural testing on the restaurant_bill.py module using Pytest "
        "and Pytest-Cov to achieve 90%–100% statement and branch coverage. The laboratory emphasizes deriving comprehensive "
        "test cases using Equivalence Class Partitioning (ECP) and Boundary Value Analysis (BVA) to exercise all computational "
        "and decision paths."
    )
    add_body(doc, "Key Learning Objectives:")
    add_bullet(doc, "Understand statement coverage (C0) and branch coverage (C1) criteria.", "Coverage Concepts: ")
    add_bullet(doc, "Derive test inputs targeting decision outcomes in conditional and branching logic.", "Test Case Design: ")
    add_bullet(doc, "Measure, monitor, and analyze coverage using pytest-cov and coverage.py.", "Tool Proficiency: ")
    add_bullet(doc, "Validate exception handling for domain-specific boundary conditions and constraints.", "Exception Testing: ")

    # 2. Theoretical Concepts: Statement & Branch Coverage
    add_h1(doc, "2. Theoretical Concepts")
    add_body(
        doc,
        "Structural white-box testing uses the source code structure to design test cases and assess test completeness:"
    )
    add_bullet(
        doc,
        "Measures the percentage of executable statements in the program executed by the test suite: "
        "Statement Coverage = (Executed Statements / Total Statements) * 100%. "
        "While fundamental, 100% statement coverage does not guarantee that every decision path or condition has been tested.",
        "Statement Coverage (C0): "
    )
    add_bullet(
        doc,
        "Measures whether each decision point (such as if/else statements) evaluates to both True and False outcomes across test executions: "
        "Branch Coverage = (Executed Decision Outcomes / Total Decision Outcomes) * 100%. "
        "Branch coverage subsumes statement coverage and uncovers bugs hidden within conditional logic.",
        "Branch Coverage (C1): "
    )
    add_callout(
        doc,
        "Pytest with pytest-cov and coverage.py enables precise tracking of both executed statements and individual branch paths. "
        "Running with the --cov-branch flag verifies that every branch evaluates to both branches (True and False) during the test session.",
        title="WHITE-BOX TESTING METRIC"
    )

    # 3. System & Module Specification
    add_h1(doc, "3. System & Module Specification")
    add_body(
        doc,
        "The module under test (restaurant_bill.py) manages a restaurant billing and seating workflow through three distinct functions:"
    )

    add_h2(doc, "3.1 bill_total(section, food_total, guests, is_happy_hour)")
    add_body(
        doc,
        "Calculates the total payable bill based on dining section, order total, guest count, and happy hour status:",
        bold_prefix="Purpose: "
    )
    add_bullet(doc, "dine_in (10%), takeaway (0%), rooftop (15%). If unknown, raises ValueError('unknown section').", "Service Charges: ")
    add_bullet(doc, "food_total must be >= 0 and guests >= 1; otherwise raises ValueError('invalid order').", "Input Validation: ")
    add_bullet(doc, "If is_happy_hour is True and section != 'takeaway', food_total is discounted by 20% (food_total * 80 // 100).", "Happy Hour Discount: ")
    add_bullet(doc, "If guests >= 8, an additional surcharge of 200 is added to service charge.", "Large Party Surcharge: ")
    add_bullet(doc, "5% tax is added to (food_total + service): total + total * 5 // 100.", "Tax Calculation: ")

    add_h2(doc, "3.2 split_bill(total, people, tip_percent)")
    add_body(
        doc,
        "Splits a bill evenly among patrons, incorporating tips and remainder adjustments:",
        bold_prefix="Purpose: "
    )
    add_bullet(doc, "people must be >= 1; otherwise raises ValueError('invalid people count').", "People Validation: ")
    add_bullet(doc, "tip_percent must be >= 0; otherwise raises ValueError('invalid tip').", "Tip Validation: ")
    add_bullet(doc, "Calculates grand_total = total + total * tip_percent // 100.", "Grand Total: ")
    add_bullet(doc, "Calculates share = grand_total // people. If grand_total % people != 0, share is rounded up by 1 (+1).", "Even / Uneven Split: ")

    add_h2(doc, "3.3 table_wait(party_size, free_tables, has_reservation)")
    add_body(
        doc,
        "Determines the seating decision based on party size, table availability, and reservation status:",
        bold_prefix="Purpose: "
    )
    add_bullet(doc, "If party_size > 12, returns 'call ahead' immediately.", "Rule 1 (Oversized Party): ")
    add_bullet(doc, "If has_reservation is True, returns 'seat now'.", "Rule 2 (Reservation): ")
    add_bullet(doc, "If free_tables == 0, returns 'waitlist'.", "Rule 3 (No Capacity): ")
    add_bullet(doc, "If party_size > 6 and free_tables < 2, returns 'waitlist'.", "Rule 4 (Large Walk-in): ")
    add_bullet(doc, "Otherwise, returns 'seat now'.", "Rule 5 (Default): ")

    # 4. Branch Analysis & Decision Table
    add_h1(doc, "4. Branch Coverage & Decision Table")
    add_body(
        doc,
        "To achieve 100% branch coverage, every decision predicate must evaluate to both True and False outcomes. "
        "The table below details each decision branch in restaurant_bill.py and the corresponding test case:"
    )

    # Branch analysis table
    tbl_branch = doc.add_table(rows=12, cols=4)
    tbl_branch.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_branch)

    col_widths = [Inches(1.8), Inches(1.5), Inches(1.2), Inches(2.0)]
    b_headers = ["Function & Predicate", "Branch Condition", "Outcome", "Exercising Test Case"]
    for c_idx, h in enumerate(b_headers):
        cell = tbl_branch.cell(0, c_idx)
        cell.width = col_widths[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    branch_data = [
        ["bill_total: section not in SERVICE_CHARGE", "section in {'dine_in', ...}\nsection == 'balcony'", "False\nTrue", "test_dine_in_bill\ntest_bill_total_unknown_section"],
        ["bill_total: food_total < 0", "food_total = 1000\nfood_total = -50", "False\nTrue", "test_dine_in_bill\ntest_bill_total_negative_food_total"],
        ["bill_total: guests < 1", "guests >= 1 (e.g. 2)\nguests = 0", "False\nTrue", "test_dine_in_bill\ntest_bill_total_invalid_guests"],
        ["bill_total: is_happy_hour and section != 'takeaway'", "Happy hour dine_in\nHappy hour takeaway", "True\nFalse", "test_dine_in_happy_hour\ntest_takeaway_happy_hour_no_discount"],
        ["bill_total: guests >= 8", "guests = 8\nguests = 2", "True\nFalse", "test_large_party_service_surcharge\ntest_dine_in_bill"],
        ["split_bill: people < 1", "people = 4\npeople = 0", "False\nTrue", "test_even_split\ntest_split_bill_invalid_people"],
        ["split_bill: tip_percent < 0", "tip_percent = 10\ntip_percent = -5", "False\nTrue", "test_even_split\ntest_split_bill_negative_tip"],
        ["split_bill: grand_total % people != 0", "1100 % 4 == 0\n100 % 3 != 0", "False\nTrue", "test_even_split\ntest_uneven_split_rounds_up"],
        ["table_wait: party_size > 12", "party_size = 13\nparty_size = 4", "True\nFalse", "test_table_wait_large_party_calls_ahead\ntest_table_wait_with_reservation"],
        ["table_wait: has_reservation", "has_reservation = True\nhas_reservation = False", "True\nFalse", "test_table_wait_with_reservation\ntest_table_wait_no_free_tables"],
        ["table_wait: free_tables == 0", "free_tables = 0\nfree_tables >= 1", "True\nFalse", "test_table_wait_no_free_tables\ntest_table_wait_small_party_one_table"],
    ]

    for r_idx, row in enumerate(branch_data):
        for c_idx, val in enumerate(row):
            cell = tbl_branch.cell(r_idx + 1, c_idx)
            cell.width = col_widths[c_idx]
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if c_idx in (1, 2):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            r.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. Test Suite Implementation (21 Test Cases)
    add_h1(doc, "5. Test Suite Implementation (21 Test Cases)")
    add_body(
        doc,
        "The test suite test_restaurant_bill.py was authored with 21 unit tests systematically verifying all boundary values, "
        "valid operational scenarios, and exception conditions:"
    )

    # Table of all 21 test cases
    tbl_tests = doc.add_table(rows=22, cols=4)
    tbl_tests.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tests)

    t_widths = [Inches(0.6), Inches(2.2), Inches(2.0), Inches(1.7)]
    t_headers = ["ID", "Test Function Name", "Inputs / Scenario", "Expected Result"]
    for c_idx, h in enumerate(t_headers):
        cell = tbl_tests.cell(0, c_idx)
        cell.width = t_widths[c_idx]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    test_case_records = [
        ("TC01", "test_dine_in_bill", "'dine_in', 1000, 2, False", "1155 (10% svc + 5% tax)"),
        ("TC02", "test_dine_in_happy_hour", "'dine_in', 1000, 2, True", "924 (20% food disc + svc + tax)"),
        ("TC03", "test_large_party_service_surcharge", "'dine_in', 1000, 8, False", "1365 (+200 svc for 8 guests)"),
        ("TC04", "test_large_party_happy_hour", "'dine_in', 1000, 8, True", "1134 (disc + 200 svc + tax)"),
        ("TC05", "test_takeaway_happy_hour_no_discount", "'takeaway', 1000, 2, True", "1050 (no happy hour disc)"),
        ("TC06", "test_takeaway_regular", "'takeaway', 1000, 2, False", "1050 (0% svc + 5% tax)"),
        ("TC07", "test_rooftop_service_charge", "'rooftop', 1000, 4, False", "1207 (15% svc + 5% tax)"),
        ("TC08", "test_rooftop_happy_hour", "'rooftop', 1000, 4, True", "966 (20% disc + 15% svc + tax)"),
        ("TC09", "test_bill_total_unknown_section", "'balcony', 1000, 2, False", "raises ValueError('unknown section')"),
        ("TC10", "test_bill_total_negative_food_total", "'dine_in', -50, 2, False", "raises ValueError('invalid order')"),
        ("TC11", "test_bill_total_invalid_guests", "'dine_in', 1000, 0, False", "raises ValueError('invalid order')"),
        ("TC12", "test_even_split", "1000, 4, 10% tip", "275 (1100 // 4 exact)"),
        ("TC13", "test_uneven_split_rounds_up", "100, 3, 0% tip", "34 (100 // 3 + 1 remainder)"),
        ("TC14", "test_split_bill_invalid_people", "1000, 0, 10% tip", "raises ValueError('invalid people count')"),
        ("TC15", "test_split_bill_negative_tip", "1000, 4, -5% tip", "raises ValueError('invalid tip')"),
        ("TC16", "test_table_wait_large_party_calls_ahead", "party=13, tables=5, res=True", "'call ahead'"),
        ("TC17", "test_table_wait_with_reservation", "party=4, tables=0, res=True", "'seat now'"),
        ("TC18", "test_table_wait_no_free_tables", "party=4, tables=0, res=False", "'waitlist'"),
        ("TC19", "test_table_wait_party_over_6_few_tables", "party=7, tables=1, res=False", "'waitlist'"),
        ("TC20", "test_table_wait_party_over_6_sufficient_tables", "party=7, tables=2, res=False", "'seat now'"),
        ("TC21", "test_table_wait_small_party_one_table", "party=4, tables=1, res=False", "'seat now'"),
    ]

    for r_idx, (tc_id, tc_name, tc_input, tc_exp) in enumerate(test_case_records):
        row_cells = tbl_tests.cell(r_idx + 1, 0), tbl_tests.cell(r_idx + 1, 1), tbl_tests.cell(r_idx + 1, 2), tbl_tests.cell(r_idx + 1, 3)
        for c_idx, cell in enumerate(row_cells):
            cell.width = t_widths[c_idx]
            set_cell_margins(cell, top=45, bottom=45, left=70, right=70)
            p = cell.paragraphs[0]
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(tc_id)
                r.bold = True
            elif c_idx == 1:
                r = p.add_run(tc_name)
                r.font.name = "Consolas"
                r.font.size = Pt(8)
            else:
                r = p.add_run(tc_input if c_idx == 2 else tc_exp)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. Test Execution & Coverage Results
    add_h1(doc, "6. Test Execution & Coverage Results")
    add_body(
        doc,
        "The test suite was executed from the LAB6 working directory using the Pytest command line with coverage flags:\n"
        "pytest -v --cov=restaurant_bill --cov-branch --cov-report=term-missing"
    )

    add_image_with_caption(
        doc,
        "LAB6/screenshots/SS_lab6_coverage.png",
        "Figure 1: Lab 6 Pytest Execution – Demonstrating 21/21 passed tests with 100% statement and 100% branch coverage."
    )

    # Coverage summary table
    add_h2(doc, "Coverage Summary Table")
    tbl_cov = doc.add_table(rows=3, cols=6)
    tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_cov)

    cov_h = ["File Name", "Statements", "Missed Stmts", "Branches", "Partial Branches", "Total Coverage"]
    for c_idx, h in enumerate(cov_h):
        cell = tbl_cov.cell(0, c_idx)
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    cov_rows = [
        ["restaurant_bill.py", "33", "0", "22", "0", "100%"],
        ["TOTAL", "33", "0", "22", "0", "100%"]
    ]
    for r_idx, row in enumerate(cov_rows):
        for c_idx, val in enumerate(row):
            cell = tbl_cov.cell(r_idx + 1, c_idx)
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if r_idx == 1 or c_idx == 5:
                r.bold = True
                if c_idx == 5:
                    r.font.color.rgb = COLOR_GREEN

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 7. Conclusion & Analysis
    add_h1(doc, "7. Conclusion & Key Observations")
    add_body(
        doc,
        "1. Complete Structural Verification: By systematically constructing test cases for every predicate and boundary value, "
        "100% statement coverage (33/33 statements) and 100% branch coverage (22/22 branches) were achieved."
    )
    add_body(
        doc,
        "2. Subtleties in Branch Logic: Branch coverage exposed subtle conditional dependencies that simple statement coverage could miss. "
        "Specifically, verifying that takeaway orders are not discounted during happy hour (section != 'takeaway') required testing both "
        "is_happy_hour=True with takeaway and is_happy_hour=True with dine_in."
    )
    add_body(
        doc,
        "3. Boundary Condition Rigor: Testing boundary conditions (guests == 8 for large party surcharge, people == 0 for division error, "
        "and remainder round-ups for split bills) ensured complete robustness against off-by-one errors."
    )

    # 8. Source Code Appendix
    add_h1(doc, "8. Source Code Appendix")
    add_h2(doc, "8.1 restaurant_bill.py (Module Under Test)")
    with open("LAB6/restaurant_bill.py", "r") as f:
        add_code_block(doc, f.read())

    add_h2(doc, "8.2 test_restaurant_bill.py (Complete 21 Test Cases)")
    with open("LAB6/test_restaurant_bill.py", "r") as f:
        add_code_block(doc, f.read())

    doc.save(filename)
    print(f"Saved {filename}")

if __name__ == "__main__":
    generate_lab6_document("LAB6/PES2UG24CS307_NavyakishoreDash_LAB6.docx")
    generate_lab6_document("LAB6/LAB6_Report.docx")
    generate_lab6_document("LAB6_Report_PES2UG24CS307.docx")
