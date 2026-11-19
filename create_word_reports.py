import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# Color palette
COLOR_PRIMARY = RGBColor(26, 35, 126)     # Deep Navy #1A237E
COLOR_SECONDARY = RGBColor(63, 81, 181)   # Indigo #3F51B5
COLOR_DARK = RGBColor(33, 33, 33)         # Near Black #212121
COLOR_MUTED = RGBColor(102, 102, 102)     # Gray #666666
COLOR_GREEN = RGBColor(46, 125, 50)       # Green #2E7D32
COLOR_RED = RGBColor(198, 40, 40)         # Red #C62828

HEX_PRIMARY = "1A237E"
HEX_BG_LIGHT = "F4F6F9"
HEX_CODE_BG = "F8F9FA"
HEX_BORDER = "D0D7DE"
HEX_BORDER_DARK = "1A237E"

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
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

def add_header(doc, title_text, subtitle_text):
    # Department / Lab Category
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(2)
    p_meta.paragraph_format.space_before = Pt(0)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
    run_meta.font.name = "Calibri"
    run_meta.font.size = Pt(10)
    run_meta.font.bold = True
    run_meta.font.color.rgb = COLOR_MUTED

    # Document Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(4)
    p_title.paragraph_format.space_before = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(title_text)
    run_title.font.name = "Arial"
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run(subtitle_text)
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = COLOR_SECONDARY

def add_student_table(doc, rows_data):
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="D0D7DE")
    
    for r_idx, (col1, col2) in enumerate(rows_data):
        cell1 = table.cell(r_idx, 0)
        cell2 = table.cell(r_idx, 1)
        cell1.width = Inches(3.25)
        cell2.width = Inches(3.25)
        set_cell_background(cell1, HEX_BG_LIGHT)
        set_cell_background(cell2, HEX_BG_LIGHT)
        set_cell_margins(cell1, top=80, bottom=80, left=140, right=140)
        set_cell_margins(cell2, top=80, bottom=80, left=140, right=140)

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
    run.font.size = Pt(14)
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
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_SECONDARY
    return p

def add_body(doc, text, bold_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_DARK
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
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
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_DARK
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_DARK
    return p

def add_callout(doc, text, title="NOTE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=100, bottom=100, left=180, right=180)
    
    # Left border only
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
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
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


# ===========================================================================
# 1. GENERATE LAB 7 WORD DOCUMENT
# ===========================================================================
def build_lab7_doc(filename):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    add_header(
        doc,
        "LAB 7: Fuzzing for Input Validation Bugs",
        "Property-Based Fuzz Testing with Hypothesis on Restaurant Bill Parser"
    )

    metadata = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course", "Software Engineering (SE)"), ("Course Code", "UE23CS301")],
        [("Topic", "Lab 7 Fuzz Testing"), ("Repository", "SETAPESU26/L6_L7_58_restaurant_bill")],
        [("Platform", "Python 3.12, Pytest 9.2, Hypothesis"), ("Date", "November 2026")]
    ]
    add_student_table(doc, metadata)

    # 1. Objective & Learning Outcomes
    add_h1(doc, "1. Objective & Learning Outcomes")
    add_body(
        doc,
        "The objective of this laboratory is to explore fuzz testing as a dynamic security and quality assurance "
        "technique by detecting, analyzing, and fixing input validation flaws in a Python application using the Hypothesis "
        "property-based fuzzing library."
    )
    add_body(doc, "By the end of this lab, the student is expected to:")
    add_bullet(doc, "Understand the methodology, goals, and advantages of dynamic fuzz testing.", "Methodology: ")
    add_bullet(doc, "Apply property-based fuzzing using Hypothesis (@given, strategies as st).", "Hypothesis: ")
    add_bullet(doc, "Detect hidden bugs, edge cases, and unexpected crash conditions across diverse inputs.", "Bug Detection: ")
    add_bullet(doc, "Harden application source code using defensive programming patterns.", "Defensive Programming: ")

    # 2. Introduction to Fuzzing and Hypothesis
    add_h1(doc, "2. Introduction to Fuzzing & Hypothesis")
    add_body(
        doc,
        "Fuzzing (or fuzz testing) is a dynamic testing technique where random, malformed, or unexpected data is fed to a "
        "program to uncover crashes, memory leaks, unhandled exceptions, and security vulnerabilities. Traditional unit testing "
        "only validates scenarios that the engineer specifically envisioned; fuzzing exposes edge cases that developers never anticipated."
    )
    add_callout(
        doc,
        "Hypothesis is a modern property-based testing framework for Python. Instead of testing static inputs against expected outputs, "
        "you assert general properties (invariants) that must always hold true—such as 'the function must never crash on any Unicode string or None'. "
        "When Hypothesis finds a failing input, it automatically minimizes ('shrinks') it to the simplest reproducing example.",
        title="HYPOTHESIS PROPERTY-BASED FUZZING"
    )

    # 3. Lab Execution Steps & Screenshots
    add_h1(doc, "3. Lab Execution Steps & Deliverable Screenshots")
    
    add_h2(doc, "Step 1: Installing Dependencies")
    add_body(doc, "The required dependencies were installed via pip:")
    add_code_block(doc, "pip install hypothesis\npip install pytest")

    add_h2(doc, "Step 2: Running the Buggy Application (SS1)")
    add_body(
        doc,
        "Command executed: python restaurant_bill_parser.py\n"
        "The application was run with standard 'happy-path' user inputs (*paneer tikka*, 250,180,320, and 1200/4). "
        "Under ideal conditions, the application succeeds as shown below:"
    )
    add_image_with_caption(
        doc,
        "LAB7/screenshots/SS1_buggy_run.png",
        "Figure 1: Screenshot SS1 – Execution of the original restaurant_bill_parser.py with standard happy-path inputs."
    )

    add_h2(doc, "Step 3: Completing Property-Based Tests in test_restaurant_bill_parser.py")
    add_body(
        doc,
        "Using Hypothesis strategies (st.text() | st.none()), tests were defined for all three functions. "
        "Each test asserts that the function must never raise an unhandled exception when invoked with any valid string, "
        "arbitrary character sequence, or None:"
    )
    add_code_block(
        doc,
        "from hypothesis import given, strategies as st\n"
        "from restaurant_bill_parser import clean_dish_name, parse_item_prices, split_amount\n\n"
        "@given(st.text() | st.none())\n"
        "def test_clean_dish_name_no_crash(s):\n"
        "    try:\n"
        "        clean_dish_name(s)\n"
        "    except Exception as e:\n"
        "        assert False, f\"clean_dish_name crashed with: {e}\"\n\n"
        "@given(st.text() | st.none())\n"
        "def test_parse_item_prices_no_crash(s):\n"
        "    try:\n"
        "        parse_item_prices(s)\n"
        "    except Exception as e:\n"
        "        assert False, f\"parse_item_prices crashed with: {e}\"\n\n"
        "@given(st.text() | st.none())\n"
        "def test_split_amount_no_crash(s):\n"
        "    try:\n"
        "        split_amount(s)\n"
        "    except Exception as e:\n"
        "        assert False, f\"split_amount crashed with: {e}\""
    )

    add_h2(doc, "Step 4: Running Fuzz Tests on Buggy Module (SS2)")
    add_body(
        doc,
        "Command executed: pytest test_restaurant_bill_parser.py\n"
        "Hypothesis immediately generated counterexamples exposing critical unhandled exceptions in all three functions. "
        "All 3 tests failed, demonstrating crashes on None, empty strings, and missing delimiters:"
    )
    add_image_with_caption(
        doc,
        "LAB7/screenshots/SS2_hypothesis_failing.png",
        "Figure 2: Screenshot SS2 – Hypothesis fuzz testing failing on the buggy module, pinpointing minimal failing cases."
    )

    add_h2(doc, "Step 5: Hardening the Application & Running Fixed Module (SS3)")
    add_body(
        doc,
        "Defensive programming guards, type checks, and exception containment were implemented across all three functions. "
        "The fixed module was then re-run with standard user inputs to verify that valid functionality is completely preserved:"
    )
    add_image_with_caption(
        doc,
        "LAB7/screenshots/SS3_fixed_run.png",
        "Figure 3: Screenshot SS3 – Execution of the hardened restaurant_bill_parser.py with standard happy-path inputs."
    )

    add_h2(doc, "Step 6: Re-Running Fuzz Tests on Fixed Module (SS4)")
    add_body(
        doc,
        "Command executed: pytest -v test_restaurant_bill_parser.py\n"
        "Hypothesis executed hundreds of randomized iterations over arbitrary strings and null inputs. "
        "All tests passed with zero crashes:"
    )
    add_image_with_caption(
        doc,
        "LAB7/screenshots/SS4_hypothesis_passing.png",
        "Figure 4: Screenshot SS4 – 100% of Hypothesis fuzz tests passing on the hardened module."
    )

    # 4. Reflection Questions & Technical Analysis
    add_h1(doc, "4. Reflection Questions & Technical Analysis")

    add_h2(doc, "Reflection Question 1: What bugs or vulnerabilities did Hypothesis uncover in each function?")
    add_body(
        doc,
        "Hypothesis systematically generated adversarial inputs across the domain st.text() | st.none(), exposing three main categories of vulnerabilities:"
    )
    add_bullet(
        doc,
        "When name is None or non-string, calling name.strip() raised AttributeError: 'NoneType' object has no attribute 'strip'. "
        "In addition, empty strings or strings made solely of chef markers (e.g. '***') were not defensively handled.",
        "clean_dish_name(name): "
    )
    add_bullet(
        doc,
        "Passing None triggered AttributeError on text.split(','). When passed an empty string ('') or trailing commas ('100,'), "
        "int(price) raised ValueError: invalid literal for int() with base 10: ''. When passed non-integer or float strings ('12.5', 'abc'), "
        "the function unceremoniously crashed.",
        "parse_item_prices(text): "
    )
    add_bullet(
        doc,
        "Passing None triggered AttributeError. When passed strings lacking exactly one slash delimiter (e.g., '', '1200', '100/2/3'), "
        "tuple unpacking raised ValueError: not enough values to unpack or too many values to unpack. "
        "When people was '0' (e.g. '1200/0'), int(total) // int(people) triggered ZeroDivisionError. Non-numeric values triggered ValueError.",
        "split_amount(text): "
    )

    add_h2(doc, "Reflection Question 2: How did you fix each function using defensive programming?")
    add_body(
        doc,
        "Each function was hardened using structured defensive programming principles (validation, containment, and safe defaults):"
    )
    add_bullet(
        doc,
        "Added explicit type validation using isinstance(name, str). If the input is not a string (including None), it safely returns an empty string \"\". "
        "Marker characters (*, #, ~) and whitespace are safely removed before title-casing.",
        "Hardening clean_dish_name: "
    )
    add_bullet(
        doc,
        "Added string type validation returning [] on invalid input. Split by comma and parsed each element inside a try-except (ValueError, TypeError) block, "
        "gracefully filtering out empty or non-numeric tokens while preserving valid integer prices.",
        "Hardening parse_item_prices: "
    )
    add_bullet(
        doc,
        "Added string validation and verified len(parts) == 2. Wrapped integer conversion and division inside try-except (ValueError, TypeError, ZeroDivisionError), "
        "explicitly checking people > 0 before division, returning a safe default of 0 on any invalid input format.",
        "Hardening split_amount: "
    )

    add_h2(doc, "Reflection Question 3: What are the key advantages of property-based fuzz testing over traditional unit testing?")
    add_bullet(
        doc,
        "Traditional unit tests only examine specific edge cases that the developer personally considered. Hypothesis tests automatically generate "
        "diverse, adversarial inputs (null bytes, surrogates, empty strings, delimiter floods) that unmask unknown blind spots.",
        "Automated Edge-Case Discovery: "
    )
    add_bullet(
        doc,
        "When Hypothesis encounters a failing input, it automatically minimizes ('shrinks') the failing input down to the simplest reproducible example "
        "(e.g., reducing a long Unicode string failure to s='' or s=None), making root-cause diagnosis immediate.",
        "Test Case Shrinking: "
    )
    add_bullet(
        doc,
        "Unit testing evaluates individual point examples. Property-based fuzzing tests universal program invariants (such as robustness and absence of crashes) "
        "across hundreds of random variations on every execution run.",
        "Invariant Verification: "
    )

    # 5. Source Code Appendix
    add_h1(doc, "5. Source Code Appendix")
    add_h2(doc, "Fixed Application: restaurant_bill_parser.py")
    with open("LAB7/restaurant_bill_parser.py", "r") as f:
        add_code_block(doc, f.read())

    add_h2(doc, "Hypothesis Test Suite: test_restaurant_bill_parser.py")
    with open("LAB7/test_restaurant_bill_parser.py", "r") as f:
        add_code_block(doc, f.read())

    doc.save(filename)
    print(f"Saved {filename}")


# ===========================================================================
# 2. GENERATE LAB 6 WORD DOCUMENT
# ===========================================================================
def build_lab6_doc(filename):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    add_header(
        doc,
        "LAB 6: Code and Branch Coverage",
        "Restaurant Bill Module Testing (100% Statement & Branch Coverage)"
    )

    metadata = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course", "Software Engineering (SE)"), ("Course Code", "UE23CS301")],
        [("Topic", "Lab 6 Code & Branch Coverage"), ("Repository", "SETAPESU26/L6_L7_58_restaurant_bill")],
        [("Platform", "Python 3.12, Pytest 9.2, Pytest-Cov"), ("Date", "November 2026")]
    ]
    add_student_table(doc, metadata)

    # 1. Objective & Description
    add_h1(doc, "1. Objective & Module Description")
    add_body(
        doc,
        "The objective of Lab 6 is to implement a comprehensive test suite for restaurant_bill.py that achieves "
        "at least 90–100% code coverage, with complete branch coverage across all decision paths in the module. "
        "The module implements restaurant billing, tip splitting, and table seating operations across three primary functions:"
    )
    add_bullet(
        doc,
        "Calculates bill total by applying section service charges (dine_in: 10%, takeaway: 0%, rooftop: 15%), "
        "happy hour discount (20% discount on food total, except for takeaway), large party surcharge (+200 for 8+ guests), "
        "and a 5% tax. Raises ValueError on invalid section or order parameters.",
        "bill_total(section, food_total, guests, is_happy_hour): "
    )
    add_bullet(
        doc,
        "Calculates each diner's share including tip percentage, rounding up (+1) if the division has a remainder. "
        "Raises ValueError if people < 1 or tip_percent < 0.",
        "split_bill(total, people, tip_percent): "
    )
    add_bullet(
        doc,
        "Determines seating action: parties > 12 must 'call ahead'; parties with reservations are 'seat now'; "
        "walk-in parties with 0 free tables are 'waitlist'; walk-in parties > 6 with fewer than 2 tables are 'waitlist'; "
        "all other parties are 'seat now'.",
        "table_wait(party_size, free_tables, has_reservation): "
    )

    # 2. Coverage Results & Execution Screenshot
    add_h1(doc, "2. Coverage Results & Pytest Execution Screenshot")
    add_body(
        doc,
        "Command executed: pytest -v --cov=restaurant_bill --cov-branch --cov-report=term-missing\n"
        "All 21 test cases passed, achieving 100% statement coverage and 100% branch coverage:"
    )
    add_image_with_caption(
        doc,
        "LAB6/screenshots/SS_lab6_coverage.png",
        "Figure 1: Pytest execution demonstrating all 21 test cases passing with 100% statement and 100% branch coverage."
    )

    # Coverage Table
    table = doc.add_table(rows=3, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    headers = ["File", "Statements", "Miss", "Branches", "BrPart", "Coverage"]
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    data_rows = [
        ["restaurant_bill.py", "33", "0", "22", "0", "100%"],
        ["TOTAL", "33", "0", "22", "0", "100%"]
    ]
    for r_idx, row in enumerate(data_rows):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx + 1, c_idx)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if r_idx == 1 or c_idx == 5:
                r.bold = True
                if c_idx == 5:
                    r.font.color.rgb = COLOR_GREEN

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Decision Path Breakdown
    add_h1(doc, "3. Decision Path & Branch Coverage Breakdown")
    add_h2(doc, "bill_total Decision Paths")
    add_bullet(doc, "Covered with invalid section 'balcony' raising ValueError(\"unknown section\").", "Branch 1 (section not in SERVICE_CHARGE): ")
    add_bullet(doc, "Covered with food_total = -50 and guests = 0 respectively, raising ValueError(\"invalid order\").", "Branch 2 (food_total < 0 or guests < 1): ")
    add_bullet(doc, "Covered with dine_in + happy hour (discount applied) vs takeaway + happy hour (discount skipped per specification).", "Branch 3 (is_happy_hour and section != 'takeaway'): ")
    add_bullet(doc, "Covered with guests = 8 (+200 service charge) vs guests = 2 (normal service charge).", "Branch 4 (guests >= 8): ")
    add_bullet(doc, "Covered dine_in (10%), takeaway (0%), and rooftop (15%) seating sections.", "Branch 5 (SERVICE_CHARGE[section]): ")

    add_h2(doc, "split_bill Decision Paths")
    add_bullet(doc, "Covered with people = 0 raising ValueError(\"invalid people count\").", "Branch 1 (people < 1): ")
    add_bullet(doc, "Covered with tip_percent = -5 raising ValueError(\"invalid tip\").", "Branch 2 (tip_percent < 0): ")
    add_bullet(doc, "Covered exact split (1000 with 10% tip among 4 diners = 275) vs uneven split (100 among 3 diners = 34 due to remainder round-up).", "Branch 3 (grand_total % people != 0): ")

    add_h2(doc, "table_wait Decision Paths")
    add_bullet(doc, "Covered with party sizes 13 and 14 returning 'call ahead'.", "Branch 1 (party_size > 12): ")
    add_bullet(doc, "Covered with parties having reservation returning 'seat now' regardless of tables.", "Branch 2 (has_reservation): ")
    add_bullet(doc, "Covered walk-in with 0 free tables returning 'waitlist'.", "Branch 3 (free_tables == 0): ")
    add_bullet(doc, "Covered party_size = 7 and free_tables = 1 returning 'waitlist'.", "Branch 4 (party_size > 6 and free_tables < 2): ")
    add_bullet(doc, "Covered party_size = 7 and free_tables = 2 returning 'seat now'.", "Branch 5 (party_size > 6 and free_tables >= 2): ")
    add_bullet(doc, "Covered party_size <= 6 and free_tables >= 1 returning 'seat now'.", "Branch 6 (small party with available table): ")

    # 4. Source Code Appendix
    add_h1(doc, "4. Source Code Appendix")
    add_h2(doc, "Target Module: restaurant_bill.py")
    with open("LAB6/restaurant_bill.py", "r") as f:
        add_code_block(doc, f.read())

    add_h2(doc, "Test Suite: test_restaurant_bill.py (21 Test Cases)")
    with open("LAB6/test_restaurant_bill.py", "r") as f:
        add_code_block(doc, f.read())

    doc.save(filename)
    print(f"Saved {filename}")


# ===========================================================================
# 3. GENERATE COMBINED LAB 6 & LAB 7 WORD DOCUMENT
# ===========================================================================
def build_combined_doc(filename):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    add_header(
        doc,
        "Software Engineering Laboratory Report",
        "Lab 6: Code & Branch Coverage | Lab 7: Fuzzing for Input Validation Bugs"
    )

    metadata = [
        [("Student Name", "Navyakishore Dash"), ("SRN", "PES2UG24CS307")],
        [("Course", "Software Engineering (SE)"), ("Course Code", "UE23CS301")],
        [("Topic", "Lab 6 Coverage & Lab 7 Fuzzing"), ("Repository", "SETAPESU26/L6_L7_58_restaurant_bill")],
        [("Platform", "Python 3.12, Pytest 9.2, Hypothesis, Pytest-Cov"), ("Date", "November 2026")]
    ]
    add_student_table(doc, metadata)

    # PART 1: LAB 6
    add_h1(doc, "PART I — LAB 6: CODE AND BRANCH COVERAGE")
    add_body(
        doc,
        "Lab 6 addresses rigorous structural unit testing, ensuring that every executable statement and decision branch "
        "in the restaurant_bill.py module is exercised and validated."
    )
    
    add_h2(doc, "1.1 Coverage Results & Pytest Execution Screenshot")
    add_body(doc, "Command: pytest -v --cov=restaurant_bill --cov-branch --cov-report=term-missing")
    add_image_with_caption(
        doc,
        "LAB6/screenshots/SS_lab6_coverage.png",
        "Figure 1: Lab 6 Pytest execution showing 21 tests passed with 100% Statement and Branch coverage."
    )

    # Coverage Table
    table = doc.add_table(rows=3, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    headers = ["File", "Statements", "Miss", "Branches", "BrPart", "Coverage"]
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    data_rows = [
        ["restaurant_bill.py", "33", "0", "22", "0", "100%"],
        ["TOTAL", "33", "0", "22", "0", "100%"]
    ]
    for r_idx, row in enumerate(data_rows):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx + 1, c_idx)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if r_idx == 1 or c_idx == 5:
                r.bold = True
                if c_idx == 5:
                    r.font.color.rgb = COLOR_GREEN

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    add_h2(doc, "1.2 Lab 6 Summary of Tests (21 Test Cases)")
    add_bullet(doc, "test_dine_in_bill: Standard dine-in calculation without happy hour.", "Test 1: ")
    add_bullet(doc, "test_dine_in_happy_hour: 20% discount on food total during happy hour.", "Test 2: ")
    add_bullet(doc, "test_large_party_service_surcharge: Party size >= 8 incurs 200 service charge.", "Test 3: ")
    add_bullet(doc, "test_large_party_happy_hour: Combined happy hour and large party fee.", "Test 4: ")
    add_bullet(doc, "test_takeaway_happy_hour_no_discount: Takeaway excludes happy hour discount.", "Test 5: ")
    add_bullet(doc, "test_takeaway_regular: Standard takeaway with 0% service charge.", "Test 6: ")
    add_bullet(doc, "test_rooftop_service_charge: Rooftop section with 15% service charge.", "Test 7: ")
    add_bullet(doc, "test_rooftop_happy_hour: Rooftop section with happy hour discount.", "Test 8: ")
    add_bullet(doc, "test_bill_total_unknown_section: Raises ValueError on unknown section.", "Test 9: ")
    add_bullet(doc, "test_bill_total_negative_food_total: Raises ValueError on negative food total.", "Test 10: ")
    add_bullet(doc, "test_bill_total_invalid_guests: Raises ValueError on guests < 1.", "Test 11: ")
    add_bullet(doc, "test_even_split: Exact split with tip percentage.", "Test 12: ")
    add_bullet(doc, "test_uneven_split_rounds_up: Split with remainder rounds share up by 1.", "Test 13: ")
    add_bullet(doc, "test_split_bill_invalid_people: Raises ValueError on people < 1.", "Test 14: ")
    add_bullet(doc, "test_split_bill_negative_tip: Raises ValueError on negative tip.", "Test 15: ")
    add_bullet(doc, "test_table_wait_large_party_calls_ahead: Party size > 12 must 'call ahead'.", "Test 16: ")
    add_bullet(doc, "test_table_wait_with_reservation: Party with reservation seated immediately.", "Test 17: ")
    add_bullet(doc, "test_table_wait_no_free_tables: Walk-in with 0 tables goes to 'waitlist'.", "Test 18: ")
    add_bullet(doc, "test_table_wait_party_over_6_few_tables: Party > 6 with < 2 tables waitlisted.", "Test 19: ")
    add_bullet(doc, "test_table_wait_party_over_6_sufficient_tables: Party > 6 with >= 2 tables seated.", "Test 20: ")
    add_bullet(doc, "test_table_wait_small_party_one_table: Party <= 6 with 1 table seated.", "Test 21: ")

    doc.add_page_break()

    # PART 2: LAB 7
    add_h1(doc, "PART II — LAB 7: FUZZING FOR INPUT VALIDATION BUGS")
    add_body(
        doc,
        "Lab 7 explores dynamic property-based fuzz testing using Hypothesis to discover input validation vulnerabilities "
        "and crashes in restaurant_bill_parser.py, followed by code hardening via defensive programming."
    )

    add_h2(doc, "2.1 Deliverable Screenshots SS1 to SS4")
    add_body(doc, "Screenshot SS1: Running original buggy module with standard inputs (python restaurant_bill_parser.py):")
    add_image_with_caption(doc, "LAB7/screenshots/SS1_buggy_run.png", "Figure 2: Screenshot SS1 – Buggy module running on standard inputs.")

    add_body(doc, "Screenshot SS2: Hypothesis fuzz tests failing on buggy module (pytest test_restaurant_bill_parser.py):")
    add_image_with_caption(doc, "LAB7/screenshots/SS2_hypothesis_failing.png", "Figure 3: Screenshot SS2 – Hypothesis exposing crashes on None, empty strings, and missing delimiters.")

    add_body(doc, "Screenshot SS3: Running fixed, hardened module with standard inputs (python restaurant_bill_parser.py):")
    add_image_with_caption(doc, "LAB7/screenshots/SS3_fixed_run.png", "Figure 4: Screenshot SS3 – Hardened module preserving expected happy-path outputs.")

    add_body(doc, "Screenshot SS4: Hypothesis fuzz tests passing 100% on fixed module (pytest -v test_restaurant_bill_parser.py):")
    add_image_with_caption(doc, "LAB7/screenshots/SS4_hypothesis_passing.png", "Figure 5: Screenshot SS4 – All Hypothesis fuzz tests passing without unhandled exceptions.")

    add_h2(doc, "2.2 Reflection Questions & Technical Analysis")
    add_body(doc, "Reflection Question 1: What bugs or vulnerabilities did Hypothesis uncover in each function?", bold_prefix="Q1: ")
    add_bullet(doc, "clean_dish_name: None caused AttributeError on name.strip(). Empty or marker-only strings were unvalidated.", "Bugs in clean_dish_name: ")
    add_bullet(doc, "parse_item_prices: None caused AttributeError. Empty string '' or trailing commas caused ValueError: invalid literal for int(). Non-numeric strings crashed int conversion.", "Bugs in parse_item_prices: ")
    add_bullet(doc, "split_amount: None caused AttributeError. Missing or extra slashes ('', '1200', '10/2/3') caused ValueError in tuple unpacking. '1200/0' caused ZeroDivisionError.", "Bugs in split_amount: ")

    add_body(doc, "Reflection Question 2: How did you fix each function using defensive programming?", bold_prefix="Q2: ")
    add_bullet(doc, "clean_dish_name: Validated isinstance(name, str), returning \"\" on non-string/None inputs. Safely stripped markers and whitespace.", "Fix for clean_dish_name: ")
    add_bullet(doc, "parse_item_prices: Validated string type; parsed each token inside try-except (ValueError, TypeError), ignoring bad tokens while collecting valid integers.", "Fix for parse_item_prices: ")
    add_bullet(doc, "split_amount: Validated string type and len(parts) == 2; verified people > 0; caught ValueError, TypeError, ZeroDivisionError; returned safe default 0.", "Fix for split_amount: ")

    add_body(doc, "Reflection Question 3: What are the key advantages of property-based fuzz testing over traditional unit testing?", bold_prefix="Q3: ")
    add_bullet(doc, "Automates discovery of unanticipated blind spots, edge cases, and crash conditions.", "Advantage 1: ")
    add_bullet(doc, "Automatically shrinks failing test cases to the minimal reproducing example for rapid debugging.", "Advantage 2: ")
    add_bullet(doc, "Validates global system invariants across hundreds of random variations per run.", "Advantage 3: ")

    doc.save(filename)
    print(f"Saved {filename}")

if __name__ == "__main__":
    build_lab7_doc("LAB7/PES2UG24CS307_NavyakishoreDash_LAB7.docx")
    build_lab7_doc("LAB7/SRN_NAME_LAB7.docx")
    build_lab6_doc("LAB6/PES2UG24CS307_NavyakishoreDash_LAB6.docx")
    build_lab6_doc("LAB6/SRN_NAME_LAB6.docx")
    build_combined_doc("PES2UG24CS307_NavyakishoreDash_SE_Lab6_Lab7_Report.docx")
    build_combined_doc("SE_Lab6_Lab7_Report.docx")
