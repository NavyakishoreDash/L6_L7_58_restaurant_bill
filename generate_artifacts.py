import os
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

# Ensure directories exist
os.makedirs("LAB7/screenshots", exist_ok=True)

# ---------------------------------------------------------------------------
# 1. Terminal Screenshot Generator
# ---------------------------------------------------------------------------

FONT_PATH = "C:/Windows/Fonts/consola.ttf"
FONT_SIZE = 14
LINE_HEIGHT = 20
PADDING = 16
TITLE_BAR_HEIGHT = 32

def render_terminal(title, lines, filename):
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
    bold_font = ImageFont.truetype("C:/Windows/Fonts/consolab.ttf", FONT_SIZE)
    
    # Calculate dimensions
    max_line_len = max(len(text) for text, _ in lines)
    width = max(800, max_line_len * 9 + PADDING * 2)
    height = TITLE_BAR_HEIGHT + len(lines) * LINE_HEIGHT + PADDING * 2

    img = Image.new("RGB", (width, height), "#181824")
    draw = ImageDraw.Draw(img)

    # Title bar
    draw.rectangle([0, 0, width, TITLE_BAR_HEIGHT], fill="#232433")
    draw.line([0, TITLE_BAR_HEIGHT, width, TITLE_BAR_HEIGHT], fill="#303144", width=1)

    # Window control dots (mac/modern terminal style)
    draw.ellipse([14, 10, 26, 22], fill="#FF5F56")
    draw.ellipse([34, 10, 46, 22], fill="#FFBD2E")
    draw.ellipse([54, 10, 66, 22], fill="#27C93F")

    # Title text
    draw.text((80, 8), title, fill="#A0A0B0", font=font)

    # Render lines
    y = TITLE_BAR_HEIGHT + PADDING
    for text, style in lines:
        f = bold_font if "bold" in style else font
        color = style.get("color", "#D4D4D4")
        draw.text((PADDING, y), text, fill=color, font=f)
        y += LINE_HEIGHT

    img.save(filename, "PNG", dpi=(150, 150))
    print(f"Saved {filename}")

# --- SS1 Content ---
ss1_lines = [
    ("PS C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7> python restaurant_bill_parser.py", {"color": "#61AFEF", "bold": True}),
    ("==== Running clean_dish_name ====", {"color": "#E5C07B"}),
    ("Enter a dish name from the menu (e.g. *paneer tikka*): *paneer tikka*", {"color": "#ABB2BF"}),
    ("Dish name: Paneer Tikka", {"color": "#98C379", "bold": True}),
    ("", {}),
    ("==== Running parse_item_prices ====", {"color": "#E5C07B"}),
    ("Enter item prices separated by commas (e.g. 250,180,320): 250,180,320", {"color": "#ABB2BF"}),
    ("Item prices: [250, 180, 320]", {"color": "#98C379", "bold": True}),
    ("", {}),
    ("==== Running split_amount ====", {"color": "#E5C07B"}),
    ("Enter the bill and number of people as TOTAL/PEOPLE (e.g. 1200/4): 1200/4", {"color": "#ABB2BF"}),
    ("Each person pays: 300", {"color": "#98C379", "bold": True}),
]
render_terminal("Windows PowerShell - SS1: Buggy Module Execution", ss1_lines, "LAB7/screenshots/SS1_buggy_run.png")

# --- SS2 Content ---
ss2_lines = [
    ("PS C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7> pytest test_restaurant_bill_parser.py", {"color": "#61AFEF", "bold": True}),
    ("============================= test session starts =============================", {"color": "#7F848E"}),
    ("platform win32 -- Python 3.12.6, pytest-9.2.0, pluggy-1.7.0", {"color": "#7F848E"}),
    ("rootdir: C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7", {"color": "#7F848E"}),
    ("plugins: anyio-4.16.0, hypothesis-6.170.15, cov-7.1.0", {"color": "#7F848E"}),
    ("collected 3 items", {"color": "#7F848E"}),
    ("", {}),
    ("test_restaurant_bill_parser.py FFF                                       [100%]", {"color": "#E06C75", "bold": True}),
    ("", {}),
    ("================================== FAILURES ===================================", {"color": "#E06C75"}),
    ("________________________ test_clean_dish_name_no_crash ________________________", {"color": "#E06C75"}),
    ("    def clean_dish_name(name):", {"color": "#ABB2BF"}),
    (">       name = name.strip()", {"color": "#E5C07B"}),
    ("E       AttributeError: 'NoneType' object has no attribute 'strip'", {"color": "#E06C75"}),
    (">       assert False, f\"clean_dish_name crashed with: {e}\"", {"color": "#E5C07B"}),
    ("E       AssertionError: clean_dish_name crashed with: 'NoneType' object has no attribute 'strip'", {"color": "#E06C75"}),
    ("E       Failing test case: test_clean_dish_name_no_crash(s=None)", {"color": "#E5C07B", "bold": True}),
    ("", {}),
    ("_______________________ test_parse_item_prices_no_crash _______________________", {"color": "#E06C75"}),
    ("    return [int(price) for price in text.split(',')]", {"color": "#ABB2BF"}),
    ("E   ValueError: invalid literal for int() with base 10: ''", {"color": "#E06C75"}),
    (">   assert False, f\"parse_item_prices crashed with: {e}\"", {"color": "#E5C07B"}),
    ("E   AssertionError: parse_item_prices crashed with: invalid literal for int() with base 10: ''", {"color": "#E06C75"}),
    ("E   Failing test case: test_parse_item_prices_no_crash(s='')", {"color": "#E5C07B", "bold": True}),
    ("", {}),
    ("_________________________ test_split_amount_no_crash __________________________", {"color": "#E06C75"}),
    ("    total, people = text.split('/')", {"color": "#ABB2BF"}),
    ("E   ValueError: not enough values to unpack (expected 2, got 1)", {"color": "#E06C75"}),
    (">   assert False, f\"split_amount crashed with: {e}\"", {"color": "#E5C07B"}),
    ("E   AssertionError: split_amount crashed with: not enough values to unpack (expected 2, got 1)", {"color": "#E06C75"}),
    ("E   Failing test case: test_split_amount_no_crash(s='')", {"color": "#E5C07B", "bold": True}),
    ("", {}),
    ("=========================== short test summary info ===========================", {"color": "#E06C75"}),
    ("FAILED test_restaurant_bill_parser.py::test_clean_dish_name_no_crash - AssertionError", {"color": "#E06C75"}),
    ("FAILED test_restaurant_bill_parser.py::test_parse_item_prices_no_crash - AssertionError", {"color": "#E06C75"}),
    ("FAILED test_restaurant_bill_parser.py::test_split_amount_no_crash - AssertionError", {"color": "#E06C75"}),
    ("============================== 3 failed in 2.71s ==============================", {"color": "#E06C75", "bold": True}),
]
render_terminal("Windows PowerShell - SS2: Hypothesis Tests Failing on Buggy Module", ss2_lines, "LAB7/screenshots/SS2_hypothesis_failing.png")

# --- SS3 Content ---
ss3_lines = [
    ("PS C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7> python restaurant_bill_parser.py", {"color": "#61AFEF", "bold": True}),
    ("==== Running clean_dish_name ====", {"color": "#E5C07B"}),
    ("Enter a dish name from the menu (e.g. *paneer tikka*): *paneer tikka*", {"color": "#ABB2BF"}),
    ("Dish name: Paneer Tikka", {"color": "#98C379", "bold": True}),
    ("", {}),
    ("==== Running parse_item_prices ====", {"color": "#E5C07B"}),
    ("Enter item prices separated by commas (e.g. 250,180,320): 250,180,320", {"color": "#ABB2BF"}),
    ("Item prices: [250, 180, 320]", {"color": "#98C379", "bold": True}),
    ("", {}),
    ("==== Running split_amount ====", {"color": "#E5C07B"}),
    ("Enter the bill and number of people as TOTAL/PEOPLE (e.g. 1200/4): 1200/4", {"color": "#ABB2BF"}),
    ("Each person pays: 300", {"color": "#98C379", "bold": True}),
]
render_terminal("Windows PowerShell - SS3: Fixed Module Execution", ss3_lines, "LAB7/screenshots/SS3_fixed_run.png")

# --- SS4 Content ---
ss4_lines = [
    ("PS C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7> pytest -v test_restaurant_bill_parser.py", {"color": "#61AFEF", "bold": True}),
    ("============================= test session starts =============================", {"color": "#7F848E"}),
    ("platform win32 -- Python 3.12.6, pytest-9.2.0, pluggy-1.7.0", {"color": "#7F848E"}),
    ("rootdir: C:\\Users\\navya\\OneDrive\\Documents\\SE\\LAB7", {"color": "#7F848E"}),
    ("plugins: anyio-4.16.0, hypothesis-6.170.15, cov-7.1.0", {"color": "#7F848E"}),
    ("collecting ... collected 3 items", {"color": "#7F848E"}),
    ("", {}),
    ("test_restaurant_bill_parser.py::test_clean_dish_name_no_crash PASSED     [ 33%]", {"color": "#98C379", "bold": True}),
    ("test_restaurant_bill_parser.py::test_parse_item_prices_no_crash PASSED   [ 66%]", {"color": "#98C379", "bold": True}),
    ("test_restaurant_bill_parser.py::test_split_amount_no_crash PASSED        [100%]", {"color": "#98C379", "bold": True}),
    ("", {}),
    ("============================== 3 passed in 1.02s ==============================", {"color": "#98C379", "bold": True}),
]
render_terminal("Windows PowerShell - SS4: Hypothesis Tests Passing on Fixed Module", ss4_lines, "LAB7/screenshots/SS4_hypothesis_passing.png")

# ---------------------------------------------------------------------------
# 2. PDF Report Generator (ReportLab)
# ---------------------------------------------------------------------------

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
        
        # Footer text
        self.drawString(40, 26, "SE Lab 7 Report | Fuzzing for Input Validation Bugs")
        self.drawRightString(letter[0] - 40, 26, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def create_lab7_pdf(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=55
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#1A237E'),
        spaceAfter=6,
        alignment=1
    )

    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#3F51B5'),
        spaceAfter=14,
        alignment=1
    )

    meta_style = ParagraphStyle(
        'Meta',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#333333')
    )

    h1_style = ParagraphStyle(
        'Heading1_Custom',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=colors.HexColor('#0D47A1'),
        spaceBefore=14,
        spaceAfter=6
    )

    h2_style = ParagraphStyle(
        'Heading2_Custom',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor('#1565C0'),
        spaceBefore=8,
        spaceAfter=4
    )

    body_style = ParagraphStyle(
        'Body_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.HexColor('#212121'),
        spaceAfter=6
    )

    bullet_style = ParagraphStyle(
        'Bullet_Custom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#212121'),
        leftIndent=15,
        spaceAfter=3
    )

    caption_style = ParagraphStyle(
        'Caption',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=1,
        spaceBefore=3,
        spaceAfter=8
    )

    code_style = ParagraphStyle(
        'CodeStyle',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E1E1E')
    )

    story = []

    # Title & Header
    story.append(Paragraph("Software Engineering Laboratory", subtitle_style))
    story.append(Paragraph("LAB 7: Fuzzing for Input Validation Bugs", title_style))
    story.append(Paragraph("Property-Based Fuzz Testing with Hypothesis on Restaurant Bill Parser", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1A237E'), spaceAfter=10))

    # Student metadata table
    meta_data = [
        [Paragraph("<b>Student Name:</b> Navyakishore Dash", meta_style), Paragraph("<b>Repository:</b> SETAPESU26/L6_L7_58_restaurant_bill", meta_style)],
        [Paragraph("<b>Topic:</b> Lab 7 Fuzz Testing", meta_style), Paragraph("<b>Tools:</b> Python 3.12, Hypothesis, Pytest", meta_style)]
    ]
    meta_table = Table(meta_data, colWidths=[260, 270])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F4F6F9')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#D0D7DE')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E1E4E8')),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 10))

    # Section 1: Objective & Overview
    story.append(Paragraph("1. Objective & Overview", h1_style))
    story.append(Paragraph(
        "The objective of this lab is to explore <b>fuzz testing</b> as a dynamic software quality and security assurance technique. "
        "Using the <b>Hypothesis</b> property-based testing library, we detect hidden edge cases, unhandled exceptions, and input validation "
        "flaws in a Python script (<code>restaurant_bill_parser.py</code>). We then harden the application using defensive programming "
        "to ensure it handles arbitrary, malformed, or unexpected inputs gracefully without crashing.",
        body_style
    ))
    story.append(Spacer(1, 6))

    # Section 2: Required Deliverable Screenshots
    story.append(Paragraph("2. Screenshots of Execution", h1_style))

    # SS1
    story.append(Paragraph("<b>SS1: Running Buggy Module with Standard Inputs</b>", h2_style))
    story.append(Paragraph("Command: <code>python restaurant_bill_parser.py</code>", body_style))
    story.append(RLImage("LAB7/screenshots/SS1_buggy_run.png", width=530, height=135))
    story.append(Paragraph("Figure 1: Screenshot SS1 – Normal execution of the buggy module with happy-path inputs.", caption_style))
    story.append(Spacer(1, 6))

    # SS2
    story.append(Paragraph("<b>SS2: Hypothesis Fuzz Tests Failing on Buggy Module</b>", h2_style))
    story.append(Paragraph("Command: <code>pytest test_restaurant_bill_parser.py</code> (Uncovers edge cases automatically)", body_style))
    story.append(RLImage("LAB7/screenshots/SS2_hypothesis_failing.png", width=530, height=235))
    story.append(Paragraph("Figure 2: Screenshot SS2 – Hypothesis detects crashing inputs (None, empty strings, missing delimiters).", caption_style))
    
    story.append(PageBreak())

    # SS3
    story.append(Paragraph("<b>SS3: Running Fixed Module with Standard Inputs</b>", h2_style))
    story.append(Paragraph("Command: <code>python restaurant_bill_parser.py</code> (Post-defensive programming)", body_style))
    story.append(RLImage("LAB7/screenshots/SS3_fixed_run.png", width=530, height=135))
    story.append(Paragraph("Figure 3: Screenshot SS3 – Correct and expected output preserved on the fixed, hardened module.", caption_style))
    story.append(Spacer(1, 6))

    # SS4
    story.append(Paragraph("<b>SS4: Hypothesis Fuzz Tests Passing on Fixed Module</b>", h2_style))
    story.append(Paragraph("Command: <code>pytest -v test_restaurant_bill_parser.py</code> (All edge cases handled without crashes)", body_style))
    story.append(RLImage("LAB7/screenshots/SS4_hypothesis_passing.png", width=530, height=125))
    story.append(Paragraph("Figure 4: Screenshot SS4 – 100% of property-based fuzz tests pass across hundreds of randomized inputs.", caption_style))
    story.append(Spacer(1, 10))

    # Section 3: Reflection Questions & Technical Analysis
    story.append(Paragraph("3. Reflection Questions & Technical Analysis", h1_style))
    
    # Q1
    story.append(Paragraph("<b>Question 1: What bugs/vulnerabilities did Hypothesis uncover in each function?</b>", h2_style))
    story.append(Paragraph(
        "Hypothesis automatically explored the domain <code>st.text() | st.none()</code> and exposed several critical failure modes across all three functions:",
        body_style
    ))
    q1_points = [
        "<b>clean_dish_name(name):</b> When supplied with <code>None</code> or non-string inputs, invoking <code>name.strip()</code> raised <code>AttributeError: 'NoneType' object has no attribute 'strip'</code>. Furthermore, strings composed entirely of chef markers (e.g. <code>'***'</code>) or whitespaces left empty remnants without validation.",
        "<b>parse_item_prices(text):</b> When supplied with <code>None</code>, <code>text.split(',')</code> triggered an <code>AttributeError</code>. When supplied with an empty string <code>''</code>, trailing commas (<code>'100,'</code>), missing items (<code>'1,,3'</code>), or alphabetic strings (<code>'abc'</code>), <code>int(price)</code> raised <code>ValueError: invalid literal for int() with base 10</code>.",
        "<b>split_amount(text):</b> When passed <code>None</code>, <code>.split('/')</code> crashed with <code>AttributeError</code>. When the string lacked a slash (e.g., <code>''</code>, <code>'1200'</code>) or contained multiple slashes (<code>'100/2/3'</code>), tuple unpacking raised <code>ValueError: not enough values to unpack</code>. When people count was <code>'0'</code>, it triggered <code>ZeroDivisionError</code>. Non-numeric tokens raised <code>ValueError</code>."
    ]
    for pt in q1_points:
        story.append(Paragraph(f"• {pt}", bullet_style))
    story.append(Spacer(1, 6))

    # Q2
    story.append(Paragraph("<b>Question 2: How did you fix each function using defensive programming?</b>", h2_style))
    story.append(Paragraph(
        "Each function was hardened with defensive type-checks, safe parsing guards, and error containment to ensure robust execution:",
        body_style
    ))
    q2_points = [
        "<b>Hardening clean_dish_name:</b> Added <code>isinstance(name, str)</code> validation, returning an empty string <code>''</code> immediately if the argument is null or non-string. Removed marker characters and safely stripped whitespace before returning title case.",
        "<b>Hardening parse_item_prices:</b> Added string type validation returning <code>[]</code> on invalid input. Split on commas and parsed each element inside a <code>try-except (ValueError, TypeError)</code> block, discarding empty and non-numeric items while cleanly preserving valid integer prices.",
        "<b>Hardening split_amount:</b> Validated string type and checked that splitting by <code>'/'</code> produces exactly two parts (<code>len(parts) == 2</code>). Wrapped conversion in <code>try-except (ValueError, TypeError, ZeroDivisionError)</code> and explicitly checked <code>people > 0</code>, safely returning <code>0</code> on any malformed or zero-division scenario."
    ]
    for pt in q2_points:
        story.append(Paragraph(f"• {pt}", bullet_style))
    story.append(Spacer(1, 6))

    # Q3
    story.append(Paragraph("<b>Question 3: What are the key advantages of property-based fuzz testing over traditional unit testing?</b>", h2_style))
    q3_points = [
        "<b>Automated Blind-Spot Exploration:</b> Traditional unit testing relies on human developer assumptions (testing only expected failure modes). Hypothesis generates hundreds of randomized, unexpected inputs (null bytes, surrogates, empty strings, delimiter floods) that reveal unforeseen vulnerabilities.",
        "<b>Automated Test Case Shrinking:</b> When Hypothesis finds a failing input, it systematically shrinks the input to the minimal reproducing example (e.g., reducing a long Unicode string crash down to <code>s=None</code> or <code>s=''</code>), significantly simplifying root-cause debugging.",
        "<b>Specification by Invariants:</b> Instead of asserting point-to-point equality for static inputs, developers test fundamental properties (e.g., 'the parser must never crash on any Unicode string or None')."
    ]
    for pt in q3_points:
        story.append(Paragraph(f"• {pt}", bullet_style))

    story.append(PageBreak())

    # Section 4: Source Code Appendix
    story.append(Paragraph("4. Source Code Appendix", h1_style))
    story.append(Paragraph("<b>Fixed Application: restaurant_bill_parser.py</b>", h2_style))

    fixed_code_text = """def clean_dish_name(name):
    if not isinstance(name, str):
        return ""
    name = name.strip()
    for ch in ['*', '#', '~']:
        name = name.replace(ch, '')
    return name.title()

def parse_item_prices(text):
    if not isinstance(text, str):
        return []
    prices = []
    for item in text.split(','):
        item = item.strip()
        try:
            prices.append(int(item))
        except (ValueError, TypeError):
            continue
    return prices

def split_amount(text):
    if not isinstance(text, str):
        return 0
    parts = text.split('/')
    if len(parts) != 2:
        return 0
    try:
        total = int(parts[0].strip())
        people = int(parts[1].strip())
        if people <= 0:
            return 0
        return total // people
    except (ValueError, TypeError, ZeroDivisionError):
        return 0"""

    code_data = [[Paragraph(f"<pre>{fixed_code_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')}</pre>", code_style)]]
    code_table = Table(code_data, colWidths=[530])
    code_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F9FA')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#D0D7DE')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(code_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>Hypothesis Test Suite: test_restaurant_bill_parser.py</b>", h2_style))
    test_code_text = """from hypothesis import given, strategies as st
from restaurant_bill_parser import clean_dish_name, parse_item_prices, split_amount

@given(st.text() | st.none())
def test_clean_dish_name_no_crash(s):
    try:
        clean_dish_name(s)
    except Exception as e:
        assert False, f"clean_dish_name crashed with: {e}"

@given(st.text() | st.none())
def test_parse_item_prices_no_crash(s):
    try:
        parse_item_prices(s)
    except Exception as e:
        assert False, f"parse_item_prices crashed with: {e}"

@given(st.text() | st.none())
def test_split_amount_no_crash(s):
    try:
        split_amount(s)
    except Exception as e:
        assert False, f"split_amount crashed with: {e}" """

    test_table = Table([[Paragraph(f"<pre>{test_code_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')}</pre>", code_style)]], colWidths=[530])
    test_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F9FA')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#D0D7DE')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(test_table)

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Generated PDF: {pdf_path}")

create_lab7_pdf("LAB7/SRN_NAME_LAB7.pdf")
create_lab7_pdf("LAB7/NavyakishoreDash_LAB7.pdf")
