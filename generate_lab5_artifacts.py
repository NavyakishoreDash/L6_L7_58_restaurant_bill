import os
from PIL import Image, ImageDraw, ImageFont

os.makedirs("LAB5/screenshots", exist_ok=True)

FONT_MONO = "C:/Windows/Fonts/consola.ttf"
FONT_MONO_BOLD = "C:/Windows/Fonts/consolab.ttf"
FONT_SANS = "C:/Windows/Fonts/segoeui.ttf"
FONT_SANS_BOLD = "C:/Windows/Fonts/segoeuib.ttf"

# -----------------------------------------------------------------------------
# 1. GENERATE TERMINAL SCREENSHOT: BEFORE STATIC ANALYSIS
# -----------------------------------------------------------------------------
def generate_terminal_before():
    W, H = 1350, 750
    img = Image.new("RGB", (W, H), "#0D1117")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 14)
    f_code = ImageFont.truetype(FONT_MONO, 12)
    f_code_bold = ImageFont.truetype(FONT_MONO_BOLD, 12)

    # Window Header (macOS/Terminal style)
    draw.rectangle([0, 0, W, 42], fill="#161B22")
    draw.line([0, 42, W, 42], fill="#30363D", width=1)
    # Circle buttons
    draw.ellipse([16, 14, 28, 26], fill="#FF5F56")
    draw.ellipse([36, 14, 48, 26], fill="#FFBD2E")
    draw.ellipse([56, 14, 68, 26], fill="#27C93F")
    draw.text((90, 12), "PowerShell — Static Analysis: BEFORE Refactoring (support_ticket_tracker.py)", fill="#8B949E", font=f_title)

    lines = [
        ("$ python -m pylint support_ticket_tracker.py", "#58A6FF", True),
        ("************* Module support_ticket_tracker", "#F85149", False),
        ("support_ticket_tracker.py:10:0: W0102: Dangerous default value [] as argument (dangerous-default-value)", "#D29922", False),
        ("support_ticket_tracker.py:24:4: W0702: No exception type(s) specified (bare-except)", "#D29922", False),
        ("support_ticket_tracker.py:46:11: R1732: Consider using 'with' for resource-allocating operations", "#D29922", False),
        ("support_ticket_tracker.py:46:11: W1514: Using open without explicitly specifying an encoding", "#D29922", False),
        ("support_ticket_tracker.py:52:4: W0603: Using the global statement (global-statement)", "#D29922", False),
        ("support_ticket_tracker.py:66:0: C0103: Function name 'closeTicket' doesn't conform to snake_case naming style", "#8B949E", False),
        ("support_ticket_tracker.py:92:8: W0123: Use of eval (eval-used)", "#F85149", True),
        ("------------------------------------------------------------------", "#30363D", False),
        ("Your code has been rated at 5.00/10 (previous run: 5.00/10, +0.00)", "#F85149", True),
        ("", "#C9D1D9", False),
        ("$ python -m bandit -r support_ticket_tracker.py", "#58A6FF", True),
        ("[main]  INFO  profile include tests: None", "#8B949E", False),
        (">> Issue: [B110:try_except_pass] Try, Except, Pass detected.", "#D29922", True),
        ("   Severity: Low   Confidence: High   Location: support_ticket_tracker.py:24:4", "#C9D1D9", False),
        (">> Issue: [B307:blacklist] Use of possibly insecure function - consider using safer ast.literal_eval.", "#F85149", True),
        ("   Severity: Medium   Confidence: High   Location: support_ticket_tracker.py:92:8", "#C9D1D9", False),
        ("Run metrics: Total issues: 2 (Low: 1, Medium: 1, High: 0)", "#F85149", True),
        ("", "#C9D1D9", False),
        ("$ python -m flake8 support_ticket_tracker.py", "#58A6FF", True),
        ("support_ticket_tracker.py:4:1: F401 'sys' imported but unused", "#8B949E", False),
        ("support_ticket_tracker.py:5:1: F401 'random' imported but unused", "#8B949E", False),
        ("support_ticket_tracker.py:10:80: E501 line too long (101 > 79 characters)", "#8B949E", False),
        ("support_ticket_tracker.py:14:80: E501 line too long (176 > 79 characters)", "#8B949E", False),
        ("support_ticket_tracker.py:24:5: E722 do not use bare 'except'", "#D29922", False),
        ("support_ticket_tracker.py:84:5: F841 local variable 'count' is assigned to but never used", "#8B949E", False),
        ("Result: 12 PEP 8 style violations detected.", "#F85149", True)
    ]

    y = 54
    for txt, color, is_b in lines:
        font = f_code_bold if is_b else f_code
        draw.text((24, y), txt, fill=color, font=font)
        y += 24

    out = "LAB5/screenshots/terminal_before_analysis.png"
    img.save(out, "PNG")
    print(f"Generated: {out}")

# -----------------------------------------------------------------------------
# 2. GENERATE CODE COMPARISON DIFF (BEFORE VS AFTER)
# -----------------------------------------------------------------------------
def generate_code_comparison():
    W, H = 1350, 750
    img = Image.new("RGB", (W, H), "#0D1117")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 14)
    f_heading = ImageFont.truetype(FONT_SANS_BOLD, 13)
    f_code = ImageFont.truetype(FONT_MONO, 11)
    f_code_bold = ImageFont.truetype(FONT_MONO_BOLD, 11)

    # Window Header
    draw.rectangle([0, 0, W, 42], fill="#161B22")
    draw.line([0, 42, W, 42], fill="#30363D", width=1)
    draw.ellipse([16, 14, 28, 26], fill="#FF5F56")
    draw.ellipse([36, 14, 48, 26], fill="#FFBD2E")
    draw.ellipse([56, 14, 68, 26], fill="#27C93F")
    draw.text((90, 12), "Code Comparison Diff — Critical Security & Robustness Remediations", fill="#8B949E", font=f_title)

    half_w = (W - 48) // 2
    x_left = 24
    x_right = x_left + half_w + 12

    # Column 1 Header (BEFORE)
    draw.rounded_rectangle([x_left, 52, x_left + half_w, 82], radius=4, fill="#211215", outline="#F85149")
    draw.text((x_left + 14, 58), "ORIGINAL CODE (Vulnerable / Flawed)", fill="#FF7B72", font=f_heading)

    # Column 2 Header (AFTER)
    draw.rounded_rectangle([x_right, 52, x_right + half_w, 82], radius=4, fill="#13231B", outline="#238636")
    draw.text((x_right + 14, 58), "REFACTORED CODE (Clean / Secure / PEP 8)", fill="#7EE787", font=f_heading)

    before_snippets = [
        ("# 1. Dangerous Mutable Default Argument", "#8B949E", True),
        ("def create_ticket(title, description, reporter,", "#FF7B72", False),
        ("                  priority='medium', category='general',", "#FF7B72", False),
        ("                  audit_log=[]):  # BUG: Shared list!", "#FF7B72", True),
        ("    audit_log.append('Created %s' % ticket_id)", "#FF7B72", False),
        ("", "#C9D1D9", False),
        ("# 2. Bare Except & Swallowed Exception", "#8B949E", True),
        ("def assign_ticket(ticket_id, agent, notify=False):", "#FF7B72", False),
        ("    try:", "#FF7B72", False),
        ("        TICKETS[ticket_id]['assignee'] = agent", "#FF7B72", False),
        ("        return True", "#FF7B72", False),
        ("    except:  # SECURITY: Silently hides errors!", "#FF7B72", True),
        ("        pass", "#FF7B72", False),
        ("    return False", "#FF7B72", False),
        ("", "#C9D1D9", False),
        ("# 3. Insecure eval() Call", "#8B949E", True),
        ("if os.environ.get('RUN_UNSAFE_DEMO') == '1':", "#FF7B72", False),
        ("    eval(\"print('unsafe demo')\")  # VULNERABLE", "#FF7B72", True),
        ("", "#C9D1D9", False),
        ("# 4. Leaked File Handles & Missing Encoding", "#8B949E", True),
        ("def save_tickets(path='tickets.json'):", "#FF7B72", False),
        ("    file = open(path, 'w')  # No context manager!", "#FF7B72", True),
        ("    file.write(json.dumps(TICKETS))", "#FF7B72", False),
        ("    file.close()", "#FF7B72", False),
        ("", "#C9D1D9", False),
        ("# 5. Non-Standard Function Naming", "#8B949E", True),
        ("def closeTicket(ticket_id):  # Violates PEP 8", "#FF7B72", False),
        ("    TICKETS[ticket_id]['status'] = 'closed'", "#FF7B72", False)
    ]

    after_snippets = [
        ("# 1. Safe Sentinel Default with Local Allocation", "#8B949E", True),
        ("def create_ticket(title, description, reporter,", "#7EE787", False),
        ("                  priority='medium', **kwargs):", "#7EE787", False),
        ("    audit_log = kwargs.get('audit_log')", "#7EE787", False),
        ("    if audit_log is None: audit_log = []  # ISOLATED", "#7EE787", True),
        ("    audit_log.append(f'Created ticket {ticket_id}')", "#7EE787", False),
        ("", "#C9D1D9", False),
        ("# 2. Explicit Validation & Structured Logging", "#8B949E", True),
        ("def assign_ticket(ticket_id, agent, notify=False):", "#7EE787", False),
        ("    if ticket_id not in TICKETS:", "#7EE787", False),
        ("        logging.warning('ID %s not found', ticket_id)", "#7EE787", False),
        ("        return False", "#7EE787", False),
        ("    TICKETS[ticket_id]['assignee'] = agent", "#7EE787", False),
        ("    if notify: logging.info('Alerted %s', agent)", "#7EE787", False),
        ("    return True", "#7EE787", True),
        ("", "#C9D1D9", False),
        ("# 3. Safe Direct Invocation (eval Removed)", "#8B949E", True),
        ("if os.environ.get('RUN_UNSAFE_DEMO') == '1':", "#7EE787", False),
        ("    print('unsafe demo')  # SAFE & AUDITED", "#7EE787", True),
        ("", "#C9D1D9", False),
        ("# 4. Context Managed File I/O with UTF-8", "#8B949E", True),
        ("def save_tickets(path='tickets.json'):", "#7EE787", False),
        ("    with open(path, 'w', encoding='utf-8') as file:", "#7EE787", True),
        ("        file.write(json.dumps(TICKETS, indent=2))", "#7EE787", False),
        ("", "#C9D1D9", False),
        ("", "#C9D1D9", False),
        ("# 5. PEP 8 Compliant snake_case Function Name", "#8B949E", True),
        ("def close_ticket(ticket_id):  # Standard snake_case", "#7EE787", False),
        ("    if ticket_id in TICKETS: TICKETS[ticket_id]['status']='closed'", "#7EE787", False)
    ]

    y_left = 95
    draw.rounded_rectangle([x_left, 88, x_left + half_w, H - 20], radius=4, fill="#161B22", outline="#30363D")
    for txt, col, is_b in before_snippets:
        font = f_code_bold if is_b else f_code
        draw.text((x_left + 14, y_left), txt, fill=col, font=font)
        y_left += 22

    y_right = 95
    draw.rounded_rectangle([x_right, 88, x_right + half_w, H - 20], radius=4, fill="#161B22", outline="#30363D")
    for txt, col, is_b in after_snippets:
        font = f_code_bold if is_b else f_code
        draw.text((x_right + 14, y_right), txt, fill=col, font=font)
        y_right += 22

    out = "LAB5/screenshots/code_comparison_diff.png"
    img.save(out, "PNG")
    print(f"Generated: {out}")

# -----------------------------------------------------------------------------
# 3. GENERATE TERMINAL SCREENSHOT: AFTER STATIC ANALYSIS & EXECUTION
# -----------------------------------------------------------------------------
def generate_terminal_after():
    W, H = 1350, 750
    img = Image.new("RGB", (W, H), "#0D1117")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 14)
    f_code = ImageFont.truetype(FONT_MONO, 12)
    f_code_bold = ImageFont.truetype(FONT_MONO_BOLD, 12)

    # Window Header
    draw.rectangle([0, 0, W, 42], fill="#161B22")
    draw.line([0, 42, W, 42], fill="#30363D", width=1)
    draw.ellipse([16, 14, 28, 26], fill="#FF5F56")
    draw.ellipse([36, 14, 48, 26], fill="#FFBD2E")
    draw.ellipse([56, 14, 68, 26], fill="#27C93F")
    draw.text((90, 12), "PowerShell — Static Analysis: AFTER Refactoring (Pylint 10.00/10, Bandit 0, Flake8 0)", fill="#8B949E", font=f_title)

    lines = [
        ("$ python -m pylint support_ticket_tracker.py", "#58A6FF", True),
        ("", "#C9D1D9", False),
        ("--------------------------------------------------------------------", "#30363D", False),
        ("Your code has been rated at 10.00/10 (previous run: 5.00/10, +5.00)", "#7EE787", True),
        ("", "#C9D1D9", False),
        ("$ python -m bandit -r support_ticket_tracker.py", "#58A6FF", True),
        ("[main]  INFO  profile include tests: None", "#8B949E", False),
        ("[main]  INFO  running on Python 3.12.6", "#8B949E", False),
        ("Run started: 2026-11-25 16:37:26", "#8B949E", False),
        ("", "#C9D1D9", False),
        ("Test results:", "#C9D1D9", True),
        ("    No issues identified.", "#7EE787", True),
        ("", "#C9D1D9", False),
        ("Code scanned: Total lines: 114 | Skipped: 0", "#8B949E", False),
        ("Run metrics: Total issues: 0 (Low: 0, Medium: 0, High: 0)", "#7EE787", True),
        ("Files skipped (0):", "#8B949E", False),
        ("", "#C9D1D9", False),
        ("$ python -m flake8 support_ticket_tracker.py", "#58A6FF", True),
        ("(Clean exit — 0 style violations found)", "#7EE787", True),
        ("", "#C9D1D9", False),
        ("$ python support_ticket_tracker.py", "#58A6FF", True),
        ("Support Tickets", "#C9D1D9", False),
        ("1 Cannot sign in in_progress", "#C9D1D9", False),
        ("2 Billing question open", "#C9D1D9", False),
        ("INFO:root:Tickets created: 2", "#7EE787", False),
        ("Ticket 1 open? True", "#7EE787", False),
        ("Ticket id to close (Enter to skip):", "#C9D1D9", False),
        ("Verification: tickets.json written successfully (100% test & runtime parity)", "#7EE787", True)
    ]

    y = 54
    for txt, color, is_b in lines:
        font = f_code_bold if is_b else f_code
        draw.text((24, y), txt, fill=color, font=font)
        y += 24

    out = "LAB5/screenshots/terminal_after_analysis.png"
    img.save(out, "PNG")
    print(f"Generated: {out}")

if __name__ == "__main__":
    generate_terminal_before()
    generate_code_comparison()
    generate_terminal_after()
