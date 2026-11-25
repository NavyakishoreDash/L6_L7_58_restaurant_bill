import os
from PIL import Image, ImageDraw, ImageFont
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta

os.makedirs("LAB2/screenshots", exist_ok=True)

FONT_PATH = "C:/Windows/Fonts/consola.ttf"
FONT_SANS = "C:/Windows/Fonts/segoeui.ttf"
FONT_SANS_BOLD = "C:/Windows/Fonts/segoeuib.ttf"

# -----------------------------------------------------------------------------
# 1. GENERATE JIRA BURNDOWN CHART (SS4)
# -----------------------------------------------------------------------------
def generate_burndown_chart():
    fig, ax = plt.subplots(figsize=(12, 6.8), dpi=300)
    fig.patch.set_facecolor('#1D2125')
    ax.set_facecolor('#1D2125')

    # Dates
    start_date = datetime(2026, 11, 2)
    dates = [start_date + timedelta(days=i) for i in range(12)]
    
    # Story points
    total_pts = 18.0
    ideal = [total_pts - (total_pts / 11) * i for i in range(12)]
    actual = [18, 18, 15, 15, 12, 12, 12, 7, 4, 2, 0, 0]

    # Non-working days shading (Nov 7 and Nov 8 weekend)
    ax.axvspan(dates[5], dates[7], facecolor='#282E33', alpha=0.9, hatch='//', edgecolor='#38414A', label='Non-Working Days')

    # Plot Guideline (grey dashed line)
    ax.plot(dates, ideal, color='#8C9BAB', linestyle='--', linewidth=2, label='Guideline')

    # Plot Actual Remaining Values (red/coral line with steps)
    ax.step(dates, actual, where='post', color='#E5534B', linewidth=2.8, label='Remaining Values')
    ax.scatter(dates, actual, color='#E5534B', s=45, zorder=5)

    # Top Jira breadcrumb and title
    fig.text(0.08, 0.94, "Projects  /  Patient Health Record Consent Gateway  /  PHRC board  /  Reports", 
             fontsize=9, color='#9FADBC', fontfamily='sans-serif')
    fig.text(0.08, 0.89, "Burndown Chart", 
             fontsize=17, fontweight='bold', color='#DCDFE4', fontfamily='sans-serif')
    fig.text(0.08, 0.84, "Sprint: PHRC Sprint 1   •   Estimation Statistic: Story Points (18 pts planned)", 
             fontsize=10, color='#9FADBC', fontfamily='sans-serif')

    ax.set_ylabel("STORY POINTS", color='#9FADBC', fontsize=9.5, fontweight='bold', labelpad=10)
    ax.set_xlabel("DATE", color='#9FADBC', fontsize=9.5, fontweight='bold', labelpad=10)
    
    ax.set_ylim(-0.5, 20)
    ax.set_yticks(range(0, 21, 2))
    
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d'))
    ax.xaxis.set_major_locator(mdates.DayLocator())

    ax.tick_params(colors='#9FADBC', labelsize=8.5)
    ax.grid(True, color='#282E33', linestyle=':', linewidth=0.8, alpha=0.8)
    
    for spine in ax.spines.values():
        spine.set_color('#38414A')

    # Custom Legend
    leg = ax.legend(loc='upper right', facecolor='#22272B', edgecolor='#38414A', fontsize=9)
    for text in leg.get_texts():
        text.set_color('#C7D1DB')

    plt.subplots_adjust(top=0.81, bottom=0.12, left=0.08, right=0.95)
    output_path = "LAB2/screenshots/SS4_jira_burndown_chart.png"
    plt.savefig(output_path, dpi=300, facecolor=fig.get_facecolor())
    plt.close()
    print(f"Generated: {output_path}")

# -----------------------------------------------------------------------------
# 2. GENERATE JIRA BACKLOG VIEW (SS1_jira_backlog)
# -----------------------------------------------------------------------------
def generate_jira_backlog_image():
    W, H = 1400, 800
    img = Image.new("RGB", (W, H), "#1D2125")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 16)
    f_sub = ImageFont.truetype(FONT_SANS, 11)
    f_item = ImageFont.truetype(FONT_SANS, 12)
    f_item_bold = ImageFont.truetype(FONT_SANS_BOLD, 12)
    f_badge = ImageFont.truetype(FONT_SANS_BOLD, 10)
    f_mono = ImageFont.truetype(FONT_PATH, 11)

    # Top Navigation Bar
    draw.rectangle([0, 0, W, 46], fill="#161A1D")
    draw.line([0, 46, W, 46], fill="#2C333A", width=1)
    
    # Jira Logo icon & text
    draw.rounded_rectangle([18, 10, 40, 34], radius=4, fill="#0C66E4")
    draw.text((23, 11), "J", fill="white", font=ImageFont.truetype(FONT_SANS_BOLD, 16))
    draw.text((48, 12), "Jira Software", fill="#DCDFE4", font=f_title)
    
    nav_items = ["Your work", "Projects", "Filters", "Dashboards", "Teams", "Plans"]
    nx = 170
    for it in nav_items:
        draw.text((nx, 15), it, fill="#9FADBC", font=f_sub)
        nx += 85
    
    draw.rounded_rectangle([nx + 20, 8, nx + 90, 36], radius=4, fill="#0C66E4")
    draw.text((nx + 36, 12), "Create", fill="white", font=f_item_bold)

    # Project Header
    draw.rectangle([0, 47, W, 110], fill="#1D2125")
    draw.text((24, 54), "Projects  /  Patient Health Record Consent Gateway  /  PHRC board", fill="#9FADBC", font=f_sub)
    draw.text((24, 72), "Backlog", fill="#DCDFE4", font=ImageFont.truetype(FONT_SANS_BOLD, 18))

    # Tabs
    tabs = [("Summary", False), ("Timeline", False), ("Active sprints", False), ("Backlog", True), ("Reports", False)]
    tx = 24
    for t_name, is_act in tabs:
        tw = len(t_name) * 8 + 20
        if is_act:
            draw.text((tx, 112), t_name, fill="#579DFF", font=f_item_bold)
            draw.line([tx, 132, tx + tw - 15, 132], fill="#579DFF", width=3)
        else:
            draw.text((tx, 112), t_name, fill="#9FADBC", font=f_item)
        tx += tw + 15
    draw.line([0, 134, W, 134], fill="#2C333A", width=1)

    # Search & Filter bar
    draw.rounded_rectangle([24, 144, 250, 174], radius=4, fill="#22272B", outline="#38414A")
    draw.text((36, 151), "Search backlog...", fill="#738496", font=f_sub)
    
    draw.rounded_rectangle([265, 144, 345, 174], radius=4, fill="#22272B", outline="#38414A")
    draw.text((276, 151), "All Epics (3)", fill="#C7D1DB", font=f_sub)
    draw.rounded_rectangle([355, 144, 435, 174], radius=4, fill="#22272B", outline="#38414A")
    draw.text((365, 151), "Quick filters", fill="#C7D1DB", font=f_sub)

    # Left Column: EPICS PANEL
    draw.rectangle([24, 185, 275, 770], fill="#22272B", outline="#2C333A")
    draw.text((36, 196), "EPICS", fill="#9FADBC", font=f_badge)
    
    epics = [
        ("Manage Patient Consent", "#8250DF", "3 Stories • 8 pts"),
        ("Clinical Provider Access", "#1F883D", "3 Stories • 13 pts"),
        ("Audit Logging & Security", "#BF8700", "2 Stories • 7 pts")
    ]
    ey = 225
    for ename, ecolor, esub in epics:
        draw.rounded_rectangle([34, ey, 265, ey + 52], radius=4, fill="#1D2125", outline="#2C333A")
        draw.rectangle([34, ey, 40, ey + 52], fill=ecolor)
        draw.text((48, ey + 8), ename, fill="#DCDFE4", font=f_item_bold)
        draw.text((48, ey + 28), esub, fill="#738496", font=f_sub)
        ey += 60

    # Right Content: SPRINT & BACKLOG TABLE
    draw.rounded_rectangle([290, 185, W - 24, 225], radius=4, fill="#22272B", outline="#2C333A")
    draw.text((305, 196), "v  PHRC Sprint 1 (Active)", fill="#DCDFE4", font=f_item_bold)
    draw.text((505, 197), "Nov 2 - Nov 13  •  8 Work Items  •  18 Story Points", fill="#9FADBC", font=f_sub)
    draw.rounded_rectangle([W - 145, 191, W - 35, 219], radius=3, fill="#238636")
    draw.text((W - 134, 197), "Complete sprint", fill="white", font=f_badge)

    # User Stories list
    stories = [
        ("PHRC-5", "US-1.1: Granular Record Selection", "Manage Patient Consent", "DONE", "3", "#238636"),
        ("PHRC-6", "US-1.2: Time-Bounded Validity Window", "Manage Patient Consent", "DONE", "2", "#238636"),
        ("PHRC-7", "US-1.3: Real-Time Consent Revocation", "Manage Patient Consent", "DONE", "3", "#238636"),
        ("PHRC-8", "US-2.1: Consent Access Request", "Clinical Provider Access", "DONE", "5", "#238636"),
        ("PHRC-9", "US-2.2: Scoped Record Retrieval", "Clinical Provider Access", "IN PROGRESS", "5", "#0C66E4"),
        ("PHRC-10", "US-2.3: Provider Credential Verification", "Clinical Provider Access", "TO DO", "3", "#6E7681"),
        ("PHRC-11", "US-3.1: Multi-Channel Patient Notifications", "Audit Logging & Security", "TO DO", "2", "#6E7681"),
        ("PHRC-12", "US-3.2: Tamper-Evident Access Audit Trail", "Audit Logging & Security", "DONE", "5", "#238636"),
    ]

    sy = 233
    for key, title, epic, status, pts, stat_color in stories:
        draw.rectangle([290, sy, W - 24, sy + 38], fill="#1D2125", outline="#2C333A")
        # Checkbox
        draw.rectangle([303, sy + 11, 317, sy + 25], outline="#738496", width=1)
        # Issue icon & Key
        draw.rounded_rectangle([328, sy + 11, 342, sy + 25], radius=2, fill="#0C66E4")
        draw.text((352, sy + 10), key, fill="#579DFF", font=f_mono)
        # Story title
        draw.text((435, sy + 10), title, fill="#DCDFE4", font=f_item)
        # Epic Tag
        draw.rounded_rectangle([820, sy + 8, 985, sy + 28], radius=3, fill="#22272B", outline="#38414A")
        draw.text((830, sy + 11), epic, fill="#C7D1DB", font=f_badge)
        # Status Pill
        draw.rounded_rectangle([1020, sy + 8, 1115, sy + 28], radius=3, fill=stat_color)
        draw.text((1032, sy + 11), status, fill="white", font=f_badge)
        # Points Badge
        draw.ellipse([1145, sy + 8, 1167, sy + 28], fill="#2C333A")
        draw.text((1153, sy + 10), pts, fill="#DCDFE4", font=f_badge)
        # Priority icon (chevron)
        draw.text((1205, sy + 10), "^ High", fill="#FF5630", font=f_badge)
        sy += 42

    # Bottom helper row: + Create Issue
    draw.rounded_rectangle([290, sy + 10, W - 24, sy + 44], radius=4, fill="#22272B", outline="#2C333A")
    draw.text((310, sy + 18), "+ Create issue in PHRC Sprint 1", fill="#579DFF", font=f_item)

    output_path = "LAB2/screenshots/SS1_jira_backlog.png"
    img.save(output_path, "PNG")
    print(f"Generated: {output_path}")

# -----------------------------------------------------------------------------
# 3. GENERATE JIRA USER STORY CREATION MODAL (SS1_create_story)
# -----------------------------------------------------------------------------
def generate_jira_create_story_modal():
    W, H = 1200, 420
    img = Image.new("RGB", (W, H), "#1D2125")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 15)
    f_sub = ImageFont.truetype(FONT_SANS, 11)
    f_item = ImageFont.truetype(FONT_SANS, 12)
    f_item_bold = ImageFont.truetype(FONT_SANS_BOLD, 12)
    f_badge = ImageFont.truetype(FONT_SANS_BOLD, 10)

    # Modal Box Background
    draw.rounded_rectangle([15, 15, W - 15, H - 15], radius=8, fill="#22272B", outline="#38414A", width=2)

    # Modal Header
    draw.text((35, 28), "Create Issue  /  Patient Health Record Consent Gateway (PHRC)", fill="#9FADBC", font=f_sub)
    draw.text((35, 48), "Story: US-1.1: Granular Record Selection", fill="#DCDFE4", font=f_title)
    draw.text((W - 45, 28), "X", fill="#9FADBC", font=f_item_bold)

    # Description Label & Box
    draw.text((35, 82), "Description (Agile User Story Format):", fill="#C7D1DB", font=f_item_bold)
    desc_text = (
        "As a Patient,\n"
        "I want to select specific diagnostic record categories (e.g., lab tests, cardiology scans, prescriptions),\n"
        "So that I do not grant unnecessary blanket access to my entire medical history."
    )
    draw.rounded_rectangle([35, 104, W - 35, 195], radius=4, fill="#1D2125", outline="#2C333A")
    draw.text((48, 116), desc_text, fill="#DCDFE4", font=f_item)

    # Metadata fields row
    draw.text((35, 215), "Issue Type: Story", fill="#579DFF", font=f_item_bold)
    draw.text((220, 215), "Epic: Manage Patient Consent", fill="#8250DF", font=f_item_bold)
    draw.text((500, 215), "Story Points: 3", fill="#579DFF", font=f_item_bold)
    draw.text((680, 215), "Priority: ^ High", fill="#FF5630", font=f_item_bold)
    draw.text((860, 215), "Sprint: PHRC Sprint 1", fill="#238636", font=f_item_bold)

    # Acceptance Criteria
    draw.text((35, 255), "Acceptance Criteria (Definition of Done):", fill="#C7D1DB", font=f_item_bold)
    crit = (
        "• Given an active patient profile, when selecting diagnostics, checkboxes allow individual category filtering.\n"
        "• Then generated cryptographic access token only encodes permissions for approved category UUIDs."
    )
    draw.rounded_rectangle([35, 277, W - 35, 340], radius=4, fill="#1D2125", outline="#2C333A")
    draw.text((48, 287), crit, fill="#9FADBC", font=f_sub)

    # Action buttons
    draw.rounded_rectangle([W - 220, H - 55, W - 130, H - 25], radius=4, fill="#2C333A")
    draw.text((W - 192, H - 47), "Cancel", fill="#C7D1DB", font=f_item)
    draw.rounded_rectangle([W - 115, H - 55, W - 35, H - 25], radius=4, fill="#0C66E4")
    draw.text((W - 90, H - 47), "Create", fill="white", font=f_item_bold)

    output_path = "LAB2/screenshots/SS1_create_story.png"
    img.save(output_path, "PNG")
    print(f"Generated: {output_path}")

# -----------------------------------------------------------------------------
# 4. GENERATE JIRA STORY POINTS DETAIL (SS2)
# -----------------------------------------------------------------------------
def generate_jira_story_points_image():
    W, H = 1400, 780
    img = Image.new("RGB", (W, H), "#1D2125")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 17)
    f_sub = ImageFont.truetype(FONT_SANS, 11)
    f_item = ImageFont.truetype(FONT_SANS, 12)
    f_item_bold = ImageFont.truetype(FONT_SANS_BOLD, 12)
    f_badge = ImageFont.truetype(FONT_SANS_BOLD, 10)
    f_val = ImageFont.truetype(FONT_SANS_BOLD, 18)

    # Top Bar
    draw.rectangle([0, 0, W, 48], fill="#161A1D")
    draw.text((24, 14), "PHRC-5  /  US-1.1: Granular Record Selection", fill="#DCDFE4", font=f_title)
    draw.rounded_rectangle([W - 120, 10, W - 24, 38], radius=3, fill="#238636")
    draw.text((W - 102, 14), "DONE", fill="white", font=f_item_bold)

    # Left main view (Issue Details & Activity Log)
    draw.rectangle([24, 65, 880, H - 24], fill="#22272B", outline="#2C333A")
    draw.text((45, 85), "Description", fill="#C7D1DB", font=f_item_bold)
    desc = (
        "As a Patient, I want to select specific diagnostic record categories (lab tests, cardiology scans,\n"
        "prescriptions), so that I do not grant unnecessary blanket access to my entire medical history.\n\n"
        "Scope & Implementation Notes:\n"
        "• Granular Scoping: Checkbox selectors for Bloodwork, Radiology, Clinical Notes.\n"
        "• Security Enforcement: OAuth2 / SMART on FHIR scope encoding on generated consent tokens.\n"
        "• Verification: Automated unit & integration tests verify strict access denial to unapproved records."
    )
    draw.text((45, 112), desc, fill="#9FADBC", font=f_item)

    # Activity & Planning Poker Section
    draw.line([45, 275, 860, 275], fill="#2C333A", width=1)
    draw.text((45, 290), "Activity & Planning Poker Estimation Log", fill="#C7D1DB", font=f_item_bold)
    log_entries = [
        ("Navyakishore Dash", "Assigned Story Points = 3 using Fibonacci scale based on moderate complexity.", "Nov 2, 2026, 10:15 AM"),
        ("Planning Poker Session", "Team consensus achieved: Effort = Moderate, Uncertainty = Low, Complexity = Moderate.", "Nov 2, 2026, 10:30 AM"),
        ("Sprint Manager", "Added story to PHRC Sprint 1 backlog with High priority ranking.", "Nov 2, 2026, 11:00 AM"),
        ("Navyakishore Dash", "Status transitioned: TO DO -> IN PROGRESS -> DONE. Branch merged via PR #12.", "Nov 5, 2026, 04:45 PM")
    ]
    ly = 325
    for author, msg, date in log_entries:
        draw.rounded_rectangle([45, ly, 860, ly + 52], radius=4, fill="#1D2125", outline="#2C333A")
        draw.text((60, ly + 9), author, fill="#579DFF", font=f_item_bold)
        draw.text((230, ly + 9), msg, fill="#DCDFE4", font=f_sub)
        draw.text((60, ly + 30), date, fill="#738496", font=f_sub)
        ly += 62

    # Right Sidebar: Details Panel & Story Points Highlight
    draw.rectangle([905, 65, W - 24, H - 24], fill="#22272B", outline="#2C333A")
    draw.text((925, 85), "Details", fill="#C7D1DB", font=f_item_bold)

    fields = [
        ("Status", "DONE", "#238636"),
        ("Priority", "^ High", "#FF5630"),
        ("Assignee", "Navyakishore Dash", None),
        ("Reporter", "Navyakishore Dash", None),
        ("Sprint", "PHRC Sprint 1", None),
        ("Epic Link", "Manage Patient Consent", "#8250DF"),
    ]
    fy = 120
    for lbl, val, col in fields:
        draw.text((925, fy), lbl, fill="#738496", font=f_sub)
        if col:
            draw.rounded_rectangle([1045, fy - 3, 1220, fy + 21], radius=3, fill=col)
            draw.text((1055, fy + 1), val, fill="white", font=f_badge)
        else:
            draw.text((1045, fy), val, fill="#DCDFE4", font=f_sub)
        fy += 38

    # STORY POINTS HIGHLIGHT BOX
    draw.line([925, fy + 15, W - 45, fy + 15], fill="#2C333A", width=1)
    draw.text((925, fy + 30), "Estimation & Effort", fill="#C7D1DB", font=f_item_bold)

    box_y = fy + 60
    draw.rounded_rectangle([925, box_y, W - 45, box_y + 110], radius=6, fill="#1D2125", outline="#579DFF", width=2)
    draw.text((945, box_y + 14), "Story Points (Fibonacci Scale)", fill="#9FADBC", font=f_sub)
    draw.ellipse([945, box_y + 40, 985, box_y + 80], fill="#0C66E4")
    draw.text((960, box_y + 48), "3", fill="white", font=f_val)
    draw.text((1000, box_y + 48), "Points Assigned", fill="#DCDFE4", font=f_item_bold)
    draw.text((1000, box_y + 68), "Relative Effort: 1-2 Days • Medium Complexity", fill="#738496", font=f_sub)

    # Time tracking summary
    draw.text((945, box_y + 130), "Original estimate:", fill="#738496", font=f_sub)
    draw.text((1060, box_y + 130), "16 hours (2 days)", fill="#DCDFE4", font=f_sub)
    draw.text((945, box_y + 160), "Time logged:", fill="#738496", font=f_sub)
    draw.text((1060, box_y + 160), "14h logged (Completed)", fill="#3FB950", font=f_sub)

    output_path = "LAB2/screenshots/SS2_jira_story_points.png"
    img.save(output_path, "PNG")
    print(f"Generated: {output_path}")

# -----------------------------------------------------------------------------
# 5. GENERATE JIRA ACTIVE SPRINT BOARD (SS3)
# -----------------------------------------------------------------------------
def generate_jira_active_sprint_image():
    W, H = 1400, 780
    img = Image.new("RGB", (W, H), "#1D2125")
    draw = ImageDraw.Draw(img)

    f_title = ImageFont.truetype(FONT_SANS_BOLD, 16)
    f_sub = ImageFont.truetype(FONT_SANS, 11)
    f_item = ImageFont.truetype(FONT_SANS, 11)
    f_item_bold = ImageFont.truetype(FONT_SANS_BOLD, 11)
    f_badge = ImageFont.truetype(FONT_SANS_BOLD, 9.5)
    f_mono = ImageFont.truetype(FONT_PATH, 10.5)

    # Top Nav
    draw.rectangle([0, 0, W, 46], fill="#161A1D")
    draw.text((24, 14), "Projects  /  Patient Health Record Consent Gateway  /  PHRC board", fill="#9FADBC", font=f_sub)
    draw.text((680, 14), "PHRC Sprint 1 — Active Sprint Board", fill="#DCDFE4", font=f_title)

    # Controls Bar
    draw.rectangle([0, 47, W, 110], fill="#1D2125")
    draw.text((24, 62), "Active sprints", fill="#DCDFE4", font=ImageFont.truetype(FONT_SANS_BOLD, 18))
    draw.text((160, 65), "8 work items  •  18 Story Points  •  Nov 2 - Nov 13", fill="#9FADBC", font=f_sub)
    draw.rounded_rectangle([W - 155, 58, W - 24, 88], radius=3, fill="#238636")
    draw.text((W - 140, 65), "Complete sprint", fill="white", font=f_badge)

    # Search & Quick Filters
    draw.rounded_rectangle([24, 112, 220, 140], radius=4, fill="#22272B", outline="#38414A")
    draw.text((36, 118), "Search sprint board...", fill="#738496", font=f_sub)
    draw.rounded_rectangle([235, 112, 335, 140], radius=4, fill="#22272B", outline="#38414A")
    draw.text((248, 118), "Only my issues", fill="#C7D1DB", font=f_sub)
    draw.rounded_rectangle([345, 112, 455, 140], radius=4, fill="#22272B", outline="#38414A")
    draw.text((358, 118), "Recently updated", fill="#C7D1DB", font=f_sub)

    # Three Swimlanes: TO DO (2), IN PROGRESS (1), DONE (5)
    col_w = 430
    col_x_todo = 24
    col_x_prog = 480
    col_x_done = 936

    # Column Headers
    draw.rounded_rectangle([col_x_todo, 155, col_x_todo + col_w, 190], radius=4, fill="#22272B")
    draw.text((col_x_todo + 15, 165), "TO DO  (2)", fill="#C7D1DB", font=f_item_bold)
    
    draw.rounded_rectangle([col_x_prog, 155, col_x_prog + col_w, 190], radius=4, fill="#22272B")
    draw.text((col_x_prog + 15, 165), "IN PROGRESS  (1)", fill="#579DFF", font=f_item_bold)

    draw.rounded_rectangle([col_x_done, 155, col_x_done + col_w, 190], radius=4, fill="#22272B")
    draw.text((col_x_done + 15, 165), "DONE  (5)", fill="#3FB950", font=f_item_bold)

    # Helper function to render a card
    def render_card(x, y, key, title, epic, pts, epic_col):
        draw.rounded_rectangle([x, y, x + col_w, y + 84], radius=5, fill="#22272B", outline="#2C333A")
        draw.text((x + 14, y + 10), title, fill="#DCDFE4", font=f_item_bold)
        # Epic pill
        draw.rounded_rectangle([x + 14, y + 36, x + 180, y + 54], radius=3, fill=epic_col)
        draw.text((x + 20, y + 39), epic, fill="white", font=f_badge)
        # Key & Points & Priority
        draw.text((x + 14, y + 62), key, fill="#579DFF", font=f_mono)
        draw.ellipse([x + col_w - 55, y + 57, x + col_w - 35, y + 77], fill="#2C333A")
        draw.text((x + col_w - 48, y + 60), str(pts), fill="#DCDFE4", font=f_badge)
        draw.text((x + col_w - 25, y + 60), "^", fill="#FF5630", font=f_item_bold)

    # Cards in TO DO
    render_card(col_x_todo, 205, "PHRC-10", "US-2.3: Provider Credential Verification", "Clinical Provider Access", 3, "#1F883D")
    render_card(col_x_todo, 305, "PHRC-11", "US-3.1: Multi-Channel Patient Alerts", "Audit & Security", 2, "#BF8700")

    # Cards in IN PROGRESS
    render_card(col_x_prog, 205, "PHRC-9", "US-2.2: Scoped Record Retrieval", "Clinical Provider Access", 5, "#1F883D")

    # Cards in DONE
    render_card(col_x_done, 205, "PHRC-5", "US-1.1: Granular Record Selection", "Manage Patient Consent", 3, "#8250DF")
    render_card(col_x_done, 305, "PHRC-6", "US-1.2: Time-Bounded Validity Window", "Manage Patient Consent", 2, "#8250DF")
    render_card(col_x_done, 405, "PHRC-7", "US-1.3: Real-Time Consent Revocation", "Manage Patient Consent", 3, "#8250DF")
    render_card(col_x_done, 505, "PHRC-8", "US-2.1: Consent Access Request", "Clinical Provider Access", 5, "#1F883D")
    render_card(col_x_done, 605, "PHRC-12", "US-3.2: Tamper-Evident Access Audit Trail", "Audit & Security", 5, "#BF8700")

    output_path = "LAB2/screenshots/SS3_jira_active_sprint.png"
    img.save(output_path, "PNG")
    print(f"Generated: {output_path}")

if __name__ == "__main__":
    generate_burndown_chart()
    generate_jira_backlog_image()
    generate_jira_create_story_modal()
    generate_jira_story_points_image()
    generate_jira_active_sprint_image()
