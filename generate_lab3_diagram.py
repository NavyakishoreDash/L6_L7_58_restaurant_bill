import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.path import Path

os.makedirs("LAB3/diagrams", exist_ok=True)

fig, ax = plt.subplots(figsize=(14, 10.5), dpi=300)
ax.set_xlim(0, 14)
ax.set_ylim(0, 10.5)
ax.axis('off')

# Background
fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FFFFFF')

# Title
ax.text(7, 10.15, "Coffee Kiosk System — UML Component Diagram", 
        fontsize=16, fontweight='bold', ha='center', va='top', color='#1A237E', fontfamily='sans-serif')
ax.text(7, 9.85, "Architectural Style: Layered Architecture (3-Tier)", 
        fontsize=11, fontstyle='italic', ha='center', va='top', color='#3F51B5', fontfamily='sans-serif')

# Function to draw UML 2.0 Component Icon in top-right corner of component
def draw_component_icon(ax, x, y, size=0.22):
    # Main box
    rect = patches.Rectangle((x, y), size*1.2, size*1.4, facecolor='white', edgecolor='#333333', linewidth=1, zorder=5)
    ax.add_patch(rect)
    # Protruding tabs
    tab_w = size * 0.4
    tab_h = size * 0.35
    t1 = patches.Rectangle((x - tab_w/2, y + size*0.8), tab_w, tab_h, facecolor='white', edgecolor='#333333', linewidth=1, zorder=6)
    t2 = patches.Rectangle((x - tab_w/2, y + size*0.25), tab_w, tab_h, facecolor='white', edgecolor='#333333', linewidth=1, zorder=6)
    ax.add_patch(t1)
    ax.add_patch(t2)

# Function to draw a Layer Package Container
def draw_layer_box(ax, x, y, w, h, label):
    # Tab
    tab_w = 2.4
    tab_h = 0.35
    tab = patches.Rectangle((x, y + h), tab_w, tab_h, facecolor='#E8EAF6', edgecolor='#7986CB', linewidth=1.2, zorder=1)
    ax.add_patch(tab)
    ax.text(x + tab_w/2, y + h + tab_h/2, label, fontsize=9.5, fontweight='bold', ha='center', va='center', color='#1A237E')
    # Main Layer Box
    box = patches.Rectangle((x, y), w, h, facecolor='#FAFAFC', edgecolor='#C5CAE9', linewidth=1.2, linestyle='--', zorder=1)
    ax.add_patch(box)

# Function to draw a Component Box
def draw_component(ax, cx, cy, w, h, name, subtext=""):
    x = cx - w/2
    y = cy - h/2
    box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.08", 
                                 facecolor='#FFF9C4', edgecolor='#FBC02D', linewidth=1.5, zorder=3)
    ax.add_patch(box)
    draw_component_icon(ax, x + w - 0.42, y + h - 0.42)
    ax.text(cx, cy + 0.18, "<<component>>", fontsize=8.5, fontstyle='italic', ha='center', va='center', color='#5D4037', zorder=4)
    ax.text(cx, cy - 0.05, name, fontsize=11, fontweight='bold', ha='center', va='center', color='#212121', zorder=4)
    if subtext:
        ax.text(cx, cy - 0.28, subtext, fontsize=8, ha='center', va='center', color='#616161', zorder=4)
    return x, y, w, h

# Function to draw a small Port square
def draw_port(ax, px, py, size=0.14):
    p = patches.Rectangle((px - size/2, py - size/2), size, size, facecolor='#ECEFF1', edgecolor='#455A64', linewidth=1.2, zorder=7)
    ax.add_patch(p)

# Draw Layer Boxes
draw_layer_box(ax, 0.8, 7.3, 12.4, 2.2, "Presentation Layer")
draw_layer_box(ax, 0.8, 3.6, 12.4, 3.2, "Business / Domain Layer")
draw_layer_box(ax, 0.8, 0.6, 12.4, 2.5, "Data Access Layer")

# Draw Components
# 1. User Interface Component (Presentation Layer)
draw_component(ax, 7.0, 8.35, 4.4, 1.25, "User Interface Component", "(Touchscreen Menu, Customization & Payment UI)")

# 2. Business Layer Components:
# Left: Payment Service Component
draw_component(ax, 2.8, 5.2, 3.6, 1.3, "Payment Service Component", "(Card Reader Gateway & Authorization)")
# Center: Order Manager Component
draw_component(ax, 7.0, 5.2, 3.8, 1.3, "Order Manager Component", "(Order State Machine, Sizing & Workflow)")
# Right: Receipt Printer Component
draw_component(ax, 11.2, 5.2, 3.6, 1.3, "Receipt Printer Component", "(Thermal Printer Hardware Controller)")

# 3. Data Layer Component:
draw_component(ax, 7.0, 1.8, 4.2, 1.3, "Database Component", "(Beverage Catalog, Pricing & Audit Store)")

# ==============================================================================
# Draw Interfaces (Lollipop = Ball = Provided; Socket = Semicircle = Required)
# ==============================================================================

# Interface 1: Order Interface (Between User Interface Component and Order Manager)
# Port on UI bottom: (7.0, 7.725)
# Port on Order Manager top: (7.0, 5.85)
draw_port(ax, 7.0, 7.725)
draw_port(ax, 7.0, 5.85)

# Assembly Connector at y = 6.8
# Solid line from Order Manager up to Ball
ax.plot([7.0, 7.0], [5.85, 6.72], color='#1565C0', linewidth=1.5, zorder=5)
# Ball (Provided by Order Manager)
ball1 = patches.Circle((7.0, 6.72), 0.12, facecolor='white', edgecolor='#1565C0', linewidth=1.8, zorder=8)
ax.add_patch(ball1)
# Solid line from UI down to Socket
ax.plot([7.0, 7.0], [7.725, 6.94], color='#1565C0', linewidth=1.5, zorder=5)
# Socket (Required by UI) - semicircle facing downwards
arc1 = patches.Arc((7.0, 6.72), 0.38, 0.38, angle=0, theta1=20, theta2=160, color='#1565C0', linewidth=2.0, zorder=7)
ax.add_patch(arc1)

# Label for Interface 1
ax.text(7.2, 6.75, "IOrderService (Order Interface)\n[In-Memory API: selectDrink(), pay(), print()]", 
        fontsize=8.5, fontweight='bold', color='#0D47A1', va='center', ha='left', zorder=9,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="#E3F2FD", edgecolor="#90CAF9", alpha=0.9))

# ------------------------------------------------------------------------------
# Interface 2: Payment Interface (Between Order Manager and Payment Service)
# Port on Order Manager left: (5.1, 5.2)
# Port on Payment Service right: (4.6, 5.2)
draw_port(ax, 5.1, 5.2)
draw_port(ax, 4.6, 5.2)

# Assembly Connector at x = 4.85
# Line from Payment Service right to Ball
ax.plot([4.6, 4.8], [5.2, 5.2], color='#2E7D32', linewidth=1.5, zorder=5)
ball2 = patches.Circle((4.8, 5.2), 0.12, facecolor='white', edgecolor='#2E7D32', linewidth=1.8, zorder=8)
ax.add_patch(ball2)
# Line from Order Manager left to Socket
ax.plot([5.1, 4.96], [5.2, 5.2], color='#2E7D32', linewidth=1.5, zorder=5)
arc2 = patches.Arc((4.8, 5.2), 0.38, 0.38, angle=0, theta1=110, theta2=250, color='#2E7D32', linewidth=2.0, zorder=7)
ax.add_patch(arc2)

# Label for Interface 2
ax.text(4.85, 5.55, "IPaymentService (Payment Interface)\n[PCI-DSS Auth: processCard(total)]", 
        fontsize=8.5, fontweight='bold', color='#1B5E20', va='bottom', ha='center', zorder=9,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="#E8F5E9", edgecolor="#A5D6A7", alpha=0.9))

# ------------------------------------------------------------------------------
# Interface 3: Receipt Interface (Between Order Manager and Receipt Printer)
# Port on Order Manager right: (8.9, 5.2)
# Port on Receipt Printer left: (9.4, 5.2)
draw_port(ax, 8.9, 5.2)
draw_port(ax, 9.4, 5.2)

# Assembly Connector at x = 9.15
# Line from Receipt Printer left to Ball
ax.plot([9.4, 9.2], [5.2, 5.2], color='#C2185B', linewidth=1.5, zorder=5)
ball3 = patches.Circle((9.2, 5.2), 0.12, facecolor='white', edgecolor='#C2185B', linewidth=1.8, zorder=8)
ax.add_patch(ball3)
# Line from Order Manager right to Socket
ax.plot([8.9, 9.04], [5.2, 5.2], color='#C2185B', linewidth=1.5, zorder=5)
arc3 = patches.Arc((9.2, 5.2), 0.38, 0.38, angle=0, theta1=290, theta2=70, color='#C2185B', linewidth=2.0, zorder=7)
ax.add_patch(arc3)

# Label for Interface 3
ax.text(9.15, 5.55, "IPrinterService (Receipt Interface)\n[ESC/POS Driver: printReceipt(details)]", 
        fontsize=8.5, fontweight='bold', color='#880E4F', va='bottom', ha='center', zorder=9,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="#FCE4EC", edgecolor="#F48FB1", alpha=0.9))

# ------------------------------------------------------------------------------
# Interface 4: Database Interface (Between Order Manager and Database Component)
# Port on Order Manager bottom: (7.0, 4.55)
# Port on Database Component top: (7.0, 2.45)
draw_port(ax, 7.0, 4.55)
draw_port(ax, 7.0, 2.45)

# Assembly Connector at y = 3.5
# Solid line from Database Component up to Ball
ax.plot([7.0, 7.0], [2.45, 3.42], color='#E65100', linewidth=1.5, zorder=5)
ball4 = patches.Circle((7.0, 3.42), 0.12, facecolor='white', edgecolor='#E65100', linewidth=1.8, zorder=8)
ax.add_patch(ball4)
# Solid line from Order Manager down to Socket
ax.plot([7.0, 7.0], [4.55, 3.64], color='#E65100', linewidth=1.5, zorder=5)
arc4 = patches.Arc((7.0, 3.42), 0.38, 0.38, angle=0, theta1=200, theta2=340, color='#E65100', linewidth=2.0, zorder=7)
ax.add_patch(arc4)

# Label for Interface 4
ax.text(7.2, 3.45, "IDatabaseService (Database Interface)\n[DAO / SQL: getMenu(), saveOrder()]", 
        fontsize=8.5, fontweight='bold', color='#BF360C', va='center', ha='left', zorder=9,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="#FFF3E0", edgecolor="#FFCC80", alpha=0.9))

# ------------------------------------------------------------------------------
# Interface 5: Direct Menu Query Interface from UI to Database (Read-only Cache)
# Curved line from UI left to Database left
port_ui_left = (4.8, 8.0)
port_db_left = (4.9, 1.8)
draw_port(ax, port_ui_left[0], port_ui_left[1])
draw_port(ax, port_db_left[0], port_db_left[1])

# Curved route along the left side
ax.annotate("",
            xy=port_db_left, xycoords='data',
            xytext=port_ui_left, textcoords='data',
            arrowprops=dict(arrowstyle="->", color="#78909C",
                            connectionstyle="arc3,rad=-0.35",
                            linestyle="dashed", linewidth=1.4),
            zorder=4)

ax.text(1.2, 6.8, "<<use>>\nIMenuCatalog\n(Read-Only Menu & Pricing)", 
        fontsize=8, fontstyle='italic', ha='left', va='center', color='#37474F', zorder=9,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="#ECEFF1", edgecolor="#CFD8DC", alpha=0.85))

# ==============================================================================
# Legend / Notation Box in Bottom Left
# ==============================================================================
leg_x, leg_y, leg_w, leg_h = 0.9, 0.75, 4.2, 0.95
leg_box = patches.FancyBboxPatch((leg_x, leg_y), leg_w, leg_h, boxstyle="round,pad=0.03", 
                                facecolor='#FFFFFF', edgecolor='#B0BEC5', linewidth=1, zorder=10)
ax.add_patch(leg_box)
ax.text(leg_x + 0.15, leg_y + leg_h - 0.18, "UML 2.0 Notation Legend:", fontsize=8.5, fontweight='bold', color='#37474F', zorder=11)

# Ball icon
ax.plot([leg_x + 0.25, leg_x + 0.5], [leg_y + 0.5, leg_y + 0.5], color='#333333', linewidth=1.2, zorder=11)
ax.add_patch(patches.Circle((leg_x + 0.5, leg_y + 0.5), 0.08, facecolor='white', edgecolor='#333333', linewidth=1.4, zorder=12))
ax.text(leg_x + 0.7, leg_y + 0.5, "Provided Interface (Ball)", fontsize=7.5, va='center', color='#263238', zorder=11)

# Socket icon
ax.plot([leg_x + 2.35, leg_x + 2.6], [leg_y + 0.5, leg_y + 0.5], color='#333333', linewidth=1.2, zorder=11)
ax.add_patch(patches.Arc((leg_x + 2.6, leg_y + 0.5), 0.25, 0.25, angle=0, theta1=270, theta2=90, color='#333333', linewidth=1.6, zorder=12))
ax.text(leg_x + 2.8, leg_y + 0.5, "Required Interface (Socket)", fontsize=7.5, va='center', color='#263238', zorder=11)

# Port square
ax.add_patch(patches.Rectangle((leg_x + 0.38, leg_y + 0.18), 0.12, 0.12, facecolor='#ECEFF1', edgecolor='#455A64', linewidth=1, zorder=12))
ax.text(leg_x + 0.7, leg_y + 0.24, "Interaction Port", fontsize=7.5, va='center', color='#263238', zorder=11)

# Assembly icon
ax.plot([leg_x + 2.15, leg_x + 2.6], [leg_y + 0.24, leg_y + 0.24], color='#1565C0', linewidth=1.2, zorder=11)
ax.add_patch(patches.Circle((leg_x + 2.45, leg_y + 0.24), 0.07, facecolor='white', edgecolor='#1565C0', linewidth=1.4, zorder=12))
ax.add_patch(patches.Arc((leg_x + 2.45, leg_y + 0.24), 0.22, 0.22, angle=0, theta1=270, theta2=90, color='#1565C0', linewidth=1.6, zorder=12))
ax.text(leg_x + 2.8, leg_y + 0.24, "Assembly Connector", fontsize=7.5, va='center', color='#263238', zorder=11)

plt.tight_layout()
png_path = "LAB3/diagrams/coffee_kiosk_component_diagram.png"
pdf_path = "LAB3/diagrams/coffee_kiosk_component_diagram.pdf"
plt.savefig(png_path, dpi=300, bbox_inches='tight')
plt.savefig(pdf_path, bbox_inches='tight')
plt.close()

print(f"Generated diagram: {png_path} and {pdf_path}")
