import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

os.makedirs("LAB1/diagrams", exist_ok=True)

fig, ax = plt.subplots(figsize=(15, 11), dpi=300)
ax.set_xlim(0, 15)
ax.set_ylim(0, 11)
ax.axis('off')

# Background
fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FFFFFF')

# Diagram Header
ax.text(7.5, 10.65, "UML Use-Case Diagram — Patient Health Record Consent Management System",
        fontsize=15, fontweight='bold', ha='center', va='top', color='#1A237E', fontfamily='sans-serif')
ax.text(7.5, 10.35, "Problem Statement #13: Healthcare & Telemedicine | PES University — Dept. of CSE",
        fontsize=10.5, fontstyle='italic', ha='center', va='top', color='#3F51B5', fontfamily='sans-serif')

# System Boundary Rectangle
sys_rect = patches.FancyBboxPatch((3.1, 0.6), 8.8, 9.4, boxstyle="round,pad=0.08,rounding_size=0.15",
                                 facecolor='#FAFAFD', edgecolor='#303F9F', linewidth=1.8, zorder=1)
ax.add_patch(sys_rect)
ax.text(7.5, 9.75, "System: Patient Health Record Consent Management System",
        fontsize=11.5, fontweight='bold', ha='center', va='center', color='#1A237E', zorder=2)

# Function to draw an Actor (Stick Figure)
def draw_actor(ax, x, y, name, color='#1A237E'):
    # Head
    head = patches.Circle((x, y + 0.35), 0.16, facecolor='white', edgecolor=color, linewidth=1.8, zorder=5)
    ax.add_patch(head)
    # Body
    ax.plot([x, x], [y + 0.19, y - 0.22], color=color, linewidth=1.8, zorder=5)
    # Arms
    ax.plot([x - 0.28, x + 0.28], [y + 0.05, y + 0.05], color=color, linewidth=1.8, zorder=5)
    # Legs
    ax.plot([x, x - 0.22], [y - 0.22, y - 0.55], color=color, linewidth=1.8, zorder=5)
    ax.plot([x, x + 0.22], [y - 0.22, y - 0.55], color=color, linewidth=1.8, zorder=5)
    # Name Label
    ax.text(x, y - 0.75, name, fontsize=9.5, fontweight='bold', ha='center', va='top', color='#212121', zorder=6)

# Function to draw a Use Case (Ellipse)
def draw_use_case(ax, cx, cy, w, h, text, subtext="", facecolor='#E8EAF6', edgecolor='#3949AB'):
    ellipse = patches.Ellipse((cx, cy), w, h, facecolor=facecolor, edgecolor=edgecolor, linewidth=1.4, zorder=3)
    ax.add_patch(ellipse)
    if subtext:
        ax.text(cx, cy + 0.09, text, fontsize=8.8, fontweight='bold', ha='center', va='center', color='#1A237E', zorder=4)
        ax.text(cx, cy - 0.11, subtext, fontsize=7.5, fontstyle='italic', ha='center', va='center', color='#455A64', zorder=4)
    else:
        ax.text(cx, cy, text, fontsize=8.8, fontweight='bold', ha='center', va='center', color='#1A237E', zorder=4)
    return cx, cy

# Function to draw dependency relationship (dashed arrow with stereotype)
def draw_dependency(ax, x1, y1, x2, y2, stereotype="<<include>>", text_offset=(0, 0.16), rad=0.0):
    ax.annotate("",
                xy=(x2, y2), xycoords='data',
                xytext=(x1, y1), textcoords='data',
                arrowprops=dict(arrowstyle="->", color="#C2185B" if "extend" in stereotype else "#0288D1",
                                linestyle="dashed", linewidth=1.4,
                                connectionstyle=f"arc3,rad={rad}"),
                zorder=6)
    mid_x = (x1 + x2) / 2 + text_offset[0]
    mid_y = (y1 + y2) / 2 + text_offset[1]
    color = "#AD1457" if "extend" in stereotype else "#01579B"
    bg = "#FCE4EC" if "extend" in stereotype else "#E1F5FE"
    border = "#F48FB1" if "extend" in stereotype else "#81D4FA"
    ax.text(mid_x, mid_y, stereotype, fontsize=7.8, fontweight='bold', ha='center', va='center',
            color=color, zorder=7,
            bbox=dict(boxstyle="round,pad=0.2", facecolor=bg, edgecolor=border, alpha=0.95))

# Draw Actors
draw_actor(ax, 1.5, 6.2, "Patient\n(Primary Actor)", color='#0D47A1')
draw_actor(ax, 13.5, 7.8, "Clinic Administrator\n(Institutional Actor)", color='#2E7D32')
draw_actor(ax, 13.5, 4.4, "Doctor / Care Provider\n(Healthcare Actor)", color='#D81B60')
draw_actor(ax, 13.5, 1.8, "Audit & Logging System\n(Secondary System)", color='#455A64')

# Draw Use Cases inside boundary
# UC 1: Core Grant Consent
uc1 = draw_use_case(ax, 6.2, 8.4, 2.7, 0.85, "Grant Granular Time-Bound\nConsent", "UC-01 (Core Flow)")
# UC 2: Authenticate Identity
uc2 = draw_use_case(ax, 6.2, 6.8, 2.4, 0.75, "Authenticate User Identity", "MFA / Biometric", facecolor='#FFF9C4', edgecolor='#FBC02D')
# UC 3: Emergency Access Override (Extension of Grant Consent)
uc3 = draw_use_case(ax, 9.8, 8.4, 2.5, 0.8, "Set Emergency Access\nOverride", "Extension Case", facecolor='#FFEBEE', edgecolor='#E53935')
# UC 4: Revoke / Modify Active Consent
uc4 = draw_use_case(ax, 5.8, 5.3, 2.6, 0.75, "Revoke / Modify Active\nConsent", "Real-Time Revocation")
# UC 5: View Access Audit History
uc5 = draw_use_case(ax, 5.8, 3.8, 2.6, 0.75, "View Access Audit History", "Immutable Patient Logs")
# UC 6: Request Health Record Access
uc6 = draw_use_case(ax, 9.5, 5.3, 2.6, 0.8, "Request Health Record\nAccess", "Episode-Based Request")
# UC 7: Access Diagnostic Records
uc7 = draw_use_case(ax, 7.5, 2.4, 2.7, 0.85, "Access Patient Diagnostic\nRecords", "Scoped Clinical Retrieval")
# UC 8: Send Access Notification
uc8 = draw_use_case(ax, 10.3, 2.4, 2.3, 0.75, "Send Access Notification", "Multi-Channel Alert", facecolor='#F3E5F5', edgecolor='#8E24AA')
# UC 9: Verify Doctor Credentials
uc9 = draw_use_case(ax, 9.5, 6.8, 2.6, 0.75, "Verify Doctor Credentials", "Medical License Registry")
# UC 10: Register Clinic Profile
uc10 = draw_use_case(ax, 7.8, 9.4, 2.5, 0.7, "Register & Manage Clinic", "Institutional Onboarding")

# Actor Association Lines (Solid lines)
# Patient associations
ax.plot([1.8, 4.85], [6.2, 8.4], color='#1565C0', linewidth=1.3, zorder=2) # to UC1
ax.plot([1.8, 4.5], [6.2, 5.3], color='#1565C0', linewidth=1.3, zorder=2)  # to UC4
ax.plot([1.8, 4.5], [6.2, 3.8], color='#1565C0', linewidth=1.3, zorder=2)  # to UC5

# Clinic Admin associations
ax.plot([13.1, 9.0], [7.8, 9.4], color='#2E7D32', linewidth=1.3, zorder=2) # to UC10
ax.plot([13.1, 10.8], [7.8, 6.8], color='#2E7D32', linewidth=1.3, zorder=2) # to UC9

# Doctor associations
ax.plot([13.1, 10.8], [4.4, 5.3], color='#D81B60', linewidth=1.3, zorder=2) # to UC6
ax.plot([13.1, 8.85], [4.4, 2.4], color='#D81B60', linewidth=1.3, zorder=2) # to UC7

# Audit System association
ax.plot([13.1, 7.1], [1.8, 3.8], color='#455A64', linewidth=1.3, linestyle=':', zorder=2) # to UC5
ax.plot([13.1, 8.85], [1.8, 2.4], color='#455A64', linewidth=1.3, linestyle=':', zorder=2) # to UC7

# <<include>> Relationships
# 1. Grant Consent includes Authenticate Identity
draw_dependency(ax, 6.2, 7.97, 6.2, 7.18, stereotype="<<include>>", text_offset=(0.55, 0))
# 2. Revoke Consent includes Authenticate Identity
draw_dependency(ax, 5.8, 5.68, 6.1, 6.42, stereotype="<<include>>", text_offset=(0.2, 0.38), rad=-0.15)
# 3. Access Diagnostic Records includes Authenticate Identity
draw_dependency(ax, 7.5, 2.83, 6.4, 6.42, stereotype="<<include>>", text_offset=(-0.6, -0.6), rad=0.25)

# <<extend>> Relationships
# 1. Emergency Access Override extends Grant Consent
draw_dependency(ax, 8.55, 8.4, 7.55, 8.4, stereotype="<<extend>>", text_offset=(0, 0.22))
# 2. Send Access Notification extends Access Diagnostic Records
draw_dependency(ax, 9.15, 2.4, 8.85, 2.4, stereotype="<<extend>>", text_offset=(0, 0.22))

# Notation Legend
leg_box = patches.FancyBboxPatch((0.4, 0.6), 2.5, 2.4, boxstyle="round,pad=0.03",
                                facecolor='#FFFFFF', edgecolor='#B0BEC5', linewidth=1, zorder=10)
ax.add_patch(leg_box)
ax.text(0.5, 2.8, "UML Legend:", fontsize=8.5, fontweight='bold', color='#37474F', zorder=11)

# Use case item
ax.add_patch(patches.Ellipse((0.75, 2.45), 0.5, 0.22, facecolor='#E8EAF6', edgecolor='#3949AB', linewidth=1, zorder=11))
ax.text(1.1, 2.45, "Use Case", fontsize=7.5, va='center', color='#263238', zorder=11)

# Association item
ax.plot([0.55, 0.95], [2.15, 2.15], color='#1565C0', linewidth=1.3, zorder=11)
ax.text(1.1, 2.15, "Actor Association", fontsize=7.5, va='center', color='#263238', zorder=11)

# Include item
ax.annotate("", xy=(0.95, 1.8), xytext=(0.55, 1.8),
            arrowprops=dict(arrowstyle="->", color="#0288D1", linestyle="dashed", linewidth=1.2), zorder=11)
ax.text(1.1, 1.8, "<<include>> Rel.", fontsize=7.5, va='center', color='#01579B', fontweight='bold', zorder=11)

# Extend item
ax.annotate("", xy=(0.95, 1.45), xytext=(0.55, 1.45),
            arrowprops=dict(arrowstyle="->", color="#C2185B", linestyle="dashed", linewidth=1.2), zorder=11)
ax.text(1.1, 1.45, "<<extend>> Rel.", fontsize=7.5, va='center', color='#AD1457', fontweight='bold', zorder=11)

# System Boundary
ax.plot([0.55, 0.95], [1.1, 1.1], color='#303F9F', linewidth=1.4, linestyle='-', zorder=11)
ax.text(1.1, 1.1, "System Boundary", fontsize=7.5, va='center', color='#263238', zorder=11)

plt.tight_layout()
png_path = "LAB1/diagrams/patient_consent_use_case_diagram.png"
pdf_path = "LAB1/diagrams/patient_consent_use_case_diagram.pdf"
plt.savefig(png_path, dpi=300, bbox_inches='tight')
plt.savefig(pdf_path, bbox_inches='tight')
plt.close()

print(f"Generated UML Use-Case Diagram: {png_path} and {pdf_path}")
