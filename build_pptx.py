from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import copy

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── COLORS ──────────────────────────────
NAVY   = RGBColor(0x0D, 0x1B, 0x2A)
NAVY2  = RGBColor(0x15, 0x23, 0x36)
TEAL   = RGBColor(0x00, 0xC9, 0xA7)
TEAL2  = RGBColor(0x00, 0xA9, 0x8F)
BLUE   = RGBColor(0x1B, 0x6C, 0xA8)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GRAY   = RGBColor(0x64, 0x74, 0x8B)
LIGHT  = RGBColor(0xE2, 0xEA, 0xF4)
OFF    = RGBColor(0xF4, 0xF7, 0xFB)
RED    = RGBColor(0xEF, 0x44, 0x44)
ORANGE = RGBColor(0xF5, 0x9E, 0x0B)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
PINK   = RGBColor(0xEC, 0x48, 0x99)
GREEN  = RGBColor(0x10, 0xB9, 0x81)
DARK   = RGBColor(0x1F, 0x29, 0x37)

W = prs.slide_width
H = prs.slide_height

blank = prs.slide_layouts[6]   # blank layout


# ── HELPER FUNCTIONS ─────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None, line_w=Pt(0)):
    shape = slide.shapes.add_shape(1, l, t, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background() if line is None else None
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
             italic=False, wrap=True):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def add_ml_text(slide, lines, l, t, w, h,
                size=14, color=WHITE, align=PP_ALIGN.LEFT,
                bold_first=False, spacing=Pt(4)):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_before = spacing
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold_first and (line == lines[0])
    return txb

def col(r,g,b): return RGBColor(r,g,b)

M = Inches(0.18)   # margin unit

# ══════════════════════════════════════
# SLIDE 1 — TITLE
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)

# Full background navy
add_rect(sl, 0, 0, W, H, fill=NAVY)

# Right gradient panel (teal→blue)
add_rect(sl, Inches(7.3), 0, Inches(6.03), H, fill=BLUE)
add_rect(sl, Inches(8.5), 0, Inches(4.83), H, fill=col(0x00,0x8C,0x76))

# Decorative circles (right side)
for r,a in [(2.2,0.06),(1.6,0.10),(0.9,0.16)]:
    s = slide_circle = sl.shapes.add_shape(9,  # oval
        Inches(10.16 - r), Inches(3.75 - r),
        Inches(r*2), Inches(r*2))
    s.fill.background()
    s.line.color.rgb = col(0xFF,0xFF,0xFF)
    s.line.width = Pt(1)

# Left content
# Badge
add_rect(sl, Inches(0.6), Inches(0.55), Inches(2.6), Inches(0.34),
         fill=col(0x00,0x40,0x35), line=TEAL, line_w=Pt(1))
add_text(sl, "● CONFIDENTIAL PROPOSAL · 2026",
         Inches(0.62), Inches(0.55), Inches(2.6), Inches(0.34),
         size=8, bold=True, color=TEAL)

# Main heading
add_text(sl, "Scaling Care at Fakeeh",
         Inches(0.55), Inches(1.1), Inches(6.5), Inches(0.75),
         size=36, bold=True, color=WHITE)
add_text(sl, "with AI Agents",
         Inches(0.55), Inches(1.8), Inches(6.5), Inches(0.65),
         size=36, bold=True, color=TEAL)

# Sub-text
add_text(sl,
    "How Amigo's patient-facing clinical AI platform can multiply\n"
    "Fakeeh Care Group's clinical capacity — without scaling headcount.",
    Inches(0.55), Inches(2.65), Inches(6.2), Inches(1.1),
    size=13, color=col(0xAA,0xBB,0xCC))

# Chips
for i, chip in enumerate(["🏥  Fakeeh Care Group", "🤖  Amigo AI", "🇸🇦  Vision 2030 Aligned"]):
    add_rect(sl, Inches(0.55 + i*2.15), Inches(4.0), Inches(2.0), Inches(0.36),
             fill=col(0x1E,0x30,0x45), line=col(0x33,0x55,0x77), line_w=Pt(1))
    add_text(sl, chip, Inches(0.6 + i*2.15), Inches(4.0), Inches(2.0), Inches(0.36),
             size=9, color=col(0xCC,0xDD,0xEE))

# Amigo logo placeholder text (right panel)
add_text(sl, "🤝", Inches(9.5), Inches(3.1), Inches(1.0), Inches(0.8),
         size=40, align=PP_ALIGN.CENTER)
add_text(sl, "amigo.ai", Inches(9.2), Inches(3.95), Inches(1.5), Inches(0.4),
         size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl, "×  Fakeeh Care Group", Inches(8.8), Inches(4.35), Inches(2.4), Inches(0.35),
         size=11, color=col(0xCC,0xFF,0xEE), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════
# SLIDE 2 — ABOUT AMIGO
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=OFF)

# Header bar
add_rect(sl, 0, 0, W, Inches(1.05), fill=NAVY)
add_rect(sl, Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38), fill=TEAL)
add_text(sl, "01", Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38),
         size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(sl, "About Amigo AI",
         Inches(0.88), Inches(0.28), Inches(5), Inches(0.5),
         size=22, bold=True, color=WHITE)

# Tagline box (left top)
add_rect(sl, Inches(0.4), Inches(1.2), Inches(5.9), Inches(0.95), fill=NAVY)
add_text(sl, '"Turn Clinical Expertise into Exponential Capacity."',
         Inches(0.5), Inches(1.25), Inches(5.7), Inches(0.9),
         size=15, bold=True, color=WHITE)

# Stats row
stats = [("3M+","Autonomous patient encounters"),
         ("0","Safety incidents to date"),
         ("6 wk","Concept → production"),
         ("100+","Languages incl. Arabic")]
for i,(num,lbl) in enumerate(stats):
    x = Inches(0.4 + i*1.5)
    add_rect(sl, x, Inches(2.28), Inches(1.42), Inches(0.95),
             fill=WHITE, line=LIGHT, line_w=Pt(1))
    add_text(sl, num, x, Inches(2.32), Inches(1.42), Inches(0.45),
             size=19, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(sl, lbl, x, Inches(2.72), Inches(1.42), Inches(0.45),
             size=7.5, color=GRAY, align=PP_ALIGN.CENTER)

# Description box
add_rect(sl, Inches(0.4), Inches(3.38), Inches(5.9), Inches(1.1),
         fill=WHITE, line=LIGHT, line_w=Pt(1))
add_text(sl,
    "Amigo is the trusted platform for building, training & deploying patient-facing clinical AI agents. "
    "Founded 2023 · Headquartered NYC · Backed by Madrona & Optum Ventures ($17M raised). "
    "Partners define their AI strategy; Amigo builds & launches custom agents safely in 6 weeks — no engineering required.",
    Inches(0.55), Inches(3.43), Inches(5.6), Inches(1.05),
    size=10, color=DARK)

# Investors
add_rect(sl, Inches(0.4), Inches(4.6), Inches(5.9), Inches(0.42),
         fill=WHITE, line=LIGHT, line_w=Pt(1))
add_text(sl, "Backed by:  Madrona · Optum Ventures   |   Total Raised: $17M",
         Inches(0.55), Inches(4.62), Inches(5.7), Inches(0.38),
         size=10, color=GRAY)

# Right column — features
features = [
    ("🎓", "Digital Residency™ Training",
     "Millions of simulated patient scenarios — adversarial & edge cases — before any real patient. 100% safety pass rate required."),
    ("⚡", "Amigo Actions (Real-World)",
     "Agents order labs, write to EHR, schedule appointments, message care team — not just advisory chatbots."),
    ("🧠", "Context Graph & 4-Tier Memory",
     "Living, evolving understanding of each patient shared across multi-agent care team in real time."),
    ("🔒", "Enterprise Compliance",
     "HIPAA · SOC 2 Type II · GDPR · Epic · Oracle Health · Athenahealth integrations."),
]
for i,(icon,title,desc) in enumerate(features):
    y = Inches(1.2 + i*1.5)
    add_rect(sl, Inches(6.55), y, Inches(6.4), Inches(1.35),
             fill=WHITE, line=TEAL, line_w=Pt(2))
    add_text(sl, icon, Inches(6.62), y+Inches(0.05), Inches(0.5), Inches(0.5), size=18)
    add_text(sl, title, Inches(7.15), y+Inches(0.08), Inches(5.7), Inches(0.38),
             size=11, bold=True, color=NAVY)
    add_text(sl, desc, Inches(7.15), y+Inches(0.48), Inches(5.7), Inches(0.8),
             size=9, color=GRAY)

# ══════════════════════════════════════
# SLIDE 3 — FAKEEH AT A GLANCE
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=NAVY)

add_rect(sl, 0, 0, W, Inches(1.05), fill=NAVY2)
add_rect(sl, Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38), fill=TEAL)
add_text(sl, "02", Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38),
         size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(sl, "Fakeeh Care Group at a Glance",
         Inches(0.88), Inches(0.28), Inches(8), Inches(0.5),
         size=22, bold=True, color=WHITE)

# KPI cards
kpis = [
    ("SAR 3.1B","FY2025 Revenue","↑ +11% YoY"),
    ("1.89M","Annual Patients","↑ +8% YoY"),
    ("4 Hospitals","KSA Network (→7 by 2028)","835 beds & growing"),
    ("900+","Doctors in Network","50+ specialties"),
    ("47 Years","Of Clinical Excellence","Founded 1978, Jeddah"),
    ("JCI x6","Consecutive Accreditation","+ ANCC Magnet Nursing"),
]
for i,(num,lbl,sub) in enumerate(kpis):
    x = Inches(0.35 + (i%3)*4.35)
    y = Inches(1.15 + (i//3)*1.7)
    add_rect(sl, x, y, Inches(4.1), Inches(1.52),
             fill=col(0x14,0x26,0x38), line=col(0x1E,0x40,0x58), line_w=Pt(1))
    add_rect(sl, x, y, Inches(4.1), Inches(0.055), fill=TEAL)
    add_text(sl, num, x+Inches(0.18), y+Inches(0.14), Inches(3.7), Inches(0.55),
             size=22, bold=True, color=WHITE)
    add_text(sl, lbl, x+Inches(0.18), y+Inches(0.66), Inches(3.7), Inches(0.38),
             size=10, color=col(0x88,0xAA,0xCC))
    add_text(sl, sub, x+Inches(0.18), y+Inches(1.0), Inches(3.7), Inches(0.38),
             size=9, bold=True, color=TEAL)

# Two info boxes bottom
for col_i, (heading, bullets) in enumerate([
    ("Established Strength", [
        "→  JCI accredited — 6th consecutive reaccreditation",
        "→  CAP Lab: zero deficiencies | ANCC Magnet Nursing",
        "→  #1 private hospital in KSA — Newsweek 2022–2025",
        "→  First private-sector robotic surgery in Saudi Arabia",
        "→  Largest NICU in Western Province of KSA",
    ]),
    ("Ambitious Growth Agenda", [
        "→  Target: 7 hospitals, 1,675 beds, 9 centers by 2028",
        "→  SAR 938M Alinma Bank credit facility secured",
        "→  HEAL Hospital (SAR 461M Neuroscience) underway",
        "→  Acquired Saudia Airlines Medical Services",
        "→  UAE expansion: FUH + Medicentres acquisition",
    ]),
]):
    x = Inches(0.35 + col_i*6.7)
    add_rect(sl, x, Inches(4.62), Inches(6.4), Inches(2.6),
             fill=col(0x0A,0x16,0x24), line=col(0x1E,0x40,0x58), line_w=Pt(1))
    add_text(sl, heading, x+Inches(0.2), Inches(4.72), Inches(6.0), Inches(0.38),
             size=11, bold=True, color=TEAL)
    add_ml_text(sl, bullets, x+Inches(0.15), Inches(5.15), Inches(6.1), Inches(2.0),
                size=9.5, color=col(0xBB,0xCC,0xDD), spacing=Pt(2))

# ══════════════════════════════════════
# SLIDE 4 — PAIN POINTS
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=OFF)

add_rect(sl, 0, 0, W, Inches(1.05), fill=NAVY)
add_rect(sl, Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38), fill=RED)
add_text(sl, "03", Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38),
         size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl, "Identified Challenges at Scale",
         Inches(0.88), Inches(0.28), Inches(8), Inches(0.5),
         size=22, bold=True, color=WHITE)

pain_data = [
    (RED,    "📈", "Explosive Patient Volume Growth",
     "1.89M annual patients growing 8%+ YoY. Clinical staff cannot\nscale at the same pace — every new hospital multiplies the\npatient engagement burden.", "Capacity Constraint"),
    (ORANGE, "🔄", "Post-Visit Follow-Up Gap",
     "1.89M discharges/consultations annually. Systematically\nfollowing up every patient for medication adherence, recovery,\nand chronic care is operationally impossible.", "Care Continuity"),
    (PURPLE, "🌍", "Multi-Language Patient Population",
     "Diverse expatriate & national population requiring care in\nArabic dialects, English, Urdu, Tagalog, and more — across\nall Fakeeh facilities.", "Accessibility"),
    (BLUE,   "🏗️", "Rapid Multi-Site Expansion",
     "Scaling 4→7 hospitals and 9 medical centers by 2028 means\nreplicating patient engagement operations across Jeddah,\nRiyadh, Madinah, Makkah, NEOM without a scalable model.", "Operational Scale"),
    (PINK,   "💊", "Chronic Disease Management Burden",
     "With 50+ specialties incl. cardiology, oncology, nephrology,\nmanaging ongoing patient engagement consumes\ndisproportionate clinical bandwidth.", "Clinical Efficiency"),
    (NAVY2,  "🏠", "Home Healthcare Coordination",
     "Fakeeh Home Health Care operates across 6 cities. Remote\npatient monitoring & check-ins rely on manual outreach —\ncreating inconsistent patient experience.", "Remote Care"),
]

for i,(accent,icon,title,desc,tag) in enumerate(pain_data):
    col_i = i % 3
    row_i = i // 3
    x = Inches(0.3 + col_i*4.35)
    y = Inches(1.15 + row_i*3.05)
    add_rect(sl, x, y, Inches(4.1), Inches(2.85),
             fill=WHITE, line=LIGHT, line_w=Pt(1))
    add_rect(sl, x, y+Inches(2.79), Inches(4.1), Inches(0.06), fill=accent)
    add_text(sl, icon, x+Inches(0.15), y+Inches(0.12), Inches(0.55), Inches(0.5), size=20)
    add_text(sl, title, x+Inches(0.15), y+Inches(0.65), Inches(3.85), Inches(0.42),
             size=11, bold=True, color=NAVY)
    add_text(sl, desc, x+Inches(0.15), y+Inches(1.1), Inches(3.85), Inches(1.45),
             size=9, color=GRAY)
    add_rect(sl, x+Inches(0.15), y+Inches(2.5), Inches(1.6), Inches(0.26),
             fill=OFF, line=LIGHT, line_w=Pt(1))
    add_text(sl, tag, x+Inches(0.15), y+Inches(2.5), Inches(1.6), Inches(0.26),
             size=7.5, color=GRAY, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════
# SLIDE 5 — SOLUTION MAP
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=NAVY)

add_rect(sl, 0, 0, W, Inches(1.05), fill=NAVY2)
add_rect(sl, Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38), fill=TEAL)
add_text(sl, "04", Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38),
         size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(sl, "Amigo Solution Mapping",
         Inches(0.88), Inches(0.28), Inches(8), Inches(0.5),
         size=22, bold=True, color=WHITE)

# Table header
add_rect(sl, Inches(0.35), Inches(1.12), Inches(12.63), Inches(0.42),
         fill=col(0x00,0x40,0x35))
for xi,(lbl,w) in enumerate([("Fakeeh Challenge",4.5),("Amigo Solution",6.0),("Impact",2.1)]):
    add_text(sl, lbl.upper(),
             Inches(0.5 + sum([4.5,6.0,2.1][:xi])), Inches(1.14),
             Inches(w), Inches(0.38),
             size=8.5, bold=True, color=TEAL)

rows = [
    (RED,    "Explosive patient volume growth",
     "24/7 AI Agents handle routine inquiries, triage & scheduling — freeing clinicians for high-complexity cases",
     "80–90% deflection"),
    (ORANGE, "Post-visit follow-up gap",
     "Outbound AI Engagement — automated post-visit check-ins, medication reminders, recovery tracking at population scale",
     "100% coverage"),
    (PURPLE, "Multi-language patient population",
     "100+ Language Support including Arabic dialects — natively deployed, no translation layer required",
     "Full inclusivity"),
    (BLUE,   "Multi-site rapid expansion",
     "Scalable Agent Architecture — one platform centrally managed, deployed across all Fakeeh facilities and brands",
     "Instant replication"),
    (PINK,   "Chronic disease management burden",
     "Chronic Care AI — ongoing engagement, care plan adherence, lab result debriefs, medication guidance",
     "Significant reduction"),
    (NAVY2,  "Home healthcare coordination",
     "Home Health AI Agents — proactive remote check-ins, vitals follow-up, escalation routing across 6 cities",
     "Consistent quality"),
]

for i,(dot_c,pain,sol,impact) in enumerate(rows):
    y = Inches(1.54 + i*0.97)
    bg = col(0xFF,0xFF,0xFF) if i%2==0 else col(0xF8,0xFA,0xFC)
    add_rect(sl, Inches(0.35), y, Inches(12.63), Inches(0.94), fill=col(0x10,0x1E,0x2C))
    # dot
    circ = sl.shapes.add_shape(9, Inches(0.5), y+Inches(0.35), Inches(0.18), Inches(0.18))
    circ.fill.solid(); circ.fill.fore_color.rgb = dot_c
    circ.line.fill.background()
    add_text(sl, pain, Inches(0.75), y+Inches(0.05), Inches(4.1), Inches(0.85),
             size=9.5, color=WHITE)
    add_text(sl, sol, Inches(4.85), y+Inches(0.05), Inches(5.85), Inches(0.85),
             size=9, color=col(0xBB,0xCC,0xDD))
    add_rect(sl, Inches(10.75), y+Inches(0.28), Inches(2.0), Inches(0.32),
             fill=col(0x00,0x40,0x35), line=TEAL, line_w=Pt(1))
    add_text(sl, impact, Inches(10.75), y+Inches(0.28), Inches(2.0), Inches(0.32),
             size=8, bold=True, color=TEAL, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════
# SLIDE 6 — USE CASES DEEP DIVE
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=OFF)

add_rect(sl, 0, 0, W, Inches(1.05), fill=NAVY)
add_rect(sl, Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38), fill=TEAL)
add_text(sl, "05", Inches(0.4), Inches(0.33), Inches(0.38), Inches(0.38),
         size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(sl, "Amigo Use Cases for Fakeeh",
         Inches(0.88), Inches(0.28), Inches(8), Inches(0.5),
         size=22, bold=True, color=WHITE)

use_cases = [
    ("🚦","AI Triage & Intake","Across 120+ outpatient clinics",
     "Intelligent AI agents screen patients before appointments — collecting symptoms, history & urgency. Routes to the right specialty, reducing no-shows and misrouting across all facilities.",
     "Reduces admin time per patient by 40%+"),
    ("📋","Chronic Care Management","Cardiology · Oncology · Nephrology",
     "Persistent agents manage ongoing patient relationships. Proactively reach out for medication adherence, lifestyle monitoring, and early warning escalation — at full population scale.",
     "80–90% reduction in routine chronic care calls"),
    ("🔬","Lab Result Debriefs","Integrated with Fakeeh's YASASII HIS",
     "AI agents explain lab & radiology results in plain language and the patient's preferred language — reducing the thousands of 'what does this mean?' calls that consume clinician time daily.",
     "Frees 2–4 hours/day per clinical department"),
    ("🏠","Home Health AI Support","All 6 home health cities covered",
     "Automated check-ins for Fakeeh Home Health Care patients — vitals prompts, medication reminders, wound care guidance, and escalation to nursing staff when deterioration is detected.",
     "Consistent care quality across all cities"),
    ("🎗️","Oncology Survivorship","DSFH Madinah Cancer Centre + Jeddah",
     "Dedicated cancer survivorship agents — side effect monitoring, psychological support check-ins, appointment reminders, and connecting patients to palliative or social services.",
     "Improved survivorship outcomes & CSAT"),
    ("🧬","NEOM AI-Native Care","Personalized · Predictive · Genomics-Ready",
     "For Fakeeh's NEOM hospital — Amigo's advanced personalized agents align perfectly with NEOM's vision for AI-driven, genomics-integrated, preventive healthcare.",
     "First AI-native hospital care model in KSA"),
]

for i,(icon,title,sub,desc,outcome) in enumerate(use_cases):
    ci = i % 3
    ri = i // 3
    x = Inches(0.3 + ci*4.35)
    y = Inches(1.15 + ri*3.05)
    add_rect(sl, x, y, Inches(4.1), Inches(2.85),
             fill=WHITE, line=LIGHT, line_w=Pt(1))
    add_rect(sl, x, y, Inches(4.1), Inches(0.06), fill=NAVY)
    # icon circle
    circ2 = sl.shapes.add_shape(9, x+Inches(0.15), y+Inches(0.14), Inches(0.48), Inches(0.48))
    circ2.fill.solid(); circ2.fill.fore_color.rgb = NAVY
    circ2.line.fill.background()
    add_text(sl, icon, x+Inches(0.15), y+Inches(0.12), Inches(0.48), Inches(0.48), size=16)
    add_text(sl, title, x+Inches(0.72), y+Inches(0.14), Inches(3.25), Inches(0.3),
             size=10.5, bold=True, color=NAVY)
    add_text(sl, sub, x+Inches(0.72), y+Inches(0.44), Inches(3.25), Inches(0.26),
             size=8, color=GRAY)
    add_text(sl, desc, x+Inches(0.15), y+Inches(0.82), Inches(3.8), Inches(1.45),
             size=8.5, color=DARK)
    add_rect(sl, x+Inches(0.15), y+Inches(2.42), Inches(3.8), Inches(0.32),
             fill=col(0xE6,0xF9,0xF5), line=TEAL, line_w=Pt(1))
    add_text(sl, "✓  " + outcome, x+Inches(0.2), y+Inches(2.43), Inches(3.75), Inches(0.3),
             size=7.5, bold=True, color=TEAL2)

# ══════════════════════════════════════
# SLIDE 7 — WHY AMIGO
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=NAVY)

# Left panel gradient
add_rect(sl, 0, 0, Inches(5.5), H, fill=col(0x00,0x85,0x70))
add_rect(sl, 0, 0, Inches(4.2), H, fill=col(0x00,0xA0,0x88))
add_rect(sl, 0, 0, Inches(3.0), H, fill=TEAL2)

# Decorative circles on left
for r in [3.0, 2.2, 1.4]:
    c2 = sl.shapes.add_shape(9, Inches(-r+1), Inches(3.75-r), Inches(r*2), Inches(r*2))
    c2.fill.background(); c2.line.color.rgb = col(0xFF,0xFF,0xFF); c2.line.width = Pt(0.75)

add_text(sl, "Why Amigo?", Inches(0.3), Inches(0.8), Inches(5.0), Inches(0.65),
         size=26, bold=True, color=WHITE)
add_text(sl, "Built for exactly this scale.", Inches(0.3), Inches(1.45), Inches(5.0), Inches(0.42),
         size=14, color=col(0xDD,0xFF,0xF5))

# Big stat
add_text(sl, "3M+", Inches(0.3), Inches(2.1), Inches(5.0), Inches(1.1),
         size=56, bold=True, color=WHITE)
add_text(sl, "Autonomous patient encounters completed worldwide",
         Inches(0.3), Inches(3.15), Inches(4.8), Inches(0.55),
         size=11, color=col(0xCC,0xFF,0xEE))

# Safety badge
add_rect(sl, Inches(0.3), Inches(3.9), Inches(4.8), Inches(0.44),
         fill=col(0xFF,0xFF,0xFF) if False else col(0x00,0x55,0x48),
         line=col(0xFF,0xFF,0xFF), line_w=Pt(1))
add_text(sl, "🛡️  Zero safety incidents · 99.9%+ safety metrics",
         Inches(0.38), Inches(3.91), Inches(4.7), Inches(0.42),
         size=10, bold=True, color=WHITE)

add_rect(sl, Inches(0.3), Inches(4.45), Inches(4.8), Inches(0.44),
         fill=col(0x00,0x55,0x48), line=col(0xFF,0xFF,0xFF), line_w=Pt(1))
add_text(sl, "🇸🇦  Active Middle East presence via Heal partnership",
         Inches(0.38), Inches(4.46), Inches(4.7), Inches(0.42),
         size=10, bold=True, color=WHITE)

# Right panel header
add_rect(sl, Inches(5.7), 0, Inches(7.63), H, fill=NAVY)
add_text(sl, "What Makes Amigo Different for Fakeeh",
         Inches(5.85), Inches(0.25), Inches(7.3), Inches(0.55),
         size=15, bold=True, color=WHITE)

diffs = [
    ("1","Trained Like Doctors, Not Like Chatbots",
     "Digital Residency™ simulates millions of patient scenarios — including adversarial & edge cases — before any real patient interaction. Aligned with Fakeeh's JCI standards."),
    ("2","YASASII-Ready EHR Integration",
     "Amigo integrates natively with major EHRs and can bridge to Fakeeh's proprietary YASASII HIS — ensuring agents read & write to the same patient record clinicians use."),
    ("3","Arabic-First, Multi-Dialect",
     "Native Arabic support across dialects — Gulf, Egyptian, Levantine — plus 100+ languages covering Fakeeh's full patient demographics in KSA and UAE."),
    ("4","Real Actions, Not Just Advice",
     "Agents order labs, schedule appointments, write clinical notes — directly supporting Fakeeh's digital transformation and Vision 2030 paperless hospital targets."),
]
for i,(num,title,desc) in enumerate(diffs):
    y = Inches(0.9 + i*1.3)
    add_rect(sl, Inches(5.85), y, Inches(7.1), Inches(1.18),
             fill=col(0x14,0x26,0x38), line=col(0x1E,0x40,0x58), line_w=Pt(1))
    circ3 = sl.shapes.add_shape(9, Inches(5.95), y+Inches(0.3), Inches(0.5), Inches(0.5))
    circ3.fill.solid(); circ3.fill.fore_color.rgb = TEAL
    circ3.line.fill.background()
    add_text(sl, num, Inches(5.95), y+Inches(0.3), Inches(0.5), Inches(0.5),
             size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(sl, title, Inches(6.55), y+Inches(0.08), Inches(6.3), Inches(0.38),
             size=10.5, bold=True, color=WHITE)
    add_text(sl, desc, Inches(6.55), y+Inches(0.48), Inches(6.3), Inches(0.65),
             size=8.5, color=col(0x88,0xAA,0xCC))

# Partner note
add_rect(sl, Inches(5.85), Inches(6.1), Inches(7.1), Inches(0.9),
         fill=col(0x00,0x30,0x28), line=TEAL, line_w=Pt(1))
add_text(sl,
    "🤝  Established Middle East Presence: Amigo has an active partnership with Heal (KSA/UAE) "
    "to expand clinical AI across the Middle East — Fakeeh joins a proven regional deployment track.",
    Inches(5.98), Inches(6.12), Inches(6.9), Inches(0.86),
    size=8.5, color=TEAL)

# ══════════════════════════════════════
# SLIDE 8 — NEXT STEPS
# ══════════════════════════════════════
sl = prs.slides.add_slide(blank)
add_rect(sl, 0, 0, W, H, fill=NAVY)

# Radial glow effects (simulated with large soft rects)
add_rect(sl, Inches(8.0), Inches(-1.5), Inches(6.0), Inches(6.0), fill=col(0x00,0x2A,0x22))
add_rect(sl, Inches(-1.5), Inches(3.5), Inches(6.0), Inches(6.0), fill=col(0x10,0x20,0x35))

add_text(sl, "THE PATH FORWARD",
         Inches(0), Inches(0.35), W, Inches(0.38),
         size=9, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
add_text(sl, "Let's Build the Future of Healthcare Together",
         Inches(0), Inches(0.7), W, Inches(0.8),
         size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl,
    "Amigo goes from concept to production in 6 weeks.\nHere's how we propose to begin the partnership with Fakeeh Care Group.",
    Inches(2.0), Inches(1.45), Inches(9.33), Inches(0.65),
    size=11, color=col(0x88,0xAA,0xCC), align=PP_ALIGN.CENTER)

steps = [
    ("1","Discovery\nWorkshop","Map Fakeeh's top 3 patient engagement bottlenecks with clinical teams"),
    ("2","Pilot\nDesign","Co-design a focused 6-week pilot — e.g. post-discharge follow-up at DSFH Jeddah"),
    ("3","Agent\nTraining","Digital Residency™ — trained on Fakeeh's patient population, Arabic dialects, protocols"),
    ("4","Safe\nLaunch","Deploy to live patients with full monitoring, safety dashboards, clinical oversight"),
    ("5","Scale Across\nNetwork","Expand across all hospitals, home health, UAE → 7-hospital 2028 vision"),
]

step_x_start = Inches(0.55)
step_w = Inches(2.35)
gap = Inches(0.2)

for i,(num,title,desc) in enumerate(steps):
    x = step_x_start + i*(step_w+gap)
    y = Inches(2.35)
    # Circle
    circ4 = sl.shapes.add_shape(9, x+Inches(0.78), y, Inches(0.78), Inches(0.78))
    circ4.fill.solid(); circ4.fill.fore_color.rgb = col(0x00,0x30,0x28)
    circ4.line.color.rgb = TEAL; circ4.line.width = Pt(2)
    add_text(sl, num, x+Inches(0.78), y, Inches(0.78), Inches(0.78),
             size=18, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    # Title
    add_text(sl, title, x, y+Inches(0.92), step_w, Inches(0.65),
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Desc
    add_text(sl, desc, x, y+Inches(1.58), step_w, Inches(1.1),
             size=8.5, color=col(0x88,0xAA,0xCC), align=PP_ALIGN.CENTER)
    # Connector arrow (not last)
    if i < 4:
        add_text(sl, "→", x+step_w+Inches(0.02), y+Inches(0.2), gap+Inches(0.1), Inches(0.42),
                 size=14, color=TEAL, align=PP_ALIGN.CENTER)

# CTA button (simulated)
add_rect(sl, Inches(4.36), Inches(6.0), Inches(3.0), Inches(0.52), fill=TEAL)
add_text(sl, "Schedule Discovery Workshop",
         Inches(4.36), Inches(6.02), Inches(3.0), Inches(0.48),
         size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

add_text(sl, "Contact: amigo.ai  ·  Backed by Madrona & Optum Ventures",
         Inches(0), Inches(6.65), W, Inches(0.4),
         size=9, color=col(0x55,0x77,0x99), align=PP_ALIGN.CENTER)

# ── SAVE ──────────────────────────────
out = "/home/user/Population-Health-Management-/Amigo_x_Fakeeh_Presentation.pptx"
prs.save(out)
print("Saved:", out)
