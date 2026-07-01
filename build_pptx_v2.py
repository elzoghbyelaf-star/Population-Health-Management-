"""
Amigo × Fakeeh Care Group — PPTX Presentation
Brand: Amigo actual identity (terracotta/rust red + dark navy + off-white + serif style)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── AMIGO BRAND COLORS ─────────────────────────────────────
RUST    = RGBColor(0xB4, 0x40, 0x30)   # Amigo primary terracotta/rust
RUST2   = RGBColor(0x96, 0x33, 0x24)   # darker rust
RUST_LT = RGBColor(0xCC, 0x55, 0x40)   # lighter rust
NAVY    = RGBColor(0x0D, 0x2B, 0x3D)   # Amigo dark navy/teal
NAVY2   = RGBColor(0x16, 0x3A, 0x50)   # slightly lighter navy
OFF     = RGBColor(0xF7, 0xF5, 0xF0)   # off-white background
CREAM   = RGBColor(0xEE, 0xEA, 0xE2)   # card background
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
NEAR_BLK= RGBColor(0x1A, 0x1A, 0x1A)
DARK_GR = RGBColor(0x2C, 0x3E, 0x50)
MID_GR  = RGBColor(0x5A, 0x6A, 0x7A)
LIGHT_GR= RGBColor(0xC8, 0xD0, 0xD8)
CARD_BG = RGBColor(0xF0, 0xED, 0xE8)   # warm card background

W = prs.slide_width
H = prs.slide_height
blank = prs.slide_layouts[6]

def rgb(r,g,b): return RGBColor(r,g,b)

# ── HELPERS ────────────────────────────────────────────────

def rect(slide, l, t, w, h, fill=None, line=None, lw=Pt(0.75), radius=0):
    s = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line; s.line.width = lw
    else:
        s.line.fill.background()
    return s

def oval(slide, l, t, w, h, fill=None, line=None, lw=Pt(1)):
    s = slide.shapes.add_shape(9, l, t, w, h)
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line; s.line.width = lw
    else:
        s.line.fill.background()
    return s

def txt(slide, text, l, t, w, h,
        size=12, bold=False, color=NEAR_BLK,
        align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color
    return txb

def mtxt(slide, lines, l, t, w, h,
         size=11, color=NEAR_BLK, align=PP_ALIGN.LEFT,
         bold=False, sp=Pt(3)):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame; tf.word_wrap = True
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align; p.space_before = sp
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.color.rgb = color
        r.font.bold = bold
    return txb

def logo_placeholder(slide, x, y):
    """Amigo logo — stacked squares icon + wordmark"""
    # Icon (simplified layered squares)
    for i, alpha in [(0,0.4),(1,0.7),(2,1.0)]:
        offset = Inches(i*0.04)
        s = slide.shapes.add_shape(1,
            x + offset, y + offset,
            Inches(0.28), Inches(0.28))
        s.fill.solid(); s.fill.fore_color.rgb = WHITE
        s.line.fill.background()
    txt(slide, "Amigo", x+Inches(0.33), y+Inches(0.04),
        Inches(0.8), Inches(0.28), size=13, bold=False,
        color=WHITE, align=PP_ALIGN.LEFT)

def logo_dark(slide, x, y):
    """Amigo logo on light background"""
    for i in range(3):
        offset = Inches(i*0.04)
        s = slide.shapes.add_shape(1,
            x + offset, y + offset,
            Inches(0.26), Inches(0.26))
        s.fill.solid(); s.fill.fore_color.rgb = RUST
        s.line.fill.background()
    txt(slide, "Amigo", x+Inches(0.31), y+Inches(0.03),
        Inches(0.8), Inches(0.28), size=13, bold=False,
        color=NEAR_BLK, align=PP_ALIGN.LEFT)

# ═══════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# Mimics Amigo's rust-red full-bleed title slide
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
# Full rust background
rect(sl, 0, 0, W, H, fill=RUST)
# Subtle bottom accent bars (from brand style)
for i,w_,a in [(0,Inches(3.0),0.6),(1,Inches(4.5),0.4),(2,Inches(2.0),0.3)]:
    shade = rgb(0x96,0x33,0x24)
    rect(sl, Inches(0.5+i*4.1), Inches(6.0), w_, Inches(1.6), fill=RUST2)

# Amigo logo top-right
logo_placeholder(sl, Inches(11.6), Inches(0.3))

# Partnership badge
rect(sl, Inches(0.55), Inches(1.4), Inches(3.8), Inches(0.38),
     fill=rgb(0x8A,0x30,0x22), line=None)
txt(sl, "CONFIDENTIAL  ·  PRESENTED TO FAKEEH CARE GROUP",
    Inches(0.65), Inches(1.42), Inches(3.6), Inches(0.34),
    size=8, bold=True, color=rgb(0xFF,0xCC,0xBB), align=PP_ALIGN.LEFT)

# Main title — large serif style
txt(sl, "Scaling Care at Fakeeh.",
    Inches(0.55), Inches(2.1), Inches(9.0), Inches(1.2),
    size=46, bold=False, color=WHITE, align=PP_ALIGN.LEFT)
txt(sl, "Powered by Amigo.",
    Inches(0.55), Inches(3.2), Inches(9.0), Inches(0.9),
    size=36, bold=False, color=rgb(0xFF,0xCC,0xBB), align=PP_ALIGN.LEFT)

# Subtitle
txt(sl,
    "How Amigo's patient-facing clinical AI agents can multiply Fakeeh Care Group's\n"
    "capacity — without scaling headcount.",
    Inches(0.55), Inches(4.25), Inches(8.5), Inches(0.85),
    size=14, color=rgb(0xFF,0xE0,0xD5), align=PP_ALIGN.LEFT)

# Meta chips
for i,(lbl) in enumerate(["🏥  Fakeeh Care Group", "🤖  Amigo AI", "🇸🇦  Vision 2030"]):
    rect(sl, Inches(0.55 + i*2.55), Inches(5.3), Inches(2.35), Inches(0.38),
         fill=rgb(0x8A,0x30,0x22))
    txt(sl, lbl, Inches(0.65 + i*2.55), Inches(5.32), Inches(2.2), Inches(0.34),
        size=9, color=rgb(0xFF,0xCC,0xBB))

# ═══════════════════════════════════════════════════════════
# SLIDE 2 — ABOUT AMIGO
# Off-white background with rust accents — matches brand
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=OFF)
# Top rule
rect(sl, 0, 0, W, Inches(0.06), fill=RUST)

logo_dark(sl, Inches(0.45), Inches(0.2))
txt(sl, "About Amigo", Inches(0.45), Inches(0.65), Inches(7.0), Inches(0.65),
    size=28, bold=False, color=NEAR_BLK)
txt(sl, "The trusted platform for building, training, and deploying patient-facing clinical AI agents.",
    Inches(0.45), Inches(1.25), Inches(7.5), Inches(0.45),
    size=13, color=MID_GR, italic=True)

# Divider
rect(sl, Inches(0.45), Inches(1.78), Inches(12.4), Inches(0.02), fill=LIGHT_GR)

# Left column — context block
rect(sl, Inches(0.45), Inches(1.92), Inches(5.8), Inches(1.5), fill=NAVY)
txt(sl, '"We understand the challenges faced\nby healthcare organizations."',
    Inches(0.65), Inches(2.02), Inches(5.4), Inches(1.3),
    size=14, bold=False, color=WHITE, italic=True)

# Challenges mini-grid
challenges = [
    ("Provider capacity outpaced\nby patient demand", "$1M annual revenue\nlost per provider"),
    ("43% physician\nburnout rate",                  "$500K replacement\ncost per physician"),
    ("2 hrs wasted per\nhour of patient care",       "$200 lost revenue\nper unfilled visit"),
]
for i,(prob,cost) in enumerate(challenges):
    x = Inches(0.45 + (i%2)*3.0)
    y = Inches(3.6 + (i//2)*0.8) if i < 2 else Inches(4.4)
    if i == 2: x = Inches(0.45)

for i, row in enumerate([
    ("Lost Revenue",      "Provider capacity outpaced by demand",    "$1M / year / provider"),
    ("High Turnover",     "43% physician burnout rate",               "$500K replacement cost"),
    ("High Costs",        "2 hrs wasted per hour of patient care",    "Clinical inefficiency"),
    ("Unfilled Schedules","No-shows & cancellations",                 "$200 lost per unfilled visit"),
    ("Patient Drop-Off",  "Cannot influence experience outside visit","$15K lifetime value lost"),
]):
    y = Inches(2.05 + i*0.98)
    rect(sl, Inches(0.45), y, Inches(5.8), Inches(0.88), fill=CREAM,
         line=LIGHT_GR, lw=Pt(0.5))
    rect(sl, Inches(0.45), y, Inches(0.06), Inches(0.88), fill=RUST)
    txt(sl, row[0], Inches(0.62), y+Inches(0.08), Inches(1.8), Inches(0.3),
        size=9, bold=True, color=RUST)
    txt(sl, row[1], Inches(0.62), y+Inches(0.4), Inches(3.0), Inches(0.4),
        size=8.5, color=DARK_GR)
    txt(sl, row[2], Inches(3.7), y+Inches(0.28), Inches(2.4), Inches(0.3),
        size=8, color=MID_GR, align=PP_ALIGN.RIGHT)

# Right column — Amigo stats + description
rect(sl, Inches(6.55), Inches(1.92), Inches(6.35), Inches(5.3), fill=WHITE,
     line=LIGHT_GR, lw=Pt(1))

txt(sl, "Amigo expands provider capacity.", Inches(6.75), Inches(2.0),
    Inches(5.9), Inches(0.5), size=16, bold=False, color=NEAR_BLK)

stat_data = [
    ("5X",    "CAPACITY",  "patient capacity per provider"),
    ("400%",  "REVENUE",   "increase in billable revenue"),
    ("25%",   "OUTCOMES",  "lift in patient outcome measures"),
    ("22.2x", "ROI",       "total return on investment (Ivim Health)"),
]
for i,(num,tag,lbl) in enumerate(stat_data):
    y = Inches(2.6 + i*1.1)
    rect(sl, Inches(6.75), y, Inches(5.9), Inches(1.0), fill=OFF,
         line=LIGHT_GR, lw=Pt(0.5))
    rect(sl, Inches(6.75), y, Inches(0.7), Inches(0.32),
         fill=RUST if i==0 else NAVY if i==1 else rgb(0x2C,0x6E,0x5A) if i==2 else rgb(0x4A,0x3A,0x70))
    txt(sl, tag, Inches(6.77), y+Inches(0.05), Inches(0.66), Inches(0.22),
        size=6, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, num, Inches(7.55), y+Inches(0.06), Inches(2.5), Inches(0.55),
        size=26, bold=False, color=NEAR_BLK)
    txt(sl, lbl, Inches(7.55), y+Inches(0.6), Inches(4.8), Inches(0.32),
        size=9, color=MID_GR)

# Footer
rect(sl, 0, Inches(7.15), W, Inches(0.35), fill=RUST)
txt(sl, "amigo.ai  ·  Backed by General Catalyst, GSV Ventures, Madrona Ventures & United Healthcare  ·  $20M raised",
    Inches(0.4), Inches(7.18), Inches(12.5), Inches(0.28),
    size=8, color=WHITE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════
# SLIDE 3 — FAKEEH AT A GLANCE
# Navy background — premium hospital feel
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=NAVY)
rect(sl, 0, 0, W, Inches(0.06), fill=RUST)

logo_placeholder(sl, Inches(0.45), Inches(0.22))
txt(sl, "Fakeeh Care Group at a Glance",
    Inches(0.45), Inches(0.62), Inches(9.0), Inches(0.6),
    size=26, bold=False, color=WHITE)
txt(sl, "Dubai-based private healthcare conglomerate with regional leadership — founded 1978, listed 2024.",
    Inches(0.45), Inches(1.2), Inches(9.5), Inches(0.38),
    size=12, color=LIGHT_GR, italic=True)
rect(sl, Inches(0.45), Inches(1.65), Inches(12.4), Inches(0.02), fill=rgb(0x2A,0x4A,0x60))

# KPI row (4 cards)
kpis = [
    ("SAR 3.1B", "+11% YoY", "FY2025 Revenue"),
    ("1.89M",    "+8% YoY",  "Annual Patients"),
    ("4 → 7",    "by 2028",  "Hospitals (beds: 835→1,675)"),
    ("900+",     "50+ specialties", "Doctors in Network"),
]
for i,(num,sub,lbl) in enumerate(kpis):
    x = Inches(0.4 + i*3.25)
    rect(sl, x, Inches(1.78), Inches(3.05), Inches(1.45),
         fill=rgb(0x14,0x32,0x48), line=rgb(0x25,0x50,0x70), lw=Pt(1))
    rect(sl, x, Inches(1.78), Inches(3.05), Inches(0.05), fill=RUST)
    txt(sl, num, x+Inches(0.18), Inches(1.88), Inches(2.7), Inches(0.58),
        size=24, bold=False, color=WHITE)
    txt(sl, sub, x+Inches(0.18), Inches(2.42), Inches(2.7), Inches(0.28),
        size=9, bold=True, color=RUST_LT)
    txt(sl, lbl, x+Inches(0.18), Inches(2.72), Inches(2.7), Inches(0.38),
        size=9, color=LIGHT_GR)

# Two info panels
for ci, (heading, bullets) in enumerate([
    ("Established Excellence", [
        "→  Founded 1978 by Dr. Soliman Fakeeh — Saudi Arabia's first private hospital",
        "→  JCI Accredited 6th consecutive time | ANCC Magnet Nursing Recognition",
        "→  CAP Lab: zero deficiencies | CEBAHI accredited",
        "→  Fakeeh University Hospital — flagship Dubai facility (Dubai Silicon Oasis)",
        "→  First private-sector robotic surgery program in the GCC",
        "→  Medicentres acquisition — 22-clinic primary care network across Dubai & UAE",
    ]),
    ("Ambitious Growth Agenda", [
        "→  Target: 7 hospitals · 1,675 beds · 9 medical centers by 2028",
        "→  HEAL Hospital (AED 500M+ Neuroscience centre) — construction underway 2025",
        "→  Expanding oncology & tertiary care across Dubai and the Northern Emirates",
        "→  Partnership with Dubai Health Authority & DHA-accredited facilities",
        "→  AED credit facility secured for GCC-wide expansion",
        "→  Headquarters: Dubai Silicon Oasis — positioned as UAE's healthcare hub",
    ]),
]):
    x = Inches(0.4 + ci*6.5)
    rect(sl, x, Inches(3.38), Inches(6.2), Inches(3.9),
         fill=rgb(0x0A,0x1E,0x2C), line=rgb(0x25,0x50,0x70), lw=Pt(1))
    txt(sl, heading, x+Inches(0.22), Inches(3.48), Inches(5.8), Inches(0.4),
        size=11, bold=True, color=RUST_LT)
    rect(sl, x+Inches(0.22), Inches(3.9), Inches(5.6), Inches(0.02), fill=rgb(0x25,0x50,0x70))
    mtxt(sl, bullets, x+Inches(0.18), Inches(3.98), Inches(5.85), Inches(3.1),
         size=9.5, color=LIGHT_GR, sp=Pt(4))

rect(sl, 0, Inches(7.15), W, Inches(0.35), fill=RUST)
txt(sl, "amigo.ai", Inches(0.4), Inches(7.18), Inches(12.5), Inches(0.28),
    size=8, color=WHITE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════
# SLIDE 4 — CHALLENGES FAKEEH FACES
# Off-white with dark card grid
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=OFF)
rect(sl, 0, 0, W, Inches(0.06), fill=RUST)

logo_dark(sl, Inches(0.45), Inches(0.22))
txt(sl, "We understand the challenges faced by Fakeeh",
    Inches(0.45), Inches(0.62), Inches(11.0), Inches(0.65),
    size=26, bold=False, color=NEAR_BLK)
txt(sl, "Healthcare organizations at Fakeeh's scale struggle to efficiently deliver care as demand outpaces capacity.",
    Inches(0.45), Inches(1.22), Inches(11.0), Inches(0.38),
    size=12, color=MID_GR, italic=True)
rect(sl, Inches(0.45), Inches(1.65), Inches(12.4), Inches(0.02), fill=LIGHT_GR)

# 6 challenge cards — dark cards with rust accent (brand style)
pain_data = [
    ("Lost Revenue",      "Explosive Patient Volume Growth",
     "1.89M patients growing 8%+ YoY. Clinical staff cannot scale at the same pace — every new hospital multiplies patient engagement burden.",
     "$1M annual revenue lost per provider at capacity"),
    ("Unfilled Schedules","Post-Visit Follow-Up Gap",
     "1.89M discharges annually. Systematically following up every patient for medication adherence and recovery is operationally impossible at current staffing.",
     "$200 lost per unfilled visit from no-shows"),
    ("High Costs",        "Multi-Language Patient Population",
     "Dubai serves a diverse population requiring care in Arabic, English, Urdu, Tagalog, Hindi and more — across all Fakeeh facilities.",
     "2 hours wasted per hour of patient care"),
    ("High Turnover",     "Rapid Multi-Site Expansion",
     "Scaling 4→7 hospitals and 9 medical centers by 2028 means replicating patient engagement operations without a scalable digital model.",
     "43% physician burnout rate nationally"),
    ("Patient Drop-Off",  "Chronic Disease Management Burden",
     "With 50+ specialties incl. cardiology, oncology, nephrology — managing ongoing patient engagement for chronic conditions consumes disproportionate bandwidth.",
     "$15K lifetime value lost per patient drop-off"),
    ("Missed Quality",    "Home Healthcare Coordination",
     "Fakeeh Home Health Care spans 6 cities. Remote monitoring and check-ins rely on manual outreach — creating inconsistent patient experience.",
     "Missed quality measures from care gaps"),
]
for i,(tag,title,desc,cost) in enumerate(pain_data):
    ci = i % 3; ri = i // 3
    x = Inches(0.4 + ci*4.32)
    y = Inches(1.82 + ri*2.72)
    rect(sl, x, y, Inches(4.1), Inches(2.55), fill=NEAR_BLK)
    # top rust label
    rect(sl, x, y, Inches(4.1), Inches(0.35), fill=RUST)
    txt(sl, tag, x+Inches(0.12), y+Inches(0.07), Inches(3.85), Inches(0.25),
        size=9, bold=True, color=WHITE)
    txt(sl, title, x+Inches(0.14), y+Inches(0.44), Inches(3.82), Inches(0.42),
        size=11, bold=False, color=WHITE)
    txt(sl, desc, x+Inches(0.14), y+Inches(0.9), Inches(3.82), Inches(1.1),
        size=8.5, color=LIGHT_GR)
    rect(sl, x+Inches(0.14), y+Inches(2.1), Inches(3.82), Inches(0.02), fill=rgb(0x40,0x40,0x40))
    txt(sl, cost, x+Inches(0.14), y+Inches(2.17), Inches(3.82), Inches(0.28),
        size=7.5, color=RUST_LT)

rect(sl, 0, Inches(7.15), W, Inches(0.35), fill=RUST)
txt(sl, "amigo.ai", Inches(0.4), Inches(7.18), Inches(12.5), Inches(0.28),
    size=8, color=WHITE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════
# SLIDE 5 — HOW AMIGO TRANSFORMS THE PATIENT JOURNEY
# Rust background — directly from Amigo deck slide 06
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=RUST)

logo_placeholder(sl, Inches(11.6), Inches(0.22))
txt(sl, "How Amigo transforms the patient care journey at Fakeeh",
    Inches(0.45), Inches(0.35), Inches(10.8), Inches(0.85),
    size=24, bold=False, color=WHITE)
txt(sl, "8 AI-powered touchpoints across the complete patient journey",
    Inches(0.45), Inches(1.18), Inches(10.0), Inches(0.35),
    size=12, color=rgb(0xFF,0xCC,0xBB), italic=True)

# 8 use cases in 2x4 grid (rust-tinted dark cards — exact Amigo style)
use_cases = [
    ("1","Proactive Patient Outreach",
     "Reactivate dormant patients, fill schedule gaps, outbound appointment booking"),
    ("2","Scheduling & Reminders",
     "Cut no-shows by 50% and free up front-desk staff across 120+ Fakeeh clinics"),
    ("3","Clinical Intake & Triage",
     "Eliminate wait times and route Fakeeh patients to the right specialty in seconds"),
    ("4","Patient Education",
     "Prep patients before arrival; cut visit times by 30% across all hospitals"),
    ("5","Provider Copilot",
     "Deliver clinical decision support and elevate patient outcomes for 900+ Fakeeh doctors"),
    ("6","Post-Visit Follow-ups",
     "Answer patient questions and free providers from callbacks — 1.89M patients annually"),
    ("7","Ongoing 24/7 Virtual Care",
     "Engage Fakeeh patients beyond their visit to improve chronic disease adherence"),
    ("8","RPM & Patient Insights",
     "Flag at-risk patients and gain visibility into population health across all 6 home health cities"),
]
for i,(num,title,desc) in enumerate(use_cases):
    ci = i % 4; ri = i // 4
    x = Inches(0.38 + ci*3.26)
    y = Inches(1.65 + ri*2.68)
    rect(sl, x, y, Inches(3.1), Inches(2.5), fill=RUST2)
    # number badge
    oval(sl, x+Inches(2.58), y+Inches(0.1), Inches(0.42), Inches(0.42),
         fill=rgb(0xFF,0xFF,0xFF,))
    txt(sl, num, x+Inches(2.58), y+Inches(0.12), Inches(0.42), Inches(0.35),
        size=9, bold=True, color=RUST, align=PP_ALIGN.CENTER)
    txt(sl, title, x+Inches(0.14), y+Inches(0.14), Inches(2.35), Inches(0.65),
        size=12, bold=False, color=WHITE)
    rect(sl, x+Inches(0.14), y+Inches(0.9), Inches(2.75), Inches(0.02),
         fill=rgb(0xCC,0x55,0x40))
    txt(sl, desc, x+Inches(0.14), y+Inches(1.0), Inches(2.82), Inches(1.35),
        size=8.5, color=rgb(0xFF,0xDD,0xD5))

rect(sl, 0, Inches(7.15), W, Inches(0.35), fill=RUST2)
txt(sl, "amigo.ai", Inches(0.4), Inches(7.18), Inches(12.5), Inches(0.28),
    size=8, color=rgb(0xFF,0xCC,0xBB), align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════
# SLIDE 6 — ONE PATIENT, EVERY TOUCHPOINT (Fakeeh version)
# White/off-white — narrative journey slide
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=OFF)
rect(sl, 0, 0, W, Inches(0.06), fill=RUST)

logo_dark(sl, Inches(0.45), Inches(0.22))
txt(sl, "One patient, every touchpoint.",
    Inches(0.45), Inches(0.58), Inches(7.0), Inches(0.65),
    size=28, bold=False, color=NEAR_BLK)

# Patient scenario box
rect(sl, Inches(7.3), Inches(0.18), Inches(5.6), Inches(1.2), fill=CREAM,
     line=LIGHT_GR, lw=Pt(1))
txt(sl, "Ahmed's chronic care journey at Fakeeh University Hospital, Dubai",
    Inches(7.5), Inches(0.26), Inches(5.2), Inches(0.38),
    size=11, bold=True, color=NEAR_BLK, italic=True)
txt(sl, "Ahmed, 58  ·  Diabetic & hypertensive  ·  Regular Fakeeh University Hospital patient, Dubai",
    Inches(7.5), Inches(0.63), Inches(5.2), Inches(0.38),
    size=9, color=MID_GR)
txt(sl, "Powered by Amigo",
    Inches(7.5), Inches(0.95), Inches(5.2), Inches(0.28),
    size=8, color=RUST, italic=True)

# Journey columns
stages = ["Onboarding:", "Pre-visit:", "Visit:", "Post-visit:", "Ongoing care:"]
stage_x = [Inches(0.35), Inches(2.82), Inches(5.28), Inches(7.75), Inches(10.22)]
stage_w = Inches(2.3)

for i,(stage,x) in enumerate(zip(stages,stage_x)):
    txt(sl, stage, x, Inches(1.48), stage_w, Inches(0.3),
        size=9, color=MID_GR, align=PP_ALIGN.CENTER)
    rect(sl, x, Inches(1.8), stage_w, Inches(0.02), fill=LIGHT_GR)

# Amigo speech bubbles
amigo_msgs = [
    '"Ahmed, you\'re overdue for a diabetes check. I\'ve found a slot Thursday at Fakeeh University Hospital. Want me to book it?"',
    '"Your appointment is tomorrow. Any new symptoms for Dr. Al-Rashidi? Here\'s what to bring..."',
    "Dr. Al-Mansoori reviews Ahmed's full history, flagged care gaps, and new HbA1c concerns before entering.",
    '"Your A1C is 7.2 — manageable! Dr. Al-Rashidi has started you on adjusted Metformin. Dietitian booked for next week."',
    '"You hit your 8,000 steps goal yesterday! How\'s the new Metformin dose feeling? Any stomach issues?"',
]
rect(sl, Inches(0.35), Inches(1.95), Inches(12.6), Inches(2.35), fill=CREAM,
     line=LIGHT_GR, lw=Pt(0.5))
txt(sl, "Amigo agent", Inches(0.45), Inches(2.0), Inches(1.0), Inches(0.28),
    size=8, bold=True, color=RUST)
for i,(msg,x) in enumerate(zip(amigo_msgs,stage_x)):
    txt(sl, msg, x+Inches(0.06), Inches(2.28), stage_w-Inches(0.12), Inches(2.0),
        size=8, color=DARK_GR, italic=True, align=PP_ALIGN.LEFT)

# Touchpoint rows
touchpoint_rows = [
    ["Overdue visit flagged\nOutbound appt. booked\nInsurance verified",
     "Intake completed\nAppt. reminder sent\nPatient prep info shared",
     "Provider fully briefed\nCopilot ensures nothing missed\nPersonalized session",
     "Lab results in plain Arabic\nMetformin Rx drafted\nDietitian referred",
     "24/7 health coaching\nProgress check-ins\nMedication adherence"],
]
rect(sl, Inches(0.35), Inches(4.45), Inches(12.6), Inches(2.38), fill=WHITE,
     line=LIGHT_GR, lw=Pt(0.5))
txt(sl, "Touchpoints", Inches(0.45), Inches(4.5), Inches(1.2), Inches(0.28),
    size=8, bold=True, color=MID_GR)
for i,(pts,x) in enumerate(zip(touchpoint_rows[0],stage_x)):
    for j,pt in enumerate(pts.split('\n')):
        # bullet dot
        oval(sl, x+Inches(0.08), Inches(4.78+j*0.5)+Inches(0.08),
             Inches(0.1), Inches(0.1), fill=RUST)
        txt(sl, pt, x+Inches(0.26), Inches(4.76+j*0.5),
            stage_w-Inches(0.28), Inches(0.42),
            size=7.5, color=DARK_GR)

# Bottom summary
rect(sl, Inches(0.35), Inches(6.95), Inches(12.6), Inches(0.38), fill=RUST)
txt(sl,
    "Amigo agents handle front- and back-office tasks to free up Fakeeh staff and make patients feel cared for throughout their journey.",
    Inches(0.5), Inches(6.97), Inches(12.4), Inches(0.34),
    size=9, bold=False, color=WHITE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════
# SLIDE 7 — THE AMIGO ADVANTAGE
# White — comparison table (exact Amigo style)
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=OFF)
rect(sl, 0, 0, W, Inches(0.06), fill=RUST)

logo_dark(sl, Inches(0.45), Inches(0.22))
txt(sl, "The Amigo advantage",
    Inches(0.45), Inches(0.6), Inches(6.0), Inches(0.65),
    size=28, bold=False, color=NEAR_BLK)
txt(sl, "Why Amigo — not a generic chatbot, not a point solution.",
    Inches(0.45), Inches(1.22), Inches(9.0), Inches(0.38),
    size=12, color=MID_GR, italic=True)
rect(sl, Inches(0.45), Inches(1.62), Inches(12.4), Inches(0.02), fill=LIGHT_GR)

# Column headers
txt(sl, "Amigo", Inches(7.2), Inches(1.72), Inches(2.5), Inches(0.4),
    size=14, bold=False, color=NAVY, align=PP_ALIGN.CENTER)
txt(sl, "Other Healthcare Agents", Inches(9.9), Inches(1.72), Inches(3.0), Inches(0.4),
    size=12, color=MID_GR, align=PP_ALIGN.CENTER)

# Table rows
rows = [
    "One unified platform for all agent use cases",
    "Custom-built for your patient population",
    "Deploy in weeks with deep EHR integrations",
    "Dedicated Agent Engineer",
    "Transparent and auditable reasoning",
    "Simulation-validated evaluations (Digital Residency™)",
    "Continuous agent self-improvement",
    "You own all agent IP and patient data",
]
for i,row in enumerate(rows):
    y = Inches(2.22 + i*0.6)
    bg = WHITE if i%2==0 else rgb(0xF0,0xED,0xE8)
    rect(sl, Inches(0.45), y, Inches(12.4), Inches(0.56), fill=bg,
         line=LIGHT_GR, lw=Pt(0.5))
    txt(sl, row, Inches(0.62), y+Inches(0.1), Inches(6.4), Inches(0.38),
        size=10.5, color=DARK_GR)
    # Amigo checkmark
    oval(sl, Inches(8.0), y+Inches(0.1), Inches(0.36), Inches(0.36), fill=NAVY)
    txt(sl, "✓", Inches(8.0), y+Inches(0.1), Inches(0.36), Inches(0.36),
        size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Others X
    oval(sl, Inches(10.9), y+Inches(0.1), Inches(0.36), Inches(0.36),
         fill=rgb(0xE8,0xE8,0xE8))
    txt(sl, "✕", Inches(10.9), y+Inches(0.1), Inches(0.36), Inches(0.36),
        size=10, bold=True, color=MID_GR, align=PP_ALIGN.CENTER)

# GCC-specific note
rect(sl, Inches(0.45), Inches(7.0), Inches(12.4), Inches(0.38),
     fill=rgb(0xF5,0xEC,0xE8), line=RUST, lw=Pt(1))
txt(sl,
    "🇦🇪  Amigo has an active partnership with Heal (UAE/GCC) — proven regional deployment track. Arabic-first, 100+ languages, HIPAA · SOC 2 · GDPR compliant.",
    Inches(0.62), Inches(7.02), Inches(12.1), Inches(0.34),
    size=9, color=RUST2)

rect(sl, 0, Inches(7.4), W, Inches(0.1), fill=RUST)

# ═══════════════════════════════════════════════════════════
# SLIDE 8 — YOUR AI PARTNER, END TO END
# White with rust accents — partnership process
# ═══════════════════════════════════════════════════════════
sl = prs.slides.add_slide(blank)
rect(sl, 0, 0, W, H, fill=NAVY)

logo_placeholder(sl, Inches(11.6), Inches(0.22))
txt(sl, "Your AI partner, end to end.",
    Inches(0.45), Inches(0.55), Inches(9.0), Inches(0.75),
    size=30, bold=False, color=WHITE)
txt(sl,
    "From discovery to post-development, our team handles the complexity so Fakeeh can stay focused on care.",
    Inches(0.45), Inches(1.28), Inches(10.0), Inches(0.42),
    size=13, color=LIGHT_GR, italic=True)
rect(sl, Inches(0.45), Inches(1.75), Inches(12.4), Inches(0.02), fill=rgb(0x25,0x50,0x70))

# 6-step process (2 columns x 3 rows)
steps = [
    ("1","Discovery",
     "We deeply understand Fakeeh's unique workflows and goals — not pre-built templates."),
    ("2","Custom Agent Build",
     "Tailor agents for your needs: Arabic intake, post-visit follow-up, chronic care, home health."),
    ("3","Rigorous Testing",
     "Millions of simulated sessions before launch — accuracy, brand consistency, edge case handling."),
    ("4","Development in 6 Weeks",
     "Full engineering and IT support. Integrate with Fakeeh's YASASII HIS & EHRs. Go live in 6 weeks."),
    ("5","Monitoring",
     "Continuous monitoring with real-time risk flags, full visibility into agent performance and ROI."),
    ("6","Continuous Improvement",
     "We refine your agents based on real Fakeeh patient interactions so they get smarter over time."),
]
for i,(num,title,desc) in enumerate(steps):
    ci = i % 2; ri = i // 2
    x = Inches(0.45 + ci*6.55)
    y = Inches(2.0 + ri*1.62)
    rect(sl, x, y, Inches(6.2), Inches(1.48),
         fill=rgb(0x14,0x32,0x48), line=rgb(0x25,0x50,0x70), lw=Pt(1))
    # number circle
    oval(sl, x+Inches(0.15), y+Inches(0.15), Inches(0.68), Inches(0.68), fill=RUST)
    txt(sl, num, x+Inches(0.15), y+Inches(0.17), Inches(0.68), Inches(0.6),
        size=14, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
    txt(sl, title, x+Inches(1.0), y+Inches(0.1), Inches(5.0), Inches(0.42),
        size=13, bold=False, color=WHITE)
    txt(sl, desc, x+Inches(1.0), y+Inches(0.55), Inches(5.0), Inches(0.8),
        size=9.5, color=LIGHT_GR)

# Timeline arrow
rect(sl, Inches(0.45), Inches(6.92), Inches(12.4), Inches(0.03), fill=RUST)
txt(sl, "6 weeks from concept to live agents →",
    Inches(0.45), Inches(6.95), Inches(5.0), Inches(0.3),
    size=9, bold=True, color=RUST_LT)
txt(sl, "Schedule your discovery session:  amigo.ai",
    Inches(8.5), Inches(6.95), Inches(4.4), Inches(0.3),
    size=9, color=LIGHT_GR, align=PP_ALIGN.RIGHT)

# CTA bar
rect(sl, Inches(3.8), Inches(7.08), Inches(5.73), Inches(0.38), fill=RUST)
txt(sl, "Let's build the future of care at Fakeeh together.",
    Inches(3.95), Inches(7.1), Inches(5.4), Inches(0.34),
    size=10, bold=False, color=WHITE, align=PP_ALIGN.CENTER)

# Save
out = "/home/user/Population-Health-Management-/Amigo_x_Fakeeh_v2.pptx"
prs.save(out)
print("Saved:", out)
