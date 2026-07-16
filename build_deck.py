#!/usr/bin/env python3
"""
Amigo x Cleveland Diagnostics — partnership opportunity deck.
Theme matched to amigo_summary.svg (navy #1E3A8A + sky #0EA5E9, white cards,
soft shadows, pill tags, stat tiles, "Population Health Intelligence").
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ---- Brand palette (from amigo_summary.svg) ----
NAVY      = RGBColor(0x1E, 0x3A, 0x8A)
SKY       = RGBColor(0x0E, 0xA5, 0xE9)
BG1       = RGBColor(0xF8, 0xFA, 0xFF)
BG2       = RGBColor(0xEE, 0xF2, 0xFB)
CARD_TOP  = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BOT  = RGBColor(0xF0, 0xF5, 0xFF)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
INK       = RGBColor(0x33, 0x41, 0x55)   # body text
MUTE      = RGBColor(0x64, 0x74, 0x8B)   # secondary
FAINT     = RGBColor(0x94, 0xA3, 0xB8)   # tertiary
PALEBLUE  = RGBColor(0xBF, 0xDB, 0xFE)   # on-navy subtext
CORAL     = RGBColor(0xEF, 0x44, 0x44)   # pain-point accent

SERIF = "Georgia"
SANS  = "Segoe UI"

EMU = 914400
SW, SH = 13.333, 7.5

prs = Presentation()
prs.slide_width  = Emu(int(SW * EMU))
prs.slide_height = Emu(int(SH * EMU))
BLANK = prs.slide_layouts[6]


# ---------------- helpers ----------------
def _set_grad(shape, c1, c2, angle=45):
    """Linear gradient fill on a shape via raw XML."""
    sp = shape.fill._xPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        e = sp.find(qn(tag))
        if e is not None:
            sp.remove(e)
    grad = sp.makeelement(qn("a:gradFill"), {})
    lst = grad.makeelement(qn("a:gsLst"), {})
    for pos, col in ((0, c1), (100000, c2)):
        gs = grad.makeelement(qn("a:gs"), {"pos": str(pos)})
        clr = grad.makeelement(qn("a:srgbClr"), {"val": "%02X%02X%02X" % (col[0], col[1], col[2])})
        gs.append(clr)
        lst.append(gs)
    grad.append(lst)
    lin = grad.makeelement(qn("a:lin"), {"ang": str(int(angle * 60000)), "scaled": "1"})
    grad.append(lin)
    ln = sp.find(qn("a:ln"))
    sp.insert(list(sp).index(ln) if ln is not None else len(sp), grad)


def _shadow(shape, blur=0.06, dist=0.03, alpha=88):
    spPr = shape._element.spPr
    ef = spPr.makeelement(qn("a:effectLst"), {})
    sh = spPr.makeelement(qn("a:outerShdw"), {
        "blurRad": str(int(blur * EMU)), "dist": str(int(dist * EMU)),
        "dir": "5400000", "rotWithShape": "0"})
    clr = spPr.makeelement(qn("a:srgbClr"), {"val": "1E3A8A"})
    a = spPr.makeelement(qn("a:alpha"), {"val": str((100 - alpha) * 1000)})
    clr.append(a); sh.append(clr); ef.append(sh); spPr.append(ef)


def slide():
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.line.fill.background()
    _set_grad(bg, BG1, BG2, 45)
    bg.shadow.inherit = False
    return s


def rect(s, x, y, w, h, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def card(s, x, y, w, h, radius=0.06, accent=False, shadow=True):
    sp = rect(s, x, y, w, h, radius=radius)
    _set_grad(sp, CARD_TOP, CARD_BOT, 90)
    sp.line.color.rgb = NAVY
    sp.line.width = Pt(0.75)
    ln = sp.line._get_or_add_ln()
    ln.find(qn("a:solidFill")).find(qn("a:srgbClr")).append(
        ln.makeelement(qn("a:alpha"), {"val": "18000"}))
    if shadow:
        _shadow(sp)
    if accent:
        bar = rect(s, x, y, w, 0.07, radius=0.5)
        _set_grad(bar, NAVY, SKY, 0)
        bar.line.fill.background()
    return sp


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=4, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph = list of (text, size, color, bold, font, italic)."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (t, sz, col, bold, fnt, ital) in para:
            r = p.add_run(); r.text = t
            f = r.font
            f.size = Pt(sz); f.bold = bold; f.italic = ital
            f.color.rgb = col; f.name = fnt
    return tb


def R(t, sz, col=INK, bold=False, fnt=SANS, ital=False):
    return (t, sz, col, bold, fnt, ital)


def topbar(s):
    b = rect(s, 0, 0, SW, 0.07, shape=MSO_SHAPE.RECTANGLE)
    _set_grad(b, NAVY, SKY, 0)
    b.line.fill.background()
    bb = rect(s, 0, SH - 0.07, SW, 0.07, shape=MSO_SHAPE.RECTANGLE)
    _set_grad(bb, NAVY, SKY, 0)
    bb.line.fill.background()


def brandmark(s):
    txt(s, 0.55, 0.22, 3.0, 0.5, [[R("Amigo", 20, NAVY, False, SERIF)]])
    u = rect(s, 0.57, 0.68, 0.85, 0.03, shape=MSO_SHAPE.RECTANGLE)
    _set_grad(u, NAVY, SKY, 0); u.line.fill.background()


def footer(s, n):
    txt(s, 0.55, SH - 0.5, 6, 0.3,
        [[R("Amigo  ·  Population Health Intelligence", 8.5, FAINT)]])
    txt(s, SW - 2.05, SH - 0.5, 1.5, 0.3,
        [[R(f"{n} / 8", 8.5, FAINT)]], align=PP_ALIGN.RIGHT)


def eyebrow(s, kicker, title, x=0.55, y=0.95, w=12.2):
    txt(s, x, y, w, 0.35, [[R(kicker.upper(), 11.5, SKY, True)]])
    txt(s, x, y + 0.32, w, 0.75, [[R(title, 27, NAVY, True)]])
    u = rect(s, x + 0.02, y + 0.95, 0.85, 0.035, shape=MSO_SHAPE.RECTANGLE)
    _set_grad(u, NAVY, SKY, 0); u.line.fill.background()


def pill(s, x, y, w, text, on_navy=False):
    p = rect(s, x, y, w, 0.42, radius=0.5)
    if on_navy:
        _set_grad(p, NAVY, SKY, 0); p.line.fill.background()
        col = WHITE
    else:
        p.fill.solid(); p.fill.fore_color.rgb = WHITE
        p.line.color.rgb = NAVY; p.line.width = Pt(1)
        _shadow(p, blur=0.04, dist=0.02, alpha=90)
        col = NAVY
    txt(s, x, y + 0.02, w, 0.4, [[R(text, 11, col, True)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# SLIDE 1 — TITLE
# ============================================================
s = slide(); topbar(s)
txt(s, 0.7, 0.55, 4, 0.6, [[R("Amigo", 30, NAVY, False, SERIF)]])
u = rect(s, 0.72, 1.18, 1.1, 0.035, shape=MSO_SHAPE.RECTANGLE)
_set_grad(u, NAVY, SKY, 0); u.line.fill.background()

# center hero
txt(s, 1, 2.35, 11.33, 0.5,
    [[R("PARTNERSHIP  OPPORTUNITY", 14, SKY, True)]], align=PP_ALIGN.CENTER)
txt(s, 0.8, 2.8, 11.73, 1.7,
    [[R("Amigo × Cleveland Diagnostics", 42, NAVY, True)],
     [R("Scaling early cancer detection from a validated test", 21, INK, False)],
     [R("to a population-scale screening engine", 21, INK, False)]],
    align=PP_ALIGN.CENTER, line_spacing=1.05, space_after=2)

band = rect(s, 3.17, 4.95, 7.0, 0.045, shape=MSO_SHAPE.RECTANGLE)
_set_grad(band, NAVY, SKY, 0); band.line.fill.background()

txt(s, 1, 5.2, 11.33, 0.5,
    [[R("AI agents that identify eligible patients, drive adoption, and generate real-world evidence — for IsoPSA and beyond",
        13, MUTE, False, SANS, True)]], align=PP_ALIGN.CENTER)

# footer credit
txt(s, 0.7, SH - 0.55, 8, 0.3, [[R("Prepared for Cleveland Diagnostics, Inc.  ·  2026", 10.5, FAINT)]])
txt(s, SW - 4.7, SH - 0.55, 4, 0.3, [[R("elzoghby.elaf@amigo.ai", 10.5, SKY, True)]], align=PP_ALIGN.RIGHT)


# ============================================================
# SLIDE 2 — CLEVELAND DIAGNOSTICS SUMMARY (them)
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 2)
eyebrow(s, "The Company", "Cleveland Diagnostics at a Glance")

# stat tiles row
stats = [("2025", "FDA PMA approval of IsoPSA (Dec 1)"),
         ("$75M+", "Growth capital raised (2024, Novo Holdings)"),
         ("~90%", "Sensitivity for high-grade cancer"),
         ("0.78", "Validation AUC (n≈1,093, 8 sites)")]
tw, gap, x0, ty = 2.85, 0.2, 0.55, 1.95
for i, (big, lab) in enumerate(stats):
    x = x0 + i * (tw + gap)
    card(s, x, ty, tw, 1.15)
    txt(s, x, ty + 0.17, tw, 0.55, [[R(big, 30, NAVY, True)]], align=PP_ALIGN.CENTER)
    txt(s, x + 0.1, ty + 0.72, tw - 0.2, 0.4, [[R(lab, 10, MUTE)]], align=PP_ALIGN.CENTER)

# left: profile
card(s, 0.55, 3.3, 6.1, 3.5, accent=True)
txt(s, 0.85, 3.55, 5.6, 0.4, [[R("WHO THEY ARE", 12, NAVY, True)]])
prof = [
    ("Precision-oncology biotech", "Cleveland, OH — rooted in predecessor AnalizaDx."),
    ("Core platform: SIA / IsoClear", "Reads protein structure, not concentration, to find cancer isoforms."),
    ("Lead product: IsoPSA", "Stratifies high-grade prostate-cancer risk; aids the biopsy decision."),
    ("Regulatory milestone", "FDA Premarket Approval as an IVD kit (Dec 2025)."),
    ("Leadership", "Michael Iskra named CEO (Jan 2026) to drive scale-up."),
]
yy = 3.92
for h, d in prof:
    txt(s, 0.9, yy, 5.5, 0.3, [[R("●  ", 10, SKY, True), R(h, 11.5, NAVY, True)]])
    txt(s, 1.15, yy + 0.25, 5.3, 0.3, [[R(d, 10, INK)]], line_spacing=1.0)
    yy += 0.57

# right: evidence + focus
card(s, 6.85, 3.3, 5.93, 3.5, accent=True)
txt(s, 7.15, 3.55, 5.3, 0.4, [[R("EVIDENCE & POSITIONING", 12, NAVY, True)]])
ev = [
    "Prospective, multicenter validation (Klein et al., 2022)",
    "Outperformed total PSA and % free PSA on AUC & specificity",
    "SUO 2025: accuracy shown with and without mpMRI",
    "Competes with mpMRI pathways, 4Kscore, PHI",
]
yy = 3.98
for e in ev:
    txt(s, 7.2, yy, 5.4, 0.4, [[R("✓  ", 11, SKY, True), R(e, 10.5, INK)]])
    yy += 0.44
txt(s, 7.15, 5.9, 5.3, 0.3, [[R("SINGLE-PRODUCT TODAY — PIPELINE AHEAD", 10.5, NAVY, True)]])
txt(s, 7.15, 6.2, 5.5, 0.5,
    [[R("IsoPSA is the sole commercial product; the SIA platform can extend to other cancers.", 10, MUTE, False, SANS, True)]])


# ============================================================
# SLIDE 3 — THEIR PAIN POINTS
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 3)
eyebrow(s, "The Challenge", "Strategic Pain Points After FDA Approval")
txt(s, 0.55, 1.95, 12.2, 0.4,
    [[R("Approval is won. The harder problem is now commercial — turning a validated test into standard-of-care at national scale.",
        13, MUTE, False, SANS, True)]])

pains = [
    ("Commercial scale-up", "Moving from a niche lab-developed test to nationwide IVD adoption — with a salesforce that can't reach every urologist and PCP."),
    ("Patient identification", "The right patients (men 50+, rising PSA) surface in primary care, where risk-stratification is inconsistent and IsoPSA is rarely top-of-mind."),
    ("Physician & payer adoption", "Ordering behavior is hard to shift; reimbursement and payer coverage remain unclear and slow the funnel."),
    ("Real-world evidence gap", "Payers and guideline bodies want real-world outcomes; key data (e.g. SUO 2025) is conference-stage, not yet peer-reviewed."),
    ("Competitive pressure", "mpMRI pathways, 4Kscore and PHI compete for the same biopsy-decision moment — differentiation must be continually proven."),
    ("Single-product concentration", "Revenue rests entirely on IsoPSA; pipeline expansion needs large longitudinal datasets and validation partners."),
]
cw, ch, gx, gy = 4.0, 1.75, 0.28, 0.28
x0, y0 = 0.55, 2.55
for i, (h, d) in enumerate(pains):
    r, c = divmod(i, 3)
    x = x0 + c * (cw + gx); y = y0 + r * (ch + gy)
    cd = card(s, x, y, cw, ch)
    tab = rect(s, x, y, 0.09, ch, shape=MSO_SHAPE.RECTANGLE)
    tab.fill.solid(); tab.fill.fore_color.rgb = CORAL; tab.line.fill.background()
    txt(s, x + 0.28, y + 0.18, cw - 0.45, 0.4, [[R(f"{i+1}.  ", 13, CORAL, True), R(h, 13.5, NAVY, True)]])
    txt(s, x + 0.28, y + 0.62, cw - 0.5, 1.05, [[R(d, 10.3, INK)]], line_spacing=1.02)


# ============================================================
# SLIDE 4 — PAIN POINTS -> AMIGO USE CASES (solve them)
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 4)
eyebrow(s, "The Fit", "Solving the Pain Points with Amigo Agents")
txt(s, 0.55, 1.95, 12.2, 0.4,
    [[R("Each Amigo AI agent maps directly onto a Cleveland Diagnostics gap — not a dashboard, an agent that acts.",
        13, MUTE, False, SANS, True)]])

# header row
hy = 2.5
for x, w, t in [(0.55, 4.3, "PAIN POINT"), (4.95, 3.35, "AMIGO AGENT"), (8.4, 4.38, "HOW IT SOLVES IT")]:
    hd = rect(s, x, hy, w, 0.45, radius=0.12)
    _set_grad(hd, NAVY, SKY, 0); hd.line.fill.background()
    txt(s, x + 0.15, hy + 0.02, w - 0.2, 0.4, [[R(t, 11, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)

rows = [
    ("Finding eligible patients", "Risk Stratification", "Predictive scoring flags men 50+ with rising PSA across the population — a ready demand funnel for IsoPSA."),
    ("Patients slip through primary care", "Observation", "24/7 surveillance detects newly elevated PSA labs in real time and triggers an IsoPSA recommendation at the point of care."),
    ("Adoption & ordering behavior", "Campaign Tracking", "Closed-loop screening campaigns with live dashboards drive physician ordering and patient follow-through."),
    ("Real-world evidence gap", "Patient Profile (HIE)", "360° longitudinal records generate real-world outcomes data for payers, guidelines and label expansion."),
    ("Payer value & pipeline", "Simulation", "Digital-twin engine models biopsy-avoidance ROI for payer dossiers and forecasts new SIA biomarker programs."),
]
ry = hy + 0.55
rh = 0.78
for i, (pn, ag, hw) in enumerate(rows):
    y = ry + i * (rh + 0.06)
    c1 = card(s, 0.55, y, 4.3, rh, shadow=True)
    txt(s, 0.75, y + 0.05, 4.0, rh - 0.1, [[R(pn, 11.5, NAVY, True)]], anchor=MSO_ANCHOR.MIDDLE)
    c2 = rect(s, 4.95, y, 3.35, rh, radius=0.1)
    _set_grad(c2, CARD_TOP, CARD_BOT, 90)
    c2.line.color.rgb = SKY; c2.line.width = Pt(1.25); _shadow(c2)
    txt(s, 5.1, y + 0.05, 3.1, rh - 0.1, [[R("⚡  ", 12, SKY, True), R(ag, 12, NAVY, True)]], anchor=MSO_ANCHOR.MIDDLE)
    c3 = card(s, 8.4, y, 4.38, rh, shadow=True)
    txt(s, 8.6, y + 0.06, 4.05, rh - 0.12, [[R(hw, 9.8, INK)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)


# ============================================================
# SLIDE 5 — HIGH-VALUE USE CASES (beneficial)
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 5)
eyebrow(s, "The Opportunity", "High-Value Use Cases for Cleveland Diagnostics")

ucs = [
    ("◎", "Population-scale demand generation",
     "Screen millions of longitudinal records to surface every IsoPSA-eligible man and route them to ordering providers — a continuous, self-refilling top of funnel."),
    ("⟳", "Reflex screening pathway",
     "Elevated PSA in primary care automatically triggers an IsoPSA reflex recommendation, closing the gap between lab result and risk-stratified biopsy decision."),
    ("♡", "Real-world evidence engine",
     "HIE-powered outcomes tracking builds the peer-review-grade real-world dataset payers and guideline committees require for coverage and inclusion."),
    ("⚡", "Payer value modeling",
     "Simulate biopsy-avoidance and cost savings per population to arm reimbursement negotiations with quantified, defensible ROI."),
    ("▶", "Guideline & literature intelligence",
     "The NLP engine (2,000+ papers/month) keeps medical affairs ahead of evidence, competitors and guideline shifts in prostate diagnostics."),
    ("✦", "Pipeline acceleration beyond IsoPSA",
     "Reuse Amigo's cancer focus areas + longitudinal cohorts to validate future SIA biomarkers — de-risking the single-product concentration."),
]
cw, ch, gx, gy = 4.0, 1.85, 0.28, 0.25
x0, y0 = 0.55, 2.2
for i, (ic, h, d) in enumerate(ucs):
    r, c = divmod(i, 3)
    x = x0 + c * (cw + gx); y = y0 + r * (ch + gy)
    card(s, x, y, cw, ch, accent=True)
    bdg = rect(s, x + 0.25, y + 0.28, 0.55, 0.55, radius=0.5)
    _set_grad(bdg, NAVY, SKY, 45); bdg.line.fill.background()
    txt(s, x + 0.25, y + 0.3, 0.55, 0.5, [[R(ic, 16, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 0.95, y + 0.28, cw - 1.15, 0.55, [[R(h, 12.5, NAVY, True)]], line_spacing=0.95, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 0.28, y + 0.92, cw - 0.5, 0.85, [[R(d, 10, INK)]], line_spacing=1.0)


# ============================================================
# SLIDE 6 — PARTNERSHIP VALUE / IMPACT & ROADMAP
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 6)
eyebrow(s, "The Impact", "What the Partnership Delivers")

# expected impact tiles
imp = [("↑ Funnel", "More eligible patients identified and routed to IsoPSA"),
       ("↑ Adoption", "Closed-loop campaigns lift provider ordering"),
       ("↑ Evidence", "Real-world data for payers & guidelines"),
       ("↓ Risk", "Pipeline diversification beyond a single test")]
tw, gap, x0, ty = 2.95, 0.18, 0.55, 2.05
for i, (big, lab) in enumerate(imp):
    x = x0 + i * (tw + gap)
    card(s, x, ty, tw, 1.15, accent=True)
    txt(s, x, ty + 0.22, tw, 0.5, [[R(big, 19, NAVY, True)]], align=PP_ALIGN.CENTER)
    txt(s, x + 0.12, ty + 0.68, tw - 0.24, 0.42, [[R(lab, 9.7, MUTE)]], align=PP_ALIGN.CENTER, line_spacing=0.95)

# roadmap
txt(s, 0.55, 3.55, 12, 0.35, [[R("A PHASED ROADMAP", 12, NAVY, True)]])
phases = [
    ("Weeks 0–6", "Deploy & integrate", "Connect to HIE / lab feeds and stand up the eligible-patient model — Amigo deploys in as little as 6 weeks."),
    ("Quarter 1", "Reflex pathway live", "Elevated-PSA surveillance triggers IsoPSA recommendations; first adoption campaign runs with live dashboards."),
    ("Quarters 2–3", "Evidence & payers", "Longitudinal outcomes accrue into an RWE dataset; simulation powers payer value dossiers."),
    ("Quarter 4+", "Expand the platform", "Extend agents to new SIA biomarkers and cancer programs across the national footprint."),
]
pw, gx, x0, py = 3.02, 0.18, 0.55, 3.95
for i, (ph, h, d) in enumerate(phases):
    x = x0 + i * (pw + gx)
    card(s, x, py, pw, 2.55)
    bar = rect(s, x, py, pw, 0.07, radius=0.5)
    _set_grad(bar, NAVY, SKY, 0); bar.line.fill.background()
    txt(s, x + 0.22, py + 0.24, pw - 0.4, 0.35, [[R(ph, 11, SKY, True)]])
    txt(s, x + 0.22, py + 0.58, pw - 0.4, 0.55, [[R(h, 13.5, NAVY, True)]], line_spacing=0.95)
    txt(s, x + 0.22, py + 1.18, pw - 0.42, 1.2, [[R(d, 10, INK)]], line_spacing=1.05)
    if i < 3:
        txt(s, x + pw - 0.02, py + 0.9, 0.3, 0.4, [[R("→", 18, SKY, True)]], align=PP_ALIGN.CENTER)


# ============================================================
# SLIDE 7 — AMIGO SUMMARY (us)
# ============================================================
s = slide(); topbar(s); brandmark(s); footer(s, 7)
eyebrow(s, "The Partner", "Amigo — Population Health Intelligence")

# stat row
stats = [("22.2x", "Return on investment"), ("6M+", "Patient interactions"),
         ("6 wks", "To full deployment"), ("100%", "Hospital coverage")]
tw, gap, x0, ty = 2.95, 0.18, 0.55, 1.95
for i, (big, lab) in enumerate(stats):
    x = x0 + i * (tw + gap)
    card(s, x, ty, tw, 1.05)
    txt(s, x, ty + 0.16, tw, 0.5, [[R(big, 27, NAVY, True)]], align=PP_ALIGN.CENTER)
    txt(s, x + 0.1, ty + 0.66, tw - 0.2, 0.35, [[R(lab, 10, MUTE)]], align=PP_ALIGN.CENTER)

# left: agent architecture
card(s, 0.55, 3.2, 6.1, 3.45, accent=True)
txt(s, 0.85, 3.45, 5.6, 0.4, [[R("AI AGENT ARCHITECTURE", 12, NAVY, True)]])
agents = [
    ("Observation", "Live population surveillance, 24/7"),
    ("Risk Stratification", "Identify high-risk patients before crisis"),
    ("Simulation", "Forecast intervention outcomes (digital twin)"),
    ("Campaign Tracking", "Closed-loop action with live dashboards"),
    ("Patient Profile", "360° longitudinal view, HIE-powered"),
]
yy = 3.9
for h, d in agents:
    txt(s, 0.9, yy, 5.6, 0.3, [[R("⚡  ", 11, SKY, True), R(h + " — ", 11.5, NAVY, True), R(d, 10.5, INK)]])
    yy += 0.53

# right: why amigo + compliance
card(s, 6.85, 3.2, 5.93, 3.45, accent=True)
txt(s, 7.15, 3.45, 5.3, 0.4, [[R("WHY AMIGO", 12, NAVY, True)]])
why = [
    "Not dashboards — AI agents that act",
    "Built on billions of longitudinal health records",
    "NLP engine processing 2,000+ papers / month",
    "Clinical + administrative intelligence, unified",
    "8 clinical focus areas — incl. 4 cancers",
]
yy = 3.9
for w in why:
    txt(s, 7.2, yy, 5.5, 0.3, [[R("✦  ", 11, SKY, True), R(w, 11, INK)]])
    yy += 0.42
txt(s, 7.15, 6.02, 5.5, 0.3,
    [[R("SOC 2 Type II · HIPAA · GDPR · Local Data Residency", 10, NAVY, True)]])
txt(s, 7.15, 6.3, 5.5, 0.3, [[R("Backed by General Catalyst · Madrona", 9.5, FAINT, False, SANS, True)]])


# ============================================================
# SLIDE 8 — WHY AMIGO + NEXT STEPS (CTA)
# ============================================================
s = slide(); topbar(s)
brandmark(s)
eyebrow(s, "The Ask", "Why Amigo for Cleveland Diagnostics")

# left: three reasons
reasons = [
    ("Purpose-built for cancer screening", "Amigo already runs population-scale cancer focus areas — prostate risk-stratification is a natural extension, not a new build."),
    ("From test to standard-of-care", "Agents close every gap in the IsoPSA journey: find the patient, prompt the order, prove the outcome."),
    ("Proven, compliant, fast", "22.2x ROI, deployment in ~6 weeks, and SOC 2 / HIPAA / GDPR out of the box."),
]
cw, x0, y0 = 7.1, 0.55, 2.15
for i, (h, d) in enumerate(reasons):
    y = y0 + i * 1.35
    card(s, x0, y, cw, 1.2, accent=True)
    txt(s, x0 + 0.3, y + 0.2, 0.7, 0.7, [[R(f"{i+1}", 30, SKY, True, SERIF)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x0 + 1.05, y + 0.18, cw - 1.3, 0.4, [[R(h, 14, NAVY, True)]])
    txt(s, x0 + 1.05, y + 0.58, cw - 1.3, 0.55, [[R(d, 11, INK)]], line_spacing=1.02)

# right: CTA panel
cta = rect(s, 8.0, 2.15, 4.78, 4.05, radius=0.05)
_set_grad(cta, NAVY, SKY, 60); cta.line.fill.background(); _shadow(cta, blur=0.1, dist=0.05, alpha=80)
txt(s, 8.3, 2.6, 4.2, 0.4, [[R("READY TO TRANSFORM", 12, PALEBLUE, True)]], align=PP_ALIGN.CENTER)
txt(s, 8.2, 3.0, 4.4, 1.3,
    [[R("Prostate cancer", 22, WHITE, True)], [R("detection at", 22, WHITE, True)], [R("national scale", 22, WHITE, True)]],
    align=PP_ALIGN.CENTER, line_spacing=1.02, space_after=0)
btn = rect(s, 8.75, 4.75, 3.28, 0.6, radius=0.5)
btn.fill.solid(); btn.fill.fore_color.rgb = WHITE; btn.line.fill.background(); _shadow(btn, blur=0.05, dist=0.03, alpha=80)
txt(s, 8.75, 4.77, 3.28, 0.56, [[R("Let's Meet  →", 15, NAVY, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, 8.3, 5.55, 4.2, 0.35, [[R("elzoghby.elaf@amigo.ai", 12, WHITE, True)]], align=PP_ALIGN.CENTER)
txt(s, 8.3, 5.9, 4.2, 0.3, [[R("Predict.  Prevent.  Act.", 11, PALEBLUE, False, SANS, True)]], align=PP_ALIGN.CENTER)

footer(s, 8)

out = "Amigo_x_ClevelandDiagnostics.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
