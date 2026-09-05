"""
Victorious Christian Life - Professional Bilingual PowerPoint Generator
Generates ONE combined presentation with English + Malayalam (BSI) side by side.
Each sub-point = one slide. Keyboard Next triggers animations.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ── COLOR PALETTE ──
BG       = RGBColor(0x0A, 0x11, 0x28)
BG_CARD  = RGBColor(0x15, 0x20, 0x3B)
BORDER   = RGBColor(0x2E, 0x3E, 0x66)
GOLD     = RGBColor(0xF5, 0x9E, 0x0B)
GOLD_L   = RGBColor(0xFB, 0xBF, 0x24)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
MUTED    = RGBColor(0x94, 0xA3, 0xB8)
CYAN     = RGBColor(0x38, 0xBD, 0xF8)
# Bright readable ring colors: animation order (center→outer)
RING_CLR = [
    RGBColor(0xFF, 0xD7, 0x00),  # 1 center: bright gold
    RGBColor(0xFF, 0xAE, 0x42),  # 2 amber
    RGBColor(0xFF, 0x8C, 0x00),  # 3 dark orange
    RGBColor(0xFF, 0x63, 0x47),  # 4 tomato
    RGBColor(0xFF, 0x45, 0x00),  # 5 orange-red
    RGBColor(0xEF, 0x44, 0x44),  # 6 red
    RGBColor(0xDC, 0x26, 0x26),  # 7 outer: crimson
]
SW = Inches(13.33)
SH = Inches(7.5)
EN_FONT = "Calibri"
EN_HEAD = "Palatino Linotype"
ML_FONT = "Nirmala UI"

# ══════════════════════════════════════════════
#  ANIMATION (supports groups of shapes per click)
# ══════════════════════════════════════════════
def apply_click_animations(slide, groups):
    """groups: list of items; each item is a shape or list of shapes.
       Shapes in the same list animate TOGETHER on one click."""
    norm = []
    for g in groups:
        norm.append(g if isinstance(g, (list, tuple)) else [g])
    if not norm:
        return
    sp = slide.shapes._spTree
    sld = sp.getparent().getparent()
    timing = etree.SubElement(sld, qn('p:timing'))
    tnLst = etree.SubElement(timing, qn('p:tnLst'))
    pr = etree.SubElement(tnLst, qn('p:par'))
    cr = etree.SubElement(pr, qn('p:cTn'))
    cr.set('id','1'); cr.set('dur','indefinite')
    cr.set('restart','whenNotActive'); cr.set('nodeType','tmRoot')
    crl = etree.SubElement(cr, qn('p:childTnLst'))
    seq = etree.SubElement(crl, qn('p:seq'))
    seq.set('concurrent','1'); seq.set('nextAc','seek')
    cs = etree.SubElement(seq, qn('p:cTn'))
    cs.set('id','2'); cs.set('dur','indefinite'); cs.set('nodeType','mainSeq')
    scl = etree.SubElement(cs, qn('p:childTnLst'))
    for evt_name, lst_name in [('onPrev','p:prevCondLst'),('onNext','p:nextCondLst')]:
        cl = etree.SubElement(seq, qn(lst_name))
        cd = etree.SubElement(cl, qn('p:cond'))
        cd.set('evt', evt_name); cd.set('delay','0')
        te = etree.SubElement(cd, qn('p:tgtEl'))
        etree.SubElement(te, qn('p:sldTgt'))
    cid = 3
    for grp in norm:
        gp = etree.SubElement(scl, qn('p:par'))
        gc = etree.SubElement(gp, qn('p:cTn'))
        gc.set('id', str(cid)); cid += 1; gc.set('fill','hold')
        gs = etree.SubElement(gc, qn('p:stCondLst'))
        gsc = etree.SubElement(gs, qn('p:cond')); gsc.set('delay','indefinite')
        gcl = etree.SubElement(gc, qn('p:childTnLst'))
        mp = etree.SubElement(gcl, qn('p:par'))
        mc = etree.SubElement(mp, qn('p:cTn'))
        mc.set('id', str(cid)); cid += 1; mc.set('fill','hold')
        ms = etree.SubElement(mc, qn('p:stCondLst'))
        msc = etree.SubElement(ms, qn('p:cond')); msc.set('delay','0')
        mcl = etree.SubElement(mc, qn('p:childTnLst'))
        for si, shape in enumerate(grp):
            spid = str(shape.shape_id)
            nt = 'clickEffect' if si == 0 else 'withEffect'
            ap = etree.SubElement(mcl, qn('p:par'))
            ac = etree.SubElement(ap, qn('p:cTn'))
            ac.set('id', str(cid)); cid += 1
            ac.set('presetID','10'); ac.set('presetClass','entr')
            ac.set('presetSubtype','0'); ac.set('fill','hold')
            ac.set('grpId','0'); ac.set('nodeType', nt)
            acs = etree.SubElement(ac, qn('p:stCondLst'))
            acc = etree.SubElement(acs, qn('p:cond')); acc.set('delay','0')
            acl = etree.SubElement(ac, qn('p:childTnLst'))
            se = etree.SubElement(acl, qn('p:set'))
            sb = etree.SubElement(se, qn('p:cBhvr'))
            st = etree.SubElement(sb, qn('p:cTn'))
            st.set('id', str(cid)); cid += 1; st.set('dur','1'); st.set('fill','hold')
            ste = etree.SubElement(sb, qn('p:tgtEl'))
            stt = etree.SubElement(ste, qn('p:spTgt')); stt.set('spid', spid)
            sal = etree.SubElement(sb, qn('p:attrNameLst'))
            san = etree.SubElement(sal, qn('p:attrName')); san.text = 'style.visibility'
            sto = etree.SubElement(se, qn('p:to'))
            stv = etree.SubElement(sto, qn('p:strVal')); stv.set('val','visible')
            ae = etree.SubElement(acl, qn('p:animEffect'))
            ae.set('transition','in'); ae.set('filter','fade')
            ab = etree.SubElement(ae, qn('p:cBhvr'))
            at = etree.SubElement(ab, qn('p:cTn'))
            at.set('id', str(cid)); cid += 1; at.set('dur','500')
            ate = etree.SubElement(ab, qn('p:tgtEl'))
            att = etree.SubElement(ate, qn('p:spTgt')); att.set('spid', spid)

# ══════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════
def set_bg(s):
    s.background.fill.solid(); s.background.fill.fore_color.rgb = BG

def bar(slide, y, h, color):
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y, SW, h)
    b.fill.solid(); b.fill.fore_color.rgb = color; b.line.fill.background()
    return b

def txtbox(slide, txt, x, y, w, h, sz=Pt(18), bold=False, color=WHITE,
           align=PP_ALIGN.LEFT, italic=False, font=EN_FONT):
    t = slide.shapes.add_textbox(x, y, w, h)
    t.text_frame.word_wrap = True
    p = t.text_frame.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = txt
    r.font.size = sz; r.font.bold = bold; r.font.color.rgb = color
    r.font.italic = italic; r.font.name = font
    etree.SubElement(r._r.get_or_add_rPr(), qn('a:cs')).set('typeface', font)
    return t

def add_header(slide, en_title, ml_title):
    bar(slide, 0, Inches(0.06), GOLD)
    hb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.06), SW, Inches(1.05))
    hb.fill.solid(); hb.fill.fore_color.rgb = RGBColor(0x0E,0x18,0x36); hb.line.fill.background()
    bar(slide, Inches(1.11), Inches(0.02), BORDER)
    txtbox(slide, en_title, Inches(0.5), Inches(0.12), Inches(6.2), Inches(0.48),
           sz=Pt(24), bold=True, color=WHITE, font=EN_HEAD)
    txtbox(slide, ml_title, Inches(6.9), Inches(0.12), Inches(6.2), Inches(0.48),
           sz=Pt(22), bold=True, color=WHITE, font=ML_FONT)
    txtbox(slide, "ENGLISH", Inches(0.5), Inches(0.64), Inches(1.5), Inches(0.3),
           sz=Pt(10), bold=True, color=GOLD)
    txtbox(slide, "MALAYALAM", Inches(6.9), Inches(0.64), Inches(1.8), Inches(0.3),
           sz=Pt(10), bold=True, color=GOLD)

def make_card(slide, x, y, w, h, title, points, verse, font_body, font_head):
    c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    c.fill.solid(); c.fill.fore_color.rgb = BG_CARD
    c.line.color.rgb = BORDER; c.line.width = Pt(1.5)
    tf = c.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.15); tf.margin_bottom = Inches(0.15)
    pt = tf.paragraphs[0]; pt.space_after = Pt(6)
    rt = pt.add_run(); rt.text = title
    rt.font.size = Pt(19); rt.font.bold = True; rt.font.color.rgb = GOLD; rt.font.name = font_head
    etree.SubElement(rt._r.get_or_add_rPr(), qn('a:cs')).set('typeface', font_head)
    for item in points:
        p = tf.add_paragraph(); p.space_after = Pt(5)
        r = p.add_run(); r.text = "  " + item
        r.font.size = Pt(15); r.font.color.rgb = WHITE; r.font.name = font_body
        etree.SubElement(r._r.get_or_add_rPr(), qn('a:cs')).set('typeface', font_body)
    if verse:
        pv = tf.add_paragraph(); pv.space_before = Pt(6)
        rv = pv.add_run(); rv.text = verse
        rv.font.size = Pt(13); rv.font.italic = True; rv.font.color.rgb = CYAN; rv.font.name = font_body
        etree.SubElement(rv._r.get_or_add_rPr(), qn('a:cs')).set('typeface', font_body)
    return c

# ══════════════════════════════════════════════
#  BILINGUAL CONTENT SLIDE
# ══════════════════════════════════════════════
def make_content_slide(prs, en_topic, ml_topic, sp):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide)
    add_header(slide, en_topic, ml_topic)
    bar(slide, SH - Inches(0.06), Inches(0.06), GOLD)
    cw = Inches(6.16); cy = Inches(1.25); ch = Inches(5.85)
    en = make_card(slide, Inches(0.25), cy, cw, ch,
                   sp['en_title'], sp['en_points'], sp.get('en_verse',''), EN_FONT, EN_HEAD)
    dv = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.56), Inches(1.45), Inches(0.02), Inches(5.4))
    dv.fill.solid(); dv.fill.fore_color.rgb = BORDER; dv.line.fill.background()
    ml = make_card(slide, Inches(6.72), cy, cw, ch,
                   sp['ml_title'], sp['ml_points'], sp.get('ml_verse',''), ML_FONT, ML_FONT)
    apply_click_animations(slide, [[en, ml]])

# ══════════════════════════════════════════════
#  TITLE SLIDE
# ══════════════════════════════════════════════
def make_title(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    bar(s, 0, Inches(0.1), GOLD)
    p = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.2), Inches(11.73), Inches(5.0))
    p.fill.solid(); p.fill.fore_color.rgb = BG_CARD; p.line.color.rgb = GOLD; p.line.width = Pt(2)
    cv = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.46), Inches(1.5), Inches(0.4), Inches(1.0))
    ch = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.06), Inches(1.8), Inches(1.2), Inches(0.4))
    for x in (cv,ch): x.fill.solid(); x.fill.fore_color.rgb = GOLD; x.line.fill.background()
    txtbox(s, "VICTORIOUS CHRISTIAN LIFE", Inches(1), Inches(2.8), Inches(11.33), Inches(1.0),
           sz=Pt(46), bold=True, color=GOLD, align=PP_ALIGN.CENTER, font=EN_HEAD)
    txtbox(s, "\u0d35\u0d3f\u0d1c\u0d2f\u0d15\u0d30\u0d2e\u0d3e\u0d2f \u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d40\u0d2f \u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02",
           Inches(1), Inches(3.85), Inches(11.33), Inches(0.8),
           sz=Pt(36), bold=True, color=GOLD_L, align=PP_ALIGN.CENTER, font=ML_FONT)
    txtbox(s, "ICPF Camp at Trivandrum 2026  |  ICPF \u0d15\u0d4d\u0d2f\u0d3e\u0d2e\u0d4d\u0d2a\u0d4d, \u0d24\u0d3f\u0d30\u0d41\u0d35\u0d28\u0d28\u0d4d\u0d24\u0d2a\u0d41\u0d30\u0d02 2026",
           Inches(1), Inches(4.85), Inches(11.33), Inches(0.5),
           sz=Pt(20), color=CYAN, align=PP_ALIGN.CENTER)
    bar(s, SH - Inches(0.1), Inches(0.1), GOLD)

# ══════════════════════════════════════════════
#  BLANK INTRODUCTION SLIDE
# ══════════════════════════════════════════════
def make_intro(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    bar(s, 0, Inches(0.06), GOLD)
    bar(s, SH - Inches(0.06), Inches(0.06), GOLD)
    txtbox(s, "INTRODUCTION", Inches(1), Inches(3.0), Inches(11.33), Inches(1.0),
           sz=Pt(50), bold=True, color=GOLD, align=PP_ALIGN.CENTER, font=EN_HEAD)
    txtbox(s, "ആമുഖം", Inches(1), Inches(4.0), Inches(11.33), Inches(0.8),
           sz=Pt(40), bold=True, color=GOLD_L, align=PP_ALIGN.CENTER, font=ML_FONT)

# ══════════════════════════════════════════════
#  SECTION DIVIDER
# ══════════════════════════════════════════════
def make_divider(prs, num, en_t, ml_t):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    lb = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.25), SH)
    lb.fill.solid(); lb.fill.fore_color.rgb = GOLD; lb.line.fill.background()
    ci = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.0), Inches(2.4), Inches(1.6), Inches(1.6))
    ci.fill.solid(); ci.fill.fore_color.rgb = GOLD; ci.line.fill.background()
    tf = ci.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num; r.font.size = Pt(50); r.font.bold = True; r.font.color.rgb = BG
    etree.SubElement(r._r.get_or_add_rPr(), qn('a:cs')).set('typeface', EN_FONT)
    txtbox(s, en_t, Inches(3.0), Inches(2.2), Inches(9.5), Inches(0.9),
           sz=Pt(42), bold=True, color=WHITE, font=EN_HEAD)
    txtbox(s, ml_t, Inches(3.0), Inches(3.2), Inches(9.5), Inches(0.9),
           sz=Pt(34), bold=True, color=GOLD_L, font=ML_FONT)
    bar(s, SH - Inches(0.08), Inches(0.08), GOLD)

# ══════════════════════════════════════════════
#  7 BINDINGS SLIDE (Eph 2:2-3)  Center → Outer
# ══════════════════════════════════════════════
# Labels in ANIMATION ORDER (center first → outer last)
BIND_EN = [
    "Dead in Trespasses\n& Sins",
    "Course of\nthis World",
    "Prince of the\nPower of the Air",
    "Spirit of\nDisobedience",
    "Passions / Lusts\nof the Flesh",
    "Desires of Body\n& Mind",
    "Children of Wrath\nby Nature",
]
BIND_ML = [
    "\u0d05\u0d24\u0d3f\u0d15\u0d4d\u0d30\u0d2e\u0d19\u0d4d\u0d19\u0d33\u0d3f\u0d32\u0d41\u0d02\n\u0d2a\u0d3e\u0d2a\u0d19\u0d4d\u0d19\u0d33\u0d3f\u0d32\u0d41\u0d02 \u0d2e\u0d30\u0d3f\u0d1a\u0d4d\u0d1a\u0d35\u0d30\u0d4d\u200d",
    "\u0d08 \u0d32\u0d4b\u0d15\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46\n\u0d35\u0d34\u0d3f\u0d2f\u0d3f\u0d32\u0d4d\u200d \u0d28\u0d1f\u0d28\u0d4d\u0d28\u0d41",
    "\u0d06\u0d15\u0d3e\u0d36\u0d3e\u0d27\u0d3f\u0d2a\u0d24\u0d3f\u0d2f\u0d3e\u0d2f\n\u0d05\u0d27\u0d3f\u0d15\u0d3e\u0d30\u0d3f",
    "\u0d05\u0d28\u0d41\u0d38\u0d30\u0d23\u0d15\u0d4d\u0d15\u0d47\u0d1f\u0d3f\u0d28\u0d4d\u0d31\u0d46\n\u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d",
    "\u0d1c\u0d21\u0d2e\u0d4b\u0d39\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d",
    "\u0d36\u0d30\u0d40\u0d30\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46\u0d2f\u0d41\u0d02\n\u0d2e\u0d28\u0d38\u0d4d\u0d38\u0d3f\u0d28\u0d4d\u0d31\u0d46\u0d2f\u0d41\u0d02 \u0d06\u0d17\u0d4d\u0d30\u0d39\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d",
    "\u0d38\u0d4d\u0d35\u0d2d\u0d3e\u0d35\u0d2e\u0d3e\u0d2f\u0d3f\n\u0d15\u0d4d\u0d30\u0d4b\u0d27\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2e\u0d15\u0d4d\u0d15\u0d33\u0d4d\u200d",
]

def make_bindings_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    add_header(s, "4. Love \u2014 7 Bindings of Sin",
               "4. \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02 \u2014 \u0d2a\u0d3e\u0d2a\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 7 \u0d2c\u0d28\u0d4d\u0d27\u0d28\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d")
    bar(s, SH - Inches(0.06), Inches(0.06), GOLD)
    txtbox(s, "Ephesians 2:2-3 | \u0d0e\u0d2b\u0d46\u0d38\u0d4d\u0d2f\u0d30\u0d4d\u200d 2:2-3",
           Inches(4.5), Inches(0.64), Inches(4.5), Inches(0.3),
           sz=Pt(11), bold=True, color=CYAN)

    cx = Inches(6.665); cy_c = Inches(4.15)
    # Ring sizes: drawing order outer(idx0)→inner(idx6)
    rw = [Inches(10.0), Inches(8.4), Inches(7.0), Inches(5.6), Inches(4.2), Inches(2.8), Inches(1.6)]
    rh = [Inches(5.0),  Inches(4.2), Inches(3.4), Inches(2.6), Inches(2.0), Inches(1.4), Inches(0.9)]

    rings = []; lbl_en = []; lbl_ml = []
    for di in range(7):  # di=0 outermost, di=6 innermost
        ai = 6 - di  # animation index: inner(0)→outer(6)
        clr = RING_CLR[ai]
        rx = cx - rw[di] / 2; ry = cy_c - rh[di] / 2
        ov = s.shapes.add_shape(MSO_SHAPE.OVAL, rx, ry, rw[di], rh[di])
        ov.fill.background()
        ov.line.color.rgb = clr; ov.line.width = Pt(3.5)
        rings.append(ov)
        # EN label (left)
        ly = Inches(1.25) + Inches(0.72) * ai
        le = txtbox(s, BIND_EN[ai], Inches(0.15), ly, Inches(2.9), Inches(0.6),
                    sz=Pt(12), bold=True, color=clr, font=EN_FONT)
        lbl_en.append(le)
        # ML label (right)
        lm = txtbox(s, BIND_ML[ai], Inches(10.4), ly, Inches(2.9), Inches(0.6),
                    sz=Pt(12), bold=True, color=clr, font=ML_FONT)
        lbl_ml.append(lm)

    # Animation: center(di=6) first → outer(di=0) last
    anim_groups = []
    for di in range(6, -1, -1):
        anim_groups.append([rings[di], lbl_en[di], lbl_ml[di]])

    # BREAKING: golden cross + BUT GOD
    cr_v = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx-Inches(0.3), cy_c-Inches(2.0), Inches(0.6), Inches(4.0))
    cr_h = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx-Inches(2.0), cy_c-Inches(0.3), Inches(4.0), Inches(0.6))
    for x in (cr_v, cr_h):
        x.fill.solid(); x.fill.fore_color.rgb = GOLD; x.line.fill.background()
    ov_p = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(5.4), Inches(8.93), Inches(1.7))
    ov_p.fill.solid(); ov_p.fill.fore_color.rgb = BG_CARD; ov_p.line.color.rgb = GOLD; ov_p.line.width = Pt(2)
    be = txtbox(s, 'BUT GOD, being rich in mercy, because of His great love... (Eph 2:4)',
                Inches(2.4), Inches(5.5), Inches(4.2), Inches(1.5),
                sz=Pt(15), bold=True, color=GOLD, font=EN_HEAD)
    bm = txtbox(s, '\u0d0e\u0d28\u0d4d\u0d28\u0d3e\u0d32\u0d4d\u200d \u0d26\u0d48\u0d35\u0d02 \u0d15\u0d30\u0d41\u0d23\u0d3e\u0d38\u0d2e\u0d4d\u0d2a\u0d28\u0d4d\u0d28\u0d28\u0d3e\u0d15\u0d2f\u0d3e\u0d32\u0d4d\u200d, \u0d28\u0d2e\u0d4d\u0d2e\u0d46 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d1a\u0d4d\u0d1a \u0d2e\u0d39\u0d3e\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02... (\u0d0e\u0d2b\u0d46 2:4)',
                Inches(6.8), Inches(5.5), Inches(4.2), Inches(1.5),
                sz=Pt(15), bold=True, color=GOLD, font=ML_FONT)
    anim_groups.append([cr_v, cr_h, ov_p, be, bm])
    apply_click_animations(s, anim_groups)

# ══════════════════════════════════════════════
#  CLOSING SLIDE (7 pillars - aligned columns)
# ══════════════════════════════════════════════
def make_closing(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    bar(s, 0, Inches(0.08), GOLD)
    txtbox(s, "LIVE VICTORIOUSLY IN CHRIST!", Inches(0.3), Inches(0.2), Inches(12.7), Inches(0.7),
           sz=Pt(32), bold=True, color=GOLD, align=PP_ALIGN.CENTER, font=EN_HEAD)
    txtbox(s, "\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d32\u0d4d\u200d \u0d35\u0d3f\u0d1c\u0d2f\u0d15\u0d30\u0d2e\u0d3e\u0d2f\u0d3f \u0d1c\u0d40\u0d35\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15!",
           Inches(0.3), Inches(0.8), Inches(12.7), Inches(0.6),
           sz=Pt(24), bold=True, color=GOLD_L, align=PP_ALIGN.CENTER, font=ML_FONT)
    pillars = [
        ("1. Relationship\nwith Christ", "1. \u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d41\u0d2e\u0d3e\u0d2f\u0d41\u0d33\u0d4d\u0d33\n\u0d2c\u0d28\u0d4d\u0d27\u0d02"),
        ("2. Bible\nReading", "2. \u0d2c\u0d48\u0d2c\u0d3f\u0d33\u0d4d\u200d\n\u0d35\u0d3e\u0d2f\u0d28"),
        ("3. Prayer\nLife", "3. \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d3e\n\u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02"),
        ("4. Love", "4. \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02"),
        ("5. Receive\nHoly Spirit", "5. പരിശുദ്ധാത്മാവിനെ\nപ്രാപിക്കുക"),
        ("6. Holy\nLife", "6. \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\n\u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02"),
        ("7. Eternal\nHope", "7. \u0d28\u0d3f\u0d24\u0d4d\u0d2f\n\u0d2a\u0d4d\u0d30\u0d24\u0d4d\u0d2f\u0d3e\u0d36"),
    ]
    # Evenly spaced grid: 4 top + 3 bottom (centered)
    margin = Inches(0.3)
    gap = Inches(0.2)
    usable = SW - 2 * margin  # ~12.73"
    cw = (usable - 3 * gap) / 4  # 4 boxes per row
    ch = Inches(1.9)
    # Top row: 4 pillars
    for i in range(4):
        cx = margin + i * (cw + gap)
        _pillar_box(s, cx, Inches(1.6), cw, ch, pillars[i][0], pillars[i][1])
    # Bottom row: 3 pillars centered
    bot_total = 3 * cw + 2 * gap
    bot_x0 = (SW - bot_total) / 2
    for i in range(3):
        cx = bot_x0 + i * (cw + gap)
        _pillar_box(s, cx, Inches(3.7), cw, ch, pillars[4+i][0], pillars[4+i][1])
    # Key verse
    txtbox(s, '"I can do all things through Christ who strengthens me." \u2014 Philippians 4:13',
           Inches(0.3), Inches(5.8), Inches(6.2), Inches(0.7),
           sz=Pt(16), italic=True, color=CYAN, font=EN_FONT)
    txtbox(s, '"\u0d0e\u0d28\u0d4d\u0d28\u0d46 \u0d36\u0d15\u0d4d\u0d24\u0d28\u0d3e\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d35\u0d28\u0d4d\u200d \u0d2e\u0d41\u0d16\u0d3e\u0d28\u0d4d\u0d24\u0d30\u0d02 \u0d1e\u0d3e\u0d28\u0d4d\u200d \u0d38\u0d15\u0d32\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d28\u0d41\u0d02 \u0d2e\u0d24\u0d3f\u0d2f\u0d3e\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41." \u2014 \u0d2b\u0d3f\u0d32\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d2f\u0d30\u0d4d\u200d 4:13',
           Inches(6.8), Inches(5.8), Inches(6.2), Inches(0.7),
           sz=Pt(15), italic=True, color=CYAN, font=ML_FONT)
    bar(s, SH - Inches(0.08), Inches(0.08), GOLD)

def _pillar_box(s, x, y, w, h, en, ml):
    bx = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bx.fill.solid(); bx.fill.fore_color.rgb = BG_CARD
    bx.line.color.rgb = GOLD; bx.line.width = Pt(1.5)
    tf = bx.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.12); tf.margin_top = Inches(0.1)
    p1 = tf.paragraphs[0]; p1.space_after = Pt(4)
    r1 = p1.add_run(); r1.text = en
    r1.font.size = Pt(14); r1.font.bold = True; r1.font.color.rgb = WHITE; r1.font.name = EN_FONT
    etree.SubElement(r1._r.get_or_add_rPr(), qn('a:cs')).set('typeface', EN_FONT)
    p2 = tf.add_paragraph()
    r2 = p2.add_run(); r2.text = ml
    r2.font.size = Pt(13); r2.font.color.rgb = GOLD_L; r2.font.name = ML_FONT
    etree.SubElement(r2._r.get_or_add_rPr(), qn('a:cs')).set('typeface', ML_FONT)

# ══════════════════════════════════════════════
#  BIBLE ACRONYM SLIDE (Before Prayer)
# ══════════════════════════════════════════════
def make_bible_acronym_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    bar(s, 0, Inches(0.06), GOLD)
    bar(s, SH - Inches(0.06), Inches(0.06), GOLD)
    
    # Title
    txtbox(s, "WHAT IS THE BIBLE?", Inches(1), Inches(0.8), Inches(11.33), Inches(1.0),
           sz=Pt(40), bold=True, color=CYAN, align=PP_ALIGN.CENTER, font=EN_HEAD)
           
    # Acronym
    start_y = Inches(2.2)
    y_step = Inches(0.9)
    letters = ["B", "I", "B", "L", "E"]
    words = ["Basic", "Instructions", "Before", "Leaving", "Earth"]
    
    for i in range(5):
        y = start_y + i * y_step
        txtbox(s, letters[i], Inches(4.0), y, Inches(1.0), Inches(0.9),
               sz=Pt(48), bold=True, color=GOLD, align=PP_ALIGN.RIGHT, font=EN_HEAD)
        txtbox(s, f" \u2014  {words[i]}", Inches(5.0), y, Inches(6.0), Inches(0.9),
               sz=Pt(48), bold=True, color=WHITE, align=PP_ALIGN.LEFT, font=EN_HEAD)

# ══════════════════════════════════════════════
#  CONTENT DATA  (BSI Malayalam Bible text)
# ══════════════════════════════════════════════
TOPICS = [
    # ─── 1. RELATIONSHIP WITH CHRIST ───
    {
        'num': '01',
        'en': 'Relationship with Christ',
        'ml': '\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d41\u0d2e\u0d3e\u0d2f\u0d41\u0d33\u0d4d\u0d33 \u0d2c\u0d28\u0d4d\u0d27\u0d02',
        'subs': [
            {
                'en_title': 'Abiding in Christ',
                'ml_title': '\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d32\u0d4d\u200d \u0d35\u0d38\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15',
                'en_points': [
                    '>> I am the vine; you are the branches.',
                    '>> Whoever abides in Me bears much fruit.',
                    '>> Apart from Christ, we can do nothing.',
                    '>> If you abide in Me and My words abide in you, ask whatever you wish.'
                ],
                'ml_points': [
                    '>> \u0d1e\u0d3e\u0d28\u0d4d\u200d \u0d2e\u0d41\u0d28\u0d4d\u0d24\u0d3f\u0d30\u0d3f\u0d35\u0d33\u0d4d\u0d33\u0d3f\u0d2f\u0d41\u0d02 \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d15\u0d4a\u0d2e\u0d4d\u0d2a\u0d41\u0d15\u0d33\u0d41\u0d02 \u0d06\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d0e\u0d28\u0d4d\u0d28\u0d3f\u0d32\u0d4d\u200d \u0d35\u0d38\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d35\u0d28\u0d4d\u200d \u0d35\u0d33\u0d30\u0d46 \u0d2b\u0d32\u0d02 \u0d15\u0d3e\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d41\u0d02.',
                    '>> \u0d0e\u0d28\u0d4d\u0d28\u0d46\u0d15\u0d4d\u0d15\u0d42\u0d1f\u0d3e\u0d24\u0d46 \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d\u0d15\u0d4d\u0d15\u0d4d \u0d12\u0d28\u0d4d\u0d28\u0d41\u0d02 \u0d1a\u0d46\u0d2f\u0d4d\u0d2f\u0d41\u0d35\u0d3e\u0d28\u0d4d\u200d \u0d15\u0d34\u0d3f\u0d2f\u0d41\u0d15\u0d2f\u0d3f\u0d32\u0d4d\u0d32.',
                    '>> \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d0e\u0d28\u0d4d\u0d28\u0d3f\u0d32\u0d41\u0d02 \u0d0e\u0d28\u0d4d\u0d31\u0d46 \u0d35\u0d1a\u0d28\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d3f\u0d32\u0d41\u0d02 \u0d35\u0d38\u0d3f\u0d1a\u0d4d\u0d1a\u0d3e\u0d32\u0d4d\u200d \u0d07\u0d1a\u0d4d\u0d1b\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d41 \u0d1a\u0d4b\u0d26\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d28\u0d4d\u200d.'
                ],
                'en_verse': 'John 15:5, 7-8',
                'ml_verse': '\u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 15:5, 7-8',
            },
            {
                'en_title': 'The Mind of Christ (Phil 2:5)',
                'ml_title': '\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2d\u0d3e\u0d35\u0d02 (\u0d2b\u0d3f\u0d32\u0d3f 2:5)',
                'en_points': [
                    '>> Have this mind among yourselves, which is yours in Christ Jesus.',
                    '>> Cultivate the humble, obedient attitude of Jesus.',
                    '>> Walk in selfless servanthood and love for others.'
                ],
                'ml_points': [
                    '>> \u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d32\u0d41\u0d33\u0d4d\u0d33 \u0d08 \u0d2d\u0d3e\u0d35\u0d02 \u0d24\u0d28\u0d4d\u0d28\u0d46 \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d3f\u0d32\u0d41\u0d02 \u0d09\u0d23\u0d4d\u0d1f\u0d3e\u0d2f\u0d3f\u0d30\u0d3f\u0d15\u0d4d\u0d15\u0d1f\u0d4d\u0d1f\u0d46.',
                    '>> \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d24\u0d3e\u0d34\u0d4d\u0d2e\u0d2f\u0d41\u0d02 \u0d35\u0d3f\u0d28\u0d2f\u0d35\u0d41\u0d02 \u0d38\u0d4d\u0d35\u0d28\u0d4d\u0d24\u0d2e\u0d3e\u0d15\u0d4d\u0d15\u0d41\u0d15.',
                    '>> \u0d24\u0d4d\u0d2f\u0d3e\u0d17\u0d24\u0d4d\u0d24\u0d4b\u0d1f\u0d46 \u0d2e\u0d31\u0d4d\u0d31\u0d41\u0d33\u0d4d\u0d33\u0d35\u0d30\u0d46 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15\u0d2f\u0d41\u0d02 \u0d38\u0d47\u0d35\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15\u0d2f\u0d41\u0d02 \u0d1a\u0d46\u0d2f\u0d4d\u0d2f\u0d41\u0d15.'
                ],
                'en_verse': 'Philippians 2:5',
                'ml_verse': '\u0d2b\u0d3f\u0d32\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d2f\u0d30\u0d4d\u200d 2:5',
            },
            {
                'en_title': "Bearing Fruit for God's Glory (John 15:7-8)",
                'ml_title': '\u0d2a\u0d3f\u0d24\u0d3e\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2e\u0d39\u0d24\u0d4d\u0d24\u0d4d\u0d35\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d3e\u0d2f\u0d3f \u0d2b\u0d32\u0d02 \u0d15\u0d3e\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d41\u0d15 (\u0d2f\u0d4b\u0d39 15:7-8)',
                'en_points': [
                    '>> By this my Father is glorified, that you bear much fruit.',
                    '>> Fruitfulness is the evidence of authentic discipleship.',
                    '>> Our lives bring honor and glory to the Father.'
                ],
                'ml_points': [
                    '>> \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d35\u0d33\u0d30\u0d46 \u0d2b\u0d32\u0d02 \u0d15\u0d3e\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d3f\u0d28\u0d3e\u0d32\u0d4d\u200d \u0d0e\u0d28\u0d4d\u0d31\u0d46 \u0d2a\u0d3f\u0d24\u0d3e\u0d35\u0d4d \u0d2e\u0d39\u0d24\u0d4d\u0d24\u0d4d\u0d35\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d0f\u0d31\u0d46 \u0d2b\u0d32\u0d02 \u0d15\u0d3e\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d3e\u0d23\u0d4d \u0d2f\u0d25\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25 \u0d36\u0d3f\u0d37\u0d4d\u0d2f\u0d24\u0d4d\u0d35\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d24\u0d46\u0d33\u0d3f\u0d35\u0d4d.',
                    '>> \u0d28\u0d2e\u0d4d\u0d2e\u0d41\u0d1f\u0d46 \u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02 \u0d35\u0d34\u0d3f \u0d38\u0d4d\u0d35\u0d30\u0d4d\u200d\u0d17\u0d4d\u0d17\u0d40\u0d2f \u0d2a\u0d3f\u0d24\u0d3e\u0d35\u0d4d \u0d2e\u0d39\u0d24\u0d4d\u0d24\u0d4d\u0d35\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d28\u0d4d\u0d28\u0d41.'
                ],
                'en_verse': 'John 15:7-8',
                'ml_verse': '\u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 15:7-8',
            },
        ]
    },
    # ─── 2. BIBLE READING ───
    {
        'num': '02',
        'en': 'Bible Reading',
        'ml': '\u0d2c\u0d48\u0d2c\u0d3f\u0d33\u0d4d\u200d \u0d35\u0d3e\u0d2f\u0d28',
        'subs': [
            {
                'en_title': 'Word of God Has Creating Power',
                'ml_title': '\u0d26\u0d48\u0d35\u0d35\u0d1a\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d \u0d38\u0d43\u0d37\u0d4d\u0d1f\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28 \u0d36\u0d15\u0d4d\u0d24\u0d3f\u0d2f\u0d41\u0d23\u0d4d\u0d1f\u0d4d',
                'en_points': [
                    '>> In the beginning, God created the heavens and the earth by His spoken Word.',
                    '>> God SAID "Let there be light" \u2014 and there was light.',
                    '>> The Word of God carries creative, life-giving power.'
                ],
                'ml_points': [
                    '>> \u0d06\u0d26\u0d3f\u0d2f\u0d3f\u0d32\u0d4d\u200d \u0d26\u0d48\u0d35\u0d02 \u0d06\u0d15\u0d3e\u0d36\u0d35\u0d41\u0d02 \u0d2d\u0d42\u0d2e\u0d3f\u0d2f\u0d41\u0d02 \u0d38\u0d43\u0d37\u0d4d\u0d1f\u0d3f\u0d1a\u0d4d\u0d1a\u0d41.',
                    '>> \u0d26\u0d48\u0d35\u0d02 "\u0d35\u0d46\u0d33\u0d3f\u0d1a\u0d4d\u0d1a\u0d02 \u0d09\u0d23\u0d4d\u0d1f\u0d3e\u0d15\u0d1f\u0d4d\u0d1f\u0d46" \u0d0e\u0d28\u0d4d\u0d28\u0d4d \u0d05\u0d30\u0d41\u0d33\u0d3f\u0d1a\u0d4d\u0d1a\u0d46\u0d2f\u0d4d\u0d24\u0d41; \u0d35\u0d46\u0d33\u0d3f\u0d1a\u0d4d\u0d1a\u0d02 \u0d09\u0d23\u0d4d\u0d1f\u0d3e\u0d2f\u0d3f.',
                    '>> \u0d26\u0d48\u0d35\u0d35\u0d1a\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d \u0d38\u0d43\u0d37\u0d4d\u0d1f\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d41\u0d02 \u0d1c\u0d40\u0d35\u0d28\u0d4d\u200d \u0d28\u0d32\u0d4d\u200d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d41\u0d2e\u0d3e\u0d2f \u0d36\u0d15\u0d4d\u0d24\u0d3f\u0d2f\u0d41\u0d23\u0d4d\u0d1f\u0d4d.'
                ],
                'en_verse': 'Genesis 1:1 \u2014 "In the beginning God created the heavens and the earth."',
                'ml_verse': '\u0d09\u0d32\u0d4d\u200d\u0d2a\u0d24\u0d4d\u0d24\u0d3f 1:1 \u2014 "\u0d06\u0d26\u0d3f\u0d2f\u0d3f\u0d32\u0d4d\u200d \u0d26\u0d48\u0d35\u0d02 \u0d06\u0d15\u0d3e\u0d36\u0d35\u0d41\u0d02 \u0d2d\u0d42\u0d2e\u0d3f\u0d2f\u0d41\u0d02 \u0d38\u0d43\u0d37\u0d4d\u0d1f\u0d3f\u0d1a\u0d4d\u0d1a\u0d41."',
            },
            {
                'en_title': 'Word of God is Jesus Christ',
                'ml_title': '\u0d26\u0d48\u0d35\u0d35\u0d1a\u0d28\u0d02 \u0d2f\u0d47\u0d36\u0d41\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3e\u0d23\u0d4d',
                'en_points': [
                    '>> In the beginning was the Word, and the Word was with God, and the Word was God.',
                    '>> Jesus Christ IS the living Word of God made flesh.',
                    '>> Scripture reveals the person and heart of Jesus to us daily.'
                ],
                'ml_points': [
                    '>> \u0d06\u0d26\u0d3f\u0d2f\u0d3f\u0d32\u0d4d\u200d \u0d35\u0d1a\u0d28\u0d02 \u0d09\u0d23\u0d4d\u0d1f\u0d3e\u0d2f\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41; \u0d35\u0d1a\u0d28\u0d02 \u0d26\u0d48\u0d35\u0d24\u0d4d\u0d24\u0d4b\u0d1f\u0d41\u0d15\u0d42\u0d1f\u0d46 \u0d06\u0d2f\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41; \u0d35\u0d1a\u0d28\u0d02 \u0d26\u0d48\u0d35\u0d02 \u0d06\u0d2f\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d2f\u0d47\u0d36\u0d41\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41 \u0d1c\u0d21\u0d2e\u0d3e\u0d2f\u0d3f\u0d24\u0d4d\u0d24\u0d40\u0d30\u0d4d\u200d\u0d28\u0d4d\u0d28 \u0d1c\u0d40\u0d35\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d26\u0d48\u0d35\u0d35\u0d1a\u0d28\u0d2e\u0d3e\u0d23\u0d4d.',
                    '>> \u0d24\u0d3f\u0d30\u0d41\u0d35\u0d46\u0d34\u0d41\u0d24\u0d4d\u0d24\u0d41\u0d15\u0d33\u0d4d\u200d \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d39\u0d43\u0d26\u0d2f\u0d35\u0d41\u0d02 \u0d38\u0d24\u0d4d\u0d2f\u0d35\u0d41\u0d02 \u0d28\u0d2e\u0d41\u0d15\u0d4d\u0d15\u0d4d \u0d35\u0d46\u0d33\u0d3f\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d24\u0d4d\u0d24\u0d41\u0d28\u0d4d\u0d28\u0d41.'
                ],
                'en_verse': 'John 1:1 \u2014 "The Word was with God, and the Word was God."',
                'ml_verse': '\u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 1:1 \u2014 "\u0d35\u0d1a\u0d28\u0d02 \u0d26\u0d48\u0d35\u0d24\u0d4d\u0d24\u0d4b\u0d1f\u0d41\u0d15\u0d42\u0d1f\u0d46 \u0d06\u0d2f\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41; \u0d35\u0d1a\u0d28\u0d02 \u0d26\u0d48\u0d35\u0d02 \u0d06\u0d2f\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41."',
            },
            {
                'en_title': 'Good News of the Kingdom of God',
                'ml_title': 'ദൈവരാജ്യത്തിന്റെ സുവിശേഷം',
                'en_points': [
                    '>> The beginning of the Gospel of Jesus Christ, the Son of God.',
                    '>> Without understanding, the evil one snatches away the sown Word.',
                    '>> Guard, meditate upon, and obey the Truth in your heart.'
                ],
                'ml_points': [
                    '>> \u0d26\u0d48\u0d35\u0d2a\u0d41\u0d24\u0d4d\u0d30\u0d28\u0d3e\u0d2f \u0d2f\u0d47\u0d36\u0d41\u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d38\u0d41\u0d35\u0d3f\u0d36\u0d47\u0d37\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d06\u0d30\u0d02\u0d2d\u0d02.',
                    '>> \u0d17\u0d4d\u0d30\u0d39\u0d3f\u0d15\u0d4d\u0d15\u0d3e\u0d24\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d3e\u0d32\u0d4d\u200d \u0d26\u0d41\u0d37\u0d4d\u0d1f\u0d28\u0d4d\u200d \u0d35\u0d28\u0d4d\u0d28\u0d4d \u0d39\u0d43\u0d26\u0d2f\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d46 \u0d35\u0d1a\u0d28\u0d02 \u0d31\u0d3e\u0d1e\u0d4d\u0d1a\u0d3f \u0d0e\u0d1f\u0d41\u0d15\u0d4d\u0d15\u0d41\u0d02.',
                    '>> \u0d35\u0d1a\u0d28\u0d02 \u0d36\u0d4d\u0d30\u0d26\u0d4d\u0d27\u0d2f\u0d4b\u0d1f\u0d46 \u0d15\u0d47\u0d1f\u0d4d\u0d1f\u0d4d \u0d39\u0d43\u0d26\u0d2f\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d4d\u200d \u0d15\u0d3e\u0d24\u0d4d\u0d24\u0d41\u0d38\u0d42\u0d15\u0d4d\u0d37\u0d3f\u0d15\u0d4d\u0d15\u0d23\u0d02.'
                ],
                'en_verse': 'Mark 1:1 & Matthew 13:19',
                'ml_verse': '\u0d2e\u0d30\u0d4d\u200d\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 1:1 & \u0d2e\u0d24\u0d4d\u0d24\u0d3e\u0d2f\u0d3f 13:19',
            },
            {
                'en_title': 'Written Word of God is the Bible',
                'ml_title': 'എഴുതപ്പെട്ട ദൈവവചനമാണ് ബൈബിൾ',
                'en_points': [
                    '>> All Scripture is breathed out by God (God-breathed / inspired).',
                    '>> Profitable for teaching, reproof, correction, and training in righteousness.',
                    '>> The Bible is the inspired, written, authoritative Word of God.'
                ],
                'ml_points': [
                    '>> \u0d0e\u0d32\u0d4d\u0d32\u0d3e \u0d24\u0d3f\u0d30\u0d41\u0d35\u0d46\u0d34\u0d41\u0d24\u0d4d\u0d24\u0d41\u0d02 \u0d26\u0d48\u0d35\u0d36\u0d4d\u0d35\u0d3e\u0d38\u0d40\u0d2f\u0d2e\u0d3e\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d09\u0d2a\u0d26\u0d47\u0d36\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d41\u0d02 \u0d36\u0d3e\u0d38\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d41\u0d02 \u0d24\u0d3f\u0d30\u0d41\u0d24\u0d4d\u0d24\u0d32\u0d3f\u0d28\u0d41\u0d02 \u0d28\u0d40\u0d24\u0d3f\u0d2f\u0d3f\u0d32\u0d46 \u0d2a\u0d30\u0d3f\u0d36\u0d40\u0d32\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d41\u0d02 \u0d2a\u0d4d\u0d30\u0d2f\u0d4b\u0d1c\u0d28\u0d2e\u0d41\u0d33\u0d4d\u0d33\u0d24\u0d4d.',
                    '>> \u0d38\u0d24\u0d4d\u0d2f\u0d35\u0d47\u0d26\u0d2a\u0d41\u0d38\u0d4d\u0d24\u0d15\u0d02 \u0d26\u0d48\u0d35\u0d24\u0d4d\u0d24\u0d3e\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d36\u0d4d\u0d35\u0d38\u0d3f\u0d15\u0d4d\u0d15\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d4d\u0d1f \u0d05\u0d27\u0d3f\u0d15\u0d3e\u0d30\u0d2a\u0d42\u0d30\u0d4d\u200d\u0d23\u0d4d\u0d23\u0d2e\u0d3e\u0d2f \u0d35\u0d1a\u0d28\u0d2e\u0d3e\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.'
                ],
                'en_verse': '2 Timothy 3:16-17',
                'ml_verse': '2 \u0d24\u0d3f\u0d2e\u0d4a\u0d25\u0d46\u0d2f\u0d4a\u0d38\u0d4d 3:16-17',
            },
        ]
    },
    # ─── 3. PRAYER ───
    {
        'num': '03',
        'en': 'Prayer Life',
        'ml': '\u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d3e \u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02',
        'subs': [
            {
                'en_title': 'Lord, Teach Us to Pray (Luke 11:1)',
                'ml_title': '\u0d15\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d3e\u0d35\u0d47, \u0d1e\u0d19\u0d4d\u0d19\u0d33\u0d46 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d3f\u0d2a\u0d4d\u0d2a\u0d3e\u0d28\u0d4d\u200d \u0d2a\u0d20\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d15\u0d4d\u0d15\u0d47\u0d23\u0d2e\u0d47 (\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a 11:1)',
                'en_points': [
                    '>> The disciples saw the power of prayer in Jesus and asked Him to teach them.',
                    '>> Prayer is not a religious duty but intimate communion with the Father.',
                ],
                'ml_points': [
                    '>> \u0d36\u0d3f\u0d37\u0d4d\u0d2f\u0d28\u0d4d\u0d2e\u0d3e\u0d30\u0d4d\u200d \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d3e\u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02 \u0d15\u0d23\u0d4d\u0d1f\u0d4d \u0d2a\u0d20\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d15\u0d4d\u0d15\u0d3e\u0d28\u0d4d\u200d \u0d05\u0d2a\u0d47\u0d15\u0d4d\u0d37\u0d3f\u0d1a\u0d4d\u0d1a\u0d41.',
                    '>> \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28 \u0d15\u0d47\u0d35\u0d32\u0d02 \u0d1a\u0d1f\u0d19\u0d4d\u0d19\u0d32\u0d4d\u0d32, \u0d2a\u0d3f\u0d24\u0d3e\u0d35\u0d41\u0d2e\u0d3e\u0d2f\u0d41\u0d33\u0d4d\u0d33 \u0d06\u0d34\u0d2e\u0d47\u0d31\u0d3f\u0d2f \u0d2c\u0d28\u0d4d\u0d27\u0d2e\u0d3e\u0d23\u0d4d.',
                ],
                'en_verse': '\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 11:1', 'ml_verse': '\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 11:1',
            },
            {
                'en_title': 'Persistence in Prayer (Luke 11:5)',
                'ml_title': '\u0d28\u0d3f\u0d30\u0d28\u0d4d\u0d24\u0d30 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28 (\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a 11:5)',
                'en_points': [
                    '>> The Friend at Midnight: bold persistence in prayer yields results.',
                    '>> Ask, seek, knock \u2014 the Heavenly Father gives the Holy Spirit to those who ask.',
                ],
                'ml_points': [
                    '>> \u0d05\u0d30\u0d4d\u200d\u0d26\u0d4d\u0d27\u0d30\u0d3e\u0d24\u0d4d\u0d30\u0d3f\u0d2f\u0d3f\u0d32\u0d46 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d24\u0d28\u0d4d\u200d: \u0d32\u0d1c\u0d4d\u0d1c\u0d2f\u0d3f\u0d32\u0d4d\u0d32\u0d3e\u0d24\u0d4d\u0d24 \u0d2f\u0d3e\u0d1a\u0d28\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d4d \u0d09\u0d24\u0d4d\u0d24\u0d30\u0d02 \u0d32\u0d2d\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d02.',
                    '>> \u0d1a\u0d4b\u0d26\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d28\u0d4d\u200d, \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d\u0d15\u0d4d\u0d15\u0d4d \u0d28\u0d32\u0d4d\u200d\u0d15\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d02; \u0d24\u0d3f\u0d30\u0d2f\u0d41\u0d35\u0d3f\u0d28\u0d4d\u200d, \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d15\u0d23\u0d4d\u0d1f\u0d46\u0d24\u0d4d\u0d24\u0d41\u0d02.',
                ],
                'en_verse': 'Luke 11:5-13', 'ml_verse': '\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 11:5-13',
            },
            {
                'en_title': 'Prayer Life of Jesus',
                'ml_title': '\u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d3e \u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02',
                'en_points': [
                    '>> Jesus regularly withdrew to solitary places to pray (Luke 5:16).',
                    '>> He prayed all night before appointing the 12 apostles (Luke 6:12).',
                    '>> His High Priestly prayer \u2014 John 17.',
                ],
                'ml_points': [
                    '>> \u0d2f\u0d47\u0d36\u0d41 \u0d0f\u0d15\u0d3e\u0d28\u0d4d\u0d24\u0d38\u0d4d\u0d25\u0d32\u0d19\u0d4d\u0d19\u0d33\u0d3f\u0d32\u0d4d\u200d \u0d2a\u0d4b\u0d2f\u0d3f \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d3f\u0d1a\u0d4d\u0d1a\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28\u0d41 (\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a 5:16).',
                    '>> \u0d36\u0d3f\u0d37\u0d4d\u0d2f\u0d30\u0d46 \u0d24\u0d46\u0d30\u0d1e\u0d4d\u0d1e\u0d46\u0d1f\u0d41\u0d15\u0d4d\u0d15\u0d41\u0d02 \u0d2e\u0d41\u0d2e\u0d4d\u0d2a\u0d4d \u0d30\u0d3e\u0d24\u0d4d\u0d30\u0d3f \u0d2e\u0d41\u0d34\u0d41\u0d35\u0d28\u0d4d\u200d \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d3f\u0d1a\u0d4d\u0d1a\u0d41 (\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a 6:12).',
                    '>> \u0d2e\u0d39\u0d3e\u0d2a\u0d4c\u0d30\u0d4b\u0d39\u0d3f\u0d24\u0d4d\u0d2f \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28 \u2014 \u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 17.',
                ],
                'en_verse': 'Luke 5:16, 6:12, John 17', 'ml_verse': '\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 5:16, 6:12, \u0d2f\u0d4b\u0d39 17',
            },
            {
                'en_title': 'Prayer Warriors \u2014 Recommended Reading',
                'ml_title': '\u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d3e \u0d2f\u0d4b\u0d26\u0d4d\u0d27\u0d3e\u0d15\u0d4d\u0d15\u0d33\u0d4d\u200d \u2014 \u0d35\u0d3e\u0d2f\u0d28\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d4d',
                'en_points': [
                    '>> Leonard Ravenhill \u2014 "Why Revival Tarries"',
                    '>> E.M. Bounds \u2014 "Power Through Prayer"',
                    '>> C.H. Spurgeon \u2014 Known for fervent morning prayers; "The Prince of Preachers"',
                    '>> Martin Luther \u2014 "I have so much to do that I shall spend the first three hours in prayer."',
                ],
                'ml_points': [
                    '>> \u0d32\u0d46\u0d23\u0d3e\u0d30\u0d4d\u200d\u0d21\u0d4d \u0d31\u0d47\u0d35\u0d28\u0d4d\u200d\u0d39\u0d3f\u0d32\u0d4d\u200d \u2014 "Why Revival Tarries"',
                    '>> E.M. \u0d2c\u0d4c\u0d23\u0d4d\u0d1f\u0d4d\u200c\u0d38\u0d4d \u2014 "Power Through Prayer"',
                    '>> C.H. \u0d38\u0d4d\u200c\u0d2a\u0d30\u0d4d\u200d\u0d1c\u0d28\u0d4d\u200d \u2014 \u0d09\u0d1c\u0d4d\u0d1c\u0d4d\u0d35\u0d32\u0d2e\u0d3e\u0d2f \u0d2a\u0d4d\u0d30\u0d2d\u0d3e\u0d24 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d4d \u0d2a\u0d4d\u0d30\u0d36\u0d38\u0d4d\u0d24\u0d28\u0d4d\u200d',
                    '>> \u0d2e\u0d3e\u0d30\u0d4d\u200d\u0d1f\u0d4d\u0d1f\u0d3f\u0d28\u0d4d\u200d \u0d32\u0d42\u0d25\u0d30\u0d4d\u200d \u2014 "\u0d0e\u0d28\u0d3f\u0d15\u0d4d\u0d15\u0d4d \u0d07\u0d24\u0d4d\u0d30 \u0d1c\u0d4b\u0d32\u0d3f\u0d2f\u0d41\u0d33\u0d4d\u0d33\u0d24\u0d41\u0d15\u0d4a\u0d23\u0d4d\u0d1f\u0d4d \u0d06\u0d26\u0d4d\u0d2f \u0d2e\u0d42\u0d28\u0d4d\u0d28\u0d4d \u0d2e\u0d23\u0d3f\u0d15\u0d4d\u0d15\u0d42\u0d30\u0d4d\u200d \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d28\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d4d \u0d1a\u0d46\u0d32\u0d35\u0d3f\u0d1f\u0d41\u0d02."',
                ],
                'en_verse': '', 'ml_verse': '',
            },
        ]
    },
    # ─── 4. LOVE ───
    {
        'num': '04',
        'en': 'Love',
        'ml': '\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02',
        'subs': [
            {
                'en_title': 'The Greatest Commandment (Deut 6:4)',
                'ml_title': '\u0d0f\u0d31\u0d4d\u0d31\u0d35\u0d41\u0d02 \u0d35\u0d32\u0d3f\u0d2f \u0d15\u0d32\u0d4d\u200d\u0d2a\u0d28 (\u0d06\u0d35 6:4)',
                'en_points': [
                    '>> Hear O Israel: The LORD our God, the LORD is one.',
                    '>> Love the LORD your God with all your heart, soul, mind, and strength.',
                ],
                'ml_points': [
                    '>> \u0d2f\u0d3f\u0d38\u0d4d\u0d30\u0d3e\u0d2f\u0d47\u0d32\u0d47 \u0d15\u0d47\u0d33\u0d4d\u200d\u0d15\u0d4d\u0d15; \u0d28\u0d2e\u0d4d\u0d2e\u0d41\u0d1f\u0d46 \u0d26\u0d48\u0d35\u0d2e\u0d3e\u0d2f \u0d2f\u0d39\u0d4b\u0d35 \u0d0f\u0d15 \u0d2f\u0d39\u0d4b\u0d35 \u0d06\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d28\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d26\u0d48\u0d35\u0d2e\u0d3e\u0d2f \u0d2f\u0d39\u0d4b\u0d35\u0d2f\u0d46 \u0d2a\u0d42\u0d30\u0d4d\u200d\u0d23\u0d4d\u0d23 \u0d39\u0d43\u0d26\u0d2f\u0d24\u0d4d\u0d24\u0d4b\u0d1f\u0d41\u0d02 \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4b\u0d1f\u0d41\u0d02 \u0d36\u0d15\u0d4d\u0d24\u0d3f\u0d2f\u0d4b\u0d1f\u0d41\u0d02 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d15\u0d4d\u0d15\u0d47\u0d23\u0d02.',
                ],
                'en_verse': 'Deuteronomy 6:4-5', 'ml_verse': '\u0d06\u0d35\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d28\u0d02 6:4-5',
            },
            {
                'en_title': 'Love Your Neighbour (Luke 10:25-37 / Eph 2:1-4)',
                'ml_title': '\u0d05\u0d2f\u0d32\u0d4d\u200d\u0d15\u0d4d\u0d15\u0d3e\u0d30\u0d28\u0d46 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15 (\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a 10:25-37 / \u0d0e\u0d2b\u0d46 2:1-4)',
                'en_points': [
                    '>> The Good Samaritan: practical compassion is the proof of love.',
                    '>> God, being rich in mercy, loved us even when we were dead in sins.',
                ],
                'ml_points': [
                    '>> \u0d28\u0d32\u0d4d\u0d32 \u0d36\u0d2e\u0d30\u0d4d\u0d2f\u0d15\u0d4d\u0d15\u0d3e\u0d30\u0d28\u0d4d\u200d: \u0d2a\u0d4d\u0d30\u0d3e\u0d2f\u0d4b\u0d17\u0d3f\u0d15 \u0d05\u0d28\u0d41\u0d15\u0d2e\u0d4d\u0d2a\u0d2f\u0d3e\u0d23\u0d4d \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d24\u0d46\u0d33\u0d3f\u0d35\u0d4d.',
                    '>> \u0d15\u0d30\u0d41\u0d23\u0d3e\u0d38\u0d2e\u0d4d\u0d2a\u0d28\u0d4d\u0d28\u0d28\u0d3e\u0d2f \u0d26\u0d48\u0d35\u0d02 \u0d2a\u0d3e\u0d2a\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d4d\u200d \u0d2e\u0d30\u0d3f\u0d1a\u0d4d\u0d1a\u0d3f\u0d30\u0d41\u0d28\u0d4d\u0d28 \u0d28\u0d2e\u0d4d\u0d2e\u0d46\u0d2f\u0d41\u0d02 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d1a\u0d4d\u0d1a\u0d41.',
                ],
                'en_verse': 'Luke 10:25-37 & Ephesians 2:1-4', 'ml_verse': '\u0d32\u0d42\u0d15\u0d4d\u0d15\u0d4a\u0d38\u0d4d 10:25-37 & \u0d0e\u0d2b\u0d46\u0d38\u0d4d\u0d2f\u0d30\u0d4d\u200d 2:1-4',
            },
            # 7 BINDINGS inserted by builder between index 1 and 2
            {
                'en_title': 'Love Never Fails (1 Cor 13 & Rom 8:35-39)',
                'ml_title': '\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02 \u0d12\u0d30\u0d41\u0d28\u0d3e\u0d33\u0d41\u0d02 \u0d09\u0d24\u0d3f\u0d30\u0d4d\u200d\u0d28\u0d4d\u0d28\u0d41\u0d2a\u0d4b\u0d15\u0d2f\u0d3f\u0d32\u0d4d\u0d32 (1 \u0d15\u0d4a\u0d30\u0d3f 13 & \u0d31\u0d4b\u0d2e 8:35-39)',
                'en_points': [
                    '>> Love is patient, kind, does not envy, and never ends.',
                    '>> Nothing in all creation can separate us from the love of God in Christ.',
                ],
                'ml_points': [
                    '>> \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02 \u0d26\u0d40\u0d30\u0d4d\u200d\u0d18\u0d2e\u0d3e\u0d2f\u0d3f \u0d15\u0d4d\u0d37\u0d2e\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41, \u0d26\u0d2f \u0d15\u0d3e\u0d23\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41; \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d02 \u0d12\u0d1f\u0d41\u0d19\u0d4d\u0d19\u0d41\u0d15\u0d2f\u0d3f\u0d32\u0d4d\u0d32.',
                    '>> \u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d32\u0d41\u0d33\u0d4d\u0d33 \u0d26\u0d48\u0d35\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d28\u0d4d\u0d28\u0d4d \u0d28\u0d2e\u0d4d\u0d2e\u0d46 \u0d35\u0d47\u0d30\u0d4d\u200d\u0d2a\u0d3f\u0d30\u0d3f\u0d2a\u0d4d\u0d2a\u0d3e\u0d28\u0d4d\u200d \u0d12\u0d28\u0d4d\u0d28\u0d3f\u0d28\u0d41\u0d02 \u0d15\u0d34\u0d3f\u0d2f\u0d41\u0d15\u0d2f\u0d3f\u0d32\u0d4d\u0d32.',
                ],
                'en_verse': '1 Corinthians 13 & Romans 8:35-39', 'ml_verse': '1 \u0d15\u0d4a\u0d30\u0d3f\u0d28\u0d4d\u0d24\u0d4d\u0d2f\u0d30\u0d4d\u200d 13 & \u0d31\u0d4b\u0d2e\u0d30\u0d4d\u200d 8:35-39',
            },
            {
                'en_title': 'The New Commandment (John 13:34-35)',
                'ml_title': '\u0d2a\u0d41\u0d24\u0d3f\u0d2f \u0d15\u0d32\u0d4d\u200d\u0d2a\u0d28 (\u0d2f\u0d4b\u0d39 13:34-35)',
                'en_points': [
                    '>> Love one another as Christ has loved us.',
                    '>> By this love, the world will know we are His true disciples.',
                ],
                'ml_points': [
                    '>> \u0d1e\u0d3e\u0d28\u0d4d\u200d \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d46 \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d1a\u0d4d\u0d1a\u0d24\u0d41\u0d2a\u0d4b\u0d32\u0d46 \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d41\u0d02 \u0d24\u0d2e\u0d4d\u0d2e\u0d3f\u0d32\u0d4d\u200d \u0d24\u0d2e\u0d4d\u0d2e\u0d3f\u0d32\u0d4d\u200d \u0d38\u0d4d\u0d28\u0d47\u0d39\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d35\u0d3f\u0d28\u0d4d\u200d.',
                    '>> \u0d38\u0d39\u0d4b\u0d26\u0d30\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d24\u0d4d\u0d24\u0d3e\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d15\u0d4d\u0d30\u0d3f\u0d38\u0d4d\u0d24\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d36\u0d3f\u0d37\u0d4d\u0d2f\u0d30\u0d46\u0d28\u0d4d\u0d28\u0d4d \u0d32\u0d4b\u0d15\u0d02 \u0d05\u0d31\u0d3f\u0d2f\u0d41\u0d02.',
                ],
                'en_verse': 'John 13:34-35', 'ml_verse': '\u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 13:34-35',
            },
        ]
    },
    # ─── 5. RECEIVE HOLY SPIRIT ───
    {
        'num': '05',
        'en': 'Receive Holy Spirit',
        'ml': 'പരിശുദ്ധാത്മാവിനെ പ്രാപിക്കുക',
        'subs': [
            {
                'en_title': 'The Helper & Spirit of Truth (John 16:7-8, 13)',
                'ml_title': '\u0d15\u0d3e\u0d30\u0d4d\u0d2f\u0d38\u0d4d\u0d25\u0d28\u0d3e\u0d2f \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d3e\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d (\u0d2f\u0d4b\u0d39 16:7-8, 13)',
                'en_points': [
                    '>> The Holy Spirit convicts the world of sin, righteousness, and judgment.',
                    '>> He guides believers into all truth and glorifies Jesus Christ.',
                ],
                'ml_points': [
                    '>> \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d3e\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d \u0d32\u0d4b\u0d15\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d \u0d2a\u0d3e\u0d2a\u0d02, \u0d28\u0d40\u0d24\u0d3f, \u0d28\u0d4d\u0d2f\u0d3e\u0d2f\u0d35\u0d3f\u0d27\u0d3f \u0d07\u0d35\u0d2f\u0d46\u0d15\u0d4d\u0d15\u0d41\u0d31\u0d3f\u0d1a\u0d4d\u0d1a\u0d4d \u0d2c\u0d4b\u0d27\u0d4d\u0d2f\u0d02 \u0d35\u0d30\u0d41\u0d24\u0d4d\u0d24\u0d41\u0d02.',
                    '>> \u0d05\u0d35\u0d28\u0d4d\u200d \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d46 \u0d38\u0d15\u0d32 \u0d38\u0d24\u0d4d\u0d2f\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d41\u0d02 \u0d35\u0d34\u0d3f\u0d28\u0d1f\u0d24\u0d4d\u0d24\u0d41\u0d02.',
                ],
                'en_verse': 'John 16:7-8, 13', 'ml_verse': '\u0d2f\u0d4b\u0d39\u0d28\u0d4d\u0d28\u0d3e\u0d28\u0d4d\u200d 16:7-8, 13',
            },
            {
                'en_title': 'Power from On High (Acts 1:4, 8)',
                'ml_title': '\u0d09\u0d2f\u0d30\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d28\u0d4d\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d36\u0d15\u0d4d\u0d24\u0d3f (\u0d2a\u0d4d\u0d30\u0d35\u0d43 1:4, 8)',
                'en_points': [
                    '>> Wait for the promise of the Father \u2014 baptism in the Holy Spirit.',
                    '>> You will receive power to be bold witnesses to the ends of the earth.',
                ],
                'ml_points': [
                    '>> \u0d2a\u0d3f\u0d24\u0d3e\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d35\u0d3e\u0d17\u0d4d\u0d26\u0d24\u0d4d\u0d24\u0d2e\u0d3e\u0d2f \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d3e\u0d24\u0d4d\u0d2e \u0d38\u0d4d\u0d28\u0d3e\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d3e\u0d2f\u0d3f \u0d15\u0d3e\u0d24\u0d4d\u0d24\u0d3f\u0d30\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d15.',
                    '>> \u0d36\u0d15\u0d4d\u0d24\u0d3f \u0d32\u0d2d\u0d3f\u0d1a\u0d4d\u0d1a\u0d3f\u0d1f\u0d4d\u0d1f\u0d4d \u0d2d\u0d42\u0d2e\u0d3f\u0d2f\u0d41\u0d1f\u0d46 \u0d05\u0d31\u0d4d\u0d31\u0d24\u0d4d\u0d24\u0d4b\u0d33\u0d02 \u0d38\u0d3e\u0d15\u0d4d\u0d37\u0d3f\u0d15\u0d33\u0d3e\u0d15\u0d41\u0d02.',
                ],
                'en_verse': 'Acts 1:4, 8', 'ml_verse': '\u0d2a\u0d4d\u0d30\u0d35\u0d43\u0d24\u0d4d\u0d24\u0d3f\u0d15\u0d33\u0d4d\u200d 1:4, 8',
            },
            {
                'en_title': 'Outpouring of the Spirit in Acts',
                'ml_title': '\u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2a\u0d15\u0d30\u0d4d\u200d\u0d1a\u0d4d\u0d1a \u2014 \u0d2a\u0d4d\u0d30\u0d35\u0d43\u0d24\u0d4d\u0d24\u0d3f\u0d15\u0d33\u0d4d\u200d',
                'en_points': [
                    '>> Jerusalem: All filled, spoke in tongues (Acts 2:1-4).',
                    '>> Samaria: Received the Spirit through laying on of hands (Acts 8:17).',
                    '>> Caesarea \u2014 Cornelius: Holy Spirit fell on all who heard the Word (Acts 10:44).',
                    '>> Ephesus: Filled with the Spirit and prophesied (Acts 19:6).',
                ],
                'ml_points': [
                    '>> യെരൂശലേം: എല്ലാവരും പരിശുദ്ധാത്മാവ് നിറഞ്ഞു സംസാരിച്ചു (പ്രവൃ 2:1-4).',
                    '>> \u0d36\u0d2e\u0d30\u0d4d\u0d2f: \u0d15\u0d48\u0d35\u0d46\u0d1a\u0d4d\u0d1a\u0d41 \u0d2a\u0d4d\u0d30\u0d3e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d25\u0d3f\u0d1a\u0d4d\u0d1a\u0d2a\u0d4d\u0d2a\u0d4b\u0d33\u0d4d\u200d \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d \u0d32\u0d2d\u0d3f\u0d1a\u0d4d\u0d1a\u0d41 (\u0d2a\u0d4d\u0d30\u0d35\u0d43 8:17).',
                    '>> \u0d15\u0d48\u0d38\u0d30\u0d4d\u0d2f \u2014 \u0d15\u0d4a\u0d30\u0d4d\u200d\u0d28\u0d4d\u0d28\u0d47\u0d32\u0d4d\u0d2f\u0d4a\u0d38\u0d4d: \u0d35\u0d1a\u0d28\u0d02 \u0d15\u0d47\u0d1f\u0d4d\u0d1f \u0d0f\u0d35\u0d30\u0d41\u0d1f\u0d46\u0d2f\u0d41\u0d02 \u0d2e\u0d47\u0d32\u0d4d\u200d \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d \u0d07\u0d31\u0d19\u0d4d\u0d19\u0d3f (\u0d2a\u0d4d\u0d30\u0d35\u0d43 10:44).',
                    '>> \u0d0e\u0d2b\u0d46\u0d38\u0d4a\u0d38\u0d4d: \u0d15\u0d48\u0d35\u0d46\u0d1a\u0d4d\u0d1a\u0d2a\u0d4d\u0d2a\u0d4b\u0d33\u0d4d\u200d \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d \u0d35\u0d28\u0d4d\u0d28\u0d41 \u0d2a\u0d4d\u0d30\u0d35\u0d1a\u0d3f\u0d1a\u0d4d\u0d1a\u0d41 (\u0d2a\u0d4d\u0d30\u0d35\u0d43 19:6).',
                ],
                'en_verse': 'Acts 2:1-4, 8:17, 10:44, 19:6',
                'ml_verse': '\u0d2a\u0d4d\u0d30\u0d35\u0d43\u0d24\u0d4d\u0d24\u0d3f\u0d15\u0d33\u0d4d\u200d 2:1-4, 8:17, 10:44, 19:6',
            },
        ]
    },
    # ─── 6. HOLY LIFE ───
    {
        'num': '06',
        'en': 'Holy Life',
        'ml': '\u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27 \u0d1c\u0d40\u0d35\u0d3f\u0d24\u0d02',
        'subs': [
            {
                'en_title': 'Be Holy as God is Holy (1 Peter 1:14-16)',
                'ml_title': '\u0d26\u0d48\u0d35\u0d02 \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d3e\u0d15\u0d2f\u0d3e\u0d32\u0d4d\u200d \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d30\u0d3e\u0d15\u0d41\u0d15 (1 \u0d2a\u0d24\u0d4d\u0d30\u0d4a 1:14-16)',
                'en_points': [
                    '>> Do not conform to the former worldly passions.',
                    '>> Be holy in all your conduct and daily lifestyle.',
                ],
                'ml_points': [
                    '>> \u0d2e\u0d41\u0d28\u0d4d\u200d\u0d15\u0d3e\u0d32 \u0d32\u0d4b\u0d15 \u0d2e\u0d4b\u0d39\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d\u0d15\u0d4d\u0d15\u0d4d \u0d05\u0d28\u0d41\u0d30\u0d42\u0d2a\u0d30\u0d3e\u0d15\u0d30\u0d41\u0d24\u0d4d.',
                    '>> \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d41\u0d1f\u0d46 \u0d0e\u0d32\u0d4d\u0d32\u0d3e \u0d28\u0d1f\u0d24\u0d4d\u0d24\u0d2f\u0d3f\u0d32\u0d41\u0d02 \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d30\u0d3e\u0d2f\u0d3f\u0d30\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d28\u0d4d\u200d.',
                ],
                'en_verse': '1 Peter 1:14-16 \u2014 "Be holy, for I am holy."',
                'ml_verse': '1 \u0d2a\u0d24\u0d4d\u0d30\u0d4a\u0d38\u0d4d 1:14-16 \u2014 "\u0d1e\u0d3e\u0d28\u0d4d\u200d \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d \u0d06\u0d15\u0d2f\u0d3e\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d41\u0d02 \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d30\u0d3e\u0d2f\u0d3f\u0d30\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d28\u0d4d\u200d."',
            },
            {
                'en_title': 'Vision of God\'s Holiness (Isaiah 6:1-3)',
                'ml_title': '\u0d26\u0d48\u0d35\u0d40\u0d15 \u0d35\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d3f\u0d2f\u0d41\u0d1f\u0d46 \u0d26\u0d30\u0d4d\u200d\u0d36\u0d28\u0d02 (\u0d2f\u0d46\u0d36 6:1-3)',
                'en_points': [
                    '>> Isaiah saw the Lord high and exalted; Seraphim cried "Holy, Holy, Holy."',
                    '>> The heavenly throne glorifies the Holy and Righteous Lord.',
                ],
                'ml_points': [
                    '>> \u0d2f\u0d46\u0d36\u0d2f\u0d4d\u0d2f\u0d3e\u0d35\u0d4d \u0d15\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d3e\u0d35\u0d3f\u0d28\u0d46 \u0d09\u0d28\u0d4d\u0d28\u0d24\u0d2e\u0d3e\u0d2f\u0d3f \u0d26\u0d30\u0d4d\u200d\u0d36\u0d3f\u0d1a\u0d4d\u0d1a\u0d41; \u0d38\u0d46\u0d31\u0d3e\u0d2b\u0d41\u0d15\u0d33\u0d4d\u200d "\u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d" \u0d0e\u0d28\u0d4d\u0d28\u0d4d \u0d06\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d41.',
                    '>> \u0d38\u0d4d\u0d35\u0d30\u0d4d\u200d\u0d17\u0d4d\u0d17\u0d40\u0d2f \u0d38\u0d3f\u0d02\u0d39\u0d3e\u0d38\u0d28\u0d02 \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d3e\u0d2f \u0d15\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d3e\u0d35\u0d3f\u0d28\u0d46 \u0d2a\u0d41\u0d15\u0d34\u0d4d\u0d24\u0d4d\u0d24\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                ],
                'en_verse': 'Isaiah 6:1-3', 'ml_verse': '\u0d2f\u0d46\u0d36\u0d2f\u0d4d\u0d2f\u0d3e\u0d35\u0d41 6:1-3',
            },
            {
                'en_title': 'Holy, Holy, Holy (Rev 4:8, 11 & Acts 3:14)',
                'ml_title': '\u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d, \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d, \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d (\u0d35\u0d46\u0d33\u0d3f 4:8, 11 & \u0d2a\u0d4d\u0d30\u0d35\u0d43 3:14)',
                'en_points': [
                    '>> "Holy, holy, holy, is the Lord God Almighty, who was and is and is to come!"',
                    '>> Jesus is the Holy and Righteous One who set us apart.',
                ],
                'ml_points': [
                    '>> "\u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d, \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d, \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d4d\u200d; \u0d38\u0d30\u0d4d\u200d\u0d35\u0d4d\u0d35\u0d36\u0d15\u0d4d\u0d24\u0d3f\u0d2f\u0d41\u0d33\u0d4d\u0d33 \u0d26\u0d48\u0d35\u0d2e\u0d3e\u0d2f \u0d15\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d24\u0d3e\u0d35\u0d4d!"',
                    '>> \u0d2f\u0d47\u0d36\u0d41 \u0d28\u0d2e\u0d4d\u0d2e\u0d46 \u0d24\u0d28\u0d3f\u0d15\u0d4d\u0d15\u0d3e\u0d2f\u0d3f \u0d35\u0d47\u0d30\u0d4d\u200d\u0d24\u0d3f\u0d30\u0d3f\u0d1a\u0d4d\u0d1a \u0d2a\u0d30\u0d3f\u0d36\u0d41\u0d26\u0d4d\u0d27\u0d28\u0d41\u0d02 \u0d28\u0d40\u0d24\u0d3f\u0d2e\u0d3e\u0d28\u0d41\u0d2e\u0d3e\u0d23\u0d4d.',
                ],
                'en_verse': 'Revelation 4:8, 11 & Acts 3:14',
                'ml_verse': '\u0d35\u0d46\u0d33\u0d3f\u0d2a\u0d4d\u0d2a\u0d3e\u0d1f\u0d41 4:8, 11 & \u0d2a\u0d4d\u0d30\u0d35\u0d43\u0d24\u0d4d\u0d24\u0d3f\u0d15\u0d33\u0d4d\u200d 3:14',
            },
            {
                'en_title': 'Living Sacrifice & Transformed Mind (Rom 12:1-2)',
                'ml_title': '\u0d1c\u0d40\u0d35\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d2f\u0d3e\u0d17\u0d35\u0d41\u0d02 \u0d30\u0d42\u0d2a\u0d3e\u0d28\u0d4d\u0d24\u0d30\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d28\u0d4d\u0d28 \u0d2e\u0d28\u0d38\u0d4d\u0d38\u0d41\u0d02 (\u0d31\u0d4b\u0d2e 12:1-2)',
                'en_points': [
                    '>> Present your bodies as a living sacrifice, holy and acceptable to God.',
                    '>> Do not be conformed to this world, but be transformed by renewing your mind.',
                ],
                'ml_points': [
                    '>> \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d41\u0d1f\u0d46 \u0d36\u0d30\u0d40\u0d30\u0d19\u0d4d\u0d19\u0d33\u0d46 \u0d26\u0d48\u0d35\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d28\u0d4d \u0d2a\u0d4d\u0d30\u0d38\u0d3e\u0d26\u0d2e\u0d41\u0d33\u0d4d\u0d33 \u0d1c\u0d40\u0d35\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d2f\u0d3e\u0d17\u0d2e\u0d3e\u0d2f\u0d3f \u0d38\u0d2e\u0d30\u0d4d\u200d\u0d2a\u0d4d\u0d2a\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d28\u0d4d\u200d.',
                    '>> \u0d08 \u0d32\u0d4b\u0d15\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d28\u0d4d \u0d05\u0d28\u0d41\u0d30\u0d42\u0d2a\u0d2e\u0d3e\u0d15\u0d3e\u0d24\u0d46 \u0d2e\u0d28\u0d38\u0d4d\u0d38\u0d4d \u0d2a\u0d41\u0d24\u0d41\u0d15\u0d4d\u0d15\u0d3f \u0d30\u0d42\u0d2a\u0d3e\u0d28\u0d4d\u0d24\u0d30\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d35\u0d3f\u0d28\u0d4d\u200d.',
                ],
                'en_verse': 'Romans 12:1-2', 'ml_verse': '\u0d31\u0d4b\u0d2e\u0d30\u0d4d\u200d 12:1-2',
            },
        ]
    },
    # ─── 7. ETERNAL HOPE ───
    {
        'num': '07',
        'en': 'Eternal Hope (Eternity)',
        'ml': '\u0d28\u0d3f\u0d24\u0d4d\u0d2f\u0d2a\u0d4d\u0d30\u0d24\u0d4d\u0d2f\u0d3e\u0d36 (\u0d28\u0d3f\u0d24\u0d4d\u0d2f\u0d24)',
        'subs': [
            {
                'en_title': 'Temporary Troubles vs. Eternal Glory (2 Cor 4:17-18)',
                'ml_title': 'ക്ഷണികമായ കഷ്ടവും നിത്യ തേജസ്സിന്റെ ഘനവും (2 കൊരി 4:17-18)',
                'en_points': [
                    '>> Light momentary troubles prepare us for an eternal weight of glory.',
                    '>> We look not to things seen, but to the eternal unseen reality.',
                ],
                'ml_points': [
                    '>> \u0d28\u0d4a\u0d1f\u0d3f\u0d28\u0d47\u0d30\u0d24\u0d4d\u0d24\u0d47\u0d15\u0d4d\u0d15\u0d41\u0d33\u0d4d\u0d33 \u0d32\u0d18\u0d41\u0d35\u0d3e\u0d2f \u0d15\u0d37\u0d4d\u0d1f\u0d02 \u0d28\u0d3f\u0d24\u0d4d\u0d2f\u0d2e\u0d3e\u0d2f \u0d24\u0d47\u0d1c\u0d4b\u0d2d\u0d3e\u0d30\u0d02 \u0d09\u0d33\u0d35\u0d3e\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d28\u0d3e\u0d02 \u0d15\u0d3e\u0d23\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d3f\u0d28\u0d46\u0d2f\u0d32\u0d4d\u0d32, \u0d15\u0d3e\u0d23\u0d3e\u0d24\u0d4d\u0d24 \u0d28\u0d3f\u0d24\u0d4d\u0d2f\u0d24\u0d2f\u0d46\u0d2f\u0d24\u0d4d\u0d30\u0d47 \u0d28\u0d4b\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d4d.',
                ],
                'en_verse': '2 Corinthians 4:17-18', 'ml_verse': '2 \u0d15\u0d4a\u0d30\u0d3f\u0d28\u0d4d\u0d24\u0d4d\u0d2f\u0d30\u0d4d\u200d 4:17-18',
            },
            {
                'en_title': 'Living Hope & Imperishable Inheritance (1 Pet 1:4-6)',
                'ml_title': '\u0d1c\u0d40\u0d35\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d2a\u0d4d\u0d30\u0d24\u0d4d\u0d2f\u0d3e\u0d36\u0d2f\u0d41\u0d02 \u0d35\u0d3e\u0d1f\u0d3e\u0d24\u0d4d\u0d24 \u0d05\u0d35\u0d15\u0d3e\u0d36\u0d35\u0d41\u0d02 (1 \u0d2a\u0d24\u0d4d\u0d30\u0d4a 1:4-6)',
                'en_points': [
                    '>> Born again into a living hope through the resurrection of Jesus.',
                    '>> An inheritance imperishable, undefiled, and unfading in heaven.',
                ],
                'ml_points': [
                    '>> \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d2a\u0d41\u0d28\u0d30\u0d41\u0d24\u0d4d\u0d25\u0d3e\u0d28\u0d24\u0d4d\u0d24\u0d3e\u0d32\u0d4d\u200d \u0d1c\u0d40\u0d35\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d2a\u0d4d\u0d30\u0d24\u0d4d\u0d2f\u0d3e\u0d36\u0d2f\u0d4d\u0d15\u0d4d\u0d15\u0d3e\u0d2f\u0d3f \u0d35\u0d40\u0d23\u0d4d\u0d1f\u0d41\u0d02 \u0d1c\u0d28\u0d3f\u0d1a\u0d4d\u0d1a\u0d41.',
                    '>> \u0d05\u0d15\u0d4d\u0d37\u0d2f\u0d35\u0d41\u0d02 \u0d28\u0d3f\u0d30\u0d4d\u200d\u0d2e\u0d4d\u0d2e\u0d32\u0d35\u0d41\u0d02 \u0d35\u0d3e\u0d1f\u0d3e\u0d24\u0d4d\u0d24\u0d24\u0d41\u0d2e\u0d3e\u0d2f \u0d05\u0d35\u0d15\u0d3e\u0d36\u0d02 \u0d38\u0d4d\u0d35\u0d30\u0d4d\u200d\u0d17\u0d4d\u0d17\u0d24\u0d4d\u0d24\u0d3f\u0d32\u0d4d\u200d \u0d38\u0d42\u0d15\u0d4d\u0d37\u0d3f\u0d1a\u0d4d\u0d1a\u0d3f\u0d30\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                ],
                'en_verse': '1 Peter 1:4-6', 'ml_verse': '1 \u0d2a\u0d24\u0d4d\u0d30\u0d4a\u0d38\u0d4d 1:4-6',
            },
            {
                'en_title': 'Fix Our Eyes on Jesus (Heb 12:2-3, Rom 8:18)',
                'ml_title': '\u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d32\u0d47\u0d15\u0d4d\u0d15\u0d4d \u0d28\u0d4b\u0d15\u0d4d\u0d15\u0d3f \u0d13\u0d1f\u0d41\u0d15 (\u0d0e\u0d2c\u0d4d\u0d30\u0d3e 12:2-3, \u0d31\u0d4b\u0d2e 8:18)',
                'en_points': [
                    '>> Looking to Jesus, the founder and perfecter of our faith.',
                    '>> Present sufferings are not worth comparing with the glory to be revealed.',
                ],
                'ml_points': [
                    '>> \u0d35\u0d3f\u0d36\u0d4d\u0d35\u0d3e\u0d38\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d28\u0d3e\u0d2f\u0d15\u0d28\u0d3e\u0d2f \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d46 \u0d28\u0d4b\u0d15\u0d4d\u0d15\u0d3f \u0d38\u0d4d\u0d25\u0d3f\u0d30\u0d24\u0d2f\u0d4b\u0d1f\u0d46 \u0d13\u0d1f\u0d41\u0d15.',
                    '>> \u0d07\u0d2a\u0d4d\u0d2a\u0d4b\u0d34\u0d24\u0d4d\u0d24\u0d46 \u0d15\u0d37\u0d4d\u0d1f\u0d19\u0d4d\u0d19\u0d33\u0d4d\u200d \u0d35\u0d46\u0d33\u0d3f\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d3e\u0d28\u0d3f\u0d30\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28 \u0d24\u0d47\u0d1c\u0d38\u0d4d\u0d38\u0d41\u0d2e\u0d3e\u0d2f\u0d3f \u0d24\u0d41\u0d32\u0d28\u0d02 \u0d1a\u0d46\u0d2f\u0d4d\u0d2f\u0d3e\u0d28\u0d3e\u0d15\u0d3f\u0d32\u0d4d\u0d32.',
                ],
                'en_verse': 'Hebrews 12:2-3 & Romans 8:18', 'ml_verse': '\u0d0e\u0d2c\u0d4d\u0d30\u0d3e\u0d2f\u0d30\u0d4d\u200d 12:2-3 & \u0d31\u0d4b\u0d2e\u0d30\u0d4d\u200d 8:18',
            },
            {
                'en_title': 'The Resurrection Life (Romans 8:11)',
                'ml_title': '\u0d2a\u0d41\u0d28\u0d30\u0d41\u0d24\u0d4d\u0d25\u0d3e\u0d28\u0d24\u0d4d\u0d24\u0d3f\u0d28\u0d4d\u0d31\u0d46 \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d (\u0d31\u0d4b\u0d2e 8:11)',
                'en_points': [
                    '>> The same Spirit who raised Jesus from the dead dwells in us.',
                    '>> He will give eternal life and glory to our mortal bodies.',
                ],
                'ml_points': [
                    '>> \u0d2f\u0d47\u0d36\u0d41\u0d35\u0d3f\u0d28\u0d46 \u0d2e\u0d30\u0d3f\u0d1a\u0d4d\u0d1a\u0d35\u0d30\u0d3f\u0d32\u0d4d\u200d \u0d28\u0d3f\u0d28\u0d4d\u0d28\u0d4d \u0d09\u0d2f\u0d3f\u0d30\u0d4d\u200d\u0d2a\u0d4d\u0d2a\u0d3f\u0d1a\u0d4d\u0d1a\u0d35\u0d28\u0d4d\u0d31\u0d46 \u0d06\u0d24\u0d4d\u0d2e\u0d3e\u0d35\u0d4d \u0d28\u0d2e\u0d4d\u0d2e\u0d3f\u0d32\u0d4d\u200d \u0d35\u0d38\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d41.',
                    '>> \u0d05\u0d35\u0d28\u0d4d\u200d \u0d28\u0d2e\u0d4d\u0d2e\u0d41\u0d1f\u0d46 \u0d2e\u0d30\u0d4d\u200d\u0d24\u0d4d\u0d2f\u0d36\u0d30\u0d40\u0d30\u0d19\u0d4d\u0d19\u0d33\u0d46\u0d2f\u0d41\u0d02 \u0d1c\u0d40\u0d35\u0d3f\u0d2a\u0d4d\u0d2a\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d02.',
                ],
                'en_verse': 'Romans 8:11', 'ml_verse': '\u0d31\u0d4b\u0d2e\u0d30\u0d4d\u200d 8:11',
            },
        ]
    },
]



def make_thank_you_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s)
    bar(s, 0, Inches(0.1), GOLD)
    bar(s, SH - Inches(0.1), Inches(0.1), GOLD)
    
    ci = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.66), Inches(2.0), Inches(10.0), Inches(3.5))
    ci.fill.solid(); ci.fill.fore_color.rgb = RGBColor(0x0A,0x11,0x28) 
    ci.line.color.rgb = GOLD; ci.line.width = Pt(3)
    
    en = txtbox(s, "THANK YOU!", Inches(2.0), Inches(2.5), Inches(9.33), Inches(1.0),
           sz=Pt(65), bold=True, color=GOLD, align=PP_ALIGN.CENTER, font=EN_HEAD)
    ml = txtbox(s, "ക്ലാസ്സിൽ പങ്കെടുത്ത എല്ലാവർക്കും നന്ദി", Inches(2.0), Inches(3.8), Inches(9.33), Inches(0.8),
           sz=Pt(36), bold=True, color=CYAN, align=PP_ALIGN.CENTER, font=ML_FONT)
           
    apply_click_animations(s, [[ci, en, ml]])

# ══════════════════════════════════════════════
#  MAIN BUILDER
# ══════════════════════════════════════════════
def generate():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    make_title(prs)
    make_intro(prs)  # Blank introduction slide

    for topic in TOPICS:
        if topic['num'] == '03':
            make_bible_acronym_slide(prs)

        make_divider(prs, topic['num'], topic['en'], topic['ml'])
        en_hdr = f"{topic['num']}. {topic['en']}"
        ml_hdr = f"{topic['num']}. {topic['ml']}"

        for idx, sp in enumerate(topic['subs']):
            # Insert 7 bindings slide after Love sub-point 2 (index 1)
            if topic['num'] == '04' and idx == 2:
                make_bindings_slide(prs)
            make_content_slide(prs, en_hdr, ml_hdr, sp)

    make_closing(prs)

    make_thank_you_slide(prs)

    out = "Victorious_Christian_Life.pptx"
    prs.save(out)
    print(f"[OK] Generated: {out} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    print("=" * 60)
    print("  Victorious Christian Life - Bilingual PPT Generator")
    print("=" * 60)
    generate()
    print("\nDone! Single combined EN+ML presentation created.")
