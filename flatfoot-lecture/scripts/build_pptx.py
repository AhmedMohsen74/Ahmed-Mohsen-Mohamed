#!/usr/bin/env python3
"""Professional 50-slide Flat Foot lecture — tight layouts, no overflow."""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
OUT = ROOT / "Flat_Foot_Kaplan_Orthopedic_Lecture.pptx"

NAVY = RGBColor(0x1B, 0x28, 0x38)
TEAL = RGBColor(0x0D, 0x73, 0x77)
TEAL2 = RGBColor(0x2E, 0x8B, 0x8C)
CORAL = RGBColor(0xE0, 0x7A, 0x5F)
CORAL2 = RGBColor(0xB4, 0x46, 0x46)
SOFT = RGBColor(0xF4, 0xF7, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x6B, 0x7C, 0x8A)
SLATE = RGBColor(0x2C, 0x3E, 0x50)
LIGHT = RGBColor(0xE4, 0xF1, 0xF0)
LINE = RGBColor(0xD0, 0xDC, 0xE0)

W, H = Inches(13.333), Inches(7.5)
# Safe content band under header / above footer
CONTENT_TOP = Inches(1.45)
CONTENT_BOTTOM = Inches(6.85)
FOOTER_Y = Inches(7.05)


def fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def no_line(shape):
    shape.line.fill.background()


def set_run(run, size=16, bold=False, color=NAVY, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def textbox(slide, l, t, w, h, text, size=16, bold=False, color=NAVY,
            align=PP_ALIGN.LEFT, name="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf.auto_size = None
    except Exception:
        pass
    # vertical anchor
    body = tf._txBody
    bodyPr = body.find(qn("a:bodyPr"))
    if bodyPr is not None:
        anchor_map = {
            MSO_ANCHOR.TOP: "t",
            MSO_ANCHOR.MIDDLE: "ctr",
            MSO_ANCHOR.BOTTOM: "b",
        }
        bodyPr.set("anchor", anchor_map.get(anchor, "t"))
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, name=name)
    return box


def bullets(slide, l, t, w, h, lines, size=15, color=SLATE, spacing=6, bold_first=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = ("•  " + line) if line.strip() else ""
        set_run(run, size=size, bold=(bold_first and i == 0), color=color)
    return box


def rect(slide, l, t, w, h, color, lined=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    fill(sh, color)
    if lined is None:
        no_line(sh)
    else:
        sh.line.color.rgb = lined
        sh.line.width = Pt(1.25)
    return sh


def rrect(slide, l, t, w, h, color, lined=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill(sh, color)
    # tighter corner
    try:
        sh.adjustments[0] = 0.08
    except Exception:
        pass
    if lined is None:
        no_line(sh)
    else:
        sh.line.color.rgb = lined
        sh.line.width = Pt(1.5)
    return sh


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, SOFT)  # soft page background
    return s


def footer(slide, n, total=50):
    rect(slide, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.015), LINE)
    textbox(slide, Inches(0.5), FOOTER_Y, Inches(8), Inches(0.3),
            "Flat Foot — Kaplan Orthopedic Lecture", size=10, color=MUTED)
    textbox(slide, Inches(11.0), FOOTER_Y, Inches(1.9), Inches(0.3),
            f"{n}  /  {total}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)


def header(slide, title, subtitle=None):
    rect(slide, 0, 0, Inches(0.14), H, TEAL)
    rect(slide, 0, 0, W, Inches(1.2), NAVY)
    textbox(slide, Inches(0.45), Inches(0.28), Inches(12.4), Inches(0.5),
            title, size=26, bold=True, color=WHITE, name="Georgia")
    if subtitle:
        textbox(slide, Inches(0.45), Inches(0.78), Inches(12.4), Inches(0.32),
                subtitle, size=13, color=LIGHT)


def card(slide, l, t, w, h, title, lines, accent=TEAL, title_size=15, body_size=13):
    """Filled content card — title strip + bullets, no huge empty space."""
    rrect(slide, l, t, w, h, WHITE, LINE)
    rect(slide, l, t, w, Inches(0.06), accent)
    textbox(slide, l + Inches(0.18), t + Inches(0.16), w - Inches(0.3), Inches(0.35),
            title, size=title_size, bold=True, color=accent, name="Georgia")
    if lines:
        bullets(slide, l + Inches(0.18), t + Inches(0.55), w - Inches(0.32), h - Inches(0.7),
                lines, size=body_size, color=SLATE, spacing=4)


def pill(slide, l, t, w, h, text, bg=TEAL, fg=WHITE, size=13):
    rrect(slide, l, t, w, h, bg)
    textbox(slide, l, t + Inches(0.06), w, h - Inches(0.08),
            text, size=size, bold=True, color=fg, align=PP_ALIGN.CENTER)


def row_item(slide, y, left_label, body, accent=TEAL, label_w=Inches(2.5)):
    rrect(slide, Inches(0.45), y, label_w, Inches(0.58), accent)
    textbox(slide, Inches(0.5), y + Inches(0.12), label_w - Inches(0.1), Inches(0.4),
            left_label, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rrect(slide, Inches(0.45) + label_w + Inches(0.12), y,
          Inches(12.4) - label_w - Inches(0.12), Inches(0.58), WHITE, LINE)
    textbox(slide, Inches(0.7) + label_w, y + Inches(0.12),
            Inches(11.8) - label_w, Inches(0.4), body, size=13, color=SLATE)


def two_col_compare(slide, left_title, left_lines, right_title, right_lines,
                    left_color=TEAL, right_color=CORAL, left_note=None, right_note=None):
    """Side-by-side compare cards sized to content (no huge empty bottoms)."""
    n = max(len(left_lines), len(right_lines))
    # header 0.65 + bullets (~0.42 each) + optional note 0.7 + padding
    body_h = Inches(0.65) + Inches(0.48 * n) + Inches(0.35)
    if left_note or right_note:
        body_h += Inches(0.85)
    card_h = min(body_h, Inches(5.2))
    # left
    rrect(slide, Inches(0.4), CONTENT_TOP, Inches(6.1), card_h, WHITE, left_color)
    rect(slide, Inches(0.4), CONTENT_TOP, Inches(6.1), Inches(0.65), left_color)
    textbox(slide, Inches(0.4), CONTENT_TOP + Inches(0.15), Inches(6.1), Inches(0.4),
            left_title, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Georgia")
    bullets(slide, Inches(0.7), CONTENT_TOP + Inches(0.85), Inches(5.5), Inches(0.48 * len(left_lines) + 0.2),
            left_lines, size=14, spacing=6)
    if left_note:
        ny = CONTENT_TOP + card_h - Inches(0.8)
        rrect(slide, Inches(0.6), ny, Inches(5.7), Inches(0.55), LIGHT, left_color)
        textbox(slide, Inches(0.75), ny + Inches(0.1), Inches(5.4), Inches(0.4),
                left_note, size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    # right
    rrect(slide, Inches(6.8), CONTENT_TOP, Inches(6.1), card_h, WHITE, right_color)
    rect(slide, Inches(6.8), CONTENT_TOP, Inches(6.1), Inches(0.65), right_color)
    textbox(slide, Inches(6.8), CONTENT_TOP + Inches(0.15), Inches(6.1), Inches(0.4),
            right_title, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Georgia")
    bullets(slide, Inches(7.1), CONTENT_TOP + Inches(0.85), Inches(5.5), Inches(0.48 * len(right_lines) + 0.2),
            right_lines, size=14, spacing=6)
    if right_note:
        ny = CONTENT_TOP + card_h - Inches(0.8)
        rrect(slide, Inches(7.0), ny, Inches(5.7), Inches(0.55), RGBColor(0xFD, 0xF0, 0xEB), right_color)
        textbox(slide, Inches(7.15), ny + Inches(0.1), Inches(5.4), Inches(0.4),
                right_note, size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)


def grid_cards(slide, items, cols=3, accent=TEAL, footer_note=None):
    """Content-sized cards (no tall empty interiors). Note docks to bottom of band."""
    n = len(items)
    rows = (n + cols - 1) // cols
    gap_x, gap_y = Inches(0.22), Inches(0.2)
    usable_w = Inches(12.4)
    usable_h = Inches(5.2)
    card_w = (usable_w - gap_x * (cols - 1)) / cols
    max_lines = max((len(lines) for _, lines in items), default=3)
    # Tight to content — never stretch into empty white boxes
    card_h = Inches(0.72) + Inches(0.42 * max_lines) + Inches(0.28)

    if footer_note is None and rows == 1:
        footer_note = "Kaplan: classify flexibility and stage before choosing treatment."

    note_h = Inches(0.8) if footer_note else Inches(0)
    # If cards + gaps would collide with note/footer band, shrink slightly
    total = rows * card_h + (rows - 1) * gap_y
    limit = usable_h - (note_h + gap_y if footer_note else Inches(0))
    if total > limit:
        card_h = (limit - gap_y * (rows - 1)) / rows
    for i, (title, lines) in enumerate(items):
        c, r = i % cols, i // cols
        x = Inches(0.45) + c * (card_w + gap_x)
        y = CONTENT_TOP + r * (card_h + gap_y)
        acc = accent if isinstance(accent, RGBColor) else accent[i % len(accent)]
        card(slide, x, y, card_w, card_h, title, lines, acc,
             title_size=15, body_size=13)
    if footer_note:
        # Sit note just under the cards (not at slide bottom) to avoid a dead middle gap
        y = CONTENT_TOP + rows * (card_h + gap_y) + Inches(0.15)
        # If that would collide with footer, pull up slightly
        max_y = CONTENT_TOP + usable_h - note_h
        if y > max_y:
            y = max_y
        rrect(slide, Inches(0.45), y, Inches(12.4), note_h, LIGHT, TEAL)
        textbox(slide, Inches(0.7), y + Inches(0.2), Inches(11.9), Inches(0.45),
                footer_note, size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)


def add_pic(slide, path, l, t, w=None, h=None):
    if not Path(path).exists():
        return None
    kw = {"left": l, "top": t}
    if w:
        kw["width"] = w
    if h:
        kw["height"] = h
    return slide.shapes.add_picture(str(path), **kw)


# ───────────────────── SLIDES ─────────────────────

def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, NAVY)
    add_pic(s, ASSETS / "title_hero.png", 0, 0, w=W)
    textbox(s, Inches(0.55), Inches(6.55), Inches(10), Inches(0.35),
            "Orthopedic Surgery  ·  Foot & Ankle Module", size=13, color=LIGHT)
    return s


def s02(prs, n):
    s = blank(prs)
    header(s, "Learning Objectives", "By the end of this lecture you will be able to…")
    objs = [
        "Define pes planus and distinguish flexible vs rigid deformity",
        "Describe anatomy & biomechanics of the medial longitudinal arch",
        "Classify adult acquired flatfoot (Johnson–Strom / Myerson)",
        "Perform a focused flatfoot exam and interpret key signs",
        "Select imaging and radiographic measurements that matter",
        "Build a stage-based nonoperative and surgical treatment plan",
        "Recognize pediatric coalitions and when to operate",
        "Recall Kaplan high-yield pearls for boards & OR",
    ]
    for i, text in enumerate(objs):
        y = CONTENT_TOP + Inches(i * 0.62)
        pill(s, Inches(0.45), y, Inches(0.62), Inches(0.48), f"{i+1:02d}", size=14)
        rrect(s, Inches(1.2), y, Inches(11.6), Inches(0.48), WHITE, LINE)
        textbox(s, Inches(1.4), y + Inches(0.1), Inches(11.2), Inches(0.35),
                text, size=14, color=SLATE)
    footer(s, n)


def s03(prs, n):
    s = blank(prs)
    header(s, "Lecture Roadmap", "50 slides · anatomy → pathology → exam → treatment")
    sections = [
        ("A", "Foundation", "Anatomy, arches,\nsupports, definitions", TEAL),
        ("B", "Pathology", "Classification, AAFD,\nPTTD stages, peds", CORAL),
        ("C", "Workup", "History, exam,\nimaging, DDx", TEAL2),
        ("D", "Treatment", "Nonop, surgery,\nalgorithms, cases", RGBColor(0x3D, 0x8B, 0x8E)),
        ("E", "Mastery", "Pearls, boards,\nsummary, refs", NAVY),
    ]
    for i, (letter, title, body, color) in enumerate(sections):
        x = Inches(0.4) + Inches(i * 2.55)
        rrect(s, x, CONTENT_TOP, Inches(2.4), Inches(5.1), WHITE, LINE)
        rect(s, x, CONTENT_TOP, Inches(2.4), Inches(1.0), color)
        textbox(s, x, CONTENT_TOP + Inches(0.25), Inches(2.4), Inches(0.55),
                letter, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Georgia")
        textbox(s, x + Inches(0.1), Inches(2.75), Inches(2.2), Inches(0.45),
                title, size=16, bold=True, color=color, align=PP_ALIGN.CENTER, name="Georgia")
        textbox(s, x + Inches(0.15), Inches(3.4), Inches(2.1), Inches(2.2),
                body, size=13, color=SLATE, align=PP_ALIGN.CENTER)
    footer(s, n)


def s04(prs, n):
    s = blank(prs)
    header(s, "Why Flatfoot Matters", "Common · Progressive · Reconstructable")
    grid_cards(s, [
        ("Epidemiology", [
            "Flexible flatfoot near-universal in toddlers",
            "AAFD peaks in middle-aged women",
            "PTT dysfunction = leading adult cause",
        ]),
        ("Clinical Impact", [
            "Progressive pain & deformity",
            "Loss of push-off power",
            "Late disease needs fusion",
        ]),
        ("Surgeon’s Role", [
            "Stage accurately → right operation",
            "Flexible ≠ fusion candidate",
            "Understand the cascade early",
        ]),
    ], cols=3, accent=[TEAL, CORAL, TEAL2])
    footer(s, n)


def s05(prs, n):
    s = blank(prs)
    header(s, "Foot Anatomy — The Three Arches", "Structural scaffolding of the plantigrade foot")
    arches = [
        ("Medial Longitudinal", TEAL,
         "Calcaneus → talus → navicular → cuneiforms → 1st–3rd MTs",
         "Primary arch lost in pes planus"),
        ("Lateral Longitudinal", TEAL2,
         "Calcaneus → cuboid → 4th–5th MTs",
         "Lower, more rigid; secondary support"),
        ("Transverse", CORAL,
         "Cuneiforms & cuboid across midfoot",
         "Maintains forefoot column relationships"),
    ]
    for i, (title, color, a, note) in enumerate(arches):
        y = CONTENT_TOP + Inches(i * 1.65)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(1.45), WHITE, LINE)
        rect(s, Inches(0.45), y, Inches(0.14), Inches(1.45), color)
        textbox(s, Inches(0.85), y + Inches(0.2), Inches(11.7), Inches(0.35),
                title, size=18, bold=True, color=color, name="Georgia")
        textbox(s, Inches(0.85), y + Inches(0.6), Inches(11.7), Inches(0.3),
                a, size=14, color=SLATE)
        textbox(s, Inches(0.85), y + Inches(0.95), Inches(11.7), Inches(0.3),
                note, size=13, bold=True, color=NAVY)
    footer(s, n)


def s06(prs, n):
    s = blank(prs)
    header(s, "Medial Longitudinal Arch (MLA)", "The arch that defines “flat” vs “normal”")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(3.35), WHITE, LINE)
    add_pic(s, ASSETS / "arch_comparison.png", Inches(0.7), CONTENT_TOP + Inches(0.15),
            w=Inches(11.9))
    # pearl band
    rrect(s, Inches(0.45), Inches(5.0), Inches(12.4), Inches(1.65), LIGHT, TEAL)
    textbox(s, Inches(0.7), Inches(5.15), Inches(11.9), Inches(0.35),
            "Kaplan Pearl", size=15, bold=True, color=TEAL, name="Georgia")
    bullets(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.0), [
        "Flexible flatfoot: arch reconstitutes on toe-standing / non-weightbearing",
        "Rigid flatfoot: arch stays collapsed — think tarsal coalition or arthritis",
    ], size=14, spacing=5)
    footer(s, n)


def s07(prs, n):
    s = blank(prs)
    header(s, "Keystone — Talonavicular Joint", "Where collapse begins clinically")
    grid_cards(s, [
        ("Biomechanics", [
            "Talus = keystone of the MLA",
            "Navicular suspended by spring ligament + PTT",
            "Supports fail → talar head plantarflexes",
            "Result: TN uncoverage on AP XR",
            "Clinically: abduction + lost arch",
        ]),
        ("Clinical Translation", [
            "Measure TN uncoverage % / angle",
            ">30–40% uncovering = significant",
            "Spring ligament repair often paired",
            "TN arthritis in Stage III → fusion",
            "Kaplan: don’t ignore medial column",
        ]),
    ], cols=2, accent=[TEAL, CORAL])
    footer(s, n)


def s08(prs, n):
    s = blank(prs)
    header(s, "Supports of the Medial Arch", "Static ligaments + dynamic tendons")
    two_col_compare(
        s,
        "STATIC",
        [
            "Spring ligament (calcaneonavicular)",
            "Plantar fascia",
            "Long & short plantar ligaments",
            "Bony architecture (talus–navicular)",
            "Joint capsules of TN / ST complex",
        ],
        "DYNAMIC",
        [
            "Posterior tibial tendon (key)",
            "Flexor digitorum longus",
            "Flexor hallucis longus",
            "Intrinsic foot muscles",
            "Peroneus longus (1st ray plantarflexion)",
        ],
        TEAL, CORAL,
        left_note="Static failure → TN uncoverage",
        right_note="PTT failure → unopposed eversion",
    )
    footer(s, n)


def s09(prs, n):
    s = blank(prs)
    header(s, "Posterior Tibial Tendon — Anatomy", "Primary dynamic guardian of the arch")
    items = [
        ("Origin", "Posterior tibia, fibula, interosseous membrane"),
        ("Course", "Behind medial malleolus — hypovascular zone ~1–1.5 cm distal"),
        ("Insertion", "Navicular tuberosity + slips to cuneiforms, MTs, cuboid"),
        ("Action", "Invert midfoot, adduct forefoot, lock transverse tarsal joint"),
        ("Antagonist", "Peroneus brevis — unopposed eversion when PTT fails"),
        ("Innervation", "Tibial nerve (L4–S1)"),
    ]
    for i, (t, b) in enumerate(items):
        col, row = i % 2, i // 2
        x = Inches(0.4) + Inches(col * 6.45)
        y = CONTENT_TOP + Inches(row * 1.65)
        rrect(s, x, y, Inches(6.25), Inches(1.45), WHITE, LINE)
        textbox(s, x + Inches(0.25), y + Inches(0.25), Inches(5.8), Inches(0.35),
                t, size=15, bold=True, color=TEAL, name="Georgia")
        textbox(s, x + Inches(0.25), y + Inches(0.7), Inches(5.8), Inches(0.55),
                b, size=14, color=SLATE)
    footer(s, n)


def s10(prs, n):
    s = blank(prs)
    header(s, "Spring Ligament Complex", "Static sling under the talar head")
    lines = [
        "Also called plantar calcaneonavicular ligament complex",
        "Superomedial + inferoplantar bundles cradle the talar head",
        "PTT = dynamic support; spring ligament = static support",
        "Attenuation → plantar-medial escape of talar head",
        "MRI: thickening, signal change, or discontinuity in AAFD",
        "Intra-op: inspect and repair when attenuated (Kaplan emphasis)",
        "Greater spring failure ↔ greater TN uncoverage",
        "Options: direct repair, allograft, or FDL reinforcement",
    ]
    for i, line in enumerate(lines):
        y = CONTENT_TOP + Inches(i * 0.62)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.52), WHITE, LINE)
        rect(s, Inches(0.45), y, Inches(0.1), Inches(0.52), TEAL if i % 2 == 0 else CORAL)
        textbox(s, Inches(0.8), y + Inches(0.1), Inches(11.8), Inches(0.35),
                line, size=14, color=SLATE)
    footer(s, n)


def s11(prs, n):
    s = blank(prs)
    header(s, "Definition — Pes Planus", "What we mean by “flat foot”")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(1.55), LIGHT, TEAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.35), Inches(11.9), Inches(1.0),
            "Pes planus = loss of the medial longitudinal arch on weightbearing, "
            "often with hindfoot valgus and forefoot abduction (planovalgus foot).",
            size=17, color=NAVY, name="Georgia")
    grid_cards(s, [
        ("Morphologic", ["Decreased arch height", "Heel valgus", "Forefoot abduction"]),
        ("Functional", ["Impaired windlass", "Poor push-off", "Fatigue / pain"]),
        ("Not Always Disease", ["Kids often flexible & OK", "Treat symptoms + stage", "Not XR alone"]),
    ], cols=3)
    # shift cards down - grid starts at CONTENT_TOP; need custom for definition
    footer(s, n)


def s11b(prs, n):
    """Replace s11 with better spacing - redefine properly."""
    pass


def s12(prs, n):
    s = blank(prs)
    header(s, "Epidemiology & Natural History", "Who gets flatfoot — and what happens next")
    rows = [
        ("Pediatric", "Nearly all infants flat; arch develops by ~5–6 yrs; most flexible & asymptomatic"),
        ("Adolescent", "Rigid painful flatfoot → rule out tarsal coalition"),
        ("Adult AAFD", "Middle-aged / older; ♀ > ♂; obese; diabetes; inflammatory arthritis"),
        ("Risk factors", "Obesity, HTN, diabetes, steroids, trauma, RA / seronegative SpA"),
        ("Natural history", "PTTD often progresses: tenosynovitis → flexible → rigid → ankle"),
        ("Key message", "Early Stage I–II care can delay or avoid fusion surgery"),
    ]
    for i, (lab, body) in enumerate(rows):
        row_item(s, CONTENT_TOP + Inches(i * 0.82), lab, body, TEAL if i % 2 == 0 else TEAL2)
    footer(s, n)


def s13(prs, n):
    s = blank(prs)
    header(s, "Classification Overview", "Organize before you operate")
    grid_cards(s, [
        ("By Flexibility", [
            "Flexible — arch on tip-toe / NWB",
            "Rigid — no reconstitution",
            "Drives reconstruct vs fuse",
            "Always examine both ways",
        ]),
        ("By Onset", [
            "Congenital (CVT, coalitions)",
            "Acquired (PTTD, trauma, Charcot)",
            "Ask age at onset",
            "Progressive vs lifelong",
        ]),
        ("By Age", [
            "Pediatric physiologic / pathologic",
            "Adolescent rigid → coalition",
            "Adult acquired flatfoot",
            "Different algorithms by age",
        ]),
        ("By Stage (AAFD)", [
            "I Tenosynovitis — no deformity",
            "II Flexible planovalgus",
            "III Rigid ± arthritis",
            "IV Ankle / deltoid involvement",
        ]),
    ], cols=2, accent=[TEAL, TEAL2, CORAL, NAVY],
       footer_note="Kaplan: name flexibility + stage in every assessment.")
    footer(s, n)


def s14(prs, n):
    s = blank(prs)
    header(s, "Flexible vs Rigid — The Fork in the Road", "This single distinction drives the entire plan")
    two_col_compare(
        s,
        "FLEXIBLE",
        [
            "Arch appears on toe-standing",
            "Correctable heel valgus",
            "Typical: physiologic kids, Stage II AAFD",
            "Nonop first-line often succeeds",
            "Surgery = joint-sparing reconstruction",
            "Prognosis better if treated early",
        ],
        "RIGID",
        [
            "No arch on tip-toe or NWB",
            "Fixed hindfoot valgus",
            "Think: coalition, Stage III, arthritis",
            "Orthotics may still help symptoms",
            "Surgery often requires arthrodesis",
            "Image for coalition / degeneration",
        ],
        left_note="→ Reconstruct, don’t fuse",
        right_note="→ Image early · often fuse",
    )
    footer(s, n)


def s15(prs, n):
    s = blank(prs)
    header(s, "Congenital vs Acquired", "Different biology — different playbook")
    two_col_compare(
        s,
        "Congenital / Developmental",
        [
            "Physiologic flexible flatfoot",
            "Tarsal coalition",
            "Congenital vertical talus",
            "Accessory navicular (medial pain)",
            "Skewfoot / other malformations",
        ],
        "Acquired",
        [
            "Posterior tibial tendon dysfunction",
            "Inflammatory arthritis (RA)",
            "Neuropathic (Charcot)",
            "Trauma / spring ligament tear",
            "Iatrogenic / overcorrected clubfoot",
        ],
        TEAL, CORAL,
        left_note="Ask: onset age? rigidity?",
        right_note="Ask: progression? systemic disease?",
    )
    footer(s, n)


def s16(prs, n):
    s = blank(prs)
    header(s, "Pediatric Flatfoot Spectrum", "Most kids need reassurance — some need CT")
    cards_data = [
        ("Physiologic Flexible", TEAL, [
            "Most common", "Asymptomatic", "Arch on tip-toe", "Observe / educate",
        ]),
        ("Flexible Symptomatic", TEAL2, [
            "Activity pain", "Arch reconstitutes", "Orthotics / PT", "Rarely surgery",
        ]),
        ("Rigid (Coalition)", CORAL, [
            "Painful, stiff", "Peroneal spasm", "Calcaneonavicular", "Talocalcaneal",
        ]),
        ("Congenital Vertical Talus", CORAL2, [
            "Rigid rocker-bottom", "Irreducible equinus", "Needs early surgery", "≠ flexible FF",
        ]),
    ]
    for i, (title, color, lines) in enumerate(cards_data):
        x = Inches(0.35) + Inches(i * 3.2)
        rrect(s, x, CONTENT_TOP, Inches(3.05), Inches(5.15), WHITE, color)
        rect(s, x, CONTENT_TOP, Inches(3.05), Inches(1.0), color)
        textbox(s, x + Inches(0.12), CONTENT_TOP + Inches(0.25), Inches(2.8), Inches(0.6),
                title, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bullets(s, x + Inches(0.2), CONTENT_TOP + Inches(1.25), Inches(2.65), Inches(3.6),
                lines, size=14, spacing=10)
    footer(s, n)


def s17(prs, n):
    s = blank(prs)
    header(s, "Physiologic Flexible Flatfoot", "Most common “diagnosis” that isn’t a disease")
    pearls = [
        ("Reassurance", "Asymptomatic flexible flatfoot in children does not require treatment"),
        ("Exam", "Arch reconstitutes on tip-toe; heel goes into varus — normal windlass"),
        ("Orthotics", "Do not create a permanent arch; may help symptomatic activity pain"),
        ("Red flags", "Pain, rigidity, progressive deformity, neuro findings, asymmetry"),
        ("When to image", "Rigid or painful → XR ± oblique; CT if coalition suspected"),
        ("Kaplan pearl", "Avoid unnecessary surgery in flexible asymptomatic pediatric feet"),
    ]
    for i, (t, b) in enumerate(pearls):
        y = CONTENT_TOP + Inches(i * 0.82)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.7), WHITE, LINE)
        textbox(s, Inches(0.7), y + Inches(0.18), Inches(2.6), Inches(0.4),
                t, size=13, bold=True, color=TEAL)
        textbox(s, Inches(3.5), y + Inches(0.18), Inches(9.1), Inches(0.4),
                b, size=13, color=SLATE)
    footer(s, n)


def s18(prs, n):
    s = blank(prs)
    header(s, "Adult Acquired Flatfoot (AAFD)", "The core adult pathology of this lecture")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(1.5), LIGHT, TEAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.3), Inches(11.9), Inches(1.0),
            "AAFD is a progressive planovalgus deformity most often driven by PTT dysfunction, "
            "with secondary spring ligament failure and progressive joint degeneration.",
            size=16, color=NAVY, name="Georgia")
    items = [
        ("Primary driver", "PTT degeneration / tear"),
        ("Static failure", "Spring ligament complex"),
        ("Deformity triad", "Valgus + abduction + collapse"),
        ("Staging system", "Johnson & Strom (+ Myerson IV)"),
    ]
    for i, (t, b) in enumerate(items):
        x = Inches(0.45) + Inches(i * 3.2)
        rrect(s, x, Inches(3.3), Inches(3.05), Inches(3.2), WHITE, LINE)
        textbox(s, x + Inches(0.15), Inches(3.55), Inches(2.75), Inches(0.7),
                t, size=13, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.15), Inches(4.5), Inches(2.75), Inches(1.4),
                b, size=15, color=NAVY, align=PP_ALIGN.CENTER)
    footer(s, n)


def s19(prs, n):
    s = blank(prs)
    header(s, "Etiology of AAFD / PTTD", "Why the tendon fails")
    grid_cards(s, [
        ("Intrinsic", ["Age-related degeneration", "Hypovascular watershed zone", "Pre-existing flatfoot"]),
        ("Mechanical", ["Obesity / overuse", "Repetitive microtrauma", "Equinus contracture"]),
        ("Systemic", ["RA / seronegative SpA", "Diabetes, steroids", "Fluoroquinolones (rare)"]),
        ("Other", ["Trauma to spring lig / PTT", "Accessory navicular complex", "Charcot midfoot collapse"]),
    ], cols=4, accent=[TEAL, TEAL2, CORAL, NAVY])
    footer(s, n)


def s20(prs, n):
    s = blank(prs)
    header(s, "PTTD Staging — Big Picture", "Johnson & Strom I–III + Myerson Stage IV")
    stages = [
        ("I", "Tenosynovitis", ["Normal alignment", "Medial pain / swelling", "No deformity yet"],
         "Conservative\n± débridement", TEAL),
        ("II", "Flexible", ["Pes planovalgus", "Correctable", "+ Too-many-toes"],
         "FDL transfer\n+ calc osteotomy", TEAL2),
        ("III", "Rigid", ["Fixed valgus", "Midfoot arthritis", "Not correctable"],
         "Triple\narthrodesis", CORAL),
        ("IV", "Ankle", ["Valgus talar tilt", "Deltoid failure", "Ankle overload"],
         "Triple + ankle\nreconstruction", CORAL2),
    ]
    for i, (num, title, findings, tx, color) in enumerate(stages):
        x = Inches(0.35) + Inches(i * 3.25)
        rrect(s, x, CONTENT_TOP, Inches(3.1), Inches(5.15), WHITE, color)
        # number circle via rounded square
        pill(s, x + Inches(1.05), CONTENT_TOP + Inches(0.25), Inches(1.0), Inches(0.7),
             num, bg=color, size=22)
        textbox(s, x + Inches(0.15), CONTENT_TOP + Inches(1.15), Inches(2.8), Inches(0.4),
                title, size=16, bold=True, color=color, align=PP_ALIGN.CENTER, name="Georgia")
        bullets(s, x + Inches(0.25), CONTENT_TOP + Inches(1.7), Inches(2.6), Inches(1.8),
                findings, size=13, spacing=6)
        rrect(s, x + Inches(0.2), CONTENT_TOP + Inches(3.6), Inches(2.7), Inches(1.3), LIGHT, color)
        textbox(s, x + Inches(0.3), CONTENT_TOP + Inches(3.75), Inches(2.5), Inches(0.3),
                "Treatment", size=11, bold=True, color=color, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.3), CONTENT_TOP + Inches(4.1), Inches(2.5), Inches(0.7),
                tx, size=13, color=NAVY, align=PP_ALIGN.CENTER)
    footer(s, n)


def s21(prs, n):
    s = blank(prs)
    header(s, "Stage I — Tenosynovitis", "Painful tendon, preserved alignment")
    two_col_compare(
        s, "Clinical Features",
        ["Medial ankle / arch pain & swelling", "Tenderness along PTT course",
         "Alignment still normal", "Single heel rise possible (painful)",
         "XR: usually normal alignment"],
        "Management",
        ["NSAID, rest, activity modification", "Boot / AFO short term",
         "Medial posting orthosis", "PT once acute pain settles",
         "Surgery rare: tenosynovectomy"],
        TEAL, CORAL,
        left_note="No deformity yet = Stage I",
        right_note="Most Stage I never need surgery",
    )
    footer(s, n)


def s22(prs, n):
    s = blank(prs)
    header(s, "Stage II — Flexible Deformity", "The reconstructable sweet spot")
    lines = [
        "Pes planovalgus present but passively correctable",
        "Positive too-many-toes; weak / absent single heel rise inversion",
        "Often IIA (minimal abduction) vs IIB (significant TN uncoverage)",
        "Nonop: custom orthosis / AFO, PT, Achilles stretching",
        "Surgical workhorse: FDL transfer + medializing calcaneal osteotomy",
        "Add-ons: spring ligament repair, Evans LCL, Cotton osteotomy",
        "Address equinus: gastroc recession or TAL",
        "Kaplan: Stage II is where joint-preserving surgery shines",
    ]
    for i, line in enumerate(lines):
        y = CONTENT_TOP + Inches(i * 0.62)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.52), WHITE, LINE)
        pill(s, Inches(0.55), y + Inches(0.08), Inches(0.4), Inches(0.36), str(i + 1), size=11)
        textbox(s, Inches(1.15), y + Inches(0.1), Inches(11.4), Inches(0.35),
                line, size=13, color=SLATE)
    footer(s, n)


def s23(prs, n):
    s = blank(prs)
    header(s, "Stage III — Rigid Deformity", "Fixed planovalgus ± midfoot arthritis")
    grid_cards(s, [
        ("Findings", ["Fixed heel valgus", "No arch reconstitution", "ST ± TN ± CC arthritis", "Lateral impingement pain"]),
        ("Imaging", ["Severe Meary angle", "Decreased calc pitch", "Degenerative changes", "CT for joints / coalition"]),
        ("Treatment", ["Brace for poor hosts", "Triple arthrodesis standard", "Correct all planes", "± first TMT if needed"]),
    ], cols=3, accent=[TEAL, TEAL2, CORAL])
    footer(s, n)


def s24(prs, n):
    s = blank(prs)
    header(s, "Stage IV — Ankle Involvement", "Myerson: talar tilt in the mortise")
    lines = [
        "Hindfoot deformity + valgus tilt of talus in ankle mortise",
        "Deltoid ligament / deep deltoid incompetence",
        "Lateral ankle overload → early ankle arthritis",
        "Flexible ankle: reconstruct deltoid ± hindfoot correction",
        "Rigid / arthritic ankle: TTC fusion or pantalar fusion",
        "Select cases: ankle replacement + hindfoot realignment (specialized)",
        "Kaplan pearl: always scrutinize the ankle mortise on every AAFD series",
    ]
    for i, line in enumerate(lines):
        y = CONTENT_TOP + Inches(i * 0.7)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.58), WHITE, LINE)
        rect(s, Inches(0.45), y, Inches(0.1), Inches(0.58), CORAL if i >= 4 else TEAL)
        textbox(s, Inches(0.8), y + Inches(0.12), Inches(11.8), Inches(0.4),
                line, size=14, color=SLATE)
    footer(s, n)


def s25(prs, n):
    s = blank(prs)
    header(s, "Pathomechanics Cascade", "How a tendon problem becomes a foot shape problem")
    steps = [
        ("1", "PTT failure", "Loss of dynamic\ninversion"),
        ("2", "Spring lig.", "TN joint\nuncoverage"),
        ("3", "Heel valgus", "Calcaneus\neverts"),
        ("4", "Abduction", "Too-many-toes\nsign"),
        ("5", "Arch collapse", "Meary angle\nincreases"),
        ("6", "Arthritis", "ST / TN / ankle\ninvolvement"),
    ]
    for i, (num, title, body) in enumerate(steps):
        x = Inches(0.35) + Inches(i * 2.15)
        rrect(s, x, CONTENT_TOP, Inches(2.0), Inches(3.3), WHITE, LINE)
        pill(s, x + Inches(0.55), CONTENT_TOP + Inches(0.25), Inches(0.9), Inches(0.65),
             num, bg=TEAL if i < 4 else CORAL, size=18)
        textbox(s, x + Inches(0.1), CONTENT_TOP + Inches(1.15), Inches(1.8), Inches(0.7),
                title, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.1), CONTENT_TOP + Inches(2.0), Inches(1.8), Inches(1.0),
                body, size=12, color=MUTED, align=PP_ALIGN.CENTER)
        if i < 5:
            textbox(s, x + Inches(1.85), CONTENT_TOP + Inches(1.4), Inches(0.35), Inches(0.4),
                    "›", size=20, bold=True, color=MUTED, align=PP_ALIGN.CENTER)
    rrect(s, Inches(0.45), Inches(5.05), Inches(12.4), Inches(1.55), LIGHT, TEAL)
    textbox(s, Inches(0.7), Inches(5.2), Inches(11.9), Inches(0.3),
            "Biomechanical Pearl", size=14, bold=True, color=TEAL, name="Georgia")
    textbox(s, Inches(0.7), Inches(5.6), Inches(11.9), Inches(0.8),
            "Once the hindfoot is in valgus, the Achilles becomes an everting force. "
            "That is why gastroc recession / TAL is commonly added to Stage II reconstructions.",
            size=13, color=SLATE)
    footer(s, n)


def s26(prs, n):
    s = blank(prs)
    header(s, "Clinical Presentation", "What the patient tells you")
    grid_cards(s, [
        ("Medial pain", ["Along PTT course", "Worse with activity", "Swelling behind malleolus"]),
        ("Lateral pain", ["Subfibular impingement", "Advanced valgus", "“Lateral overload”"]),
        ("Fatigue", ["Aching after standing", "Loss of endurance", "Difficulty on uneven ground"]),
        ("Deformity", ["“Fallen arch” noticed", "Uneven shoe wear", "Progressive shape change"]),
        ("Function", ["Hard stairs / push-off", "Single-leg support weak", "Gait inefficiency"]),
        ("Late disease", ["Diffuse arthritis pain", "Ankle symptoms (IV)", "Brace dependence"]),
    ], cols=3, accent=[TEAL, CORAL, TEAL2, TEAL, CORAL, NAVY])
    footer(s, n)


def s27(prs, n):
    s = blank(prs)
    header(s, "History Taking Checklist", "Don’t miss systemic and progressive clues")
    qs = [
        "Onset & progression — sudden vs insidious? weeks vs years?",
        "Prior trauma, steroid injections, fluoroquinolone use?",
        "Inflammatory disease, diabetes, neuropathy?",
        "Prior orthotics / PT / injections — what helped?",
        "Activity goals, smoking, BMI, bone health",
        "Shoe wear pattern; difficulty with heel rise",
        "Pediatric history if young adult — coalition?",
        "Red flags: night pain, fever, open wounds, rapid Charcot-like collapse",
    ]
    for i, q in enumerate(qs):
        y = CONTENT_TOP + Inches(i * 0.62)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.52), WHITE, LINE)
        pill(s, Inches(0.55), y + Inches(0.08), Inches(0.4), Inches(0.36), str(i + 1), size=11)
        textbox(s, Inches(1.15), y + Inches(0.1), Inches(11.4), Inches(0.35), q, size=13, color=SLATE)
    footer(s, n)


def s28(prs, n):
    s = blank(prs)
    header(s, "Physical Examination Framework", "Look · Feel · Move · Special tests · NV")
    grid_cards(s, [
        ("1. Inspection", ["Standing alignment", "Arch height", "Heel valgus", "Too-many-toes", "Shoe wear"]),
        ("2. Palpation", ["PTT course", "Navicular", "Sinus tarsi", "Lateral gutter", "Achilles"]),
        ("3. Motion", ["Ankle DF / PF", "ST inv / ever", "TN mobility", "Equinus test", "Rigidity check"]),
        ("4. Function", ["Single heel rise", "Double heel rise", "Tip-toe arch", "Gait", "Strength grade"]),
    ], cols=4, accent=TEAL)
    footer(s, n)


def s29(prs, n):
    s = blank(prs)
    header(s, "High-Yield Special Tests", "Too-many-toes & single-limb heel rise")
    # Left — Too many toes
    rrect(s, Inches(0.4), CONTENT_TOP, Inches(6.1), Inches(5.15), WHITE, TEAL)
    rect(s, Inches(0.4), CONTENT_TOP, Inches(6.1), Inches(0.65), TEAL)
    textbox(s, Inches(0.4), CONTENT_TOP + Inches(0.15), Inches(6.1), Inches(0.4),
            "“Too-Many-Toes” Sign", size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Georgia")
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.9), Inches(5.5), Inches(0.35),
            "How to perform", size=13, bold=True, color=TEAL)
    bullets(s, Inches(0.7), CONTENT_TOP + Inches(1.3), Inches(5.5), Inches(1.3), [
        "Patient stands, examiner behind",
        "Look at how many toes show laterally",
        "Compare to contralateral foot",
    ], size=13, spacing=5)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(2.7), Inches(5.5), Inches(0.35),
            "Positive finding", size=13, bold=True, color=TEAL)
    bullets(s, Inches(0.7), CONTENT_TOP + Inches(3.1), Inches(5.5), Inches(1.0), [
        "Extra lateral toes visible from behind",
        "Means forefoot abduction / TN uncoverage",
    ], size=13, spacing=5)
    rrect(s, Inches(0.7), CONTENT_TOP + Inches(4.25), Inches(5.5), Inches(0.65), LIGHT, TEAL)
    textbox(s, Inches(0.85), CONTENT_TOP + Inches(4.4), Inches(5.2), Inches(0.4),
            "Classic in Stage II–III AAFD", size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # Right — Heel rise
    rrect(s, Inches(6.8), CONTENT_TOP, Inches(6.1), Inches(5.15), WHITE, CORAL)
    rect(s, Inches(6.8), CONTENT_TOP, Inches(6.1), Inches(0.65), CORAL)
    textbox(s, Inches(6.8), CONTENT_TOP + Inches(0.15), Inches(6.1), Inches(0.4),
            "Single-Limb Heel Rise", size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Georgia")
    textbox(s, Inches(7.1), CONTENT_TOP + Inches(0.9), Inches(5.5), Inches(0.35),
            "How to perform", size=13, bold=True, color=CORAL)
    bullets(s, Inches(7.1), CONTENT_TOP + Inches(1.3), Inches(5.5), Inches(1.3), [
        "Patient rises on one forefoot",
        "Watch heel for inversion (varus)",
        "Compare strength & ability side-to-side",
    ], size=13, spacing=5)
    textbox(s, Inches(7.1), CONTENT_TOP + Inches(2.7), Inches(5.5), Inches(0.35),
            "Interpretation", size=13, bold=True, color=CORAL)
    bullets(s, Inches(7.1), CONTENT_TOP + Inches(3.1), Inches(5.5), Inches(1.0), [
        "Normal: heel inverts via intact PTT",
        "PTTD: weak/absent inversion or can’t rise",
    ], size=13, spacing=5)
    rrect(s, Inches(7.1), CONTENT_TOP + Inches(4.25), Inches(5.5), Inches(0.65), RGBColor(0xFD, 0xF0, 0xEB), CORAL)
    textbox(s, Inches(7.25), CONTENT_TOP + Inches(4.4), Inches(5.2), Inches(0.4),
            "Key functional test for PTT integrity", size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    footer(s, n)


def s30(prs, n):
    s = blank(prs)
    header(s, "More Exam Pearls", "Details that change staging")
    pearls = [
        ("Silfverskiöld", "Gastroc vs Achilles equinus — guides recession vs TAL"),
        ("Passive correction", "Heel to neutral? If no → Stage III territory"),
        ("First ray mobility", "Residual FF varus after heel correction → Cotton candidate"),
        ("Peroneal spasm", "Rigid painful adolescent → coalition until proven otherwise"),
        ("Neurovascular", "Protect against missed Charcot / neuropathy"),
        ("Opposite foot", "Compare; bilateral disease common, staging may differ"),
    ]
    for i, (t, b) in enumerate(pearls):
        y = CONTENT_TOP + Inches(i * 0.82)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.7), WHITE, LINE)
        textbox(s, Inches(0.7), y + Inches(0.18), Inches(3.2), Inches(0.4),
                t, size=13, bold=True, color=CORAL if i % 2 else TEAL)
        textbox(s, Inches(4.1), y + Inches(0.18), Inches(8.5), Inches(0.4),
                b, size=13, color=SLATE)
    footer(s, n)


def s31(prs, n):
    s = blank(prs)
    header(s, "Imaging — Weightbearing Radiographs", "Standing films are mandatory")
    grid_cards(s, [
        ("AP foot", ["TN coverage", "Forefoot abduction", "Midfoot arthritis"]),
        ("Lateral foot", ["Meary’s angle", "Calcaneal pitch", "Cyma line / TN sag"]),
        ("Hindfoot alignment", ["Saltzman view", "Quantifies heel valgus", "Guides osteotomy"]),
        ("Ankle series", ["Mortise tilt (Stage IV)", "Ankle arthritis", "Deltoid competence clues"]),
        ("Oblique foot", ["CN coalition", "Anteater nose sign", "Adolescent workup"]),
        ("Compare priors", ["Progression over time", "Guides urgency", "Post-op baseline"]),
    ], cols=3, accent=TEAL)
    footer(s, n)


def s32(prs, n):
    s = blank(prs)
    header(s, "Key Radiographic Measurements", "Numbers you quote on rounds")
    headers = ["Parameter", "View", "Normal", "Flatfoot Finding"]
    rows = [
        ["Meary’s angle", "AP / Lat talar–1st MT", "0°", ">4° plantar convex = planus"],
        ["Calcaneal pitch", "Lateral", "20–30°", "Decreased in flatfoot"],
        ["TN uncoverage", "AP", "<30% uncover", "Increased → abduction"],
        ["Hindfoot alignment", "Saltzman / axial", "0–5° valgus", "Increased heel valgus"],
        ["Talocalcaneal (Kite)", "AP / Lat", "15–35°", "Increased in planovalgus"],
    ]
    # header row
    y0 = CONTENT_TOP
    rrect(s, Inches(0.4), y0, Inches(12.5), Inches(0.55), TEAL)
    widths = [Inches(3.2), Inches(3.2), Inches(2.4), Inches(3.7)]
    x = Inches(0.55)
    for hdg, w in zip(headers, widths):
        textbox(s, x, y0 + Inches(0.12), w, Inches(0.35), hdg, size=13, bold=True, color=WHITE)
        x += w
    for r, row in enumerate(rows):
        y = y0 + Inches(0.65) + Inches(r * 0.9)
        bgc = WHITE if r % 2 == 0 else LIGHT
        rrect(s, Inches(0.4), y, Inches(12.5), Inches(0.8), bgc, LINE)
        x = Inches(0.55)
        for j, (cell, w) in enumerate(zip(row, widths)):
            textbox(s, x, y + Inches(0.22), w, Inches(0.4), cell, size=13,
                    bold=(j == 0), color=NAVY if j == 0 else SLATE)
            x += w
    footer(s, n)


def s33(prs, n):
    s = blank(prs)
    header(s, "MRI · CT · Ultrasound", "When plain films aren’t enough")
    grid_cards(s, [
        ("MRI", ["PTT signal / tears", "Spring ligament", "Synovitis", "Early arthritis", "Deltoid integrity"]),
        ("CT", ["Tarsal coalition", "Joint degeneration map", "Osteotomy planning", "Subtalar morphology", "Fracture sequelae"]),
        ("Ultrasound", ["Dynamic PTT exam", "Tenosynovial fluid", "Accessible / cheap", "Operator dependent", "Limited bony detail"]),
    ], cols=3, accent=[TEAL, CORAL, TEAL2])
    footer(s, n)


def s34(prs, n):
    s = blank(prs)
    header(s, "Differential Diagnosis", "Not every flat painful foot is PTTD")
    ddx = [
        ("Tarsal coalition", "Rigid adolescent; peroneal spasm; CT"),
        ("Accessory navicular", "Medial prominence pain; may coexist"),
        ("Spring ligament tear", "Acute medial trauma; MRI"),
        ("Charcot arthropathy", "Neuropathy; fragmentation; warmth"),
        ("RA / inflammatory", "Polyarticular; multi-tendon synovitis"),
        ("Lisfranc injury", "Trauma; midfoot instability / arthritis"),
        ("Peroneal pathology", "Lateral symptoms; opposite tendon"),
        ("Subtalar arthritis", "Post-calcaneus fracture history"),
        ("Congenital vertical talus", "Rigid rocker-bottom from infancy"),
        ("Neuromuscular", "CP, myelomeningocele — tone & strength"),
    ]
    for i, (t, b) in enumerate(ddx):
        col, row = i % 2, i // 2
        x = Inches(0.4) + Inches(col * 6.45)
        y = CONTENT_TOP + Inches(row * 0.95)
        rrect(s, x, y, Inches(6.25), Inches(0.82), WHITE, LINE)
        textbox(s, x + Inches(0.2), y + Inches(0.2), Inches(2.5), Inches(0.45),
                t, size=12, bold=True, color=TEAL)
        textbox(s, x + Inches(2.8), y + Inches(0.2), Inches(3.2), Inches(0.45),
                b, size=12, color=SLATE)
    footer(s, n)


def s35(prs, n):
    s = blank(prs)
    header(s, "Conservative Treatment Overview", "First-line for Stage I–II and poor surgical hosts")
    layers = [
        (Inches(0.8), Inches(11.7), TEAL, "1   NSAIDs  ·  Rest  ·  Weight optimization"),
        (Inches(1.5), Inches(10.3), TEAL2, "2   Custom orthotics / brace (AFO, UCBL)"),
        (Inches(2.2), Inches(8.9), RGBColor(0x46, 0x96, 0x8C), "3   PT — strengthen PTT / Achilles stretch"),
        (Inches(2.9), Inches(7.5), CORAL, "4   Operative reconstruction if refractory"),
    ]
    y = CONTENT_TOP
    for left, width, color, text in layers:
        rrect(s, left, y, width, Inches(0.95), color)
        textbox(s, left, y + Inches(0.28), width, Inches(0.45),
                text, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        y += Inches(1.1)
    textbox(s, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.4),
            "Kaplan: exhaust nonop care in flexible deformity before osteotomy + tendon transfer.",
            size=12, color=MUTED, align=PP_ALIGN.CENTER)
    footer(s, n)


def s36(prs, n):
    s = blank(prs)
    header(s, "Orthotics & Footwear", "Brace the deformity you can’t yet reconstruct")
    items = [
        ("Medial posting", "Supports MLA; reduces PTT strain in mild disease"),
        ("UCBL orthosis", "Deeper heel cup; controls flexible heel valgus"),
        ("AFO / Arizona", "Advanced flexible deformity; limits ST motion"),
        ("Footwear", "Stable heel counter, wide toe box; avoid floppy shoes"),
        ("Goal", "Pain relief + slow progression — not anatomic “cure”"),
        ("Compliance", "Counsel that braces work only when worn"),
    ]
    for i, (t, b) in enumerate(items):
        y = CONTENT_TOP + Inches(i * 0.82)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.7), WHITE, LINE)
        textbox(s, Inches(0.7), y + Inches(0.18), Inches(3.5), Inches(0.4),
                t, size=13, bold=True, color=TEAL)
        textbox(s, Inches(4.4), y + Inches(0.18), Inches(8.2), Inches(0.4),
                b, size=13, color=SLATE)
    footer(s, n)


def s37(prs, n):
    s = blank(prs)
    header(s, "Physical Therapy Strategies", "Strengthen what remains — stretch what drives valgus")
    grid_cards(s, [
        ("Strengthen", ["PTT / invertors", "Intrinsics", "Hip abductors", "Core for gait"]),
        ("Stretch", ["Gastroc–soleus", "Silfverskiöld-guided", "Avoid painful ST forcing"]),
        ("Functional", ["Proprioception", "Gait retraining", "Activity pacing", "Weight referral"]),
        ("Limits", ["Won’t reverse Stage III", "Adjunct to orthosis", "Useful prehab", "Post-op protocols"]),
    ], cols=4, accent=TEAL)
    footer(s, n)


def s38(prs, n):
    s = blank(prs)
    header(s, "Surgical Indications", "When to put the scalpel in the plan")
    lines = [
        "Persistent pain after adequate nonop care (typically ≥3–6 months)",
        "Progressive deformity despite bracing",
        "Stage II flexible deformity → joint-preserving reconstruction",
        "Stage III rigid symptomatic deformity → realignment arthrodesis",
        "Stage IV — reconstruct or fuse based on ankle joint status",
        "Rigid pediatric coalition failing conservative care",
        "Optimize host: nicotine cessation, BMI, glycemic control",
        "Goal: plantigrade, stable, preferably brace-free foot",
    ]
    for i, line in enumerate(lines):
        y = CONTENT_TOP + Inches(i * 0.62)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.52), WHITE, LINE)
        pill(s, Inches(0.55), y + Inches(0.08), Inches(0.4), Inches(0.36), str(i + 1), size=11)
        textbox(s, Inches(1.15), y + Inches(0.1), Inches(11.4), Inches(0.35), line, size=13, color=SLATE)
    footer(s, n)


def s39(prs, n):
    s = blank(prs)
    header(s, "Soft-Tissue Procedures", "Tendon & ligament work in Stage I–II")
    procs = [
        ("PTT débridement / tenosynovectomy", "Stage I refractory synovitis"),
        ("FDL transfer to navicular", "Workhorse dynamic replacement for PTT"),
        ("FHL transfer", "Alternative when FDL unavailable"),
        ("Spring ligament repair", "Restore static TN support"),
        ("Deltoid reconstruction", "Stage IV flexible ankle valgus"),
        ("Gastroc recession / TAL", "Correct equinus driving valgus"),
    ]
    for i, (t, b) in enumerate(procs):
        y = CONTENT_TOP + Inches(i * 0.82)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.7), WHITE, LINE)
        textbox(s, Inches(0.7), y + Inches(0.18), Inches(6.8), Inches(0.4),
                t, size=14, bold=True, color=NAVY)
        textbox(s, Inches(7.6), y + Inches(0.18), Inches(5.0), Inches(0.4),
                b, size=13, color=TEAL)
    footer(s, n)


def s40(prs, n):
    s = blank(prs)
    header(s, "Osteotomies for Flexible Flatfoot", "Bony realignment that protects soft-tissue repairs")
    osteos = [
        ("Medializing calcaneal osteotomy (MCO)", "Shifts Achilles medially; reduces valgus moment — Stage II staple"),
        ("Evans lateral column lengthening", "Corrects severe forefoot abduction / TN uncoverage (IIB)"),
        ("Cotton (medial cuneiform opening wedge)", "Plantarflexes first ray; treats residual forefoot varus"),
        ("First TMT fusion (modified Lapidus)", "When first ray unstable or arthritic"),
        ("Extra-articular osteotomies preferred", "Preserve motion in flexible disease — Kaplan principle"),
    ]
    for i, (t, b) in enumerate(osteos):
        y = CONTENT_TOP + Inches(i * 0.95)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.82), WHITE, LINE)
        rect(s, Inches(0.45), y, Inches(0.12), Inches(0.82), TEAL if i % 2 == 0 else CORAL)
        textbox(s, Inches(0.85), y + Inches(0.12), Inches(11.7), Inches(0.3),
                t, size=14, bold=True, color=TEAL)
        textbox(s, Inches(0.85), y + Inches(0.45), Inches(11.7), Inches(0.3),
                b, size=12, color=SLATE)
    footer(s, n)


def s41(prs, n):
    s = blank(prs)
    header(s, "Arthrodesis Options", "For rigidity, arthritis, and salvage")
    grid_cards(s, [
        ("Subtalar fusion", ["Isolated ST arthritis", "Selected coalitions", "Less motion loss than triple"]),
        ("Double / Triple", ["TN + ST ± CC", "Stage III classic", "Correct rotation all planes", "Nonunion risk if smoking"]),
        ("Midfoot fusions", ["TN / NC / TMT arthritis", "Often adjunctive", "Restore medial column"]),
        ("TTC / Pantalar", ["Stage IV salvage", "Neuropathic caution", "Highest morbidity"]),
    ], cols=4, accent=[TEAL, TEAL2, CORAL, CORAL2])
    footer(s, n)


def s42(prs, n):
    s = blank(prs)
    header(s, "Stage-Based Surgical Algorithm", "A Kaplan-oriented decision map")
    # Start
    pill(s, Inches(4.4), CONTENT_TOP, Inches(4.5), Inches(0.7),
         "Symptomatic FF · failed nonop", bg=TEAL, size=13)
    # Decision
    rrect(s, Inches(4.4), Inches(2.4), Inches(4.5), Inches(0.85), WHITE, TEAL)
    textbox(s, Inches(4.4), Inches(2.55), Inches(4.5), Inches(0.55),
            "Flexible? (correctable)", size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # YES branch
    rrect(s, Inches(0.4), Inches(3.6), Inches(5.8), Inches(1.5), WHITE, TEAL)
    textbox(s, Inches(0.6), Inches(3.75), Inches(5.4), Inches(0.35),
            "YES — Stage II", size=14, bold=True, color=TEAL)
    textbox(s, Inches(0.6), Inches(4.2), Inches(5.4), Inches(0.7),
            "FDL transfer + MCO  ± spring lig  ± Evans  ± Cotton  + equinus release",
            size=13, color=SLATE)

    # NO branch
    rrect(s, Inches(7.1), Inches(3.6), Inches(5.8), Inches(1.5), WHITE, CORAL)
    textbox(s, Inches(7.3), Inches(3.75), Inches(5.4), Inches(0.35),
            "NO — Stage III", size=14, bold=True, color=CORAL)
    textbox(s, Inches(7.3), Inches(4.2), Inches(5.4), Inches(0.7),
            "Triple arthrodesis (ST + TN + CC). If ankle valgus → Stage IV pathway",
            size=13, color=SLATE)

    # Pearls
    rrect(s, Inches(0.4), Inches(5.35), Inches(12.5), Inches(1.35), LIGHT, TEAL)
    textbox(s, Inches(0.65), Inches(5.5), Inches(12), Inches(0.3),
            "Adjuncts & Pearls", size=13, bold=True, color=TEAL, name="Georgia")
    textbox(s, Inches(0.65), Inches(5.9), Inches(12), Inches(0.6),
            "Never fuse a flexible foot when reconstruction can restore motion. "
            "Always correct equinus. Address residual forefoot varus with Cotton.",
            size=13, color=SLATE)
    footer(s, n)


def s43(prs, n):
    s = blank(prs)
    header(s, "Pediatric Rigid Flatfoot — Coalitions", "Resect early coalitions; fuse late arthritis")
    two_col_compare(
        s,
        "Calcaneonavicular",
        [
            "Most common overall",
            "Ages ~8–12 years",
            "Oblique XR: anteater nose",
            "Rx: resection + fat interposition",
            "Good prognosis if early",
        ],
        "Talocalcaneal",
        [
            "Middle facet most common",
            "Ages ~12–16 years",
            "Harris / CT; C-sign on lateral",
            "<50% facet: resection possible",
            "Larger / arthritic → arthrodesis",
        ],
        TEAL, CORAL,
        left_note="Resection usually succeeds",
        right_note="Size + arthritis decide surgery",
    )
    footer(s, n)


def s44(prs, n):
    s = blank(prs)
    header(s, "Complications & Pitfalls", "Know them before you consent")
    two_col_compare(
        s,
        "Intraop / Early",
        [
            "Under-/overcorrection of heel",
            "Nonunion of osteotomy or fusion",
            "FDL failure if bony deformity ignored",
            "Lateral column pain after Evans",
            "Medial wound issues / nerve irritation",
        ],
        "Late / Planning",
        [
            "Progressive arthritis after reconstruction",
            "Malunion → residual valgus",
            "TAL overlengthening → weak push-off",
            "Smoking → nonunion",
            "Missing Stage IV ankle disease",
        ],
        CORAL, TEAL,
        left_note="Correct all planes in OR",
        right_note="Counsel risk · optimize host",
    )
    footer(s, n)


def s45(prs, n):
    s = blank(prs)
    header(s, "Outcomes & Prognosis", "What to tell patients honestly")
    lines = [
        "Stage I–II nonop: many improve with bracing + PT",
        "Stage II reconstruction: high satisfaction when alignment restored",
        "FDL + MCO is the most validated combo for flexible AAFD",
        "Triple relieves pain but sacrifices hindfoot motion — counsel gait change",
        "Obesity & smoking worsen every outcome metric",
        "Goal: plantigrade painless foot — mild residual stiffness is common",
        "Kaplan: correct all planes — residual FF varus sinks reconstructions",
    ]
    for i, line in enumerate(lines):
        y = CONTENT_TOP + Inches(i * 0.7)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.58), WHITE, LINE)
        rect(s, Inches(0.45), y, Inches(0.1), Inches(0.58), TEAL)
        textbox(s, Inches(0.8), y + Inches(0.12), Inches(11.8), Inches(0.4), line, size=14, color=SLATE)
    footer(s, n)


def s46(prs, n):
    s = blank(prs)
    header(s, "Clinical Pearls — Kaplan High Yield", "Memorize these before boards & clinic")
    pearls = [
        "Flexible vs rigid is the first branch point — tip-toe & passive correction",
        "Too-many-toes = abduction; single heel rise tests PTT function",
        "Weightbearing XR only — Meary, calc pitch, TN coverage, ankle mortise",
        "Stage II → joint preserving (FDL + MCO ± LCL ± Cotton ± spring lig)",
        "Stage III → triple; Stage IV → address ankle / deltoid",
        "Always treat equinus or valgus recurs",
        "Rigid adolescent flatfoot = coalition until proven otherwise",
        "Don’t fuse a flexible foot; don’t reconstruct a rigid arthritic one",
    ]
    for i, p in enumerate(pearls):
        y = CONTENT_TOP + Inches(i * 0.62)
        pill(s, Inches(0.45), y, Inches(0.5), Inches(0.48), str(i + 1),
             bg=TEAL if i < 4 else CORAL, size=13)
        rrect(s, Inches(1.1), y, Inches(11.7), Inches(0.48), WHITE, LINE)
        textbox(s, Inches(1.3), y + Inches(0.1), Inches(11.3), Inches(0.35), p, size=13, color=NAVY)
    footer(s, n)


def s47(prs, n):
    s = blank(prs)
    header(s, "Case Vignette 1 — Stage II AAFD", "Apply the algorithm")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(2.2), LIGHT, TEAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.15), Inches(11.9), Inches(0.3),
            "Case", size=12, bold=True, color=TEAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.5), Inches(11.9), Inches(1.5),
            "52-year-old woman, 18 months medial ankle pain. BMI 34. Too-many-toes +. "
            "Unable to single heel rise. Heel valgus corrects passively. "
            "WB XR: Meary −15°, TN uncoverage 45%, ankle mortise normal. Failed 5 months AFO + PT.",
            size=14, color=NAVY, name="Georgia")
    answers = [
        "Diagnosis: Stage IIB adult acquired flatfoot (PTTD)",
        "Plan: FDL transfer + medializing calcaneal osteotomy + spring ligament repair",
        "Consider Evans LCL (high TN uncoverage); Cotton if residual FF varus",
        "Address equinus; counsel weight loss & nicotine avoidance",
    ]
    for i, a in enumerate(answers):
        y = Inches(3.95) + Inches(i * 0.65)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.55), WHITE, LINE)
        pill(s, Inches(0.55), y + Inches(0.08), Inches(0.4), Inches(0.38), str(i + 1), size=11)
        textbox(s, Inches(1.15), y + Inches(0.1), Inches(11.4), Inches(0.35), a, size=13, color=SLATE)
    footer(s, n)


def s48(prs, n):
    s = blank(prs)
    header(s, "Case Vignette 2 — Rigid Adolescent Flatfoot", "Don’t miss the coalition")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(2.2), LIGHT, CORAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.15), Inches(11.9), Inches(0.3),
            "Case", size=12, bold=True, color=CORAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.5), Inches(11.9), Inches(1.5),
            "13-year-old boy, activity pain 1 year. Rigid planovalgus, peroneal spasm, "
            "no arch on tip-toe. Lateral XR suggests C-sign; Harris view abnormal. "
            "CT: middle-facet talocalcaneal coalition ~30% of facet. No arthritis.",
            size=14, color=NAVY, name="Georgia")
    answers = [
        "Diagnosis: Symptomatic talocalcaneal coalition → rigid flatfoot",
        "Nonop trial: activity mod, NSAID, casting/boot — if fails → surgery",
        "Surgery: coalition resection + fat interposition (± calc osteotomy)",
        "Kaplan: <50% involvement without arthritis favors resection over fusion",
    ]
    for i, a in enumerate(answers):
        y = Inches(3.95) + Inches(i * 0.65)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.55), WHITE, LINE)
        pill(s, Inches(0.55), y + Inches(0.08), Inches(0.4), Inches(0.38), str(i + 1),
             bg=CORAL, size=11)
        textbox(s, Inches(1.15), y + Inches(0.1), Inches(11.4), Inches(0.35), a, size=13, color=SLATE)
    footer(s, n)


def s49(prs, n):
    s = blank(prs)
    header(s, "Board-Style Rapid Fire", "Say the answer out loud")
    qa = [
        ("Most common cause of AAFD?", "Posterior tibial tendon dysfunction"),
        ("Test for PTT function?", "Single-limb heel rise (inversion)"),
        ("Flexible Stage II surgery?", "FDL transfer + medializing calc osteotomy"),
        ("Rigid Stage III surgery?", "Triple arthrodesis"),
        ("Adolescent rigid flatfoot?", "Tarsal coalition (CN or TC)"),
        ("Anteater nose sign?", "Calcaneonavicular coalition on oblique XR"),
        ("Why add gastroc recession?", "Equinus drives recurrent valgus"),
        ("Stage IV hallmark?", "Valgus talar tilt in ankle mortise"),
    ]
    for i, (q, a) in enumerate(qa):
        col, row = i % 2, i // 2
        x = Inches(0.4) + Inches(col * 6.45)
        y = CONTENT_TOP + Inches(row * 1.2)
        rrect(s, x, y, Inches(6.25), Inches(1.05), WHITE, LINE)
        textbox(s, x + Inches(0.2), y + Inches(0.12), Inches(5.85), Inches(0.35),
                q, size=12, bold=True, color=TEAL)
        textbox(s, x + Inches(0.2), y + Inches(0.52), Inches(5.85), Inches(0.4),
                a, size=13, color=NAVY)
    footer(s, n)


def s50(prs, n):
    s = blank(prs)
    header(s, "Summary & Take-Home", "One page to remember tomorrow")
    points = [
        "Define the deformity: flexible vs rigid; pediatric vs adult; stage I–IV",
        "PTT + spring ligament failure drive the AAFD cascade",
        "Exam: too-many-toes, single heel rise, Silfverskiöld, passive correctability",
        "Image standing: Meary, calc pitch, TN coverage, ankle mortise ± CT/MRI",
        "Nonop first for flexible disease; reconstruct II; fuse III; add ankle for IV",
        "Coalitions: resect early, fuse late — never ignore rigid kids",
    ]
    for i, p in enumerate(points):
        y = CONTENT_TOP + Inches(i * 0.8)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.65), WHITE, LINE)
        pill(s, Inches(0.6), y + Inches(0.1), Inches(0.45), Inches(0.45), str(i + 1), size=13)
        textbox(s, Inches(1.3), y + Inches(0.15), Inches(11.3), Inches(0.4), p, size=14, color=NAVY)
    footer(s, n)


def fix_s11(prs, n):
    """Definition slide with proper lower cards (no overlap with quote)."""
    s = blank(prs)
    header(s, "Definition — Pes Planus", "What we mean by “flat foot”")
    rrect(s, Inches(0.45), CONTENT_TOP, Inches(12.4), Inches(1.7), LIGHT, TEAL)
    textbox(s, Inches(0.7), CONTENT_TOP + Inches(0.4), Inches(11.9), Inches(1.1),
            "Pes planus = loss of the medial longitudinal arch on weightbearing, "
            "often with hindfoot valgus and forefoot abduction (planovalgus foot).",
            size=17, color=NAVY, name="Georgia")
    items = [
        ("Morphologic", ["Decreased arch height", "Heel valgus", "Forefoot abduction"]),
        ("Functional", ["Impaired windlass", "Poor push-off", "Fatigue / pain"]),
        ("Not Always Disease", ["Kids often flexible & OK", "Treat symptoms + stage", "Not XR alone"]),
    ]
    for i, (title, lines) in enumerate(items):
        x = Inches(0.45) + Inches(i * 4.2)
        card(s, x, Inches(3.5), Inches(4.0), Inches(3.1), title, lines, TEAL if i != 1 else CORAL)
    footer(s, n)


def build():
    # regenerate compact diagrams
    import subprocess, sys
    subprocess.check_call([sys.executable, str(ROOT / "scripts" / "generate_diagrams.py")])

    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    builders = [
        s01, s02, s03, s04, s05, s06, s07, s08, s09, s10,
        fix_s11, s12, s13, s14, s15, s16, s17, s18, s19, s20,
        s21, s22, s23, s24, s25, s26, s27, s28, s29, s30,
        s31, s32, s33, s34, s35, s36, s37, s38, s39, s40,
        s41, s42, s43, s44, s45, s46, s47, s48, s49, s50,
    ]
    assert len(builders) == 50

    for i, fn in enumerate(builders, start=1):
        if fn is s01:
            fn(prs)
        else:
            fn(prs, i)

    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")
    return OUT


if __name__ == "__main__":
    build()
