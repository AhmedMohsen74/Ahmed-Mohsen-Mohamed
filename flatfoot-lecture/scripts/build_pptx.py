#!/usr/bin/env python3
"""Build a 50-slide Kaplan-based Flat Foot lecture for orthopedic trainees."""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
OUT = ROOT / "Flat_Foot_Kaplan_Orthopedic_Lecture.pptx"

# Palette — medical teal (avoid purple / cream-terracotta clichés)
NAVY = RGBColor(0x1B, 0x28, 0x38)
TEAL = RGBColor(0x0D, 0x73, 0x77)
TEAL_DARK = RGBColor(0x0A, 0x55, 0x58)
CORAL = RGBColor(0xE0, 0x7A, 0x5F)
SOFT = RGBColor(0xF7, 0xF9, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x5A, 0x6A, 0x7A)
LIGHT = RGBColor(0xE8, 0xF5, 0xF4)
SLATE = RGBColor(0x2C, 0x3E, 0x50)

W, H = Inches(13.333), Inches(7.5)  # widescreen 16:9


def rgb(shape_fill, color):
    shape_fill.solid()
    shape_fill.fore_color.rgb = color


def set_run(run, size=18, bold=False, color=NAVY, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_textbox(slide, left, top, width, height, text, size=18, bold=False, color=NAVY,
                align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._fFramePr = None
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def add_paras(slide, left, top, width, height, lines, size=16, color=NAVY, bold_first=False,
              spacing=6, bullet=False, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(spacing)
        run = p.add_run()
        prefix = "•  " if bullet and line.strip() else ""
        run.text = prefix + line
        set_run(run, size=size, bold=(bold_first and i == 0), color=color, font=font)
    return box


def rect(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    rgb(sh.fill, fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
    return sh


def round_rect(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    rgb(sh.fill, fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.5)
    return sh


def accent_bar(slide, left=Inches(0), top=Inches(0), width=Inches(0.18), height=H):
    return rect(slide, left, top, width, height, TEAL)


def footer(slide, num, total=50):
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(8), Inches(0.35),
                "Flat Foot — Kaplan Orthopedic Lecture", size=11, color=MUTED, font="Calibri")
    add_textbox(slide, Inches(11.2), Inches(7.05), Inches(1.8), Inches(0.35),
                f"{num}  /  {total}", size=11, color=MUTED, align=PP_ALIGN.RIGHT)


def section_header(slide, title, subtitle=None):
    accent_bar(slide)
    rect(slide, Inches(0), Inches(0), W, Inches(1.35), NAVY)
    add_textbox(slide, Inches(0.55), Inches(0.35), Inches(12), Inches(0.55),
                title, size=32, bold=True, color=WHITE, font="Georgia")
    if subtitle:
        add_textbox(slide, Inches(0.55), Inches(0.9), Inches(12), Inches(0.35),
                    subtitle, size=14, color=LIGHT, font="Calibri")


def card(slide, left, top, width, height, title, body_lines, accent=TEAL):
    round_rect(slide, left, top, width, height, WHITE, LIGHT)
    rect(slide, left, top, width, Inches(0.08), accent)
    add_textbox(slide, left + Inches(0.2), top + Inches(0.2), width - Inches(0.3), Inches(0.4),
                title, size=16, bold=True, color=accent, font="Georgia")
    add_paras(slide, left + Inches(0.2), top + Inches(0.65), width - Inches(0.35), height - Inches(0.8),
              body_lines, size=14, color=SLATE, bullet=True, spacing=4)


def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


def add_image(slide, path, left, top, width=None, height=None):
    if not Path(path).exists():
        return None
    kwargs = dict(left=left, top=top)
    if width:
        kwargs["width"] = width
    if height:
        kwargs["height"] = height
    return slide.shapes.add_picture(str(path), **kwargs)


# ───────────────────────── SLIDE BUILDERS ─────────────────────────

def s01_title(prs):
    s = blank_slide(prs)
    rect(s, 0, 0, W, H, NAVY)
    add_image(s, ASSETS / "title_hero.png", 0, 0, width=W)
    # Overlay text in case image text is small on some viewers — already in image
    add_textbox(s, Inches(0.6), Inches(6.5), Inches(8), Inches(0.4),
                "Orthopedic Surgery  ·  Foot & Ankle Module", size=14, color=LIGHT)
    return s


def s02_objectives(prs, n):
    s = blank_slide(prs)
    section_header(s, "Learning Objectives", "By the end of this lecture you will be able to…")
    objs = [
        ("01", "Define pes planus and distinguish flexible vs rigid deformity"),
        ("02", "Describe anatomy & biomechanics of the medial longitudinal arch"),
        ("03", "Classify adult acquired flatfoot (Johnson–Strom / Myerson)"),
        ("04", "Perform a focused flatfoot exam and interpret key signs"),
        ("05", "Select imaging and radiographic measurements that matter"),
        ("06", "Build a stage-based nonoperative and surgical treatment plan"),
        ("07", "Recognize pediatric coalitions and when to operate"),
        ("08", "Recall Kaplan high-yield pearls for boards & OR"),
    ]
    for i, (num, text) in enumerate(objs):
        y = Inches(1.6) + Inches(i * 0.62)
        round_rect(s, Inches(0.55), y, Inches(0.7), Inches(0.5), TEAL)
        add_textbox(s, Inches(0.62), y + Inches(0.08), Inches(0.6), Inches(0.4),
                    num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.45), y + Inches(0.08), Inches(11), Inches(0.4),
                    text, size=17, color=NAVY)
    footer(s, n)
    return s


def s03_roadmap(prs, n):
    s = blank_slide(prs)
    section_header(s, "Lecture Roadmap", "50 slides · anatomy → pathology → exam → treatment")
    sections = [
        ("A", "Foundation", "Anatomy, arches,\nsupports, definitions", TEAL),
        ("B", "Pathology", "Classification, AAFD,\nPTTD stages, peds", CORAL),
        ("C", "Workup", "History, exam,\nimaging, DDx", TEAL_DARK),
        ("D", "Treatment", "Nonop, surgery,\nalgorithms, cases", RGBColor(0x3D, 0x8B, 0x8E)),
        ("E", "Mastery", "Pearls, boards,\nsummary, refs", NAVY),
    ]
    for i, (letter, title, body, color) in enumerate(sections):
        x = Inches(0.45) + Inches(i * 2.55)
        round_rect(s, x, Inches(2.0), Inches(2.35), Inches(4.2), WHITE, LIGHT)
        rect(s, x, Inches(2.0), Inches(2.35), Inches(1.1), color)
        add_textbox(s, x, Inches(2.15), Inches(2.35), Inches(0.45),
                    letter, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
        add_textbox(s, x + Inches(0.1), Inches(3.3), Inches(2.15), Inches(0.5),
                    title, size=18, bold=True, color=color, align=PP_ALIGN.CENTER, font="Georgia")
        add_textbox(s, x + Inches(0.15), Inches(4.0), Inches(2.05), Inches(1.8),
                    body, size=14, color=SLATE, align=PP_ALIGN.CENTER)
    footer(s, n)
    return s


def s04_why(prs, n):
    s = blank_slide(prs)
    section_header(s, "Why Flatfoot Matters", "Common • Progressive • Reconstructable")
    points = [
        ("Epidemiology", ["Flexible flatfoot is near-universal in toddlers",
                          "Adult acquired flatfoot peaks in middle-aged women",
                          "PTT dysfunction is the leading cause of AAFD"]),
        ("Clinical Impact", ["Progressive pain, deformity, and arthritis",
                             "Loss of push-off power & walking endurance",
                             "Late disease needs fusion — early care preserves joints"]),
        ("Surgeon’s Role", ["Stage accurately → choose the right operation",
                            "Flexible ≠ fusion candidate",
                            "Kaplan emphasis: understand the cascade"]),
    ]
    for i, (title, lines) in enumerate(points):
        card(s, Inches(0.45) + Inches(i * 4.2), Inches(1.8), Inches(3.95), Inches(4.5),
             title, lines, accent=TEAL if i != 1 else CORAL)
    footer(s, n)
    return s


def s05_anatomy_overview(prs, n):
    s = blank_slide(prs)
    section_header(s, "Foot Anatomy — The Three Arches", "Structural scaffolding of the plantigrade foot")
    arches = [
        ("Medial Longitudinal", "Calcaneus → talus → navicular → cuneiforms → 1st–3rd MTs",
         "Primary arch lost in pes planus", TEAL),
        ("Lateral Longitudinal", "Calcaneus → cuboid → 4th–5th MTs",
         "Lower, more rigid; secondary support", TEAL_DARK),
        ("Transverse", "Cuneiforms & cuboid across midfoot",
         "Maintains forefoot column relationships", CORAL),
    ]
    for i, (t, a, note, c) in enumerate(arches):
        y = Inches(1.7) + Inches(i * 1.6)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(1.4), WHITE, LIGHT)
        rect(s, Inches(0.5), y, Inches(0.15), Inches(1.4), c)
        add_textbox(s, Inches(0.9), y + Inches(0.2), Inches(11.5), Inches(0.4),
                    t, size=20, bold=True, color=c, font="Georgia")
        add_textbox(s, Inches(0.9), y + Inches(0.6), Inches(11.5), Inches(0.35),
                    a, size=15, color=SLATE)
        add_textbox(s, Inches(0.9), y + Inches(0.95), Inches(11.5), Inches(0.3),
                    note, size=14, bold=True, color=NAVY)
    footer(s, n)
    return s


def s06_mla(prs, n):
    s = blank_slide(prs)
    section_header(s, "Medial Longitudinal Arch (MLA)", "The arch that defines “flat” vs “normal”")
    add_image(s, ASSETS / "arch_comparison.png", Inches(0.4), Inches(1.5), width=Inches(12.5))
    footer(s, n)
    return s


def s07_keystone(prs, n):
    s = blank_slide(prs)
    section_header(s, "Keystone Concept — Talonavicular Joint", "Where collapse begins clinically")
    left = [
        "Talus is the “keystone” of the MLA",
        "Navicular is suspended by spring ligament + PTT",
        "When supports fail → talar head plantarflexes & adducts relative to navicular",
        "Result: TN uncoverage on AP radiograph",
        "Clinically: midfoot abduction + loss of arch",
    ]
    right = [
        "Kaplan focus: measure TN uncoverage angle / percentage",
        ">30–40% uncovering = significant abduction",
        "Spring ligament repair often paired with PTT reconstruction",
        "TN arthritis in Stage III → fusion territory",
    ]
    card(s, Inches(0.45), Inches(1.7), Inches(6.1), Inches(4.8), "Biomechanics", left, TEAL)
    card(s, Inches(6.8), Inches(1.7), Inches(6.1), Inches(4.8), "Clinical Translation", right, CORAL)
    footer(s, n)
    return s


def s08_supports(prs, n):
    s = blank_slide(prs)
    section_header(s, "Supports of the Medial Arch", "Static ligaments + dynamic tendons")
    add_image(s, ASSETS / "arch_supports.png", Inches(0.5), Inches(1.45), width=Inches(12.3))
    footer(s, n)
    return s


def s09_ptt_anatomy(prs, n):
    s = blank_slide(prs)
    section_header(s, "Posterior Tibial Tendon — Anatomy", "The primary dynamic guardian of the arch")
    items = [
        ("Origin", "Posterior tibia, fibula, interosseous membrane"),
        ("Course", "Posterior to medial malleolus (hypovascular zone ~1–1.5 cm distal)"),
        ("Insertion", "Navicular tuberosity + expansive slips to cuneiforms, MTs, cuboid"),
        ("Action", "Invert midfoot, adduct forefoot, lock transverse tarsal joint, support MLA"),
        ("Antagonist", "Peroneus brevis — unopposed eversion when PTT fails"),
        ("Innervation", "Tibial nerve (L4–S1)"),
    ]
    for i, (t, b) in enumerate(items):
        col = i % 2
        row = i // 2
        x = Inches(0.45) + Inches(col * 6.4)
        y = Inches(1.65) + Inches(row * 1.6)
        round_rect(s, x, y, Inches(6.15), Inches(1.4), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.25), y + Inches(0.25), Inches(5.6), Inches(0.35),
                    t, size=16, bold=True, color=TEAL, font="Georgia")
        add_textbox(s, x + Inches(0.25), y + Inches(0.7), Inches(5.6), Inches(0.55),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s10_spring(prs, n):
    s = blank_slide(prs)
    section_header(s, "Spring Ligament Complex", "Static sling under the talar head")
    add_paras(s, Inches(0.6), Inches(1.7), Inches(12), Inches(5), [
        "Also called the plantar calcaneonavicular ligament complex",
        "Superomedial and inferoplantar bundles cradle the talar head",
        "Works with PTT: tendon = dynamic; spring ligament = static",
        "Attenuation → plantar-medial escape of talar head",
        "MRI: look for thickening, signal change, discontinuity in AAFD",
        "Intra-op: inspect and repair/reconstruct when attenuated (Kaplan emphasis)",
        "Failure of spring ligament correlates with greater TN uncoverage",
        "Reconstruction options: direct repair, allograft, or FDL transfer reinforcement",
    ], size=17, bullet=True, spacing=8)
    footer(s, n)
    return s


def s11_definition(prs, n):
    s = blank_slide(prs)
    section_header(s, "Definition — Pes Planus", "What we mean by “flat foot”")
    round_rect(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.0), LIGHT)
    add_textbox(s, Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.4),
                "Pes planus = loss of the medial longitudinal arch on weightbearing, "
                "often with hindfoot valgus and forefoot abduction (planovalgus foot).",
                size=20, color=NAVY, font="Georgia")
    cards = [
        ("Morphologic", ["Decreased arch height", "Heel valgus", "Forefoot abduction"]),
        ("Functional", ["Impaired windlass", "Poor push-off", "Fatigue / pain"]),
        ("Not Always Pathologic", ["Kids often flexible & OK", "Treat symptoms + stage", "Not radiographs alone"]),
    ]
    for i, (t, lines) in enumerate(cards):
        card(s, Inches(0.5) + Inches(i * 4.2), Inches(4.0), Inches(4.0), Inches(2.5), t, lines)
    footer(s, n)
    return s


def s12_epi(prs, n):
    s = blank_slide(prs)
    section_header(s, "Epidemiology & Natural History", "Who gets flatfoot — and what happens next")
    rows = [
        ("Pediatric", "Nearly all infants flat; arch develops by ~5–6 yrs; most flexible & asymptomatic"),
        ("Adolescent", "Rigid painful flatfoot → rule out tarsal coalition"),
        ("Adult AAFD", "Middle-aged / older; ♀ > ♂; obese; diabetics; inflammatory arthritis"),
        ("Risk factors", "Obesity, hypertension, diabetes, steroids, prior trauma, RA / seronegative"),
        ("Natural Hx", "PTTD often progressive from tenosynovitis → flexible → rigid → ankle"),
        ("Key message", "Early Stage I–II care can delay or avoid fusion surgery"),
    ]
    for i, (t, b) in enumerate(rows):
        y = Inches(1.55) + Inches(i * 0.82)
        round_rect(s, Inches(0.5), y, Inches(2.4), Inches(0.7), TEAL if i % 2 == 0 else TEAL_DARK)
        add_textbox(s, Inches(0.6), y + Inches(0.18), Inches(2.2), Inches(0.4),
                    t, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        round_rect(s, Inches(3.05), y, Inches(9.7), Inches(0.7), WHITE, LIGHT)
        add_textbox(s, Inches(3.25), y + Inches(0.18), Inches(9.3), Inches(0.4),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s13_class_overview(prs, n):
    s = blank_slide(prs)
    section_header(s, "Classification Overview", "Organize before you operate")
    matrix = [
        ("By Flexibility", ["Flexible — arch on tip-toe / NWB", "Rigid — no reconstitution"]),
        ("By Onset", ["Congenital (vertical talus, coalitions)", "Acquired (PTTD, trauma, Charcot)"]),
        ("By Age", ["Pediatric physiologic / pathologic", "Adult acquired flatfoot deformity"]),
        ("By Stage (AAFD)", ["I Tenosynovitis", "II Flexible", "III Rigid", "IV Ankle"]),
    ]
    for i, (t, lines) in enumerate(matrix):
        x = Inches(0.4) + Inches((i % 2) * 6.45)
        y = Inches(1.65) + Inches((i // 2) * 2.5)
        card(s, x, y, Inches(6.2), Inches(2.25), t, lines, TEAL if i < 2 else CORAL)
    footer(s, n)
    return s


def s14_flex_vs_rigid(prs, n):
    s = blank_slide(prs)
    section_header(s, "Flexible vs Rigid — The Fork in the Road", "This single distinction drives the entire plan")
    # two columns
    round_rect(s, Inches(0.4), Inches(1.65), Inches(6.1), Inches(4.9), WHITE, TEAL)
    rect(s, Inches(0.4), Inches(1.65), Inches(6.1), Inches(0.8), TEAL)
    add_textbox(s, Inches(0.4), Inches(1.8), Inches(6.1), Inches(0.5),
                "FLEXIBLE", size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    add_paras(s, Inches(0.7), Inches(2.7), Inches(5.5), Inches(3.5), [
        "Arch appears on toe-standing",
        "Correctable heel valgus",
        "Typical: physiologic kids, Stage II AAFD",
        "Nonop first-line often succeeds",
        "Surgery = joint-sparing reconstruction",
        "Prognosis better if treated early",
    ], size=16, bullet=True, spacing=8)

    round_rect(s, Inches(6.8), Inches(1.65), Inches(6.1), Inches(4.9), WHITE, CORAL)
    rect(s, Inches(6.8), Inches(1.65), Inches(6.1), Inches(0.8), CORAL)
    add_textbox(s, Inches(6.8), Inches(1.8), Inches(6.1), Inches(0.5),
                "RIGID", size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    add_paras(s, Inches(7.1), Inches(2.7), Inches(5.5), Inches(3.5), [
        "No arch on tip-toe or NWB",
        "Fixed hindfoot valgus",
        "Think: coalition, Stage III, arthritis",
        "Orthotics may still help symptoms",
        "Surgery often requires arthrodesis",
        "Image for coalition / degeneration",
    ], size=16, bullet=True, spacing=8)
    footer(s, n)
    return s


def s15_congenital_vs_acq(prs, n):
    s = blank_slide(prs)
    section_header(s, "Congenital vs Acquired", "Different biology — different playbook")
    add_paras(s, Inches(0.6), Inches(1.7), Inches(5.8), Inches(5), [
        "Congenital / Developmental",
        "Physiologic flexible flatfoot",
        "Tarsal coalition",
        "Congenital vertical talus (rocker-bottom)",
        "Accessory navicular (contribute to medial pain)",
        "Skewfoot / other malformations",
        "",
        "Ask: onset age? trauma? rigidity?",
    ], size=16, bullet=True, spacing=5, bold_first=True)
    add_paras(s, Inches(7.0), Inches(1.7), Inches(5.8), Inches(5), [
        "Acquired",
        "Posterior tibial tendon dysfunction",
        "Inflammatory arthritis (RA)",
        "Neuropathic (Charcot)",
        "Trauma (Lisfranc, calcaneus, spring lig tear)",
        "Iatrogenic / overcorrected clubfoot",
        "Tumor / space-occupying (rare)",
        "",
        "Ask: progression? systemic disease?",
    ], size=16, bullet=True, spacing=5, bold_first=True)
    footer(s, n)
    return s


def s16_pediatric(prs, n):
    s = blank_slide(prs)
    section_header(s, "Pediatric Flatfoot Spectrum", "Most kids need reassurance — some need CT")
    add_image(s, ASSETS / "pediatric_spectrum.png", Inches(0.5), Inches(1.45), width=Inches(12.3))
    footer(s, n)
    return s


def s17_physio_kids(prs, n):
    s = blank_slide(prs)
    section_header(s, "Physiologic Flexible Flatfoot", "The most common “diagnosis” that isn’t a disease")
    pearls = [
        ("Reassurance", "Asymptomatic flexible flatfoot in children does not require treatment"),
        ("Exam", "Arch reconstitutes on tip-toe; heel goes into varus — normal windlass"),
        ("Orthotics", "Do not “create” a permanent arch; may help symptomatic activity pain"),
        ("Red flags", "Pain, rigidity, progressive deformity, neuro findings, asymmetry"),
        ("When to image", "Rigid or painful → XR ± oblique; CT for coalition suspicion"),
        ("Kaplan pearl", "Avoid unnecessary surgery in flexible asymptomatic pediatric feet"),
    ]
    for i, (t, b) in enumerate(pearls):
        y = Inches(1.55) + Inches(i * 0.82)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), WHITE, LIGHT)
        add_textbox(s, Inches(0.75), y + Inches(0.18), Inches(2.5), Inches(0.4),
                    t, size=15, bold=True, color=TEAL)
        add_textbox(s, Inches(3.4), y + Inches(0.18), Inches(9.1), Inches(0.4),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s18_aafd_intro(prs, n):
    s = blank_slide(prs)
    section_header(s, "Adult Acquired Flatfoot Deformity (AAFD)", "The core adult pathology of this lecture")
    round_rect(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(1.6), LIGHT)
    add_textbox(s, Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.1),
                "AAFD is a progressive planovalgus deformity most often driven by posterior tibial "
                "tendon dysfunction, with secondary spring ligament failure and progressive joint degeneration.",
                size=18, color=NAVY, font="Georgia")
    triples = [
        ("Primary driver", "PTT degeneration / tear"),
        ("Static failure", "Spring ligament complex"),
        ("Deformity triad", "Valgus + abduction + collapse"),
        ("Staging system", "Johnson & Strom (+ Myerson IV)"),
    ]
    for i, (t, b) in enumerate(triples):
        x = Inches(0.5) + Inches(i * 3.2)
        round_rect(s, x, Inches(3.7), Inches(3.0), Inches(2.5), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.2), Inches(4.0), Inches(2.6), Inches(0.8),
                    t, size=14, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        add_textbox(s, x + Inches(0.2), Inches(4.8), Inches(2.6), Inches(1.0),
                    b, size=16, color=NAVY, align=PP_ALIGN.CENTER)
    footer(s, n)
    return s


def s19_etiology(prs, n):
    s = blank_slide(prs)
    section_header(s, "Etiology of AAFD / PTTD", "Why the tendon fails")
    cols = [
        ("Intrinsic", ["Age-related degeneration", "Hypovascular watershed zone", "Pre-existing flatfoot"]),
        ("Mechanical", ["Obesity / overuse", "Repetitive microtrauma", "Equinus contracture"]),
        ("Systemic", ["RA / seronegative SpA", "Diabetes, steroids", "Fluoroquinolones (rare)"]),
        ("Other", ["Trauma to spring lig / PTT", "Accessory navicular pain complex", "Charcot midfoot collapse"]),
    ]
    for i, (t, lines) in enumerate(cols):
        card(s, Inches(0.4) + Inches(i * 3.25), Inches(1.7), Inches(3.1), Inches(4.7), t, lines,
             TEAL if i % 2 == 0 else CORAL)
    footer(s, n)
    return s


def s20_stages_overview(prs, n):
    s = blank_slide(prs)
    section_header(s, "PTTD Staging — Big Picture", "Johnson & Strom Stages I–III + Myerson Stage IV")
    add_image(s, ASSETS / "pttd_stages.png", Inches(0.15), Inches(1.35), width=Inches(13.0))
    footer(s, n)
    return s


def s21_stage1(prs, n):
    s = blank_slide(prs)
    section_header(s, "Stage I — Tenosynovitis without Deformity", "Painful tendon, preserved alignment")
    left = ["Medial ankle / arch pain & swelling", "Tenderness along PTT course",
            "Alignment still normal", "Single heel rise possible (may be painful)",
            "XR: usually normal alignment"]
    right = ["NSAID, rest, activity mod", "Boot / AFO short term", "Medial posting orthosis",
             "PT once acute pain settles", "Surgery rare: tenosynovectomy ± debridement"]
    card(s, Inches(0.45), Inches(1.7), Inches(6.1), Inches(4.6), "Clinical Features", left, TEAL)
    card(s, Inches(6.8), Inches(1.7), Inches(6.1), Inches(4.6), "Management", right, CORAL)
    footer(s, n)
    return s


def s22_stage2(prs, n):
    s = blank_slide(prs)
    section_header(s, "Stage II — Flexible Deformity", "The reconstructable sweet spot")
    add_paras(s, Inches(0.6), Inches(1.6), Inches(12), Inches(5.2), [
        "Pes planovalgus present but passively correctable",
        "Positive too-many-toes sign; weak or absent single heel rise inversion",
        "Often subdivided: IIA (minimal abduction) vs IIB (significant TN uncoverage)",
        "Nonop: custom orthosis / AFO, PT, Achilles stretching",
        "Surgical workhorse: FDL transfer to navicular + medializing calcaneal osteotomy (MCO)",
        "Add-ons: spring ligament repair, lateral column lengthening (Evans), Cotton osteotomy",
        "Address equinus: gastroc recession or TAL",
        "Kaplan: Stage II is where joint-preserving surgery shines — don’t wait for rigidity",
    ], size=16, bullet=True, spacing=7)
    footer(s, n)
    return s


def s23_stage3(prs, n):
    s = blank_slide(prs)
    section_header(s, "Stage III — Rigid Deformity", "Fixed planovalgus ± midfoot arthritis")
    cards_data = [
        ("Findings", ["Fixed heel valgus", "No arch reconstitution", "ST ± TN ± CC arthritis", "Painful lateral impingement"]),
        ("Imaging", ["WB XR: severe Meary", "Decreased calc pitch", "Degenerative changes", "CT for coalition / joints"]),
        ("Treatment", ["Nonop brace for poor hosts", "Triple arthrodesis gold standard", "Correct rotation carefully", "Consider first TMT if needed"]),
    ]
    for i, (t, lines) in enumerate(cards_data):
        card(s, Inches(0.4) + Inches(i * 4.25), Inches(1.7), Inches(4.05), Inches(4.7), t, lines,
             TEAL if i != 2 else CORAL)
    footer(s, n)
    return s


def s24_stage4(prs, n):
    s = blank_slide(prs)
    section_header(s, "Stage IV — Ankle Involvement", "Myerson addition: talar tilt in the mortise")
    add_paras(s, Inches(0.6), Inches(1.7), Inches(12), Inches(5), [
        "Hindfoot deformity + valgus tilt of talus in ankle mortise",
        "Deltoid ligament / deep deltoid incompetence",
        "Lateral ankle joint overload → early ankle arthritis",
        "Treatment individualized and complex:",
        "    • Flexible ankle deformity: reconstruct deltoid ± hindfoot correction",
        "    • Rigid / arthritic ankle: tibiotalocalcaneal (TTC) fusion or pantalar fusion",
        "    • Select cases: ankle replacement + hindfoot realignment (specialized)",
        "Kaplan pearl: always scrutinize the ankle mortise on every AAFD series",
    ], size=17, bullet=True, spacing=7)
    footer(s, n)
    return s


def s25_pathomech(prs, n):
    s = blank_slide(prs)
    section_header(s, "Pathomechanics Cascade", "How a tendon problem becomes a foot shape problem")
    add_image(s, ASSETS / "pathomechanics.png", Inches(0.5), Inches(1.45), width=Inches(12.3))
    footer(s, n)
    return s


def s26_clinical(prs, n):
    s = blank_slide(prs)
    section_header(s, "Clinical Presentation", "What the patient tells you")
    items = [
        ("Medial pain", "Along PTT, worse with activity; swelling behind medial malleolus"),
        ("Lateral pain", "Subfibular impingement in advanced valgus — “lateral overload”"),
        ("Fatigue", "Aching arch after standing; loss of endurance"),
        ("Deformity", "Patient notices “fallen arch,” shoes wear medially / unevenly"),
        ("Function", "Difficulty on uneven ground, stairs, single-leg support"),
        ("Late", "Diffuse arthritis pain; ankle symptoms if Stage IV"),
    ]
    for i, (t, b) in enumerate(items):
        col, row = i % 3, i // 3
        x = Inches(0.4) + Inches(col * 4.25)
        y = Inches(1.7) + Inches(row * 2.4)
        round_rect(s, x, y, Inches(4.05), Inches(2.15), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.25), y + Inches(0.3), Inches(3.55), Inches(0.4),
                    t, size=18, bold=True, color=TEAL, font="Georgia")
        add_textbox(s, x + Inches(0.25), y + Inches(0.9), Inches(3.55), Inches(1.0),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s27_history(prs, n):
    s = blank_slide(prs)
    section_header(s, "History Taking Checklist", "Don’t miss the systemic and progressive clues")
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
    add_paras(s, Inches(0.7), Inches(1.7), Inches(11.5), Inches(5), qs, size=17, bullet=True, spacing=8)
    footer(s, n)
    return s


def s28_exam(prs, n):
    s = blank_slide(prs)
    section_header(s, "Physical Examination Framework", "Look • Feel • Move • Special tests • Neurovascular")
    steps = [
        ("1. Inspection", ["Standing alignment", "Arch height", "Heel valgus", "Too-many-toes", "Shoe wear"]),
        ("2. Palpation", ["PTT course", "Navicular", "Sinus tarsi", "Lateral gutter", "Achilles"]),
        ("3. Motion", ["Ankle DF/PF", "ST inversion/eversion", "TN mobility", "Equinus test", "Rigidity check"]),
        ("4. Function", ["Single heel rise", "Double heel rise", "Tip-toe arch", "Gait", "Strength grading"]),
    ]
    for i, (t, lines) in enumerate(steps):
        card(s, Inches(0.35) + Inches(i * 3.25), Inches(1.65), Inches(3.1), Inches(4.8), t, lines, TEAL)
    footer(s, n)
    return s


def s29_special_tests(prs, n):
    s = blank_slide(prs)
    section_header(s, "High-Yield Special Tests", "Too-many-toes & single-limb heel rise")
    add_image(s, ASSETS / "exam_signs.png", Inches(0.5), Inches(1.4), width=Inches(12.3))
    footer(s, n)
    return s


def s30_more_exam(prs, n):
    s = blank_slide(prs)
    section_header(s, "More Exam Pearls", "Details that change staging")
    pearls = [
        ("Silfverskiöld test", "Differentiates gastroc vs Achilles equinus — guides recession vs TAL"),
        ("Passive correction", "Can you correct heel to neutral? If no → Stage III territory"),
        ("First ray mobility", "Residual forefoot varus after heel correction → Cotton candidate"),
        ("Peroneal spasm", "Rigid painful adolescent foot → coalition until proven otherwise"),
        ("Neurovascular", "Protect against Charcot / neuropathy missed diagnoses"),
        ("Opposite foot", "Compare; bilateral disease common but asymmetric staging possible"),
    ]
    for i, (t, b) in enumerate(pearls):
        y = Inches(1.55) + Inches(i * 0.82)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), WHITE, LIGHT)
        add_textbox(s, Inches(0.75), y + Inches(0.18), Inches(3.2), Inches(0.4),
                    t, size=14, bold=True, color=CORAL if i % 2 else TEAL)
        add_textbox(s, Inches(4.1), y + Inches(0.18), Inches(8.4), Inches(0.4),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s31_imaging_xr(prs, n):
    s = blank_slide(prs)
    section_header(s, "Imaging — Weightbearing Radiographs", "Standing films are mandatory")
    views = [
        ("AP foot", "TN coverage, metatarsus adductus/abductus, midfoot arthritis"),
        ("Lateral foot", "Meary’s angle, calcaneal pitch, cyma line, TN sag"),
        ("Hindfoot alignment", "Saltzman view — quantifies heel valgus"),
        ("Ankle series", "Mortise tilt for Stage IV; arthritis"),
        ("Oblique foot", "Calcaneonavicular coalition (anteater nose)"),
        ("Compare prior", "Progression over time guides urgency"),
    ]
    for i, (t, b) in enumerate(views):
        col, row = i % 3, i // 3
        x = Inches(0.4) + Inches(col * 4.25)
        y = Inches(1.7) + Inches(row * 2.4)
        round_rect(s, x, y, Inches(4.05), Inches(2.15), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.25), y + Inches(0.35), Inches(3.55), Inches(0.4),
                    t, size=17, bold=True, color=TEAL, font="Georgia")
        add_textbox(s, x + Inches(0.25), y + Inches(0.95), Inches(3.55), Inches(0.9),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s32_measurements(prs, n):
    s = blank_slide(prs)
    section_header(s, "Key Radiographic Measurements", "Numbers you quote on rounds")
    add_image(s, ASSETS / "xray_measurements.png", Inches(0.55), Inches(1.4), width=Inches(12.2))
    footer(s, n)
    return s


def s33_advanced_imaging(prs, n):
    s = blank_slide(prs)
    section_header(s, "MRI · CT · Ultrasound", "When plain films aren’t enough")
    cols = [
        ("MRI", ["PTT signal / tears", "Spring ligament", "Synovitis", "Early arthritis", "Deltoid integrity"]),
        ("CT", ["Tarsal coalition", "Joint degeneration map", "Osteotomy planning", "Subtalar morphology", "Fracture sequelae"]),
        ("Ultrasound", ["Dynamic PTT exam", "Tenosynovial fluid", "Cheap & accessible", "Operator dependent", "Limited bony detail"]),
    ]
    for i, (t, lines) in enumerate(cols):
        card(s, Inches(0.45) + Inches(i * 4.2), Inches(1.7), Inches(4.0), Inches(4.7), t, lines,
             TEAL if i == 0 else (CORAL if i == 1 else TEAL_DARK))
    footer(s, n)
    return s


def s34_ddx(prs, n):
    s = blank_slide(prs)
    section_header(s, "Differential Diagnosis", "Not every flat painful foot is PTTD")
    ddx = [
        ("Tarsal coalition", "Rigid adolescent; peroneal spasm; CT"),
        ("Accessory navicular", "Medial prominence pain; may coexist"),
        ("Spring ligament tear", "Acute medial trauma; MRI"),
        ("Charcot arthropathy", "Neuropathy; fragmentation; warmth"),
        ("RA / inflammatory", "Polyarticular; tenosynovitis multi-tendon"),
        ("Lisfranc injury", "Trauma; midfoot instability / arthritis"),
        ("Peroneal pathology", "Lateral symptoms; opposite tendon"),
        ("Subtalar arthritis", "Post-calcaneus fracture history"),
        ("Congenital vertical talus", "Rigid rocker-bottom from infancy"),
        ("Neuromuscular", "CP, myelomeningocele — tone & strength clues"),
    ]
    for i, (t, b) in enumerate(ddx):
        col, row = i % 2, i // 2
        x = Inches(0.4) + Inches(col * 6.45)
        y = Inches(1.55) + Inches(row * 0.95)
        round_rect(s, x, y, Inches(6.25), Inches(0.85), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.2), y + Inches(0.22), Inches(2.4), Inches(0.45),
                    t, size=13, bold=True, color=TEAL)
        add_textbox(s, x + Inches(2.7), y + Inches(0.22), Inches(3.3), Inches(0.45),
                    b, size=13, color=SLATE)
    footer(s, n)
    return s


def s35_nonop(prs, n):
    s = blank_slide(prs)
    section_header(s, "Conservative Treatment Overview", "First-line for Stage I–II and poor surgical hosts")
    add_image(s, ASSETS / "treatment_pyramid.png", Inches(0.55), Inches(1.4), width=Inches(12.2))
    footer(s, n)
    return s


def s36_orthotics(prs, n):
    s = blank_slide(prs)
    section_header(s, "Orthotics & Footwear", "Brace the deformity you can’t yet reconstruct")
    items = [
        ("Medial longitudinal posting", "Supports MLA; reduces PTT strain in mild disease"),
        ("UCBL orthosis", "Deeper heel cup; controls flexible heel valgus"),
        ("AFO / Arizona brace", "Advanced flexible deformity; immobilizes ST motion"),
        ("Footwear", "Stable heel counter, wide toe box, skip flexible flip-flops"),
        ("Goal", "Pain relief + slow progression — not “cure” of anatomy"),
        ("Compliance", "Counsel that braces work only when worn"),
    ]
    for i, (t, b) in enumerate(items):
        y = Inches(1.55) + Inches(i * 0.82)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), WHITE, LIGHT)
        add_textbox(s, Inches(0.75), y + Inches(0.18), Inches(4.0), Inches(0.4),
                    t, size=14, bold=True, color=TEAL)
        add_textbox(s, Inches(5.0), y + Inches(0.18), Inches(7.5), Inches(0.4),
                    b, size=14, color=SLATE)
    footer(s, n)
    return s


def s37_pt(prs, n):
    s = blank_slide(prs)
    section_header(s, "Physical Therapy Strategies", "Strengthen what remains — stretch what pulls into valgus")
    cards_data = [
        ("Strengthen", ["PTT / invertors", "Intrinsics", "Hip abductors", "Core for gait"]),
        ("Stretch / Mobilize", ["Gastroc–soleus", "Silfverskiöld-guided", "Avoid aggressive ST mobilization if painful arthritis"]),
        ("Functional", ["Proprioception", "Gait retraining", "Activity pacing", "Weight management referral"]),
        ("Limits", ["Won’t reverse Stage III", "Adjunct to orthosis", "Prehab before surgery", "Post-op protocols matter"]),
    ]
    for i, (t, lines) in enumerate(cards_data):
        card(s, Inches(0.35) + Inches(i * 3.25), Inches(1.7), Inches(3.1), Inches(4.6), t, lines, TEAL)
    footer(s, n)
    return s


def s38_sx_indications(prs, n):
    s = blank_slide(prs)
    section_header(s, "Surgical Indications", "When to put the scalpel in the plan")
    add_paras(s, Inches(0.7), Inches(1.7), Inches(11.5), Inches(5), [
        "Persistent pain & dysfunction after adequate nonoperative care (typically ≥3–6 months)",
        "Progressive deformity despite bracing",
        "Stage II flexible deformity suitable for joint-preserving reconstruction",
        "Stage III rigid symptomatic deformity → realignment arthrodesis",
        "Stage IV with ankle involvement — reconstruct or fuse based on joint status",
        "Rigid pediatric coalition failing conservative care",
        "Patient optimized: nicotine cessation, BMI discussion, glycemic control",
        "Clear goals: pain relief, plantigrade stable foot, brace-free ambulation when possible",
    ], size=17, bullet=True, spacing=8)
    footer(s, n)
    return s


def s39_soft_tissue(prs, n):
    s = blank_slide(prs)
    section_header(s, "Soft-Tissue Procedures", "Tendon & ligament work in Stage I–II")
    procs = [
        ("PTT débridement / tenosynovectomy", "Stage I refractory synovitis"),
        ("FDL transfer to navicular", "Workhorse dynamic replacement for PTT"),
        ("FHL transfer", "Alternative when FDL insufficient / prior harvest"),
        ("Spring ligament repair", "Address static TN support"),
        ("Deltoid reconstruction", "Stage IV flexible ankle valgus"),
        ("Gastroc recession / TAL", "Correct equinus driving valgus"),
    ]
    for i, (t, b) in enumerate(procs):
        y = Inches(1.55) + Inches(i * 0.82)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), WHITE, LIGHT)
        add_textbox(s, Inches(0.75), y + Inches(0.18), Inches(6.5), Inches(0.4),
                    t, size=15, bold=True, color=NAVY)
        add_textbox(s, Inches(7.4), y + Inches(0.18), Inches(5.1), Inches(0.4),
                    b, size=14, color=TEAL)
    footer(s, n)
    return s


def s40_osteotomies(prs, n):
    s = blank_slide(prs)
    section_header(s, "Osteotomies for Flexible Flatfoot", "Bony realignment that protects soft-tissue repairs")
    osteos = [
        ("Medializing calcaneal osteotomy (MCO)", "Shifts Achilles medially; reduces valgus moment — Stage II staple"),
        ("Evans lateral column lengthening", "Corrects severe forefoot abduction / TN uncoverage (IIB)"),
        ("Cotton (medial cuneiform opening wedge)", "Plantarflexes first ray; treats residual forefoot varus"),
        ("First TMT fusion (modified Lapidus)", "When first ray unstable / arthritic"),
        ("Extra-articular osteotomies preferred", "Preserve motion in flexible disease — Kaplan principle"),
    ]
    for i, (t, b) in enumerate(osteos):
        y = Inches(1.6) + Inches(i * 0.95)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.85), WHITE, LIGHT)
        rect(s, Inches(0.5), y, Inches(0.12), Inches(0.85), TEAL if i % 2 == 0 else CORAL)
        add_textbox(s, Inches(0.9), y + Inches(0.12), Inches(11.6), Inches(0.3),
                    t, size=15, bold=True, color=TEAL)
        add_textbox(s, Inches(0.9), y + Inches(0.45), Inches(11.6), Inches(0.3),
                    b, size=13, color=SLATE)
    footer(s, n)
    return s


def s41_arthrodesis(prs, n):
    s = blank_slide(prs)
    section_header(s, "Arthrodesis Options", "For rigidity, arthritis, and salvage")
    cards_data = [
        ("Subtalar fusion", ["Isolated ST arthritis", "Selected coalitions", "Less motion loss than triple"]),
        ("Double / Triple", ["TN + ST ± CC", "Stage III classic", "Correct rotation in all planes", "Nonunion risk if smoking"]),
        ("Midfoot fusions", ["TN / NC / TMT arthritis", "Often adjunctive", "Restore medial column"]),
        ("TTC / Pantalar", ["Stage IV salvage", "Neuropathic caution", "Highest morbidity"]),
    ]
    for i, (t, lines) in enumerate(cards_data):
        card(s, Inches(0.35) + Inches(i * 3.25), Inches(1.7), Inches(3.1), Inches(4.6), t, lines,
             CORAL if i == 3 else TEAL)
    footer(s, n)
    return s


def s42_algorithm(prs, n):
    s = blank_slide(prs)
    section_header(s, "Stage-Based Surgical Algorithm", "A Kaplan-oriented decision map")
    add_image(s, ASSETS / "surgical_algorithm.png", Inches(0.3), Inches(1.35), width=Inches(12.7))
    footer(s, n)
    return s


def s43_coalition_sx(prs, n):
    s = blank_slide(prs)
    section_header(s, "Pediatric Rigid Flatfoot — Coalitions", "Resect early coalitions; fuse late arthritis")
    add_image(s, ASSETS / "coalition.png", Inches(1.0), Inches(1.4), width=Inches(11.2))
    footer(s, n)
    return s


def s44_complications(prs, n):
    s = blank_slide(prs)
    section_header(s, "Complications & Pitfalls", "Know them before you consent")
    left = ["Under-/overcorrection of heel", "Nonunion of osteotomy or fusion",
            "Survivorship failure of FDL transfer if bony deformity ignored",
            "Lateral column pain after Evans", "Wound issues medially", "CRPS / nerve irritation"]
    right = ["Progressive arthritis after reconstruction", "Malunion → residual valgus",
             "Calcaneal osteotomy: sural / medial neurovascular risk",
             "TAL overlengthening → push-off weakness",
             "Smoking → nonunion (counsel hard)", "Missing Stage IV ankle disease"]
    card(s, Inches(0.45), Inches(1.7), Inches(6.1), Inches(4.7), "Intraop / Early", left, CORAL)
    card(s, Inches(6.8), Inches(1.7), Inches(6.1), Inches(4.7), "Late / Planning", right, TEAL)
    footer(s, n)
    return s


def s45_outcomes(prs, n):
    s = blank_slide(prs)
    section_header(s, "Outcomes & Prognosis", "What to tell patients honestly")
    add_paras(s, Inches(0.7), Inches(1.7), Inches(11.5), Inches(5), [
        "Stage I–II nonop: many improve functionally with bracing + PT",
        "Stage II reconstruction: high satisfaction when alignment restored & soft tissues balanced",
        "FDL + MCO is the most validated combination for flexible AAFD",
        "Triple arthrodesis relieves pain but sacrifices hindfoot motion — counsel gait changes",
        "Obesity & smoking worsen every outcome metric",
        "Residual mild stiffness is common; plantigrade painless foot is the goal",
        "Kaplan teaching: correct all planes — residual forefoot varus sinks reconstructions",
    ], size=17, bullet=True, spacing=8)
    footer(s, n)
    return s


def s46_pearls(prs, n):
    s = blank_slide(prs)
    section_header(s, "Clinical Pearls — Kaplan High Yield", "Memorize these before boards & clinic")
    pearls = [
        "Flexible vs rigid is the first branch point — tip-toe & passive correction",
        "Too-many-toes = forefoot abduction; single heel rise tests PTT function",
        "Weightbearing XR only — Meary, calc pitch, TN uncoverage, ankle mortise",
        "Stage II → joint preserving (FDL + MCO ± LCL ± Cotton ± spring lig)",
        "Stage III → triple; Stage IV → address ankle / deltoid",
        "Always treat equinus or valgus recurs",
        "Rigid adolescent flatfoot = coalition until proven otherwise",
        "Don’t fuse a flexible foot; don’t reconstruct a rigid arthritic one",
    ]
    for i, p in enumerate(pearls):
        y = Inches(1.5) + Inches(i * 0.62)
        round_rect(s, Inches(0.5), y, Inches(0.55), Inches(0.48), TEAL if i < 4 else CORAL)
        add_textbox(s, Inches(0.55), y + Inches(0.08), Inches(0.45), Inches(0.35),
                    str(i + 1), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.25), y + Inches(0.08), Inches(11.5), Inches(0.4),
                    p, size=14, color=NAVY)
    footer(s, n)
    return s


def s47_case1(prs, n):
    s = blank_slide(prs)
    section_header(s, "Case Vignette 1 — Stage II AAFD", "Apply the algorithm")
    round_rect(s, Inches(0.5), Inches(1.6), Inches(12.3), Inches(2.3), LIGHT)
    add_textbox(s, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.9),
                "52-year-old woman with 18 months of medial ankle pain. BMI 34. Too-many-toes positive. "
                "Unable to perform single heel rise. Heel valgus corrects passively. "
                "WB XR: Meary −15°, TN uncoverage 45%, ankle mortise normal. Failed 5 months of AFO + PT.",
                size=16, color=NAVY, font="Georgia")
    add_paras(s, Inches(0.7), Inches(4.2), Inches(11.5), Inches(2.5), [
        "Diagnosis: Stage IIB adult acquired flatfoot (PTTD)",
        "Plan: FDL transfer + medializing calcaneal osteotomy + spring ligament repair",
        "Consider Evans LCL given high TN uncoverage; Cotton if residual FF varus intra-op",
        "Address equinus; counsel weight loss & nicotine avoidance",
    ], size=15, bullet=True, spacing=6)
    footer(s, n)
    return s


def s48_case2(prs, n):
    s = blank_slide(prs)
    section_header(s, "Case Vignette 2 — Rigid Adolescent Flatfoot", "Don’t miss the coalition")
    round_rect(s, Inches(0.5), Inches(1.6), Inches(12.3), Inches(2.3), LIGHT)
    add_textbox(s, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.9),
                "13-year-old boy with activity-related foot pain for 1 year. Rigid planovalgus, "
                "peroneal spasm, no arch on tip-toe. Lateral XR suggests C-sign; Harris view abnormal. "
                "CT confirms middle-facet talocalcaneal coalition involving ~30% of the facet. No arthritis.",
                size=16, color=NAVY, font="Georgia")
    add_paras(s, Inches(0.7), Inches(4.2), Inches(11.5), Inches(2.5), [
        "Diagnosis: Symptomatic talocalcaneal coalition causing rigid flatfoot",
        "Nonop trial: activity mod, NSAID, casting/boot — if fails → surgery",
        "Surgery: coalition resection + fat interposition (± calcaneal osteotomy if valgus remains)",
        "Kaplan: <50% involvement without arthritis favors resection over primary fusion",
    ], size=15, bullet=True, spacing=6)
    footer(s, n)
    return s


def s49_boards(prs, n):
    s = blank_slide(prs)
    section_header(s, "Board-Style Rapid Fire", "Say the answer out loud")
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
        y = Inches(1.5) + Inches(row * 1.25)
        round_rect(s, x, y, Inches(6.25), Inches(1.1), WHITE, LIGHT)
        add_textbox(s, x + Inches(0.2), y + Inches(0.15), Inches(5.85), Inches(0.35),
                    q, size=13, bold=True, color=TEAL)
        add_textbox(s, x + Inches(0.2), y + Inches(0.55), Inches(5.85), Inches(0.4),
                    a, size=14, color=NAVY)
    footer(s, n)
    return s


def s50_summary(prs, n):
    s = blank_slide(prs)
    section_header(s, "Summary & Take-Home", "One page to remember tomorrow")
    points = [
        "Define the deformity: flexible vs rigid; pediatric vs adult; stage I–IV",
        "PTT + spring ligament failure drive the AAFD cascade",
        "Exam: too-many-toes, single heel rise, Silfverskiöld, passive correctability",
        "Image standing: Meary, calc pitch, TN coverage, ankle mortise ± CT/MRI",
        "Nonop first for flexible disease; reconstruct Stage II; fuse Stage III; add ankle for IV",
        "Coalitions: resect early, fuse late — never ignore rigid kids",
    ]
    for i, p in enumerate(points):
        y = Inches(1.55) + Inches(i * 0.75)
        round_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.65), WHITE, LIGHT)
        round_rect(s, Inches(0.65), y + Inches(0.12), Inches(0.42), Inches(0.42), TEAL)
        add_textbox(s, Inches(0.7), y + Inches(0.18), Inches(0.35), Inches(0.3),
                    str(i + 1), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.3), y + Inches(0.15), Inches(11.2), Inches(0.4),
                    p, size=15, color=NAVY)
    footer(s, n)
    return s


def s51_refs_qa(prs, n):
    """Closing slide — take-home + references + Q&A."""
    s = blank_slide(prs)
    rect(s, 0, 0, W, H, NAVY)
    accent_bar(s, width=Inches(0.25))
    add_textbox(s, Inches(0.7), Inches(0.4), Inches(12), Inches(0.55),
                "Take-Home · References · Questions", size=28, bold=True, color=WHITE, font="Georgia")
    takeaways = [
        "Flexible vs rigid first — then stage I–IV",
        "PTT + spring ligament drive AAFD",
        "Stage II = reconstruct; III = fuse; IV = add ankle",
        "Rigid adolescent foot → coalition until proven otherwise",
    ]
    for i, t in enumerate(takeaways):
        x = Inches(0.7) + Inches((i % 2) * 6.2)
        y = Inches(1.15) + Inches((i // 2) * 0.7)
        round_rect(s, x, y, Inches(5.9), Inches(0.55), TEAL)
        add_textbox(s, x + Inches(0.2), y + Inches(0.1), Inches(5.5), Inches(0.4),
                    t, size=13, bold=True, color=WHITE)
    add_textbox(s, Inches(0.7), Inches(2.75), Inches(12), Inches(0.35),
                "Core references", size=16, bold=True, color=RGBColor(0x7F, 0xDB, 0xDA), font="Georgia")
    refs = [
        "Kaplan Medical — Orthopedic / USMLE flatfoot & PTTD review sections",
        "Johnson KA, Strom DE. Tibialis posterior tendon dysfunction.",
        "Myerson MS. Adult acquired flatfoot deformity (Stage IV addition).",
        "AAOS / AOFAS ICL chapters on AAFD; Coughlin–Mann Foot & Ankle",
    ]
    for i, r in enumerate(refs):
        add_textbox(s, Inches(0.9), Inches(3.2) + Inches(i * 0.42), Inches(11.5), Inches(0.4),
                    f"•  {r}", size=13, color=LIGHT)
    add_textbox(s, Inches(0.7), Inches(5.2), Inches(12), Inches(0.5),
                "Questions?  Let’s discuss cases.", size=24, bold=True, color=RGBColor(0x7F, 0xDB, 0xDA), font="Georgia")
    add_textbox(s, Inches(0.7), Inches(5.9), Inches(12), Inches(0.4),
                "Educational lecture for orthopedic surgery trainees  ·  Flat Foot (Pes Planus)", size=13, color=MUTED)
    add_textbox(s, Inches(11.2), Inches(7.05), Inches(1.8), Inches(0.35),
                f"{n}  /  50", size=11, color=MUTED, align=PP_ALIGN.RIGHT)
    return s


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    builders = [
        s01_title,
        s02_objectives,
        s03_roadmap,
        s04_why,
        s05_anatomy_overview,
        s06_mla,
        s07_keystone,
        s08_supports,
        s09_ptt_anatomy,
        s10_spring,
        s11_definition,
        s12_epi,
        s13_class_overview,
        s14_flex_vs_rigid,
        s15_congenital_vs_acq,
        s16_pediatric,
        s17_physio_kids,
        s18_aafd_intro,
        s19_etiology,
        s20_stages_overview,
        s21_stage1,
        s22_stage2,
        s23_stage3,
        s24_stage4,
        s25_pathomech,
        s26_clinical,
        s27_history,
        s28_exam,
        s29_special_tests,
        s30_more_exam,
        s31_imaging_xr,
        s32_measurements,
        s33_advanced_imaging,
        s34_ddx,
        s35_nonop,
        s36_orthotics,
        s37_pt,
        s38_sx_indications,
        s39_soft_tissue,
        s40_osteotomies,
        s41_arthrodesis,
        s42_algorithm,
        s43_coalition_sx,
        s44_complications,
        s45_outcomes,
        s46_pearls,
        s47_case1,
        s48_case2,
        s49_boards,
        s50_summary,
    ]

    assert len(builders) == 50, f"Expected 50 builders, got {len(builders)}"

    for i, fn in enumerate(builders, start=1):
        if fn is s01_title:
            fn(prs)
        else:
            fn(prs, i)

    prs.save(OUT)
    print(f"Saved {OUT} with {len(prs.slides)} slides")
    return OUT


if __name__ == "__main__":
    build()
