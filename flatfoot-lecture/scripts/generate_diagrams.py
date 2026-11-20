#!/usr/bin/env python3
"""Generate medical illustration assets for Flat Foot lecture."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parents[1] / "assets"
OUT.mkdir(parents=True, exist_ok=True)

TEAL = (13, 115, 119)
NAVY = (27, 40, 56)
CORAL = (224, 122, 95)
SOFT = (247, 249, 252)
LIGHT_TEAL = (180, 220, 218)
BONE = (236, 230, 220)
LINE = (55, 70, 90)
WHITE = (255, 255, 255)
MUTED = (120, 135, 150)


def font(size, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def rounded_bg(size, color=SOFT):
    img = Image.new("RGBA", size, color + (255,))
    return img


def label(draw, xy, text, size=18, fill=NAVY, bold=True):
    draw.text(xy, text, font=font(size, bold=bold), fill=fill)


def draw_foot_side(draw, origin, scale=1.0, collapsed=False, color=NAVY):
    ox, oy = origin
    # Heel
    heel = [
        (ox + 20 * scale, oy + 80 * scale),
        (ox + 70 * scale, oy + 95 * scale),
        (ox + 90 * scale, oy + 70 * scale),
        (ox + 55 * scale, oy + 40 * scale),
        (ox + 25 * scale, oy + 50 * scale),
    ]
    # Midfoot arch
    if collapsed:
        mid = [
            (ox + 90 * scale, oy + 70 * scale),
            (ox + 160 * scale, oy + 78 * scale),
            (ox + 230 * scale, oy + 72 * scale),
        ]
        arch_y = oy + 78 * scale
    else:
        mid = [
            (ox + 90 * scale, oy + 70 * scale),
            (ox + 160 * scale, oy + 35 * scale),
            (ox + 230 * scale, oy + 55 * scale),
        ]
        arch_y = oy + 35 * scale
    # Forefoot
    toes = [
        (ox + 230 * scale, oy + (72 if collapsed else 55) * scale),
        (ox + 300 * scale, oy + 65 * scale),
        (ox + 320 * scale, oy + 85 * scale),
        (ox + 280 * scale, oy + 95 * scale),
        (ox + 160 * scale, oy + 95 * scale),
        (ox + 70 * scale, oy + 95 * scale),
    ]
    poly = heel + mid[1:] + toes
    draw.polygon(poly, fill=BONE, outline=color)
    # Arch marker
    draw.ellipse(
        [ox + 145 * scale, arch_y - 8 * scale, ox + 175 * scale, arch_y + 8 * scale],
        fill=CORAL if collapsed else TEAL,
    )
    # Ground line
    gy = oy + 100 * scale
    draw.line([(ox, gy), (ox + 340 * scale, gy)], fill=MUTED, width=2)
    return arch_y


def make_arch_comparison():
    img = rounded_bg((1200, 700))
    d = ImageDraw.Draw(img)
    label(d, (40, 30), "Medial Longitudinal Arch — Normal vs Flatfoot", 28)
    label(d, (80, 120), "NORMAL", 22, TEAL)
    draw_foot_side(d, (80, 180), 1.4, collapsed=False)
    d.line([(280, 280), (280, 360)], fill=TEAL, width=3)
    label(d, (210, 370), "Preserved MLA", 18, TEAL)

    label(d, (680, 120), "PES PLANUS", 22, CORAL)
    draw_foot_side(d, (650, 180), 1.4, collapsed=True)
    d.line([(850, 330), (850, 360)], fill=CORAL, width=3)
    label(d, (760, 370), "Collapsed MLA", 18, CORAL)

    # Key points box
    d.rounded_rectangle([60, 480, 1140, 650], radius=16, fill=WHITE, outline=LIGHT_TEAL, width=3)
    label(d, (90, 500), "Kaplan Pearl", 20, TEAL)
    label(d, (90, 545), "• Flexible flatfoot: arch reconstitutes on toe-standing / non-weightbearing", 18, NAVY, False)
    label(d, (90, 585), "• Rigid flatfoot: arch remains collapsed — think tarsal coalition / arthritis", 18, NAVY, False)
    img.save(OUT / "arch_comparison.png")


def make_anatomy_supports():
    img = rounded_bg((1200, 800))
    d = ImageDraw.Draw(img)
    label(d, (40, 30), "Dynamic & Static Supports of the Medial Arch", 28)

    boxes = [
        (60, 120, "STATIC", TEAL, [
            "Spring ligament (calcaneonavicular)",
            "Plantar fascia",
            "Long & short plantar ligaments",
            "Bony architecture (talus–navicular)",
        ]),
        (620, 120, "DYNAMIC", CORAL, [
            "Posterior tibial tendon (key)",
            "Flexor digitorum longus",
            "Flexor hallucis longus",
            "Intrinsic foot muscles",
        ]),
    ]
    for x, y, title, color, items in boxes:
        d.rounded_rectangle([x, y, x + 500, y + 380], radius=20, fill=WHITE, outline=color, width=4)
        d.rounded_rectangle([x, y, x + 500, y + 70], radius=20, fill=color)
        # Fix bottom corners of header
        d.rectangle([x, y + 40, x + 500, y + 70], fill=color)
        label(d, (x + 30, y + 20), title, 26, WHITE)
        for i, item in enumerate(items):
            label(d, (x + 40, y + 110 + i * 55), f"▸  {item}", 20, NAVY, False)

    d.rounded_rectangle([60, 540, 1140, 740], radius=16, fill=(232, 245, 244), outline=TEAL, width=2)
    label(d, (90, 570), "Clinical Correlation (Kaplan)", 22, TEAL)
    label(d, (90, 620), "PTT dysfunction → unopposed eversion → midfoot abduction + heel valgus → progressive collapse.", 18, NAVY, False)
    label(d, (90, 670), "Spring ligament failure often accompanies advanced AAFD — check for talonavicular uncoverage.", 18, NAVY, False)
    img.save(OUT / "arch_supports.png")


def make_pttd_stages():
    img = rounded_bg((1400, 900))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Adult Acquired Flatfoot — Johnson & Strom (+ Myerson IV)", 26)

    stages = [
        ("I", "Tenosynovitis", "Normal alignment\nPain / swelling\nmedial ankle", TEAL, "Conservative\n± debridement"),
        ("II", "Flexible\ndeformity", "Pes planovalgus\nCorrectable\n+Too-many-toes", (46, 139, 140), "FDL transfer\n+calc osteotomy"),
        ("III", "Rigid\ndeformity", "Fixed valgus\nMidfoot arthritis\nNo correction", CORAL, "Triple\narthrodesis"),
        ("IV", "Ankle\ninvolvement", "Valgus talar\ntilt in ankle\nDeltoid failure", (180, 70, 70), "Triple + ankle\nreconstruction"),
    ]
    w = 300
    for i, (num, title, body, color, tx) in enumerate(stages):
        x = 50 + i * 335
        y = 100
        d.rounded_rectangle([x, y, x + w, y + 720], radius=18, fill=WHITE, outline=color, width=4)
        d.ellipse([x + 110, y + 25, x + 190, y + 105], fill=color)
        label(d, (x + 135, y + 42), num, 36, WHITE)
        label(d, (x + 30, y + 130), title, 22, color)
        for j, line in enumerate(body.split("\n")):
            label(d, (x + 30, y + 220 + j * 36), line, 18, NAVY, False)
        d.rounded_rectangle([x + 20, y + 420, x + w - 20, y + 560], radius=12, fill=(color[0], color[1], color[2], 40) if False else (240, 248, 247), outline=color, width=2)
        label(d, (x + 35, y + 440), "Treatment", 16, color)
        for j, line in enumerate(tx.split("\n")):
            label(d, (x + 35, y + 475 + j * 32), line, 17, NAVY, False)
        if i < 3:
            d.polygon([(x + w + 8, 420), (x + w + 28, 440), (x + w + 8, 460)], fill=MUTED)
    img.save(OUT / "pttd_stages.png")


def make_exam_signs():
    img = rounded_bg((1200, 750))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "High-Yield Clinical Signs", 28)

    # Too many toes
    d.rounded_rectangle([50, 100, 570, 680], radius=18, fill=WHITE, outline=TEAL, width=3)
    label(d, (80, 130), "“Too-Many-Toes” Sign", 24, TEAL)
    # Foot from behind schematic
    cx, cy = 310, 380
    d.ellipse([cx - 40, cy - 100, cx + 40, cy + 100], fill=BONE, outline=NAVY, width=2)  # heel
    d.ellipse([cx - 10, cy - 140, cx + 50, cy - 60], fill=LIGHT_TEAL, outline=TEAL, width=2)  # forefoot abducted
    for i, t in enumerate(["2", "3", "4", "5"]):
        d.ellipse([cx + 45, cy - 130 + i * 35, cx + 75, cy - 105 + i * 35], fill=CORAL)
        label(d, (cx + 85, cy - 128 + i * 35), f"Toe {t}", 14, CORAL, False)
    label(d, (80, 560), "Posterior view: lateral toes visible", 16, NAVY, False)
    label(d, (80, 600), "→ forefoot abduction / midfoot uncoverage", 16, NAVY, False)

    # Single heel rise
    d.rounded_rectangle([630, 100, 1150, 680], radius=18, fill=WHITE, outline=CORAL, width=3)
    label(d, (660, 130), "Single-Limb Heel Rise", 24, CORAL)
    # Standing heel
    d.polygon([(820, 480), (880, 300), (940, 480)], fill=BONE, outline=NAVY, width=2)
    d.ellipse([860, 270, 900, 310], fill=TEAL)  # inverted heel
    d.arc([780, 500, 980, 620], 200, 340, fill=MUTED, width=3)
    label(d, (860, 250), "Varus", 16, TEAL)
    label(d, (660, 540), "Normal: heel inverts via PTT", 16, NAVY, False)
    label(d, (660, 580), "PTTD: weak / absent inversion", 16, NAVY, False)
    label(d, (660, 620), "or inability to rise on toes", 16, NAVY, False)
    img.save(OUT / "exam_signs.png")


def make_radiographic_angles():
    img = rounded_bg((1200, 800))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Key Radiographic Measurements (Weightbearing)", 26)

    rows = [
        ("Meary’s angle", "AP / Lat talar–1st MT", "0° normal", ">4° convex plantar = planus"),
        ("Calcaneal pitch", "Lateral", "20–30°", "Decreased in flatfoot"),
        ("TN coverage / uncoverage", "AP", "<30% uncover", "Increased → abduction"),
        ("Hindfoot alignment", "Saltzman / axial", "0–5° valgus", "Increased heel valgus"),
        ("Talocalcaneal angle", "AP / Lat (Kite)", "15–35°", "Increased in planovalgus"),
    ]
    d.rounded_rectangle([40, 90, 1160, 160], radius=10, fill=TEAL)
    for i, h in enumerate(["Parameter", "View", "Normal", "Flatfoot Finding"]):
        label(d, (70 + i * 280, 110), h, 18, WHITE)
    for r, row in enumerate(rows):
        y = 175 + r * 100
        fill = WHITE if r % 2 == 0 else (236, 245, 244)
        d.rounded_rectangle([40, y, 1160, y + 90], radius=8, fill=fill, outline=LIGHT_TEAL)
        for i, cell in enumerate(row):
            label(d, (70 + i * 280, y + 30), cell, 17, NAVY, bold=(i == 0))
    img.save(OUT / "xray_measurements.png")


def make_treatment_pyramid():
    img = rounded_bg((1200, 800))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Treatment Pyramid — Flexible Adult Flatfoot", 26)
    layers = [
        (200, 120, 1000, 220, TEAL, "Activity modification • NSAIDs • Weight loss"),
        (260, 250, 940, 350, (46, 139, 140), "UCBL / AFO / medial posting orthosis"),
        (320, 380, 880, 480, (70, 150, 140), "Physical therapy — PTT eccentric strengthening"),
        (380, 510, 820, 610, CORAL, "Surgery if refractory symptomatic deformity"),
    ]
    for x1, y1, x2, y2, color, text in layers:
        d.polygon([(x1, y2), ((x1 + x2) // 2, y1), (x2, y2)], fill=color)
        # Use trapezoid instead
    # Better trapezoids
    traps = [
        [(150, 200), (1050, 200), (980, 290), (220, 290), TEAL, "1  NSAIDs • Rest • Weight optimization"],
        [(220, 320), (980, 320), (900, 410), (300, 410), (46, 139, 140), "2  Custom orthotics / brace (AFO, UCBL)"],
        [(300, 440), (900, 440), (820, 530), (380, 530), (70, 150, 140), "3  PT — strengthen PTT / Achilles stretch"],
        [(380, 560), (820, 560), (740, 680), (460, 680), CORAL, "4  Operative reconstruction"],
    ]
    img2 = rounded_bg((1200, 800))
    d = ImageDraw.Draw(img2)
    label(d, (40, 25), "Treatment Pyramid — Flexible Adult Flatfoot", 26)
    for x1, y1, x2, y2, x3, y3, x4, y4, color, text in [
        (150, 200, 1050, 200, 980, 290, 220, 290, TEAL, "1  NSAIDs • Rest • Weight optimization"),
        (220, 320, 980, 320, 900, 410, 300, 410, (46, 139, 140), "2  Custom orthotics / brace (AFO, UCBL)"),
        (300, 440, 900, 440, 820, 530, 380, 530, (70, 150, 140), "3  PT — strengthen PTT / Achilles stretch"),
        (380, 560, 820, 560, 740, 680, 460, 680, CORAL, "4  Operative reconstruction"),
    ]:
        d.polygon([(x1, y1), (x2, y2), (x3, y3), (x4, y4)], fill=color)
        label(d, (420 if "4" not in text else 470, (y1 + y3) // 2 - 12), text, 18, WHITE)
    label(d, (80, 720), "Kaplan: exhaust nonop care in flexible deformity before osteotomy + tendon transfer.", 16, NAVY, False)
    img2.save(OUT / "treatment_pyramid.png")


def make_surgical_algorithm():
    img = rounded_bg((1300, 850))
    d = ImageDraw.Draw(img)
    label(d, (40, 20), "Surgical Decision Algorithm (Kaplan-Oriented)", 26)

    def box(x, y, w, h, text, color=TEAL, fill=WHITE):
        d.rounded_rectangle([x, y, x + w, y + h], radius=12, fill=fill, outline=color, width=3)
        lines = text.split("\n")
        for i, line in enumerate(lines):
            label(d, (x + 16, y + 18 + i * 28), line, 16, NAVY if fill == WHITE else WHITE, bold=(i == 0))

    box(480, 70, 340, 70, "Symptomatic Flatfoot\nFailed nonoperative care", TEAL, TEAL)
    label(d, (620, 95), "", 1)  # keep layout
    # rewrite header text in white
    d.rounded_rectangle([480, 70, 820, 140], radius=12, fill=TEAL)
    label(d, (510, 85), "Symptomatic Flatfoot", 18, WHITE)
    label(d, (510, 112), "Failed nonoperative care", 15, WHITE, False)

    d.line([(650, 140), (650, 180)], fill=NAVY, width=3)
    box(480, 180, 340, 80, "Flexible?\n(correctable deformity)")

    d.line([(480, 220), (200, 220)], fill=NAVY, width=3)
    d.line([(820, 220), (1100, 220)], fill=NAVY, width=3)
    d.line([(200, 220), (200, 280)], fill=NAVY, width=3)
    d.line([(1100, 220), (1100, 280)], fill=NAVY, width=3)

    box(50, 280, 300, 100, "YES — Stage II\nFDL transfer +\nMedializing calc osteotomy\n± Cotton / lateral column")
    box(950, 280, 300, 100, "NO — Stage III\nTriple arthrodesis\n(ST + TN + CC)")

    d.line([(200, 380), (200, 440)], fill=NAVY, width=3)
    d.line([(1100, 380), (1100, 440)], fill=NAVY, width=3)
    box(50, 440, 300, 90, "Equinus?\n→ Gastroc recession\nor Achilles lengthening")
    box(950, 440, 300, 90, "Ankle valgus?\n→ Stage IV pathway\nDeltoid ± ankle fusion")

    d.rounded_rectangle([350, 580, 950, 780], radius=16, fill=(232, 245, 244), outline=TEAL, width=2)
    label(d, (380, 610), "Adjuncts & Pearls", 20, TEAL)
    label(d, (380, 660), "• Spring ligament repair if attenuated    • Lateral column lengthening if severe abduction", 16, NAVY, False)
    label(d, (380, 700), "• Cotton (medial cuneiform) opening wedge for residual forefoot varus", 16, NAVY, False)
    label(d, (380, 740), "• Never fuse a flexible foot when reconstruction can restore motion", 16, NAVY, False)
    img.save(OUT / "surgical_algorithm.png")


def make_pediatric_types():
    img = rounded_bg((1200, 750))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Pediatric Flatfoot Spectrum", 28)
    cards = [
        (50, "Physiologic\nFlexible", TEAL, "Most common\nAsymptomatic\nArch on tip-toe\nObserve / educate"),
        (340, "Flexible\nSymptomatic", (46, 139, 140), "Activity pain\nArch reconstitutes\nOrthotics / PT\nRarely surgery"),
        (630, "Rigid\n(Coalition)", CORAL, "Painful, stiff\nPeroneal spasm\nCalcaneonavicular\nTalocalcaneal"),
        (920, "Congenital\nVertical Talus", (180, 70, 70), "Rigid rocker-bottom\nIrreducible equinus\nNeeds early surgery\nOpposite of CVT ≠ FF"),
    ]
    for x, title, color, body in cards:
        d.rounded_rectangle([x, 100, x + 260, 680], radius=18, fill=WHITE, outline=color, width=4)
        d.rounded_rectangle([x, 100, x + 260, 220], radius=18, fill=color)
        d.rectangle([x, 180, x + 260, 220], fill=color)
        for i, line in enumerate(title.split("\n")):
            label(d, (x + 25, 125 + i * 32), line, 20, WHITE)
        for i, line in enumerate(body.split("\n")):
            label(d, (x + 25, 260 + i * 55), line, 17, NAVY, False)
    img.save(OUT / "pediatric_spectrum.png")


def make_title_hero():
    img = Image.new("RGB", (1600, 900), NAVY)
    d = ImageDraw.Draw(img)
    # Gradient-like bands
    for i in range(37):
        y0 = 600 + i * 8
        y1 = min(900, y0 + 8)
        c = int(13 + i * 1.5)
        d.rectangle([0, y0, 1600, y1], fill=(c, min(115 + i, 160), min(119 + i, 160)))
    # Abstract foot silhouette
    draw_foot_side(d, (950, 280), 1.6, collapsed=True, color=LIGHT_TEAL)
    draw_foot_side(d, (920, 200), 1.6, collapsed=False, color=TEAL)
    label(d, (80, 180), "FLAT FOOT", 72, WHITE)
    label(d, (80, 280), "Pes Planus & Adult Acquired Flatfoot", 32, LIGHT_TEAL, False)
    label(d, (80, 360), "A Kaplan-Based Lecture for Orthopedic Surgery Trainees", 22, (180, 200, 210), False)
    d.rounded_rectangle([80, 450, 420, 520], radius=8, fill=TEAL)
    label(d, (110, 470), "50 Educational Slides", 22, WHITE)
    img.save(OUT / "title_hero.png")


def make_pathomechanics():
    img = rounded_bg((1200, 700))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Pathomechanics Cascade", 28)
    steps = [
        ("1", "PTT failure", "Loss of dynamic\ninversion support"),
        ("2", "Spring lig.\nattenuation", "TN joint\nuncoverage"),
        ("3", "Hindfoot\nvalgus", "Calcaneus\neverts"),
        ("4", "Forefoot\nabduction", "Too-many-toes\nsign"),
        ("5", "Arch\ncollapse", "Meary angle\nincreases"),
        ("6", "Secondary\narthritis", "ST / TN / ankle\ninvolvement"),
    ]
    for i, (n, t, b) in enumerate(steps):
        x = 40 + i * 195
        d.ellipse([x + 55, 120, x + 115, 180], fill=TEAL if i < 4 else CORAL)
        label(d, (x + 75, 135), n, 28, WHITE)
        label(d, (x + 20, 210), t, 16, NAVY)
        for j, line in enumerate(b.split("\n")):
            label(d, (x + 20, 300 + j * 28), line, 14, MUTED, False)
        if i < 5:
            d.polygon([(x + 160, 140), (x + 185, 150), (x + 160, 160)], fill=MUTED)
    d.rounded_rectangle([40, 420, 1160, 640], radius=14, fill=WHITE, outline=LIGHT_TEAL, width=3)
    label(d, (70, 450), "Biomechanical Pearl", 20, TEAL)
    label(d, (70, 500), "Once the hindfoot is in valgus, the Achilles becomes an everting force and accelerates deformity.", 17, NAVY, False)
    label(d, (70, 545), "That is why gastrocnemius recession / TAL is commonly added to Stage II reconstructions.", 17, NAVY, False)
    label(d, (70, 590), "Lateral column shortening relative to medial column drives abduction — lengthen laterally if needed.", 17, NAVY, False)
    img.save(OUT / "pathomechanics.png")


def make_coalition_map():
    img = rounded_bg((1100, 700))
    d = ImageDraw.Draw(img)
    label(d, (40, 25), "Tarsal Coalition — Rigid Pediatric Flatfoot", 26)
    # Two big cards
    d.rounded_rectangle([50, 110, 520, 620], radius=16, fill=WHITE, outline=TEAL, width=3)
    label(d, (80, 140), "Calcaneonavicular", 24, TEAL)
    label(d, (80, 200), "• Most common overall", 18, NAVY, False)
    label(d, (80, 250), "• Ages ~8–12 years", 18, NAVY, False)
    label(d, (80, 300), "• Oblique XR: anteater nose", 18, NAVY, False)
    label(d, (80, 350), "• Rx: resection + fat interposition", 18, NAVY, False)
    label(d, (80, 400), "• Good prognosis if caught early", 18, NAVY, False)

    d.rounded_rectangle([580, 110, 1050, 620], radius=16, fill=WHITE, outline=CORAL, width=3)
    label(d, (610, 140), "Talocalcaneal", 24, CORAL)
    label(d, (610, 200), "• Middle facet most common", 18, NAVY, False)
    label(d, (610, 250), "• Ages ~12–16 years", 18, NAVY, False)
    label(d, (610, 300), "• Harris / CT: C-sign on lat XR", 18, NAVY, False)
    label(d, (610, 350), "• <50% facet: resection possible", 18, NAVY, False)
    label(d, (610, 400), "• Larger / arthritic → arthrodesis", 18, NAVY, False)
    img.save(OUT / "coalition.png")


if __name__ == "__main__":
    make_title_hero()
    make_arch_comparison()
    make_anatomy_supports()
    make_pttd_stages()
    make_exam_signs()
    make_radiographic_angles()
    make_treatment_pyramid()
    make_surgical_algorithm()
    make_pediatric_types()
    make_pathomechanics()
    make_coalition_map()
    print("Generated diagrams in", OUT)
    for p in sorted(OUT.glob("*.png")):
        print(" -", p.name, p.stat().st_size)
