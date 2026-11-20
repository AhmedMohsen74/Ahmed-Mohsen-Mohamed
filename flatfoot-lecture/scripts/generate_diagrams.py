#!/usr/bin/env python3
"""Generate compact, title-free diagram assets that fit under slide headers."""
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
MUTED = (120, 135, 150)
WHITE = (255, 255, 255)
SLATE = (44, 62, 80)


def font(size, bold=False):
    paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for p in paths:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def label(d, xy, text, size=16, fill=NAVY, bold=True):
    d.text(xy, text, font=font(size, bold=bold), fill=fill)


def bg(size, color=SOFT):
    return Image.new("RGBA", size, color + (255,))


def draw_foot_side(d, origin, scale=1.0, collapsed=False, color=NAVY):
    ox, oy = origin
    heel = [
        (ox + 20 * scale, oy + 80 * scale),
        (ox + 70 * scale, oy + 95 * scale),
        (ox + 90 * scale, oy + 70 * scale),
        (ox + 55 * scale, oy + 40 * scale),
        (ox + 25 * scale, oy + 50 * scale),
    ]
    if collapsed:
        mid = [(ox + 90 * scale, oy + 70 * scale), (ox + 160 * scale, oy + 78 * scale), (ox + 230 * scale, oy + 72 * scale)]
        arch_y = oy + 78 * scale
    else:
        mid = [(ox + 90 * scale, oy + 70 * scale), (ox + 160 * scale, oy + 35 * scale), (ox + 230 * scale, oy + 55 * scale)]
        arch_y = oy + 35 * scale
    toes = [
        (ox + 230 * scale, oy + (72 if collapsed else 55) * scale),
        (ox + 300 * scale, oy + 65 * scale),
        (ox + 320 * scale, oy + 85 * scale),
        (ox + 280 * scale, oy + 95 * scale),
        (ox + 160 * scale, oy + 95 * scale),
        (ox + 70 * scale, oy + 95 * scale),
    ]
    poly = heel + mid[1:] + toes
    d.polygon(poly, fill=BONE, outline=color, width=2)
    d.ellipse([ox + 145 * scale, arch_y - 8 * scale, ox + 175 * scale, arch_y + 8 * scale],
              fill=CORAL if collapsed else TEAL)
    gy = oy + 100 * scale
    d.line([(ox, gy), (ox + 340 * scale, gy)], fill=MUTED, width=2)
    return arch_y


def make_arch_comparison():
    img = bg((1100, 420))
    d = ImageDraw.Draw(img)
    label(d, (40, 20), "NORMAL", 20, TEAL)
    draw_foot_side(d, (40, 70), 1.15, False)
    label(d, (150, 360), "Preserved MLA", 16, TEAL)

    label(d, (580, 20), "PES PLANUS", 20, CORAL)
    draw_foot_side(d, (560, 70), 1.15, True)
    label(d, (680, 360), "Collapsed MLA", 16, CORAL)
    img.save(OUT / "arch_comparison.png")


def make_title_hero():
    img = Image.new("RGB", (1600, 900), NAVY)
    d = ImageDraw.Draw(img)
    for i in range(35):
        y0 = 620 + i * 8
        if y0 >= 900:
            break
        y1 = min(900, y0 + 8)
        if y1 <= y0:
            break
        c = int(13 + i * 1.5)
        d.rectangle([0, y0, 1600, y1], fill=(c, min(115 + i, 160), min(119 + i, 160)))
    draw_foot_side(d, (980, 220), 1.45, True, LIGHT_TEAL)
    draw_foot_side(d, (950, 140), 1.45, False, TEAL)
    label(d, (90, 200), "FLAT FOOT", 70, WHITE)
    label(d, (90, 300), "Pes Planus & Adult Acquired Flatfoot", 30, LIGHT_TEAL, False)
    label(d, (90, 370), "Kaplan-Based Lecture for Orthopedic Surgery Trainees", 20, (180, 200, 210), False)
    d.rounded_rectangle([90, 460, 430, 530], radius=10, fill=TEAL)
    label(d, (120, 480), "50 Educational Slides", 22, WHITE)
    img.save(OUT / "title_hero.png")


if __name__ == "__main__":
    make_title_hero()
    make_arch_comparison()
    print("Generated compact assets in", OUT)
