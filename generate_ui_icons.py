"""MailWeave — generate the sidebar/UI icon set.

Renders a coherent set of flat, line-style icons at 96px so they antialias
cleanly when scaled to the 22-26px the sidebar uses. Each icon is drawn in
the brand lime accent on a transparent background. Run from build.bat (or
manually) whenever the icon design changes.

The PyInstaller spec bundles every `icon_*.png` in the project root, so the
files just need to land there.
"""

from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw

HERE = Path(__file__).parent
ICON_SIZE = 96
LINE = 6
ACCENT = (132, 204, 22, 255)   # #84cc16 — lime accent
ACCENT_FILL = (132, 204, 22, 40)


def _new_canvas() -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new('RGBA', (ICON_SIZE, ICON_SIZE), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)


def _save(img: Image.Image, name: str) -> None:
    img.save(HERE / f'icon_{name}.png', format='PNG')


def _rounded_rect(draw, box, radius, outline=ACCENT, width=LINE, fill=None):
    draw.rounded_rectangle(box, radius=radius, outline=outline, width=width, fill=fill)


def files_icon():
    """Two stacked document sheets — for 'Open Files'."""
    img, d = _new_canvas()
    # Back doc — solid lime so it reads at 22px
    d.rounded_rectangle((14, 10, 66, 74), radius=8, fill=ACCENT)
    # Front doc — offset, also solid; the gap between the two reads as a stack
    d.rounded_rectangle((26, 24, 84, 88), radius=8, fill=ACCENT)
    # White index lines on the front doc
    for y in (46, 60, 74):
        d.line((36, y, 74, y), fill=(255, 255, 255, 255), width=LINE - 1)
    _save(img, 'files')


def folder_icon():
    """Folder with a tab — for 'Open Folder'."""
    img, d = _new_canvas()
    # Tab
    d.polygon([(14, 26), (40, 26), (48, 36), (14, 36)], fill=ACCENT)
    # Body
    _rounded_rect(d, (10, 30, 86, 80), radius=8, fill=ACCENT_FILL)
    # Subtle fold line
    d.line((10, 42, 86, 42), fill=ACCENT, width=LINE - 3)
    _save(img, 'folder')


def outlook_icon():
    """Envelope — for 'Import from Outlook'."""
    img, d = _new_canvas()
    _rounded_rect(d, (12, 24, 84, 76), radius=6, fill=ACCENT_FILL)
    # Flap (V) — two diagonals from top corners meeting in the middle
    d.line((12, 28, 48, 54), fill=ACCENT, width=LINE)
    d.line((84, 28, 48, 54), fill=ACCENT, width=LINE)
    _save(img, 'outlook')


def bundle_icon():
    """Stack of bound pages — for 'Indexed Bundle'."""
    img, d = _new_canvas()
    # Three stacked rectangles, the front one with index lines
    _rounded_rect(d, (20, 22, 76, 36), radius=4)
    _rounded_rect(d, (16, 38, 80, 54), radius=4)
    _rounded_rect(d, (12, 56, 84, 84), radius=6, fill=ACCENT_FILL)
    # Index tick marks on the front sheet
    for y in (66, 76):
        d.line((22, y, 30, y), fill=ACCENT, width=LINE - 2)
        d.line((36, y, 72, y), fill=ACCENT, width=LINE - 3)
    _save(img, 'bundle')


def settings_icon():
    """Gear — for 'Settings'."""
    img, d = _new_canvas()
    cx = cy = ICON_SIZE // 2
    # 8 teeth as small rounded rectangles around a centre disc
    import math
    teeth_outer = 40
    teeth_inner = 28
    teeth_w = 10
    for i in range(8):
        angle = (math.pi * 2) * (i / 8)
        # Draw a tooth as a small rectangle rotated; emulate by 4-point polygon
        ca, sa = math.cos(angle), math.sin(angle)
        # Centre of tooth midline:
        x1 = cx + teeth_inner * ca
        y1 = cy + teeth_inner * sa
        x2 = cx + teeth_outer * ca
        y2 = cy + teeth_outer * sa
        # Perpendicular offset
        px, py = -sa, ca
        w = teeth_w / 2
        poly = [
            (x1 + px * w, y1 + py * w),
            (x1 - px * w, y1 - py * w),
            (x2 - px * w, y2 - py * w),
            (x2 + px * w, y2 + py * w),
        ]
        d.polygon(poly, fill=ACCENT)
    # Outer disc and inner hole
    d.ellipse((cx - 30, cy - 30, cx + 30, cy + 30), outline=ACCENT, width=LINE,
              fill=ACCENT_FILL)
    d.ellipse((cx - 10, cy - 10, cx + 10, cy + 10), outline=ACCENT, width=LINE,
              fill=(0, 0, 0, 0))
    _save(img, 'settings')


def info_icon():
    """Circled 'i' — for 'Diagnostics'."""
    img, d = _new_canvas()
    cx = cy = ICON_SIZE // 2
    d.ellipse((cx - 36, cy - 36, cx + 36, cy + 36), outline=ACCENT, width=LINE,
              fill=ACCENT_FILL)
    # Dot
    d.ellipse((cx - 5, cy - 22, cx + 5, cy - 12), fill=ACCENT)
    # Stem
    d.rounded_rectangle((cx - 5, cy - 4, cx + 5, cy + 26), radius=3, fill=ACCENT)
    _save(img, 'info')


def annexure_icon():
    """List with a tick — for 'Annexure List'."""
    img, d = _new_canvas()
    _rounded_rect(d, (16, 14, 80, 86), radius=8, fill=ACCENT_FILL)
    # Row checks
    for y, ticked in ((30, True), (50, True), (70, False)):
        # Box
        d.rounded_rectangle((24, y - 6, 36, y + 6), radius=2, outline=ACCENT,
                            width=LINE - 3)
        if ticked:
            d.line((26, y, 30, y + 4), fill=ACCENT, width=LINE - 3)
            d.line((30, y + 4, 36, y - 4), fill=ACCENT, width=LINE - 3)
        # Line
        d.line((44, y, 72, y), fill=ACCENT, width=LINE - 3)
    _save(img, 'annexure')


def main():
    files_icon()
    folder_icon()
    outlook_icon()
    bundle_icon()
    settings_icon()
    info_icon()
    annexure_icon()
    print('generate_ui_icons: wrote icon_files / folder / outlook / bundle / '
          'settings / info / annexure to', HERE)


if __name__ == '__main__':
    main()
