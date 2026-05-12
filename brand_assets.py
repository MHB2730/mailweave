"""Shared brand asset helpers for MailWeave."""

from __future__ import annotations

import os
import sys
import tkinter as tk

from PIL import Image, ImageOps, ImageTk

LOGO_FILE = 'mailweave_logo.png'
ICON_FILE = 'mailweave.ico'


def resource_path(relative: str) -> str:
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def logo_path() -> str:
    return resource_path(LOGO_FILE)


def icon_path() -> str:
    return resource_path(ICON_FILE)


def load_logo_image(max_size: tuple[int, int]) -> Image.Image | None:
    path = logo_path()
    if not os.path.isfile(path):
        return None
    image = Image.open(path).convert('RGBA')
    image.thumbnail(max_size, Image.LANCZOS)
    return image


def load_logo_photo(max_size: tuple[int, int]) -> ImageTk.PhotoImage | None:
    image = load_logo_image(max_size)
    if image is None:
        return None
    return ImageTk.PhotoImage(image)


def load_ui_icon(name: str, size: tuple[int, int] = (18, 18)) -> ImageTk.PhotoImage | None:
    path = resource_path(f'icon_{name}.png')
    if not os.path.isfile(path):
        return None
    try:
        img = Image.open(path).convert('RGBA')
        img.thumbnail(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


def extract_mark_square(size: int = 256, padding: int = 16) -> Image.Image:
    path = logo_path() if os.path.isfile(logo_path()) else LOGO_FILE
    image = Image.open(path).convert('RGBA')

    upper = image.crop((0, 0, image.width, int(image.height * 0.58)))
    alpha = Image.new('L', upper.size, 0)
    pixels = upper.load()
    for y in range(upper.height):
        for x in range(upper.width):
            r, g, b, _a = pixels[x, y]
            if max(r, g, b) < 245:
                alpha.putpixel((x, y), 255)
    bbox = alpha.getbbox() or (0, 0, upper.width, upper.height)
    mark = upper.crop(bbox)
    canvas = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    effective_padding = min(padding, max(2, size // 8))
    inner_size = max(1, size - effective_padding * 2)
    mark = ImageOps.contain(mark, (inner_size, inner_size), Image.LANCZOS)
    x = (size - mark.width) // 2
    y = (size - mark.height) // 2
    canvas.paste(mark, (x, y), mark)
    return canvas


def extract_logo_square(size: int = 256, padding: int = 12) -> Image.Image:
    path = logo_path() if os.path.isfile(logo_path()) else LOGO_FILE
    image = Image.open(path).convert('RGBA')

    # Detect the mark: any pixel that is NOT white
    alpha = Image.new('L', image.size, 0)
    pixels = image.load()
    for y in range(image.height):
        for x in range(image.width):
            r, g, b, a = pixels[x, y]
            # If not white (using a safer threshold for JPEG-like artifacts)
            if r < 250 or g < 250 or b < 250:
                alpha.putpixel((x, y), 255)

    bbox = alpha.getbbox() or (0, 0, image.width, image.height)
    logo = image.crop(bbox)
    
    # Create a new transparent canvas
    effective_padding = min(padding, max(2, size // 10))
    inner_size = max(1, size - effective_padding * 2)
    logo = ImageOps.contain(logo, (inner_size, inner_size), Image.LANCZOS)

    canvas = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    x = (size - logo.width) // 2
    y = (size - logo.height) // 2
    
    # We need to preserve the colors but ensure the background we cropped is transparent
    # The logo we cropped still has white pixels around the edges if it wasn't perfectly cropped
    # So we'll use the alpha we just calculated to mask it
    logo_alpha = Image.new('L', logo.size, 0)
    l_pix = logo.load()
    for y_l in range(logo.height):
        for x_l in range(logo.width):
            r, g, b, a = l_pix[x_l, y_l]
            if r < 250 or g < 250 or b < 250:
                logo_alpha.putpixel((x_l, y_l), 255)
    
    canvas.paste(logo, (x, y), logo_alpha)
    return canvas


def set_window_icon(root: tk.Tk | tk.Toplevel):
    path = icon_path()
    if os.path.isfile(path):
        try:
            root.iconbitmap(path)
        except Exception:
            pass
