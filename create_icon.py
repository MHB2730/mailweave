"""Generate MailWeave icon assets from the reference logo."""

from __future__ import annotations

import os

from brand_assets import extract_logo_square

SIZES = [256, 128, 64, 48, 32, 16]


def create_ico(out_path: str = 'mailweave.ico'):
    frames = [extract_logo_square(size) for size in SIZES]
    frames[0].save(
        out_path,
        format='ICO',
        sizes=[(size, size) for size in SIZES],
        append_images=frames[1:],
    )
    print(f'Created {out_path}')


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    create_ico()
