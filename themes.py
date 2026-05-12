"""MailWeave — colour themes."""

from __future__ import annotations

THEMES: dict[str, dict[str, str]] = {
    'dark': {
        'name':     'dark',
        'bg':       '#09090B',
        'sidebar':  '#111111',
        'surface':  '#18181B',
        'surface2': '#27272A',
        'surface3': '#3F3F46',
        'accent':   '#84cc16',        # Lime Green
        'accent2':  '#a3e635',
        'accent3':  '#10B981',
        'fg':       '#FAFAFA',
        'fgsub':    '#A1A1AA',
        'fgdim':    '#52525B',
        'border':   '#27272A',
        'error':    '#EF4444',
        'ok':       '#10B981',
        'sel_bg':   '#27272A',
        'sel_fg':   '#FFFFFF',
        'menu_bg':  '#111111',
        'menu_fg':  '#F4F4F5',
        'hero':     '#09090B',
        'hero2':    '#18181B',
        'card':     '#09090B',
        'glow':     '#84cc16',
    },
    'light': {
        'name':     'light',
        'bg':       '#FFFFFF',
        'sidebar':  '#F7FEE7',        # Very light lime sidebar
        'surface':  '#FFFFFF',
        'surface2': '#ECFCCB',        # Light lime surface
        'surface3': '#D9F99D',
        'accent':   '#84cc16',        # Lime Green
        'accent2':  '#65a30d',
        'accent3':  '#4d7c0f',
        'fg':       '#14532d',        # Dark Green text
        'fgsub':    '#166534',
        'fgdim':    '#15803d',
        'border':   '#BEF264',
        'error':    '#DC2626',
        'ok':       '#166534',
        'sel_bg':   '#ECFCCB',
        'sel_fg':   '#14532d',
        'menu_bg':  '#FFFFFF',
        'menu_fg':  '#14532d',
        'hero':     '#F7FEE7',
        'hero2':    '#ECFCCB',
        'card':     '#FFFFFF',
        'glow':     '#84cc16',
    },
}


def get_theme(name: str) -> dict[str, str]:
    return THEMES.get(name, THEMES['light'])
