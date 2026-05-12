# -*- mode: python ; coding: utf-8 -*-
"""
MailWeave — PyInstaller spec file.
Produces a one-directory bundle in dist/MailWeave/.
"""

import os, sys
from PyInstaller.utils.hooks import collect_data_files, collect_all

# ── Source directory (where this spec lives) ───────────────────────────────────
HERE = os.path.abspath(SPECPATH)

# ── Collect extra data / hidden imports for tricky packages ───────────────────

datas      = []
binaries   = []
hiddenimps = []

# tkinterdnd2 — needs its tkdnd DLL + Tcl scripts bundled
datas += collect_data_files('tkinterdnd2', include_py_files=False)
datas += [
    (os.path.join(HERE, 'mailweave_logo.png'), '.'),
    (os.path.join(HERE, 'mailweave.ico'), '.'),
    (os.path.join(HERE, 'icon_*.png'), '.'),
]

# reportlab — ships fonts and XML colour configs as package data
rl_d, rl_b, rl_h = collect_all('reportlab')
datas     += rl_d
binaries  += rl_b
hiddenimps += rl_h

# extract_msg + its dependencies
em_d, em_b, em_h = collect_all('extract_msg')
datas     += em_d
binaries  += em_b
hiddenimps += em_h

# RTFDE (RTF parsing used by extract_msg)
try:
    rtf_d, rtf_b, rtf_h = collect_all('RTFDE')
    datas += rtf_d; binaries += rtf_b; hiddenimps += rtf_h
except Exception:
    pass

# bs4 (beautifulsoup4 — used by extract_msg)
bs4_d, bs4_b, bs4_h = collect_all('bs4')
datas += bs4_d; binaries += bs4_b; hiddenimps += bs4_h

# python-docx
try:
    docx_d, docx_b, docx_h = collect_all('docx')
    datas += docx_d; binaries += docx_b; hiddenimps += docx_h
except Exception:
    pass

# pypdf — required by the Indexed Bundle feature
try:
    pp_d, pp_b, pp_h = collect_all('pypdf')
    datas += pp_d; binaries += pp_b; hiddenimps += pp_h
except Exception:
    pass

# Pillow — required for image sources in the bundle builder
try:
    pil_d, pil_b, pil_h = collect_all('PIL')
    datas += pil_d; binaries += pil_b; hiddenimps += pil_h
except Exception:
    pass

# docx2pdf — optional Word→PDF conversion in the bundle builder
try:
    d2p_d, d2p_b, d2p_h = collect_all('docx2pdf')
    datas += d2p_d; binaries += d2p_b; hiddenimps += d2p_h
except Exception:
    pass

# tzdata / tzlocal (timezones, used by extract_msg date parsing)
try:
    tz_d, _, tz_h = collect_all('tzdata')
    datas += tz_d; hiddenimps += tz_h
except Exception:
    pass

# Additional hidden imports that static analysis misses
hiddenimps += [
    'tkinterdnd2',
    'extract_msg',
    'olefile',
    'compressed_rtf',
    'ebcdic',
    'colorclass',
    'tzlocal',
    'msoffcrypto',
    'email',
    'email.mime',
    'email.mime.multipart',
    'email.mime.text',
    'email.mime.base',
    'email.utils',
    'email.parser',
    'email.policy',
    'html',
    'html.parser',
    'hashlib',
    'json',
    'dataclasses',
    'copy',
    # Our own modules (spec-based builds may need them explicit)
    'settings',
    'themes',
    'email_parser',
    'exporters',
    'splash',
    'options',
    'app',
    'diagnostics',
    'session_store',
    'annexures',
    'outlook_import',
    'brand_assets',
    'bundle_builder',
    'bundle_dialog',
]

# ── Analysis ───────────────────────────────────────────────────────────────────

a = Analysis(
    [os.path.join(HERE, 'main.py')],
    pathex=[HERE],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimps,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'scipy', 'pandas',
        'IPython', 'jupyter', 'notebook',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'wx', 'gi',
        '_tkinter.test',
        'test', 'tests', 'unittest',
    ],
    noarchive=False,
)

# ── Packaging ──────────────────────────────────────────────────────────────────

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # onedir: keep DLLs separate
    name='MailWeave',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # UPX not required; skip to keep it simple
    console=False,                  # no black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(HERE, 'mailweave.ico'),
    version=os.path.join(HERE, 'version_info.txt'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='MailWeave',
)
