"""MailWeave — persistent application settings."""

import json
import logging
import os
from copy import deepcopy
from dataclasses import dataclass, asdict, fields as dc_fields

_LOGGER = logging.getLogger('MailWeave')

_APPDATA     = os.environ.get('APPDATA') or os.path.expanduser('~')
SETTINGS_DIR  = os.path.join(_APPDATA, 'MailWeave')
SETTINGS_FILE = os.path.join(SETTINGS_DIR, 'settings.json')


# Bumped whenever a field is renamed or its meaning changes. Old settings
# files are migrated on load; if migration fails we reset rather than crash.
SETTINGS_SCHEMA_VERSION = 2


@dataclass
class AppSettings:
    # General
    strip_quotes:        bool = True
    sort_oldest_first:   bool = True
    date_format:         str  = 'uk'     # 'uk' | 'us' | 'iso'
    # Export
    pdf_page_size:       str  = 'A4'    # 'A4' | 'Letter' | 'Legal'
    include_cover_page:  bool = True
    include_email_index: bool = False
    safe_mode_import:    bool = True
    autosave_session:    bool = True
    confirm_before_clear: bool = True
    confirm_export_overwrite: bool = True
    document_author:     str  = ''
    default_export_dir:  str  = ''
    remember_export_dir: bool = True
    # Appearance
    theme:               str  = 'light'   # 'dark' | 'light'
    font_size:           str  = 'medium' # 'small' | 'medium' | 'large'
    show_preview:        bool = True

    # Persistence of window state — survives reinstalls.
    window_geometry:     str  = ''      # e.g. '1300x840+120+80'
    column_widths:       str  = ''      # 'num=42,date=166,...'

    # Update / telemetry — both opt-in. The check_updates flag controls a
    # single GET on startup; nothing is uploaded.
    check_updates:       bool = True
    update_url:          str  = ''      # JSON endpoint returning {"version":"…","url":"…"}

    # Schema version for forward-compat migrations.
    schema_version:      int  = SETTINGS_SCHEMA_VERSION

    # Legal
    disclaimer_accepted: bool = False


_VALID_KEYS = {f.name for f in dc_fields(AppSettings)}


def _migrate(raw: dict) -> dict:
    """Apply forward migrations so older settings.json files load cleanly."""
    version = int(raw.get('schema_version') or 1)
    if version < 2:
        # v1 → v2: introduced window_geometry, column_widths, check_updates,
        # update_url. Nothing to rename; defaults from the dataclass apply.
        version = 2
    raw['schema_version'] = SETTINGS_SCHEMA_VERSION
    return raw


def load_settings() -> AppSettings:
    try:
        with open(SETTINGS_FILE, encoding='utf-8') as fh:
            raw = json.load(fh)
        raw = _migrate(raw)
        s = AppSettings()
        for k, v in raw.items():
            if k in _VALID_KEYS:
                setattr(s, k, v)
        return s
    except Exception:
        return AppSettings()


def save_settings(s: AppSettings) -> None:
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as fh:
            json.dump(asdict(s), fh, indent=2)
    except Exception as exc:
        _LOGGER.warning('settings-save-failed error=%s', exc)


def reset_settings() -> AppSettings:
    return deepcopy(AppSettings())
