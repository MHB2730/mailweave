"""Outlook selection import helpers for MailWeave."""

from __future__ import annotations

import logging
import shutil
import tempfile
from pathlib import Path

from email_parser import EmailData, parse_email_file

_LOGGER = logging.getLogger('MailWeave')

try:
    import pythoncom
    import win32com.client

    HAS_OUTLOOK_IMPORT = True
except ImportError:
    HAS_OUTLOOK_IMPORT = False


class OutlookImportError(RuntimeError):
    """Raised when Outlook import is unavailable or cannot proceed."""


def _get_selection():
    if not HAS_OUTLOOK_IMPORT:
        raise OutlookImportError('Outlook import support is not installed.')

    try:
        app = win32com.client.Dispatch('Outlook.Application')
    except Exception as exc:
        # Check if it's a COM error (Invalid class string usually means Outlook missing)
        if 'Invalid class string' in str(exc) or '-2147221005' in str(exc):
            raise OutlookImportError('Microsoft Outlook is not installed or not properly registered.') from exc
        raise OutlookImportError(f'Could not connect to Outlook: {exc}') from exc

    try:
        explorer = app.ActiveExplorer()
    except Exception as exc:
        raise OutlookImportError(f'Could not access Outlook explorer: {exc}') from exc

    if explorer is None:
        raise OutlookImportError('Open Outlook and select one or more emails first.')

    try:
        selection = explorer.Selection
        count = int(getattr(selection, 'Count', 0) or 0)
    except Exception as exc:
        raise OutlookImportError(f'Could not access Outlook selection: {exc}') from exc

    if count <= 0:
        raise OutlookImportError('Select one or more emails in Outlook first.')
    return selection, count


def get_outlook_selection_count() -> int:
    if not HAS_OUTLOOK_IMPORT:
        raise OutlookImportError('Outlook import support is not installed.')

    pythoncom.CoInitialize()
    try:
        _selection, count = _get_selection()
        return count
    finally:
        pythoncom.CoUninitialize()


def iter_selected_outlook_emails():
    """Yield parsed EmailData objects from the current Outlook selection."""
    if not HAS_OUTLOOK_IMPORT:
        raise OutlookImportError('Outlook import support is not installed.')

    temp_dir = Path(tempfile.mkdtemp(prefix='mailweave_outlook_'))
    pythoncom.CoInitialize()
    try:
        selection, count = _get_selection()
        for index in range(1, count + 1):
            item = selection.Item(index)
            if int(getattr(item, 'Class', 0) or 0) != 43:
                yield None
                continue

            msg_path = temp_dir / f'selected_{index}.msg'
            try:
                item.SaveAs(str(msg_path), 3)
                yield parse_email_file(str(msg_path))
            except Exception as exc:
                _LOGGER.warning(
                    'outlook-item-save-failed index=%s error=%s', index, exc,
                )
                yield None
            finally:
                try:
                    msg_path.unlink(missing_ok=True)
                except Exception:
                    pass
    finally:
        pythoncom.CoUninitialize()
        shutil.rmtree(temp_dir, ignore_errors=True)
