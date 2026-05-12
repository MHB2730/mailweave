"""Session autosave and recovery for MailWeave."""

from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime

import diagnostics
from email_parser import AttachmentData, EmailData


_EMAIL_FIELDS = (
    'message_id',
    'subject',
    'sender',
    'recipients',
    'date',
    'date_str',
    'body_plain',
    'body_clean',
    'source_file',
    'attachments',
)


def _session_file():
    diagnostics.ensure_runtime_dirs()
    return diagnostics.RECOVERY_DIR / 'last_session.json'


def _lock_file():
    diagnostics.ensure_runtime_dirs()
    return diagnostics.RECOVERY_DIR / 'session.lock'


def mark_session_open():
    diagnostics.ensure_runtime_dirs()
    _lock_file().write_text(datetime.now().isoformat(timespec='seconds'), encoding='utf-8')


def clear_session_lock():
    try:
        _lock_file().unlink(missing_ok=True)
    except Exception:
        diagnostics.LOGGER.exception('session-lock-clear-failed')


def save_session(emails: list[EmailData], settings_dict: dict | None = None):
    diagnostics.ensure_runtime_dirs()
    serialized = []
    for email in emails:
        raw = asdict(email)
        if raw.get('date') is not None:
            raw['date'] = raw['date'].isoformat()

        # Exclude attachment content (bytes) from JSON serialization. We also
        # record `had_content` so restore can tell users which attachments
        # need re-importing to be openable/exportable.
        if raw.get('attachments'):
            for att in raw['attachments']:
                content = att.pop('content', None)
                att['had_content'] = bool(content)

        serialized.append(raw)
    payload = {
        'saved_at': datetime.now().isoformat(timespec='seconds'),
        'count': len(emails),
        'emails': serialized,
        'settings': settings_dict or {},
    }
    # Write atomically via a temp file so a crash mid-write can't leave a
    # truncated JSON on disk.
    target = _session_file()
    tmp = target.with_suffix(target.suffix + '.tmp')
    try:
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        tmp.replace(target)
    except Exception:
        diagnostics.LOGGER.exception('session-save-failed')
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def has_recovery_session() -> bool:
    return _session_file().is_file() and _lock_file().is_file()


def load_session() -> dict | None:
    session_file = _session_file()
    if not session_file.is_file():
        return None
    try:
        payload = json.loads(session_file.read_text(encoding='utf-8'))
        emails = []
        for raw in payload.get('emails', []):
            date_value = raw.get('date')
            if date_value:
                try:
                    raw['date'] = datetime.fromisoformat(date_value)
                except Exception:
                    raw['date'] = None
            attachments = []
            for att_raw in raw.get('attachments') or []:
                try:
                    attachments.append(AttachmentData(
                        filename=att_raw.get('filename', ''),
                        size=int(att_raw.get('size') or 0),
                        content_type=att_raw.get('content_type', ''),
                        cid=att_raw.get('cid'),
                    ))
                except Exception:
                    continue
            kwargs = {field: raw.get(field) for field in _EMAIL_FIELDS if field != 'attachments'}
            kwargs['attachments'] = attachments
            try:
                emails.append(EmailData(**kwargs))
            except Exception:
                diagnostics.LOGGER.exception('session-email-restore-failed')
                continue
        payload['emails'] = emails
        return payload
    except Exception:
        diagnostics.LOGGER.exception('session-load-failed')
        return None


def clear_session():
    try:
        _session_file().unlink(missing_ok=True)
    except Exception:
        diagnostics.LOGGER.exception('session-file-clear-failed')
    clear_session_lock()
