"""MailWeave — email parsing utilities."""

import logging
import os
import re
import html
import hashlib
import email.utils
from email.header import decode_header, make_header
from dataclasses import dataclass, field
from datetime import datetime
from email.parser import BytesParser
from email import policy as email_policy
from typing import Optional, List

try:
    import extract_msg
    HAS_MSG = True
except ImportError:
    HAS_MSG = False

_LOGGER = logging.getLogger('MailWeave')


# ── Data model ─────────────────────────────────────────────────────────────────

@dataclass
class AttachmentData:
    filename: str
    size: int
    content_type: str
    content: bytes = field(default=b'', repr=False)
    cid: Optional[str] = None
    annexure_id: Optional[str] = None  # e.g., 'Annexure A'

@dataclass
class EmailData:
    message_id:  str
    subject:     str
    sender:      str
    recipients:  str
    date:        Optional[datetime]
    date_str:    str
    body_plain:  str        # original body text
    body_clean:  str        # body with quoted replies stripped
    source_file: str
    attachments: List[AttachmentData] = field(default_factory=list)
    unique_key:  str = field(default='', init=False)
    duplicate_fingerprint: str = field(default='', init=False)
    thread_id: str = field(default='', init=False)
    # Pre-lowercased concatenation of searchable fields. Computed once at
    # parse time so type-ahead search is O(1) per email per keystroke
    # instead of re-lowercasing four fields each frame.
    search_index: str = field(default='', init=False, repr=False)

    def __post_init__(self):
        mid = (self.message_id or '').strip('<> ')
        if mid and mid.lower() not in ('', 'unknown'):
            self.unique_key = mid
        else:
            raw = f'{self.sender}|{self.date_str}|{self.subject}'
            self.unique_key = hashlib.md5(
                raw.encode('utf-8', errors='replace')).hexdigest()
        digest_source = (
            f'{self.sender}|{self.recipients}|{self.subject}|'
            f'{self.date_str}|{(self.body_clean or self.body_plain or "")[:2000]}'
        )
        self.duplicate_fingerprint = hashlib.sha1(
            digest_source.encode('utf-8', errors='replace')
        ).hexdigest()
        
        # Normalize subject for threading
        s = self.subject or ''
        while True:
            old = s
            s = re.sub(r'^(Re|Fwd|Aw|Wg|Antwort|Forward|FW):\s*', '', s, flags=re.IGNORECASE).strip()
            if s == old: break
        self.thread_id = s.lower()

        self.search_index = (
            f'{self.subject or ""}\n{self.sender or ""}\n'
            f'{self.recipients or ""}\n{self.body_plain or ""}'
        ).lower()


# ── Text helpers ───────────────────────────────────────────────────────────────

def _strip_html(text: str) -> str:
    """Minimal HTML → plain-text conversion.

    Drops the *contents* of <script> and <style> tags (a naive `<[^>]+>`
    sub leaks CSS/JS text into the body), then collapses remaining tags.
    """
    if not text:
        return ''
    # Remove script/style with their contents (case-insensitive, multiline).
    text = re.sub(r'<script\b[^>]*>.*?</script\s*>', '',
                  text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<style\b[^>]*>.*?</style\s*>', '',
                  text, flags=re.IGNORECASE | re.DOTALL)
    # Drop HTML comments (including conditional ones).
    text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p\s*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    return html.unescape(text)


# Per-attachment size cap (bytes). Attachments larger than this are kept as
# metadata-only so we don't hold hundreds of MB per email in memory.
MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024  # 50 MB


def _decode_mime_header(value: str, fallback: str = '') -> str:
    try:
        if not value:
            return fallback
        return str(make_header(decode_header(value))).strip() or fallback
    except Exception:
        return str(value or fallback).strip()


def _decode_bytes(payload: bytes, charset: str | None) -> str:
    if payload is None:
        return ''

    candidates = [charset, 'utf-8', 'utf-8-sig', 'cp1252', 'latin-1']
    seen: set[str] = set()
    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return payload.decode(candidate)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode('utf-8', errors='replace')


def strip_quoted_text(body: str) -> str:
    """Strip quoted-reply content from an email body.

    Cuts everything from the first detected quote boundary to the end of the
    string.  Quote boundaries recognised:
      - Lines beginning with '>'
      - 'On … wrote:' reply headers
      - '--- Original/Forwarded Message ---' dividers
      - Outlook reply/forward header blocks (From: … Sent: … To: … Subject:)
      - Underline separators (5+ underscores/dashes)
    """
    lines = body.splitlines()

    _simple = [
        re.compile(r'^>'),
        re.compile(r'^On .{5,200} wrote:\s*$', re.IGNORECASE),
        re.compile(r'^-{3,}[\s\w]*(original|forwarded)[\s\w]*-{3,}',
                   re.IGNORECASE),
        re.compile(r'^[_\-]{5,}'),
    ]

    def _is_outlook_header(i: int) -> bool:
        # Require an exact `From:` reply-header line (with an address-looking
        # value), AND a `Sent:`/`Date:` line within the next four lines, AND a
        # `To:` or `Subject:` line shortly after. This avoids cutting normal
        # signatures or forwarded text the user wants to keep.
        if not re.match(r'^From:\s+\S', lines[i], re.IGNORECASE):
            return False
        window = lines[i + 1: i + 6]
        has_sent = any(re.match(r'^(Sent|Date):\s+', l, re.IGNORECASE) for l in window)
        has_to_or_subj = any(re.match(r'^(To|Subject):\s+', l, re.IGNORECASE) for l in window)
        return has_sent and has_to_or_subj

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        for pat in _simple:
            if pat.match(stripped):
                result = '\n'.join(lines[:i]).strip()
                return re.sub(r'\n{3,}', '\n\n', result)
        if i + 1 < len(lines) and _is_outlook_header(i):
            result = '\n'.join(lines[:i]).strip()
            return re.sub(r'\n{3,}', '\n\n', result)

    return re.sub(r'\n{3,}', '\n\n', body.strip())


# ── Parsers ────────────────────────────────────────────────────────────────────

def _parse_date(s: str) -> Optional[datetime]:
    if not s:
        return None
    try:
        return email.utils.parsedate_to_datetime(s)
    except Exception:
        return None


def parse_eml_file(filepath: str) -> Optional[EmailData]:
    try:
        with open(filepath, 'rb') as fh:
            msg = BytesParser(policy=email_policy.default).parse(fh)

        message_id = _decode_mime_header(msg.get('Message-ID', 'unknown'), 'unknown')
        subject    = _decode_mime_header(msg.get('Subject', '(No Subject)'), '(No Subject)')
        sender     = _decode_mime_header(msg.get('From', 'Unknown'), 'Unknown')
        recipients = _decode_mime_header(msg.get('To', ''), '')
        date_str   = _decode_mime_header(msg.get('Date', ''), '')
        date_obj   = _parse_date(date_str)

        plain = html_body = ''
        attachments = []
        
        if msg.is_multipart():
            for part in msg.walk():
                ct      = part.get_content_type()
                disp    = (part.get_content_disposition() or '').lower()
                
                if disp == 'attachment' or (part.get_filename() and ct not in ('text/plain', 'text/html')):
                    fname = _decode_mime_header(part.get_filename() or 'unnamed_attachment')
                    payload = part.get_payload(decode=True) or b''
                    size = len(payload)
                    cid = part.get('Content-ID')
                    if size > MAX_ATTACHMENT_BYTES:
                        _LOGGER.warning(
                            'attachment-too-large file=%s size=%s cap=%s — keeping metadata only',
                            fname, size, MAX_ATTACHMENT_BYTES,
                        )
                        payload = b''
                    attachments.append(AttachmentData(
                        filename=fname,
                        size=size,
                        content_type=ct,
                        content=payload,
                        cid=cid,
                    ))
                    continue
                    
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                decoded = _decode_bytes(payload, part.get_content_charset())
                if ct == 'text/plain' and not plain:
                    plain = decoded
                elif ct == 'text/html' and not html_body:
                    html_body = decoded
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                content = _decode_bytes(payload, msg.get_content_charset())
                if msg.get_content_type() == 'text/html':
                    html_body = content
                else:
                    plain = content

        body       = plain if plain.strip() else _strip_html(html_body)
        body_clean = strip_quoted_text(body)

        return EmailData(
            message_id=message_id, subject=subject,
            sender=sender, recipients=recipients,
            date=date_obj, date_str=date_str,
            body_plain=body, body_clean=body_clean,
            source_file=filepath,
            attachments=attachments,
        )
    except Exception as exc:
        _LOGGER.warning('eml-parse-failed path=%s error=%s', filepath, exc)
        return None


def parse_msg_file(filepath: str) -> Optional[EmailData]:
    if not HAS_MSG:
        return None
    try:
        msg = extract_msg.Message(filepath)

        message_id = str(getattr(msg, 'messageId', None) or 'unknown')
        subject    = str(msg.subject or '(No Subject)')
        sender     = str(msg.sender or 'Unknown')
        recipients = str(msg.to or '')
        date_obj   = msg.date
        date_str   = str(date_obj) if date_obj else ''

        body_raw = msg.body or ''
        html_raw = getattr(msg, 'htmlBody', None) or b''
        if isinstance(html_raw, bytes):
            html_raw = html_raw.decode('utf-8', errors='replace')

        body = body_raw if body_raw.strip() else _strip_html(html_raw)
        body_clean = strip_quoted_text(body)

        attachments = []
        for att in getattr(msg, 'attachments', []):
            try:
                # extract_msg attachment object
                fname = att.longFilename or att.shortFilename or 'unnamed'
                data = getattr(att, 'data', b'') or b''
                size = len(data)
                ctype = getattr(att, 'mimetype', 'application/octet-stream')
                if size > MAX_ATTACHMENT_BYTES:
                    _LOGGER.warning(
                        'attachment-too-large file=%s size=%s cap=%s — keeping metadata only',
                        fname, size, MAX_ATTACHMENT_BYTES,
                    )
                    data = b''
                attachments.append(AttachmentData(
                    filename=str(fname),
                    size=size,
                    content_type=str(ctype),
                    content=data,
                ))
            except Exception:
                continue

        msg.close()

        return EmailData(
            message_id=message_id, subject=subject,
            sender=sender, recipients=recipients,
            date=date_obj, date_str=date_str,
            body_plain=body, body_clean=body_clean,
            source_file=filepath,
            attachments=attachments,
        )
    except Exception as exc:
        _LOGGER.warning('msg-parse-failed path=%s error=%s', filepath, exc)
        return None


def parse_email_file(filepath: str) -> Optional[EmailData]:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.msg':
        return parse_msg_file(filepath)
    if ext == '.eml':
        return parse_eml_file(filepath)
    return None


# ── Sorting ────────────────────────────────────────────────────────────────────

def email_timestamp(e: EmailData) -> float:
    if e.date is None:
        return 0.0
    try:
        return e.date.timestamp()
    except Exception:
        return 0.0
