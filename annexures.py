"""MailWeave — Annexure management logic."""

from __future__ import annotations
from typing import TYPE_CHECKING, List, Optional
import string

if TYPE_CHECKING:
    from email_parser import EmailData, AttachmentData

def get_annexure_label(index: int) -> str:
    """Return a label like A, B, ... Z, AA, AB..."""
    if index < 0:
        return ""
    res = ""
    while index >= 0:
        res = string.ascii_uppercase[index % 26] + res
        index = (index // 26) - 1
    return res


def annexure_anchor_id(att: 'AttachmentData') -> str:
    """Stable in-document anchor id for a PDF/Word cross-reference."""
    label = (att.annexure_id or '').replace(' ', '_')
    return f'mw_annex_{label or id(att)}'

def auto_assign_annexures(emails: List[EmailData]):
    """
    Automatically assign Annexure IDs to all attachments in a list of emails.
    Follows the order of emails and their internal attachment order.
    """
    counter = 0
    for email in emails:
        for att in email.attachments:
            att.annexure_id = f"Annexure {get_annexure_label(counter)}"
            counter += 1

def get_all_attachments(emails: List[EmailData]) -> List[tuple[EmailData, AttachmentData]]:
    """Return a flattened list of (email, attachment) pairs."""
    results = []
    for email in emails:
        for att in email.attachments:
            results.append((email, att))
    return results

def format_annexure_reference(att: AttachmentData) -> str:
    """Format a reference string for inclusion in text."""
    if att.annexure_id:
        return f"[{att.annexure_id}: {att.filename}]"
    return f"[Attachment: {att.filename}]"
