"""MailWeave — single source of truth for the version string.

Bump this on every release. The PyInstaller spec, the Inno Setup installer,
the About dialog, and the auto-updater all read from here.
"""

from __future__ import annotations

__version__ = '1.0.1'

# GitHub repository the updater queries for new releases.
# Set to 'owner/repo'. Leave empty to disable the update check.
GITHUB_REPO = 'MHB2730/mailweave'

# Tag prefix used when cutting releases (we publish tags like 'v1.0.0').
RELEASE_TAG_PREFIX = 'v'


def version_tuple() -> tuple[int, ...]:
    """Return the version as a tuple of ints for ordered comparison."""
    parts = []
    for piece in __version__.split('.'):
        try:
            parts.append(int(piece))
        except ValueError:
            # Pre-release suffixes (e.g. '1.0.0rc1') compare as the integer
            # prefix; the suffix is ignored. Good enough for our use case.
            digits = ''.join(c for c in piece if c.isdigit())
            parts.append(int(digits) if digits else 0)
    return tuple(parts)
