"""MailWeave — GitHub Releases auto-updater.

Two public entry points:

  check_for_update(current_version) -> UpdateInfo | None
      Fetch the latest release from `version.GITHUB_REPO`. Return metadata
      iff strictly newer than `current_version`. Network errors return None.

  download_installer(info, on_progress=None) -> Path
      Download the .exe asset to %TEMP%. Returns the file path on success.

Fails silently on network errors so a background check on app start can't
ever crash the UI.
"""

from __future__ import annotations

import json
import logging
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from urllib.error import URLError, HTTPError
from urllib.request import Request, urlopen

from version import GITHUB_REPO, __version__, version_tuple

_LOGGER = logging.getLogger('MailWeave.updater')

_API_URL = 'https://api.github.com/repos/{repo}/releases/latest'
_USER_AGENT = f'MailWeave/{__version__} (+https://github.com/{GITHUB_REPO})'
_INSTALLER_ASSET_PATTERN = re.compile(r'MailWeave-Setup-.*\.exe$', re.IGNORECASE)


@dataclass
class UpdateInfo:
    version: str          # '1.0.1' (no 'v' prefix)
    tag: str              # 'v1.0.1'
    download_url: str     # direct .exe asset URL
    asset_name: str       # 'MailWeave-Setup-1.0.1.exe'
    asset_size: int       # bytes
    release_notes: str    # markdown body of the release
    html_url: str         # human-friendly release page


def _parse_version(s: str) -> tuple[int, ...]:
    s = s.strip().lstrip('vV')
    parts: list[int] = []
    for piece in s.split('.'):
        digits = ''.join(c for c in piece if c.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _is_newer(latest: str, current: str) -> bool:
    try:
        return _parse_version(latest) > _parse_version(current)
    except Exception:
        return False


def check_for_update(current_version: str | None = None,
                     timeout: float = 5.0) -> Optional[UpdateInfo]:
    """Ask GitHub for the latest release. Return info iff strictly newer.

    `current_version` may be passed in (e.g. for tests); defaults to the
    bundled __version__. Returns None on any network/parsing error so callers
    can poll quietly.
    """
    if not GITHUB_REPO:
        return None
    current = current_version or __version__

    url = _API_URL.format(repo=GITHUB_REPO)
    req = Request(url, headers={
        'User-Agent': _USER_AGENT,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
    })
    try:
        with urlopen(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
    except (URLError, HTTPError, TimeoutError, ValueError) as exc:
        _LOGGER.info('update-check-failed url=%s err=%s', url, exc)
        return None
    except Exception:
        _LOGGER.exception('update-check-unexpected-error')
        return None

    tag = str(payload.get('tag_name') or '').strip()
    if not tag:
        return None
    latest_version = tag.lstrip('vV')

    if not _is_newer(latest_version, current):
        return None

    # Find the installer .exe asset. Releases without one mean the publisher
    # forgot to attach the binary — surface nothing rather than half-info.
    assets = payload.get('assets') or []
    installer = None
    for asset in assets:
        name = str(asset.get('name') or '')
        if _INSTALLER_ASSET_PATTERN.match(name):
            installer = asset
            break
    if installer is None:
        _LOGGER.info('update-found-but-no-asset tag=%s assets=%s',
                     tag, [a.get('name') for a in assets])
        return None

    return UpdateInfo(
        version=latest_version,
        tag=tag,
        download_url=str(installer.get('browser_download_url') or ''),
        asset_name=str(installer.get('name') or 'MailWeave-Setup.exe'),
        asset_size=int(installer.get('size') or 0),
        release_notes=str(payload.get('body') or ''),
        html_url=str(payload.get('html_url') or ''),
    )


def download_installer(info: UpdateInfo,
                       on_progress: Optional[Callable[[int, int], None]] = None,
                       chunk_size: int = 64 * 1024) -> Path:
    """Stream the installer .exe to %TEMP% and return its path.

    `on_progress(done, total)` is called periodically with byte counts; total
    may be 0 if the server omits Content-Length.
    """
    if not info.download_url:
        raise RuntimeError('No download URL on UpdateInfo.')

    dest = Path(tempfile.gettempdir()) / info.asset_name
    req = Request(info.download_url, headers={'User-Agent': _USER_AGENT})
    with urlopen(req, timeout=30) as resp:
        total = int(resp.headers.get('Content-Length') or info.asset_size or 0)
        done = 0
        with open(dest, 'wb') as fh:
            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                fh.write(chunk)
                done += len(chunk)
                if on_progress:
                    try:
                        on_progress(done, total)
                    except Exception:
                        pass

    if not dest.is_file() or dest.stat().st_size == 0:
        raise RuntimeError('Downloaded installer is empty.')
    return dest


def launch_installer(installer_path: Path) -> None:
    """Run the installer and let it take over. The current process should
    exit shortly after so the new build can replace files."""
    if os.name != 'nt':
        raise RuntimeError('Installer launch is Windows-only.')
    # Use os.startfile so the UAC prompt comes from the shell, not us.
    os.startfile(str(installer_path))  # type: ignore[attr-defined]
