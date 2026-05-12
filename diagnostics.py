"""Diagnostics, logging, and startup checks for MailWeave."""

from __future__ import annotations

import json
import logging
import os
import platform
import sys
import tempfile
import traceback
from dataclasses import asdict, dataclass
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Iterable

APP_NAME = 'MailWeave'
APPDATA_DIR = Path(os.environ.get('APPDATA') or Path.home()) / APP_NAME
LOG_DIR = APPDATA_DIR / 'logs'
RECOVERY_DIR = APPDATA_DIR / 'recovery'
CRASH_DIR = APPDATA_DIR / 'crash_reports'
TEMP_DIR = Path(tempfile.gettempdir()) / 'MailWeave'


@dataclass
class CheckResult:
    name: str
    ok: bool
    detail: str


def ensure_runtime_dirs():
    global APPDATA_DIR, LOG_DIR, RECOVERY_DIR, CRASH_DIR
    for path in (APPDATA_DIR, LOG_DIR, RECOVERY_DIR, CRASH_DIR, TEMP_DIR):
        try:
            path.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            fallback_root = TEMP_DIR / 'runtime'
            APPDATA_DIR = fallback_root
            LOG_DIR = fallback_root / 'logs'
            RECOVERY_DIR = fallback_root / 'recovery'
            CRASH_DIR = fallback_root / 'crash_reports'
            for fallback in (APPDATA_DIR, LOG_DIR, RECOVERY_DIR, CRASH_DIR, TEMP_DIR):
                fallback.mkdir(parents=True, exist_ok=True)
            return


def _log_path() -> Path:
    ensure_runtime_dirs()
    return LOG_DIR / 'mailweave.log'


def setup_logging() -> logging.Logger:
    ensure_runtime_dirs()
    logger = logging.getLogger(APP_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '%(asctime)s %(levelname)s %(name)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    file_handler = RotatingFileHandler(
        _log_path(),
        maxBytes=1_000_000,
        backupCount=5,
        encoding='utf-8',
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    logger.info('logging-initialized')
    return logger


LOGGER = setup_logging()


def prune_old_logs(days: int = 14):
    ensure_runtime_dirs()
    threshold = datetime.now().timestamp() - (days * 86400)
    for path in LOG_DIR.glob('*'):
        try:
            if path.is_file() and path.stat().st_mtime < threshold:
                path.unlink()
        except Exception:
            LOGGER.exception('log-prune-failed path=%s', path)


def log_exception(context: str, exc: BaseException):
    LOGGER.error('%s: %s\n%s', context, exc, traceback.format_exc())


def install_exception_hooks():
    def _hook(exc_type, exc, tb):
        LOGGER.critical('uncaught-exception: %s\n%s', exc, ''.join(traceback.format_exception(exc_type, exc, tb)))
        try:
            from tkinter import messagebox

            messagebox.showerror(
                'Unexpected error',
                'MailWeave hit an unexpected error.\n\nA diagnostic log has been written to:\n'
                f'{_log_path()}',
            )
        except Exception:
            pass

    sys.excepthook = _hook


def run_startup_checks() -> list[CheckResult]:
    ensure_runtime_dirs()
    results: list[CheckResult] = []

    def _add(name: str, ok: bool, detail: str):
        results.append(CheckResult(name, ok, detail))

    try:
        ensure_runtime_dirs()
        _add('runtime_dirs', True, f'Using {APPDATA_DIR}')
    except Exception as exc:
        _add('runtime_dirs', False, f'Could not create runtime directories: {exc}')

    try:
        temp_file = TEMP_DIR / 'write_test.tmp'
        temp_file.write_text('ok', encoding='utf-8')
        temp_file.unlink(missing_ok=True)
        _add('temp_dir', True, f'Writable temp directory: {TEMP_DIR}')
    except Exception as exc:
        _add('temp_dir', False, f'Temp directory is not writable: {exc}')

    try:
        from exporters import HAS_DOCX, HAS_PDF

        _add('pdf_engine', HAS_PDF, 'ReportLab available' if HAS_PDF else 'ReportLab missing')
        _add('docx_engine', HAS_DOCX, 'python-docx available' if HAS_DOCX else 'python-docx missing')
    except Exception as exc:
        _add('export_engines', False, f'Could not inspect export engines: {exc}')

    try:
        from outlook_import import HAS_OUTLOOK_IMPORT

        _add(
            'outlook_import',
            HAS_OUTLOOK_IMPORT,
            'Outlook import available' if HAS_OUTLOOK_IMPORT else 'pywin32 / Outlook import unavailable',
        )
    except Exception as exc:
        _add('outlook_import', False, f'Could not inspect Outlook integration: {exc}')

    LOGGER.info('startup-checks %s', json.dumps([asdict(result) for result in results]))
    return results


def summarize_checks(results: Iterable[CheckResult]) -> tuple[list[str], list[str]]:
    ok = [f'{item.name}: {item.detail}' for item in results if item.ok]
    failed = [f'{item.name}: {item.detail}' for item in results if not item.ok]
    return ok, failed


def write_crash_report(exc: BaseException, tb_text: str, context: dict | None = None) -> Path:
    """Persist a crash report so the next launch can surface it.

    Reports are plain text — no upload, no PII the user didn't already have.
    The app's "View crash reports" dialog handles disclosure.
    """
    ensure_runtime_dirs()
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = CRASH_DIR / f'crash_{stamp}.txt'
    payload = build_diagnostic_report({'exception_type': type(exc).__name__,
                                       'exception': str(exc),
                                       'context': context or {}})
    body = f'{payload}\n\nTraceback:\n{tb_text}\n'
    try:
        path.write_text(body, encoding='utf-8')
    except Exception:
        LOGGER.exception('crash-report-write-failed')
    return path


def pending_crash_reports() -> list[Path]:
    ensure_runtime_dirs()
    try:
        return sorted(CRASH_DIR.glob('crash_*.txt'))
    except Exception:
        return []


def clear_crash_reports() -> int:
    count = 0
    for path in pending_crash_reports():
        try:
            path.unlink()
            count += 1
        except Exception:
            pass
    return count


def check_for_update(current_version: str, url: str, timeout: float = 4.0) -> dict | None:
    """Fetch update metadata. Returns dict with 'version' and optional 'url'
    if a strictly newer version is available, otherwise None.

    Designed to fail silently: no exception escapes, no logging at error level
    for the common cases (network down, endpoint not configured)."""
    if not url:
        return None
    try:
        import json as _json
        from urllib.request import Request, urlopen
        req = Request(url, headers={
            'User-Agent': f'MailWeave/{current_version}',
            'Accept': 'application/json',
        })
        with urlopen(req, timeout=timeout) as resp:
            data = _json.loads(resp.read().decode('utf-8'))
    except Exception:
        # Network errors are not interesting — the user opted in to a check,
        # not to a guarantee.
        return None

    latest = str(data.get('version') or '').strip()
    if not latest:
        return None

    def _tuplise(v: str) -> tuple:
        parts = []
        for chunk in v.lstrip('v').split('.'):
            digits = ''.join(c for c in chunk if c.isdigit())
            parts.append(int(digits) if digits else 0)
        return tuple(parts)

    try:
        if _tuplise(latest) <= _tuplise(current_version):
            return None
    except Exception:
        return None
    return {'version': latest, 'url': str(data.get('url') or '')}


def build_diagnostic_report(extra: dict | None = None) -> str:
    payload = {
        'app': APP_NAME,
        'python': sys.version,
        'platform': platform.platform(),
        'executable': sys.executable,
        'cwd': os.getcwd(),
        'appdata_dir': str(APPDATA_DIR),
        'log_file': str(_log_path()),
        'timestamp': datetime.now().isoformat(timespec='seconds'),
    }
    if extra:
        payload.update(extra)
    return json.dumps(payload, indent=2)
