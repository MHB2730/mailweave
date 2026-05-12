"""MailWeave — Index-driven document bundler.

Takes a list of `BundleEntry(label, title, source_path)` and produces a single
PDF: optional cover → clickable index → each source document, stamped per-page
with the annexure label, and (optionally) a continuous bundle-wide page number.

Source types supported:
  - .pdf                 native passthrough
  - .docx                converted via docx2pdf (Word COM, Windows-only)
  - .png/.jpg/.jpeg/.tif image → single-page PDF via Pillow + ReportLab

The engine is pure logic. UI lives in bundle_dialog.py.
"""

from __future__ import annotations

import io
import logging
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from annexures import get_annexure_label

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import (
        ArrayObject,
        DictionaryObject,
        Fit,
        NameObject,
        NumberObject,
        TextStringObject,
    )
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from reportlab.lib.pagesizes import A4, letter, legal
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.platypus import (
        BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
    )
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from docx2pdf import convert as docx2pdf_convert
    HAS_DOCX2PDF = True
except Exception:
    HAS_DOCX2PDF = False


_LOGGER = logging.getLogger('MailWeave')

PAGE_SIZES = {'A4': None, 'Letter': None, 'Legal': None}


@dataclass
class BundleEntry:
    label: str                # 'A', 'B', 'AA' — without the 'Annexure ' prefix
    title: str                # Description for the index page
    source_path: str          # Original file the user picked
    manual_label: bool = False  # True when the user pinned this label explicitly
    # Filled in during build:
    start_page: int = 0       # 1-based bundle page where this annexure begins
    page_count: int = 0
    error: str = ''


@dataclass
class BundleOptions:
    output_path: str
    page_size: str = 'A4'                 # 'A4' | 'Letter' | 'Legal'
    include_cover: bool = True
    cover_title: str = 'Indexed Bundle'
    cover_author: str = 'MailWeave'
    include_bundle_pagination: bool = True
    stamp_per_annexure_pagination: bool = True
    footer_text: str = ''                 # extra left-aligned text on every page
    # Build-time outputs (filled in by build_bundle):
    total_pages: int = 0
    entries: list[BundleEntry] = field(default_factory=list)


class BundleBuildError(RuntimeError):
    pass


# ── Source conversion ──────────────────────────────────────────────────────────

_SUPPORTED_EXTS = {'.pdf', '.docx', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif'}


def supported_source(path: str) -> bool:
    return Path(path).suffix.lower() in _SUPPORTED_EXTS


def _convert_docx(src: Path, workdir: Path) -> Path:
    if not HAS_DOCX2PDF:
        raise BundleBuildError(
            f'Cannot convert "{src.name}": docx2pdf is not installed. '
            'Install it with: pip install docx2pdf  (requires Microsoft Word).'
        )
    dest = workdir / f'{src.stem}.pdf'
    try:
        docx2pdf_convert(str(src), str(dest))
    except Exception as exc:
        raise BundleBuildError(f'Word conversion failed for "{src.name}": {exc}') from exc
    if not dest.is_file():
        raise BundleBuildError(f'Word conversion produced no file for "{src.name}".')
    return dest


def _convert_image(src: Path, workdir: Path, page_size) -> Path:
    if not HAS_PIL:
        raise BundleBuildError(
            f'Cannot embed image "{src.name}": Pillow is not installed.'
        )
    dest = workdir / f'{src.stem}.pdf'
    page_w, page_h = page_size
    try:
        with Image.open(src) as img:
            img = img.convert('RGB')
            iw, ih = img.size
        # Fit the image into the page with a 1cm margin, preserving aspect.
        margin = 1 * cm
        max_w = page_w - 2 * margin
        max_h = page_h - 2 * margin
        scale = min(max_w / iw, max_h / ih)
        dw, dh = iw * scale, ih * scale
        x = (page_w - dw) / 2
        y = (page_h - dh) / 2

        c = rl_canvas.Canvas(str(dest), pagesize=page_size)
        c.drawImage(str(src), x, y, width=dw, height=dh,
                    preserveAspectRatio=True, anchor='c', mask='auto')
        c.showPage()
        c.save()
    except Exception as exc:
        raise BundleBuildError(f'Image conversion failed for "{src.name}": {exc}') from exc
    return dest


def _ensure_pdf(src_path: str, workdir: Path, page_size) -> Path:
    """Return a PDF path for the given source, converting if necessary."""
    src = Path(src_path)
    if not src.is_file():
        raise BundleBuildError(f'Source not found: {src_path}')
    ext = src.suffix.lower()
    if ext == '.pdf':
        return src
    if ext == '.docx':
        return _convert_docx(src, workdir)
    if ext in {'.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif'}:
        return _convert_image(src, workdir, page_size)
    raise BundleBuildError(f'Unsupported source type: {src.name}')


# ── Stamping (overlay) ────────────────────────────────────────────────────────

def _build_stamp(page_w: float,
                 page_h: float,
                 annexure_label: str,
                 annexure_page: int,
                 annexure_total: int,
                 bundle_page: Optional[int],
                 left_text: str) -> 'PdfReader':
    """Render a single-page transparent overlay carrying the annexure label
    and page numbering in the TOP-RIGHT corner.

    Layout (top-right corner, stacked):
        Annexure <X>
        Page <annexure_page> of <annexure_total>
        Bundle page <bundle_page>          (only if enabled)

    Any optional `left_text` is placed in the top-left of the same band so
    legal teams can include a matter reference without it competing with the
    annexure marker.
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setFillColor(colors.HexColor('#1A2840'))

    margin_x = 1.5 * cm
    # Anchor the topmost line ~1cm down from the page top so it sits within
    # standard header margins of most office documents.
    top_baseline = page_h - 1.0 * cm
    line_gap = 0.42 * cm

    # Top-left: optional matter/case reference.
    if left_text:
        c.setFont('Helvetica', 8)
        c.setFillColor(colors.HexColor('#3F3F46'))
        c.drawString(margin_x, top_baseline, left_text)
        c.setFillColor(colors.HexColor('#1A2840'))

    # Top-right: annexure label (prominent), then page-of-N, then bundle page.
    right_x = page_w - margin_x

    c.setFont('Helvetica-Bold', 10)
    c.drawRightString(right_x, top_baseline, f'Annexure {annexure_label}')

    c.setFont('Helvetica', 8)
    c.setFillColor(colors.HexColor('#3F3F46'))

    next_y = top_baseline - line_gap
    if annexure_page > 0:
        c.drawRightString(
            right_x,
            next_y,
            f'Page {annexure_page} of {annexure_total}',
        )
        next_y -= line_gap

    if bundle_page is not None:
        c.drawRightString(
            right_x,
            next_y,
            f'Bundle page {bundle_page}',
        )

    c.save()
    buf.seek(0)
    return PdfReader(buf)


def _stamp_pdf_pages(reader: 'PdfReader',
                     writer: 'PdfWriter',
                     entry: BundleEntry,
                     options: BundleOptions,
                     bundle_page_start: int,
                     page_size: tuple) -> int:
    """Append every page of `reader` to `writer`, stamped with footer text.
    Returns the count of pages added.
    """
    added = 0
    total = len(reader.pages)
    for i, src_page in enumerate(reader.pages, start=1):
        page = src_page
        # Use the source page's own size for stamping so we don't distort it.
        try:
            box = page.mediabox
            pw = float(box.width)
            ph = float(box.height)
        except Exception:
            pw, ph = page_size

        bundle_page = bundle_page_start + added if options.include_bundle_pagination else None
        annex_page = i if options.stamp_per_annexure_pagination else 0
        if options.stamp_per_annexure_pagination or options.include_bundle_pagination or options.footer_text:
            stamp = _build_stamp(
                pw, ph,
                annexure_label=entry.label,
                annexure_page=annex_page,
                annexure_total=total,
                bundle_page=bundle_page,
                left_text=options.footer_text,
            )
            try:
                page.merge_page(stamp.pages[0])
            except Exception:
                _LOGGER.exception('stamp-merge-failed entry=%s page=%s', entry.label, i)
        writer.add_page(page)
        added += 1
    return added


# ── Cover + index PDF ─────────────────────────────────────────────────────────

def _build_cover_and_index_pdf(options: BundleOptions,
                                entries: list[BundleEntry],
                                page_size: tuple,
                                workdir: Path) -> Path:
    """Build a one-or-more-page PDF containing the cover (optional) and a
    clickable index. The index links point to anchors we'll embed by name.

    Returns the path of the generated PDF.
    """
    out = workdir / '_cover_index.pdf'

    class _Doc(BaseDocTemplate):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            frame = Frame(self.leftMargin, self.bottomMargin,
                          self.width, self.height, id='normal')
            self.addPageTemplates([PageTemplate(id='cover_index', frames=[frame])])

    doc = _Doc(str(out), pagesize=page_size,
               leftMargin=2.2 * cm, rightMargin=2.2 * cm,
               topMargin=2.2 * cm, bottomMargin=2.2 * cm,
               title=options.cover_title, author=options.cover_author)

    title_sty = ParagraphStyle('B_T', fontName='Helvetica-Bold', fontSize=22,
                               textColor=colors.HexColor('#1A2840'), spaceAfter=4)
    meta_sty = ParagraphStyle('B_M', fontName='Helvetica', fontSize=9,
                              textColor=colors.HexColor('#6B7280'), spaceAfter=18)
    idx_title_sty = ParagraphStyle('B_IT', fontName='Helvetica-Bold', fontSize=16,
                                   textColor=colors.HexColor('#1A2840'), spaceAfter=10)
    row_sty = ParagraphStyle('B_R', fontName='Helvetica', fontSize=10,
                             textColor=colors.HexColor('#1F2937'), leading=14)
    link_sty = ParagraphStyle('B_L', fontName='Helvetica-Bold', fontSize=10,
                              textColor=colors.HexColor('#1A56DB'), leading=14)

    story = []
    if options.include_cover:
        story.append(Paragraph(options.cover_title, title_sty))
        story.append(Paragraph(
            f'Compiled by {options.cover_author} &nbsp;&bull;&nbsp; '
            f'{len(entries)} annexure{"s" if len(entries) != 1 else ""}',
            meta_sty,
        ))
        story.append(Spacer(1, 8))

    story.append(Paragraph('Index of Annexures', idx_title_sty))

    rows = [[
        Paragraph('<b>Annexure</b>', row_sty),
        Paragraph('<b>Description</b>', row_sty),
        Paragraph('<b>Pages</b>', row_sty),
    ]]
    for entry in entries:
        rows.append([
            Paragraph(f'<b>{entry.label}</b>', link_sty),
            Paragraph(_escape(entry.title), row_sty),
            Paragraph(str(entry.page_count or '—'), row_sty),
        ])

    tbl = Table(rows, colWidths=[2.4 * cm, None, 1.8 * cm])
    tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('LINEBELOW', (0, 0), (-1, 0), 0.6, colors.HexColor('#D1D5DB')),
        ('LINEBELOW', (0, 1), (-1, -1), 0.25, colors.HexColor('#E5E7EB')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F9FAFB')),
    ]))
    story.append(tbl)
    doc.build(story)
    return out


def annexure_anchor_id_for_label(label: str) -> str:
    safe = (label or '').replace(' ', '_')
    return f'mw_bundle_annex_{safe}'


def _escape(text: str) -> str:
    import html
    return html.escape(text or '', quote=False)


# ── Anchor injection ──────────────────────────────────────────────────────────

def _embed_anchor(writer: 'PdfWriter', page_index: int, anchor_name: str):
    """Add a named destination so `<link href="#name">` from the cover/index
    page resolves to `page_index` (0-based) in the merged document."""
    try:
        page_ref = writer.pages[page_index].indirect_reference
        dest = ArrayObject([
            page_ref,
            NameObject('/XYZ'),
            NumberObject(0),
            NumberObject(842),  # top-left of A4-ish; renderers clamp anyway
            NumberObject(0),
        ])
        writer.add_named_destination_array(anchor_name, dest)
    except Exception:
        _LOGGER.exception('anchor-embed-failed name=%s page=%s', anchor_name, page_index)


# ── Public entry point ───────────────────────────────────────────────────────

def auto_label_entries(entries: list[BundleEntry]) -> None:
    """Fill in label for any entry where the user didn't override one.
    Assigns A, B, … in current order. Preserves user-set labels."""
    used = {e.label.strip().upper() for e in entries if e.label.strip()}
    next_idx = 0
    for entry in entries:
        if entry.label.strip():
            continue
        while True:
            candidate = get_annexure_label(next_idx)
            next_idx += 1
            if candidate not in used:
                entry.label = candidate
                used.add(candidate)
                break


def build_bundle(entries: list[BundleEntry],
                 options: BundleOptions,
                 progress_cb: Optional[Callable[[int, int, str], None]] = None) -> BundleOptions:
    """Build the bundle PDF. Returns the (mutated) options with stats filled.

    Raises BundleBuildError on fatal problems; per-entry errors are recorded on
    the entry itself so the UI can show a per-row report after a partial build.
    """
    if not HAS_PYPDF:
        raise BundleBuildError(
            'pypdf is not installed.\n\nRun:  pip install pypdf'
        )
    if not HAS_REPORTLAB:
        raise BundleBuildError(
            'reportlab is not installed.\n\nRun:  pip install reportlab'
        )
    if not entries:
        raise BundleBuildError('No annexures to bundle.')

    size_map = {'A4': A4, 'Letter': letter, 'Legal': legal}
    page_size = size_map.get(options.page_size, A4)

    workdir = Path(tempfile.mkdtemp(prefix='mailweave_bundle_'))
    try:
        # Pass 1 — convert/probe sources to learn page counts (needed for the
        # index column AND for per-annexure "Page X of N" stamps).
        converted: list[tuple[BundleEntry, Path, int]] = []
        for idx, entry in enumerate(entries):
            if progress_cb:
                progress_cb(idx, len(entries) * 2, f'Reading {Path(entry.source_path).name}')
            try:
                pdf_path = _ensure_pdf(entry.source_path, workdir, page_size)
                reader = PdfReader(str(pdf_path))
                entry.page_count = len(reader.pages)
                converted.append((entry, pdf_path, entry.page_count))
            except BundleBuildError as exc:
                entry.error = str(exc)
                _LOGGER.warning('bundle-entry-failed label=%s error=%s', entry.label, exc)
                converted.append((entry, None, 0))

        # Build the cover + index now that we know page counts.
        cover_index_pdf = _build_cover_and_index_pdf(options, entries, page_size, workdir)

        writer = PdfWriter()

        # Append cover/index pages first.
        for page in PdfReader(str(cover_index_pdf)).pages:
            writer.add_page(page)
        cover_pages = len(writer.pages)

        # Append each annexure, stamped.
        bundle_page = cover_pages + 1
        for idx, (entry, pdf_path, _count) in enumerate(converted):
            if progress_cb:
                progress_cb(len(entries) + idx, len(entries) * 2,
                            f'Stamping {entry.label}: {Path(entry.source_path).name}')
            if pdf_path is None or entry.page_count == 0:
                continue
            entry.start_page = bundle_page
            try:
                reader = PdfReader(str(pdf_path))
                added = _stamp_pdf_pages(reader, writer, entry, options, bundle_page, page_size)
                bundle_page += added
            except Exception as exc:
                entry.error = f'Failed to merge: {exc}'
                _LOGGER.exception('bundle-merge-failed label=%s', entry.label)

        options.total_pages = bundle_page - 1
        options.entries = entries

        # Set basic metadata.
        try:
            writer.add_metadata({
                '/Title': options.cover_title,
                '/Author': options.cover_author,
                '/Producer': 'MailWeave',
            })
        except Exception:
            pass

        out = Path(options.output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, 'wb') as fh:
            writer.write(fh)

        if not out.is_file() or out.stat().st_size == 0:
            raise BundleBuildError('Bundle write produced an empty file.')

        return options
    finally:
        # Best-effort cleanup; intermediate PDFs aren't useful after build.
        try:
            import shutil
            shutil.rmtree(workdir, ignore_errors=True)
        except Exception:
            pass


# Re-export so callers can build anchor IDs without importing pypdf side.
__all__ = [
    'BundleEntry', 'BundleOptions', 'BundleBuildError',
    'build_bundle', 'auto_label_entries', 'supported_source',
    'HAS_PYPDF', 'HAS_DOCX2PDF', 'HAS_PIL',
]


# Reference these so static-checkers don't flag the imports.
_ = (DictionaryObject, TextStringObject, Fit, get_annexure_label)
