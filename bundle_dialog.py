"""MailWeave — Indexed Bundle wizard.

Simplified flow: add documents (button or drag-and-drop), reorder them, rename
titles inline, then click Build. Labels A, B, C… are assigned automatically
from display order and re-flow when rows are reordered.
"""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from threading import Thread
from tkinter import filedialog, messagebox, ttk

from annexures import get_annexure_label
from bundle_builder import (
    HAS_DOCX2PDF,
    HAS_PIL,
    HAS_PYPDF,
    BundleBuildError,
    BundleEntry,
    BundleOptions,
    build_bundle,
    supported_source,
)
from diagnostics import LOGGER

try:
    from tkinterdnd2 import DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


class BundleDialog:
    """Modal wizard for building an indexed document bundle."""

    def __init__(self, parent: tk.Tk, theme: dict,
                 default_author: str = '',
                 toast_cb=None):
        self.parent = parent
        self.T = theme
        self.default_author = default_author or 'MailWeave'
        self._toast = toast_cb or (lambda *_a, **_kw: None)
        self.entries: list[BundleEntry] = []

        self._win = tk.Toplevel(parent)
        self._win.title('Build Indexed Bundle — MailWeave')
        self._win.transient(parent)
        self._win.grab_set()
        self._win.configure(bg=theme['bg'])
        self._win.geometry('860x600')
        self._win.minsize(700, 500)

        self._busy = False
        self._build_ui()
        self._register_drop_target(self.tree)
        self._register_drop_target(self._win)
        self._refresh_table()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        T = self.T

        # Header
        head = tk.Frame(self._win, bg=T['surface'])
        head.pack(fill=tk.X)
        tk.Label(head, text='Build Indexed Bundle', bg=T['surface'], fg=T['fg'],
                 font=('Segoe UI Semibold', 14)).pack(anchor='w', padx=20, pady=(14, 0))

        caveats = []
        if not HAS_PYPDF:
            caveats.append('pypdf is not installed — bundling will not work. Run: pip install pypdf')
        if not HAS_DOCX2PDF:
            caveats.append('docx2pdf is not installed — Word documents cannot be converted.')
        if not HAS_PIL:
            caveats.append('Pillow is not installed — image sources cannot be embedded.')
        sub = (
            'Add documents below. The bundle will contain a cover page, a clickable '
            'index, and your documents — labelled A, B, C… in the order shown.'
        )
        if caveats:
            sub = sub + '\n\n• ' + '\n• '.join(caveats)
        tk.Label(head, text=sub, bg=T['surface'], fg=T['fgsub'],
                 font=('Segoe UI', 9), wraplength=820, justify='left').pack(
                 anchor='w', padx=20, pady=(2, 14))

        # Toolbar
        bar = tk.Frame(self._win, bg=T['bg'])
        bar.pack(fill=tk.X, padx=16, pady=(12, 6))
        ttk.Button(bar, text='Add documents…', command=self._add_files,
                   style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(bar, text='↑ Move up', command=lambda: self._move(-1),
                   style='Toolbar.TButton').pack(side=tk.LEFT, padx=(12, 4))
        ttk.Button(bar, text='↓ Move down', command=lambda: self._move(+1),
                   style='Toolbar.TButton').pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text='Rename…', command=self._rename_selected,
                   style='Toolbar.TButton').pack(side=tk.LEFT, padx=(12, 4))
        ttk.Button(bar, text='Remove', command=self._remove_selected,
                   style='Toolbar.TButton').pack(side=tk.LEFT, padx=4)

        hint = 'Drop files anywhere on this window to add them.' if HAS_DND else ''
        if hint:
            tk.Label(bar, text=hint, bg=T['bg'], fg=T['fgdim'],
                     font=('Segoe UI', 9)).pack(side=tk.RIGHT)

        # Table
        table_wrap = tk.Frame(self._win, bg=T['surface2'])
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        self.tree = ttk.Treeview(
            table_wrap,
            columns=('label', 'title', 'source', 'status'),
            show='headings',
            selectmode='extended',
        )
        self.tree.heading('label', text='Annexure')
        self.tree.heading('title', text='Title (double-click to edit label / title)')
        self.tree.heading('source', text='Source file')
        self.tree.heading('status', text='Status')
        self.tree.column('label', width=80, anchor='center', stretch=False)
        self.tree.column('title', width=280)
        self.tree.column('source', width=300)
        self.tree.column('status', width=120, anchor='w', stretch=False)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(table_wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind('<Double-1>', self._on_double_click)
        self.tree.bind('<Delete>', lambda _e: self._remove_selected())

        # Options
        opt = tk.LabelFrame(self._win, text='  Bundle options  ',
                            bg=T['bg'], fg=T['fg'],
                            font=('Segoe UI', 9, 'bold'), relief='flat', borderwidth=1,
                            highlightbackground=T['border'], highlightthickness=1)
        opt.pack(fill=tk.X, padx=16, pady=(8, 4))

        inner = tk.Frame(opt, bg=T['bg'])
        inner.pack(fill=tk.X, padx=10, pady=8)

        self.v_title = tk.StringVar(value='Indexed Bundle')
        self.v_author = tk.StringVar(value=self.default_author)
        self.v_page = tk.StringVar(value='A4')
        self.v_cover = tk.BooleanVar(value=True)
        self.v_bundle_pagination = tk.BooleanVar(value=True)
        self.v_annex_pagination = tk.BooleanVar(value=True)

        tk.Label(inner, text='Title:', bg=T['bg'], fg=T['fg'], font=('Segoe UI', 9)).grid(row=0, column=0, sticky='e', padx=4, pady=3)
        tk.Entry(inner, textvariable=self.v_title, bg=T['surface2'], fg=T['fg'],
                 insertbackground=T['fg'], relief='flat', font=('Segoe UI', 9), width=26).grid(row=0, column=1, sticky='we', padx=4)
        tk.Label(inner, text='Author:', bg=T['bg'], fg=T['fg'], font=('Segoe UI', 9)).grid(row=0, column=2, sticky='e', padx=4)
        tk.Entry(inner, textvariable=self.v_author, bg=T['surface2'], fg=T['fg'],
                 insertbackground=T['fg'], relief='flat', font=('Segoe UI', 9), width=22).grid(row=0, column=3, sticky='we', padx=4)
        tk.Label(inner, text='Page size:', bg=T['bg'], fg=T['fg'], font=('Segoe UI', 9)).grid(row=0, column=4, sticky='e', padx=4)
        ttk.Combobox(inner, textvariable=self.v_page, values=['A4', 'Letter', 'Legal'],
                     state='readonly', width=10).grid(row=0, column=5, sticky='w', padx=4)

        cb_kwargs = dict(bg=T['bg'], fg=T['fg'], selectcolor='#FFFFFF',
                         activebackground=T['bg'], activeforeground=T['fg'],
                         font=('Segoe UI', 9))
        tk.Checkbutton(inner, text='Include cover page', variable=self.v_cover,
                       **cb_kwargs).grid(row=1, column=0, columnspan=2, sticky='w', padx=2, pady=(6, 0))
        tk.Checkbutton(inner, text='Stamp bundle page numbers', variable=self.v_bundle_pagination,
                       **cb_kwargs).grid(row=1, column=2, columnspan=2, sticky='w', padx=2, pady=(6, 0))
        tk.Checkbutton(inner, text='Stamp per-annexure "Page X of N"', variable=self.v_annex_pagination,
                       **cb_kwargs).grid(row=1, column=4, columnspan=2, sticky='w', padx=2, pady=(6, 0))
        inner.columnconfigure(1, weight=1)
        inner.columnconfigure(3, weight=1)

        # Progress + actions
        bottom = tk.Frame(self._win, bg=T['bg'])
        bottom.pack(fill=tk.X, padx=16, pady=(4, 14))

        self.status_lbl = tk.Label(bottom, text='Ready.', bg=T['bg'], fg=T['fgsub'],
                                    font=('Segoe UI', 9), anchor='w')
        self.status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress = ttk.Progressbar(bottom, orient=tk.HORIZONTAL, length=240, mode='determinate')
        self.progress.pack(side=tk.LEFT, padx=10)

        ttk.Button(bottom, text='Close', command=self._on_close).pack(side=tk.RIGHT)
        self.build_btn = ttk.Button(bottom, text='Build PDF…', command=self._on_build,
                                    style='Accent.TButton')
        self.build_btn.pack(side=tk.RIGHT, padx=8)

        self._win.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── Labelling ─────────────────────────────────────────────────────────────

    def _relabel(self):
        """Assign A, B, C… in display order, skipping labels the user has pinned.

        Manual labels stay where the user put them; auto rows fill the next
        unused slot in alphabetical order.
        """
        used = {e.label.strip().upper() for e in self.entries
                if e.manual_label and e.label.strip()}
        next_idx = 0
        for entry in self.entries:
            if entry.manual_label and entry.label.strip():
                continue
            while True:
                candidate = get_annexure_label(next_idx)
                next_idx += 1
                if candidate not in used:
                    entry.label = candidate
                    used.add(candidate)
                    break

    def _refresh_table(self):
        self._relabel()
        self.tree.delete(*self.tree.get_children())
        for idx, entry in enumerate(self.entries):
            status = 'OK' if entry.source_path and os.path.exists(entry.source_path) else 'Missing'
            if entry.error:
                status = entry.error[:40]
            src = os.path.basename(entry.source_path) if entry.source_path else '(none)'
            self.tree.insert('', 'end', iid=str(idx),
                             values=(entry.label, entry.title or '(untitled)', src, status))

    # ── Row operations ────────────────────────────────────────────────────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            parent=self._win,
            title='Add documents to bundle',
            filetypes=[
                ('Supported', '*.pdf *.docx *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.gif'),
                ('PDF', '*.pdf'),
                ('Word', '*.docx'),
                ('Images', '*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.gif'),
            ],
        )
        if paths:
            self._absorb_paths(paths)

    def _absorb_paths(self, paths):
        added = skipped = 0
        for path in paths:
            if not path:
                continue
            if not supported_source(path):
                skipped += 1
                continue
            self.entries.append(BundleEntry(
                label='',
                title=Path(path).stem,
                source_path=path,
            ))
            added += 1
        if added:
            self._refresh_table()
            self._toast(f'Added {added} document(s).', kind='success')
        if skipped and not added:
            messagebox.showinfo(
                'Unsupported files',
                f'{skipped} file(s) were skipped (unsupported type). '
                'Supported: PDF, DOCX, PNG, JPG, TIFF, BMP, GIF.',
                parent=self._win,
            )

    def _remove_selected(self):
        targets = sorted({int(i) for i in self.tree.selection() if i.isdigit()}, reverse=True)
        for idx in targets:
            if 0 <= idx < len(self.entries):
                self.entries.pop(idx)
        self._refresh_table()

    def _move(self, delta: int):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        new_idx = idx + delta
        if not (0 <= new_idx < len(self.entries)):
            return
        self.entries[idx], self.entries[new_idx] = self.entries[new_idx], self.entries[idx]
        self._refresh_table()
        self.tree.selection_set(str(new_idx))
        self.tree.focus(str(new_idx))
        self.tree.see(str(new_idx))

    def _on_double_click(self, event):
        row = self.tree.identify_row(event.y)
        if not row:
            return
        try:
            idx = int(row)
        except ValueError:
            return
        self._rename_row(idx)

    def _rename_selected(self):
        sel = self.tree.selection()
        if sel:
            self._rename_row(int(sel[0]))

    def _rename_row(self, idx: int):
        if not (0 <= idx < len(self.entries)):
            return
        T = self.T
        entry = self.entries[idx]

        top = tk.Toplevel(self._win)
        top.title('Edit annexure')
        top.transient(self._win)
        top.grab_set()
        top.configure(bg=T['bg'])
        top.resizable(False, False)

        f = tk.Frame(top, bg=T['bg'], padx=18, pady=14)
        f.pack(fill=tk.BOTH, expand=True)

        tk.Label(f, text='Label (leave blank for automatic A, B, C…):',
                 bg=T['bg'], fg=T['fg'], font=('Segoe UI', 9)).pack(anchor='w')
        v_label = tk.StringVar(value=entry.label if entry.manual_label else '')
        lbl_ent = tk.Entry(f, textvariable=v_label, bg=T['surface2'], fg=T['fg'],
                           insertbackground=T['fg'], relief='flat',
                           font=('Segoe UI', 10), width=10)
        lbl_ent.pack(anchor='w', pady=(4, 10))

        tk.Label(f, text='Title:',
                 bg=T['bg'], fg=T['fg'], font=('Segoe UI', 9)).pack(anchor='w')
        v_title = tk.StringVar(value=entry.title)
        ent = tk.Entry(f, textvariable=v_title, bg=T['surface2'], fg=T['fg'],
                       insertbackground=T['fg'], relief='flat',
                       font=('Segoe UI', 10), width=50)
        ent.pack(fill=tk.X, pady=(4, 10))
        ent.icursor(tk.END)
        ent.focus_set()
        ent.selection_range(0, tk.END)

        def _save(_e=None):
            label_text = v_label.get().strip().upper()
            entry.title = v_title.get().strip() or Path(entry.source_path).stem
            if label_text:
                entry.label = label_text
                entry.manual_label = True
            else:
                entry.manual_label = False
            self._refresh_table()
            top.destroy()

        btns = tk.Frame(f, bg=T['bg'])
        btns.pack(fill=tk.X)
        ttk.Button(btns, text='Cancel', command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text='Save', command=_save,
                   style='Accent.TButton').pack(side=tk.RIGHT, padx=6)
        for widget in (lbl_ent, ent):
            widget.bind('<Return>', _save)
            widget.bind('<Escape>', lambda _e: top.destroy())

    # ── Drag and drop ─────────────────────────────────────────────────────────

    def _register_drop_target(self, widget):
        if not HAS_DND:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self._on_drop)
        except Exception:
            LOGGER.exception('bundle-dnd-register-failed')

    def _on_drop(self, event):
        try:
            paths = self._win.tk.splitlist(event.data)
        except Exception:
            return
        files = [p for p in paths if os.path.isfile(p)]
        if files:
            self._absorb_paths(files)

    # ── Build ─────────────────────────────────────────────────────────────────

    def _on_build(self):
        if self._busy:
            return
        if not self.entries:
            messagebox.showinfo('Nothing to build',
                                'Add at least one document first.',
                                parent=self._win)
            return
        missing = [e for e in self.entries if not e.source_path or not os.path.exists(e.source_path)]
        if missing:
            messagebox.showerror(
                'Missing sources',
                f'{len(missing)} row(s) have no valid source file.',
                parent=self._win,
            )
            return
        if not HAS_PYPDF:
            messagebox.showerror('pypdf missing',
                                 'pypdf is not installed.\n\nRun: pip install pypdf',
                                 parent=self._win)
            return

        self._relabel()

        out_path = filedialog.asksaveasfilename(
            parent=self._win,
            title='Save bundle PDF',
            defaultextension='.pdf',
            filetypes=[('PDF', '*.pdf')],
            initialfile='bundle.pdf',
        )
        if not out_path:
            return

        options = BundleOptions(
            output_path=out_path,
            page_size=self.v_page.get(),
            include_cover=self.v_cover.get(),
            cover_title=self.v_title.get().strip() or 'Indexed Bundle',
            cover_author=self.v_author.get().strip() or self.default_author,
            include_bundle_pagination=self.v_bundle_pagination.get(),
            stamp_per_annexure_pagination=self.v_annex_pagination.get(),
        )

        self._busy = True
        self.build_btn.configure(state=tk.DISABLED)
        self.progress['mode'] = 'determinate'
        self.progress['value'] = 0
        self.status_lbl.config(text='Starting build…')

        def progress_cb(done: int, total: int, label: str):
            self._win.after(0, lambda: self._update_progress(done, total, label))

        def _worker():
            try:
                result = build_bundle(self.entries, options, progress_cb=progress_cb)
                self._win.after(0, lambda: self._build_finished(result))
            except BundleBuildError as exc:
                LOGGER.exception('bundle-build-failed')
                self._win.after(0, lambda: self._build_failed(str(exc)))
            except Exception as exc:
                LOGGER.exception('bundle-build-failed-unexpected')
                self._win.after(0, lambda: self._build_failed(f'Unexpected error: {exc}'))

        Thread(target=_worker, daemon=True).start()

    def _update_progress(self, done: int, total: int, label: str):
        total = max(total, 1)
        self.progress['value'] = int(done / total * 100)
        self.status_lbl.config(text=label or 'Building…')

    def _build_finished(self, options: BundleOptions):
        self._busy = False
        self.build_btn.configure(state=tk.NORMAL)
        self.progress['value'] = 100
        self._refresh_table()
        failed = [e for e in options.entries if e.error]
        msg = (
            f'Bundle saved to:\n{options.output_path}\n\n'
            f'Total pages: {options.total_pages}\n'
            f'Annexures: {len(options.entries)} '
            f'({len(failed)} with errors)'
        )
        self.status_lbl.config(text=f'Done — {options.total_pages} pages.')
        if messagebox.askyesno('Bundle complete',
                               msg + '\n\nOpen the file now?',
                               parent=self._win):
            try:
                if hasattr(os, 'startfile'):
                    os.startfile(options.output_path)
            except Exception:
                LOGGER.exception('bundle-open-failed')

    def _build_failed(self, message: str):
        self._busy = False
        self.build_btn.configure(state=tk.NORMAL)
        self.progress['value'] = 0
        self.status_lbl.config(text='Build failed.')
        messagebox.showerror('Bundle failed', message, parent=self._win)

    def _on_close(self):
        if self._busy:
            if not messagebox.askyesno(
                'Build in progress',
                'A build is running. Close anyway? The output PDF may be incomplete.',
                parent=self._win,
            ):
                return
        self._win.destroy()
