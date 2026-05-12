"""MailWeave — Options dialog."""

from __future__ import annotations

import tkinter as tk
from dataclasses import replace
from tkinter import ttk, filedialog
from typing import Callable

from settings import AppSettings, reset_settings


class OptionsDialog:
    """Modal dialog for editing all AppSettings, organised in 3 tabs."""

    W, H = 530, 480

    def __init__(
        self,
        parent: tk.Tk,
        settings: AppSettings,
        on_apply: Callable[[AppSettings], None],
        T: dict,          # theme dict from themes.get_theme()
    ):
        self.parent   = parent
        self.settings = settings
        self.on_apply = on_apply
        self.T        = T

        self._win = tk.Toplevel(parent)
        self._win.title('Options — MailWeave')
        self._win.transient(parent)
        self._win.grab_set()
        self._win.resizable(False, False)
        self._win.configure(bg=T['bg'])
        self._center()

        self._init_vars()
        self._build_styles()
        self._build()
        self._win.protocol('WM_DELETE_WINDOW', self._cancel)
        parent.wait_window(self._win)

    # ── Setup ──────────────────────────────────────────────────────────────────

    def _center(self):
        self._win.update_idletasks()
        pw = self.parent.winfo_x() + self.parent.winfo_width() // 2
        ph = self.parent.winfo_y() + self.parent.winfo_height() // 2
        self._win.geometry(
            f'{self.W}x{self.H}+{pw - self.W // 2}+{ph - self.H // 2}')

    def _init_vars(self):
        s = self.settings
        self.v_strip        = tk.BooleanVar(value=s.strip_quotes)
        self.v_sort         = tk.StringVar(
            value='oldest' if s.sort_oldest_first else 'newest')
        self.v_date_fmt     = tk.StringVar(value=s.date_format)
        self.v_page_size    = tk.StringVar(value=s.pdf_page_size)
        self.v_cover        = tk.BooleanVar(value=s.include_cover_page)
        self.v_index        = tk.BooleanVar(value=s.include_email_index)
        self.v_safe_mode    = tk.BooleanVar(value=s.safe_mode_import)
        self.v_autosave     = tk.BooleanVar(value=s.autosave_session)
        self.v_confirm_clear = tk.BooleanVar(value=s.confirm_before_clear)
        self.v_confirm_export = tk.BooleanVar(value=s.confirm_export_overwrite)
        self.v_author       = tk.StringVar(value=s.document_author)
        self.v_export_dir   = tk.StringVar(value=s.default_export_dir)
        self.v_remember_dir = tk.BooleanVar(value=s.remember_export_dir)
        self.v_theme        = tk.StringVar(value=s.theme)
        self.v_font_size    = tk.StringVar(value=s.font_size)
        self.v_show_preview = tk.BooleanVar(value=s.show_preview)

    def _build_styles(self):
        T = self.T
        s = ttk.Style()
        s.configure('Opts.TNotebook',
                    background=T['bg'], borderwidth=0, tabmargins=[0, 0, 0, 0])
        s.configure('Opts.TNotebook.Tab',
                    background=T['surface2'], foreground=T['fgsub'],
                    font=('Segoe UI', 10), padding=(18, 7),
                    borderwidth=0)
        s.map('Opts.TNotebook.Tab',
              background=[('selected', T['surface'])],
              foreground=[('selected', T['fg'])])
        s.configure('Opts.TCombobox',
                    fieldbackground=T['surface2'],
                    foreground=T['fg'], selectbackground=T['accent'],
                    selectforeground='white', background=T['surface2'],
                    arrowcolor=T['fgsub'])

    # ── Main layout ────────────────────────────────────────────────────────────

    def _build(self):
        T = self.T

        nb = ttk.Notebook(self._win, style='Opts.TNotebook')
        nb.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self._tab_general(nb, T)
        self._tab_export(nb, T)
        self._tab_appearance(nb, T)

        # Separator
        tk.Frame(self._win, bg=T['border'], height=1).pack(fill=tk.X)

        # Button bar
        btn_row = tk.Frame(self._win, bg=T['surface'])
        btn_row.pack(fill=tk.X, padx=14, pady=10)

        self._btn(btn_row, 'Restore Defaults',
                  self._restore_defaults).pack(side=tk.LEFT)
        self._btn(btn_row, 'Cancel',
                  self._cancel).pack(side=tk.RIGHT, padx=(4, 0))
        self._btn(btn_row, 'Apply',
                  self._apply).pack(side=tk.RIGHT, padx=(4, 0))
        self._btn(btn_row, '  OK  ',
                  self._ok, accent=True).pack(side=tk.RIGHT, padx=(4, 0))

    # ── Tabs ───────────────────────────────────────────────────────────────────

    def _tab_general(self, nb, T):
        f = self._tab_frame(nb, 'General', T)

        s = self._section(f, 'Behaviour', T)
        self._chk(s, 'Strip quoted replies  (show only new content per email)',
                  self.v_strip, T)
        self._chk(s, 'Use safe-mode imports  (copy files into a temporary workspace first)',
                  self.v_safe_mode, T)
        self._chk(s, 'Autosave the current session for crash recovery',
                  self.v_autosave, T)
        self._chk(s, 'Ask before clearing all loaded emails',
                  self.v_confirm_clear, T)

        s2 = self._section(f, 'Sort Order', T)
        self._radio(s2, 'Oldest email first (chronological \u2191)',
                    self.v_sort, 'oldest', T)
        self._radio(s2, 'Newest email first (reverse chronological \u2193)',
                    self.v_sort, 'newest', T)

        s3 = self._section(f, 'Date Display Format', T)
        for val, label in [
            ('uk',  'DD/MM/YYYY HH:MM   e.g. 25/12/2024 14:30'),
            ('us',  'MM/DD/YYYY HH:MM   e.g. 12/25/2024 14:30'),
            ('iso', 'YYYY-MM-DD HH:MM   e.g. 2024-12-25 14:30'),
        ]:
            self._radio(s3, label, self.v_date_fmt, val, T)

    def _tab_export(self, nb, T):
        f = self._tab_frame(nb, 'Export', T)

        s = self._section(f, 'Default Save Location', T)
        self._chk(s, 'Remember the last used export folder',
                  self.v_remember_dir, T)

        row = tk.Frame(s, bg=T['bg'])
        row.pack(fill=tk.X, pady=(8, 0))
        tk.Label(row, text='Default folder:',
                 bg=T['bg'], fg=T['fg'],
                 font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 6))
        tk.Entry(row, textvariable=self.v_export_dir,
                 bg=T['surface2'], fg=T['fg'],
                 insertbackground=T['fg'],
                 relief='flat', font=('Segoe UI', 10),
                 highlightthickness=1,
                 highlightbackground=T['border'],
                 highlightcolor=T['accent']
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._btn(row, 'Browse\u2026',
                  self._browse_export_dir, small=True).pack(
                      side=tk.LEFT, padx=(6, 0))

        s2 = self._section(f, 'PDF Settings', T)
        row2 = tk.Frame(s2, bg=T['bg'])
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text='Page size:',
                 bg=T['bg'], fg=T['fg'],
                 font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Combobox(row2, textvariable=self.v_page_size,
                     values=['A4', 'Letter', 'Legal'],
                     state='readonly', width=11,
                     style='Opts.TCombobox').pack(side=tk.LEFT)
        self._chk(s2, 'Include cover page', self.v_cover, T)
        self._chk(
            s2,
            'Include email index with descriptions and page numbers',
            self.v_index,
            T,
        )
        self._chk(s2, 'Ask before overwriting an existing export file',
                  self.v_confirm_export, T)

        row3 = tk.Frame(s2, bg=T['bg'])
        row3.pack(fill=tk.X, pady=(8, 0))
        tk.Label(row3, text='Author name:',
                 bg=T['bg'], fg=T['fg'],
                 font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 6))
        tk.Entry(row3, textvariable=self.v_author,
                 bg=T['surface2'], fg=T['fg'],
                 insertbackground=T['fg'],
                 relief='flat', font=('Segoe UI', 10),
                 highlightthickness=1,
                 highlightbackground=T['border'],
                 highlightcolor=T['accent']
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _tab_appearance(self, nb, T):
        f = self._tab_frame(nb, 'Appearance', T)

        s = self._section(f, 'Theme', T)
        tk.Label(s,
                 text='Restart MailWeave for theme changes to apply across all panels.',
                 bg=T['bg'], fg=T['fgsub'],
                 font=('Segoe UI', 8, 'italic')).pack(anchor='w', pady=(0, 6))
        row = tk.Frame(s, bg=T['bg'])
        row.pack(fill=tk.X)
        for val, lbl in [('dark', '\u25cf  Dark'), ('light', '\u25cb  Light')]:
            self._radio(row, lbl, self.v_theme, val, T,
                        pack_kwargs={'side': tk.LEFT, 'padx': (0, 28)})

        s2 = self._section(f, 'Preview Pane', T)
        self._chk(s2, 'Show preview pane', self.v_show_preview, T)

        row2 = tk.Frame(s2, bg=T['bg'])
        row2.pack(fill=tk.X, pady=(10, 0))
        tk.Label(row2, text='Preview text size:',
                 bg=T['bg'], fg=T['fg'],
                 font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Combobox(row2, textvariable=self.v_font_size,
                     values=['small', 'medium', 'large'],
                     state='readonly', width=10,
                     style='Opts.TCombobox').pack(side=tk.LEFT)

    # ── Widget factories ───────────────────────────────────────────────────────

    def _tab_frame(self, nb, title: str, T: dict) -> tk.Frame:
        f = tk.Frame(nb, bg=T['bg'])
        nb.add(f, text=f'  {title}  ')
        return f

    def _section(self, parent, title: str, T: dict) -> tk.Frame:
        outer = tk.LabelFrame(
            parent, text=f'  {title}  ',
            bg=T['bg'], fg=T['fgsub'],
            font=('Segoe UI', 9, 'bold'),
            relief='groove', borderwidth=1,
            highlightbackground=T['border'],
            labelanchor='nw',
        )
        outer.pack(fill=tk.X, padx=12, pady=(12, 0))
        inner = tk.Frame(outer, bg=T['bg'])
        inner.pack(fill=tk.X, padx=10, pady=(6, 10))
        return inner

    def _btn(self, parent, text: str, cmd,
             accent: bool = False, small: bool = False) -> tk.Button:
        T = self.T
        bg  = T['accent']   if accent else T['surface3']
        fg  = '#FFFFFF'     if accent else T['fg']
        abg = T['accent2']  if accent else T['surface2']
        return tk.Button(
            parent, text=text, command=cmd,
            bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
            relief='flat', borderwidth=0, cursor='hand2',
            font=('Segoe UI', 9 if small else 10,
                  'bold' if accent else 'normal'),
            padx=8 if small else 14,
            pady=3 if small else 6,
        )

    def _chk(self, parent, text: str, var, T: dict) -> tk.Checkbutton:
        c = tk.Checkbutton(
            parent, text=text, variable=var,
            bg=T['bg'], fg=T['fg'],
            selectcolor=T['surface3'],
            activebackground=T['bg'], activeforeground=T['fg'],
            font=('Segoe UI', 10), anchor='w',
            relief='flat', cursor='hand2',
        )
        c.pack(fill=tk.X, pady=2)
        return c

    def _radio(self, parent, text: str, var, value,
               T: dict, pack_kwargs=None) -> tk.Radiobutton:
        r = tk.Radiobutton(
            parent, text=text, variable=var, value=value,
            bg=T['bg'], fg=T['fg'],
            selectcolor=T['surface3'],
            activebackground=T['bg'], activeforeground=T['fg'],
            font=('Segoe UI', 10), anchor='w',
            relief='flat', cursor='hand2',
        )
        kw = pack_kwargs or {'fill': tk.X, 'pady': 2}
        r.pack(**kw)
        return r

    # ── Actions ────────────────────────────────────────────────────────────────

    def _collect(self) -> AppSettings:
        # Use `replace` so we keep fields the dialog doesn't expose (notably
        # disclaimer_accepted). Adding new dialog-managed fields here is safe;
        # forgetting to copy back ones we don't manage is not.
        return replace(
            self.settings,
            strip_quotes        = self.v_strip.get(),
            sort_oldest_first   = self.v_sort.get() == 'oldest',
            date_format         = self.v_date_fmt.get(),
            pdf_page_size       = self.v_page_size.get(),
            include_cover_page  = self.v_cover.get(),
            include_email_index = self.v_index.get(),
            safe_mode_import    = self.v_safe_mode.get(),
            autosave_session    = self.v_autosave.get(),
            confirm_before_clear = self.v_confirm_clear.get(),
            confirm_export_overwrite = self.v_confirm_export.get(),
            document_author     = self.v_author.get().strip(),
            default_export_dir  = self.v_export_dir.get().strip(),
            remember_export_dir = self.v_remember_dir.get(),
            theme               = self.v_theme.get(),
            font_size           = self.v_font_size.get(),
            show_preview        = self.v_show_preview.get(),
        )

    def _apply(self):
        self.on_apply(self._collect())

    def _ok(self):
        self.on_apply(self._collect())
        self._win.destroy()

    def _cancel(self):
        self._win.destroy()

    def _restore_defaults(self):
        d = reset_settings()
        self.v_strip.set(d.strip_quotes)
        self.v_sort.set('oldest' if d.sort_oldest_first else 'newest')
        self.v_date_fmt.set(d.date_format)
        self.v_page_size.set(d.pdf_page_size)
        self.v_cover.set(d.include_cover_page)
        self.v_index.set(d.include_email_index)
        self.v_safe_mode.set(d.safe_mode_import)
        self.v_autosave.set(d.autosave_session)
        self.v_confirm_clear.set(d.confirm_before_clear)
        self.v_confirm_export.set(d.confirm_export_overwrite)
        self.v_author.set(d.document_author)
        self.v_export_dir.set(d.default_export_dir)
        self.v_remember_dir.set(d.remember_export_dir)
        self.v_theme.set(d.theme)
        self.v_font_size.set(d.font_size)
        self.v_show_preview.set(d.show_preview)

    def _browse_export_dir(self):
        folder = filedialog.askdirectory(
            parent=self._win, title='Select default export folder')
        if folder:
            self.v_export_dir.set(folder)
