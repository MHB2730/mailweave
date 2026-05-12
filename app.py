"""MailWeave — main application window."""

from __future__ import annotations

import os
import re
import shutil
import sys
import tkinter as tk
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from typing import Optional

from email_parser import EmailData, email_timestamp, parse_email_file
from exporters import HAS_DOCX, HAS_PDF, build_docx, build_pdf, verify_export_file
from brand_assets import load_logo_photo, set_window_icon
from diagnostics import (
    LOGGER,
    build_diagnostic_report,
    clear_crash_reports,
    pending_crash_reports,
    summarize_checks,
    write_crash_report,
)
from updater import (
    UpdateInfo,
    check_for_update,
    download_installer,
    launch_installer,
)
from version import GITHUB_REPO
from outlook_import import (
    HAS_OUTLOOK_IMPORT,
    OutlookImportError,
    get_outlook_selection_count,
    iter_selected_outlook_emails,
)
from annexures import auto_assign_annexures
from session_store import clear_session, clear_session_lock, has_recovery_session, load_session, mark_session_open, save_session
from settings import AppSettings, save_settings
from themes import get_theme

try:
    from tkinterdnd2 import DND_FILES

    HAS_DND = True
except ImportError:
    HAS_DND = False


_SHORT_FMTS = {
    'uk': '%d/%m/%Y  %H:%M',
    'us': '%m/%d/%Y  %H:%M',
    'iso': '%Y-%m-%d  %H:%M',
}


def _fmt_short(
    dt: Optional[datetime], fallback: str, date_format: str = 'uk'
) -> str:
    fmt = _SHORT_FMTS.get(date_format, _SHORT_FMTS['uk'])
    if dt is None:
        return fallback
    try:
        return dt.strftime(fmt)
    except Exception:
        return str(dt)[:16]


class _ExportCancelled(Exception):
    """Raised inside an export worker when the user clicks Cancel."""


def _trunc(text: str, length: int) -> str:
    text = text or ''
    return (text[: length - 1] + '\u2026') if len(text) > length else text


class MailWeaveApp:
    """Main MailWeave window. Accepts an existing tk root + AppSettings."""

    APP_TITLE = 'MailWeave — Email Bundler'
    from version import __version__ as _APP_VERSION
    VERSION = f'v{_APP_VERSION}'

    def __init__(self, root: tk.Tk, settings: AppSettings, startup_checks=None):
        self.root = root
        self.settings = settings
        self.emails: list[EmailData] = []
        self.startup_checks = startup_checks or []
        self._tw: list[tuple[tk.BaseWidget, dict[str, str]]] = []
        self._status_text = 'Ready'
        self._dnd_ready = HAS_DND
        self._loading = False
        self._load_queue: Queue | None = None
        self._load_job: dict | None = None
        self._tree_render_after: str | None = None
        self._load_cancel = Event()
        self._export_busy = False
        self._last_import_report: dict | None = None
        self._search_after: str | None = None
        self._display_emails: list[EmailData] = []
        self._current_view = 'emails'  # 'emails' | 'annexures'
        self._recent: list[str] = self._load_recent()
        self._redact_var: tk.BooleanVar | None = None
        self._preview_email: EmailData | None = None
        self._export_total = 0

        self._setup_root()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        mark_session_open()
        self._setup_styles()
        self._build_menu()

        # Layout containers
        self.main_wrap = tk.Frame(self.root, bg=self.T['bg'])
        self.main_wrap.pack(fill=tk.BOTH, expand=True)

        self.sidebar = tk.Frame(self.main_wrap, bg=self.T['sidebar'], width=240)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        self.content_wrap = tk.Frame(self.main_wrap, bg=self.T['bg'])
        self.content_wrap.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_sidebar()
        self._build_header()
        self._build_content()
        self._build_statusbar()
        
        self._bind_shortcuts()
        self._update_preview_visibility()
        self._clear_preview()
        self._update_stats()
        self._maybe_restore_session()
        self._show_startup_checks()

        if not self.settings.disclaimer_accepted:
            self.root.after(500, self._show_disclaimer)

        self._apply_saved_column_widths()
        self._update_sort_indicators()
        # Surface any prior-session crash report once, after the disclaimer.
        self.root.after(1500, self._review_pending_crashes)
        if self.settings.check_updates and GITHUB_REPO:
            Thread(target=self._background_update_check, daemon=True).start()

    @property
    def T(self) -> dict[str, str]:
        return get_theme(self.settings.theme)

    def _reg(self, widget: tk.BaseWidget, **theme_keys) -> tk.BaseWidget:
        self._tw.append((widget, theme_keys))
        return widget

    def _apply_theme(self):
        theme = self.T
        self.root.configure(bg=theme['bg'])
        self._setup_styles()
        for widget, mapping in self._tw:
            try:
                widget.configure(**{key: theme[value] for key, value in mapping.items()})
            except Exception:
                pass
        
        self.sidebar.configure(bg=theme['sidebar'])
        self.content_wrap.configure(bg=theme['bg'])
        self.header_bar.configure(bg=theme['sidebar']) # Using sidebar color for header too
        self._draw_brand_art()

        font_size = {'small': 9, 'medium': 10, 'large': 12}.get(
            self.settings.font_size, 10
        )
        self.preview.configure(
            bg=theme['surface'],
            fg=theme['fg'],
            insertbackground=theme['fg'],
            font=('Segoe UI', font_size),
        )
        self.preview.tag_configure(
            'hdr_lbl', font=('Segoe UI Semibold', 9, 'bold'), foreground=theme['accent']
        )
        self.preview.tag_configure('hdr_val', font=('Segoe UI', 9), foreground=theme['fg'])
        self.preview.tag_configure('sep', foreground=theme['fgdim'])
        self.preview.tag_configure('body', font=('Segoe UI', font_size), foreground=theme['fg'])
        self._update_stats()
        self._set_status(self._status_text)

    def _setup_root(self):
        theme = self.T
        self.root.title(self.APP_TITLE)
        # Restore last geometry if it parses; clamp to screen so a moved monitor
        # can't strand the window off-screen.
        geom = (self.settings.window_geometry or '').strip()
        if not (geom and self._apply_saved_geometry(geom)):
            self.root.geometry('1300x840')
        self.root.minsize(980, 660)
        self.root.configure(bg=theme['bg'])
        set_window_icon(self.root)
        self.root.report_callback_exception = self._report_callback_exception
        self._apply_dpi_scaling()

    def _apply_saved_geometry(self, geom: str) -> bool:
        import re as _re
        m = _re.fullmatch(r'(\d+)x(\d+)([+-]\d+)([+-]\d+)', geom)
        if not m:
            return False
        w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        if w < 600 or h < 400 or w > sw + 200 or h > sh + 200:
            return False
        # Pull off-screen positions back into the visible area.
        x = max(0, min(x, max(0, sw - 200)))
        y = max(0, min(y, max(0, sh - 200)))
        self.root.geometry(f'{w}x{h}+{x}+{y}')
        return True

    def _apply_dpi_scaling(self):
        """Match Tk's drawing scale to the system DPI so layouts don't render
        cramped on 150%/200% displays. Safe to call once at startup."""
        try:
            if sys.platform == 'win32':
                try:
                    ctypes_user32 = __import__('ctypes').windll.shcore
                    ctypes_user32.SetProcessDpiAwareness(1)
                except Exception:
                    pass
                dpi = self.root.winfo_fpixels('1i')  # pixels per inch
                if dpi and dpi > 0:
                    self.root.tk.call('tk', 'scaling', dpi / 72.0)
        except Exception:
            LOGGER.exception('dpi-scaling-failed')

    def _report_callback_exception(self, exc, val, tb):
        tb_text = ''.join(traceback.format_exception(exc, val, tb))
        LOGGER.error('tk-callback-exception: %s', tb_text)
        try:
            path = write_crash_report(
                val if isinstance(val, BaseException) else Exception(str(val)),
                tb_text,
                context={'email_count': len(self.emails),
                         'view': self._current_view,
                         'loading': self._loading},
            )
        except Exception:
            path = None
        messagebox.showerror(
            'Unexpected error',
            'MailWeave hit an unexpected UI error.\n\n'
            f'A crash report has been written to:\n{path}\n\n'
            'You can review it from Help > Crash Reports next time you start the app.',
        )

    def _setup_styles(self):
        theme = self.T
        style = ttk.Style(self.root)
        style.theme_use('clam')

        style.configure('TFrame', background=theme['bg'])
        style.configure('TLabel', background=theme['bg'], foreground=theme['fg'], font=('Segoe UI', 10))
        
        # New Progressbar styles
        style.layout('Status.Horizontal.TProgressbar', style.layout('Horizontal.TProgressbar'))
        style.configure(
            'Status.Horizontal.TProgressbar',
            troughcolor=theme['surface2'],
            background=theme['accent'],
            thickness=6,
            borderwidth=0,
        )
        style.layout('Hero.Horizontal.TProgressbar', style.layout('Horizontal.TProgressbar'))
        style.configure(
            'Hero.Horizontal.TProgressbar',
            troughcolor=theme['surface2'],
            background=theme['accent'],
            thickness=12,
            borderwidth=0,
        )
        style.configure(
            'Toolbar.TButton',
            background=theme['surface3'],
            foreground=theme['fg'],
            font=('Segoe UI Semibold', 9),
            padding=(12, 7),
            relief='flat',
            borderwidth=0,
        )
        style.map(
            'Toolbar.TButton',
            background=[('active', theme['accent'])],
            foreground=[('active', '#FFFFFF')],
        )
        style.configure(
            'TButton',
            background=theme['surface3'],
            foreground=theme['fg'],
            font=('Segoe UI Semibold', 10),
            padding=(13, 7),
            relief='flat',
            borderwidth=0,
        )
        style.map(
            'TButton',
            background=[('active', theme['accent'])],
            foreground=[('active', '#FFFFFF')],
        )
        style.configure(
            'Accent.TButton',
            background=theme['accent'],
            foreground='#FFFFFF',
            font=('Segoe UI Semibold', 10, 'bold'),
            padding=(16, 8),
            relief='flat',
            borderwidth=0,
        )
        style.map(
            'Accent.TButton',
            background=[('active', theme['accent2'])],
            foreground=[('active', '#FFFFFF')],
        )
        style.configure(
            'Treeview',
            background=theme['surface'],
            foreground=theme['fg'],
            fieldbackground=theme['surface'],
            rowheight=36,
            borderwidth=0,
            font=('Segoe UI', 9),
        )
        style.configure(
            'Treeview.Heading',
            background=theme['surface2'],
            foreground=theme['fgsub'],
            font=('Segoe UI Semibold', 9, 'bold'),
            relief='flat',
            padding=(8, 8),
        )
        style.map(
            'Treeview',
            background=[('selected', theme['sel_bg'])],
            foreground=[('selected', theme['sel_fg'])],
        )
        style.configure(
            'TScrollbar',
            background=theme['surface2'],
            troughcolor=theme['surface'],
            borderwidth=0,
            arrowsize=12,
            relief='flat',
        )
        style.configure('TPanedwindow', background=theme['bg'])
        style.configure('Sash', sashthickness=8, sashrelief='flat')

    def _build_menu(self):
        theme = self.T
        root = self.root

        def _menu(parent):
            return tk.Menu(
                parent,
                tearoff=0,
                bg=theme['menu_bg'],
                fg=theme['menu_fg'],
                activebackground=theme['accent'],
                activeforeground='#FFFFFF',
                relief='flat',
                borderwidth=1,
            )

        menubar = tk.Menu(
            root,
            bg=theme['menu_bg'],
            fg=theme['menu_fg'],
            activebackground=theme['accent'],
            activeforeground='#FFFFFF',
            relief='flat',
            borderwidth=0,
        )

        file_menu = _menu(menubar)
        file_menu.add_command(label='Open Files…', accelerator='Ctrl+O', command=self._browse_files)
        file_menu.add_command(label='Open Folder…', command=self._browse_folder)
        file_menu.add_command(
            label='Import Outlook Selection…',
            accelerator='Ctrl+Shift+O',
            command=self._import_outlook_selection,
            state=tk.NORMAL if HAS_OUTLOOK_IMPORT else tk.DISABLED,
        )

        # Recent files submenu — most-recently-used import sources.
        self._recent_menu = _menu(file_menu)
        file_menu.add_cascade(label='Open Recent', menu=self._recent_menu)
        self._rebuild_recent_menu()

        file_menu.add_separator()
        file_menu.add_command(label='Export to PDF…', accelerator='Ctrl+Shift+P', command=self._export_pdf)
        file_menu.add_command(label='Export to Word…', accelerator='Ctrl+Shift+W', command=self._export_word)
        file_menu.add_command(label='Export Selected to PDF…', command=lambda: self._export_pdf(selection_only=True))
        file_menu.add_command(label='Export Selected to Word…', command=lambda: self._export_word(selection_only=True))
        file_menu.add_separator()
        file_menu.add_command(label='Exit', command=root.quit)
        menubar.add_cascade(label='File', menu=file_menu)

        edit_menu = _menu(menubar)
        edit_menu.add_command(label='Select All', accelerator='Ctrl+A', command=self._select_all)
        edit_menu.add_command(label='Remove Selected', accelerator='Delete', command=self._remove_selected)
        edit_menu.add_separator()
        edit_menu.add_command(label='Clear All', command=self._clear_all)
        menubar.add_cascade(label='Edit', menu=edit_menu)

        view_menu = _menu(menubar)
        self._menu_preview_var = tk.BooleanVar(value=self.settings.show_preview)
        view_menu.add_checkbutton(
            label='Show Preview Pane',
            variable=self._menu_preview_var,
            command=self._toggle_preview,
        )
        view_menu.add_separator()
        view_menu.add_command(label='Sort Oldest First', command=lambda: self._set_sort(True))
        view_menu.add_command(label='Sort Newest First', command=lambda: self._set_sort(False))
        menubar.add_cascade(label='View', menu=view_menu)

        tools_menu = _menu(menubar)
        tools_menu.add_command(label='Options…', accelerator='Ctrl+,', command=self._open_options)
        tools_menu.add_command(label='Review Last Import…', command=self._show_last_import_report)
        tools_menu.add_command(label='Cancel Active Import', command=self._cancel_active_import)
        tools_menu.add_separator()
        tools_menu.add_command(label='Build Indexed Bundle…', accelerator='Ctrl+Shift+B', command=self._open_bundle_dialog)
        tools_menu.add_separator()
        tools_menu.add_command(label='Diagnostics…', command=self._show_diagnostics)
        menubar.add_cascade(label='Tools', menu=tools_menu)

        help_menu = _menu(menubar)
        help_menu.add_command(label='Startup Check Report…', command=self._show_startup_checks_dialog)
        help_menu.add_command(label='Crash Reports…', command=self._show_crash_reports_dialog)
        help_menu.add_command(label='Check for Updates…', command=self._check_for_updates_interactive)
        help_menu.add_separator()
        help_menu.add_command(label=f'About MailWeave {self.VERSION}', command=self._show_about)
        menubar.add_cascade(label='Help', menu=help_menu)

        root.config(menu=menubar)

    def _build_sidebar(self):
        theme = self.T
        side = self.sidebar

        # Sidebar Branding
        brand_f = tk.Frame(side, bg=theme['sidebar'], pady=30)
        brand_f.pack(fill=tk.X)
        
        self._logo_photo = load_logo_photo((120, 100))
        self.brand_canvas = tk.Label(brand_f, image=self._logo_photo, bg=theme['sidebar'])
        self.brand_canvas.pack()
        
        tk.Label(brand_f, text='MAILWEAVE', bg=theme['sidebar'], fg=theme['fg'], 
                 font=('Segoe UI Semibold', 14, 'bold')).pack(pady=(10, 0))

        # Navigation / Primary Actions
        nav = tk.Frame(side, bg=theme['sidebar'], padx=12)
        nav.pack(fill=tk.BOTH, expand=True)

        from brand_assets import load_ui_icon
        self._nav_icons = {}

        def _section(text):
            tk.Label(nav, text=text, bg=theme['sidebar'], fg=theme['fgdim'],
                     font=('Segoe UI', 8, 'bold')).pack(anchor='w',
                                                        padx=8, pady=(18, 6))

        def _nav_btn(text, icon_name, cmd, accent=False):
            img = load_ui_icon(icon_name, (22, 22))
            if img:
                self._nav_icons[icon_name] = img
            base_bg = theme['accent'] if accent else theme['sidebar']
            base_fg = '#FFFFFF' if accent else theme['fg']
            hover_bg = theme['accent2'] if accent else theme['surface']
            kwargs = dict(
                text=f"   {text}",
                command=cmd,
                bg=base_bg,
                fg=base_fg,
                activebackground=hover_bg,
                activeforeground='#FFFFFF' if accent else theme['fg'],
                font=('Segoe UI Semibold', 10),
                relief='flat', borderwidth=0, anchor='w',
                padx=12, pady=9, cursor='hand2',
            )
            if img is not None:
                kwargs.update(image=img, compound=tk.LEFT)
            btn = tk.Button(nav, **kwargs)
            btn.pack(fill=tk.X, pady=2)

            # Real hover effect — tk doesn't do it by default.
            def _enter(_e, b=btn, bg=hover_bg):
                if str(b['state']) != 'disabled':
                    b.configure(bg=bg)
            def _leave(_e, b=btn, bg=base_bg):
                if str(b['state']) != 'disabled':
                    b.configure(bg=bg)
            btn.bind('<Enter>', _enter)
            btn.bind('<Leave>', _leave)
            return btn

        _section('EMAILS')
        _nav_btn('Open Files', 'files', self._browse_files)
        _nav_btn('Open Folder', 'folder', self._browse_folder)
        self.btn_import_outlook = _nav_btn('Import from Outlook', 'outlook',
                                           self._import_outlook_selection)
        if not HAS_OUTLOOK_IMPORT:
            # Outlook integration was not detected at startup; disable the
            # button rather than silently failing on click.
            try:
                self.btn_import_outlook.configure(
                    state=tk.DISABLED,
                    fg=theme['fgdim'],
                    cursor='arrow',
                )
            except Exception:
                pass

        _section('ORGANISE')
        self.btn_view_emails = _nav_btn('Email View', 'files',
                                        lambda: self._set_view('emails'),
                                        accent=True)
        self.btn_view_annexures = _nav_btn('Annexures', 'annexure',
                                           lambda: self._set_view('annexures'))

        _section('BUNDLE')
        _nav_btn('Indexed Bundle', 'bundle', self._open_bundle_dialog)

        _section('SYSTEM')
        _nav_btn('Settings', 'settings', self._open_options)
        _nav_btn('Diagnostics', 'info', self._show_diagnostics)

        # Stats Card in Sidebar
        stats_f = tk.Frame(side, bg=theme['surface'], padx=20, pady=20)
        stats_f.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.side_count_lbl = tk.Label(stats_f, text='0', bg=theme['surface'], fg=theme['fg'], font=('Segoe UI Bold', 20))
        self.side_count_lbl.pack(anchor='w')
        tk.Label(stats_f, text='Emails Loaded', bg=theme['surface'], fg=theme['fgsub'], font=('Segoe UI', 9)).pack(anchor='w')

    def _build_header(self):
        theme = self.T
        self.header_bar = tk.Frame(self.content_wrap, bg=theme['sidebar'], height=64)
        self.header_bar.pack(fill=tk.X)
        self.header_bar.pack_propagate(False)

        # Left part of header: Search
        search_f = tk.Frame(self.header_bar, bg=theme['sidebar'], padx=20)
        search_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *args: self._on_search_change())
        
        search_inner = tk.Frame(search_f, bg=theme['surface2'], padx=10, pady=5)
        search_inner.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=12)
        
        tk.Label(search_inner, text='🔍', bg=theme['surface2'], fg=theme['fgsub']).pack(side=tk.LEFT)
        self.search_entry = tk.Entry(search_inner, textvariable=self.search_var, bg=theme['surface2'],
                                    fg=theme['fg'], insertbackground=theme['fg'], relief='flat', 
                                    font=('Segoe UI', 10), borderwidth=0)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Right part of header: Actions
        act_f = tk.Frame(self.header_bar, bg=theme['sidebar'], padx=20)
        act_f.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.thread_var = tk.BooleanVar(value=False)
        tk.Checkbutton(act_f, text='Thread View', variable=self.thread_var, command=self._refresh_tree_async,
                       bg=theme['sidebar'], fg=theme['fgsub'], selectcolor='#FFFFFF',
                       activebackground=theme['sidebar'], activeforeground=theme['fg'],
                       font=('Segoe UI', 9), cursor='hand2').pack(side=tk.LEFT, padx=15)

        self._redact_var = tk.BooleanVar(value=False)
        tk.Checkbutton(act_f, text='Redact PII', variable=self._redact_var,
                       command=lambda: self._on_select(),
                       bg=theme['sidebar'], fg=theme['fgsub'], selectcolor='#FFFFFF',
                       activebackground=theme['sidebar'], activeforeground=theme['fg'],
                       font=('Segoe UI', 9), cursor='hand2').pack(side=tk.LEFT, padx=15)

        ttk.Button(act_f, text='Export PDF', command=self._export_pdf, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(act_f, text='Export Word', command=self._export_word).pack(side=tk.LEFT, padx=5)

    def _build_content(self):
        theme = self.T
        paned = ttk.PanedWindow(self.content_wrap, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self._paned = paned

        # Email Card (Left)
        self.emails_card = tk.Frame(paned, bg=theme['sidebar'], highlightbackground=theme['border'], highlightthickness=1)
        paned.add(self.emails_card, weight=5)

        # Annexures Card (Alternative Left)
        self.annexures_card = tk.Frame(paned, bg=theme['sidebar'], highlightbackground=theme['border'], highlightthickness=1)
        # Not added to paned yet

        self.tree_wrap = tk.Frame(self.emails_card, bg=theme['sidebar'], padx=15, pady=15)
        self.tree_wrap.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            self.tree_wrap,
            columns=('num', 'date', 'att', 'from', 'to', 'subject'),
            show='headings',
            selectmode='extended',
        )
        # All columns sort on header click; the active column shows ▲/▼.
        # 'num' always reflects the active sort, so we don't bind it directly.
        self._sort_key = 'date'
        self._sort_ascending = self.settings.sort_oldest_first
        self.tree.heading('num', text='#')
        self.tree.heading('date', text='Date', command=lambda: self._sort_by('date'))
        self.tree.heading('att', text='📎', command=lambda: self._sort_by('att'))
        self.tree.heading('from', text='From', command=lambda: self._sort_by('from'))
        self.tree.heading('to', text='To', command=lambda: self._sort_by('to'))
        self.tree.heading('subject', text='Subject', command=lambda: self._sort_by('subject'))

        self.tree.column('num', width=42, minwidth=42, stretch=False)
        self.tree.column('date', width=166, minwidth=126)
        self.tree.column('att', width=28, minwidth=28, stretch=False, anchor='center')
        self.tree.column('from', width=220, minwidth=160)
        self.tree.column('to', width=200, minwidth=140)
        self.tree.column('subject', width=360, minwidth=180)

        tree_scroll = ttk.Scrollbar(self.tree_wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        # Right-click (Windows/Linux: Button-3; macOS reuses Button-2).
        self.tree.bind('<Button-3>', self._show_tree_context_menu)
        self.tree.bind('<Button-2>', self._show_tree_context_menu)

        if self._dnd_ready:
            self._register_drop_target(self.tree)
            self._register_drop_target(self.emails_card)

        # Welcome / Empty State Overlay
        self.welcome_overlay = tk.Frame(self.emails_card, bg=theme['sidebar'])
        self._reg(self.welcome_overlay, bg='sidebar')
        
        welcome_inner = tk.Frame(self.welcome_overlay, bg=theme['sidebar'])
        self._reg(welcome_inner, bg='sidebar')
        welcome_inner.place(relx=0.5, rely=0.5, anchor='center')

        self._logo_large = load_logo_photo((280, 220))
        self.welcome_logo = tk.Label(welcome_inner, image=self._logo_large, bg=theme['sidebar'])
        self._reg(self.welcome_logo, bg='sidebar')
        self.welcome_logo.pack(pady=(0, 20))
        
        self._reg(
            tk.Label(welcome_inner, text='Ready to weave your emails?', bg=theme['sidebar'], fg=theme['fg'], font=('Segoe UI Semibold', 20)),
            bg='sidebar', fg='fg'
        ).pack()
        
        self._reg(
            tk.Label(welcome_inner, text='Drop .msg or .eml files anywhere or use the sidebar to begin.', 
                     bg=theme['sidebar'], fg=theme['fgsub'], font=('Segoe UI', 11)),
            bg='sidebar', fg='fgsub'
        ).pack(pady=10)

        # Integrated Loading UI in Overlay
        self.overlay_loading_f = tk.Frame(welcome_inner, bg=theme['sidebar'])
        self._reg(self.overlay_loading_f, bg='sidebar')
        
        self.overlay_status_lbl = tk.Label(self.overlay_loading_f, text='Importing...', bg=theme['sidebar'], fg=theme['accent'], font=('Segoe UI Semibold', 10))
        self._reg(self.overlay_status_lbl, bg='sidebar', fg='accent')
        self.overlay_status_lbl.pack(pady=(20, 5))
        self.overlay_progress = ttk.Progressbar(
            self.overlay_loading_f, orient=tk.HORIZONTAL, length=300, mode='determinate', style='Hero.Horizontal.TProgressbar'
        )
        self.overlay_progress.pack(pady=5)
        
        # --- Annexures UI ---
        self.ann_tree_wrap = tk.Frame(self.annexures_card, bg=theme['sidebar'], padx=15, pady=15)
        self.ann_tree_wrap.pack(fill=tk.BOTH, expand=True)
        
        self.ann_tree = ttk.Treeview(
            self.ann_tree_wrap,
            columns=('id', 'filename', 'size', 'email'),
            show='headings',
            selectmode='browse'
        )
        self.ann_tree.heading('id', text='ID')
        self.ann_tree.heading('filename', text='Filename')
        self.ann_tree.heading('size', text='Size')
        self.ann_tree.heading('email', text='Parent Email')
        
        self.ann_tree.column('id', width=100, stretch=False)
        self.ann_tree.column('filename', width=250)
        self.ann_tree.column('size', width=80, stretch=False)
        self.ann_tree.column('email', width=300)
        
        ann_scroll = ttk.Scrollbar(self.ann_tree_wrap, orient='vertical', command=self.ann_tree.yview)
        self.ann_tree.configure(yscrollcommand=ann_scroll.set)
        self.ann_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ann_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.ann_tree.bind('<<TreeviewSelect>>', self._on_ann_select)
        self.ann_tree.bind('<Double-1>', lambda e: self._open_annexure())

        self._preview_frame = tk.Frame(paned, bg=theme['sidebar'], highlightbackground=theme['border'], highlightthickness=1)
        paned.add(self._preview_frame, weight=4)

        self.preview_wrap = tk.Frame(self._preview_frame, bg=theme['sidebar'], padx=15, pady=15)
        self.preview_wrap.pack(fill=tk.BOTH, expand=True)

        font_size = {'small': 9, 'medium': 10, 'large': 12}.get(self.settings.font_size, 10)
        self.preview = tk.Text(
            self.preview_wrap,
            wrap=tk.WORD,
            bg=theme['surface'],
            fg=theme['fg'],
            insertbackground=theme['fg'],
            font=('Segoe UI', font_size),
            relief='flat',
            padx=20,
            pady=18,
            state=tk.DISABLED,
            cursor='arrow',
        )
        preview_scroll = ttk.Scrollbar(self.preview_wrap, orient='vertical', command=self.preview.yview)
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.preview.tag_configure('hdr_lbl', font=('Segoe UI Semibold', 9, 'bold'), foreground=theme['accent'])
        self.preview.tag_configure('hdr_val', font=('Segoe UI', 9), foreground=theme['fg'])
        self.preview.tag_configure('sep', foreground=theme['fgdim'])
        self.preview.tag_configure('body', font=('Segoe UI', font_size), foreground=theme['fg'])

    def _build_statusbar(self):
        theme = self.T
        bar = self._reg(tk.Frame(self.root, bg=theme['surface'], height=28), bg='surface')
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        tk.Frame(bar, bg=theme['border'], height=1).pack(fill=tk.X, side=tk.TOP)

        self.status_left = self._reg(
            tk.Label(
                bar,
                text='Ready',
                anchor='w',
                bg=theme['surface'],
                fg=theme['fgsub'],
                font=('Segoe UI', 8),
            ),
            bg='surface',
            fg='fgsub',
        )
        self.status_left.pack(side=tk.LEFT, padx=10, fill=tk.Y)

        self.status_right = self._reg(
            tk.Label(
                bar,
                text=self.VERSION,
                anchor='e',
                bg=theme['surface'],
                fg=theme['fgdim'],
                font=('Segoe UI', 8),
            ),
            bg='surface',
            fg='fgdim',
        )
        self.status_right.pack(side=tk.RIGHT, padx=10, fill=tk.Y)

        self.progress = ttk.Progressbar(
            bar, orient=tk.HORIZONTAL, length=180, mode='determinate', style='Status.Horizontal.TProgressbar'
        )

    def _bind_shortcuts(self):
        root = self.root
        root.bind('<Control-o>', lambda _event: self._browse_files())
        root.bind('<Control-O>', lambda _event: self._browse_files())
        root.bind('<Control-a>', lambda _event: self._select_all())
        root.bind('<Control-A>', lambda _event: self._select_all())
        root.bind('<Delete>', lambda _event: self._remove_selected())
        root.bind('<Control-comma>', lambda _event: self._open_options())
        root.bind('<Control-Shift-O>', lambda _event: self._import_outlook_selection())
        root.bind('<Control-Shift-P>', lambda _event: self._export_pdf())
        root.bind('<Control-Shift-W>', lambda _event: self._export_word())
        root.bind('<Control-Shift-D>', lambda _event: self._show_diagnostics())
        root.bind('<Escape>', lambda _event: self._on_escape())
        root.bind('<Control-f>', lambda _event: self._open_preview_find())
        root.bind('<Control-F>', lambda _event: self._open_preview_find())
        root.bind('<Control-Shift-B>', lambda _event: self._open_bundle_dialog())

    def _on_escape(self):
        if self._export_busy:
            self._cancel_export()
        else:
            self._cancel_active_import()

    # ── Recent files ───────────────────────────────────────────────────────────

    _RECENT_LIMIT = 8

    def _recent_file_path(self) -> Path:
        from diagnostics import RECOVERY_DIR
        return Path(RECOVERY_DIR) / 'recent.json'

    def _load_recent(self) -> list[str]:
        try:
            import json
            data = json.loads(self._recent_file_path().read_text(encoding='utf-8'))
            if isinstance(data, list):
                return [str(p) for p in data if isinstance(p, str)][: self._RECENT_LIMIT]
        except Exception:
            pass
        return []

    def _save_recent(self):
        try:
            import json
            from diagnostics import ensure_runtime_dirs
            ensure_runtime_dirs()
            self._recent_file_path().write_text(
                json.dumps(self._recent[: self._RECENT_LIMIT]), encoding='utf-8',
            )
        except Exception:
            LOGGER.exception('recent-save-failed')

    def _push_recent(self, paths: list[str]):
        if not paths:
            return
        # Dedupe while preserving order; newest first.
        seen: set[str] = set()
        merged: list[str] = []
        for entry in list(paths) + self._recent:
            if entry not in seen and os.path.exists(entry):
                merged.append(entry)
                seen.add(entry)
        self._recent = merged[: self._RECENT_LIMIT]
        self._save_recent()
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self):
        if not hasattr(self, '_recent_menu') or self._recent_menu is None:
            return
        self._recent_menu.delete(0, 'end')
        if not self._recent:
            self._recent_menu.add_command(label='(no recent files)', state=tk.DISABLED)
            return
        for path in self._recent:
            label = path if len(path) <= 60 else '…' + path[-58:]
            self._recent_menu.add_command(
                label=label,
                command=lambda p=path: self._open_recent(p),
            )
        self._recent_menu.add_separator()
        self._recent_menu.add_command(label='Clear recent', command=self._clear_recent)

    def _open_recent(self, path: str):
        if not os.path.exists(path):
            messagebox.showinfo('Not found', f'The file or folder no longer exists:\n{path}')
            self._recent = [p for p in self._recent if p != path]
            self._save_recent()
            self._rebuild_recent_menu()
            return
        if os.path.isdir(path):
            found = sorted(
                str(p) for p in Path(path).rglob('*')
                if p.is_file() and p.suffix.lower() in {'.msg', '.eml'}
            )
            if found:
                self._load_files(found)
        else:
            self._load_files([path])

    def _clear_recent(self):
        self._recent = []
        self._save_recent()
        self._rebuild_recent_menu()

    # ── Find-in-preview ───────────────────────────────────────────────────────

    def _open_preview_find(self):
        if not self.settings.show_preview:
            return
        theme = self.T
        if getattr(self, '_find_win', None) and tk.Toplevel.winfo_exists(self._find_win):
            self._find_win.lift()
            self._find_entry.focus_set()
            return
        top = tk.Toplevel(self.root)
        self._find_win = top
        top.title('Find in preview')
        top.transient(self.root)
        top.configure(bg=theme['bg'])
        top.resizable(False, False)

        f = tk.Frame(top, bg=theme['bg'], padx=14, pady=10)
        f.pack(fill=tk.BOTH, expand=True)

        var = tk.StringVar()
        self._find_entry = tk.Entry(f, textvariable=var, width=32,
                                    bg=theme['surface2'], fg=theme['fg'],
                                    insertbackground=theme['fg'], relief='flat')
        self._find_entry.pack(side=tk.LEFT, padx=(0, 6))
        self._find_entry.focus_set()

        def _find_next(_event=None):
            needle = var.get()
            if not needle:
                return
            text = self.preview
            start = text.index(tk.INSERT) if text.index(tk.INSERT) != '1.0' else '1.0'
            pos = text.search(needle, start, nocase=True, stopindex=tk.END)
            if not pos:
                pos = text.search(needle, '1.0', nocase=True, stopindex=tk.END)
            if not pos:
                self._set_status(f'"{needle}" not found in preview.')
                return
            end = f'{pos}+{len(needle)}c'
            text.tag_remove('find_hit', '1.0', tk.END)
            text.tag_add('find_hit', pos, end)
            text.tag_configure('find_hit', background=theme['accent'], foreground='#FFFFFF')
            text.mark_set(tk.INSERT, end)
            text.see(pos)

        ttk.Button(f, text='Find next', command=_find_next).pack(side=tk.LEFT)
        ttk.Button(f, text='Close', command=top.destroy).pack(side=tk.LEFT, padx=6)
        top.bind('<Return>', _find_next)
        top.bind('<Escape>', lambda _e: top.destroy())

    # ── Redaction ─────────────────────────────────────────────────────────────

    _REDACT_PATTERNS = [
        # Emails
        re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+'),
        # International-ish phone numbers
        re.compile(r'\+?\d[\d \-]{7,}\d'),
        # AU TFN / generic 9-digit sequences
        re.compile(r'\b\d{3}[ \-]?\d{3}[ \-]?\d{3}\b'),
        # US SSN
        re.compile(r'\b\d{3}-\d{2}-\d{4}\b'),
    ]

    def _redact(self, text: str) -> str:
        if not self._redact_var or not self._redact_var.get() or not text:
            return text
        out = text
        for pat in self._REDACT_PATTERNS:
            out = pat.sub(lambda m: '█' * len(m.group(0)), out)
        return out

    def _on_close(self):
        try:
            self._persist_window_state()
            if self.settings.autosave_session:
                self._autosave_session()
                clear_session_lock()
            else:
                clear_session()
        finally:
            self.root.destroy()

    def _persist_window_state(self):
        try:
            geom = self.root.winfo_geometry()
            widths = ','.join(
                f'{col}={self.tree.column(col, "width")}'
                for col in ('num', 'date', 'att', 'from', 'to', 'subject')
            )
            if geom != self.settings.window_geometry or widths != self.settings.column_widths:
                self.settings.window_geometry = geom
                self.settings.column_widths = widths
                save_settings(self.settings)
        except Exception:
            LOGGER.exception('persist-window-state-failed')

    def _apply_saved_column_widths(self):
        raw = (self.settings.column_widths or '').strip()
        if not raw:
            return
        try:
            for part in raw.split(','):
                if '=' not in part:
                    continue
                col, value = part.split('=', 1)
                col = col.strip()
                width = int(value.strip())
                if col in ('num', 'date', 'att', 'from', 'to', 'subject') and 20 <= width <= 2000:
                    self.tree.column(col, width=width)
        except Exception:
            LOGGER.exception('column-widths-restore-failed')

    def _autosave_session(self):
        try:
            save_session(self.emails, settings_dict=self.settings.__dict__)
        except Exception as exc:
            LOGGER.exception('session-save-failed error=%s', exc)

    def _maybe_restore_session(self):
        if not self.settings.autosave_session or not has_recovery_session():
            return
        payload = load_session()
        if not payload or not payload.get('emails'):
            return
        count = payload.get('count', len(payload['emails']))
        if not messagebox.askyesno(
            'Restore previous session',
            f'MailWeave found a recoverable session with {count} email(s).\n\nRestore it now?',
        ):
            clear_session()
            return
        self.emails = payload['emails']
        self._sort_and_refresh()
        self._set_status(f'Restored {len(self.emails)} email(s) from the previous session.')

    def _show_disclaimer(self):
        theme = self.T
        top = tk.Toplevel(self.root)
        top.title('Terms of Use')
        top.geometry('520x420')
        top.configure(bg=theme['bg'])
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        top.geometry(f'+{rx + (rw-520)//2}+{ry + (rh-420)//2}')

        f = tk.Frame(top, bg=theme['bg'], padx=30, pady=30)
        f.pack(fill=tk.BOTH, expand=True)

        tk.Label(f, text='Terms of Use', bg=theme['bg'], fg=theme['fg'],
                 font=('Segoe UI Semibold', 14)).pack(anchor='w', pady=(0, 14))

        msg = (
            "MailWeave is provided as-is, without warranty of any kind, express or implied. "
            "By using this software, you accept that the developers are not liable for any "
            "data loss, export errors, or other damages arising from its use.\n\n"
            "MailWeave is a tool for organising and exporting your own email correspondence. "
            "It does not constitute legal advice. You remain responsible for the accuracy and "
            "appropriate handling of any material you import or export."
        )

        txt = tk.Text(f, bg=theme['surface2'], fg=theme['fg'], font=('Segoe UI', 10), wrap=tk.WORD,
                      padx=15, pady=15, height=9, relief='flat')
        txt.insert('1.0', msg)
        txt.config(state=tk.DISABLED)
        txt.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        def _accept():
            self.settings.disclaimer_accepted = True
            save_settings(self.settings)
            top.destroy()

        def _decline():
            # Ask before quitting so a stray click doesn't kill the session.
            if messagebox.askyesno(
                'Decline terms',
                'Declining will exit MailWeave. Read the terms again instead?',
                parent=top,
                default=messagebox.YES,
            ):
                return  # leave the dialog open
            top.destroy()
            self.root.after(50, self._on_close)

        top.protocol('WM_DELETE_WINDOW', _decline)

        btn_f = tk.Frame(f, bg=theme['bg'])
        btn_f.pack(fill=tk.X)

        ttk.Button(btn_f, text='Accept', command=_accept, style='Accent.TButton').pack(side=tk.RIGHT)
        ttk.Button(btn_f, text='Decline', command=_decline).pack(side=tk.RIGHT, padx=10)

    def _show_startup_checks(self):
        _ok, failed = summarize_checks(self.startup_checks)
        if failed:
            self._set_status('Startup checks reported issues. See Help > Startup Check Report.')
            LOGGER.warning('startup-checks-failed %s', failed)

    def _show_startup_checks_dialog(self):
        ok, failed = summarize_checks(self.startup_checks)
        lines = ['Startup checks', '']
        if ok:
            lines.append('Healthy')
            lines.extend(f'• {item}' for item in ok)
            lines.append('')
        if failed:
            lines.append('Attention needed')
            lines.extend(f'• {item}' for item in failed)
        messagebox.showinfo('Startup Check Report', '\n'.join(lines))

    def _show_diagnostics(self):
        report = build_diagnostic_report(
            {
                'email_count': len(self.emails),
                'loading': self._loading,
                'theme': self.settings.theme,
                'safe_mode_import': self.settings.safe_mode_import,
            }
        )
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
        except Exception:
            pass
        messagebox.showinfo(
            'Diagnostics',
            'A diagnostic report has been copied to the clipboard.\n\n'
            f'{report}',
        )

    def _show_last_import_report(self):
        report = self._last_import_report
        if not report:
            messagebox.showinfo('Import review', 'No import report is available yet.')
            return
        lines = [
            f"Source: {report.get('source', 'files')}",
            f"Loaded: {report.get('loaded', 0)}",
            f"Duplicates skipped: {report.get('duplicates', 0)}",
            f"Failed: {report.get('failed', 0)}",
        ]
        duplicate_examples = report.get('duplicate_examples') or []
        failed_examples = report.get('failed_examples') or []
        if duplicate_examples:
            lines.append('')
            lines.append('Duplicate examples')
            lines.extend(f'• {item}' for item in duplicate_examples[:8])
        if failed_examples:
            lines.append('')
            lines.append('Failed examples')
            lines.extend(f'• {item}' for item in failed_examples[:8])
        messagebox.showinfo('Import review', '\n'.join(lines))

    def _cancel_active_import(self):
        if not self._loading:
            return
        self._load_cancel.set()
        self._set_status('Cancelling import…')

    def _safe_import_paths(self, paths: list[str]) -> list[str]:
        if not self.settings.safe_mode_import:
            return paths
        safe_root = Path(os.environ.get('TEMP') or os.environ.get('TMP') or Path.home() / 'AppData' / 'Local' / 'Temp')
        # If the source is already under the temp tree (e.g. files dropped from
        # Outlook into our own staging dir) the copy is pointless and just
        # doubles I/O on large mailbox imports.
        try:
            safe_root_resolved = safe_root.resolve()
        except Exception:
            safe_root_resolved = safe_root
        if all(
            (lambda p: p == safe_root_resolved or safe_root_resolved in p.parents)(Path(src).resolve())
            for src in paths
        ):
            return paths

        target_dir = safe_root / 'MailWeave' / f'import_{datetime.now().strftime("%Y%m%d_%H%M%S_%f")}'
        target_dir.mkdir(parents=True, exist_ok=True)
        copied: list[str] = []
        for index, source in enumerate(paths, start=1):
            src_path = Path(source)
            dest = target_dir / f'{index:05d}_{src_path.name}'
            shutil.copy2(src_path, dest)
            copied.append(str(dest))
        return copied

    def _cleanup_safe_import_paths(self, paths: list[str]):
        if not self.settings.safe_mode_import or not paths:
            return
        try:
            parent = Path(paths[0]).parent
            if parent.name.startswith('import_') and parent.parent.name == 'MailWeave':
                shutil.rmtree(parent, ignore_errors=True)
        except Exception:
            LOGGER.exception('safe-import-cleanup-failed')

    def _confirm_export_target(self, path: str) -> bool:
        if not path:
            return False
        if os.path.exists(path) and self.settings.confirm_export_overwrite:
            return messagebox.askyesno(
                'Overwrite export',
                f'The file already exists:\n{path}\n\nOverwrite it?',
            )
        return True

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        file_paths = [
            path for path in paths
            if os.path.isfile(path) and path.lower().endswith(('.msg', '.eml'))
        ]
        if file_paths:
            self._load_files(file_paths)
            return
        # Outlook often drops virtual items rather than filesystem paths.
        self._import_outlook_selection(from_drop=True)

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title='Select Email Files',
            filetypes=[
                ('Email files', '*.msg *.eml'),
                ('Outlook messages', '*.msg'),
                ('EML files', '*.eml'),
                ('All files', '*.*'),
            ],
        )
        if paths:
            self._push_recent(list(paths))
            self._load_files(paths)

    def _browse_folder(self):
        folder = filedialog.askdirectory(
            title='Select a folder of .msg / .eml emails',
            parent=self.root,
        )
        if not folder:
            return
        try:
            found = sorted(
                str(path)
                for path in Path(folder).rglob('*')
                if path.is_file() and path.suffix.lower() in {'.msg', '.eml'}
            )
        except Exception as exc:
            LOGGER.exception('open-folder-scan-failed folder=%s', folder)
            messagebox.showerror(
                'Could not read folder',
                f'Failed to scan that folder:\n\n{exc}',
                parent=self.root,
            )
            return
        if found:
            self._push_recent([folder])
            self._load_files(found)
        else:
            messagebox.showinfo(
                'No emails found',
                'No .msg or .eml files were found in that folder or its '
                'subfolders.\n\nTo bundle PDFs or Word documents, use '
                'Bundle → Indexed Bundle from the sidebar instead.',
                parent=self.root,
            )

    def _import_outlook_selection(self, from_drop: bool = False):
        if self._loading:
            if not from_drop:
                messagebox.showinfo(
                    'Still loading',
                    'MailWeave is already importing emails. Please wait for the current batch to finish.',
                )
            return
        if not HAS_OUTLOOK_IMPORT:
            if not from_drop:
                messagebox.showerror(
                    'Outlook import unavailable',
                    'Outlook integration is not installed in this build.',
                )
            return

        try:
            total = get_outlook_selection_count()
        except OutlookImportError as exc:
            if not from_drop:
                messagebox.showerror('Outlook import', str(exc))
            else:
                self._set_status('Outlook drag-and-drop is not available from this Outlook view. Use Import Outlook.')
            return

        self._loading = True
        self._load_cancel.clear()
        self._load_queue = Queue()
        self._load_job = {
            'total': total,
            'known': {email.unique_key for email in self.emails},
            'known_fingerprints': {email.duplicate_fingerprint for email in self.emails},
            'loaded': [],
            'loaded_count': 0,
            'skipped': 0,
            'failed': 0,
            'done': 0,
            'source_label': 'Outlook',
            'fatal_error': '',
            'duplicate_examples': [],
            'failed_examples': [],
        }
        self._set_status(f'Importing 0 / {total} Outlook emails…')

        worker = Thread(target=self._load_outlook_worker, args=(self._load_queue,), daemon=True)
        worker.start()
        self._show_loading_ui(True)
        self.root.after(40, self._drain_load_queue)

    def _load_files(self, paths):
        if self._loading:
            messagebox.showinfo(
                'Still loading',
                'MailWeave is already importing emails. Please wait for the current batch to finish.',
            )
            return

        clean_paths = [
            raw.strip('{}').strip('"').strip()
            for raw in paths
            if raw and raw.strip('{}').strip('"').strip()
        ]
        candidates = [
            path
            for path in clean_paths
            if os.path.isfile(path) and path.lower().endswith(('.msg', '.eml'))
        ]
        if not candidates:
            self._set_status('Nothing loaded.')
            return

        if len(candidates) > 500:
            LOGGER.warning('large-batch-import count=%s', len(candidates))
            self._set_status(f'Large batch mode enabled for {len(candidates)} emails.')

        self._loading = True
        self._load_cancel.clear()
        self._load_queue = Queue()
        self._load_job = {
            'total': len(candidates),
            'known': {email.unique_key for email in self.emails},
            'known_fingerprints': {email.duplicate_fingerprint for email in self.emails},
            'loaded': [],
            'loaded_count': 0,
            'skipped': 0,
            'failed': 0,
            'done': 0,
            'source_label': 'Files',
            'fatal_error': '',
            'duplicate_examples': [],
            'failed_examples': [],
        }
        self._set_status(f'Loading 0 / {len(candidates)} emails…')

        worker = Thread(
            target=self._load_files_worker,
            args=(candidates, self._load_queue),
            daemon=True,
        )
        worker.start()
        self._show_loading_ui(True)
        self.root.after(40, self._drain_load_queue)

    def _load_files_worker(self, paths: list[str], queue: Queue):
        try:
            paths = self._safe_import_paths(paths)
        except Exception as exc:
            queue.put(('__error__', f'Safe-mode import failed: {exc}'))
            queue.put(None)
            return

        max_workers = min(8, max(2, (os.cpu_count() or 4)))

        def _parse(path: str):
            if self._load_cancel.is_set():
                return path, None
            return path, parse_email_file(path)

        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(_parse, path) for path in paths]
                for future in as_completed(futures):
                    if self._load_cancel.is_set():
                        break
                    try:
                        queue.put(future.result())
                    except Exception as exc:
                        LOGGER.warning('parse-worker-future-failed error=%s', exc)
                        queue.put((None, None))
        finally:
            self._cleanup_safe_import_paths(paths)
            queue.put(None)

    def _load_outlook_worker(self, queue: Queue):
        try:
            for data in iter_selected_outlook_emails():
                if self._load_cancel.is_set():
                    break
                queue.put(('outlook', data))
        except OutlookImportError as exc:
            queue.put(('__error__', str(exc)))
        except Exception as exc:
            queue.put(('__error__', f'Unexpected Outlook import failure: {exc}'))
        finally:
            queue.put(None)

    def _drain_load_queue(self):
        if not self._loading or self._load_queue is None or self._load_job is None:
            return

        if self._load_cancel.is_set():
            self._finish_load_job()
            return

        processed_any = False
        finished = False
        job = self._load_job
        processed_count = 0
        max_per_tick = 150

        while True:
            if processed_count >= max_per_tick:
                break
            try:
                item = self._load_queue.get_nowait()
            except Empty:
                break

            processed_any = True
            processed_count += 1
            if item is None:
                finished = True
                break

            if item[0] == '__error__':
                job['fatal_error'] = item[1]
                job['failed'] += max(0, job['total'] - job['done'])
                finished = True
                break

            _path, data = item
            job['done'] += 1
            if data is None:
                job['failed'] += 1
                if _path not in ('outlook', '__error__') and len(job['failed_examples']) < 10:
                    job['failed_examples'].append(os.path.basename(str(_path)))
                continue
            if data.unique_key in job['known'] or data.duplicate_fingerprint in job['known_fingerprints']:
                job['skipped'] += 1
                if len(job['duplicate_examples']) < 10:
                    job['duplicate_examples'].append(data.subject or os.path.basename(data.source_file))
                continue
            job['known'].add(data.unique_key)
            job['known_fingerprints'].add(data.duplicate_fingerprint)
            job['loaded'].append(data)
            job['loaded_count'] += 1

        if processed_any:
            total = job['total']
            source_label = job.get('source_label', 'Loading')
            done = job['done']
            percent = (done / total * 100) if total > 0 else 0
            
            self.progress['value'] = percent
            self.overlay_progress['value'] = percent
            
            self._set_status(
                f"{source_label} {done} / {total} emails…  "
                f"Loaded {job['loaded_count']}  •  "
                f"Skipped {job['skipped']}  •  Failed {job['failed']}"
            )
            self.overlay_status_lbl.config(
                text=f"Importing from {source_label}: {done} of {total} processed..."
            )

        if finished:
            if job['loaded']:
                self.emails.extend(job['loaded'])
                self._sort_and_refresh()
            self._finish_load_job()
            return

        self.root.after(40, self._drain_load_queue)

    def _finish_load_job(self):
        job = self._load_job or {}
        parts = []
        cancelled = job.get('done', 0) < job.get('total', 0) and not job.get('fatal_error')
        if cancelled:
            parts.append('Import cancelled')
        if job.get('loaded_count'):
            parts.append(f"loaded {job['loaded_count']} email(s)")
        if job.get('skipped'):
            parts.append(f"skipped {job['skipped']} duplicate(s)")
        if job.get('failed'):
            parts.append(f"failed {job['failed']} file(s)")
        if parts:
            message = parts[0][0].upper() + parts[0][1:] + (', ' + ', '.join(parts[1:]) if len(parts) > 1 else '') + '.'
        else:
            message = 'Nothing loaded.'

        self._loading = False
        self._show_loading_ui(False)
        self._load_queue = None
        self._load_job = None
        self._load_cancel.clear()
        self._set_status(message)
        self._last_import_report = {
            'source': job.get('source_label', 'Files'),
            'loaded': job.get('loaded_count', 0),
            'duplicates': job.get('skipped', 0),
            'failed': job.get('failed', 0),
            'duplicate_examples': job.get('duplicate_examples', []),
            'failed_examples': job.get('failed_examples', []),
        }
        if job.get('loaded_count') and self.settings.autosave_session:
            self._autosave_session()

        if job.get('fatal_error'):
            messagebox.showerror('Import failed', job['fatal_error'])

        # _sort_and_refresh was already invoked above when new emails landed.
        # Re-running it here re-assigned annexure IDs after the tree had
        # rendered with the prior set — kept only as a no-load safety net.
        if not job.get('loaded'):
            self._update_stats()

    def _sort_by_date(self):
        # Legacy entry point used by the View menu — preserves old behaviour.
        self._sort_by('date')

    def _set_sort(self, oldest_first: bool):
        self.settings.sort_oldest_first = oldest_first
        self._sort_key = 'date'
        self._sort_ascending = oldest_first
        save_settings(self.settings)
        self._sort_and_refresh()

    def _sort_by(self, key: str):
        # Re-clicking the active column flips direction.
        if self._sort_key == key:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_key = key
            # Date defaults to oldest-first; others to A→Z which feels natural.
            self._sort_ascending = True
        if key == 'date':
            self.settings.sort_oldest_first = self._sort_ascending
            save_settings(self.settings)
        self._sort_and_refresh()
        self._update_sort_indicators()

    def _update_sort_indicators(self):
        labels = {
            'num': '#', 'date': 'Date', 'att': '📎',
            'from': 'From', 'to': 'To', 'subject': 'Subject',
        }
        arrow = '  ▲' if self._sort_ascending else '  ▼'
        for col, base in labels.items():
            self.tree.heading(col, text=base + (arrow if col == self._sort_key else ''))

    def _sort_and_refresh(self):
        ascending = self._sort_ascending
        key = self._sort_key
        if key == 'date':
            self.emails.sort(key=email_timestamp, reverse=not ascending)
        elif key == 'from':
            self.emails.sort(key=lambda e: (e.sender or '').lower(), reverse=not ascending)
        elif key == 'to':
            self.emails.sort(key=lambda e: (e.recipients or '').lower(), reverse=not ascending)
        elif key == 'subject':
            self.emails.sort(key=lambda e: (e.subject or '').lower(), reverse=not ascending)
        elif key == 'att':
            # Sort by attachment count then by date so emails with attachments group.
            self.emails.sort(
                key=lambda e: (len(e.attachments), email_timestamp(e)),
                reverse=not ascending,
            )
        else:
            self.emails.sort(key=email_timestamp, reverse=not ascending)
        auto_assign_annexures(self.emails)
        self._refresh_tree_async()
        if self._current_view == 'annexures':
            self._refresh_annexures_tree()
        self._update_stats()

    def _refresh_tree_async(self):
        if self._tree_render_after:
            try:
                self.root.after_cancel(self._tree_render_after)
            except Exception:
                pass
            self._tree_render_after = None
        
        # Cancel any pending deletion cycle
        if hasattr(self, '_delete_after') and self._delete_after:
            try:
                self.root.after_cancel(self._delete_after)
            except Exception:
                pass
            self._delete_after = None

        self._display_emails = self._get_display_emails()
        self._delete_list = list(self.tree.get_children())
        self._delete_tree_chunks()
        if self._current_view == 'annexures':
            self._refresh_annexures_tree()

    def _get_display_emails(self) -> list[EmailData]:
        query = self.search_var.get().lower().strip()
        emails = self.emails

        if query:
            # Uses precomputed lowercased index — see EmailData.search_index.
            emails = [e for e in emails if query in e.search_index]

        if self.thread_var.get():
            # Naive grouping: show only the first message of each thread_id
            seen_threads = set()
            threaded = []
            for e in emails:
                if e.thread_id not in seen_threads:
                    threaded.append(e)
                    seen_threads.add(e.thread_id)
            return threaded

        return emails

    def _on_search_change(self):
        if self._search_after:
            self.root.after_cancel(self._search_after)
        self._search_after = self.root.after(250, self._refresh_tree_async)

    def _delete_tree_chunks(self):
        if not hasattr(self, '_delete_list') or not self._delete_list:
            self._delete_after = None
            self._render_tree_chunk(0)
            return
        
        chunk = self._delete_list[:800]
        self._delete_list = self._delete_list[800:]
        try:
            if chunk:
                self.tree.delete(*chunk)
        except Exception:
            pass
        
        if self._delete_list:
            self._delete_after = self.root.after(1, self._delete_tree_chunks)
        else:
            self._delete_after = None
            self._render_tree_chunk(0)

    def _render_tree_chunk(self, start: int, chunk_size: int = 250):
        end = min(len(self._display_emails), start + chunk_size)
        for index in range(start, end):
            email = self._display_emails[index]
            att_icon = '📎' if email.attachments else ''
            self.tree.insert(
                '',
                'end',
                iid=str(index),
                values=(
                    index + 1,
                    _fmt_short(email.date, email.date_str, self.settings.date_format),
                    att_icon,
                    _trunc(email.sender, 40),
                    _trunc(email.recipients, 34),
                    _trunc(email.subject, 60),
                ),
            )

        if end < len(self._display_emails):
            self._tree_render_after = self.root.after(1, lambda: self._render_tree_chunk(end, chunk_size))
        else:
            self._tree_render_after = None

    def _set_view(self, view: str):
        if view == self._current_view:
            return
        
        theme = self.T
        if view == 'emails':
            self._paned.forget(self.annexures_card)
            self._paned.insert(0, self.emails_card, weight=5)
            self.btn_view_emails.configure(bg=theme['accent'], fg='#FFFFFF')
            self.btn_view_annexures.configure(bg=theme['sidebar'], fg=theme['fgsub'])
        else:
            self._paned.forget(self.emails_card)
            self._paned.insert(0, self.annexures_card, weight=5)
            self.btn_view_annexures.configure(bg=theme['accent'], fg='#FFFFFF')
            self.btn_view_emails.configure(bg=theme['sidebar'], fg=theme['fgsub'])
            self._refresh_annexures_tree()
            
        self._current_view = view
        self._clear_preview()

    def _refresh_annexures_tree(self):
        self.ann_tree.delete(*self.ann_tree.get_children())
        from annexures import get_all_attachments
        all_atts = get_all_attachments(self.emails)
        
        for email, att in all_atts:
            self.ann_tree.insert('', 'end', values=(
                att.annexure_id or '—',
                att.filename,
                f"{att.size // 1024} KB",
                email.subject[:50]
            ), tags=(id(att),))

    def _on_ann_select(self, _event=None):
        selected = self.ann_tree.selection()
        if not selected:
            return
        
        # We need to find the attachment object
        # Since I used tags=(id(att),), I can use that to find it
        tag = self.ann_tree.item(selected[0], 'tags')[0]
        from annexures import get_all_attachments
        for email, att in get_all_attachments(self.emails):
            if str(id(att)) == str(tag):
                self._show_ann_preview(email, att)
                break

    def _show_ann_preview(self, email: EmailData, att):
        text = self.preview
        text.config(state=tk.NORMAL)
        text.delete('1.0', tk.END)
        
        text.insert(tk.END, f"ANNEXURE DETAILS\n", 'hdr_lbl')
        text.insert(tk.END, '\u2500' * 60 + '\n', 'sep')
        
        for label, value in (
            ('ID:       ', att.annexure_id or 'Not assigned'),
            ('Filename: ', att.filename),
            ('Size:     ', f"{att.size} bytes ({att.size // 1024} KB)"),
            ('Type:     ', att.content_type),
            ('Parent:   ', email.subject),
        ):
            text.insert(tk.END, label, 'hdr_lbl')
            text.insert(tk.END, value + '\n', 'hdr_val')
            
        text.insert(tk.END, '\nDouble-click in list to open this attachment.\n', 'body')
        text.config(state=tk.DISABLED)

    # Extensions that execute by file association on Windows — blocked from
    # auto-open to prevent untrusted email content from running code.
    _DANGEROUS_EXTS = {
        '.exe', '.com', '.bat', '.cmd', '.scr', '.pif', '.lnk', '.hta',
        '.js', '.jse', '.vbs', '.vbe', '.wsf', '.wsh', '.ps1', '.psm1',
        '.msi', '.msp', '.cpl', '.iso', '.img', '.reg', '.jar', '.dll',
    }

    def _sanitize_attachment_filename(self, raw: str) -> str:
        from os.path import basename
        name = basename((raw or '').replace('\\', '/'))
        for char in '<>:"/\\|?*\0':
            name = name.replace(char, '_')
        name = name.strip('. ')
        return name or 'attachment'

    def _open_annexure(self):
        selected = self.ann_tree.selection()
        if not selected:
            return

        tag = self.ann_tree.item(selected[0], 'tags')[0]
        from annexures import get_all_attachments
        import tempfile

        for _email, att in get_all_attachments(self.emails):
            if str(id(att)) != str(tag):
                continue
            if not att.content:
                messagebox.showerror(
                    'Attachment unavailable',
                    'This attachment has no stored content. It may have been '
                    'restored from a recovery session, which does not include '
                    'attachment bytes. Re-import the original email to open it.',
                )
                return

            safe_name = self._sanitize_attachment_filename(att.filename)
            ext = os.path.splitext(safe_name)[1].lower()
            if ext in self._DANGEROUS_EXTS:
                if not messagebox.askyesno(
                    'Potentially dangerous attachment',
                    f'"{safe_name}" has a file type ({ext}) that can execute '
                    'code on your computer.\n\n'
                    'Only open this if you fully trust the sender.\n\n'
                    'Open anyway?',
                    icon='warning',
                ):
                    return

            tmp_root = Path(tempfile.gettempdir()) / 'MailWeave_Preview'
            tmp_root.mkdir(exist_ok=True)
            dest = (tmp_root / safe_name).resolve()
            try:
                dest.relative_to(tmp_root.resolve())
            except ValueError:
                messagebox.showerror(
                    'Blocked',
                    'The attachment filename resolved outside the preview '
                    'sandbox and was not opened.',
                )
                return

            try:
                dest.write_bytes(att.content)
            except Exception as exc:
                messagebox.showerror('Error', f'Could not save attachment: {exc}')
                return

            try:
                if hasattr(os, 'startfile'):
                    os.startfile(str(dest))
                else:
                    import subprocess
                    opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                    subprocess.Popen([opener, str(dest)])
            except Exception as exc:
                messagebox.showerror('Error', f'Could not open attachment: {exc}')
            return

    def _update_stats(self):
        total_count = len(self.emails)
        display_count = len(getattr(self, '_display_emails', [])) or total_count

        self.side_count_lbl.config(text=str(display_count))
        # Keep the right-side selection indicator consistent across refreshes.
        try:
            self._update_selection_indicator()
        except Exception:
            pass

        if total_count == 0:
            self.welcome_overlay.place(relx=0.5, rely=0.5, anchor='center', relwidth=1, relheight=1)
        else:
            self.welcome_overlay.place_forget()

    def _select_all(self):
        for item_id in self.tree.get_children():
            self.tree.selection_add(item_id)

    def _remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        if not messagebox.askyesno('Remove selected', f'Remove {len(selected)} selected email(s) from the current bundle?'):
            return
        targets = set()
        for item in selected:
            try:
                idx = int(item)
            except ValueError:
                continue
            if 0 <= idx < len(self._display_emails):
                targets.add(id(self._display_emails[idx]))
        if not targets:
            return
        self.emails = [email for email in self.emails if id(email) not in targets]
        self._refresh_tree_async()
        self._update_stats()
        self._clear_preview()
        if self.settings.autosave_session:
            self._autosave_session()

    def _clear_all(self):
        if not self.emails:
            return
        confirmed = True
        if self.settings.confirm_before_clear:
            confirmed = messagebox.askyesno('Clear all', 'Remove all loaded emails?')
        if confirmed:
            self.emails.clear()
            self._refresh_tree_async()
            self._update_stats()
            self._clear_preview()
            self._set_status('Cleared.')
            self._toast('Cleared all loaded emails.', kind='success')
            if self.settings.autosave_session:
                self._autosave_session()

    def _on_select(self, _event=None):
        selected = self.tree.selection()
        self._update_selection_indicator()
        if not selected:
            return
        index = int(selected[0])
        if index < len(self._display_emails):
            self._show_preview(self._display_emails[index])
            self._set_status(f'Email {index + 1}  —  {self._display_emails[index].subject}')

    def _selected_emails(self) -> list[EmailData]:
        result: list[EmailData] = []
        for item in self.tree.selection():
            try:
                idx = int(item)
            except ValueError:
                continue
            if 0 <= idx < len(self._display_emails):
                result.append(self._display_emails[idx])
        return result

    def _update_selection_indicator(self):
        selected = self._selected_emails()
        try:
            if not selected:
                self.status_right.config(text=self.VERSION)
                return
            total_bytes = sum(
                sum(a.size for a in e.attachments) for e in selected
            )
            mb = total_bytes / 1024 / 1024
            size_part = f' • {mb:.1f} MB' if total_bytes else ''
            self.status_right.config(
                text=f'{len(selected)} of {len(self.emails)} selected{size_part}'
            )
        except Exception:
            pass

    # ── Tree context menu ─────────────────────────────────────────────────────

    def _show_tree_context_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
            self.tree.focus(row)
        selected = self._selected_emails()
        if not selected:
            return

        theme = self.T
        menu = tk.Menu(
            self.root,
            tearoff=0,
            bg=theme['menu_bg'],
            fg=theme['menu_fg'],
            activebackground=theme['accent'],
            activeforeground='#FFFFFF',
            relief='flat',
            borderwidth=1,
        )

        single = selected[0] if len(selected) == 1 else None
        if single:
            menu.add_command(label='Open original file',
                             command=lambda: self._open_source_file(single))
            menu.add_command(label='Show in folder',
                             command=lambda: self._reveal_in_folder(single.source_file))
            menu.add_command(label='Copy subject',
                             command=lambda: self._copy_to_clipboard(single.subject or ''))
            menu.add_command(label='Copy sender',
                             command=lambda: self._copy_to_clipboard(single.sender or ''))
            menu.add_separator()
        menu.add_command(
            label=f'Export selected to PDF…  ({len(selected)})',
            command=lambda: self._export_pdf(selection_only=True),
        )
        menu.add_command(
            label=f'Export selected to Word…  ({len(selected)})',
            command=lambda: self._export_word(selection_only=True),
        )
        menu.add_separator()
        menu.add_command(
            label=f'Remove from bundle  ({len(selected)})',
            command=self._remove_selected,
        )

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy_to_clipboard(self, text: str):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self._toast('Copied to clipboard', kind='success')
        except Exception:
            pass

    def _open_source_file(self, email: EmailData):
        path = email.source_file
        if not path or not os.path.exists(path):
            self._toast('Original file is no longer available.', kind='warn')
            return
        try:
            if hasattr(os, 'startfile'):
                os.startfile(path)
            else:
                import subprocess
                opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                subprocess.Popen([opener, path])
        except Exception as exc:
            messagebox.showerror('Open failed', str(exc))

    def _reveal_in_folder(self, path: str):
        if not path or not os.path.exists(path):
            self._toast('Path no longer exists.', kind='warn')
            return
        try:
            if sys.platform == 'win32':
                import subprocess
                subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.Popen(['open', '-R', path])
            else:
                import subprocess
                subprocess.Popen(['xdg-open', os.path.dirname(path)])
        except Exception as exc:
            messagebox.showerror('Open failed', str(exc))

    def _show_preview(self, email: EmailData):
        from exporters import _fmt_long

        text = self.preview
        text.config(state=tk.NORMAL)
        text.delete('1.0', tk.END)

        for label, value in (
            ('Subject: ', email.subject),
            ('From:    ', email.sender),
            ('To:      ', email.recipients),
            ('Date:    ', _fmt_long(email.date, email.date_str, self.settings.date_format)),
        ):
            text.insert(tk.END, label, 'hdr_lbl')
            text.insert(tk.END, value + '\n', 'hdr_val')

        text.insert(tk.END, '\u2500' * 60 + '\n', 'sep')
        
        if email.attachments:
            text.insert(tk.END, 'Attachments: ', 'hdr_lbl')
            att_list = ', '.join(f'{a.filename} ({a.size // 1024} KB)' for a in email.attachments)
            text.insert(tk.END, att_list + '\n', 'hdr_val')
            text.insert(tk.END, '\u2500' * 60 + '\n', 'sep')

        body = email.body_clean if self.settings.strip_quotes else email.body_plain
        body = self._redact(body)
        text.insert(tk.END, body or '(No body content found.)', 'body')
        text.config(state=tk.DISABLED)
        self._preview_email = email

    def _clear_preview(self):
        self.preview.config(state=tk.NORMAL)
        self.preview.delete('1.0', tk.END)
        self.preview.insert(
            tk.END,
            'Select an email to preview its cleaned body, headers, and export-ready formatting.',
            'body',
        )
        self.preview.config(state=tk.DISABLED)

    def _get_export_dir(self) -> str:
        if (
            self.settings.remember_export_dir
            and self.settings.default_export_dir
            and os.path.isdir(self.settings.default_export_dir)
        ):
            return self.settings.default_export_dir
        return ''

    def _save_export_dir(self, filepath: str):
        if self.settings.remember_export_dir:
            self.settings.default_export_dir = os.path.dirname(filepath)
            save_settings(self.settings)

    def _check_export_target(self, path: str) -> bool:
        if not self._confirm_export_target(path):
            return False
        try:
            parent = os.path.dirname(path)
            if parent:
                os.makedirs(parent, exist_ok=True)
            with open(path, 'ab'):
                pass
        except OSError:
            messagebox.showerror(
                'Export target unavailable',
                'MailWeave could not write to the selected file.\n\n'
                'Close it in Word/Acrobat if it is already open, then try again.',
            )
            return False
        return True

    def _run_export(self, kind: str, path: str, builder, emails: list[EmailData] | None = None):
        if self._export_busy:
            messagebox.showinfo('Export in progress', 'An export is already running. Please wait for it to finish.')
            return
        emails = emails if emails is not None else self.emails
        self._export_busy = True
        self._export_cancel = Event()
        self._set_status(f'Exporting {kind}… (Esc to cancel)')
        self._show_export_progress(True, kind, len(emails))

        def progress_cb(done: int, total: int, label: str = ''):
            # Builders that opt in can call this from the worker thread.
            self.root.after(0, lambda: self._update_export_progress(done, total, label))
            if self._export_cancel.is_set():
                raise _ExportCancelled()

        def _worker():
            try:
                # Builders that accept `progress_cb` / `cancel_event` will use
                # them; older signatures still work via the fallback branch.
                try:
                    builder(path, emails, self.settings,
                            progress_cb=progress_cb,
                            cancel_event=self._export_cancel)
                except TypeError:
                    builder(path, emails, self.settings)
                verify_export_file(path)
                self.root.after(0, lambda: self._finish_export(kind, path))
            except _ExportCancelled:
                self.root.after(0, self._cancelled_export)
            except Exception as exc:
                LOGGER.exception('export-failed kind=%s path=%s', kind, path)
                self.root.after(0, lambda: self._fail_export(exc))

        Thread(target=_worker, daemon=True).start()

    def _cancel_export(self):
        if self._export_busy and getattr(self, '_export_cancel', None) is not None:
            self._export_cancel.set()
            self._set_status('Cancelling export…')

    def _cancelled_export(self):
        self._export_busy = False
        self._show_export_progress(False)
        self._set_status('Export cancelled.')
        self._toast('Export cancelled.', kind='warn')

    def _show_export_progress(self, visible: bool, kind: str = '', total: int = 0):
        if visible:
            # Reuse the existing status-bar progress bar; the import flow only
            # uses it while loading so this won't clash.
            self.progress['mode'] = 'determinate'
            self.progress['value'] = 0
            self.progress.pack(side=tk.RIGHT, before=self.status_right, padx=10, pady=4)
            self._export_total = total
        else:
            self.progress.pack_forget()
            self.progress['value'] = 0
            self._export_total = 0

    def _update_export_progress(self, done: int, total: int, label: str):
        total = total or getattr(self, '_export_total', 0) or 1
        percent = max(0, min(100, int(done / total * 100)))
        self.progress['value'] = percent
        if label:
            self._set_status(f'Exporting {label}… {done}/{total} (Esc to cancel)')

    def _finish_export(self, kind: str, path: str):
        self._export_busy = False
        self._show_export_progress(False)
        self._save_export_dir(path)
        self._set_status(f'{kind} exported — {os.path.basename(path)}')

        folder_name = f"{os.path.splitext(os.path.basename(path))[0]}_Annexures"
        self._show_export_success_dialog(kind, path, folder_name)

    def _show_export_success_dialog(self, kind: str, path: str, folder_name: str):
        theme = self.T
        top = tk.Toplevel(self.root)
        top.title('Export complete')
        top.configure(bg=theme['bg'])
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)

        body = tk.Frame(top, bg=theme['bg'], padx=24, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text=f'{kind} exported', bg=theme['bg'], fg=theme['fg'],
                 font=('Segoe UI Semibold', 14)).pack(anchor='w', pady=(0, 8))
        tk.Label(body, text=path, bg=theme['bg'], fg=theme['fgsub'],
                 font=('Segoe UI', 9), wraplength=420, justify='left').pack(anchor='w')
        tk.Label(body, text=f'Annexures folder: {folder_name}',
                 bg=theme['bg'], fg=theme['fgsub'],
                 font=('Segoe UI', 9), wraplength=420, justify='left').pack(anchor='w', pady=(4, 14))

        folder = os.path.dirname(path)

        def _open_file():
            try:
                if hasattr(os, 'startfile'):
                    os.startfile(path)
                else:
                    import subprocess
                    opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                    subprocess.Popen([opener, path])
            except Exception as exc:
                messagebox.showerror('Open failed', str(exc), parent=top)

        def _open_folder():
            try:
                if sys.platform == 'win32':
                    import subprocess
                    subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
                elif hasattr(os, 'startfile'):
                    os.startfile(folder)
                else:
                    import subprocess
                    opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                    subprocess.Popen([opener, folder])
            except Exception as exc:
                messagebox.showerror('Open failed', str(exc), parent=top)

        def _copy_path():
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(path)
            except Exception:
                pass

        btn_row = tk.Frame(body, bg=theme['bg'])
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text='Open file', command=_open_file,
                   style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(btn_row, text='Open folder', command=_open_folder).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row, text='Copy path', command=_copy_path).pack(side=tk.LEFT)
        ttk.Button(btn_row, text='Close', command=top.destroy).pack(side=tk.RIGHT)

    def _fail_export(self, exc: Exception):
        self._export_busy = False
        self._set_status('Export failed.')
        messagebox.showerror('Export failed', str(exc))

    def _emails_to_export(self, selection_only: bool) -> list[EmailData]:
        if not selection_only:
            return self.emails
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo(
                'No selection',
                'Select one or more emails in the list first to use '
                '"Export Selected".',
            )
            return []
        chosen_ids = set()
        for item in selected:
            try:
                idx = int(item)
            except ValueError:
                continue
            if 0 <= idx < len(self._display_emails):
                chosen_ids.add(id(self._display_emails[idx]))
        # Preserve the master sort order rather than the click order.
        return [e for e in self.emails if id(e) in chosen_ids]

    def _export_pdf(self, selection_only: bool = False):
        emails = self._emails_to_export(selection_only)
        if not emails:
            if not selection_only:
                messagebox.showwarning('No emails', 'Load some emails first.')
            return
        if not HAS_PDF:
            messagebox.showerror('Missing library', 'reportlab not installed.\n\nRun: pip install reportlab')
            return

        path = filedialog.asksaveasfilename(
            title='Save PDF',
            defaultextension='.pdf',
            initialdir=self._get_export_dir(),
            filetypes=[('PDF files', '*.pdf')],
        )
        if not path or not self._check_export_target(path):
            return

        self._run_export('PDF', path, build_pdf, emails)

    def _export_word(self, selection_only: bool = False):
        emails = self._emails_to_export(selection_only)
        if not emails:
            if not selection_only:
                messagebox.showwarning('No emails', 'Load some emails first.')
            return
        if not HAS_DOCX:
            messagebox.showerror('Missing library', 'python-docx not installed.\n\nRun: pip install python-docx')
            return

        path = filedialog.asksaveasfilename(
            title='Save Word document',
            defaultextension='.docx',
            initialdir=self._get_export_dir(),
            filetypes=[('Word documents', '*.docx')],
        )
        if not path or not self._check_export_target(path):
            return

        self._run_export('Word document', path, build_docx, emails)

    def _open_options(self):
        try:
            from options import OptionsDialog
            OptionsDialog(self.root, self.settings,
                          on_apply=self._apply_settings, T=self.T)
        except Exception as exc:
            LOGGER.exception('open-options-failed')
            messagebox.showerror(
                'Could not open Settings',
                f'The Settings dialog failed to open:\n\n{exc}',
            )

    def _open_bundle_dialog(self):
        try:
            from bundle_dialog import BundleDialog
        except Exception as exc:
            messagebox.showerror(
                'Bundle feature unavailable',
                f'Could not load the bundle builder: {exc}\n\n'
                'Install missing dependencies with:\n'
                '  pip install pypdf docx2pdf Pillow',
            )
            return
        BundleDialog(
            self.root,
            theme=self.T,
            default_author=self.settings.document_author or 'MailWeave',
            toast_cb=self._toast,
        )

    def _apply_settings(self, new_settings: AppSettings):
        theme_changed = new_settings.theme != self.settings.theme
        font_changed = new_settings.font_size != self.settings.font_size
        preview_changed = new_settings.show_preview != self.settings.show_preview
        strip_changed = new_settings.strip_quotes != self.settings.strip_quotes

        new_settings.disclaimer_accepted = self.settings.disclaimer_accepted
        self.settings = new_settings
        save_settings(new_settings)
        if self.settings.autosave_session:
            self._autosave_session()

        if theme_changed or font_changed:
            self._apply_theme()

        if preview_changed:
            self._update_preview_visibility()
            self._menu_preview_var.set(new_settings.show_preview)

        self._refresh_tree_async()
        if strip_changed:
            self._on_select()

        self._set_status('Settings applied.')
        self._toast('Settings applied.', kind='success')

    def _toggle_preview(self):
        self.settings.show_preview = self._menu_preview_var.get()
        save_settings(self.settings)
        self._update_preview_visibility()

    def _update_preview_visibility(self):
        already_added = str(self._preview_frame) in self._paned.panes()
        if self.settings.show_preview:
            if not already_added:
                self._paned.add(self._preview_frame, weight=4)
        else:
            if already_added:
                try:
                    self._paned.forget(self._preview_frame)
                except Exception:
                    pass

    def _show_about(self):
        ok, failed = summarize_checks(self.startup_checks)
        messagebox.showinfo(
            f'About MailWeave {self.VERSION}',
            'MailWeave bundles Outlook .msg and .eml files into clean PDF or Word documents.\n\n'
            'Strip redundant reply trails, review each message, and export a single polished file.\n\n'
            f'Version: {self.VERSION}\n'
            f'Emails loaded: {len(self.emails)}\n'
            f'Startup checks: {len(ok)} ok, {len(failed)} attention item(s)',
        )

    def _set_status(self, text: str):
        self._status_text = text
        try:
            self.status_left.config(text=text)
        except Exception:
            pass

    # ── Toast notifications ───────────────────────────────────────────────────

    def _toast(self, message: str, kind: str = 'info', duration_ms: int = 2800):
        """Non-blocking, auto-dismissing notification anchored to the bottom-right.

        Use this instead of messagebox.showinfo for confirmations and other
        benign feedback. Stays out of the way; never steals focus."""
        theme = self.T
        accent = {
            'info': theme['accent'],
            'success': '#43A047',
            'warn': '#F9A825',
            'error': '#E53935',
        }.get(kind, theme['accent'])
        try:
            existing = getattr(self, '_toast_win', None)
            if existing is not None and tk.Toplevel.winfo_exists(existing):
                existing.destroy()
        except Exception:
            pass

        top = tk.Toplevel(self.root)
        self._toast_win = top
        top.overrideredirect(True)
        top.attributes('-topmost', True)
        top.configure(bg=accent)

        inner = tk.Frame(top, bg=theme['surface'])
        inner.pack(padx=2, pady=2, fill=tk.BOTH, expand=True)
        tk.Label(inner, text=message, bg=theme['surface'], fg=theme['fg'],
                 font=('Segoe UI', 10), padx=16, pady=10).pack()

        # Anchor to the bottom-right of the main window.
        self.root.update_idletasks()
        rx = self.root.winfo_rootx()
        ry = self.root.winfo_rooty()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        top.update_idletasks()
        tw = top.winfo_width()
        th = top.winfo_height()
        top.geometry(f'+{rx + rw - tw - 28}+{ry + rh - th - 56}')

        def _dismiss(*_args):
            try:
                if top.winfo_exists():
                    top.destroy()
            except tk.TclError:
                pass

        top.bind('<Button-1>', _dismiss)
        top.after(duration_ms, _dismiss)

    def _show_loading_ui(self, visible: bool):
        if visible:
            self.progress.pack(side=tk.RIGHT, before=self.status_right, padx=10, pady=4)
            self.overlay_loading_f.pack(pady=10)
            self.welcome_overlay.place(relx=0.5, rely=0.5, anchor='center', relwidth=1, relheight=1)
        else:
            self.progress.pack_forget()
            self.overlay_loading_f.pack_forget()
            self.progress['value'] = 0
            self.overlay_progress['value'] = 0
            self._update_stats()

    def _draw_brand_art(self):
        if not hasattr(self, 'brand_canvas'):
            return
        theme = self.T
        self.brand_canvas.configure(bg=theme['sidebar'])
        if hasattr(self, 'welcome_logo'):
            self.welcome_logo.configure(bg=theme['sidebar'])

    # ── Crash reports + update check ──────────────────────────────────────────

    def _review_pending_crashes(self):
        reports = pending_crash_reports()
        if not reports:
            return
        latest = reports[-1]
        self._toast(
            f'{len(reports)} crash report(s) from a prior session. '
            'Open Help → Crash Reports to view.',
            kind='warn', duration_ms=5500,
        )
        LOGGER.info('pending-crash-reports count=%s latest=%s', len(reports), latest)

    def _show_crash_reports_dialog(self):
        reports = pending_crash_reports()
        if not reports:
            self._toast('No crash reports — nice and stable.', kind='success')
            return

        theme = self.T
        top = tk.Toplevel(self.root)
        top.title('Crash Reports')
        top.configure(bg=theme['bg'])
        top.transient(self.root)
        top.grab_set()

        f = tk.Frame(top, bg=theme['bg'], padx=18, pady=14)
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text=f'{len(reports)} crash report(s)',
                 bg=theme['bg'], fg=theme['fg'],
                 font=('Segoe UI Semibold', 12)).pack(anchor='w', pady=(0, 8))

        text = tk.Text(f, width=90, height=22, wrap=tk.WORD,
                       bg=theme['surface2'], fg=theme['fg'],
                       relief='flat', font=('Consolas', 9))
        text.pack(fill=tk.BOTH, expand=True)
        try:
            for path in reports[-3:]:
                text.insert(tk.END, f'=== {path.name} ===\n')
                text.insert(tk.END, path.read_text(encoding='utf-8', errors='replace'))
                text.insert(tk.END, '\n\n')
        except Exception as exc:
            text.insert(tk.END, f'Could not read reports: {exc}')
        text.config(state=tk.DISABLED)

        btn_row = tk.Frame(f, bg=theme['bg'])
        btn_row.pack(fill=tk.X, pady=(10, 0))

        def _copy_all():
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(text.get('1.0', tk.END))
                self._toast('Crash reports copied', kind='success')
            except Exception:
                pass

        def _clear():
            removed = clear_crash_reports()
            self._toast(f'Removed {removed} crash report(s).', kind='success')
            top.destroy()

        ttk.Button(btn_row, text='Copy all', command=_copy_all, style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(btn_row, text='Open folder',
                   command=lambda: self._reveal_in_folder(str(reports[-1]))).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row, text='Delete all', command=_clear).pack(side=tk.LEFT)
        ttk.Button(btn_row, text='Close', command=top.destroy).pack(side=tk.RIGHT)

    def _background_update_check(self):
        """Quiet check at startup: just toast if an update is available."""
        if not GITHUB_REPO:
            return
        info = check_for_update(self._APP_VERSION)
        if not info:
            return
        LOGGER.info('update-available current=%s latest=%s',
                    self.VERSION, info.version)
        self.root.after(0, lambda: self._toast(
            f'Update available: v{info.version} — Help → Check for Updates…',
            kind='info', duration_ms=6500))

    def _check_for_updates_interactive(self):
        """Help → Check for Updates… — runs the check on a worker thread and
        opens a dialog with the result. Lets the user download + run the
        installer in-place."""
        if not GITHUB_REPO:
            messagebox.showinfo(
                'Updates disabled',
                'No GitHub repository is configured for update checks.',
            )
            return

        # Modal "checking…" placeholder so the user gets feedback while the
        # API call is in flight.
        theme = self.T
        win = tk.Toplevel(self.root)
        win.title('Check for updates')
        win.transient(self.root)
        win.grab_set()
        win.configure(bg=theme['bg'])
        win.geometry('480x220')
        win.resizable(False, False)

        body = tk.Frame(win, bg=theme['bg'], padx=22, pady=18)
        body.pack(fill=tk.BOTH, expand=True)

        title_lbl = tk.Label(body, text='Checking for updates…',
                             bg=theme['bg'], fg=theme['fg'],
                             font=('Segoe UI Semibold', 12), anchor='w')
        title_lbl.pack(fill=tk.X)
        detail_lbl = tk.Label(body,
                              text=f'Current version: {self.VERSION}\nContacting GitHub…',
                              bg=theme['bg'], fg=theme['fgsub'],
                              font=('Segoe UI', 9), justify='left', anchor='w')
        detail_lbl.pack(fill=tk.X, pady=(6, 12))

        progress = ttk.Progressbar(body, mode='indeterminate', length=420)
        progress.pack(fill=tk.X)
        progress.start(12)

        btn_row = tk.Frame(body, bg=theme['bg'])
        btn_row.pack(fill=tk.X, pady=(14, 0))
        close_btn = ttk.Button(btn_row, text='Close', command=win.destroy)
        close_btn.pack(side=tk.RIGHT)

        def _on_result(info: 'UpdateInfo | None'):
            progress.stop()
            progress.pack_forget()
            if info is None:
                title_lbl.config(text='You’re up to date')
                detail_lbl.config(
                    text=f'You are running the latest version ({self.VERSION}).',
                )
                return

            title_lbl.config(text=f'Update available: v{info.version}')
            size_mb = info.asset_size / (1024 * 1024) if info.asset_size else 0
            detail_lbl.config(
                text=(
                    f'Current: {self.VERSION}     Latest: v{info.version}\n'
                    f'Installer: {info.asset_name}'
                    + (f'  ({size_mb:.1f} MB)' if size_mb else '')
                ),
            )

            def _do_download():
                self._download_and_install(info, win)

            ttk.Button(btn_row, text='Release notes',
                       command=lambda: self._open_url(info.html_url),
                       style='Toolbar.TButton').pack(side=tk.LEFT)
            ttk.Button(btn_row, text='Download & install',
                       command=_do_download,
                       style='Accent.TButton').pack(side=tk.RIGHT, padx=8)

        def _worker():
            info = check_for_update(self._APP_VERSION)
            self.root.after(0, lambda: _on_result(info))

        Thread(target=_worker, daemon=True).start()

    def _download_and_install(self, info: 'UpdateInfo', parent_win: tk.Toplevel):
        """Download the installer .exe with a progress bar, then launch it."""
        theme = self.T
        for child in parent_win.winfo_children():
            child.destroy()

        body = tk.Frame(parent_win, bg=theme['bg'], padx=22, pady=18)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text=f'Downloading v{info.version}…',
                 bg=theme['bg'], fg=theme['fg'],
                 font=('Segoe UI Semibold', 12), anchor='w').pack(fill=tk.X)
        status_lbl = tk.Label(body, text=info.asset_name,
                              bg=theme['bg'], fg=theme['fgsub'],
                              font=('Segoe UI', 9), anchor='w')
        status_lbl.pack(fill=tk.X, pady=(6, 12))

        bar = ttk.Progressbar(body, mode='determinate', maximum=100, length=420)
        bar.pack(fill=tk.X)

        btn_row = tk.Frame(body, bg=theme['bg'])
        btn_row.pack(fill=tk.X, pady=(14, 0))
        cancel_btn = ttk.Button(btn_row, text='Cancel', command=parent_win.destroy)
        cancel_btn.pack(side=tk.RIGHT)

        def _on_progress(done: int, total: int):
            pct = int(done / total * 100) if total else 0
            mb_done = done / (1024 * 1024)
            mb_total = total / (1024 * 1024) if total else 0
            self.root.after(0, lambda: (
                bar.configure(value=pct),
                status_lbl.configure(
                    text=f'{info.asset_name}  —  {mb_done:.1f} / {mb_total:.1f} MB'
                         if mb_total else f'{info.asset_name}  —  {mb_done:.1f} MB',
                ),
            ))

        def _on_done(path: Path):
            bar.configure(value=100)
            status_lbl.configure(text=f'Downloaded to {path}')
            if messagebox.askyesno(
                'Install update',
                f'MailWeave v{info.version} downloaded.\n\n'
                'Launch the installer now? The app will close so it can update.',
                parent=parent_win,
            ):
                try:
                    launch_installer(path)
                except Exception as exc:
                    LOGGER.exception('update-launch-failed')
                    messagebox.showerror(
                        'Could not launch installer',
                        f'The installer was downloaded to:\n{path}\n\nError: {exc}',
                        parent=parent_win,
                    )
                    return
                # Give the installer a moment to start before we exit.
                self.root.after(800, self.root.destroy)
            else:
                parent_win.destroy()

        def _on_error(exc: Exception):
            LOGGER.exception('update-download-failed')
            messagebox.showerror(
                'Download failed',
                f'Could not download the update:\n\n{exc}',
                parent=parent_win,
            )
            parent_win.destroy()

        def _worker():
            try:
                path = download_installer(info, on_progress=_on_progress)
                self.root.after(0, lambda: _on_done(path))
            except Exception as exc:
                self.root.after(0, lambda: _on_error(exc))

        Thread(target=_worker, daemon=True).start()

    def _open_url(self, url: str):
        if not url:
            return
        try:
            import webbrowser
            webbrowser.open(url)
        except Exception:
            LOGGER.exception('open-url-failed url=%s', url)

    def _toggle_theme(self):
        new = 'light' if self.settings.theme == 'dark' else 'dark'
        self.settings.theme = new
        save_settings(self.settings)
        self._apply_theme()
        self._toast(f'Switched to {new} theme.', kind='success')

    def _register_drop_target(self, widget: tk.BaseWidget):
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self._on_drop)
        except Exception:
            self._dnd_ready = False
