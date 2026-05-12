#!/usr/bin/env python3
"""MailWeave — entry point."""

import tkinter as tk

try:
    from tkinterdnd2 import TkinterDnD
    _ROOT_CLS = TkinterDnD.Tk
except ImportError:
    _ROOT_CLS = tk.Tk

from settings import load_settings
from diagnostics import install_exception_hooks, prune_old_logs, run_startup_checks
from splash import SplashScreen
from app import MailWeaveApp


def main():
    install_exception_hooks()
    prune_old_logs()
    # Create root window but hide it until the splash finishes
    root = _ROOT_CLS()
    root.withdraw()

    settings = load_settings()
    startup_checks = run_startup_checks()

    # Build the app UI while the splash is shown (near-instant)
    app = MailWeaveApp(root, settings, startup_checks=startup_checks)

    splash = SplashScreen(root)

    def _on_splash_done():
        splash.close()
        # Centre the main window on the screen before revealing it
        root.update_idletasks()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w  = root.winfo_reqwidth()  or 1300
        h  = root.winfo_reqheight() or 840
        x  = max(0, (sw - w) // 2)
        y  = max(0, (sh - h) // 2)
        root.geometry(f'1300x840+{x}+{y}')
        root.deiconify()
        root.lift()
        root.focus_force()

    splash.run(callback=_on_splash_done, stage_pause_ms=260)
    root.mainloop()


if __name__ == '__main__':
    main()
