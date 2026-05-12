"""MailWeave — animated splash screen."""

import tkinter as tk

from brand_assets import load_logo_photo, set_window_icon


class SplashScreen:
    """Borderless splash window with animated progress bar and logo."""

    W, H = 540, 330

    STAGES = [
        (0.00, 'Loading workspace…'),
        (0.22, 'Preparing email parser…'),
        (0.45, 'Loading Outlook message support…'),
        (0.68, 'Preparing export engines…'),
        (0.86, 'Styling the workspace…'),
        (1.00, 'Ready'),
    ]

    BG = '#FFFFFF'
    BG2 = '#F8FAFC'
    GLOW = '#84cc16'  # Lime Green
    ACCENT = '#65a30d'
    ACCENT2 = '#84cc16'
    FG = '#0F172A'
    FGSUB = '#475569'
    FGDIM = '#94A3B8'

    def __init__(self, parent: tk.Tk):
        self.parent = parent
        self._progress = 0.0
        self._closing = False
        self._pulse = 0

        self.win = tk.Toplevel(parent)
        self.win.overrideredirect(True)
        self.win.resizable(False, False)
        self.win.attributes('-topmost', True)
        set_window_icon(self.win)
        self._center()

        self.canvas = tk.Canvas(
            self.win,
            width=self.W,
            height=self.H,
            bg=self.BG,
            highlightthickness=0,
        )
        self.canvas.pack()
        self._logo_photo = load_logo_photo((220, 160))
        self._draw_static()
        self._pulse_logo()

    def _center(self):
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        x = (sw - self.W) // 2
        y = (sh - self.H) // 2
        self.win.geometry(f'{self.W}x{self.H}+{x}+{y}')

    def _draw_static(self):
        canvas = self.canvas
        width, height = self.W, self.H

        # Clean light background
        canvas.create_rectangle(0, 0, width, height, fill=self.BG, outline='')
        
        # Professional borders
        canvas.create_rectangle(0, 0, width - 1, height - 1, outline=self.FGDIM, width=1)
        canvas.create_rectangle(2, 2, width - 3, height - 3, outline=self.BG2, width=2)

        # Brand text - shifted down to avoid logo overlap
        canvas.create_text(
            width // 2,
            204,
            text='MAILWEAVE',
            font=('Segoe UI Semibold', 32, 'bold'),
            fill=self.FG,
        )
        canvas.create_line(120, 240, width - 120, 240, fill=self.FGDIM, width=1)

        bx, by = 80, 260
        bw, bh = width - 160, 6
        self._bar = (bx, by, bw, bh)
        canvas.create_rectangle(bx, by, bx + bw, by + bh, fill=self.BG2, outline='')
        self._fill = canvas.create_rectangle(bx, by, bx, by + bh, fill=self.ACCENT, outline='')
        
        self._status = canvas.create_text(
            width // 2,
            284,
            text='',
            font=('Segoe UI Semibold', 9),
            fill=self.FGSUB,
        )
        canvas.create_text(width - 16, height - 12, text='v1.0', anchor='se', font=('Segoe UI', 8), fill=self.FGDIM)

    def _draw_logo(self):
        canvas = self.canvas
        canvas.delete('logo')
        pulse = abs(12 - self._pulse)
        outer = 62 + pulse
        inner = 50 + pulse // 2
        cx, cy = self.W // 2, 94 # Shifted up slightly

        canvas.create_oval(cx - outer, cy - outer, cx + outer, cy + outer, fill=self.BG2, outline='', tags='logo')
        canvas.create_oval(cx - inner, cy - inner, cx + inner, cy + inner, fill=self.BG, outline='', tags='logo')
        canvas.create_oval(cx - 56, cy - 56, cx + 56, cy + 56, outline=self.GLOW, width=2, tags='logo')
        canvas.create_arc(cx - 48, cy - 48, cx + 48, cy + 48, start=20, extent=140, style='arc', outline=self.ACCENT, width=4, tags='logo')
        if self._logo_photo is not None:
            canvas.create_image(cx, cy, image=self._logo_photo, tags='logo')

    def _pulse_logo(self):
        if self._closing:
            return
        self._pulse = (self._pulse + 1) % 24
        self._draw_logo()
        self.win.after(110, self._pulse_logo)

    def _set_progress(self, value: float, text: str):
        self._progress = max(0.0, min(1.0, value))
        bx, by, bw, bh = self._bar
        right = bx + int(bw * self._progress)
        self.canvas.coords(self._fill, bx, by, right, by + bh)
        self.canvas.itemconfig(self._status, text=text)
        try:
            self.win.update_idletasks()
        except Exception:
            pass

    def _animate_to(self, target: float, text: str, done_cb, step_ms: int = 16):
        if self._closing:
            return
        diff = target - self._progress
        if abs(diff) < 0.004:
            self._set_progress(target, text)
            done_cb()
            return
        self._set_progress(self._progress + diff * 0.22, text)
        self.win.after(step_ms, lambda: self._animate_to(target, text, done_cb, step_ms))

    def _run_stage(self, index: int, pause_ms: int, done_cb):
        if self._closing or index >= len(self.STAGES):
            self.win.after(300, done_cb)
            return
        target, text = self.STAGES[index]
        self._animate_to(
            target,
            text,
            lambda: self.win.after(pause_ms, lambda: self._run_stage(index + 1, pause_ms, done_cb)),
        )

    def run(self, callback, stage_pause_ms: int = 260):
        self._run_stage(0, stage_pause_ms, callback)

    def close(self):
        self._closing = True
        try:
            self.win.destroy()
        except Exception:
            pass
