"""
AutoMacro - Desktop Game Auto-Clicker
A configurable macro tool for automating click sequences in offline desktop games.

Requirements:
    pip install pyautogui pynput

Usage:
    python autoclicker.py
    - Add click positions manually or capture them with the "Capture" button
    - Set delays between each click and between full loops
    - Press START or use the hotkey (F6) to toggle the macro on/off
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import json
import os
import random
import math

try:
    import pyautogui
    pyautogui.FAILSAFE = True  # Move mouse to top-left corner to emergency stop
except ImportError:
    messagebox.showerror("Missing Dependency", "Please install pyautogui:\n\n  pip install pyautogui")
    exit(1)

try:
    from pynput import keyboard, mouse
except ImportError:
    messagebox.showerror("Missing Dependency", "Please install pynput:\n\n  pip install pynput")
    exit(1)


# ─── Auto-save path (sits next to the script) ───────────────────────────────────
AUTOSAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "autosave.json")

# ─── Color Palette ──────────────────────────────────────────────────────────────
BG        = "#07090f"
SURFACE   = "#0e1117"
SURFACE2  = "#151b26"
SURFACE3  = "#1c2535"
BORDER    = "#253047"
BORDER2   = "#1a2438"
ACCENT    = "#38bdf8"
ACCENT2   = "#34d399"
DANGER    = "#f87171"
WARN      = "#fb923c"
TEXT      = "#e2e8f0"
TEXT_DIM  = "#4b637a"
TEXT_MID  = "#7a94b0"
PILL_BG   = "#0f1f35"

_is_win      = os.name == "nt"
FONT_MONO    = ("Cascadia Code", 10)      if _is_win else ("DejaVu Sans Mono", 10)
FONT_MONO_SM = ("Cascadia Code", 9)       if _is_win else ("DejaVu Sans Mono", 9)
FONT_UI      = ("Segoe UI", 10)           if _is_win else ("Ubuntu", 10)
FONT_UI_SM   = ("Segoe UI", 9)            if _is_win else ("Ubuntu", 9)
FONT_UI_B    = ("Segoe UI", 10, "bold")   if _is_win else ("Ubuntu", 10, "bold")
FONT_BIG     = ("Segoe UI", 13, "bold")   if _is_win else ("Ubuntu", 13, "bold")
FONT_HEAD    = ("Segoe UI", 17, "bold")   if _is_win else ("Ubuntu", 17, "bold")


# ─── Data Model ─────────────────────────────────────────────────────────────────
class ClickStep:
    def __init__(self, x=0, y=0, delay=0.5, button="left", description=""):
        self.x = x
        self.y = y
        self.delay = delay          # seconds BEFORE this click
        self.button = button        # "left" | "right" | "middle"
        self.description = description

    def to_dict(self):
        return {"x": self.x, "y": self.y, "delay": self.delay,
                "button": self.button, "description": self.description}

    @classmethod
    def from_dict(cls, d):
        return cls(**d)


# ─── Main Application ────────────────────────────────────────────────────────────
class AutoMacroApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoMacro")
        self.geometry("840x660")
        self.minsize(740, 560)
        self.configure(bg=BG)
        self.resizable(True, True)

        # State
        self.steps: list[ClickStep] = []
        self.running = False
        self.macro_thread: threading.Thread | None = None
        self.loop_count = 0
        self.capturing = False
        self.capture_listener = None
        self._hk_listener = None
        self._rk_listener = None

        # Settings vars
        self.loop_delay_var    = tk.DoubleVar(value=1.0)
        self.repeat_var        = tk.StringVar(value="∞")   # number or ∞
        self.hotkey_var        = tk.StringVar(value="F6")
        self.record_hotkey_var = tk.StringVar(value="F8")  # bulk-add hotkey
        self.jitter_var        = tk.IntVar(value=0)        # pixel jitter radius
        self.delay_min_var     = tk.DoubleVar(value=0.3)   # random delay range min
        self.delay_max_var     = tk.DoubleVar(value=0.7)   # random delay range max
        self.delay_random_var  = tk.BooleanVar(value=False) # enable random delay
        self.status_var        = tk.StringVar(value="Idle")

        # Auto-save whenever settings change
        self.loop_delay_var.trace_add("write",    lambda *_: self._autosave())
        self.repeat_var.trace_add("write",        lambda *_: self._autosave())
        self.hotkey_var.trace_add("write",        lambda *_: self._autosave())
        self.record_hotkey_var.trace_add("write", lambda *_: self._autosave())
        self.jitter_var.trace_add("write",        lambda *_: self._autosave())
        self.delay_min_var.trace_add("write",     lambda *_: self._autosave())
        self.delay_max_var.trace_add("write",     lambda *_: self._autosave())
        self.delay_random_var.trace_add("write",  lambda *_: self._autosave())

        self._build_ui()
        self._setup_hotkey()
        self._setup_record_hotkey()
        self._update_status_color()
        self._autoload()          # restore last session silently

    # ── UI Construction ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ───────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=0, pady=0)

        inner_hdr = tk.Frame(header, bg=BG)
        inner_hdr.pack(fill="x", padx=20, pady=(14, 12))

        logo_frame = tk.Frame(inner_hdr, bg=BG)
        logo_frame.pack(side="left")
        tk.Label(logo_frame, text="⚡", font=("Segoe UI", 20) if _is_win else ("Ubuntu", 20),
                 bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(logo_frame, text=" AutoMacro", font=FONT_HEAD,
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Label(logo_frame, text="  ·  Desktop Macro Tool", font=FONT_UI_SM,
                 bg=BG, fg=TEXT_DIM).pack(side="left", pady=(3, 0))

        # Status badge (right-aligned)
        badge_frame = tk.Frame(inner_hdr, bg=PILL_BG,
                               highlightbackground=BORDER, highlightthickness=1)
        badge_frame.pack(side="right")
        self.status_dot = tk.Label(badge_frame, text="●", font=FONT_UI_SM,
                                   bg=PILL_BG, fg=TEXT_DIM, padx=6, pady=4)
        self.status_dot.pack(side="left")
        self.status_label = tk.Label(badge_frame, textvariable=self.status_var,
                                     font=FONT_UI_B, bg=PILL_BG,
                                     fg=TEXT_MID, padx=8, pady=4)
        self.status_label.pack(side="left")

        # Thin accent line below header
        tk.Frame(self, bg=BORDER2, height=1).pack(fill="x")

        # ── Main area ────────────────────────────────────────────────────────────
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # ── Left: sequence list ───────────────────────────────────────────────
        left = tk.Frame(main, bg=BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self._section(left, "CLICK SEQUENCE")

        # Treeview card
        list_card = tk.Frame(left, bg=SURFACE, highlightbackground=BORDER,
                             highlightthickness=1)
        list_card.pack(fill="both", expand=True)

        cols = ("step", "x", "y", "button", "delay", "description")
        self.tree = ttk.Treeview(list_card, columns=cols, show="headings",
                                 selectmode="browse")
        self._style_tree()

        hdrs = [("#", 32), ("X", 64), ("Y", 64), ("Btn", 58), ("Delay(s)", 76), ("Note", 180)]
        for col, (label, w) in zip(cols, hdrs):
            self.tree.heading(col, text=label)
            self.tree.column(col, width=w,
                             anchor="center" if label != "Note" else "w",
                             stretch=(label == "Note"))
        self.tree.pack(fill="both", expand=True, side="left")

        sb = ttk.Scrollbar(list_card, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        # Toolbar below sequence
        toolbar = tk.Frame(left, bg=BG, pady=7)
        toolbar.pack(fill="x")

        # Left group: mutating actions
        grp_left = tk.Frame(toolbar, bg=SURFACE2,
                            highlightbackground=BORDER, highlightthickness=1)
        grp_left.pack(side="left")
        self._tbtn(grp_left, "＋ Add",    ACCENT,   self._add_step,        first=True)
        self._tbtn_sep(grp_left)
        self._tbtn(grp_left, "✎ Edit",    WARN,     self._edit_step)
        self._tbtn_sep(grp_left)
        self._tbtn(grp_left, "✕ Remove",  DANGER,   self._remove_step)

        # Middle group: reorder
        grp_mid = tk.Frame(toolbar, bg=SURFACE2,
                           highlightbackground=BORDER, highlightthickness=1)
        grp_mid.pack(side="left", padx=(6, 0))
        self._tbtn(grp_mid, "▲",  TEXT_MID, self._move_up,   first=True)
        self._tbtn_sep(grp_mid)
        self._tbtn(grp_mid, "▼",  TEXT_MID, self._move_down)

        # Right: destructive
        grp_right = tk.Frame(toolbar, bg=SURFACE2,
                             highlightbackground=BORDER, highlightthickness=1)
        grp_right.pack(side="left", padx=(6, 0))
        self._tbtn(grp_right, "🗑 Clear", DANGER, self._clear_all_steps, first=True)

        # ── Right panel ───────────────────────────────────────────────────────
        right_outer = tk.Frame(main, bg=BG, width=256)
        right_outer.pack(side="right", fill="y")
        right_outer.pack_propagate(False)

        # Scrollable right panel
        right_canvas = tk.Canvas(right_outer, bg=BG, highlightthickness=0,
                                 width=256)
        right_canvas.pack(side="left", fill="both", expand=True)
        right_vsb = ttk.Scrollbar(right_outer, orient="vertical",
                                  command=right_canvas.yview)
        right_vsb.pack(side="right", fill="y")
        right_canvas.configure(yscrollcommand=right_vsb.set)

        right = tk.Frame(right_canvas, bg=BG)
        right_win = right_canvas.create_window((0, 0), window=right, anchor="nw")

        def _on_right_configure(e):
            right_canvas.configure(scrollregion=right_canvas.bbox("all"))
            right_canvas.itemconfig(right_win, width=right_canvas.winfo_width())
        right.bind("<Configure>", _on_right_configure)
        right_canvas.bind("<Configure>",
                          lambda e: right_canvas.itemconfig(right_win, width=e.width))

        def _on_mousewheel(e):
            right_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        right_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ── Capture card ─────────────────────────────────────────────────────
        self._section(right, "MOUSE CAPTURE")
        cap_card = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER,
                            highlightthickness=1)
        cap_card.pack(fill="x", pady=(0, 2))

        cap_inner = tk.Frame(cap_card, bg=SURFACE, padx=12, pady=10)
        cap_inner.pack(fill="x")
        tk.Label(cap_inner, text="Click anywhere on screen after\npressing Capture to record XY.",
                 bg=SURFACE, fg=TEXT_MID, font=FONT_UI_SM, justify="left").pack(anchor="w")

        self.cap_btn = self._btn(cap_inner, "🎯  Capture Position",
                                 ACCENT2, self._start_capture)
        self.cap_btn.pack(fill="x", pady=(8, 4))
        self.cap_coords = tk.Label(cap_inner, text="─  waiting ─",
                                   bg=SURFACE, fg=TEXT_DIM, font=FONT_MONO_SM)
        self.cap_coords.pack()

        # ── Settings card ────────────────────────────────────────────────────
        self._section(right, "SETTINGS")
        cfg = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER,
                       highlightthickness=1)
        cfg.pack(fill="x", pady=(0, 2))
        cfg_inner = tk.Frame(cfg, bg=SURFACE, padx=12, pady=10)
        cfg_inner.pack(fill="x")

        self._row(cfg_inner, "Loop delay (s)", self.loop_delay_var, 0, 60, 0.1)
        self._vspace(cfg_inner, 6)
        self._row(cfg_inner, "Jitter radius (px)", self.jitter_var, 0, 50, 1)

        # Divider
        self._vspace(cfg_inner, 10)
        tk.Frame(cfg_inner, bg=BORDER2, height=1).pack(fill="x")
        self._vspace(cfg_inner, 8)

        # Random delay toggle row
        rnd_hdr = tk.Frame(cfg_inner, bg=SURFACE)
        rnd_hdr.pack(fill="x")
        tk.Label(rnd_hdr, text="Randomise step delay",
                 bg=SURFACE, fg=TEXT_MID, font=FONT_UI_SM).pack(side="left")
        self.rnd_toggle = tk.Checkbutton(
            rnd_hdr, variable=self.delay_random_var,
            bg=SURFACE, fg=ACCENT, selectcolor=SURFACE3,
            activebackground=SURFACE, relief="flat", cursor="hand2",
            command=self._update_rnd_state)
        self.rnd_toggle.pack(side="right")

        # Min / Max row
        rnd_row = tk.Frame(cfg_inner, bg=SURFACE)
        rnd_row.pack(fill="x", pady=(5, 0))

        rnd_left = tk.Frame(rnd_row, bg=SURFACE)
        rnd_left.pack(side="left", fill="x", expand=True)
        tk.Label(rnd_left, text="Min (s)", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(anchor="w")
        self.rnd_min_entry = tk.Entry(
            rnd_left, textvariable=self.delay_min_var,
            bg=SURFACE2, fg=ACCENT, insertbackground=TEXT, relief="flat",
            highlightbackground=BORDER, highlightthickness=1,
            font=FONT_MONO_SM, width=6)
        self.rnd_min_entry.pack(anchor="w", pady=(2, 0))

        rnd_right = tk.Frame(rnd_row, bg=SURFACE)
        rnd_right.pack(side="right", fill="x", expand=True)
        tk.Label(rnd_right, text="Max (s)", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(anchor="w")
        self.rnd_max_entry = tk.Entry(
            rnd_right, textvariable=self.delay_max_var,
            bg=SURFACE2, fg=ACCENT, insertbackground=TEXT, relief="flat",
            highlightbackground=BORDER, highlightthickness=1,
            font=FONT_MONO_SM, width=6)
        self.rnd_max_entry.pack(anchor="w", pady=(2, 0))

        tk.Label(cfg_inner, text="Overrides per-step delay when on.",
                 bg=SURFACE, fg=TEXT_DIM, font=FONT_UI_SM).pack(anchor="w", pady=(5, 0))

        # Divider
        self._vspace(cfg_inner, 10)
        tk.Frame(cfg_inner, bg=BORDER2, height=1).pack(fill="x")
        self._vspace(cfg_inner, 8)

        # Repeat count
        rpt_row = tk.Frame(cfg_inner, bg=SURFACE)
        rpt_row.pack(fill="x")
        tk.Label(rpt_row, text="Repeat count", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(side="left")
        tk.Label(rpt_row, text="∞ = forever", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(side="right")
        tk.Entry(cfg_inner, textvariable=self.repeat_var,
                 bg=SURFACE2, fg=TEXT, insertbackground=TEXT, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 font=FONT_MONO, width=10).pack(fill="x", pady=(4, 0))

        # Divider
        self._vspace(cfg_inner, 10)
        tk.Frame(cfg_inner, bg=BORDER2, height=1).pack(fill="x")
        self._vspace(cfg_inner, 8)

        # Hotkeys side-by-side
        hk_row = tk.Frame(cfg_inner, bg=SURFACE)
        hk_row.pack(fill="x")

        hk_left = tk.Frame(hk_row, bg=SURFACE)
        hk_left.pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Label(hk_left, text="Toggle key", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(anchor="w")
        hk_entry_row = tk.Frame(hk_left, bg=SURFACE)
        hk_entry_row.pack(fill="x", pady=(3, 0))
        tk.Entry(hk_entry_row, textvariable=self.hotkey_var,
                 bg=SURFACE2, fg=ACCENT, insertbackground=TEXT, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 font=FONT_MONO_SM, width=5).pack(side="left")
        self._btn(hk_entry_row, "Set", ACCENT, self._setup_hotkey,
                  small=True).pack(side="left", padx=(4, 0))

        hk_right = tk.Frame(hk_row, bg=SURFACE)
        hk_right.pack(side="right", fill="x", expand=True, padx=(4, 0))
        tk.Label(hk_right, text="Record key", bg=SURFACE, fg=TEXT_DIM,
                 font=FONT_UI_SM).pack(anchor="w")
        rk_entry_row = tk.Frame(hk_right, bg=SURFACE)
        rk_entry_row.pack(fill="x", pady=(3, 0))
        tk.Entry(rk_entry_row, textvariable=self.record_hotkey_var,
                 bg=SURFACE2, fg=WARN, insertbackground=TEXT, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 font=FONT_MONO_SM, width=5).pack(side="left")
        self._btn(rk_entry_row, "Set", WARN, self._setup_record_hotkey,
                  small=True).pack(side="left", padx=(4, 0))

        # ── Control card ─────────────────────────────────────────────────────
        self._section(right, "CONTROL")
        ctrl_card = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER,
                             highlightthickness=1)
        ctrl_card.pack(fill="x", pady=(0, 2))
        ctrl_inner = tk.Frame(ctrl_card, bg=SURFACE, padx=12, pady=10)
        ctrl_inner.pack(fill="x")

        self.run_btn = tk.Button(
            ctrl_inner, text="▶   START", bg=ACCENT2, fg=BG,
            activebackground="#2cb885", activeforeground=BG,
            font=FONT_BIG, relief="flat", cursor="hand2",
            padx=10, pady=10, command=self._toggle)
        self.run_btn.pack(fill="x", pady=(0, 8))

        io_row = tk.Frame(ctrl_inner, bg=SURFACE)
        io_row.pack(fill="x")
        self._btn(io_row, "💾 Save", TEXT_MID, self._save_sequence).pack(
            side="left", fill="x", expand=True, padx=(0, 4))
        self._btn(io_row, "📂 Load", TEXT_MID, self._load_sequence).pack(
            side="right", fill="x", expand=True)

        # ── Footer bar ───────────────────────────────────────────────────────
        footer = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER2,
                          highlightthickness=1)
        footer.pack(fill="x", side="bottom")
        footer_inner = tk.Frame(footer, bg=SURFACE)
        footer_inner.pack(fill="x", padx=16, pady=5)

        self.loop_label = tk.Label(footer_inner, text="Loops: 0",
                                   bg=SURFACE, fg=TEXT_MID, font=FONT_UI_SM)
        self.loop_label.pack(side="left")

        self.step_label = tk.Label(footer_inner, text="Steps: 0",
                                   bg=SURFACE, fg=TEXT_MID, font=FONT_UI_SM)
        self.step_label.pack(side="left", padx=(16, 0))

        tk.Label(footer_inner,
                 text="⚠  Move mouse to top-left corner to emergency stop",
                 bg=SURFACE, fg=TEXT_DIM, font=FONT_UI_SM).pack(side="right")

        self._update_rnd_state()

    def _section(self, parent, title):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(10, 4))
        tk.Label(f, text=title, font=(FONT_UI_SM[0], 8, "bold"),
                 bg=BG, fg=TEXT_DIM).pack(side="left")
        tk.Frame(f, bg=BORDER, height=1).pack(
            side="left", fill="x", expand=True, padx=(8, 0), pady=5)

    def _vspace(self, parent, h):
        tk.Frame(parent, bg=SURFACE, height=h).pack(fill="x")

    def _tbtn(self, parent, text, color, cmd, first=False):
        """Toolbar button — flat, in a button-group container."""
        b = tk.Button(parent, text=text, bg=SURFACE2, fg=color,
                      activebackground=SURFACE3, activeforeground=color,
                      font=FONT_UI_SM, relief="flat", cursor="hand2",
                      padx=10, pady=5, bd=0, command=cmd)
        b.pack(side="left")
        b.bind("<Enter>", lambda e: b.config(bg=SURFACE3))
        b.bind("<Leave>", lambda e: b.config(bg=SURFACE2))
        return b

    def _tbtn_sep(self, parent):
        tk.Frame(parent, bg=BORDER, width=1).pack(side="left", fill="y", pady=4)

    def _btn(self, parent, text, color, cmd, small=False, big=False):
        size = FONT_UI_SM if small else (FONT_BIG if big else FONT_UI)
        b = tk.Button(parent, text=text, bg=SURFACE2, fg=color,
                      activebackground=SURFACE3, activeforeground=color,
                      font=size, relief="flat", cursor="hand2",
                      highlightbackground=BORDER, highlightthickness=1,
                      padx=8, pady=5 if not big else 8, command=cmd)
        b.bind("<Enter>", lambda e: b.config(bg=SURFACE3))
        b.bind("<Leave>", lambda e: b.config(bg=SURFACE2))
        return b

    def _row(self, parent, label, var, from_, to, res):
        tk.Label(parent, text=label, bg=SURFACE, fg=TEXT_MID,
                 font=FONT_UI_SM).pack(anchor="w")
        f = tk.Frame(parent, bg=SURFACE)
        f.pack(fill="x", pady=(2, 4))
        tk.Scale(f, variable=var, from_=from_, to=to, resolution=res,
                 orient="horizontal",
                 bg=SURFACE,
                 fg=TEXT,
                 troughcolor="#1e3a5f",
                 activebackground="#7dd3fc",
                 highlightbackground=SURFACE,
                 highlightthickness=0,
                 sliderrelief="flat",
                 sliderlength=18,
                 width=10,             # taller trough = easier to grab
                 bd=0,
                 font=FONT_UI_SM).pack(side="left", fill="x", expand=True)
        tk.Label(f, textvariable=var, bg=SURFACE2, fg=ACCENT,
                 font=FONT_MONO_SM, width=5, padx=4, pady=2,
                 highlightbackground=BORDER, highlightthickness=1).pack(side="right", padx=(6, 0))

    def _style_tree(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=SURFACE, fieldbackground=SURFACE,
                        foreground=TEXT, rowheight=28, font=FONT_MONO_SM,
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background=SURFACE2, foreground=TEXT_MID,
                        font=(FONT_UI_SM[0], 9, "bold"),
                        relief="flat", padding=(4, 7))
        style.map("Treeview",
                  background=[("selected", SURFACE3)],
                  foreground=[("selected", ACCENT)])
        style.configure("Vertical.TScrollbar",
                        background=SURFACE2, troughcolor=SURFACE,
                        arrowcolor=TEXT_DIM, borderwidth=0)

    def _update_rnd_state(self):
        """Dim the min/max entries when random delay is disabled."""
        state = "normal" if self.delay_random_var.get() else "disabled"
        fg    = ACCENT   if self.delay_random_var.get() else TEXT_DIM
        self.rnd_min_entry.config(state=state, fg=fg)
        self.rnd_max_entry.config(state=state, fg=fg)

    # ── Sequence Management ──────────────────────────────────────────────────────
    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, s in enumerate(self.steps, 1):
            self.tree.insert("", "end", iid=str(i-1),
                             values=(i, s.x, s.y, s.button.title(), f"{s.delay:.2f}", s.description))
        if hasattr(self, "step_label"):
            self.step_label.config(text=f"Steps: {len(self.steps)}")
        self._autosave()  # persist every change automatically

    def _add_step(self):
        self._step_dialog()

    def _edit_step(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select a step", "Please select a step to edit.")
            return
        idx = int(sel[0])
        self._step_dialog(idx)

    def _remove_step(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        self.steps.pop(idx)
        self._refresh_tree()

    def _move_up(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx > 0:
            self.steps[idx], self.steps[idx-1] = self.steps[idx-1], self.steps[idx]
            self._refresh_tree()
            self.tree.selection_set(str(idx-1))

    def _move_down(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx < len(self.steps) - 1:
            self.steps[idx], self.steps[idx+1] = self.steps[idx+1], self.steps[idx]
            self._refresh_tree()
            self.tree.selection_set(str(idx+1))

    def _clear_all_steps(self):
        if not self.steps:
            return
        if messagebox.askyesno("Clear All Steps",
                               f"Remove all {len(self.steps)} step(s)?\nThis cannot be undone."):
            self.steps.clear()
            self._refresh_tree()

    # ── Step Dialog ──────────────────────────────────────────────────────────────
    def _step_dialog(self, edit_idx=None):
        step = self.steps[edit_idx] if edit_idx is not None else ClickStep()
        win = tk.Toplevel(self)
        win.title("Edit Step" if edit_idx is not None else "Add Step")
        win.geometry("420x330")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        win.transient(self)

        fields = {}

        def lbl_entry(parent, label, default, row):
            tk.Label(parent, text=label, bg=BG, fg=TEXT_MID, font=FONT_UI_SM).grid(
                row=row, column=0, sticky="w", padx=16, pady=5)
            var = tk.StringVar(value=str(default))
            e = tk.Entry(parent, textvariable=var, bg=SURFACE2, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         highlightbackground=BORDER, highlightthickness=1,
                         font=FONT_MONO_SM, width=18)
            e.grid(row=row, column=1, padx=(0, 16), pady=5)
            return var

        fields["x"]     = lbl_entry(win, "X coordinate:", step.x, 0)
        fields["y"]     = lbl_entry(win, "Y coordinate:", step.y, 1)
        fields["delay"] = lbl_entry(win, "Delay before (s):", step.delay, 2)

        tk.Label(win, text="Mouse button:", bg=BG, fg=TEXT_MID,
                 font=FONT_UI_SM).grid(row=3, column=0, sticky="w", padx=16, pady=5)
        btn_var = tk.StringVar(value=step.button)
        btn_frame = tk.Frame(win, bg=BG)
        btn_frame.grid(row=3, column=1, sticky="w", pady=5)
        for b in ("left", "right", "middle"):
            tk.Radiobutton(btn_frame, text=b.title(), variable=btn_var, value=b,
                           bg=BG, fg=TEXT, selectcolor=SURFACE3,
                           activebackground=BG, font=FONT_UI_SM).pack(side="left", padx=3)

        fields["description"] = lbl_entry(win, "Note (optional):", step.description, 4)

        def fill_from_capture():
            try:
                x, y = pyautogui.position()
                fields["x"].set(str(x))
                fields["y"].set(str(y))
            except Exception:
                pass

        pos_btn = tk.Button(win, text="📍  Use Current Mouse Position",
                            bg=SURFACE2, fg=ACCENT, relief="flat",
                            font=FONT_UI_SM, cursor="hand2",
                            highlightbackground=BORDER, highlightthickness=1,
                            command=fill_from_capture)
        pos_btn.grid(row=5, column=0, columnspan=2, pady=8, padx=16, sticky="ew")
        pos_btn.bind("<Enter>", lambda e: pos_btn.config(bg=SURFACE3))
        pos_btn.bind("<Leave>", lambda e: pos_btn.config(bg=SURFACE2))

        def save():
            try:
                s = ClickStep(
                    x=int(fields["x"].get()),
                    y=int(fields["y"].get()),
                    delay=float(fields["delay"].get()),
                    button=btn_var.get(),
                    description=fields["description"].get()
                )
            except ValueError:
                messagebox.showerror("Invalid Input",
                                     "X, Y must be integers and Delay must be a number.",
                                     parent=win)
                return
            if edit_idx is not None:
                self.steps[edit_idx] = s
            else:
                self.steps.append(s)
            self._refresh_tree()
            win.destroy()

        save_btn = tk.Button(win, text="✔  Save Step", bg=ACCENT, fg=BG,
                             relief="flat", font=FONT_UI_B, cursor="hand2",
                             padx=12, pady=8, command=save)
        save_btn.grid(row=6, column=0, columnspan=2, pady=(4, 12), padx=16, sticky="ew")
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#5ccfff"))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg=ACCENT))

    # ── Capture ──────────────────────────────────────────────────────────────────
    def _start_capture(self):
        if self.capturing:
            return
        self.capturing = True
        self.cap_btn.config(text="📍  Click anywhere...", fg=WARN)
        self.cap_coords.config(text="waiting for click...", fg=TEXT_DIM)

        def on_click(x, y, button, pressed):
            if pressed:
                self.after(0, lambda: self._on_captured(int(x), int(y)))
                return False  # stop listener

        self.capture_listener = mouse.Listener(on_click=on_click)
        self.capture_listener.start()

    def _on_captured(self, x, y):
        self.capturing = False
        self.cap_btn.config(text="🎯  Capture Position", fg=ACCENT2)
        self.cap_coords.config(text=f"X: {x}   Y: {y}", fg=ACCENT)
        if messagebox.askyesno("Add Captured Position",
                               f"Captured ({x}, {y}).\nAdd as a new step?"):
            self.steps.append(ClickStep(x=x, y=y))
            self._refresh_tree()

    # ── Hotkey ───────────────────────────────────────────────────────────────────
    KEY_MAP = {
        **{f"F{i}": f"<f{i}>" for i in range(1, 13)},
        "ESCAPE": "<esc>", "ESC": "<esc>",
        "TAB": "<tab>", "SPACE": "<space>",
        "HOME": "<home>", "END": "<end>",
        "INSERT": "<insert>", "DELETE": "<delete>",
        "PAUSE": "<pause>",
    }

    def _stop_listener(self, attr):
        listener = getattr(self, attr, None)
        if listener is not None:
            try:
                listener.stop()
            except Exception:
                pass
            setattr(self, attr, None)

    def _setup_hotkey(self):
        self._stop_listener("_hk_listener")
        hk = self.hotkey_var.get().strip()
        key_str = self.KEY_MAP.get(hk.upper())
        if not key_str:
            messagebox.showwarning("Hotkey",
                f"'{hk}' is not recognised.\nUse F1–F12 or: ESC, TAB, SPACE, HOME, END, INSERT, DELETE, PAUSE.")
            return
        try:
            self._hk_listener = keyboard.GlobalHotKeys({key_str: lambda: self.after(0, self._toggle)})
            self._hk_listener.daemon = True
            self._hk_listener.start()
        except Exception as e:
            messagebox.showwarning("Hotkey", f"Could not bind hotkey '{hk}'.\n{e}")

    def _setup_record_hotkey(self):
        """Bind the bulk-add hotkey: press it anytime to snapshot current mouse position as a new step."""
        self._stop_listener("_rk_listener")
        hk = self.record_hotkey_var.get().strip()
        key_str = self.KEY_MAP.get(hk.upper())
        if not key_str:
            messagebox.showwarning("Record Hotkey",
                f"'{hk}' is not recognised.\nUse F1–F12 or: ESC, TAB, SPACE, HOME, END, INSERT, DELETE, PAUSE.")
            return
        try:
            self._rk_listener = keyboard.GlobalHotKeys({key_str: lambda: self.after(0, self._record_step)})
            self._rk_listener.daemon = True
            self._rk_listener.start()
        except Exception as e:
            messagebox.showwarning("Record Hotkey", f"Could not bind record hotkey '{hk}'.\n{e}")

    def _record_step(self):
        """Instantly add the current mouse position as a new step (no dialog)."""
        try:
            x, y = pyautogui.position()
        except Exception:
            return
        step = ClickStep(x=x, y=y, delay=0.5, button="left", description="recorded")
        self.steps.append(step)
        self._refresh_tree()
        # Select the new row and scroll to it
        new_iid = str(len(self.steps) - 1)
        self.tree.selection_set(new_iid)
        self.tree.see(new_iid)
        self._flash_record(x, y)

    def _flash_record(self, x, y):
        """Briefly highlight the status label to confirm a step was recorded."""
        orig_text  = self.status_var.get()
        orig_color = self.status_label.cget("fg")
        self.status_var.set(f"✚ Recorded ({x}, {y})")
        self.status_label.config(fg=WARN)
        self.after(900, lambda: (
            self.status_var.set(orig_text),
            self.status_label.config(fg=orig_color)
        ))

    # ── Macro Runner ─────────────────────────────────────────────────────────────
    def _toggle(self):
        if self.running:
            self._stop()
        else:
            self._start()

    def _start(self):
        if not self.steps:
            messagebox.showwarning("Empty Sequence", "Add at least one click step before starting.")
            return
        self.running = True
        self.loop_count = 0
        self.status_var.set("Running")
        self.run_btn.config(text="■   STOP", bg=DANGER, fg=TEXT,
                            activebackground="#c0392b")
        self._update_status_color()
        self.macro_thread = threading.Thread(target=self._run_loop, daemon=True)
        self.macro_thread.start()

    def _stop(self):
        self.running = False
        self.status_var.set("Idle")
        self.run_btn.config(text="▶   START", bg=ACCENT2, fg=BG,
                            activebackground="#2cb885")
        self._update_status_color()

    def _run_loop(self):
        repeat_str = self.repeat_var.get().strip()
        max_loops = None
        if repeat_str not in ("∞", "inf", ""):
            try:
                max_loops = int(repeat_str)
            except ValueError:
                max_loops = None

        while self.running:
            for step in self.steps:
                if not self.running:
                    break
                if self.delay_random_var.get():
                    lo = self.delay_min_var.get()
                    hi = self.delay_max_var.get()
                    if lo > hi:
                        lo, hi = hi, lo
                    actual_delay = random.uniform(lo, hi)
                else:
                    actual_delay = step.delay
                time.sleep(max(0, actual_delay))
                if not self.running:
                    break
                try:
                    jitter = self.jitter_var.get()
                    if jitter > 0:
                        angle  = random.uniform(0, 2 * math.pi)
                        radius = random.uniform(0, jitter)
                        cx = int(step.x + radius * math.cos(angle))
                        cy = int(step.y + radius * math.sin(angle))
                    else:
                        cx, cy = step.x, step.y
                    pyautogui.click(cx, cy, button=step.button)
                except pyautogui.FailSafeException:
                    self.after(0, self._emergency_stop)
                    return

            if not self.running:
                break

            self.loop_count += 1
            self.after(0, lambda n=self.loop_count: self.loop_label.config(text=f"Loops: {n}"))

            if max_loops is not None and self.loop_count >= max_loops:
                break

            ld = self.loop_delay_var.get()
            deadline = time.time() + ld
            while self.running and time.time() < deadline:
                time.sleep(0.05)

        self.after(0, self._stop)

    def _emergency_stop(self):
        self._stop()
        messagebox.showwarning("Emergency Stop", "Macro stopped: mouse moved to corner (failsafe triggered).")

    def _update_status_color(self):
        if self.running:
            self.status_dot.config(fg=ACCENT2)
            self.status_label.config(fg=ACCENT2)
        else:
            self.status_dot.config(fg=TEXT_DIM)
            self.status_label.config(fg=TEXT_MID)

    # ── Persistence ──────────────────────────────────────────────────────────────
    def _save_sequence(self):
        from tkinter.filedialog import asksaveasfilename
        path = asksaveasfilename(defaultextension=".json",
                                 filetypes=[("JSON", "*.json"), ("All", "*.*")],
                                 title="Save Sequence")
        if not path:
            return
        data = {
            "loop_delay": self.loop_delay_var.get(),
            "repeat": self.repeat_var.get(),
            "jitter": self.jitter_var.get(),
            "delay_random": self.delay_random_var.get(),
            "delay_min": self.delay_min_var.get(),
            "delay_max": self.delay_max_var.get(),
            "steps": [s.to_dict() for s in self.steps]
        }
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        messagebox.showinfo("Saved", f"Sequence saved to:\n{path}")

    def _load_sequence(self):
        from tkinter.filedialog import askopenfilename
        path = askopenfilename(filetypes=[("JSON", "*.json"), ("All", "*.*")],
                               title="Load Sequence")
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
            self.loop_delay_var.set(data.get("loop_delay", 1.0))
            self.repeat_var.set(data.get("repeat", "∞"))
            self.jitter_var.set(data.get("jitter", 0))
            self.delay_random_var.set(data.get("delay_random", False))
            self.delay_min_var.set(data.get("delay_min", 0.3))
            self.delay_max_var.set(data.get("delay_max", 0.7))
            self.steps = [ClickStep.from_dict(d) for d in data.get("steps", [])]
            self._refresh_tree()
            messagebox.showinfo("Loaded", f"Loaded {len(self.steps)} steps from file.")
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load file:\n{e}")

    # ── Auto-save / Auto-load ─────────────────────────────────────────────────────
    def _autosave(self):
        """Silently save current sequence and settings beside the script."""
        try:
            data = {
                "loop_delay":    self.loop_delay_var.get(),
                "repeat":        self.repeat_var.get(),
                "hotkey":        self.hotkey_var.get(),
                "record_hotkey": self.record_hotkey_var.get(),
                "jitter":        self.jitter_var.get(),
                "delay_random":  self.delay_random_var.get(),
                "delay_min":     self.delay_min_var.get(),
                "delay_max":     self.delay_max_var.get(),
                "steps":         [s.to_dict() for s in self.steps],
            }
            with open(AUTOSAVE_PATH, "w") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _autoload(self):
        """Silently restore the last session from autosave on startup."""
        if not os.path.exists(AUTOSAVE_PATH):
            return
        try:
            with open(AUTOSAVE_PATH) as f:
                data = json.load(f)
            self.loop_delay_var.set(data.get("loop_delay", 1.0))
            self.repeat_var.set(data.get("repeat", "∞"))
            self.hotkey_var.set(data.get("hotkey", "F6"))
            self.record_hotkey_var.set(data.get("record_hotkey", "F8"))
            self.jitter_var.set(data.get("jitter", 0))
            self.delay_random_var.set(data.get("delay_random", False))
            self.delay_min_var.set(data.get("delay_min", 0.3))
            self.delay_max_var.set(data.get("delay_max", 0.7))
            self.steps = [ClickStep.from_dict(d) for d in data.get("steps", [])]
            self._refresh_tree()
        except Exception:
            pass  # corrupted autosave — start fresh
        # Always re-setup hotkeys regardless of whether autosave loaded cleanly
        self._setup_hotkey()
        self._setup_record_hotkey()

    def on_close(self):
        self.running = False
        for attr in ("_hk_listener", "_rk_listener"):
            listener = getattr(self, attr, None)
            if listener:
                try:
                    listener.stop()
                except Exception:
                    pass
        self.destroy()


# ─── Entry Point ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AutoMacroApp()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()