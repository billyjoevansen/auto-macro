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
BG        = "#0d1117"
SURFACE   = "#161b22"
SURFACE2  = "#21262d"
BORDER    = "#30363d"
ACCENT    = "#58a6ff"
ACCENT2   = "#3fb950"
DANGER    = "#f85149"
WARN      = "#d29922"
TEXT      = "#e6edf3"
TEXT_DIM  = "#8b949e"
FONT_MONO = ("Cascadia Code", 10) if os.name == "nt" else ("DejaVu Sans Mono", 10)
FONT_UI   = ("Segoe UI", 10) if os.name == "nt" else ("Ubuntu", 10)
FONT_BIG  = ("Segoe UI", 13, "bold") if os.name == "nt" else ("Ubuntu", 13, "bold")


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
        self.geometry("780x640")
        self.minsize(720, 560)
        self.configure(bg=BG)
        self.resizable(True, True)

        # State
        self.steps: list[ClickStep] = []
        self.running = False
        self.macro_thread: threading.Thread | None = None
        self.loop_count = 0
        self.capturing = False
        self.capture_listener = None

        # Settings vars
        self.loop_delay_var    = tk.DoubleVar(value=1.0)
        self.repeat_var        = tk.StringVar(value="∞")   # number or ∞
        self.hotkey_var        = tk.StringVar(value="F6")
        self.record_hotkey_var = tk.StringVar(value="F8")  # bulk-add hotkey
        self.status_var        = tk.StringVar(value="Idle")

        # Auto-save whenever settings change
        self.loop_delay_var.trace_add("write",    lambda *_: self._autosave())
        self.repeat_var.trace_add("write",        lambda *_: self._autosave())
        self.hotkey_var.trace_add("write",        lambda *_: self._autosave())
        self.record_hotkey_var.trace_add("write", lambda *_: self._autosave())

        self._build_ui()
        self._setup_hotkey()
        self._setup_record_hotkey()
        self._update_status_color()
        self._autoload()          # restore last session silently

    # ── UI Construction ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # Title bar
        header = tk.Frame(self, bg=BG, pady=12)
        header.pack(fill="x", padx=20)
        tk.Label(header, text="⚡ AutoMacro", font=("Segoe UI", 18, "bold") if os.name == "nt" else ("Ubuntu", 18, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(header, text="Desktop Game Macro Tool", font=FONT_UI,
                 bg=BG, fg=TEXT_DIM).pack(side="left", padx=(10, 0), pady=(4, 0))

        # Status badge
        self.status_label = tk.Label(header, textvariable=self.status_var,
                                     font=(FONT_UI[0], 10, "bold"), bg=SURFACE2,
                                     fg=TEXT_DIM, padx=12, pady=4, relief="flat")
        self.status_label.pack(side="right")

        # Main paned area
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Left: sequence list
        left = tk.Frame(main, bg=BG)
        left.pack(side="left", fill="both", expand=True)

        self._section(left, "CLICK SEQUENCE")

        list_frame = tk.Frame(left, bg=SURFACE, highlightbackground=BORDER,
                              highlightthickness=1)
        list_frame.pack(fill="both", expand=True)

        cols = ("step", "x", "y", "button", "delay", "description")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                                 selectmode="browse")
        self._style_tree()

        hdrs = [("#", 36), ("X", 60), ("Y", 60), ("Btn", 56), ("Delay(s)", 72), ("Note", 200)]
        for col, (label, w) in zip(cols, hdrs):
            self.tree.heading(col, text=label)
            self.tree.column(col, width=w, anchor="center" if label != "Note" else "w",
                             stretch=(label == "Note"))
        self.tree.pack(fill="both", expand=True, side="left")

        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        # Sequence buttons
        seq_btns = tk.Frame(left, bg=BG, pady=6)
        seq_btns.pack(fill="x")
        self._btn(seq_btns, "＋ Add Step",    ACCENT,  self._add_step).pack(side="left", padx=(0,4))
        self._btn(seq_btns, "✎ Edit",         WARN,    self._edit_step).pack(side="left", padx=4)
        self._btn(seq_btns, "✕ Remove",       DANGER,  self._remove_step).pack(side="left", padx=4)
        self._btn(seq_btns, "▲ Up",           TEXT_DIM, self._move_up).pack(side="left", padx=4)
        self._btn(seq_btns, "▼ Down",         TEXT_DIM, self._move_down).pack(side="left", padx=4)
        self._btn(seq_btns, "🗑 Clear All",   DANGER,   self._clear_all_steps).pack(side="left", padx=4)

        # Right panel
        right = tk.Frame(main, bg=BG, width=200)
        right.pack(side="right", fill="y", padx=(14, 0))
        right.pack_propagate(False)

        # Capture
        self._section(right, "MOUSE CAPTURE")
        cap_box = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER,
                           highlightthickness=1, pady=10, padx=10)
        cap_box.pack(fill="x")
        tk.Label(cap_box, text="Click anywhere to record\nposition after pressing Capture.",
                 bg=SURFACE, fg=TEXT_DIM, font=FONT_UI, justify="left").pack(anchor="w")
        self.cap_btn = self._btn(cap_box, "🎯 Capture Position", ACCENT2, self._start_capture)
        self.cap_btn.pack(fill="x", pady=(8, 0))
        self.cap_coords = tk.Label(cap_box, text="─ waiting ─", bg=SURFACE,
                                   fg=TEXT_DIM, font=FONT_MONO)
        self.cap_coords.pack(pady=(6, 0))

        # Macro settings
        self._section(right, "MACRO SETTINGS")
        cfg = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER,
                       highlightthickness=1, padx=10, pady=10)
        cfg.pack(fill="x")

        self._row(cfg, "Loop delay (s):", self.loop_delay_var, 0, 60, 0.1)
        tk.Label(cfg, text="Repeat count (∞ = forever):", bg=SURFACE,
                 fg=TEXT_DIM, font=FONT_UI).pack(anchor="w", pady=(8, 2))
        repeat_entry = tk.Entry(cfg, textvariable=self.repeat_var, bg=SURFACE2,
                                fg=TEXT, insertbackground=TEXT, relief="flat",
                                highlightbackground=BORDER, highlightthickness=1,
                                font=FONT_MONO, width=8)
        repeat_entry.pack(anchor="w")

        tk.Label(cfg, text="Toggle hotkey:", bg=SURFACE,
                 fg=TEXT_DIM, font=FONT_UI).pack(anchor="w", pady=(8, 2))
        hk_frame = tk.Frame(cfg, bg=SURFACE)
        hk_frame.pack(fill="x")
        hk_entry = tk.Entry(hk_frame, textvariable=self.hotkey_var, bg=SURFACE2,
                            fg=ACCENT, insertbackground=TEXT, relief="flat",
                            highlightbackground=BORDER, highlightthickness=1,
                            font=FONT_MONO, width=6)
        hk_entry.pack(side="left")
        self._btn(hk_frame, "Set", ACCENT, self._setup_hotkey, small=True).pack(side="left", padx=(6,0))

        tk.Label(cfg, text="Record step hotkey:", bg=SURFACE,
                 fg=TEXT_DIM, font=FONT_UI).pack(anchor="w", pady=(8, 2))
        rk_frame = tk.Frame(cfg, bg=SURFACE)
        rk_frame.pack(fill="x")
        rk_entry = tk.Entry(rk_frame, textvariable=self.record_hotkey_var, bg=SURFACE2,
                            fg=WARN, insertbackground=TEXT, relief="flat",
                            highlightbackground=BORDER, highlightthickness=1,
                            font=FONT_MONO, width=6)
        rk_entry.pack(side="left")
        self._btn(rk_frame, "Set", WARN, self._setup_record_hotkey, small=True).pack(side="left", padx=(6,0))
        tk.Label(cfg, text="Press to snap current\nmouse pos as a new step",
                 bg=SURFACE, fg=TEXT_DIM, font=(FONT_UI[0], 8)).pack(anchor="w", pady=(2, 0))

        # Run / stop controls
        self._section(right, "CONTROL")
        ctrl = tk.Frame(right, bg=BG)
        ctrl.pack(fill="x")
        self.run_btn = self._btn(ctrl, "▶  START", ACCENT2, self._toggle, big=True)
        self.run_btn.pack(fill="x", pady=(0, 6))
        self._btn(ctrl, "💾 Save Sequence", TEXT_DIM, self._save_sequence).pack(fill="x", pady=(0,4))
        self._btn(ctrl, "📂 Load Sequence", TEXT_DIM, self._load_sequence).pack(fill="x")

        # Stats bar
        stats = tk.Frame(self, bg=SURFACE2, pady=6)
        stats.pack(fill="x", side="bottom")
        self.loop_label = tk.Label(stats, text="Loops: 0", bg=SURFACE2,
                                   fg=TEXT_DIM, font=FONT_UI, padx=14)
        self.loop_label.pack(side="left")
        tk.Label(stats, text="Emergency stop: move mouse to top-left corner of screen",
                 bg=SURFACE2, fg=TEXT_DIM, font=FONT_UI).pack(side="right", padx=14)

    def _section(self, parent, title):
        tk.Frame(parent, bg=BG, height=10).pack(fill="x")   # top spacing
        f = tk.Frame(parent, bg=BG)
        f.pack_configure(pady=(10, 4))
        f.pack(fill="x")
        tk.Label(f, text=title, font=(FONT_UI[0], 8, "bold"), bg=BG,
                 fg=TEXT_DIM).pack(side="left")
        tk.Frame(f, bg=BORDER, height=1).pack(side="left", fill="x", expand=True, padx=(8, 0), pady=5)

    def _btn(self, parent, text, color, cmd, small=False, big=False):
        size = (FONT_UI[0], 8) if small else ((FONT_BIG[0], 12, "bold") if big else FONT_UI)
        b = tk.Button(parent, text=text, bg=SURFACE2, fg=color,
                      activebackground=BORDER, activeforeground=color,
                      font=size, relief="flat", cursor="hand2",
                      highlightbackground=BORDER, highlightthickness=1,
                      padx=10, pady=6 if big else 5,
                      command=cmd)
        return b

    def _row(self, parent, label, var, from_, to, res):
        tk.Label(parent, text=label, bg=SURFACE, fg=TEXT_DIM, font=FONT_UI).pack(anchor="w")
        f = tk.Frame(parent, bg=SURFACE)
        f.pack(fill="x", pady=(2, 0))
        tk.Scale(f, variable=var, from_=from_, to=to, resolution=res,
                 orient="horizontal", bg=SURFACE, fg=TEXT, troughcolor=SURFACE2,
                 highlightthickness=0, sliderrelief="flat", bd=0,
                 activebackground=ACCENT, font=FONT_UI).pack(side="left", fill="x", expand=True)
        tk.Label(f, textvariable=var, bg=SURFACE, fg=ACCENT,
                 font=FONT_MONO, width=5).pack(side="right")

    def _style_tree(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                        foreground=TEXT, rowheight=26, font=FONT_MONO,
                        borderwidth=0)
        style.configure("Treeview.Heading", background=SURFACE2, foreground=TEXT_DIM,
                        font=(FONT_UI[0], 9, "bold"), relief="flat", padding=(4, 6))
        style.map("Treeview", background=[("selected", SURFACE2)],
                  foreground=[("selected", ACCENT)])
        style.configure("Vertical.TScrollbar", background=SURFACE2,
                        troughcolor=SURFACE, arrowcolor=TEXT_DIM, borderwidth=0)

    # ── Sequence Management ──────────────────────────────────────────────────────
    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, s in enumerate(self.steps, 1):
            self.tree.insert("", "end", iid=str(i-1),
                             values=(i, s.x, s.y, s.button.title(), f"{s.delay:.2f}", s.description))
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
        win.geometry("440x340")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        win.transient(self)

        fields = {}

        def lbl_entry(parent, label, default, row):
            tk.Label(parent, text=label, bg=BG, fg=TEXT_DIM, font=FONT_UI).grid(
                row=row, column=0, sticky="w", padx=16, pady=6)
            var = tk.StringVar(value=str(default))
            e = tk.Entry(parent, textvariable=var, bg=SURFACE2, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         highlightbackground=BORDER, highlightthickness=1,
                         font=FONT_MONO, width=18)
            e.grid(row=row, column=1, padx=(0,16), pady=6)
            return var

        fields["x"]    = lbl_entry(win, "X coordinate:", step.x, 0)
        fields["y"]    = lbl_entry(win, "Y coordinate:", step.y, 1)
        fields["delay"]= lbl_entry(win, "Delay before (s):", step.delay, 2)

        tk.Label(win, text="Mouse button:", bg=BG, fg=TEXT_DIM, font=FONT_UI).grid(
            row=3, column=0, sticky="w", padx=16, pady=6)
        btn_var = tk.StringVar(value=step.button)
        btn_frame = tk.Frame(win, bg=BG)
        btn_frame.grid(row=3, column=1, sticky="w", pady=6)
        for b in ("left", "right", "middle"):
            tk.Radiobutton(btn_frame, text=b.title(), variable=btn_var, value=b,
                           bg=BG, fg=TEXT, selectcolor=SURFACE2,
                           activebackground=BG, font=FONT_UI).pack(side="left", padx=4)

        fields["description"] = lbl_entry(win, "Note (optional):", step.description, 4)

        def fill_from_capture():
            try:
                x, y = pyautogui.position()
                fields["x"].set(str(x))
                fields["y"].set(str(y))
            except Exception:
                pass

        tk.Button(win, text="📍 Use Current Mouse Position", bg=SURFACE2, fg=ACCENT,
                  relief="flat", font=FONT_UI, cursor="hand2",
                  command=fill_from_capture).grid(row=5, column=0, columnspan=2, pady=8)

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
                messagebox.showerror("Invalid Input", "X, Y must be integers and Delay must be a number.", parent=win)
                return
            if edit_idx is not None:
                self.steps[edit_idx] = s
            else:
                self.steps.append(s)
            self._refresh_tree()
            win.destroy()

        tk.Button(win, text="✔ Save Step", bg=ACCENT, fg=BG,
                  relief="flat", font=(FONT_UI[0], 10, "bold"), cursor="hand2",
                  padx=12, pady=8, command=save).grid(row=6, column=0, columnspan=2, pady=12)

    # ── Capture ──────────────────────────────────────────────────────────────────
    def _start_capture(self):
        if self.capturing:
            return
        self.capturing = True
        self.cap_btn.config(text="📍 Click anywhere...", fg=WARN)
        self.cap_coords.config(text="waiting for click...")

        def on_click(x, y, button, pressed):
            if pressed:
                self.after(0, lambda: self._on_captured(int(x), int(y)))
                return False  # stop listener

        self.capture_listener = mouse.Listener(on_click=on_click)
        self.capture_listener.start()

    def _on_captured(self, x, y):
        self.capturing = False
        self.cap_btn.config(text="🎯 Capture Position", fg=ACCENT2)
        self.cap_coords.config(text=f"X: {x}  Y: {y}", fg=ACCENT)
        # Ask if user wants to add as new step
        if messagebox.askyesno("Add Captured Position",
                               f"Captured ({x}, {y}).\nAdd as a new step?"):
            self.steps.append(ClickStep(x=x, y=y))
            self._refresh_tree()

    # ── Hotkey ───────────────────────────────────────────────────────────────────
    def _setup_hotkey(self):
        hk = self.hotkey_var.get().strip()
        if hasattr(self, "_hk_listener") and self._hk_listener:
            try:
                self._hk_listener.stop()
            except Exception:
                pass

        try:
            key_map = {f"F{i}": getattr(keyboard.Key, f"f{i}") for i in range(1, 13)}
            key_map.update({
                "ESCAPE": keyboard.Key.esc,
                "ESC":    keyboard.Key.esc,
                "TAB":    keyboard.Key.tab,
                "SPACE":  keyboard.Key.space,
                "HOME":   keyboard.Key.home,
                "END":    keyboard.Key.end,
                "INSERT": keyboard.Key.insert,
                "DELETE": keyboard.Key.delete,
                "PAUSE":  keyboard.Key.pause,
            })
            key_obj = key_map.get(hk.upper())

            def on_press(k):
                if k == key_obj:
                    self.after(0, self._toggle)

            self._hk_listener = keyboard.Listener(on_press=on_press)
            self._hk_listener.daemon = True
            self._hk_listener.start()
        except Exception as e:
            messagebox.showwarning("Hotkey", f"Could not bind hotkey '{hk}'.\n{e}")

    def _setup_record_hotkey(self):
        """Bind the bulk-add hotkey: press it anytime to snapshot current mouse position as a new step."""
        hk = self.record_hotkey_var.get().strip()
        if hasattr(self, "_rk_listener") and self._rk_listener:
            try:
                self._rk_listener.stop()
            except Exception:
                pass

        try:
            key_map = {f"F{i}": getattr(keyboard.Key, f"f{i}") for i in range(1, 13)}
            key_map.update({
                "ESCAPE": keyboard.Key.esc, "ESC":    keyboard.Key.esc,
                "TAB":    keyboard.Key.tab, "SPACE":  keyboard.Key.space,
                "HOME":   keyboard.Key.home,"END":    keyboard.Key.end,
                "INSERT": keyboard.Key.insert,"DELETE": keyboard.Key.delete,
                "PAUSE":  keyboard.Key.pause,
            })
            key_obj = key_map.get(hk.upper())

            def on_press(k):
                if k == key_obj:
                    self.after(0, self._record_step)

            self._rk_listener = keyboard.Listener(on_press=on_press)
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
        self.status_var.set("● Running")
        self.run_btn.config(text="■  STOP", fg=DANGER)
        self._update_status_color()
        self.macro_thread = threading.Thread(target=self._run_loop, daemon=True)
        self.macro_thread.start()

    def _stop(self):
        self.running = False
        self.status_var.set("Idle")
        self.run_btn.config(text="▶  START", fg=ACCENT2)
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
                time.sleep(max(0, step.delay))
                if not self.running:
                    break
                try:
                    pyautogui.click(step.x, step.y, button=step.button)
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
        color = ACCENT2 if self.running else TEXT_DIM
        self.status_label.config(fg=color)

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
            self.steps = [ClickStep.from_dict(d) for d in data.get("steps", [])]
            self._refresh_tree()
            self._setup_hotkey()
            self._setup_record_hotkey()
        except Exception:
            pass  # corrupted autosave — start fresh

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