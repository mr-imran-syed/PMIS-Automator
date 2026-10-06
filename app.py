"""PMIS Automation Tool - desktop GUI (v0.1).

A modern single-screen Windows utility around the existing Selenium automation.
Flow: startup readiness check -> collapse -> choose Training/Live -> credentials
-> select data -> Start -> animated status + live log -> Completed/Stopped/Error.
"""

import os
import queue
import threading
import importlib
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

import config
import deps
import data_file
import settings
import credstore
from config import (APP_NAME, APP_FOOTER, ENVIRONMENTS, UPDATE_TYPES,
                    BROWSER_DOWNLOAD_URL)

config.ensure_libs_on_path()

# ---- palette -------------------------------------------------------------
BG = "#f6f7f9"
CARD = "#ffffff"
BORDER = "#e2e5ea"
TEXT = "#1f2430"
MUTED = "#6b7280"
GREEN = "#1a7f37"
YELLOW = "#c9960a"
YELLOW_BG = "#fff3cd"
ORANGE = "#c2551a"
RED = "#c02626"
NAV_BG = "#eef1f6"          # unselected nav button
NAV_SEL_BG = "#dbe6fb"      # selected nav button
NAV_SEL_FG = "#12305e"

CAT_FRAMES = ["🐱", "😺", "😸", "😻", "😽"]
HAND_FRAMES = ["✋", "🤚", "🖐️", "👋"]
PARTY_FRAMES = ["🎉", "🎊", "🥳", "✨"]

# Each status state colours the ENTIRE bar: background, border, emoji and text.
STATUS_STYLES = {
    "idle": {
        "bg": "#f3f4f6", "fg": MUTED, "bd": BORDER,
        "text": "IDLE", "emoji": "•", "frames": None,
    },
    "running": {
        "bg": "#e6f4ea", "fg": GREEN, "bd": "#9ccfb0",
        "text": "Running", "emoji": "🐱", "frames": CAT_FRAMES,
    },
    "stopping": {
        "bg": "#fdebd8", "fg": ORANGE, "bd": "#efbc86",
        "text": "Stopping…", "emoji": "✋", "frames": HAND_FRAMES,
    },
    "stopped": {
        "bg": "#fdebd8", "fg": ORANGE, "bd": "#efbc86",
        "text": "Stopped by User", "emoji": "✋", "frames": HAND_FRAMES,
    },
    "completed": {
        "bg": "#d6f2e0", "fg": "#146c2e", "bd": "#7cc79b",
        "text": "Completed", "emoji": "🎉", "frames": PARTY_FRAMES,
    },
    "error": {
        "bg": "#fde8e8", "fg": RED, "bd": "#efb0b0",
        "text": "Error", "emoji": "🚨", "frames": None,
    },
}


class ToolTip:
    """Minimal hover tooltip."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, bg="#111827", fg="white",
                 font=("Segoe UI", 8), padx=6, pady=3, justify="left").pack()

    def _hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class ScrollFrame(tk.Frame):
    """A vertically scrollable container so the window works on small screens."""
    def __init__(self, parent, bg=BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical",
                                 command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.inner,
                                              anchor="nw")
        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        # Wheel scrolling only while the pointer is over this area.
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all(
            "<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all(
            "<MouseWheel>"))

    def _wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x760")
        self.minsize(520, 420)   # content scrolls, so allow small windows
        self.configure(bg=BG)

        self.log_queue = queue.Queue()
        self.ui_queue = queue.Queue()      # callables to run on the UI thread
        self.stop_event = threading.Event()
        self.worker = None
        self.driver = None
        self.running = False

        self.readiness = []            # list of status dicts
        self._mandatory_ready = False
        self._details_open = False
        self._anim_job = None
        self._anim_frames = None
        self._anim_i = 0

        self._build_style()
        self._build_ui()
        self.after(100, self._drain_log)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Startup readiness check (runs in a thread; UI updates via after()).
        self._set_readiness_checking()
        threading.Thread(target=self._check_readiness_worker, daemon=True).start()

    # ------------------------------------------------------------- styling
    def _build_style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TFrame", background=BG)
        s.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1)
        s.configure("TLabel", background=CARD, foreground=TEXT,
                    font=("Segoe UI", 10))
        s.configure("Bg.TLabel", background=BG, foreground=TEXT)
        s.configure("Muted.TLabel", background=CARD, foreground=MUTED,
                    font=("Segoe UI", 9))
        s.configure("H.TLabel", background=CARD, foreground=TEXT,
                    font=("Segoe UI", 12, "bold"))
        s.configure("TButton", font=("Segoe UI", 10), padding=6)
        s.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        s.configure("TEntry", padding=4)
        s.configure("TCombobox", padding=4)
        # Tabs down the left-hand side ("wn" = west, aligned north).
        s.configure("Side.TNotebook", tabposition="wn", background=BG,
                    borderwidth=0)
        s.configure("Side.TNotebook.Tab", padding=(14, 10),
                    font=("Segoe UI", 10))

    def _card(self, parent):
        f = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                     highlightthickness=1, bd=0)
        return f

    # ------------------------------------------------------------- layout
    def _build_ui(self):
        root = tk.Frame(self, bg=BG)
        root.pack(fill="both", expand=True, padx=14, pady=12)

        tk.Label(root, text=APP_NAME, bg=BG, fg=TEXT,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")

        # Status bar lives outside the tabs so it is visible from any tab.
        self._build_status(root)

        # Side tabs: Setup (scrollable) | Log
        # Side nav. Buttons live in a FIXED-WIDTH sidebar and fill it, so the
        # selected one can go bold without changing size (which is what a
        # ttk.Notebook does by design).
        nav_wrap = tk.Frame(root, bg=BG)
        nav_wrap.pack(fill="both", expand=True, pady=(10, 0))

        sidebar = tk.Frame(nav_wrap, bg=BG, width=118)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)        # keep the width fixed

        page_area = tk.Frame(nav_wrap, bg=BG)
        page_area.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self._pages = {}
        self._nav_buttons = {}
        for key, label in (("setup", "Setup"), ("log", "Log"),
                           ("settings", "Settings")):
            self._pages[key] = tk.Frame(page_area, bg=BG)
            btn = tk.Button(sidebar, text=label, relief="flat", bd=0,
                            highlightthickness=0, cursor="hand2",
                            activebackground=NAV_SEL_BG, pady=9,
                            command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", pady=(0, 6))   # fill => identical widths
            self._nav_buttons[key] = btn

        self.setup_tab = self._pages["setup"]
        self.log_tab = self._pages["log"]
        self.settings_tab = self._pages["settings"]

        # Setup content scrolls, so small screens can reach every control.
        self.setup_scroll = ScrollFrame(self.setup_tab, bg=BG)
        self.setup_scroll.pack(fill="both", expand=True)
        body = self.setup_scroll.inner

        self._build_readiness(body)
        self._build_workflow(body)
        self._build_log(self.log_tab)
        self._build_settings(self.settings_tab)
        self._build_footer(root)
        self.show_page("setup")
        self._load_credentials()

    def show_page(self, key):
        """Swap the visible page and restyle the nav.

        Only weight and colour change between states - never geometry - so
        nothing shifts when a different page is selected.
        """
        for frame in self._pages.values():
            frame.pack_forget()
        self._pages[key].pack(fill="both", expand=True)
        self._current_page = key
        for k, btn in self._nav_buttons.items():
            chosen = (k == key)
            btn.config(
                font=("Segoe UI", 10, "bold" if chosen else "normal"),
                bg=NAV_SEL_BG if chosen else NAV_BG,
                fg=NAV_SEL_FG if chosen else TEXT)

    # ---- readiness ----
    def _build_readiness(self, root):
        card = self._card(root)
        card.pack(fill="x", pady=(0, 10))
        card.columnconfigure(0, weight=1)
        self.readiness_card = card

        # Compact bar (shown once ready)
        self.compact = tk.Frame(card, bg=CARD)
        self.compact.columnconfigure(0, weight=1)
        self.compact_lbl = tk.Label(self.compact, bg=CARD, fg=GREEN,
                                    font=("Segoe UI", 10, "bold"), anchor="w")
        self.compact_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=8)
        self.details_btn = ttk.Button(self.compact, text="View Details ▾",
                                      command=self._toggle_details)
        self.details_btn.grid(row=0, column=1, padx=4)
        self.update_all_btn = ttk.Button(self.compact, text="Update All",
                                         command=self._update_all)
        # (gridded only when updates exist)

        # Full checklist
        self.checklist = tk.Frame(card, bg=CARD)
        self.checklist.columnconfigure(0, weight=1)
        tk.Label(self.checklist, text="System Readiness", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        self.rows_frame = tk.Frame(self.checklist, bg=CARD)
        self.rows_frame.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))
        self.rows_frame.columnconfigure(0, weight=1)

        # Install/update output (also used for View Error)
        self.dep_output = tk.Text(self.checklist, height=6, font=("Consolas", 9),
                                  bg="#0f172a", fg="#e2e8f0", wrap="word")
        self.dep_prog = ttk.Progressbar(self.checklist, mode="indeterminate")

        self.checklist.grid(row=0, column=0, sticky="ew")   # visible first

    def _set_readiness_checking(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        tk.Label(self.rows_frame, text="Checking dependencies…", bg=CARD,
                 fg=MUTED, font=("Segoe UI", 10)).grid(row=0, column=0,
                                                       sticky="w", padx=6, pady=6)

    def _check_readiness_worker(self):
        # Fast pass first (no network) so the UI is usable immediately…
        fast = deps.check_all(check_updates=False)
        self._call(lambda: self._apply_readiness(fast))
        # …then refine with PyPI update info in the background.
        full = deps.check_all(check_updates=True)
        self._call(lambda: self._apply_readiness(full))

    def _apply_readiness(self, results):
        self.readiness = results
        mand = [r for r in results if r["mandatory"]]
        self._mandatory_ready = all(r["state"] != deps.STATE_MISSING for r in mand)
        updates = [r for r in results if r["state"] == deps.STATE_UPDATE]

        self._render_rows()

        if self._mandatory_ready:
            # Collapse to compact bar
            self.checklist.grid_remove()
            self.compact.grid(row=0, column=0, sticky="ew")
            if updates:
                self.compact_lbl.config(
                    text=f"⚠ System Ready — {len(updates)} update(s) available",
                    fg=YELLOW)
                self.update_all_btn.grid(row=0, column=2, padx=(0, 8))
            else:
                self.compact_lbl.config(
                    text="✅ System Ready — All dependencies available", fg=GREEN)
                self.update_all_btn.grid_forget()
            if self._details_open:
                self._show_details(True)
        else:
            self.compact.grid_remove()
            self.checklist.grid(row=0, column=0, sticky="ew")

        self._update_start_state()

    def _render_rows(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        for i, r in enumerate(self.readiness):
            row = tk.Frame(self.rows_frame, bg=CARD)
            row.grid(row=i, column=0, sticky="ew", padx=4, pady=2)
            row.columnconfigure(0, weight=1)
            tk.Label(row, text=r["name"], bg=CARD, fg=TEXT,
                     font=("Segoe UI", 10), anchor="w").grid(
                row=0, column=0, sticky="w")

            state, color, txt = r["state"], MUTED, ""
            if state == deps.STATE_READY:
                color, txt = GREEN, "✅ Ready"
            elif state == deps.STATE_UPDATE:
                color, txt = YELLOW, "⚠ Update Available"
            elif state == deps.STATE_MISSING:
                color, txt = RED, "❌ Missing"
            else:
                color, txt = RED, "❌ Error"
            extra = r.get("version") or r.get("detail") or ""
            if extra and state == deps.STATE_READY:
                txt = f"✅ Ready · {extra}"
            lbl = tk.Label(row, text=txt, bg=CARD, fg=color,
                           font=("Segoe UI", 10), anchor="e")
            lbl.grid(row=0, column=1, padx=8)
            ToolTip(lbl, r.get("detail") or r["name"])

            # Action button
            if state == deps.STATE_MISSING and r["kind"] in ("package",):
                b = ttk.Button(row, text="Install",
                               command=lambda d=r: self._install(d, upgrade=False))
                b.grid(row=0, column=2)
                ToolTip(b, f"Install {r['name']} into the app's libs folder")
            elif state == deps.STATE_UPDATE and r["kind"] == "package":
                b = ttk.Button(row, text="Update",
                               command=lambda d=r: self._install(d, upgrade=True))
                b.grid(row=0, column=2)
                ToolTip(b, f"Update {r['name']} to the latest version")
            elif state == deps.STATE_UPDATE and r["kind"] == "app":
                b = ttk.Button(row, text="Get Update",
                               command=self._open_releases)
                b.grid(row=0, column=2)
                ToolTip(b, "Open the GitHub releases page to download the "
                           "newer version")
            elif state == deps.STATE_MISSING and r["kind"] == "browser":
                b = ttk.Button(row, text="Install",
                               command=lambda: webbrowser.open(BROWSER_DOWNLOAD_URL))
                b.grid(row=0, column=2)
                ToolTip(b, "Open the browser download page")

    def _open_releases(self):
        repo = getattr(config, "GITHUB_REPO", None)
        if not repo or "REPLACE_ME" in repo:
            messagebox.showinfo(
                APP_NAME, "No GitHub repository is configured for updates.\n\n"
                          "Set GITHUB_REPO in config.py.")
            return
        webbrowser.open(f"https://github.com/{repo}/releases")

    def _toggle_details(self):
        self._details_open = not self._details_open
        self._show_details(self._details_open)

    def _show_details(self, show):
        if show:
            self.checklist.grid(row=1, column=0, sticky="ew")
            self.details_btn.config(text="Hide Details ▴")
        else:
            self.checklist.grid_remove()
            self.details_btn.config(text="View Details ▾")

    # ---- dependency install/update ----
    def _dep_log(self, line):
        self._call(lambda: self._append_dep_output(line))

    def _append_dep_output(self, line):
        self.dep_output.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 4))
        self.dep_output.insert("end", line + "\n")
        self.dep_output.see("end")

    def _install(self, dep_status, upgrade):
        entry = deps.dependency_by_key(dep_status["key"])
        if not entry:
            return
        self._details_open = True
        self._show_details(True)
        self.dep_output.delete("1.0", "end")
        self.dep_prog.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))
        self.dep_prog.start(12)
        verb = "Updating" if upgrade else "Installing"
        self._append_dep_output(f"{verb} {entry['name']}…")

        def work():
            ok = deps.install_package(entry["pip_name"], self._dep_log, upgrade)
            self._call(lambda: self._after_install(ok))
        threading.Thread(target=work, daemon=True).start()

    def _after_install(self, ok):
        self.dep_prog.stop()
        self.dep_prog.grid_forget()
        importlib.invalidate_caches()
        if not ok:
            self._append_dep_output("Installation failed. See output above.")
        # Re-check and re-render.
        threading.Thread(target=self._check_readiness_worker, daemon=True).start()

    def _update_all(self):
        updates = [r for r in self.readiness if r["state"] == deps.STATE_UPDATE
                   and r["kind"] == "package"]
        if not updates:
            return
        self._details_open = True
        self._show_details(True)
        self.dep_output.delete("1.0", "end")
        self.dep_prog.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))
        self.dep_prog.start(12)

        def work():
            for r in updates:
                entry = deps.dependency_by_key(r["key"])
                self._dep_log(f"--- Updating {entry['name']} ---")
                deps.install_package(entry["pip_name"], self._dep_log, True)
            self._call(lambda: self._after_install(True))
        threading.Thread(target=work, daemon=True).start()

    # ---- workflow ----
    def _build_workflow(self, root):
        card = self._card(root)
        card.pack(fill="x", pady=(0, 10))
        card.columnconfigure(1, weight=1)
        pad = dict(padx=10, pady=6)

        # Environment
        tk.Label(card, text="Environment", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).grid(row=0, column=0,
                                                     sticky="w", **pad)
        env_row = tk.Frame(card, bg=CARD)
        env_row.grid(row=0, column=1, sticky="w", **pad)
        # Training is the default. NOTE: an empty value equals tk's default
        # tristatevalue, which makes BOTH radios render as pre-selected --
        # hence a real default plus an unreachable tristatevalue.
        self.env_var = tk.StringVar(value="Training")
        self.train_rb = tk.Radiobutton(
            env_row, text="Training", value="Training", variable=self.env_var,
            bg=CARD, fg=TEXT, selectcolor=CARD, activebackground=CARD,
            tristatevalue="\x00",
            font=("Segoe UI", 10), command=self._on_env_change)
        self.train_rb.pack(side="left", padx=(0, 12))
        self.live_rb = tk.Radiobutton(
            env_row, text="Live", value="Live", variable=self.env_var,
            bg=CARD, fg=TEXT, selectcolor=CARD, activebackground=CARD,
            tristatevalue="\x00",
            font=("Segoe UI", 10, "bold"), command=self._on_env_change)
        self.live_rb.pack(side="left")
        ToolTip(self.live_rb, "Live is the real/production environment")

        # Username
        tk.Label(card, text="Username", bg=CARD, fg=TEXT).grid(
            row=1, column=0, sticky="w", **pad)
        self.user_var = tk.StringVar()
        self.user_var.trace_add("write", lambda *_: self._update_start_state())
        self.user_entry = ttk.Entry(card, textvariable=self.user_var)
        self.user_entry.grid(row=1, column=1, sticky="ew", **pad)

        # Password + hold-to-reveal
        tk.Label(card, text="Password", bg=CARD, fg=TEXT).grid(
            row=2, column=0, sticky="w", **pad)
        pw_row = tk.Frame(card, bg=CARD)
        pw_row.grid(row=2, column=1, sticky="ew", **pad)
        pw_row.columnconfigure(0, weight=1)
        self.pw_var = tk.StringVar()
        self.pw_var.trace_add("write", lambda *_: self._update_start_state())
        self.pw_entry = ttk.Entry(pw_row, textvariable=self.pw_var, show="*")
        self.pw_entry.grid(row=0, column=0, sticky="ew")
        reveal = tk.Button(pw_row, text="C", width=3, relief="groove",
                           bg="#eef1f6")
        reveal.grid(row=0, column=1, padx=(6, 0))
        reveal.bind("<ButtonPress-1>", lambda e: self.pw_entry.config(show=""))
        reveal.bind("<ButtonRelease-1>", lambda e: self.pw_entry.config(show="*"))
        ToolTip(reveal, "Hold to reveal the password; release to hide")

        # Data to be updated
        tk.Label(card, text="Select the Data to be Updated", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).grid(
            row=3, column=0, sticky="w", **pad)
        self.data_var = tk.StringVar(value="")
        opts = list(UPDATE_TYPES.keys()) + ["Other"]
        self.data_cb = ttk.Combobox(card, textvariable=self.data_var,
                                    state="readonly", values=opts)
        self.data_cb.grid(row=3, column=1, sticky="ew", **pad)
        self.data_cb.bind("<<ComboboxSelected>>",
                          lambda *_: self._update_start_state())

        # Data file row
        tk.Label(card, text="Data File", bg=CARD, fg=TEXT).grid(
            row=4, column=0, sticky="w", **pad)
        file_row = tk.Frame(card, bg=CARD)
        file_row.grid(row=4, column=1, sticky="ew", **pad)
        self.open_btn = ttk.Button(file_row, text="Open Data File",
                                   command=self._open_data_file)
        self.open_btn.pack(side="left")
        ToolTip(self.open_btn, "Open the bundled Excel to edit and save your data")
        self.file_status = tk.Label(file_row, bg=CARD, fg=MUTED,
                                    font=("Segoe UI", 9))
        self.file_status.pack(side="left", padx=10)
        # A fresh clone has only the template; make a working copy from it.
        try:
            _, created = data_file.ensure_data_file()
            if created:
                self.log(f"Created {config.DATA_FILE_NAME} from the template.")
        except Exception:
            pass
        self._refresh_file_status()

        # Start button
        btn_row = tk.Frame(card, bg=CARD)
        btn_row.grid(row=5, column=0, columnspan=2, pady=12)
        self.start_btn = tk.Button(btn_row, text="Start", command=self._start,
                                   font=("Segoe UI", 12, "bold"),
                                   bg="#d1d5db", fg="#374151", relief="flat",
                                   padx=16, pady=8, state="disabled")
        self.start_btn.pack(side="left")
        self.close_browser_btn = ttk.Button(btn_row, text="Close Browser",
                                            command=self._close_browser,
                                            state="disabled")
        self.close_browser_btn.pack(side="left", padx=(10, 0))
        ToolTip(self.close_browser_btn,
                "The browser stays open between runs so you can inspect the "
                "page. Click to close it.")

        self._on_env_change()   # apply the Training default styling

    def _on_env_change(self):
        # Highlight Live in yellow (production warning); Training normal.
        live_font = ("Segoe UI", 10, "bold")     # Live stays bold in both states
        if self.env_var.get() == "Live":
            self.live_rb.config(bg=YELLOW_BG, fg=YELLOW, selectcolor=YELLOW_BG,
                                activebackground=YELLOW_BG, font=live_font)
            self.train_rb.config(bg=CARD, fg=TEXT, selectcolor=CARD)
        else:
            self.train_rb.config(bg=CARD, fg=TEXT, selectcolor=CARD)
            self.live_rb.config(bg=CARD, fg=TEXT, selectcolor=CARD,
                                activebackground=CARD, font=live_font)
        self._update_start_state()

    def _refresh_file_status(self):
        if not data_file.exists():
            self.file_status.config(text="⚠ data file not found", fg=RED)
        elif data_file.is_open():
            self.file_status.config(text="● open in Excel", fg=YELLOW)
        else:
            self.file_status.config(text="ready", fg=MUTED)

    def _open_data_file(self):
        try:
            data_file.open_in_excel()
        except FileNotFoundError:
            messagebox.showerror(APP_NAME, "The data file was not found next to "
                                           "the application.")
            return
        self.after(1200, self._refresh_file_status)

    # ---- status ----
    def _build_status(self, root):
        card = self._card(root)
        card.pack(fill="x", pady=(10, 0))
        card.columnconfigure(1, weight=1)
        self.status_card = card
        self.status_emoji = tk.Label(card, text="•", bg=CARD, fg=MUTED,
                                     font=("Segoe UI Emoji", 18))
        self.status_emoji.grid(row=0, column=0, padx=(12, 6), pady=10)
        self.status_text = tk.Label(card, text="IDLE", bg=CARD, fg=MUTED,
                                    font=("Segoe UI", 12, "bold"), anchor="w")
        self.status_text.grid(row=0, column=1, sticky="w")
        self._set_status("idle")

    def _set_status(self, state):
        """Apply a state's full-bar colour scheme: card background, border,
        emoji and text all change together. state in STATUS_STYLES."""
        style = STATUS_STYLES.get(state, STATUS_STYLES["idle"])
        self._stop_anim()

        bg, fg = style["bg"], style["fg"]
        self.status_card.config(bg=bg, highlightbackground=style["bd"],
                                highlightcolor=style["bd"])
        self.status_emoji.config(bg=bg, fg=fg, text=style["emoji"])
        self.status_text.config(bg=bg, fg=fg, text=style["text"])

        if style["frames"]:
            self._start_anim(style["frames"])

    def _start_anim(self, frames):
        self._anim_frames = frames
        self._anim_i = 0
        self._tick_anim()

    def _tick_anim(self):
        if not self._anim_frames:
            return
        self.status_emoji.config(text=self._anim_frames[self._anim_i])
        self._anim_i = (self._anim_i + 1) % len(self._anim_frames)
        self._anim_job = self.after(450, self._tick_anim)

    def _stop_anim(self):
        if self._anim_job:
            self.after_cancel(self._anim_job)
            self._anim_job = None
        self._anim_frames = None

    # ---- log ----
    def _build_log(self, root):
        card = self._card(root)
        card.pack(fill="both", expand=True)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        head = tk.Frame(card, bg=CARD)
        head.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        head.columnconfigure(0, weight=1)
        tk.Label(head, text="Log", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")
        # Runs jump to this tab, so Stop has to be reachable from here too.
        self.log_stop_btn = tk.Button(
            head, text="Stop", command=self._stop, state="disabled",
            font=("Segoe UI", 10, "bold"), relief="flat",
            bg="#d1d5db", fg="#374151", padx=14, pady=3)
        self.log_stop_btn.grid(row=0, column=1, padx=(0, 8))
        ToolTip(self.log_stop_btn,
                "Stop after the current row (same as Stop on the Setup page)")

        copy_btn = ttk.Button(head, text="Copy Log", command=self._copy_log)
        copy_btn.grid(row=0, column=2)
        ToolTip(copy_btn, "Copy the entire current log to the clipboard")

        wrap = tk.Frame(card, bg=CARD)
        wrap.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.log_box = tk.Text(wrap, height=12, font=("Consolas", 9),
                               bg="#0f172a", fg="#e2e8f0", wrap="word",
                               state="disabled")
        self.log_box.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, command=self.log_box.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log_box.config(yscrollcommand=sb.set)

    def _copy_log(self):
        self.clipboard_clear()
        self.clipboard_append(self.log_box.get("1.0", "end-1c"))

    def log(self, msg):
        self.log_queue.put(str(msg))

    def _call(self, fn):
        """Schedule fn to run on the UI thread (safe to call from workers)."""
        self.ui_queue.put(fn)

    def _drain_log(self):
        # Run queued UI callables from worker threads.
        try:
            while True:
                fn = self.ui_queue.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass
        # Flush log lines.
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", msg + "\n")
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ---- settings ----
    def _build_settings(self, parent):
        scroll = ScrollFrame(parent, bg=BG)
        scroll.pack(fill="both", expand=True)
        body = scroll.inner

        card = self._card(body)
        card.pack(fill="x", pady=(0, 10))
        card.columnconfigure(0, weight=1)

        tk.Label(card, text="Participant Master Data", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 2))

        self.auto_save_var = tk.BooleanVar(
            value=bool(settings.get("participant_auto_save")))
        chk = tk.Checkbutton(
            card, text="Auto Save", variable=self.auto_save_var,
            command=self._on_auto_save_toggle, bg=CARD, fg=TEXT,
            selectcolor=CARD, activebackground=CARD,
            font=("Segoe UI", 10, "bold"))
        chk.grid(row=1, column=0, sticky="w", padx=6)
        ToolTip(chk, "On: save each entry to PMIS and move to the next.\n"
                     "Off: fill the form only, nothing is submitted.")

        tk.Label(card, bg=CARD, fg=MUTED, justify="left", wraplength=520,
                 font=("Segoe UI", 9),
                 text=("On  - fills the form, saves the entry to PMIS, marks "
                       "the row 'Updated' in the sheet, and moves to the next "
                       "participant.\n"
                       "Off - fills the form for the first pending "
                       "participant and stops so you can review it. Nothing "
                       "is submitted.")).grid(
            row=2, column=0, sticky="w", padx=28, pady=(0, 6))

        self.auto_save_warn = tk.Label(
            card, bg=CARD, fg=ORANGE, justify="left", wraplength=520,
            font=("Segoe UI", 9, "bold"), text="")
        self.auto_save_warn.grid(row=3, column=0, sticky="w",
                                 padx=28, pady=(0, 10))
        self._refresh_auto_save_warning()

        shg_card = self._card(body)
        shg_card.pack(fill="x", pady=(0, 10))
        shg_card.columnconfigure(0, weight=1)
        tk.Label(shg_card, text="SHG Creation", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 2))
        self.shg_auto_save_var = tk.BooleanVar(
            value=bool(settings.get("shg_auto_save")))
        shg_chk = tk.Checkbutton(
            shg_card, text="Auto Save", variable=self.shg_auto_save_var,
            command=self._on_shg_auto_save_toggle, bg=CARD, fg=TEXT,
            selectcolor=CARD, activebackground=CARD,
            font=("Segoe UI", 10, "bold"))
        shg_chk.grid(row=1, column=0, sticky="w", padx=6)
        ToolTip(shg_chk, "On: create each SHG in PMIS and move to the next.\n"
                         "Off: fill the first one only, nothing is saved.")
        tk.Label(shg_card, bg=CARD, fg=MUTED, justify="left", wraplength=520,
                 font=("Segoe UI", 9),
                 text=("The outcome of every row is written into the sheet's "
                       "PMIS-Update column - 'Updated' when it is created, or "
                       "the error PMIS reported (for example \"SHG Name "
                       "already exists\") so you can see why it "
                       "did not.")).grid(
            row=2, column=0, sticky="w", padx=28, pady=(0, 10))

        cred_card = self._card(body)
        cred_card.pack(fill="x", pady=(0, 10))
        cred_card.columnconfigure(0, weight=1)

        tk.Label(cred_card, text="Credentials", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 2))

        self.keep_creds_var = tk.BooleanVar(
            value=bool(settings.get("keep_credentials")))
        kc = tk.Checkbutton(
            cred_card, text="Keep credentials", variable=self.keep_creds_var,
            command=self._on_keep_creds_toggle, bg=CARD, fg=TEXT,
            selectcolor=CARD, activebackground=CARD,
            font=("Segoe UI", 10, "bold"))
        kc.grid(row=1, column=0, sticky="w", padx=6)
        ToolTip(kc, "Remember the username and password for next launch.\n"
                    "The password is encrypted for your Windows account.")

        tk.Label(cred_card, bg=CARD, fg=MUTED, justify="left", wraplength=520,
                 font=("Segoe UI", 9),
                 text=("Pre-fills the username and password when the app "
                       "opens. The password is encrypted with your Windows "
                       "account, so the file is useless on another PC or "
                       "under another user - but anyone signed in as you on "
                       "this PC could use it. Leave off on shared "
                       "machines.")).grid(
            row=2, column=0, sticky="w", padx=28, pady=(0, 10))

        note = self._card(body)
        note.pack(fill="x")
        tk.Label(note, bg=CARD, fg=MUTED, justify="left", wraplength=540,
                 font=("Segoe UI", 9),
                 text=(f"Settings are stored in {settings.SETTINGS_FILE} next "
                       f"to the app. Usernames and passwords are never "
                       f"saved - they are typed each time you run.")).pack(
            anchor="w", padx=10, pady=10)

    def _refresh_auto_save_warning(self):
        if self.auto_save_var.get():
            self.auto_save_warn.config(
                text="Auto Save is ON - runs will write records into PMIS.")
        else:
            self.auto_save_warn.config(text="")

    def _on_shg_auto_save_toggle(self):
        settings.set_value("shg_auto_save", bool(self.shg_auto_save_var.get()))

    def _on_keep_creds_toggle(self):
        keep = bool(self.keep_creds_var.get())
        settings.set_value("keep_credentials", keep)
        if keep:
            self._store_credentials()
        else:
            settings.set_value("saved_username", "")
            settings.set_value("saved_password", "")
            self.log("Stored credentials cleared.")

    def _store_credentials(self):
        """Persist the current fields when Keep credentials is on."""
        if not self.keep_creds_var.get():
            return
        settings.set_value("saved_username", self.user_var.get().strip())
        token = credstore.protect(self.pw_var.get())
        if token is None and self.pw_var.get():
            settings.set_value("saved_password", "")
            messagebox.showwarning(
                APP_NAME, "The username was saved, but the password could not "
                          "be encrypted on this system, so it was not stored.")
            return
        settings.set_value("saved_password", token or "")

    def _load_credentials(self):
        """Pre-fill the fields if Keep credentials was left on."""
        if not settings.get("keep_credentials"):
            return
        user = settings.get("saved_username") or ""
        if user:
            self.user_var.set(user)
        pw = credstore.unprotect(settings.get("saved_password"))
        if pw:
            self.pw_var.set(pw)
        if user or pw:
            self.log("Credentials restored from Settings.")

    def _on_auto_save_toggle(self):
        value = bool(self.auto_save_var.get())
        if not settings.set_value("participant_auto_save", value):
            messagebox.showwarning(
                APP_NAME, "Could not save the setting to disk. It will apply "
                          "for this session only.")
        self._refresh_auto_save_warning()

    # ---- footer ----
    def _build_footer(self, root):
        tk.Label(root, text=APP_FOOTER, bg=BG, fg=MUTED,
                 font=("Segoe UI", 8)).pack(pady=(8, 0))

    # ------------------------------------------------------- start / run
    def _update_start_state(self):
        ok = (self._mandatory_ready
              and self.env_var.get() in ENVIRONMENTS
              and self.user_var.get().strip()
              and self.pw_var.get()
              and self.data_var.get() in UPDATE_TYPES)   # 'Other' not runnable
        if self.running:
            ok = False
        if ok:
            self.start_btn.config(state="normal", bg="#1a7f37", fg="white")
        else:
            self.start_btn.config(state="disabled", bg="#d1d5db", fg="#374151")

    def _start(self):
        if self.running:
            return
        dtype = self.data_var.get()
        if dtype not in UPDATE_TYPES:
            messagebox.showinfo(APP_NAME, "Select a data type to update.")
            return
        spec = UPDATE_TYPES[dtype]
        if spec["runner"] is None:
            messagebox.showinfo(
                APP_NAME, f"'{dtype}' automation is not available in v0.1 yet.")
            return
        if not data_file.exists():
            messagebox.showerror(APP_NAME, "Data file not found next to the app.")
            return
        if data_file.is_open():
            messagebox.showwarning(
                APP_NAME, "The Excel data file is open.\n\nPlease close it in "
                          "Excel, then click Start again. The app needs to write "
                          "progress into the file while it runs.")
            self._refresh_file_status()
            return

        # Optional pre-flight data validation (UI thread, before the browser
        # starts) so bad input is reported as a popup instead of a failed run.
        if not self._run_validator(spec):
            return

        env = self.env_var.get()
        if env == "Live":
            if not messagebox.askokcancel(
                    "Live Environment",
                    "⚠️ You are about to run this update in the LIVE environment.\n\n"
                    "Continue?"):
                return

        self._store_credentials()
        self._enter_running_state()
        self.worker = threading.Thread(
            target=self._run_worker,
            args=(dtype, spec, env, self.user_var.get().strip(), self.pw_var.get()),
            daemon=True)
        self.worker.start()

    def _run_validator(self, spec):
        """Run an update type's 'validator' hook, if it has one.
        Returns True when it is safe to proceed."""
        dotted = spec.get("validator")
        if not dotted:
            return True
        try:
            mod_name, func_name = dotted.split(":")
            fn = getattr(importlib.import_module(mod_name), func_name)
            ok, message = fn(data_file.data_file_path(), spec["sheet"])
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Could not validate the data file:\n{e}")
            return False
        if not ok:
            messagebox.showwarning(APP_NAME, message)
            return False
        return True

    def _enter_running_state(self):
        self.running = True
        self.stop_event.clear()
        self._clear_log()
        self.show_page("log")            # show progress straight away
        self._set_status("running")
        self.start_btn.config(text="Stop", command=self._stop,
                              state="normal", bg=ORANGE, fg="white")
        self.log_stop_btn.config(state="normal", bg=ORANGE, fg="white")
        for w in (self.train_rb, self.live_rb, self.data_cb, self.open_btn,
                  self.user_entry, self.pw_entry):
            try:
                w.config(state="disabled")
            except tk.TclError:
                pass

    def _leave_running_state(self):
        self.running = False
        self.start_btn.config(text="Start", command=self._start)
        self.log_stop_btn.config(text="Stop", state="disabled",
                                 bg="#d1d5db", fg="#374151")
        for w in (self.train_rb, self.live_rb, self.open_btn,
                  self.user_entry, self.pw_entry):
            try:
                w.config(state="normal")
            except tk.TclError:
                pass
        self.data_cb.config(state="readonly")
        self._refresh_file_status()
        self._update_start_state()

    def _stop(self):
        if not self.running:
            return
        self.stop_event.set()
        self._set_status("stopping")
        for btn in (self.start_btn, self.log_stop_btn):
            try:
                btn.config(text="Stopping...", state="disabled")
            except Exception:
                pass
        self.log("Stop requested - finishing the current row.")

    def _run_worker(self, dtype, spec, env, user, pw):
        result = {"stopped": False, "error": None}
        try:
            base_url = ENVIRONMENTS[env]
            from browsers import make_driver, BrowserNotFound
            if self._driver_alive():
                self.log("Reusing the browser already open (session kept).")
            else:
                self.log(f"Launching browser for {env} ({base_url}) …")
                try:
                    self.driver, name = make_driver()
                except BrowserNotFound as e:
                    result["error"] = str(e)
                    return
                self.log(f"Using {name.title()}.")

            mod_name, func_name = spec["runner"].split(":")
            runner = getattr(importlib.import_module(mod_name), func_name)
            runner(self.driver, data_file.data_file_path(), base_url, user, pw,
                   sheet=spec["sheet"], status_col=spec["status_col"],
                   log=self.log, stop_event=self.stop_event)
            result["stopped"] = self.stop_event.is_set()
        except Exception as e:
            import traceback
            result["error"] = str(e)
            self.log(f"ERROR: {e}")
            self.log(traceback.format_exc())
        finally:
            # Deliberately NOT quitting the driver: the browser stays open so
            # an error can be inspected (and an unsaved form reviewed), and the
            # next run reuses this signed-in session.
            self._call(lambda: self._on_worker_done(result))

    def _on_worker_done(self, result):
        if result["error"]:
            self._set_status("error")
        elif result["stopped"]:
            self._set_status("stopped")
        else:
            self._set_status("completed")
        self._leave_running_state()
        self._update_browser_btn()

    def _driver_alive(self):
        """True if self.driver still points at a usable browser."""
        if self.driver is None:
            return False
        try:
            self.driver.title              # cheap round-trip to the browser
            return True
        except Exception:
            try:
                # A leftover modal dialog also blocks commands - clear it.
                self.driver.switch_to.alert.accept()
                self.driver.title
                self.log("Dismissed a leftover dialog in the open browser.")
                return True
            except Exception:
                self.driver = None
                return False

    def _update_browser_btn(self):
        try:
            state = "normal" if self._driver_alive() else "disabled"
            self.close_browser_btn.config(state=state)
        except Exception:
            pass

    def _close_browser(self):
        if self.running:
            messagebox.showinfo(APP_NAME, "A run is in progress - press Stop "
                                          "first.")
            return
        self._quit_driver()
        self._update_browser_btn()
        self.log("Browser closed.")

    def _quit_driver(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ------------------------------------------------------- close
    def _on_close(self):
        if self.running:
            if not messagebox.askokcancel(
                    APP_NAME,
                    "⚠️ A process is currently running.\n\nClosing the application "
                    "will stop the current process.\n\n[OK] Stop & Exit   "
                    "[Cancel] Keep Running"):
                return
            self.stop_event.set()
        self._quit_driver()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
