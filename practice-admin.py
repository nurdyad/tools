import ctypes
from ctypes import wintypes
import json
import logging
import os
import re
import secrets
import string
import subprocess
import sys
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    import pyautogui
    import pygetwindow as gw
    import pyperclip
    from pywinauto import Application, Desktop

    AUTOMATION_READY = True
    AUTOMATION_IMPORT_ERROR = ""
except Exception as exc:
    pyautogui = None
    gw = None
    pyperclip = None
    Application = None
    Desktop = None
    AUTOMATION_READY = False
    AUTOMATION_IMPORT_ERROR = str(exc)


if getattr(sys, "frozen", False):
    SCRIPT_DIR = sys._MEIPASS
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

PROJECT_BASE = r"C:\rpa\postie"
ROOT_FOLDERS = [
    os.path.join(PROJECT_BASE, "postie_bots_python", "devdata", "work-items-in"),
    os.path.join(PROJECT_BASE, "postie-bots", "devdata", "work-items-in"),
]
GIT_REPO_PATH = os.path.join(PROJECT_BASE, "postie-bots")
CHECK_ODS_MISMATCH_SCRIPT = os.path.join(SCRIPT_DIR, "check-ods-mismatch.ps1")
GIT_PROFILE_FILE = "git-account-profile.json"


def get_safe_log_path(filename):
    user_profile = os.environ.get("USERPROFILE", "")
    paths_to_try = [
        os.path.join(user_profile, "OneDrive", "Desktop"),
        os.path.join(user_profile, "Desktop"),
        SCRIPT_DIR,
    ]
    for path in paths_to_try:
        if os.path.isdir(path):
            return os.path.join(path, filename)
    return os.path.join(SCRIPT_DIR, filename)


LOG_FILE_PATH = get_safe_log_path("password_log.txt")
DEBUG_LOG_PATH = get_safe_log_path("debug_log.txt")
PENDING_GIT_OPS_PATH = get_safe_log_path("pending-git-ops.json")

for path in [LOG_FILE_PATH, DEBUG_LOG_PATH]:
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("")


class SafeFileHandler(logging.FileHandler):
    def emit(self, record):
        super().emit(record)
        self.flush()

    def flush(self):
        super().flush()
        try:
            if self.stream and hasattr(self.stream, "fileno"):
                os.fsync(self.stream.fileno())
        except OSError:
            pass


class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see(tk.END)

        self.text_widget.after(0, append)


def generate_strict_password(length=10):
    alphabet = string.ascii_letters + string.digits
    while True:
        password = "".join(secrets.choice(alphabet) for _ in range(length))
        if any(ch.isupper() for ch in password) and any(ch.isdigit() for ch in password):
            return password


class UnifiedToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Practice + EMIS Operations Tool")
        self.root.geometry("940x840")
        self.root.minsize(860, 680)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.generated_pwd = ""
        self.last_practices = []
        self.emis_logger = None
        self._busy_count = 0
        self._hotkeys_registered = False
        self._instant_paste_hotkey_id = 0xB001
        self._instant_copy_hotkey_id = 0xB002
        self._registered_hotkey_ids = set()
        self.pending_git_ops = {}

        self._setup_styles()
        self._show_splash()

    def _setup_styles(self):
        # NEW PALETTE: Emerald & Zinc (Modern Dark Theme)
        BG_APP = "#18181b"          # zinc-900
        BG_CARD = "#27272a"         # zinc-800
        BG_INPUT = "#09090b"        # zinc-950
        BORDER = "#3f3f46"          # zinc-700
        
        TEXT_MAIN = "#f4f4f5"       # zinc-100
        TEXT_SUB = "#a1a1aa"        # zinc-400
        
        ACCENT_PRIMARY = "#10b981"  # emerald-500
        ACCENT_HOVER = "#34d399"    # emerald-400
        ACCENT_PRESS = "#059669"    # emerald-600
        
        DANGER_PRIMARY = "#f43f5e"  # rose-500
        DANGER_HOVER = "#fb7185"    # rose-400
        DANGER_PRESS = "#e11d48"    # rose-600

        self.root.configure(bg=BG_APP)
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Root.TFrame", background=BG_APP)
        style.configure("Card.TFrame", background=BG_CARD)
        style.configure(
            "Card.TLabelframe",
            background=BG_CARD,
            borderwidth=1,
            relief="solid",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=BG_CARD,
            foreground=ACCENT_PRIMARY,
            font=("Segoe UI Semibold", 11),
            padding=(5, 0)
        )
        style.configure("TLabel", background=BG_CARD, foreground=TEXT_MAIN, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=BG_APP, foreground=TEXT_MAIN, font=("Segoe UI Semibold", 22))
        style.configure("SubHeader.TLabel", background=BG_APP, foreground=TEXT_SUB, font=("Segoe UI", 10))
        
        style.configure(
            "StatusPill.TLabel",
            background="#064e3b",      # emerald-900
            foreground="#6ee7b7",      # emerald-300
            borderwidth=1,
            relief="solid",
            lightcolor="#065f46",      # emerald-800
            darkcolor="#065f46",
            bordercolor="#065f46",
            padding=(12, 6),
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "StatusCard.TFrame",
            background="#0b3328",
            borderwidth=1,
            relief="solid",
            lightcolor="#14532d",
            darkcolor="#14532d",
            bordercolor="#14532d",
        )
        style.configure("StatusCardInner.TFrame", background="#0b3328")
        style.configure(
            "StatusTitle.TLabel",
            background="#0b3328",
            foreground="#86efac",
            font=("Segoe UI Semibold", 8),
        )
        style.configure(
            "StatusValue.TLabel",
            background="#0b3328",
            foreground="#d1fae5",
            font=("Segoe UI Semibold", 10),
        )
        
        style.configure(
            "TButton",
            font=("Segoe UI Semibold", 10),
            padding=(12, 8),
            background=BG_CARD,
            foreground=TEXT_MAIN,
            borderwidth=1,
            relief="flat",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
        )
        style.map(
            "TButton",
            background=[("active", "#3f3f46"), ("pressed", "#18181b")],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            bordercolor=[("active", ACCENT_PRIMARY)],
        )
        
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(12, 8),
            foreground="#ffffff",
            background=ACCENT_PRIMARY,
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[("active", ACCENT_HOVER), ("pressed", ACCENT_PRESS)],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        
        style.configure(
            "Accent.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(12, 8),
            foreground="#ffffff",
            background=ACCENT_PRIMARY,
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_HOVER), ("pressed", ACCENT_PRESS)],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        
        style.configure(
            "Danger.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(12, 8),
            foreground="#ffffff",
            background=DANGER_PRIMARY,
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Danger.TButton",
            background=[("active", DANGER_HOVER), ("pressed", DANGER_PRESS)],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        
        style.configure(
            "TEntry",
            fieldbackground=BG_INPUT,
            foreground=TEXT_MAIN,
            insertcolor=ACCENT_PRIMARY,
            borderwidth=1,
            relief="flat",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
            padding=8,
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", ACCENT_PRIMARY)],
            lightcolor=[("focus", ACCENT_PRIMARY)],
            darkcolor=[("focus", ACCENT_PRIMARY)],
        )
        
        style.configure(
            "TCombobox",
            fieldbackground=BG_INPUT,
            foreground=TEXT_MAIN,
            background=BG_INPUT,
            arrowcolor=ACCENT_PRIMARY,
            borderwidth=1,
            relief="flat",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
            padding=6,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", BG_INPUT)],
            selectbackground=[("readonly", BG_INPUT)],
            selectforeground=[("readonly", TEXT_MAIN)],
            bordercolor=[("focus", ACCENT_PRIMARY)],
        )
        style.configure(
            "Inline.TRadiobutton",
            background=BG_CARD,
            foreground=TEXT_MAIN,
            font=("Segoe UI", 10),
            padding=(2, 0),
        )
        style.map(
            "Inline.TRadiobutton",
            foreground=[("selected", ACCENT_PRIMARY), ("active", "#e4e4e7")],
            background=[("active", BG_CARD)],
        )
        
        style.configure("TNotebook", background=BG_APP, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI Semibold", 10),
            padding=(13, 6),
            background=BG_APP,
            foreground=TEXT_SUB,
            borderwidth=1,
            relief="flat",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", BG_CARD), ("active", "#2f2f33"), ("!selected", BG_APP)],
            foreground=[("selected", TEXT_MAIN), ("active", "#e4e4e7"), ("!selected", TEXT_SUB)],
            relief=[("selected", "raised"), ("!selected", "flat")],
            bordercolor=[("selected", "#fafaf9"), ("active", BORDER), ("!selected", BORDER)],
            lightcolor=[("selected", "#fafaf9"), ("!selected", BORDER)],
            darkcolor=[("selected", "#fafaf9"), ("!selected", BORDER)],
            padding=[("selected", [17, 10]), ("!selected", [13, 6])],
            expand=[("selected", [4, 4, 4, 4]), ("!selected", [0, 0, 0, 0])],
        )
        
        style.configure(
            "Ops.Horizontal.TProgressbar",
            troughcolor=BG_CARD,
            background=ACCENT_PRIMARY,
            bordercolor=BORDER,
            lightcolor=ACCENT_PRIMARY,
            darkcolor=ACCENT_PRIMARY,
            thickness=6,
        )
        self._status_palette_ready = ["#34d399", "#10b981", "#059669", "#10b981"]
        self._status_palette_busy = ["#f59e0b", "#fbbf24", "#f59e0b", "#fcd34d"]

    def _show_splash(self):
        self.root.withdraw()
        splash = tk.Toplevel(self.root)
        splash.overrideredirect(True)
        splash.configure(bg="#18181b")
        splash.geometry("440x240")

        sw = splash.winfo_screenwidth()
        sh = splash.winfo_screenheight()
        splash.geometry(f"+{int(sw / 2 - 220)}+{int(sh / 2 - 120)}")

        tk.Label(
            splash,
            text="Practice + EMIS",
            fg="#10b981",
            bg="#18181b",
            font=("Segoe UI Semibold", 22),
        ).pack(pady=(48, 8))
        tk.Label(
            splash,
            text="Initializing workspace...",
            fg="#a1a1aa",
            bg="#18181b",
            font=("Segoe UI", 10),
        ).pack()

        splash_bar = ttk.Progressbar(
            splash,
            style="Ops.Horizontal.TProgressbar",
            orient="horizontal",
            length=320,
            mode="determinate",
        )
        splash_bar.pack(pady=32)

        def animate(step=0):
            if step <= 100:
                splash_bar["value"] = step
                splash.after(12, lambda: animate(step + 3))
                return

            splash.destroy()
            self._build_ui()
            self._setup_logging()
            self._load_pending_git_ops()
            self._check_admin()
            self._register_global_hotkeys()
            self._bind_local_hotkeys()
            self._start_dynamic_status()
            self._load_git_account_from_global()
            self._log_info("Unified tool ready.")
            self.root.deiconify()

        animate()

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Root.TFrame", padding=20)
        container.pack(fill="both", expand=True)

        header_frame = ttk.Frame(container, style="Root.TFrame")
        header_frame.pack(fill="x", pady=(0, 16))

        left_header = ttk.Frame(header_frame, style="Root.TFrame")
        left_header.pack(side="left", fill="x", expand=True)

        right_header = ttk.Frame(header_frame, style="Root.TFrame")
        right_header.pack(side="right", anchor="ne")

        ttk.Label(left_header, text="Practice and EMIS Operations", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            left_header,
            text="Onboarding, EMIS reset automation, Git push workflow, and Git account sync in one place.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        self.status_var = tk.StringVar(value="Ready")
        self.status_card = ttk.Frame(right_header, style="StatusCard.TFrame", padding=(10, 7))
        self.status_card.pack(anchor="e")

        status_top = ttk.Frame(self.status_card, style="StatusCardInner.TFrame")
        status_top.pack(anchor="e")
        self.status_dot_canvas = tk.Canvas(
            status_top,
            width=10,
            height=10,
            bg="#0b3328",
            highlightthickness=0,
            bd=0,
        )
        self.status_dot_canvas.pack(side="left")
        self.status_dot = self.status_dot_canvas.create_oval(1, 1, 9, 9, fill="#10b981", outline="")
        ttk.Label(status_top, text="SYSTEM", style="StatusTitle.TLabel").pack(side="left", padx=(6, 0))

        self.status_label = ttk.Label(self.status_card, textvariable=self.status_var, style="StatusValue.TLabel")
        self.status_label.pack(anchor="e", pady=(2, 0))

        self.global_progress = ttk.Progressbar(
            right_header,
            style="Ops.Horizontal.TProgressbar",
            mode="indeterminate",
            length=170,
        )
        self.global_progress.pack(anchor="e", pady=(6, 0))
        self.global_progress.pack_forget()

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        self.onboarding_tab = ttk.Frame(notebook, style="Root.TFrame", padding=16)
        self.git_sync_tab = ttk.Frame(notebook, style="Root.TFrame", padding=16)

        notebook.add(self.onboarding_tab, text="Onboarding")
        notebook.add(self.git_sync_tab, text="Git Account Sync")

        self._build_onboarding_tab()
        self._build_git_sync_tab()

    def _set_busy(self, label_text):
        clean_text = str(label_text).replace("System status:", "").strip() or "Loading"
        self._busy_count += 1
        self._busy_label = clean_text
        if self._busy_count == 1 and hasattr(self, "global_progress"):
            if not self.global_progress.winfo_ismapped():
                self.global_progress.pack(anchor="e", pady=(6, 0))
            self.global_progress.start(10)
        if hasattr(self, "status_var"):
            self.status_var.set(f"{clean_text}...")
        if hasattr(self, "status_label"):
            self.status_label.configure(foreground="#fde68a")
        if hasattr(self, "status_dot_canvas") and hasattr(self, "status_dot"):
            self.status_dot_canvas.itemconfigure(self.status_dot, fill="#f59e0b")

    def _set_idle(self):
        if self._busy_count > 0:
            self._busy_count -= 1
        if self._busy_count == 0:
            if hasattr(self, "global_progress"):
                self.global_progress.stop()
                self.global_progress.pack_forget()
            if hasattr(self, "status_label"):
                self.status_label.configure(foreground="#d1fae5")
            if hasattr(self, "status_dot_canvas") and hasattr(self, "status_dot"):
                self.status_dot_canvas.itemconfigure(self.status_dot, fill="#10b981")

    def _start_dynamic_status(self):
        self._status_phase = 0
        self._animate_status()

    def _animate_status(self):
        if not hasattr(self, "status_label"):
            return
        if self._busy_count > 0:
            dots = "." * ((self._status_phase % 3) + 1)
            busy_text = getattr(self, "_busy_label", "Loading")
            self.status_var.set(f"{busy_text}{dots}")
            if hasattr(self, "status_dot_canvas") and hasattr(self, "status_dot"):
                color = self._status_palette_busy[self._status_phase % len(self._status_palette_busy)]
                self.status_dot_canvas.itemconfigure(self.status_dot, fill=color)
            self._status_phase += 1
            self.root.after(650, self._animate_status)
            return
        dots = "." * ((self._status_phase % 3) + 1)
        self.status_var.set(f"Ready{dots}")
        if hasattr(self, "status_dot_canvas") and hasattr(self, "status_dot"):
            color = self._status_palette_ready[self._status_phase % len(self._status_palette_ready)]
            self.status_dot_canvas.itemconfigure(self.status_dot, fill=color)
        self._status_phase += 1
        self.root.after(650, self._animate_status)

    def _build_onboarding_tab(self):
        self.onboarding_tab.grid_columnconfigure(0, weight=1)
        self.onboarding_tab.grid_rowconfigure(0, weight=0)
        self.onboarding_tab.grid_rowconfigure(1, weight=4)

        top_row = ttk.Frame(self.onboarding_tab, style="Root.TFrame")
        top_row.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_row.grid_rowconfigure(0, weight=1)
        top_row.grid_columnconfigure(0, weight=10, minsize=330)
        top_row.grid_columnconfigure(1, weight=12, minsize=370)

        form = ttk.LabelFrame(top_row, text="Onboard Practice", style="Card.TLabelframe", padding=16)
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        form.grid_columnconfigure(0, weight=1)

        info_row = ttk.Frame(form, style="Card.TFrame")
        info_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        info_row.grid_columnconfigure(0, weight=0)
        info_row.grid_columnconfigure(1, weight=1)
        info_row.grid_columnconfigure(2, weight=0)
        info_row.grid_columnconfigure(3, weight=0)

        ttk.Label(info_row, text="Practice Name").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.entry_practice = ttk.Entry(info_row, width=26)
        self.entry_practice.grid(row=0, column=1, sticky="ew", padx=(0, 12))

        ttk.Label(info_row, text="ODS Code").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.entry_ods = ttk.Entry(info_row, width=8)
        self.entry_ods.grid(row=0, column=3, sticky="w")

        controls_row = ttk.Frame(form, style="Card.TFrame")
        controls_row.grid(row=1, column=0, sticky="ew")
        controls_row.grid_columnconfigure(0, weight=0)
        controls_row.grid_columnconfigure(1, weight=1)

        system_row = ttk.Frame(controls_row, style="Card.TFrame")
        system_row.grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(2, 0))
        ttk.Label(system_row, text="System").pack(side="left", padx=(0, 8))
        self.system_var = tk.StringVar(value="Docman")
        ttk.Radiobutton(
            system_row,
            text="Docman",
            variable=self.system_var,
            value="Docman",
            style="Inline.TRadiobutton",
        ).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(
            system_row,
            text="EMIS",
            variable=self.system_var,
            value="EMIS",
            style="Inline.TRadiobutton",
        ).pack(side="left")

        button_row = ttk.Frame(controls_row, style="Card.TFrame")
        button_row.grid(row=0, column=1, sticky="ew")
        for col in range(2):
            button_row.grid_columnconfigure(col, weight=1, uniform="onboard_actions")
        ttk.Button(button_row, text="Create Files", command=self.create_json_files).grid(
            row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 6)
        )
        ttk.Button(button_row, text="Validate ODS", command=self.run_validation_script).grid(
            row=0, column=1, sticky="ew", padx=(6, 0), pady=(0, 6)
        )
        ttk.Button(button_row, text="Git Push", command=self.open_git_push_window).grid(
            row=1, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(button_row, text="Offboard", style="Danger.TButton", command=self.offboard_practice).grid(
            row=1, column=1, sticky="ew", padx=(6, 0)
        )

        emis_panel = ttk.LabelFrame(top_row, text="EMIS Automation", style="Card.TLabelframe", padding=16)
        emis_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        emis_panel.grid_columnconfigure(0, weight=1)
        emis_panel.grid_columnconfigure(1, weight=1)

        if not AUTOMATION_READY:
            ttk.Label(emis_panel, text=f"Automation libraries are not available: {AUTOMATION_IMPORT_ERROR}").grid(
                row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
            )
            ttk.Label(emis_panel, text="Install missing packages to enable EMIS automation.").grid(
                row=1, column=0, columnspan=2, sticky="w"
            )
        else:
            ttk.Button(emis_panel, text="Auto Detect and Run", command=self.auto_detect_and_run).grid(
                row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8)
            )
            ttk.Button(emis_panel, text="Expired Reset", command=self.run_standard_automation).grid(
                row=1, column=0, sticky="ew", padx=(0, 6), pady=4
            )
            ttk.Button(emis_panel, text="Settings Reset", command=self.run_settings_automation).grid(
                row=1, column=1, sticky="ew", padx=(6, 0), pady=4
            )
            ttk.Button(emis_panel, text="Unlock", command=self.unlock_locked_screen).grid(
                row=2, column=0, sticky="ew", padx=(0, 6), pady=4
            )
            ttk.Button(emis_panel, text="Delayed Paste (3s)", command=self.delayed_paste).grid(
                row=2, column=1, sticky="ew", padx=(6, 0), pady=4
            )

        self.pwd_entry = ttk.Entry(emis_panel, justify="center")
        self.pwd_entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 6))
        self.pwd_entry.insert(0, "[None]")
        ttk.Button(emis_panel, text="Generate Password", command=self.generate_ui).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(0, 6)
        )
        ttk.Button(emis_panel, text="Open Password Log", command=self.open_log).grid(row=5, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(emis_panel, text="Clear Logs", command=self.clear_log).grid(row=5, column=1, sticky="ew", padx=(6, 0))

        log_box = ttk.LabelFrame(self.onboarding_tab, text="Live Log", style="Card.TLabelframe", padding=12)
        log_box.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        log_box.grid_columnconfigure(0, weight=1)
        log_box.grid_rowconfigure(0, weight=1)

        self.onboarding_log_widget = scrolledtext.ScrolledText(
            log_box,
            height=34,
            font=("Consolas", 9),
            bg="#09090b",
            fg="#a7f3d0",
            insertbackground="#10b981",
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#3f3f46",
            highlightcolor="#10b981",
            padx=8,
            pady=8,
        )
        self.onboarding_log_widget.grid(row=0, column=0, sticky="nsew")
        self.onboarding_log_widget.configure(state="disabled")
        self.log_widget = self.onboarding_log_widget

    def _build_emis_tab(self):
        panel = ttk.LabelFrame(self.emis_tab, text="EMIS Automation Moved", style="Card.TLabelframe", padding=20)
        panel.pack(fill="both", expand=True)
        ttk.Label(
            panel,
            text="EMIS controls and live log are now available in the Onboarding tab.",
            style="SubHeader.TLabel",
        ).pack(anchor="w")

    def _build_git_sync_tab(self):
        card = ttk.LabelFrame(self.git_sync_tab, text="Git Identity", style="Card.TLabelframe", padding=20)
        card.pack(fill="x", pady=(0, 16))

        ttk.Label(card, text="Name (user.name)").grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
        self.git_name_entry = ttk.Entry(card, width=60)
        self.git_name_entry.grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(card, text="Email (user.email)").grid(row=1, column=0, sticky="w", pady=6, padx=(0, 10))
        self.git_email_entry = ttk.Entry(card, width=60)
        self.git_email_entry.grid(row=1, column=1, sticky="w", pady=6)

        ttk.Label(card, text="Username (optional)").grid(row=2, column=0, sticky="w", pady=6, padx=(0, 10))
        self.git_username_entry = ttk.Entry(card, width=60)
        self.git_username_entry.grid(row=2, column=1, sticky="w", pady=6)

        ttk.Label(card, text="Credential helper").grid(row=3, column=0, sticky="w", pady=6, padx=(0, 10))
        self.git_helper_entry = ttk.Entry(card, width=60)
        self.git_helper_entry.grid(row=3, column=1, sticky="w", pady=6)

        button_row = ttk.Frame(card, style="Card.TFrame")
        button_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(16, 0))
        ttk.Button(button_row, text="Load from Global Git", command=self._load_git_account_from_global).pack(
            side="left", padx=(0, 10)
        )
        ttk.Button(button_row, text="Apply to Global Git", command=self.apply_git_account).pack(
            side="left"
        )

        sync_card = ttk.LabelFrame(self.git_sync_tab, text="Cross-Machine Sync", style="Card.TLabelframe", padding=20)
        sync_card.pack(fill="x")

        ttk.Label(
            sync_card,
            text="Export your Git identity to a JSON profile and import it on another machine.",
        ).pack(anchor="w", pady=(0, 10))

        sync_row = ttk.Frame(sync_card, style="Card.TFrame")
        sync_row.pack(anchor="w")
        ttk.Button(sync_row, text="Export Profile", command=self.export_git_profile).pack(side="left", padx=(0, 10))
        ttk.Button(sync_row, text="Import Profile", command=self.import_git_profile).pack(side="left", padx=(0, 10))
        ttk.Button(sync_row, text="Open .gitconfig", command=self.open_gitconfig).pack(side="left")

    def _setup_logging(self):
        self.emis_logger = logging.getLogger("emis_tool")
        self.emis_logger.setLevel(logging.DEBUG)
        self.emis_logger.propagate = False

        for handler in list(self.emis_logger.handlers):
            self.emis_logger.removeHandler(handler)

        formatter = logging.Formatter("%(relativeCreated)d INFO: %(message)s")
        file_handler = SafeFileHandler(DEBUG_LOG_PATH, encoding="utf-8")
        file_handler.setFormatter(formatter)

        gui_handler = TextHandler(self.log_widget)
        gui_handler.setFormatter(formatter)

        self.emis_logger.addHandler(file_handler)
        self.emis_logger.addHandler(gui_handler)

    def _log_info(self, message):
        if self.emis_logger:
            self.emis_logger.info(message)

    def _log_onboarding(self, message):
        if not hasattr(self, "onboarding_log_widget"):
            return

        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {message}"
        self.onboarding_log_widget.configure(state="normal")
        self.onboarding_log_widget.insert(tk.END, line + "\n")
        self.onboarding_log_widget.configure(state="disabled")
        self.onboarding_log_widget.see(tk.END)

    def _check_admin(self):
        try:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                self._log_info("Warning: app not running as admin.")
        except Exception:
            self._log_info("Admin check unavailable on this environment.")

    def _register_global_hotkeys(self):
        if self._hotkeys_registered:
            return
        if not hasattr(ctypes, "windll"):
            return
        if pyautogui is None:
            self._log_info("Global hotkeys disabled: pyautogui is not available.")
            return

        user32 = ctypes.windll.user32
        mod_control = 0x0002
        mod_shift = 0x0004
        vk_v = 0x56
        vk_c = 0x43
        wm_hotkey = 0x0312
        pm_remove = 0x0001

        try:
            if user32.RegisterHotKey(None, self._instant_paste_hotkey_id, mod_control | mod_shift, vk_v):
                self._registered_hotkey_ids.add(self._instant_paste_hotkey_id)
            else:
                self._log_info("Global hotkey Ctrl+Shift+V unavailable (already used by another app).")

            if user32.RegisterHotKey(None, self._instant_copy_hotkey_id, mod_control | mod_shift, vk_c):
                self._registered_hotkey_ids.add(self._instant_copy_hotkey_id)
            else:
                self._log_info("Global hotkey Ctrl+Shift+C unavailable (already used by another app).")

            if not self._registered_hotkey_ids:
                return

            self._hotkeys_registered = True
            self._hotkey_constants = {"WM_HOTKEY": wm_hotkey, "PM_REMOVE": pm_remove}
            self._hotkey_msg = wintypes.MSG()
            enabled_hotkeys = []
            if self._instant_paste_hotkey_id in self._registered_hotkey_ids:
                enabled_hotkeys.append("Ctrl+Shift+V")
            if self._instant_copy_hotkey_id in self._registered_hotkey_ids and pyperclip is not None:
                enabled_hotkeys.append("Ctrl+Shift+C")
            self._log_info(f"Global hotkeys enabled: {', '.join(enabled_hotkeys)}.")
            self.root.after(80, self._poll_hotkeys)
        except Exception as exc:
            self._log_info(f"Global hotkey registration failed: {exc}")

    def _bind_local_hotkeys(self):
        self.root.bind_all("<Control-Shift-V>", self._on_local_hotkey_paste)
        self.root.bind_all("<Control-Shift-C>", self._on_local_hotkey_copy)

    def _on_local_hotkey_paste(self, _event=None):
        self._instant_paste()
        return "break"

    def _on_local_hotkey_copy(self, _event=None):
        self._instant_copy()
        return "break"

    def _poll_hotkeys(self):
        if not self._hotkeys_registered or not hasattr(ctypes, "windll"):
            return

        user32 = ctypes.windll.user32
        wm_hotkey = self._hotkey_constants["WM_HOTKEY"]
        pm_remove = self._hotkey_constants["PM_REMOVE"]

        try:
            while user32.PeekMessageW(ctypes.byref(self._hotkey_msg), None, wm_hotkey, wm_hotkey, pm_remove):
                if self._hotkey_msg.message != wm_hotkey:
                    continue
                if self._hotkey_msg.wParam == self._instant_paste_hotkey_id:
                    self._instant_paste()
                elif self._hotkey_msg.wParam == self._instant_copy_hotkey_id:
                    self._instant_copy()
        except Exception as exc:
            self._log_info(f"Global hotkey polling error: {exc}")

        if self._hotkeys_registered and self.root.winfo_exists():
            self.root.after(80, self._poll_hotkeys)

    def _unregister_global_hotkeys(self):
        if not self._hotkeys_registered or not hasattr(ctypes, "windll"):
            return
        for hotkey_id in list(self._registered_hotkey_ids):
            try:
                ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
            except Exception:
                pass
        self._registered_hotkey_ids.clear()
        self._hotkeys_registered = False

    def _on_close(self):
        self._unregister_global_hotkeys()
        self.root.destroy()

    def _run_command(self, args, cwd=None):
        result = subprocess.run(args, cwd=cwd, capture_output=True, text=True, shell=False)
        return result

    def _write_json_file(self, path, payload):
        with open(path, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, indent=4)
            handle.write("\n")

    def _load_pending_git_ops(self):
        self.pending_git_ops = {}
        if not os.path.exists(PENDING_GIT_OPS_PATH):
            return
        try:
            with open(PENDING_GIT_OPS_PATH, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
            items = payload.get("items", []) if isinstance(payload, dict) else payload
            if not isinstance(items, list):
                return
            for item in items:
                if not isinstance(item, dict):
                    continue
                ods_code = str(item.get("ods_code", "")).strip().upper()
                action = str(item.get("action", "")).strip().lower()
                if not ods_code or action not in {"onboard", "offboard"}:
                    continue
                created_at = str(item.get("created_at", "")).strip() or datetime.now().isoformat(timespec="seconds")
                updated_at = str(item.get("updated_at", "")).strip() or created_at
                self.pending_git_ops[ods_code] = {
                    "action": action,
                    "ods_code": ods_code,
                    "practice_name": str(item.get("practice_name", "")).strip(),
                    "system_type": str(item.get("system_type", "")).strip() or "Docman",
                    "created_at": created_at,
                    "updated_at": updated_at,
                }
            if self.pending_git_ops:
                self._log_info(f"Loaded {len(self.pending_git_ops)} pending Git practice update(s).")
        except Exception as exc:
            self.pending_git_ops = {}
            self._log_info(f"Pending Git queue load failed: {exc}")

    def _save_pending_git_ops(self):
        items = sorted(
            self.pending_git_ops.values(),
            key=lambda item: (item.get("updated_at", ""), item.get("ods_code", "")),
        )
        payload = {"items": items}
        try:
            self._write_json_file(PENDING_GIT_OPS_PATH, payload)
        except Exception as exc:
            self._log_info(f"Pending Git queue save failed: {exc}")

    def _clear_pending_git_ops(self, reason=None):
        if not self.pending_git_ops:
            return
        self.pending_git_ops = {}
        self._save_pending_git_ops()
        if reason:
            self._log_onboarding(reason)

    def _record_pending_git_op(self, action, practice_name, ods_code, system_type="Docman"):
        normalized_action = str(action).strip().lower()
        normalized_ods = str(ods_code).strip().upper()
        if normalized_action not in {"onboard", "offboard"} or not normalized_ods:
            return

        now = datetime.now().isoformat(timespec="seconds")
        existing = self.pending_git_ops.get(normalized_ods)
        existing_practice = str(existing.get("practice_name", "")).strip() if existing else ""
        existing_system = str(existing.get("system_type", "Docman")).strip() if existing else "Docman"
        next_practice = str(practice_name or "").strip() or existing_practice
        next_system = str(system_type or "").strip() or existing_system or "Docman"

        if existing and existing.get("action") != normalized_action:
            self.pending_git_ops.pop(normalized_ods, None)
            self._save_pending_git_ops()
            self._log_onboarding(
                f"Pending Git update for {normalized_ods} cleared (onboard/offboard canceled before push)."
            )
            return

        self.pending_git_ops[normalized_ods] = {
            "action": normalized_action,
            "ods_code": normalized_ods,
            "practice_name": next_practice,
            "system_type": next_system,
            "created_at": existing.get("created_at", now) if existing else now,
            "updated_at": now,
        }
        self._save_pending_git_ops()
        self._log_onboarding(f"Pending Git update recorded: {normalized_action} {normalized_ods}.")

    def _onboard_artifact_exists(self, ods_code, practice_name):
        normalized_ods = str(ods_code).strip().upper()
        if not normalized_ods:
            return False
        folder_suffix = f"({normalized_ods})"
        explicit_folder = f"{practice_name.title()} ({normalized_ods})" if practice_name else ""

        for root_folder in ROOT_FOLDERS:
            if not os.path.isdir(root_folder):
                continue

            if explicit_folder:
                explicit_path = os.path.join(root_folder, explicit_folder, "work-items.json")
                if os.path.exists(explicit_path):
                    return True

            try:
                for entry in os.listdir(root_folder):
                    if not entry.endswith(folder_suffix):
                        continue
                    candidate = os.path.join(root_folder, entry, "work-items.json")
                    if os.path.exists(candidate):
                        return True
            except OSError:
                continue

        return False

    def _prune_pending_git_ops(self):
        removed = []
        for ods_code, item in list(self.pending_git_ops.items()):
            action = item.get("action")
            if action == "onboard" and not self._onboard_artifact_exists(ods_code, item.get("practice_name", "")):
                removed.append(ods_code)
                self.pending_git_ops.pop(ods_code, None)

        if removed:
            self._save_pending_git_ops()
            self._log_onboarding(
                f"Removed stale onboarding item(s) from pending Git queue: {', '.join(sorted(removed))}."
            )

    def _summarize_ods_codes(self, items, limit=3):
        codes = [str(item.get("ods_code", "")).strip().upper() for item in items if item.get("ods_code")]
        codes = [code for code in codes if code]
        if not codes:
            return "none"
        if len(codes) <= limit:
            return ", ".join(codes)
        return f"{', '.join(codes[:limit])}, +{len(codes) - limit} more"

    def _slugify_branch_part(self, text, max_len=42):
        slug = re.sub(r"[^a-z0-9]+", "-", str(text).lower()).strip("-")
        if not slug:
            return "update"
        return slug[:max_len].rstrip("-")

    def _suggest_git_push_defaults(self):
        self._prune_pending_git_ops()
        pending_items = sorted(
            self.pending_git_ops.values(),
            key=lambda item: (item.get("updated_at", ""), item.get("ods_code", "")),
        )
        now = datetime.now()
        earliest_stamp = now
        for item in pending_items:
            stamp = str(item.get("created_at") or item.get("updated_at") or "").strip()
            if not stamp:
                continue
            try:
                parsed = datetime.fromisoformat(stamp)
            except Exception:
                continue
            if parsed < earliest_stamp:
                earliest_stamp = parsed
        date_tag = earliest_stamp.strftime("%Y%m%d")

        if not pending_items:
            return (
                f"practice-update/{date_tag}-manual",
                f"Practice updates ({now.strftime('%Y-%m-%d')})",
                "No tracked onboarding/offboarding items pending. Using manual defaults.",
            )

        onboard_items = [item for item in pending_items if item.get("action") == "onboard"]
        offboard_items = [item for item in pending_items if item.get("action") == "offboard"]

        if onboard_items and not offboard_items:
            if len(onboard_items) == 1:
                item = onboard_items[0]
                branch_name = f"onboard/{date_tag}-{self._slugify_branch_part(item.get('ods_code', 'practice'))}"
                commit_message = f"Onboard: {item.get('practice_name') or 'Practice'} ({item.get('ods_code')})"
            else:
                branch_name = f"onboard/{date_tag}-{len(onboard_items)}-practices"
                commit_message = f"Onboard {len(onboard_items)} practices: {self._summarize_ods_codes(onboard_items)}"
        elif offboard_items and not onboard_items:
            if len(offboard_items) == 1:
                item = offboard_items[0]
                branch_name = f"offboard/{date_tag}-{self._slugify_branch_part(item.get('ods_code', 'practice'))}"
                commit_message = f"Offboard: {item.get('practice_name') or 'Practice'} ({item.get('ods_code')})"
            else:
                branch_name = f"offboard/{date_tag}-{len(offboard_items)}-practices"
                commit_message = (
                    f"Offboard {len(offboard_items)} practices: {self._summarize_ods_codes(offboard_items)}"
                )
        else:
            branch_name = f"practice-update/{date_tag}-on{len(onboard_items)}-off{len(offboard_items)}"
            commit_message = (
                f"Practice updates: onboard {len(onboard_items)} "
                f"({self._summarize_ods_codes(onboard_items)}); offboard {len(offboard_items)} "
                f"({self._summarize_ods_codes(offboard_items)})"
            )

        summary = (
            f"Pending tracked items: onboard {len(onboard_items)} ({self._summarize_ods_codes(onboard_items)}), "
            f"offboard {len(offboard_items)} ({self._summarize_ods_codes(offboard_items)})."
        )
        return branch_name, commit_message, summary

    def _run_git(self, git_args):
        return self._run_command(["git"] + git_args, cwd=GIT_REPO_PATH)

    def _run_git_checked(self, git_args):
        result = self._run_git(git_args)
        if result.returncode != 0:
            cmd = "git " + " ".join(git_args)
            raise RuntimeError(f"{cmd} failed\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}")
        return result

    def _git_repo_ready(self):
        return os.path.isdir(GIT_REPO_PATH) and os.path.isdir(os.path.join(GIT_REPO_PATH, ".git"))

    def _suggest_branch_base(self):
        branch = "main"
        if not self._git_repo_ready():
            return branch
        try:
            result = self._run_git(["symbolic-ref", "refs/remotes/origin/HEAD"])
            if result.returncode == 0 and result.stdout.strip():
                branch = result.stdout.strip().split("/")[-1]
            else:
                fallback = self._run_git(["branch", "--list", "main"])
                if not fallback.stdout.strip():
                    branch = "master"
        except Exception:
            pass
        return branch

    def run_validation_script(self):
        self._set_busy("System status: Validating ODS")
        try:
            self._log_onboarding("Validate ODS clicked.")
            outputs = []
            if not os.path.exists(CHECK_ODS_MISMATCH_SCRIPT):
                self._log_onboarding(f"Validation script missing: {CHECK_ODS_MISMATCH_SCRIPT}")
                messagebox.showerror("Validation", f"Script not found: {CHECK_ODS_MISMATCH_SCRIPT}")
                return

            for folder in ROOT_FOLDERS:
                if not os.path.isdir(folder):
                    outputs.append(f"Skipped missing folder: {folder}")
                    continue

                result = self._run_command(
                    [
                        "powershell",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        CHECK_ODS_MISMATCH_SCRIPT,
                        "-BasePath",
                        folder,
                    ]
                )

                if result.stdout.strip():
                    outputs.append(f"[{os.path.basename(folder)}]\n{result.stdout.strip()}")
                if result.stderr.strip():
                    outputs.append(f"Errors in {os.path.basename(folder)}:\n{result.stderr.strip()}")

            final = "\n\n".join(outputs) if outputs else "No issues detected."
            self._log_onboarding("ODS validation completed.")
            messagebox.showinfo("ODS Mismatch Check", final)
        finally:
            self._set_idle()

    def _validate_current_creation(self, practice_name, ods, system_type):
        folder_name = f"{practice_name.title()} ({ods})"
        check_notes = []
        check_passed = 0
        check_issues = 0

        for root_folder in ROOT_FOLDERS:
            if not os.path.isdir(root_folder):
                check_notes.append(f"Skipped missing root folder: {root_folder}")
                continue

            root_label = os.path.basename(root_folder)
            practice_file = os.path.join(root_folder, folder_name, "work-items.json")
            if not os.path.exists(practice_file):
                check_issues += 1
                check_notes.append(f"[{root_label}] Missing practice file: {practice_file}")
                continue

            try:
                with open(practice_file, "r", encoding="utf-8") as handle:
                    practice_data = json.load(handle)
                file_ods = str(practice_data[0].get("payload", {}).get("ods_code", "")).upper()
            except Exception as exc:
                check_issues += 1
                check_notes.append(f"[{root_label}] Invalid practice JSON: {exc}")
                continue

            if file_ods == ods:
                check_passed += 1
            else:
                check_issues += 1
                check_notes.append(f"[{root_label}] ODS mismatch in practice file (found: {file_ods or '[None]'})")

            if system_type != "Docman":
                continue

            count_path = os.path.join(root_folder, "Practice Count", "work-items.json")
            if not os.path.exists(count_path):
                check_issues += 1
                check_notes.append(f"[{root_label}] Missing Practice Count file: {count_path}")
                continue

            try:
                with open(count_path, "r", encoding="utf-8") as handle:
                    count_data = json.load(handle)
            except Exception as exc:
                check_issues += 1
                check_notes.append(f"[{root_label}] Invalid Practice Count JSON: {exc}")
                continue

            found = any(str(item.get("payload", {}).get("ods_code", "")).upper() == ods for item in count_data)
            if found:
                check_passed += 1
            else:
                check_issues += 1
                check_notes.append(f"[{root_label}] ODS {ods} not found in Practice Count.")

        if system_type != "Docman":
            check_notes.append("Practice Count check skipped for EMIS mode.")

        return check_passed, check_issues, check_notes

    def create_json_files(self):
        self._set_busy("System status: Creating practice files")
        try:
            self._log_onboarding("Create Files clicked.")
            system_type = self.system_var.get().strip() or "Docman"
            practice_name = self.entry_practice.get().strip()
            ods = self.entry_ods.get().strip().upper()

            if not practice_name or not ods:
                self._log_onboarding("Create failed: Practice Name or ODS missing.")
                messagebox.showerror("Error", "Please enter both Practice Name and ODS Code.")
                return

            self.last_practices.append((practice_name, ods))

            folders_created = 0
            practice_files_written = 0
            counts_updated = 0
            notes = []

            for root_folder in ROOT_FOLDERS:
                if not os.path.isdir(root_folder):
                    notes.append(f"Skipped missing root folder: {root_folder}")
                    continue

                try:
                    folder_name = f"{practice_name.title()} ({ods})"
                    practice_folder = os.path.join(root_folder, folder_name)
                    practice_file = os.path.join(practice_folder, "work-items.json")

                    if not os.path.exists(practice_folder):
                        os.makedirs(practice_folder, exist_ok=True)
                        folders_created += 1
                        notes.append(f"Created folder: {practice_folder}")
                    elif not os.path.isdir(practice_folder):
                        raise RuntimeError(f"A file exists with the folder name: {practice_folder}")

                    self._write_json_file(practice_file, [{"payload": {"ods_code": ods}}])
                    practice_files_written += 1

                    if system_type != "Docman":
                        notes.append(f"{root_folder}: skipped Practice Count update for EMIS mode")
                        continue

                    count_path = os.path.join(root_folder, "Practice Count", "work-items.json")
                    data = []
                    if os.path.exists(count_path):
                        try:
                            with open(count_path, "r", encoding="utf-8") as handle:
                                data = json.load(handle)
                        except Exception:
                            data = []
                            notes.append(f"Reset invalid JSON in {count_path}")

                    exists = any(str(item.get("payload", {}).get("ods_code", "")).upper() == ods for item in data)
                    if not exists:
                        data.append(
                            {
                                "payload": {
                                    "ods_code": ods,
                                    "docman_practice_display_name": practice_name.upper(),
                                }
                            }
                        )
                        os.makedirs(os.path.dirname(count_path), exist_ok=True)
                        self._write_json_file(count_path, data)
                        counts_updated += 1
                    else:
                        notes.append(f"ODS {ods} already exists in Practice Count at {root_folder}")

                except Exception as exc:
                    self._log_onboarding(f"Create failed in {root_folder}: {exc}")
                    messagebox.showerror("Create Files", f"Failed in {root_folder}: {exc}")
                    return

            summary = [
                f"Onboarding Summary for {practice_name} ({system_type})",
                f"Folders created: {folders_created}",
                f"Practice Count updates: {counts_updated}",
            ]
            if notes:
                summary.append("")
                summary.append("Details:")
                summary.extend(notes)

            checks_ok, checks_failed, check_notes = self._validate_current_creation(practice_name, ods, system_type)
            summary.append("")
            summary.append("Current Creation Check:")
            if checks_failed == 0:
                summary.append(f"Passed ({checks_ok} checks).")
            else:
                summary.append(f"Completed with {checks_failed} issue(s) ({checks_ok} checks passed).")
            if check_notes:
                summary.extend(check_notes)

            self._log_onboarding(
                f"Create completed for {practice_name} ({ods}). Folders: {folders_created}, Count updates: {counts_updated}"
            )
            if practice_files_written > 0:
                self._record_pending_git_op("onboard", practice_name, ods, system_type)
            if checks_failed == 0:
                self._log_onboarding(f"Current creation validation passed for ODS {ods}.")
            else:
                self._log_onboarding(f"Current creation validation found {checks_failed} issue(s) for ODS {ods}.")
            messagebox.showinfo("Status", "\n".join(summary))
        finally:
            self._set_idle()

    def offboard_practice(self):
        self._log_onboarding("Offboard clicked.")
        ods_code = self.entry_ods.get().strip().upper()
        if not ods_code:
            self._log_onboarding("Offboard failed: ODS code missing.")
            messagebox.showerror("Offboard", "Enter ODS Code in the onboarding form first.")
            return

        removed = 0
        for root_folder in ROOT_FOLDERS:
            count_path = os.path.join(root_folder, "Practice Count", "work-items.json")
            if not os.path.exists(count_path):
                continue

            try:
                with open(count_path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)

                original_len = len(data)
                data = [
                    item
                    for item in data
                    if str(item.get("payload", {}).get("ods_code", "")).upper() != ods_code
                ]
                if len(data) < original_len:
                    self._write_json_file(count_path, data)
                    removed += 1
            except Exception as exc:
                self._log_onboarding(f"Offboard failed updating {count_path}: {exc}")
                messagebox.showerror("Offboard", f"Failed to update {count_path}: {exc}")
                return

        if removed:
            self._record_pending_git_op("offboard", self.entry_practice.get().strip(), ods_code, self.system_var.get())
            self._log_onboarding(f"Offboard success: removed {ods_code} from {removed} file(s).")
            messagebox.showinfo("Offboard", f"Removed ODS {ods_code} from {removed} Practice Count file(s).")
        else:
            self._log_onboarding(f"Offboard completed: {ods_code} not found.")
            messagebox.showinfo("Offboard", f"ODS {ods_code} was not found in Practice Count files.")

    def open_git_push_window(self):
        self._log_onboarding("Git Push window opened.")
        window = tk.Toplevel(self.root)
        window.title("Git Push")
        window.geometry("560x340")
        window.configure(bg="#18181b")
        window.resizable(False, False)

        ttk.Label(window, text="Branch Name", background="#18181b").pack(anchor="w", padx=20, pady=(20, 6))
        branch_entry = ttk.Entry(window, width=60)
        branch_entry.pack(padx=20)

        suggested_branch, suggested_commit, pending_summary = self._suggest_git_push_defaults()
        branch_entry.insert(0, suggested_branch)

        ttk.Label(window, text="Commit Message", background="#18181b").pack(anchor="w", padx=20, pady=(14, 6))
        commit_entry = ttk.Entry(window, width=60)
        commit_entry.pack(padx=20)
        commit_entry.insert(0, suggested_commit)

        tk.Label(
            window,
            text=pending_summary,
            fg="#a1a1aa",
            bg="#18181b",
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=515,
        ).pack(fill="x", padx=20, pady=(10, 0))

        push_confirm = tk.BooleanVar(value=True)
        # Fix styling slightly for the checkbutton
        cb = ttk.Checkbutton(window, text="Push to origin after commit", variable=push_confirm)
        cb.pack(anchor="w", padx=20, pady=(14, 0))

        def handle_push():
            try:
                self._log_onboarding("Git push flow started.")
                result = self.run_git_push(branch_entry.get().strip(), commit_entry.get().strip(), push_confirm.get())
                self._log_onboarding(result)
                messagebox.showinfo("Git Push", result)
                window.destroy()
            except Exception as exc:
                self._log_onboarding(f"Git push failed: {exc}")
                messagebox.showerror("Git Push", str(exc))

        ttk.Button(window, text="Run Git Flow", style="Accent.TButton", command=handle_push).pack(pady=20)

    def run_git_push(self, branch_name, commit_message, push_to_origin=True):
        self._set_busy("System status: Running git flow")
        try:
            self._log_onboarding(f"Preparing git flow for branch '{branch_name}'.")
            if not branch_name:
                raise RuntimeError("Branch name cannot be empty.")
            if not commit_message:
                raise RuntimeError("Commit message cannot be empty.")
            if not self._git_repo_ready():
                raise RuntimeError(f"Git repo not found at: {GIT_REPO_PATH}")

            default_branch = self._suggest_branch_base()
            self._run_git_checked(["fetch", "origin"])
            self._run_git_checked(["checkout", default_branch])
            self._run_git_checked(["pull", "--ff-only", "origin", default_branch])
            self._log_onboarding(f"Base branch '{default_branch}' updated.")

            branch_exists = self._run_git(["show-ref", "--verify", f"refs/heads/{branch_name}"]).returncode == 0
            if branch_exists:
                self._run_git_checked(["checkout", branch_name])
            else:
                self._run_git_checked(["checkout", "-b", branch_name])

            self._run_git_checked(["add", "-A"])
            status = self._run_git(["status", "--porcelain"])
            if not status.stdout.strip():
                self._log_onboarding("No git changes detected.")
                self._clear_pending_git_ops("Pending Git queue cleared because no working-tree changes were detected.")
                return "No changes detected. Nothing to commit."

            self._run_git_checked(["commit", "-m", commit_message])
            self._clear_pending_git_ops("Committed changes. Pending Git queue cleared.")

            if push_to_origin:
                self._run_git_checked(["push", "-u", "origin", branch_name])
                return f"Commit and push completed on branch '{branch_name}'."

            return f"Commit created locally on branch '{branch_name}'. Push was skipped."
        finally:
            self._set_idle()

    def _read_git_global(self, key):
        result = self._run_command(["git", "config", "--global", "--get", key])
        if result.returncode == 0:
            return result.stdout.strip()
        return ""

    def _set_git_global(self, key, value):
        result = self._run_command(["git", "config", "--global", key, value])
        if result.returncode != 0:
            cmd = f"git config --global {key} {value}"
            raise RuntimeError(f"{cmd} failed\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}")

    def _load_git_account_from_global(self):
        self.git_name_entry.delete(0, tk.END)
        self.git_name_entry.insert(0, self._read_git_global("user.name"))

        self.git_email_entry.delete(0, tk.END)
        self.git_email_entry.insert(0, self._read_git_global("user.email"))

        self.git_username_entry.delete(0, tk.END)
        username = self._read_git_global("credential.username")
        self.git_username_entry.insert(0, username)

        self.git_helper_entry.delete(0, tk.END)
        helper = self._read_git_global("credential.helper")
        self.git_helper_entry.insert(0, helper or "manager-core")

    def apply_git_account(self):
        self._set_busy("System status: Updating Git identity")
        try:
            name = self.git_name_entry.get().strip()
            email = self.git_email_entry.get().strip()
            username = self.git_username_entry.get().strip()
            helper = self.git_helper_entry.get().strip() or "manager-core"

            if not name or not email:
                messagebox.showerror("Git Account", "Name and Email are required.")
                return

            try:
                self._set_git_global("user.name", name)
                self._set_git_global("user.email", email)
                self._set_git_global("credential.helper", helper)
                if username:
                    self._set_git_global("credential.username", username)
                messagebox.showinfo("Git Account", "Global Git identity was updated successfully.")
            except Exception as exc:
                messagebox.showerror("Git Account", str(exc))
        finally:
            self._set_idle()

    def export_git_profile(self):
        profile = {
            "user.name": self.git_name_entry.get().strip(),
            "user.email": self.git_email_entry.get().strip(),
            "credential.username": self.git_username_entry.get().strip(),
            "credential.helper": self.git_helper_entry.get().strip() or "manager-core",
        }

        default_path = os.path.join(os.path.expanduser("~"), GIT_PROFILE_FILE)
        path = filedialog.asksaveasfilename(
            title="Export Git Profile",
            defaultextension=".json",
            initialfile=GIT_PROFILE_FILE,
            initialdir=os.path.dirname(default_path),
            filetypes=[("JSON file", "*.json")],
        )
        if not path:
            return

        try:
            self._write_json_file(path, profile)
            messagebox.showinfo("Git Account", f"Profile exported to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Git Account", f"Failed to export profile: {exc}")

    def import_git_profile(self):
        path = filedialog.askopenfilename(
            title="Import Git Profile",
            filetypes=[("JSON file", "*.json")],
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as handle:
                profile = json.load(handle)
        except Exception as exc:
            messagebox.showerror("Git Account", f"Failed to read profile: {exc}")
            return

        self.git_name_entry.delete(0, tk.END)
        self.git_name_entry.insert(0, str(profile.get("user.name", "")))

        self.git_email_entry.delete(0, tk.END)
        self.git_email_entry.insert(0, str(profile.get("user.email", "")))

        self.git_username_entry.delete(0, tk.END)
        self.git_username_entry.insert(0, str(profile.get("credential.username", "")))

        self.git_helper_entry.delete(0, tk.END)
        self.git_helper_entry.insert(0, str(profile.get("credential.helper", "manager-core")))

        if messagebox.askyesno("Git Account", "Profile loaded. Apply it to global Git config now?"):
            self.apply_git_account()

    def open_gitconfig(self):
        gitconfig = os.path.join(os.path.expanduser("~"), ".gitconfig")
        if os.path.exists(gitconfig):
            os.startfile(gitconfig)
        else:
            messagebox.showinfo("Git Account", f"No .gitconfig found at {gitconfig}")

    def log_password(self, password, context):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        entry = f"{timestamp} [{context}] - {password}\n"
        try:
            with open(LOG_FILE_PATH, "a", encoding="utf-8") as handle:
                handle.write(entry)
                handle.flush()
                os.fsync(handle.fileno())
            self._log_info(f"Password recorded for {context}.")
        except Exception as exc:
            self._log_info(f"Password save error: {exc}")

    def generate_ui(self):
        self.generated_pwd = generate_strict_password(10)
        self.pwd_entry.delete(0, tk.END)
        self.pwd_entry.insert(0, self.generated_pwd)
        if pyperclip:
            pyperclip.copy(self.generated_pwd)
        self._log_info(f"Generated password: {self.generated_pwd}")
        return self.generated_pwd

    def run_settings_automation(self):
        if not AUTOMATION_READY:
            self._log_info("Automation unavailable due to missing libraries.")
            return

        password = self.generate_ui()
        self._log_info("Searching for EMIS Edit User modal...")

        try:
            desktop = Desktop(backend="uia")
            main_emis = desktop.window(title_re=".*EMIS Web Health Care System.*")
            wizard = main_emis.child_window(auto_id="UserWizardForm", control_type="Window")

            if not wizard.exists(timeout=5):
                self._log_info("Wizard not found. Open Edit User first.")
                return

            wizard.set_focus()
            container = wizard.child_window(auto_id="loginPanel", control_type="Pane")
            pass_field = container.child_window(auto_id="passwordTextBox", control_type="Edit")
            confirm_field = container.child_window(auto_id="confirmPasswordTextBox", control_type="Edit")

            pass_field.set_focus()
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            pyautogui.write(password, interval=0.01)

            confirm_field.set_focus()
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            pyautogui.write(password, interval=0.01)

            self.log_password(password, "Settings")
            self._log_info("Settings reset fields were populated.")
        except Exception as exc:
            self._log_info(f"Settings automation failure: {exc}")

    def run_standard_automation(self):
        if not AUTOMATION_READY:
            self._log_info("Automation unavailable due to missing libraries.")
            return

        password = self.generate_ui()
        self._log_info("Standard reset started.")
        windows = gw.getWindowsWithTitle("Authentication")
        target = next((window for window in windows if 500 < window.width < 700), None)

        if not target:
            self._log_info("Authentication window not found.")
            return

        target.activate()
        time.sleep(0.5)

        pyautogui.click(target.left + 350, target.top + 215)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        pyautogui.write(password, interval=0.01)
        pyautogui.press("tab")
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        pyautogui.write(password, interval=0.01)
        pyautogui.press("enter")

        self.log_password(password, "Expired")
        self._log_info("Standard reset fields were populated.")

    def unlock_locked_screen(self):
        if not AUTOMATION_READY:
            self._log_info("Automation unavailable due to missing libraries.")
            return

        self._log_info("Unlock flow started.")
        password = pyperclip.paste()

        try:
            app = Application(backend="uia").connect(title_re=".*Locked.*", timeout=5)
            dialog = app.window(title_re=".*Locked.*")
            dialog.set_focus()
            dialog.child_window(auto_id="textBoxPassword", control_type="Edit").set_focus()
            pyautogui.write(password, interval=0.01)
            dialog.child_window(auto_id="buttonUnlock", control_type="Button").click()
            self.log_password(password, "Unlock")
            self._log_info("Unlock action sent.")
        except Exception as exc:
            self._log_info(f"Unlock failure: {exc}")

    def delayed_paste(self):
        if pyautogui is None or pyperclip is None:
            self._log_info("Delayed paste unavailable: pyautogui/pyperclip not available.")
            return

        self._log_info("Pasting in 3 seconds.")
        self.root.after(3000, lambda: pyautogui.write(pyperclip.paste(), interval=0.01))

    def _instant_paste(self):
        if pyautogui is None:
            self._log_info("Instant paste unavailable: pyautogui not available.")
            return
        try:
            # Run slightly later so the trigger modifiers are released.
            self.root.after(80, self._perform_instant_paste)
        except Exception as exc:
            self._log_info(f"Instant paste failed: {exc}")

    def _perform_instant_paste(self):
        try:
            pyautogui.hotkey("ctrl", "v")
            self._log_info("Instant paste triggered (Ctrl+Shift+V).")
        except Exception as exc:
            self._log_info(f"Instant paste failed: {exc}")

    def _instant_copy(self):
        if pyautogui is None or pyperclip is None:
            self._log_info("Instant copy unavailable: pyautogui/pyperclip not available.")
            return

        self.root.after(80, self._perform_instant_copy)

    def _perform_instant_copy(self):
        try:
            previous = pyperclip.paste()
        except Exception:
            previous = ""

        latest = previous
        try:
            for _ in range(4):
                pyautogui.hotkey("ctrl", "c")
                time.sleep(0.09)
                latest = pyperclip.paste()
                if latest != previous:
                    break
                pyautogui.hotkey("ctrl", "insert")
                time.sleep(0.09)
                latest = pyperclip.paste()
                if latest != previous:
                    break
            pyperclip.copy(latest)
            if latest:
                self._log_info("Instant copy triggered (Ctrl+Shift+C). Clipboard updated.")
            else:
                self._log_info("Instant copy triggered (Ctrl+Shift+C). Clipboard is empty.")
        except Exception as exc:
            self._log_info(f"Instant copy failed: {exc}")

    def auto_detect_and_run(self):
        if not AUTOMATION_READY:
            self._log_info("Automation unavailable due to missing libraries.")
            return

        titles = [window.title for window in gw.getAllWindows()]
        if any("Locked" in title for title in titles):
            self.unlock_locked_screen()
        elif any("Edit user" in title for title in titles):
            self.run_settings_automation()
        elif any("Authentication" in title for title in titles):
            self.run_standard_automation()
        else:
            self._log_info("No target EMIS window detected.")

    def open_log(self):
        if os.path.exists(LOG_FILE_PATH):
            os.startfile(LOG_FILE_PATH)

    def clear_log(self):
        for path in [LOG_FILE_PATH, DEBUG_LOG_PATH]:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("")
        self._log_info("Logs cleared.")


if __name__ == "__main__":
    tk_root = tk.Tk()
    app = UnifiedToolApp(tk_root)
    tk_root.mainloop()
