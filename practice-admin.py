import ctypes
import json
import logging
import os
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

# ── Default paths (used when no saved config exists) ─────────────────────────
_DEFAULT_PROJECT_BASE = r"C:\rpa\postie"
CHECK_ODS_MISMATCH_SCRIPT = os.path.join(SCRIPT_DIR, "check-ods-mismatch.ps1")
GIT_PROFILE_FILE = "git-account-profile.json"
PATHS_CONFIG_FILE = os.path.join(SCRIPT_DIR, "paths-config.json")


def _load_paths_config():
    """Return (project_base, root_folders, git_repo_path) from saved config or defaults."""
    if os.path.exists(PATHS_CONFIG_FILE):
        try:
            with open(PATHS_CONFIG_FILE, "r", encoding="utf-8") as fh:
                cfg = json.load(fh)
            base = cfg.get("project_base", _DEFAULT_PROJECT_BASE)
            root_folders = cfg.get("root_folders") or [
                os.path.join(base, "postie_bots_python", "devdata", "work-items-in"),
                os.path.join(base, "postie-bots", "devdata", "work-items-in"),
            ]
            git_repo = cfg.get("git_repo_path") or os.path.join(base, "postie-bots")
            return base, root_folders, git_repo
        except Exception:
            pass
    base = _DEFAULT_PROJECT_BASE
    return (
        base,
        [
            os.path.join(base, "postie_bots_python", "devdata", "work-items-in"),
            os.path.join(base, "postie-bots", "devdata", "work-items-in"),
        ],
        os.path.join(base, "postie-bots"),
    )


# Module-level defaults (overridden per-instance via self.* at runtime)
PROJECT_BASE, ROOT_FOLDERS, GIT_REPO_PATH = _load_paths_config()


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
        self.root.title("Practice + EMIS Operations")
        self.root.geometry("780x680")
        self.root.minsize(740, 600)

        self.generated_pwd = ""
        self.last_practices = []
        self.emis_logger = None

        # Load path config into instance variables so UI can update them live
        self._project_base, self._root_folders, self._git_repo_path = _load_paths_config()

        self._setup_styles()
        self._build_ui()
        self._setup_logging()
        self._check_admin()

        self._load_git_account_from_global()
        self._log_info("Unified tool ready.")

    def _setup_styles(self):
        # ── Colour palette ──────────────────────────────────────────────────
        self.C = {
            "bg":        "#0d1117",   # window background
            "surface":   "#161b22",   # card / frame background
            "surface2":  "#1c2128",   # card header / alternate row
            "border":    "#30363d",   # default border
            "accent":    "#3fb950",   # green (success / admin)
            "blue":      "#58a6ff",   # blue (primary action)
            "blue_dark": "#1f6feb",   # active/hover blue
            "warn":      "#d29922",   # amber warning
            "text":      "#e6edf3",   # primary text
            "text2":     "#8b949e",   # secondary / label text
            "text3":     "#484f58",   # muted / placeholder
            "log_bg":    "#010409",   # terminal bg
            "log_fg":    "#3fb950",   # terminal text
        }
        C = self.C

        self.root.configure(bg=C["bg"])
        style = ttk.Style()
        style.theme_use("clam")

        # Frames
        style.configure("Root.TFrame",    background=C["bg"])
        style.configure("Surface.TFrame", background=C["surface"])
        style.configure("TFrame",         background=C["bg"])

        # LabelFrames (cards)
        style.configure(
            "Card.TLabelframe",
            background=C["surface"],
            bordercolor=C["border"],
            borderwidth=1,
            relief="solid",
            padding=2,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=C["surface2"],
            foreground=C["text2"],
            font=("Consolas", 9, "bold"),
            padding=(6, 4),
        )

        # Labels
        style.configure("TLabel",
            background=C["bg"], foreground=C["text"], font=("Segoe UI", 10))
        style.configure("Header.TLabel",
            background=C["bg"], foreground=C["text"], font=("Segoe UI", 15, "bold"))
        style.configure("SubHeader.TLabel",
            background=C["bg"], foreground=C["text2"], font=("Segoe UI", 9))
        style.configure("Card.TLabel",
            background=C["surface"], foreground=C["text2"], font=("Segoe UI", 10))
        style.configure("CardMono.TLabel",
            background=C["surface"], foreground=C["text2"], font=("Consolas", 9))
        style.configure("Status.TLabel",
            background=C["bg"], foreground=C["accent"], font=("Consolas", 9, "bold"))
        style.configure("StatusWarn.TLabel",
            background=C["bg"], foreground=C["warn"], font=("Consolas", 9, "bold"))

        # Entries
        style.configure("TEntry",
            fieldbackground=C["bg"],
            foreground=C["text"],
            insertcolor=C["blue"],
            bordercolor=C["border"],
            lightcolor=C["border"],
            darkcolor=C["border"],
            font=("Consolas", 10),
            padding=(6, 5),
        )
        style.map("TEntry",
            bordercolor=[("focus", C["blue"])],
            lightcolor=[("focus", C["blue"])],
        )

        # Combobox
        style.configure("TCombobox",
            fieldbackground=C["bg"],
            background=C["surface2"],
            foreground=C["text"],
            arrowcolor=C["text2"],
            bordercolor=C["border"],
            font=("Consolas", 10),
            padding=(6, 5),
        )
        style.map("TCombobox",
            fieldbackground=[("readonly", C["bg"])],
            bordercolor=[("focus", C["blue"])],
        )

        # Buttons – default (ghost style)
        style.configure("TButton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            background=C["surface2"],
            foreground=C["text2"],
            bordercolor=C["border"],
            relief="flat",
        )
        style.map("TButton",
            background=[("active", C["surface"]), ("pressed", C["surface"])],
            foreground=[("active", C["text"])],
            bordercolor=[("active", C["border"])],
        )

        # Accent button – blue filled
        style.configure("Accent.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(12, 7),
            background=C["blue"],
            foreground="#0d1117",
            bordercolor=C["blue"],
            relief="flat",
        )
        style.map("Accent.TButton",
            background=[("active", C["blue_dark"]), ("pressed", C["blue_dark"])],
            foreground=[("active", "#ffffff")],
            bordercolor=[("active", C["blue_dark"])],
        )

        # Danger button – red outline
        style.configure("Danger.TButton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            background=C["surface2"],
            foreground="#f85149",
            bordercolor="#6e2720",
            relief="flat",
        )
        style.map("Danger.TButton",
            background=[("active", "#2d0f0e"), ("pressed", "#2d0f0e")],
            foreground=[("active", "#ff7b72")],
        )

        # Notebook
        style.configure("TNotebook",
            background=C["bg"],
            bordercolor=C["border"],
            tabmargins=(0, 0, 0, 0),
        )
        style.configure("TNotebook.Tab",
            font=("Consolas", 9, "bold"),
            padding=(16, 8),
            background=C["surface"],
            foreground=C["text2"],
            bordercolor=C["border"],
        )
        style.map("TNotebook.Tab",
            background=[("selected", C["surface2"])],
            foreground=[("selected", C["blue"])],
        )

        # Separator
        style.configure("TSeparator", background=C["border"])

    def _build_ui(self):
        C = self.C
        container = ttk.Frame(self.root, style="Root.TFrame", padding=(16, 14, 16, 10))
        container.pack(fill="both", expand=True)

        # ── Title bar ────────────────────────────────────────────────────────
        title_row = ttk.Frame(container, style="Root.TFrame")
        title_row.pack(fill="x", pady=(0, 2))

        ttk.Label(title_row, text="Practice & EMIS Operations",
                  style="Header.TLabel").pack(side="left", anchor="w")

        self._admin_status_label = ttk.Label(title_row, text="● CHECKING…",
                                              style="StatusWarn.TLabel")
        self._admin_status_label.pack(side="right", anchor="e", padx=(0, 4))

        ttk.Label(
            container,
            text="Onboarding · EMIS automation · Git push workflow · Account sync",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        ttk.Separator(container, orient="horizontal").pack(fill="x", pady=(0, 10))

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        # Single operational tab + git tab
        self.onboarding_tab = ttk.Frame(notebook, style="Root.TFrame", padding=(12, 10))
        self.emis_tab        = self.onboarding_tab   # same frame — merged
        self.git_sync_tab    = ttk.Frame(notebook, style="Root.TFrame", padding=(12, 10))

        notebook.add(self.onboarding_tab, text="  Operations  ")
        notebook.add(self.git_sync_tab,   text="  Git Account Sync  ")

        self._build_onboarding_tab()
        self._build_git_sync_tab()

    def _build_onboarding_tab(self):
        """Merged Operations tab: Onboard + Offboard side-by-side,
           EMIS controls row, shared log at bottom."""
        C = self.C
        tab = self.onboarding_tab
        lbl_kw = dict(style="CardMono.TLabel", anchor="e", padding=(0, 0, 10, 0))

        # ── Row 1: Onboard (left) + Offboard (right) side by side ────────────
        top_row = ttk.Frame(tab, style="Root.TFrame")
        top_row.pack(fill="x", pady=(0, 8))
        top_row.columnconfigure(0, weight=3)
        top_row.columnconfigure(1, weight=2)

        # -- Onboard card (left)
        form = ttk.LabelFrame(top_row, text=" + ONBOARD PRACTICE",
                              style="Card.TLabelframe", padding=(12, 8))
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="practice name", **lbl_kw).grid(row=0, column=0, sticky="e", pady=5)
        self.entry_practice = ttk.Entry(form, font=("Consolas", 10))
        self.entry_practice.grid(row=0, column=1, sticky="ew", pady=5)

        ttk.Label(form, text="ods code", **lbl_kw).grid(row=1, column=0, sticky="e", pady=5)
        self.entry_ods = ttk.Entry(form, font=("Consolas", 10))
        self.entry_ods.grid(row=1, column=1, sticky="ew", pady=5)

        ttk.Label(form, text="system", **lbl_kw).grid(row=2, column=0, sticky="e", pady=5)
        self.system_var = tk.StringVar(value="Docman")
        ttk.Combobox(form, textvariable=self.system_var,
                     values=["Docman", "EMIS"], state="readonly",
                     font=("Consolas", 10)).grid(row=2, column=1, sticky="ew", pady=5)

        tk.Frame(form, bg=C["border"], height=1).grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(8, 6))

        btn_row = ttk.Frame(form, style="Surface.TFrame")
        btn_row.grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Button(btn_row, text="Create Files", style="Accent.TButton",
                   command=self.create_json_files).pack(side="left", padx=(0, 5))
        ttk.Button(btn_row, text="Validate ODS",
                   command=self.run_validation_script).pack(side="left", padx=(0, 5))
        ttk.Button(btn_row, text="Git Push",
                   command=self.open_git_push_window).pack(side="left")

        # -- Offboard card (right)
        offboard = ttk.LabelFrame(top_row, text=" x OFFBOARD PRACTICE",
                                  style="Card.TLabelframe", padding=(12, 8))
        offboard.grid(row=0, column=1, sticky="nsew")
        offboard.columnconfigure(0, weight=1)

        ttk.Label(offboard, text="ods code", style="CardMono.TLabel",
                  anchor="w").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.entry_offboard_ods = ttk.Entry(offboard, font=("Consolas", 10))
        self.entry_offboard_ods.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        tk.Frame(offboard, bg=C["border"], height=1).grid(
            row=2, column=0, sticky="ew", pady=(0, 8))

        ttk.Button(offboard, text="Remove Practice", style="Danger.TButton",
                   command=self.offboard_practice).grid(row=3, column=0, sticky="w")

        # ── Row 2: EMIS automation controls ───────────────────────────────────
        emis_row = ttk.Frame(tab, style="Root.TFrame")
        emis_row.pack(fill="x", pady=(0, 8))
        emis_row.columnconfigure(0, weight=1)
        emis_row.columnconfigure(1, weight=1)

        # -- Automation controls card (left)
        controls = ttk.LabelFrame(emis_row, text=" EMIS AUTOMATION",
                                  style="Card.TLabelframe", padding=(12, 8))
        controls.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        if not AUTOMATION_READY:
            banner = tk.Frame(controls, bg="#2d2008", highlightbackground="#9e6a03",
                              highlightthickness=1)
            banner.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
            tk.Label(banner, text="Automation libraries unavailable — install pyautogui, pygetwindow, pyperclip, pywinauto",
                     bg="#2d2008", fg="#d29922", font=("Consolas", 8),
                     anchor="w", pady=5, padx=8).pack(fill="x")
            controls.grid_columnconfigure(0, weight=1)
        else:
            ttk.Button(controls, text="Auto Detect and Run",
                       style="Accent.TButton",
                       command=self.auto_detect_and_run).grid(
                row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
            ttk.Button(controls, text="Expired Reset",
                       command=self.run_standard_automation).grid(
                row=1, column=0, sticky="ew", padx=(0, 4), pady=3)
            ttk.Button(controls, text="Settings Reset",
                       command=self.run_settings_automation).grid(
                row=1, column=1, sticky="ew", padx=(4, 0), pady=3)
            ttk.Button(controls, text="Unlock",
                       command=self.unlock_locked_screen).grid(
                row=2, column=0, sticky="ew", padx=(0, 4), pady=3)
            ttk.Button(controls, text="Delayed Paste (3s)",
                       command=self.delayed_paste).grid(
                row=2, column=1, sticky="ew", padx=(4, 0), pady=3)
            controls.grid_columnconfigure(0, weight=1)
            controls.grid_columnconfigure(1, weight=1)

        # -- Password card (right)
        quick = ttk.LabelFrame(emis_row, text=" PASSWORD & LOGS",
                               style="Card.TLabelframe", padding=(12, 8))
        quick.grid(row=0, column=1, sticky="nsew")
        quick.columnconfigure(0, weight=1)
        quick.columnconfigure(1, weight=1)

        self.pwd_entry = ttk.Entry(quick, font=("Consolas", 13))
        self.pwd_entry.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        self.pwd_entry.insert(0, "— not generated —")

        ttk.Button(quick, text="Generate Password", style="Accent.TButton",
                   command=self.generate_ui).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(quick, text="Open Password Log",
                   command=self.open_log).grid(row=2, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(quick, text="Clear Logs",
                   command=self.clear_log).grid(row=2, column=1, sticky="ew", padx=(4, 0))

        # ── Row 3: Shared live log (fills remaining space) ────────────────────
        log_box = ttk.LabelFrame(tab, text=" LIVE LOG",
                                 style="Card.TLabelframe", padding=(6, 4))
        log_box.pack(fill="both", expand=True)

        # Single shared log widget — both onboarding and EMIS write here
        shared_log = scrolledtext.ScrolledText(
            log_box,
            font=("Consolas", 9),
            bg=C["log_bg"], fg=C["log_fg"],
            insertbackground=C["log_fg"],
            relief="flat", bd=0,
            selectbackground=C["surface2"],
        )
        shared_log.pack(fill="both", expand=True)
        shared_log.configure(state="disabled")

        # Both logger references point to the same widget
        self.onboarding_log_widget = shared_log
        self.log_widget = shared_log


    def _build_git_sync_tab(self):
        C = self.C
        lbl_kw = dict(style="CardMono.TLabel", anchor="e", padding=(0, 0, 10, 0))

        # -- Identity card
        card = ttk.LabelFrame(self.git_sync_tab, text=" GIT IDENTITY",
                              style="Card.TLabelframe", padding=(14, 10))
        card.pack(fill="x", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="user.name", **lbl_kw).grid(row=0, column=0, sticky="e", pady=6)
        self.git_name_entry = ttk.Entry(card, font=("Consolas", 10))
        self.git_name_entry.grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(card, text="user.email", **lbl_kw).grid(row=1, column=0, sticky="e", pady=6)
        self.git_email_entry = ttk.Entry(card, font=("Consolas", 10))
        self.git_email_entry.grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(card, text="credential.username", **lbl_kw).grid(row=2, column=0, sticky="e", pady=6)
        self.git_username_entry = ttk.Entry(card, font=("Consolas", 10))
        self.git_username_entry.grid(row=2, column=1, sticky="ew", pady=6)

        ttk.Label(card, text="credential.helper", **lbl_kw).grid(row=3, column=0, sticky="e", pady=6)
        self.git_helper_entry = ttk.Entry(card, font=("Consolas", 10))
        self.git_helper_entry.grid(row=3, column=1, sticky="ew", pady=6)

        tk.Frame(card, bg=C["border"], height=1).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(10, 8))

        button_row = ttk.Frame(card, style="Surface.TFrame")
        button_row.grid(row=5, column=0, columnspan=2, sticky="w")
        ttk.Button(button_row, text="Load from Global Git",
                   command=self._load_git_account_from_global).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="Apply to Global Git", style="Accent.TButton",
                   command=self.apply_git_account).pack(side="left")

        # -- Cross-machine sync card
        sync_card = ttk.LabelFrame(self.git_sync_tab, text=" CROSS-MACHINE SYNC",
                                   style="Card.TLabelframe", padding=(14, 10))
        sync_card.pack(fill="x")

        info_frame = tk.Frame(sync_card, bg="#0d2040", highlightbackground="#1f4068",
                              highlightthickness=1)
        info_frame.pack(fill="x", pady=(0, 10))
        tk.Label(info_frame,
                 text="Export your Git identity to a JSON profile and import it on another machine.",
                 bg="#0d2040", fg="#58a6ff",
                 font=("Segoe UI", 9), anchor="w",
                 pady=7, padx=10).pack(fill="x")

        sync_row = ttk.Frame(sync_card, style="Surface.TFrame")
        sync_row.pack(anchor="w")
        ttk.Button(sync_row, text="Export Profile",
                   command=self.export_git_profile).pack(side="left", padx=(0, 6))
        ttk.Button(sync_row, text="Import Profile",
                   command=self.import_git_profile).pack(side="left", padx=(0, 6))
        ttk.Button(sync_row, text="Open .gitconfig",
                   command=self.open_gitconfig).pack(side="left")

        # -- Paths card
        paths_card = ttk.LabelFrame(self.git_sync_tab, text=" PROJECT PATHS",
                                    style="Card.TLabelframe", padding=(14, 10))
        paths_card.pack(fill="x", pady=(10, 0))
        paths_card.columnconfigure(1, weight=1)

        lbl_kw = dict(style="CardMono.TLabel", anchor="e", padding=(0, 0, 10, 0))

        # Git repo path
        ttk.Label(paths_card, text="git repo path", **lbl_kw).grid(
            row=0, column=0, sticky="e", pady=5)
        repo_frame = ttk.Frame(paths_card, style="Surface.TFrame")
        repo_frame.grid(row=0, column=1, sticky="ew", pady=5)
        repo_frame.columnconfigure(0, weight=1)
        self._git_repo_entry = ttk.Entry(repo_frame, font=("Consolas", 10))
        self._git_repo_entry.insert(0, self._git_repo_path)
        self._git_repo_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(repo_frame, text="Browse",
                   command=lambda: self._browse_dir(self._git_repo_entry)).grid(row=0, column=1)

        # Project base (auto-fills root folders)
        ttk.Label(paths_card, text="project base", **lbl_kw).grid(
            row=1, column=0, sticky="e", pady=5)
        base_frame = ttk.Frame(paths_card, style="Surface.TFrame")
        base_frame.grid(row=1, column=1, sticky="ew", pady=5)
        base_frame.columnconfigure(0, weight=1)
        self._project_base_entry = ttk.Entry(base_frame, font=("Consolas", 10))
        self._project_base_entry.insert(0, self._project_base)
        self._project_base_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(base_frame, text="Browse",
                   command=lambda: self._browse_dir(self._project_base_entry)).grid(row=0, column=1)

        # Work-items root folders (shown read-only, derived from project base)
        ttk.Label(paths_card, text="work-items roots", **lbl_kw).grid(
            row=2, column=0, sticky="ne", pady=5)
        self._root_folders_text = tk.Text(
            paths_card, height=2, font=("Consolas", 9),
            bg=self.C["log_bg"], fg=self.C["text2"],
            insertbackground=self.C["blue"],
            relief="flat", bd=1, highlightthickness=1,
            highlightbackground=self.C["border"],
            highlightcolor=self.C["blue"],
        )
        self._root_folders_text.grid(row=2, column=1, sticky="ew", pady=5)
        for rf in self._root_folders:
            self._root_folders_text.insert(tk.END, rf + "\n")

        info = tk.Frame(paths_card, bg="#0d2040", highlightbackground="#1f4068",
                        highlightthickness=1)
        info.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 8))
        tk.Label(info,
                 text="  Tip: changing Project Base auto-derives the two work-items-in folders when you save.",
                 bg="#0d2040", fg="#58a6ff", font=("Segoe UI", 8),
                 anchor="w", pady=5, padx=4).pack(fill="x")

        tk.Frame(paths_card, bg=self.C["border"], height=1).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        paths_btn_row = ttk.Frame(paths_card, style="Surface.TFrame")
        paths_btn_row.grid(row=5, column=0, columnspan=2, sticky="w")
        ttk.Button(paths_btn_row, text="Save Paths", style="Accent.TButton",
                   command=self._save_paths_config).pack(side="left", padx=(0, 6))
        ttk.Button(paths_btn_row, text="Reset to Defaults",
                   command=self._reset_paths_to_defaults).pack(side="left")

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
            if ctypes.windll.shell32.IsUserAnAdmin():
                self._log_info("Running as administrator.")
                if hasattr(self, "_admin_status_label"):
                    self._admin_status_label.config(text="● ADMIN", style="Status.TLabel")
            else:
                self._log_info("Warning: app not running as admin.")
                if hasattr(self, "_admin_status_label"):
                    self._admin_status_label.config(text="● NOT ADMIN", style="StatusWarn.TLabel")
        except Exception:
            self._log_info("Admin check unavailable on this environment.")
            if hasattr(self, "_admin_status_label"):
                self._admin_status_label.config(text="● ADMIN N/A", style="StatusWarn.TLabel")

    def _run_command(self, args, cwd=None):
        result = subprocess.run(args, cwd=cwd, capture_output=True, text=True, shell=False)
        return result

    def _run_git(self, git_args):
        return self._run_command(["git"] + git_args, cwd=self._git_repo_path)

    def _run_git_checked(self, git_args):
        result = self._run_git(git_args)
        if result.returncode != 0:
            cmd = "git " + " ".join(git_args)
            raise RuntimeError(f"{cmd} failed\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}")
        return result

    def _git_repo_ready(self):
        p = self._git_repo_path
        return os.path.isdir(p) and os.path.isdir(os.path.join(p, ".git"))

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
        self._log_onboarding("Validate ODS clicked.")
        outputs = []
        if not os.path.exists(CHECK_ODS_MISMATCH_SCRIPT):
            self._log_onboarding(f"Validation script missing: {CHECK_ODS_MISMATCH_SCRIPT}")
            messagebox.showerror("Validation", f"Script not found: {CHECK_ODS_MISMATCH_SCRIPT}")
            return

        for folder in self._root_folders:
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

    def _validate_current_creation(self, practice_name, ods, system_type):
        folder_name = f"{practice_name.title()} ({ods})"
        check_notes = []
        check_passed = 0
        check_issues = 0

        for root_folder in self._root_folders:
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
                first_item = practice_data[0] if isinstance(practice_data, list) and practice_data else {}
                file_ods = str(first_item.get("payload", {}).get("ods_code", "")).upper()
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
        counts_updated = 0
        notes = []

        for root_folder in self._root_folders:
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

                with open(practice_file, "w", encoding="utf-8") as handle:
                    json.dump([{"payload": {"ods_code": ods}}], handle, indent=4)
                    handle.write("\n")

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
                    with open(count_path, "w", encoding="utf-8") as handle:
                        json.dump(data, handle, indent=4)
                        handle.write("\n")
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
        if checks_failed == 0:
            self._log_onboarding(f"Current creation validation passed for ODS {ods}.")
        else:
            self._log_onboarding(f"Current creation validation found {checks_failed} issue(s) for ODS {ods}.")
        messagebox.showinfo("Status", "\n".join(summary))

    def offboard_practice(self):
        self._log_onboarding("Offboard clicked.")
        ods_code = self.entry_offboard_ods.get().strip().upper()
        if not ods_code:
            self._log_onboarding("Offboard failed: ODS code missing.")
            messagebox.showerror("Offboard", "Enter the ODS code to remove.")
            return

        removed = 0
        for root_folder in self._root_folders:
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
                    with open(count_path, "w", encoding="utf-8") as handle:
                        json.dump(data, handle, indent=4)
                        handle.write("\n")
                    removed += 1
            except Exception as exc:
                self._log_onboarding(f"Offboard failed updating {count_path}: {exc}")
                messagebox.showerror("Offboard", f"Failed to update {count_path}: {exc}")
                return

        if removed:
            self._log_onboarding(f"Offboard success: removed {ods_code} from {removed} file(s).")
            messagebox.showinfo("Offboard", f"Removed ODS {ods_code} from {removed} Practice Count file(s).")
        else:
            self._log_onboarding(f"Offboard completed: {ods_code} not found.")
            messagebox.showinfo("Offboard", f"ODS {ods_code} was not found in Practice Count files.")

    def open_git_push_window(self):
        self._log_onboarding("Git Push window opened.")
        window = tk.Toplevel(self.root)
        window.title("Git Push")
        window.geometry("440x250")
        window.resizable(False, False)

        ttk.Label(window, text="Branch Name").pack(anchor="w", padx=16, pady=(16, 4))
        branch_entry = ttk.Entry(window, width=58)
        branch_entry.pack(padx=16)

        last_name, last_ods = (
            self.last_practices[-1] if self.last_practices else ("practice", datetime.now().strftime("%Y%m%d"))
        )
        branch_entry.insert(0, f"onboard/{last_ods}".lower().replace(" ", "-"))

        ttk.Label(window, text="Commit Message").pack(anchor="w", padx=16, pady=(12, 4))
        commit_entry = ttk.Entry(window, width=58)
        commit_entry.pack(padx=16)
        commit_entry.insert(0, f"Onboarded: {last_name} ({last_ods})")

        push_confirm = tk.BooleanVar(value=True)
        ttk.Checkbutton(window, text="Push to origin after commit", variable=push_confirm).pack(
            anchor="w", padx=16, pady=(12, 0)
        )

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

        ttk.Button(window, text="Run Git Flow", style="Accent.TButton", command=handle_push).pack(pady=18)

    def run_git_push(self, branch_name, commit_message, push_to_origin=True):
        self._log_onboarding(f"Preparing git flow for branch '{branch_name}'.")
        if not branch_name:
            raise RuntimeError("Branch name cannot be empty.")
        if not commit_message:
            raise RuntimeError("Commit message cannot be empty.")
        if not self._git_repo_ready():
            raise RuntimeError(
                f"Git repo not found at: {self._git_repo_path}\n\n"
                f"Go to Git Account Sync tab -> Paths to set the correct repo path."
            )

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
            return "No changes detected. Nothing to commit."

        self._run_git_checked(["commit", "-m", commit_message])

        if push_to_origin:
            self._run_git_checked(["push", "-u", "origin", branch_name])
            return f"Commit and push completed on branch '{branch_name}'."

        return f"Commit created locally on branch '{branch_name}'. Push was skipped."

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
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(profile, handle, indent=4)
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
        if not AUTOMATION_READY:
            self._log_info("Automation unavailable due to missing libraries.")
            return

        self._log_info("Pasting in 3 seconds.")
        self.root.after(3000, lambda: pyautogui.write(pyperclip.paste(), interval=0.01))

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

    # ── Path config helpers ──────────────────────────────────────────────────

    def _browse_dir(self, entry_widget):
        """Open a folder-picker and put the result in entry_widget."""
        current = entry_widget.get().strip()
        initial = current if os.path.isdir(current) else os.path.expanduser("~")
        chosen = filedialog.askdirectory(title="Select folder", initialdir=initial)
        if chosen:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, os.path.normpath(chosen))

    def _save_paths_config(self):
        git_repo = self._git_repo_entry.get().strip()
        base = self._project_base_entry.get().strip()
        if not git_repo or not base:
            messagebox.showerror("Paths", "Git repo path and project base cannot be empty.")
            return

        root_folders = [
            os.path.join(base, "postie_bots_python", "devdata", "work-items-in"),
            os.path.join(base, "postie-bots", "devdata", "work-items-in"),
        ]

        cfg = {
            "project_base": base,
            "git_repo_path": git_repo,
            "root_folders": root_folders,
        }
        try:
            with open(PATHS_CONFIG_FILE, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh, indent=4)
        except Exception as exc:
            messagebox.showerror("Paths", f"Could not save config: {exc}")
            return

        # Apply live to instance variables
        self._project_base = base
        self._git_repo_path = git_repo
        self._root_folders = root_folders

        # Refresh the read-only root folders display
        self._root_folders_text.delete("1.0", tk.END)
        for rf in root_folders:
            self._root_folders_text.insert(tk.END, rf + "\n")

        self._log_info(f"Paths saved. Git repo: {git_repo}")
        messagebox.showinfo("Paths", f"Paths saved and applied.\n\nGit repo: {git_repo}\nProject base: {base}")

    def _reset_paths_to_defaults(self):
        if not messagebox.askyesno("Paths", "Reset all paths to the original defaults?"):
            return
        if os.path.exists(PATHS_CONFIG_FILE):
            os.remove(PATHS_CONFIG_FILE)
        base = _DEFAULT_PROJECT_BASE
        self._project_base = base
        self._git_repo_path = os.path.join(base, "postie-bots")
        self._root_folders = [
            os.path.join(base, "postie_bots_python", "devdata", "work-items-in"),
            os.path.join(base, "postie-bots", "devdata", "work-items-in"),
        ]
        self._project_base_entry.delete(0, tk.END)
        self._project_base_entry.insert(0, self._project_base)
        self._git_repo_entry.delete(0, tk.END)
        self._git_repo_entry.insert(0, self._git_repo_path)
        self._root_folders_text.delete("1.0", tk.END)
        for rf in self._root_folders:
            self._root_folders_text.insert(tk.END, rf + "\n")
        self._log_info("Paths reset to defaults.")
        messagebox.showinfo("Paths", "Paths reset to defaults.")

    # ── Log helpers ──────────────────────────────────────────────────────────

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
