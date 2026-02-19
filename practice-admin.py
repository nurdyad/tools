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
        self.root.geometry("760x660")
        self.root.minsize(720, 580)

        self.generated_pwd = ""
        self.last_practices = []
        self.emis_logger = None

        self._setup_styles()
        self._build_ui()
        self._setup_logging()
        self._check_admin()

        self._load_git_account_from_global()
        self._log_info("Unified tool ready.")

    def _setup_styles(self):
        self.root.configure(bg="#eaf4ff")
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Root.TFrame", background="#eaf4ff")
        style.configure("Card.TLabelframe", background="#ffffff", borderwidth=1, relief="solid")
        style.configure(
            "Card.TLabelframe.Label",
            background="#ffffff",
            foreground="#0d3b66",
            font=("Segoe UI", 11, "bold"),
        )
        style.configure("TLabel", background="#eaf4ff", foreground="#1f2937", font=("Segoe UI", 10))
        style.configure("Header.TLabel", background="#eaf4ff", foreground="#0f172a", font=("Segoe UI", 16, "bold"))
        style.configure("SubHeader.TLabel", background="#eaf4ff", foreground="#385170", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=6, background="#d7e8ff")
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8,
            foreground="#ffffff",
            background="#0ea5e9",
        )
        style.map("Accent.TButton", background=[("active", "#0284c7")])
        style.configure("TNotebook", background="#eaf4ff", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=(10, 7),
            background="#cfe4ff",
            foreground="#0f172a",
        )
        style.map("TNotebook.Tab", background=[("selected", "#0ea5e9")], foreground=[("selected", "#ffffff")])

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Root.TFrame", padding=18)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Practice and EMIS Operations", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="Onboarding, EMIS reset automation, Git push workflow, and Git account sync in one place.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(0, 12))

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        self.onboarding_tab = ttk.Frame(notebook, style="Root.TFrame", padding=12)
        self.emis_tab = ttk.Frame(notebook, style="Root.TFrame", padding=12)
        self.git_sync_tab = ttk.Frame(notebook, style="Root.TFrame", padding=12)

        notebook.add(self.onboarding_tab, text="Onboarding")
        notebook.add(self.emis_tab, text="EMIS Automation")
        notebook.add(self.git_sync_tab, text="Git Account Sync")

        self._build_onboarding_tab()
        self._build_emis_tab()
        self._build_git_sync_tab()

    def _build_onboarding_tab(self):
        form = ttk.LabelFrame(self.onboarding_tab, text="Onboard Practice", style="Card.TLabelframe", padding=16)
        form.pack(fill="x", pady=(0, 12))

        ttk.Label(form, text="Practice Name").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_practice = ttk.Entry(form, width=46)
        self.entry_practice.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(form, text="ODS Code").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_ods = ttk.Entry(form, width=46)
        self.entry_ods.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Practice System").grid(row=2, column=0, sticky="w", pady=4)
        self.system_var = tk.StringVar(value="Docman")
        system_combo = ttk.Combobox(
            form,
            textvariable=self.system_var,
            values=["Docman", "EMIS"],
            state="readonly",
            width=43,
        )
        system_combo.grid(row=2, column=1, sticky="w", pady=4)

        button_row = ttk.Frame(form, style="Root.TFrame")
        button_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(button_row, text="Create Files", style="Accent.TButton", command=self.create_json_files).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(button_row, text="Validate ODS", command=self.run_validation_script).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Git Push", command=self.open_git_push_window).pack(side="left")

        offboard = ttk.LabelFrame(self.onboarding_tab, text="Offboard Practice", style="Card.TLabelframe", padding=16)
        offboard.pack(fill="x")

        ttk.Label(offboard, text="ODS Code to Remove").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_offboard_ods = ttk.Entry(offboard, width=46)
        self.entry_offboard_ods.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Button(offboard, text="Offboard", command=self.offboard_practice).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

        log_box = ttk.LabelFrame(self.onboarding_tab, text="Onboarding Live Log", style="Card.TLabelframe", padding=12)
        log_box.pack(fill="both", expand=True, pady=(12, 0))

        self.onboarding_log_widget = scrolledtext.ScrolledText(
            log_box, height=10, font=("Consolas", 9), bg="#0b1f33", fg="#9ef7b8"
        )
        self.onboarding_log_widget.pack(fill="both", expand=True)
        self.onboarding_log_widget.configure(state="disabled")

    def _build_emis_tab(self):
        top_row = ttk.Frame(self.emis_tab, style="Root.TFrame")
        top_row.pack(fill="x", pady=(0, 12))

        controls = ttk.LabelFrame(top_row, text="Automation Controls", style="Card.TLabelframe", padding=14)
        controls.pack(side="left", fill="both", expand=True, padx=(0, 6))

        quick = ttk.LabelFrame(top_row, text="Password and Logs", style="Card.TLabelframe", padding=14)
        quick.pack(side="left", fill="both", expand=True, padx=(6, 0))

        if not AUTOMATION_READY:
            ttk.Label(controls, text=f"Automation libraries are not available: {AUTOMATION_IMPORT_ERROR}").grid(
                row=0, column=0, sticky="w", pady=(0, 6)
            )
            ttk.Label(controls, text="Install missing packages to enable EMIS automation.").grid(
                row=1, column=0, sticky="w"
            )
        else:
            ttk.Button(
                controls,
                text="Auto Detect and Run",
                style="Accent.TButton",
                command=self.auto_detect_and_run,
            ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
            ttk.Button(controls, text="Expired Reset", command=self.run_standard_automation).grid(
                row=1, column=0, sticky="ew", padx=(0, 6), pady=4
            )
            ttk.Button(controls, text="Settings Reset", command=self.run_settings_automation).grid(
                row=1, column=1, sticky="ew", padx=(6, 0), pady=4
            )
            ttk.Button(controls, text="Unlock", command=self.unlock_locked_screen).grid(
                row=2, column=0, sticky="ew", padx=(0, 6), pady=4
            )
            ttk.Button(controls, text="Delayed Paste (3s)", command=self.delayed_paste).grid(
                row=2, column=1, sticky="ew", padx=(6, 0), pady=4
            )
            controls.grid_columnconfigure(0, weight=1)
            controls.grid_columnconfigure(1, weight=1)

        self.pwd_entry = ttk.Entry(quick, width=36)
        self.pwd_entry.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.pwd_entry.insert(0, "[None]")

        ttk.Button(quick, text="Generate Password", style="Accent.TButton", command=self.generate_ui).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8)
        )
        ttk.Button(quick, text="Open Password Log", command=self.open_log).grid(
            row=2, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(quick, text="Clear Logs", command=self.clear_log).grid(
            row=2, column=1, sticky="ew", padx=(6, 0)
        )
        quick.grid_columnconfigure(0, weight=1)
        quick.grid_columnconfigure(1, weight=1)

        log_box = ttk.LabelFrame(self.emis_tab, text="Live Log", style="Card.TLabelframe", padding=10)
        log_box.pack(fill="both", expand=True)

        self.log_widget = scrolledtext.ScrolledText(
            log_box, height=18, font=("Consolas", 9), bg="#0f172a", fg="#9ef7b8"
        )
        self.log_widget.pack(fill="both", expand=True)
        self.log_widget.configure(state="disabled")

    def _build_git_sync_tab(self):
        card = ttk.LabelFrame(self.git_sync_tab, text="Git Identity", style="Card.TLabelframe", padding=16)
        card.pack(fill="x", pady=(0, 12))

        ttk.Label(card, text="Name (user.name)").grid(row=0, column=0, sticky="w", pady=4)
        self.git_name_entry = ttk.Entry(card, width=58)
        self.git_name_entry.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(card, text="Email (user.email)").grid(row=1, column=0, sticky="w", pady=4)
        self.git_email_entry = ttk.Entry(card, width=58)
        self.git_email_entry.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(card, text="Username (optional)").grid(row=2, column=0, sticky="w", pady=4)
        self.git_username_entry = ttk.Entry(card, width=58)
        self.git_username_entry.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(card, text="Credential helper").grid(row=3, column=0, sticky="w", pady=4)
        self.git_helper_entry = ttk.Entry(card, width=58)
        self.git_helper_entry.grid(row=3, column=1, sticky="w", pady=4)

        button_row = ttk.Frame(card, style="Root.TFrame")
        button_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(button_row, text="Load from Global Git", command=self._load_git_account_from_global).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(button_row, text="Apply to Global Git", style="Accent.TButton", command=self.apply_git_account).pack(
            side="left"
        )

        sync_card = ttk.LabelFrame(self.git_sync_tab, text="Cross-Machine Sync", style="Card.TLabelframe", padding=16)
        sync_card.pack(fill="x")

        ttk.Label(
            sync_card,
            text="Export your Git identity to a JSON profile and import it on another machine.",
        ).pack(anchor="w", pady=(0, 8))

        sync_row = ttk.Frame(sync_card, style="Root.TFrame")
        sync_row.pack(anchor="w")
        ttk.Button(sync_row, text="Export Profile", command=self.export_git_profile).pack(side="left", padx=(0, 8))
        ttk.Button(sync_row, text="Import Profile", command=self.import_git_profile).pack(side="left", padx=(0, 8))
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

    def _run_command(self, args, cwd=None):
        result = subprocess.run(args, cwd=cwd, capture_output=True, text=True, shell=False)
        return result

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

                with open(practice_file, "w", encoding="utf-8") as handle:
                    json.dump([{"payload": {"ods_code": ods}}], handle, indent=4)

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

        self._log_onboarding(
            f"Create completed for {practice_name} ({ods}). Folders: {folders_created}, Count updates: {counts_updated}"
        )
        messagebox.showinfo("Status", "\n".join(summary))
        self.run_validation_script()

    def offboard_practice(self):
        self._log_onboarding("Offboard clicked.")
        ods_code = self.entry_offboard_ods.get().strip().upper()
        if not ods_code:
            self._log_onboarding("Offboard failed: ODS code missing.")
            messagebox.showerror("Offboard", "Enter the ODS code to remove.")
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
                    with open(count_path, "w", encoding="utf-8") as handle:
                        json.dump(data, handle, indent=4)
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
