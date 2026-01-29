import time
import secrets
import string
import os
import ctypes
import pyautogui
import pyperclip
import pygetwindow as gw
import tkinter as tk
import logging
from tkinter import messagebox, scrolledtext
from pywinauto import Desktop, Application
from pywinauto.timings import TimeoutError as WaitTimeoutError

# --- SMART PATH DETECTION ---
def get_safe_log_path(filename):
    user_profile = os.environ.get('USERPROFILE')
    paths_to_try = [
        os.path.join(user_profile, 'OneDrive', 'Desktop'),
        os.path.join(user_profile, 'Desktop'),
        os.path.dirname(os.path.abspath(__file__))
    ]
    for path in paths_to_try:
        if os.path.exists(path):
            return os.path.join(path, filename)
    return filename

LOG_FILE_PATH = get_safe_log_path("password_log.txt")
DEBUG_LOG_PATH = get_safe_log_path("debug_log.txt")

for path in [LOG_FILE_PATH, DEBUG_LOG_PATH]:
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f: f.write("")

class SafeFileHandler(logging.FileHandler):
    def emit(self, record):
        super().emit(record)
        self.flush()
    def flush(self):
        super().flush()
        try:
            if self.stream and hasattr(self.stream, "fileno"):
                os.fsync(self.stream.fileno())
        except OSError: pass

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

def generate_strict_password(length=10):
    alphabet = string.ascii_letters + string.digits
    while True:
        pwd = ''.join(secrets.choice(alphabet) for _ in range(length))
        if (any(c.isupper() for c in pwd) and any(c.isdigit() for c in pwd)):
            return pwd

# --- MAIN APP ---
class EMISApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EMIS Pro Dashboard v6.11")
        self.root.geometry("450x820")
        self.root.attributes("-topmost", True)
        self.generated_pwd = ""
        self.setup_ui()
        self.setup_logging()
        logging.info("v6.11 Online. Modal drilling logic enabled.")

    def setup_ui(self):
        tk.Label(self.root, text="EMIS DASHBOARD", font=("Arial", 14, "bold")).pack(pady=10)
        tk.Button(self.root, text="AUTO-DETECT & RUN", command=self.auto_detect_and_run, bg="#2196F3", fg="white", width=35, height=2, font=("Arial", 10, "bold")).pack(pady=5)
        gen_frame = tk.LabelFrame(self.root, text="Fresh Password Generator", padx=10, pady=5)
        gen_frame.pack(pady=5, fill="x", padx=20)
        self.pwd_entry = tk.Entry(gen_frame, font=("Courier", 11), justify='center')
        self.pwd_entry.pack(pady=5, fill="x")
        self.pwd_entry.insert(0, "[None]")
        tk.Button(gen_frame, text="GENERATE PASSWORD", command=self.generate_ui, bg="#9E9E9E", width=30).pack(pady=5)
        tk.Button(self.root, text="DELAYED PASTE (Wait 3s)", command=self.delayed_paste, bg="#607D8B", fg="white", width=35).pack(pady=5)
        tk.Label(self.root, text="Manual Overrides", font=("Arial", 9, "italic")).pack(pady=5)
        tk.Button(self.root, text="EXPIRED RESET", command=self.run_standard_automation, bg="#4CAF50", fg="white", width=35).pack(pady=2)
        tk.Button(self.root, text="SETTINGS RESET", command=self.run_settings_automation, bg="#FF9800", fg="white", width=35).pack(pady=2)
        tk.Button(self.root, text="UNLOCK", command=self.unlock_locked_screen, bg="#9C27B0", fg="white", width=35).pack(pady=2)
        tk.Label(self.root, text="Live System Log:", font=("Arial", 9, "bold")).pack(pady=(10, 0))
        self.log_widget = scrolledtext.ScrolledText(self.root, height=12, font=("Consolas", 8), bg="black", fg="#00FF00")
        self.log_widget.pack(pady=5, padx=20, fill="both", expand=True)
        self.log_widget.configure(state='disabled')
        log_btn_frame = tk.Frame(self.root)
        log_btn_frame.pack(pady=10)
        tk.Button(log_btn_frame, text="View Logs", command=self.open_log, width=15).grid(row=0, column=0, padx=5)
        tk.Button(log_btn_frame, text="Clear Logs", command=self.clear_log, width=15, bg="#f44336", fg="white").grid(row=0, column=1, padx=5)

    def setup_logging(self):
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(relativeCreated)d INFO: %(message)s')
        file_handler = SafeFileHandler(DEBUG_LOG_PATH, encoding='utf-8')
        file_handler.setFormatter(formatter)
        gui_handler = TextHandler(self.log_widget)
        gui_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(gui_handler)

    def log_password(self, pwd, context):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        entry = f"{timestamp} [{context}] - {pwd}\n"
        try:
            with open(LOG_FILE_PATH, "a", encoding='utf-8') as f:
                f.write(entry)
                f.flush()
                os.fsync(f.fileno())
            logging.info(f"✅ RECORD SAVED: {pwd}")
        except Exception as e: logging.error(f"Save Error: {e}")

    def generate_ui(self):
        self.generated_pwd = generate_strict_password(10)
        self.pwd_entry.delete(0, tk.END)
        self.pwd_entry.insert(0, self.generated_pwd)
        pyperclip.copy(self.generated_pwd)
        logging.info(f"Fresh Password: {self.generated_pwd}")
        return self.generated_pwd

    def run_settings_automation(self):
        """Refined drilling: Finds EMIS Main -> Wizard Child -> loginPanel."""
        pwd = self.generate_ui()
        logging.info("Scanning for EMIS Web main window...")
        try:
            # 1. Connect to main EMIS window first
            dt = Desktop(backend="uia")
            main_emis = dt.window(title_re=".*EMIS Web Health Care System.*")
            
            # 2. Drill to child UserWizardForm
            wizard = main_emis.child_window(auto_id="UserWizardForm", control_type="Window")
            if not wizard.exists(timeout=5):
                logging.error("Could not find 'Edit user' child. Ensure it is open within EMIS.")
                return
            
            wizard.set_focus()
            logging.info("Drilling into loginPanel pane...")
            
            # 3. Drill directly into loginPanel as requested
            container = wizard.child_window(auto_id="loginPanel", control_type="Pane")
            
            p_field = container.child_window(auto_id="passwordTextBox", control_type="Edit")
            c_field = container.child_window(auto_id="confirmPasswordTextBox", control_type="Edit")

            p_field.set_focus()
            pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(pwd, interval=0.01)

            c_field.set_focus()
            pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(pwd, interval=0.01)

            self.log_password(pwd, "Settings")
            messagebox.showinfo("Success", "Wizard fields populated via Drill-Down.")
        except Exception as e:
            logging.error(f"Modal Drill-down Failure: {e}")

    def unlock_locked_screen(self):
        """Restore missing function for the Purple button."""
        logging.info("Attempting unlock...")
        pwd = pyperclip.paste()
        try:
            app = Application(backend="uia").connect(title_re=".*Locked.*", timeout=5)
            dlg = app.window(title_re=".*Locked.*")
            dlg.set_focus()
            dlg.child_window(auto_id="textBoxPassword", control_type="Edit").set_focus()
            pyautogui.write(pwd, interval=0.01)
            dlg.child_window(auto_id="buttonUnlock", control_type="Button").click()
            self.log_password(pwd, "Unlock")
        except Exception as e: logging.error(f"Unlock Failure: {e}")

    def run_standard_automation(self):
        pwd = self.generate_ui()
        logging.info("Standard Reset initiated...")
        wins = gw.getWindowsWithTitle("Authentication")
        target = next((w for w in wins if 500 < w.width < 700), None)
        if target:
            target.activate()
            time.sleep(0.5)
            pyautogui.click(target.left + 350, target.top + 215)
            pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(pwd, interval=0.01)
            pyautogui.press('tab')
            pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(pwd, interval=0.01)
            self.log_password(pwd, "Expired")
            pyautogui.press('enter')
        else: logging.error("Auth window not found.")

    def delayed_paste(self):
        logging.info("Pasting in 3s...")
        self.root.after(3000, lambda: pyautogui.write(pyperclip.paste(), interval=0.01))

    def auto_detect_and_run(self):
        titles = [w.title for w in gw.getAllWindows()]
        if any("Locked" in t for t in titles): self.unlock_locked_screen()
        elif any("Edit user" in t for t in titles): self.run_settings_automation()
        elif any("Authentication" in t for t in titles): self.run_standard_automation()
        else: logging.warning("No compatible window detected.")

    def open_log(self):
        if os.path.exists(LOG_FILE_PATH): os.startfile(LOG_FILE_PATH)
    def clear_log(self):
        if messagebox.askyesno("Confirm", "Purge all logs?"):
            for p in [LOG_FILE_PATH, DEBUG_LOG_PATH]: open(p, 'w', encoding='utf-8').close()

if __name__ == "__main__":
    root = tk.Tk()
    app = EMISApp(root)
    root.mainloop()