import sys
import os
import json
import subprocess
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, Toplevel, Label, Button
from tkinter.ttk import Combobox
from tkinter import simpledialog

# === PATH HANDLING FOR EXE/SCRIPT ===
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle (EXE), use the temp folder
    SCRIPT_DIR = sys._MEIPASS
else:
    # If run as a normal script
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# === CONFIG ===
# 1. Manually set this to the folder containing your two repositories.
# Based on your logs, it looks like: C:\Users\RPA-012\AppData\Local
# Change this string if your real repos are elsewhere (e.g., C:\Repos)
PROJECT_BASE = r"C:\rpa\postie" # Ensure this path is correct for your machine!

ROOT_FOLDERS = [
    os.path.join(PROJECT_BASE, "postie_bots_python", "devdata", "work-items-in"),
    os.path.join(PROJECT_BASE, "postie-bots", "devdata", "work-items-in")
]

# This ensures Git targets the correct repo folder
GIT_REPO_PATH = os.path.join(PROJECT_BASE, "postie-bots")

# Path to PowerShell validation script - ensure this matches your filename in C:\Tools
CHECK_ODSMISMATCH_SCRIPT = os.path.join(SCRIPT_DIR, "check-ods-mismatch.ps1")

def run_validation_script():
    """Runs the PS1 script for each root folder to ensure ODS consistency."""
    all_outputs = []
    for folder in ROOT_FOLDERS:
        if not os.path.exists(folder):
            print(f"DEBUG: Skipping validation for missing folder: {folder}")
            continue
            
        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", 
                 CHECK_ODSMISMATCH_SCRIPT, "-BasePath", folder],
                capture_output=True,
                text=True
            )
            if result.stdout.strip():
                all_outputs.append(f"--- Folder: {os.path.basename(folder)} ---\n{result.stdout.strip()}")
            if result.stderr.strip():
                all_outputs.append(f"Errors in {os.path.basename(folder)}: {result.stderr.strip()}")
        except Exception as e:
            all_outputs.append(f"Failed to run check for {folder}: {e}")

    final_message = "\n\n".join(all_outputs)
    messagebox.showinfo("ODS Mismatch Check", final_message if final_message else "✅ No issues detected.")

last_practices = []

# === ONBOARDING ===
def create_json_files():
    system_type = system_var.get()  # "Docman" or "EMIS"
    p_name = entry_practice.get().strip()
    ods = entry_ods.get().strip().upper()

    if not p_name or not ods:
        messagebox.showerror("Error", "Please enter both Name and ODS.")
        return

    last_practices.append((p_name, ods))

    folders_created = 0
    counts_updated = 0
    debug_notes = []

    for ROOT_FOLDER in ROOT_FOLDERS:
        if not os.path.isdir(ROOT_FOLDER):
            debug_notes.append(f"Root folder NOT found (skipped): {ROOT_FOLDER}")
            continue

        try:
            # 1) Practice folder name: "Name (ODS)"
            f_name = f"{p_name.title()} ({ods})"
            p_folder = os.path.join(ROOT_FOLDER, f_name)
            wi_path = os.path.join(p_folder, "work-items.json")

            # Ensure folder exists
            if not os.path.exists(p_folder):
                os.makedirs(p_folder, exist_ok=True)
                folders_created += 1
                debug_notes.append(f"Created folder: {p_folder}")
            elif not os.path.isdir(p_folder):
                raise RuntimeError(f"A file already exists with this name:\n{p_folder}")
            else:
                debug_notes.append(f"Folder already existed: {p_folder}")

            # Always write practice work-items.json
            with open(wi_path, "w") as f:
                json.dump([{"payload": {"ods_code": ods}}], f, indent=4)

            # 2) Practice Count (DOCMAN ONLY)
            if system_type == "Docman":
                count_path = os.path.join(ROOT_FOLDER, "Practice Count", "work-items.json")
                data = []

                if os.path.exists(count_path):
                    with open(count_path, "r") as f:
                        try:
                            data = json.load(f)
                        except Exception:
                            debug_notes.append(f"Bad JSON in {count_path}, starting fresh.")
                            data = []

                ods_exists = any(
                    str(e.get("payload", {}).get("ods_code", "")).upper() == ods
                    for e in data
                )

                if not ods_exists:
                    data.append({
                        "payload": {
                            "ods_code": ods,
                            "docman_practice_display_name": p_name.upper()
                        }
                    })
                    os.makedirs(os.path.dirname(count_path), exist_ok=True)
                    with open(count_path, "w") as f:
                        json.dump(data, f, indent=4)
                    counts_updated += 1
                else:
                    debug_notes.append(f"ODS {ods} already in Practice Count for {ROOT_FOLDER}")
            else:
                debug_notes.append(f"EMIS selected – skipped Practice Count for {ROOT_FOLDER}")

        except Exception as e:
            messagebox.showerror("Folder Error", f"Failed at {ROOT_FOLDER}:\n{e}")
            return

    status = (
        f"Onboarding Summary for {p_name} ({system_type}):\n"
        f"✅ Folders created this run: {folders_created}\n"
        f"✅ Practice Count updates: {counts_updated}"
    )

    if folders_created == 0 and counts_updated == 0:
        status += "\n\n⚠️ No changes made. Folder/ODS may already exist."
    elif system_type == "Docman" and counts_updated < len(ROOT_FOLDERS):
        status += "\n\n⚠️ Note: ODS already existed in some 'Practice Count' files."

    if debug_notes:
        status += "\n\n--- Debug ---\n" + "\n".join(debug_notes)

    messagebox.showinfo("Status", status)
    run_validation_script()
def create_json_files():
    system_type = system_var.get()  # "Docman" or "EMIS"
    p_name = entry_practice.get().strip()
    ods = entry_ods.get().strip().upper()

    if not p_name or not ods:
        messagebox.showerror("Error", "Please enter both Name and ODS.")
        return

    last_practices.append((p_name, ods))

    folders_created = 0
    counts_updated = 0
    debug_notes = []

    for ROOT_FOLDER in ROOT_FOLDERS:
        if not os.path.isdir(ROOT_FOLDER):
            debug_notes.append(f"Root folder NOT found (skipped): {ROOT_FOLDER}")
            continue

        try:
            # 1) Practice folder name: "Name (ODS)"
            f_name = f"{p_name.title()} ({ods})"
            p_folder = os.path.join(ROOT_FOLDER, f_name)
            wi_path = os.path.join(p_folder, "work-items.json")

            # Ensure folder exists
            if not os.path.exists(p_folder):
                os.makedirs(p_folder, exist_ok=True)
                folders_created += 1
                debug_notes.append(f"Created folder: {p_folder}")
            elif not os.path.isdir(p_folder):
                raise RuntimeError(f"A file already exists with this name:\n{p_folder}")
            else:
                debug_notes.append(f"Folder already existed: {p_folder}")

            # Always write practice work-items.json
            with open(wi_path, "w") as f:
                json.dump([{"payload": {"ods_code": ods}}], f, indent=4)

            # 2) Practice Count (DOCMAN ONLY)
            if system_type == "Docman":
                count_path = os.path.join(ROOT_FOLDER, "Practice Count", "work-items.json")
                data = []

                if os.path.exists(count_path):
                    with open(count_path, "r") as f:
                        try:
                            data = json.load(f)
                        except Exception:
                            debug_notes.append(f"Bad JSON in {count_path}, starting fresh.")
                            data = []

                ods_exists = any(
                    str(e.get("payload", {}).get("ods_code", "")).upper() == ods
                    for e in data
                )

                if not ods_exists:
                    data.append({
                        "payload": {
                            "ods_code": ods,
                            "docman_practice_display_name": p_name.upper()
                        }
                    })
                    os.makedirs(os.path.dirname(count_path), exist_ok=True)
                    with open(count_path, "w") as f:
                        json.dump(data, f, indent=4)
                    counts_updated += 1
                else:
                    debug_notes.append(f"ODS {ods} already in Practice Count for {ROOT_FOLDER}")
            else:
                debug_notes.append(f"EMIS selected – skipped Practice Count for {ROOT_FOLDER}")

        except Exception as e:
            messagebox.showerror("Folder Error", f"Failed at {ROOT_FOLDER}:\n{e}")
            return

    status = (
        f"Onboarding Summary for {p_name} ({system_type}):\n"
        f"✅ Folders created this run: {folders_created}\n"
        f"✅ Practice Count updates: {counts_updated}"
    )

    if folders_created == 0 and counts_updated == 0:
        status += "\n\n⚠️ No changes made. Folder/ODS may already exist."
    elif system_type == "Docman" and counts_updated < len(ROOT_FOLDERS):
        status += "\n\n⚠️ Note: ODS already existed in some 'Practice Count' files."

    if debug_notes:
        status += "\n\n--- Debug ---\n" + "\n".join(debug_notes)

    messagebox.showinfo("Status", status)
    run_validation_script()

# === OFFBOARDING ===
def offboard_practice():
    ods_code = entry_offboard_ods.get().strip().upper()
    if not ods_code:
        messagebox.showerror("Missing Info", "Please enter the ODS Code to offboard.")
        return

    for ROOT_FOLDER in ROOT_FOLDERS:
        try:
            count_path = os.path.join(ROOT_FOLDER, "Practice Count", "work-items.json")
            if not os.path.exists(count_path): continue

            with open(count_path, "r") as f:
                data = json.load(f)
            
            entry_to_remove = next((e for e in data if e.get("payload", {}).get("ods_code") == ods_code), None)

            if entry_to_remove:
                confirm = messagebox.askyesno("Confirm", f"Remove {ods_code} from Practice Count in {os.path.basename(ROOT_FOLDER)}?")
                if confirm:
                    data = [e for e in data if e != entry_to_remove]
                    with open(count_path, "w") as f:
                        json.dump(data, f, indent=4)
                    messagebox.showinfo("Success", f"Removed {ods_code} from {os.path.basename(ROOT_FOLDER)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error during offboarding: {e}")

# === GIT LOGIC ===
def run_git_push(branch_name, commit_msg):
    try:
        os.chdir(GIT_REPO_PATH)

        # Ensure clean starting point
        subprocess.run(["git", "checkout", "main"], check=True)
        subprocess.run(["git", "pull"], check=True)

        # Create or switch to branch
        result = subprocess.run(
            ["git", "branch", "--list", branch_name],
            capture_output=True,
            text=True
        )

        if result.stdout.strip():
            subprocess.run(["git", "checkout", branch_name], check=True)
        else:
            subprocess.run(["git", "checkout", "-b", branch_name], check=True)

        subprocess.run(["git", "add", "."], check=True)

        # Avoid empty commits
        status = subprocess.run(
            ["git", "status", "--porcelain"],
            capture_output=True,
            text=True
        )
        if not status.stdout.strip():
            return "No changes detected to commit."

        subprocess.run(["git", "commit", "-m", commit_msg], check=True)

        if messagebox.askyesno("Push", f"Push branch '{branch_name}' to origin?"):
            subprocess.run(
                ["git", "push", "-u", "origin", branch_name],
                check=True
            )
            return "✅ Push successful!"

        return "Commit created locally. Push cancelled."

    except subprocess.CalledProcessError as e:
        return f"❌ Git command failed:\n{e}"
    except Exception as e:
        return f"❌ Git error:\n{e}"


def open_git_push_window():
    if not last_practices:
        messagebox.showinfo("Info", "Onboard a practice first to generate suggestions.")
        return

    win = Toplevel(root)
    win.title("Git Push")
    win.geometry("400x250")

    pname, pcode = last_practices[-1]
    
    Label(win, text="Branch Name:").pack(pady=5)
    b_entry = tk.Entry(win, width=40)
    b_entry.insert(0, f"onboard/{pcode}")
    b_entry.pack()

    Label(win, text="Commit Message:").pack(pady=5)
    c_entry = tk.Entry(win, width=40)
    c_entry.insert(0, f"Onboarded: {pname} ({pcode})")
    c_entry.pack()

    def handle_push():
        result = run_git_push(b_entry.get(), c_entry.get())
        messagebox.showinfo("Result", result)
        win.destroy()

    Button(win, text="🚀 Push to Git", command=handle_push, bg="lightgreen").pack(pady=20)

# === GUI MAIN ===
root = tk.Tk()
root.title("Practice Onboarding Tool")
root.geometry("500x550")

# Onboard Frame
f1 = tk.LabelFrame(root, text="🩺 Onboard Practice", padx=10, pady=10)
f1.pack(padx=10, pady=10, fill="both")

Label(f1, text="Practice Name:").pack()
entry_practice = tk.Entry(f1, width=40)
entry_practice.pack(pady=5)

Label(f1, text="ODS Code:").pack()
entry_ods = tk.Entry(f1, width=40)
entry_ods.pack(pady=5)

Label(f1, text="Practice System:").pack(pady=(10, 0))

system_var = tk.StringVar(value="Docman")
system_combo = Combobox(
    f1,
    textvariable=system_var,
    values=["Docman", "EMIS"],
    state="readonly",
    width=37
)
system_combo.pack(pady=5)

Button(f1, text="➕ Create Files", command=create_json_files, bg="lightblue").pack(pady=10)
Button(f1, text="🚀 Push to Git", command=open_git_push_window).pack()

# Offboard Frame
f2 = tk.LabelFrame(root, text="🧹 Offboard Practice", padx=10, pady=10)
f2.pack(padx=10, pady=10, fill="both")

Label(f2, text="ODS Code to Remove:").pack()
entry_offboard_ods = tk.Entry(f2, width=40)
entry_offboard_ods.pack(pady=5)

Button(f2, text="🗑️ Offboard", command=offboard_practice, bg="#ffcccb").pack(pady=5)

root.mainloop()