#!/usr/bin/env python3
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "emis_letter_summary.py"
VENV_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"


class EmisSummaryUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("EMIS Letter Summary")
        self.geometry("980x720")

        self._process: subprocess.Popen[str] | None = None
        self._reader_thread: threading.Thread | None = None
        self._output_queue: queue.Queue[str] = queue.Queue()

        self.practice_var = tk.StringVar(value=os.getenv("PRACTICE_NAME", ""))
        self.ods_var = tk.StringVar(value=os.getenv("ODS_CODE", ""))
        self.queues_var = tk.StringVar(value="Awaiting Filing")
        self.csv_var = tk.StringVar(value=str(BASE_DIR / "output" / "summary.csv"))
        self.timeout_var = tk.StringVar(value="120")

        self.strict_var = tk.BooleanVar(value=True)
        self.skip_login_var = tk.BooleanVar(value=False)
        self.quiet_var = tk.BooleanVar(value=False)
        self.no_kill_var = tk.BooleanVar(value=False)
        self.no_cache_var = tk.BooleanVar(value=False)

        self._build_layout()
        self.after(150, self._drain_output_queue)

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(root, text="Run Settings", padding=12)
        form.pack(fill=tk.X)

        self._add_row(form, 0, "Practice Name", self.practice_var)
        self._add_row(form, 1, "ODS Code", self.ods_var)
        self._add_row(form, 2, "Queues", self.queues_var)
        self._add_row(form, 3, "Timeout (sec)", self.timeout_var)

        csv_label = ttk.Label(form, text="CSV Output")
        csv_label.grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        csv_entry = ttk.Entry(form, textvariable=self.csv_var)
        csv_entry.grid(row=4, column=1, sticky=tk.EW, pady=4)
        csv_browse = ttk.Button(form, text="Browse", command=self._browse_csv)
        csv_browse.grid(row=4, column=2, padx=(8, 0), pady=4)

        note = ttk.Label(
            form,
            text="Mailroom API URL/Key are loaded from .env or system environment.",
        )
        note.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(2, 0))

        form.columnconfigure(1, weight=1)

        options = ttk.Frame(root, padding=(0, 8, 0, 8))
        options.pack(fill=tk.X)

        ttk.Checkbutton(options, text="Strict Sidebar Match", variable=self.strict_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(options, text="Skip Login (Attach Session)", variable=self.skip_login_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(options, text="Quiet", variable=self.quiet_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(options, text="No Kill Existing", variable=self.no_kill_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(options, text="No Clear Cache", variable=self.no_cache_var).pack(side=tk.LEFT, padx=(0, 12))

        controls = ttk.Frame(root)
        controls.pack(fill=tk.X)
        self.run_btn = ttk.Button(controls, text="Run", command=self._run)
        self.run_btn.pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(controls, text="Stop", command=self._stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(controls, text="Open Output Folder", command=self._open_output_folder).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(controls, text="Clear Log", command=self._clear_output).pack(side=tk.LEFT, padx=(8, 0))

        output_frame = ttk.LabelFrame(root, text="Console Output", padding=8)
        output_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.output = ScrolledText(output_frame, wrap=tk.WORD, height=26)
        self.output.pack(fill=tk.BOTH, expand=True)
        self.output.configure(state=tk.DISABLED)

    def _add_row(
        self,
        parent: ttk.Widget,
        row: int,
        label: str,
        variable: tk.StringVar,
        show: str | None = None,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        kwargs: dict[str, str] = {}
        if show:
            kwargs["show"] = show
        ttk.Entry(parent, textvariable=variable, **kwargs).grid(
            row=row,
            column=1,
            columnspan=2,
            sticky=tk.EW,
            pady=4,
        )

    def _browse_csv(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="CSV output path",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            initialdir=str(BASE_DIR / "output"),
        )
        if selected:
            self.csv_var.set(selected)

    def _append_output(self, text: str) -> None:
        self.output.configure(state=tk.NORMAL)
        self.output.insert(tk.END, text + "\n")
        self.output.see(tk.END)
        self.output.configure(state=tk.DISABLED)

    def _clear_output(self) -> None:
        self.output.configure(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        self.output.configure(state=tk.DISABLED)

    def _open_output_folder(self) -> None:
        output_dir = BASE_DIR / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(str(output_dir))

    def _validate(self) -> tuple[bool, str]:
        if not SCRIPT_PATH.exists():
            return False, f"Missing script: {SCRIPT_PATH}"

        if not self.skip_login_var.get() and not (
            self.practice_var.get().strip() or self.ods_var.get().strip()
        ):
            return False, "Provide Practice Name or ODS Code, or use Skip Login."

        try:
            timeout = int(self.timeout_var.get().strip())
            if timeout <= 0:
                raise ValueError
        except Exception:
            return False, "Timeout must be a positive integer."

        return True, ""

    def _build_command(self) -> list[str]:
        python_exe = str(VENV_PYTHON if VENV_PYTHON.exists() else Path(sys.executable))
        cmd = [python_exe, str(SCRIPT_PATH)]

        practice = self.practice_var.get().strip()
        if practice:
            cmd += ["--practice", practice]

        ods = self.ods_var.get().strip()
        if ods:
            cmd += ["--ods-code", ods]

        queues = self.queues_var.get().strip()
        if queues:
            cmd += ["--queues", queues]

        csv_output = self.csv_var.get().strip()
        if csv_output:
            cmd += ["--csv-output", csv_output]

        timeout = self.timeout_var.get().strip()
        if timeout:
            cmd += ["--timeout", timeout]

        if self.strict_var.get():
            cmd.append("--strict-sidebar-match")
        if self.skip_login_var.get():
            cmd.append("--skip-login")
        if self.quiet_var.get():
            cmd.append("--quiet")
        if self.no_kill_var.get():
            cmd.append("--no-kill-existing")
        if self.no_cache_var.get():
            cmd.append("--no-clear-cache")

        return cmd

    def _run(self) -> None:
        valid, message = self._validate()
        if not valid:
            messagebox.showerror("Validation", message)
            return

        if self._process is not None:
            messagebox.showwarning("Already Running", "A run is already in progress.")
            return

        cmd = self._build_command()
        pretty = " ".join(f'"{x}"' if " " in x else x for x in cmd)
        self._append_output("$ " + pretty)

        self.run_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)

        try:
            self._process = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                env=os.environ.copy(),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
        except Exception as error:
            self._append_output(f"Failed to start process: {error}")
            self._process = None
            self.run_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            return

        self._reader_thread = threading.Thread(
            target=self._read_process_output,
            daemon=True,
        )
        self._reader_thread.start()

    def _read_process_output(self) -> None:
        assert self._process is not None
        process = self._process
        assert process.stdout is not None
        for line in process.stdout:
            self._output_queue.put(line.rstrip("\n"))
        exit_code = process.wait()
        self._output_queue.put(f"\nProcess exited with code {exit_code}.")
        self._output_queue.put("__PROCESS_DONE__")

    def _stop(self) -> None:
        if self._process is None:
            return
        try:
            self._process.terminate()
            self._append_output("Termination requested...")
        except Exception as error:
            self._append_output(f"Failed to terminate process: {error}")

    def _drain_output_queue(self) -> None:
        try:
            while True:
                item = self._output_queue.get_nowait()
                if item == "__PROCESS_DONE__":
                    self._process = None
                    self._reader_thread = None
                    self.run_btn.configure(state=tk.NORMAL)
                    self.stop_btn.configure(state=tk.DISABLED)
                else:
                    self._append_output(item)
        except queue.Empty:
            pass
        self.after(150, self._drain_output_queue)


def main() -> int:
    app = EmisSummaryUI()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
