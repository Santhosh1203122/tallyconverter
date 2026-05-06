"""
Tally Sales Report Generator — desktop GUI.

Lets the user pick a Tally DayBook XML file, runs the converter, and
saves the resulting XLSX next to the chosen input file.
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    # When frozen (PyInstaller) and when run as a module
    from converter import convert
except ImportError:  # pragma: no cover
    from .converter import convert  # type: ignore


APP_TITLE = "Tally Sales Report Generator"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("640x360")
        self.minsize(560, 320)

        self._selected_path: str | None = None
        self._build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        # Use ttk for native look
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        outer = ttk.Frame(self, padding=20)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(
            outer,
            text="Tally DayBook → Sales Report",
            font=("Segoe UI", 16, "bold"),
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            outer,
            text=(
                "Pick the DayBook.xml exported from Tally. The Excel sales report "
                "will be saved next to it."
            ),
            wraplength=580,
            foreground="#555",
        )
        subtitle.pack(anchor="w", pady=(2, 16))

        # File picker row
        picker = ttk.Frame(outer)
        picker.pack(fill="x")
        self.path_var = tk.StringVar(value="No file selected")
        path_entry = ttk.Entry(picker, textvariable=self.path_var, state="readonly")
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        browse_btn = ttk.Button(picker, text="Browse…", command=self._browse)
        browse_btn.pack(side="left")

        # Convert button
        self.convert_btn = ttk.Button(
            outer,
            text="Convert to Excel",
            command=self._on_convert,
            state="disabled",
        )
        self.convert_btn.pack(pady=(20, 8), ipadx=10, ipady=4)

        # Progress bar
        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.pack(fill="x", pady=(4, 8))

        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(outer, textvariable=self.status_var, foreground="#333")
        status.pack(anchor="w")

        # Footer
        footer = ttk.Label(
            outer,
            text="Output is saved as <input_name>_sales_report.xlsx",
            foreground="#777",
            font=("Segoe UI", 9),
        )
        footer.pack(side="bottom", anchor="w", pady=(20, 0))

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Tally DayBook XML",
            filetypes=[("Tally XML files", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self._selected_path = path
            self.path_var.set(path)
            self.convert_btn.config(state="normal")
            self.status_var.set("Ready to convert.")

    def _on_convert(self) -> None:
        if not self._selected_path:
            return
        self.convert_btn.config(state="disabled")
        self.status_var.set("Converting… please wait.")
        self.progress.start(10)
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self) -> None:
        try:
            out_path = convert(self._selected_path)  # type: ignore[arg-type]
            self.after(0, self._on_success, out_path)
        except Exception:  # noqa: BLE001
            err = traceback.format_exc()
            self.after(0, self._on_failure, err)

    def _on_success(self, out_path: str) -> None:
        self.progress.stop()
        self.status_var.set(f"Done. Saved to: {out_path}")
        self.convert_btn.config(state="normal")
        if messagebox.askyesno(
            APP_TITLE,
            f"Sales report saved to:\n\n{out_path}\n\nOpen the folder now?",
        ):
            self._open_folder(os.path.dirname(out_path))

    def _on_failure(self, err: str) -> None:
        self.progress.stop()
        self.status_var.set("Conversion failed. See error dialog.")
        self.convert_btn.config(state="normal")
        messagebox.showerror(APP_TITLE, f"Conversion failed:\n\n{err}")

    @staticmethod
    def _open_folder(folder: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", folder])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", folder])
        except Exception:  # noqa: BLE001
            pass


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
