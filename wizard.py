"""
wizard.py — Multi-step tkinter wizard for configuring and running the email export.

Steps:
  1. Welcome — checks Outlook is running
  2. Folder Config — pick which folders to export and in what mode
  3. Output Directory — choose where to save exported files
  4. Confirm & Export — review settings, run export in background thread

Config is persisted to config.json so previous settings reload automatically.
"""

import json
import logging
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

logger = logging.getLogger(__name__)

CONFIG_FILE = "config.json"

EXPORT_MODES = ["skip", "attachments", "body", "both"]
EXPORT_MODE_LABELS = {
    "skip": "Skip",
    "attachments": "Attachments Only",
    "body": "Body Only",
    "both": "Both (Attachments + Body)",
}
# Reverse mapping for loading saved values
LABEL_TO_MODE = {v: k for k, v in EXPORT_MODE_LABELS.items()}


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

def load_config() -> dict:
    """Load config.json if it exists; return empty dict on any error."""
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        logger.warning("Could not load %s: %s — starting fresh.", CONFIG_FILE, e)
        return {}


def save_config(config: dict) -> None:
    """Write config dict to config.json."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except OSError as e:
        logger.warning("Could not save %s: %s", CONFIG_FILE, e)


# ---------------------------------------------------------------------------
# Wizard Application
# ---------------------------------------------------------------------------

class WizardApp(tk.Tk):
    """
    Main wizard window. Manages step transitions and accumulates config.
    result is set to the final config dict when the user clicks Export,
    or None if the user cancels.
    """

    STEPS = 4

    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook Email Export Wizard")
        self.resizable(True, True)
        self.minsize(600, 450)
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        self.config_data = load_config()
        self.result = None
        self.current_step = 0
        self._step_frame = None

        # Header
        header_frame = ttk.Frame(self, padding=(10, 8))
        header_frame.pack(fill="x")
        self._step_label = ttk.Label(
            header_frame, text="", font=("", 10, "bold")
        )
        self._step_label.pack(anchor="w")
        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # Content area
        self._container = ttk.Frame(self, padding=12)
        self._container.pack(fill="both", expand=True)

        # Footer
        ttk.Separator(self, orient="horizontal").pack(fill="x")
        footer = ttk.Frame(self, padding=(10, 6))
        footer.pack(fill="x")
        self._btn_back = ttk.Button(footer, text="< Back", command=self.prev_step)
        self._btn_back.pack(side="left")
        self._btn_cancel = ttk.Button(footer, text="Cancel", command=self.cancel)
        self._btn_cancel.pack(side="left", padx=6)
        self._btn_next = ttk.Button(footer, text="Next >", command=self.next_step)
        self._btn_next.pack(side="right")

        self.show_step(0)

    def _get_step_class(self, step: int):
        return [Step1Welcome, Step2FolderConfig, Step3OutputDir, Step4Confirm][step]

    def show_step(self, step: int) -> None:
        """Destroy current step frame and show the new one."""
        if self._step_frame is not None:
            self._step_frame.destroy()

        self._step_label.config(
            text=f"Step {step + 1} of {self.STEPS}"
        )
        self._btn_back.config(state="normal" if step > 0 else "disabled")

        # Step 4 has its own Export button in the frame; hide Next
        if step == self.STEPS - 1:
            self._btn_next.pack_forget()
        else:
            self._btn_next.pack(side="right")
            self._btn_next.config(state="normal", text="Next >")

        StepClass = self._get_step_class(step)
        self._step_frame = StepClass(self._container, wizard=self)
        self._step_frame.pack(fill="both", expand=True)
        self.current_step = step

    def next_step(self) -> None:
        """Validate current step; advance if valid."""
        frame = self._step_frame
        if hasattr(frame, "save_to_config"):
            frame.save_to_config()
        if hasattr(frame, "validate") and not frame.validate():
            return
        if self.current_step < self.STEPS - 1:
            self.show_step(self.current_step + 1)

    def prev_step(self) -> None:
        """Go back one step without validation."""
        if self.current_step > 0:
            if hasattr(self._step_frame, "save_to_config"):
                self._step_frame.save_to_config()
            self.show_step(self.current_step - 1)

    def finish(self) -> None:
        """Save config and signal completion."""
        save_config(self.config_data)
        self.result = self.config_data
        self.destroy()

    def cancel(self) -> None:
        """User cancelled; result stays None."""
        self.result = None
        self.destroy()

    def disable_nav(self) -> None:
        """Disable navigation buttons (used during export)."""
        self._btn_back.config(state="disabled")
        self._btn_cancel.config(state="disabled")


# ---------------------------------------------------------------------------
# Step 1: Welcome
# ---------------------------------------------------------------------------

class Step1Welcome(ttk.Frame):
    """Displays welcome text and Outlook connectivity status."""

    def __init__(self, parent, wizard: WizardApp, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.wizard = wizard

        ttk.Label(
            self,
            text="Welcome to the Outlook Email Export Wizard",
            font=("", 13, "bold"),
        ).pack(pady=(10, 6))

        ttk.Label(
            self,
            text=(
                "This wizard will help you export emails from Microsoft Outlook\n"
                "into organized folders named by date and transaction amount.\n\n"
                "Emails and attachments are saved exactly as they appear in Outlook.\n"
                "Dollar amounts are extracted from PDFs, Excel files, and email bodies."
            ),
            justify="center",
        ).pack(pady=6)

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=10)

        self._status_label = ttk.Label(self, text="Checking Outlook...", font=("", 10))
        self._status_label.pack(pady=4)

        self._check_outlook()

    def _check_outlook(self) -> None:
        from outlook_connector import is_outlook_running
        running = is_outlook_running()
        if running:
            self._status_label.config(
                text="Outlook is running",
                foreground="green",
            )
            self.wizard._btn_next.config(state="normal")
        else:
            self._status_label.config(
                text="Outlook is NOT running — please start Outlook and sign in, then click Refresh.",
                foreground="red",
                wraplength=460,
            )
            self.wizard._btn_next.config(state="disabled")
            ttk.Button(self, text="Refresh", command=self._check_outlook).pack(pady=4)

    def validate(self) -> bool:
        from outlook_connector import is_outlook_running
        if not is_outlook_running():
            messagebox.showerror(
                "Outlook Not Running",
                "Please start Microsoft Outlook and sign in before continuing.",
                parent=self.wizard,
            )
            return False
        return True


# ---------------------------------------------------------------------------
# Step 2: Folder Configuration
# ---------------------------------------------------------------------------

class Step2FolderConfig(ttk.Frame):
    """Scrollable folder tree with per-folder export mode dropdowns."""

    def __init__(self, parent, wizard: WizardApp, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.wizard = wizard
        self.mode_vars: dict = {}  # full_path → tk.StringVar

        ttk.Label(
            self,
            text="Select Outlook Folders to Export",
            font=("", 11, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        ttk.Label(
            self,
            text="Choose an export mode for each folder. 'Skip' means the folder is ignored.",
        ).pack(anchor="w")

        self._loading_label = ttk.Label(self, text="Loading folder list...")
        self._loading_label.pack(pady=8)

        # Build tree in background to keep UI responsive
        threading.Thread(target=self._load_folders, daemon=True).start()

    def _load_folders(self) -> None:
        try:
            from outlook_connector import get_folder_tree
            nodes = get_folder_tree()
            self.wizard.after(0, lambda: self._build_ui(nodes))
        except Exception as e:
            self.wizard.after(
                0,
                lambda: messagebox.showerror(
                    "Outlook Error",
                    f"Could not load Outlook folder list:\n{e}",
                    parent=self.wizard,
                )
            )

    def _build_ui(self, nodes: list) -> None:
        self._loading_label.destroy()

        # Scrollable canvas
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        inner.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        # Mouse wheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Header row
        hdr = ttk.Frame(inner)
        hdr.pack(fill="x", padx=4, pady=2)
        ttk.Label(hdr, text="Folder", font=("", 9, "bold"), width=50, anchor="w").pack(side="left")
        ttk.Label(hdr, text="Export Mode", font=("", 9, "bold")).pack(side="right", padx=(0, 20))
        ttk.Separator(inner, orient="horizontal").pack(fill="x")

        # Render top-level nodes (children are embedded recursively)
        for node in nodes:
            if node["depth"] == 0:
                self._render_node(inner, node)

        # Pre-populate from saved config
        saved_folders = self.wizard.config_data.get("folders", {})
        for path, mode in saved_folders.items():
            if path in self.mode_vars and mode in EXPORT_MODES:
                self.mode_vars[path].set(EXPORT_MODE_LABELS[mode])

    def _render_node(self, parent_frame, node: dict) -> None:
        """Render one folder row and recursively its children."""
        indent = node["depth"] * 20
        row = ttk.Frame(parent_frame)
        row.pack(fill="x", padx=4, pady=1)

        ttk.Label(row, text=" " * (indent // 5) + node["name"], anchor="w").pack(
            side="left", padx=(indent, 0)
        )

        var = tk.StringVar(value=EXPORT_MODE_LABELS["skip"])
        self.mode_vars[node["full_path"]] = var

        combo = ttk.Combobox(
            row,
            textvariable=var,
            values=list(EXPORT_MODE_LABELS.values()),
            state="readonly",
            width=24,
        )
        combo.pack(side="right", padx=4)

        for child in node.get("children", []):
            self._render_node(parent_frame, child)

    def validate(self) -> bool:
        modes = {path: LABEL_TO_MODE.get(var.get(), "skip")
                 for path, var in self.mode_vars.items()}
        active = [m for m in modes.values() if m != "skip"]
        if not active:
            return messagebox.askyesno(
                "No Folders Selected",
                "All folders are set to Skip. Nothing will be exported.\n"
                "Do you want to continue anyway?",
                parent=self.wizard,
            )
        return True

    def save_to_config(self) -> None:
        self.wizard.config_data["folders"] = {
            path: LABEL_TO_MODE.get(var.get(), "skip")
            for path, var in self.mode_vars.items()
        }


# ---------------------------------------------------------------------------
# Step 3: Output Directory
# ---------------------------------------------------------------------------

class Step3OutputDir(ttk.Frame):
    """Output directory picker."""

    def __init__(self, parent, wizard: WizardApp, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.wizard = wizard

        ttk.Label(
            self,
            text="Choose Output Directory",
            font=("", 11, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            self,
            text=(
                "Exported emails will be saved here, in subfolders named:\n"
                "  <FolderPath>\\YYYY-MM-DD_$X,XXX.XX\\\n\n"
                "Example:\n"
                "  C:\\Export\\Inbox\\Orders\\2026-03-04_$1,234.56\\"
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        pick_frame = ttk.Frame(self)
        pick_frame.pack(fill="x")

        self.dir_var = tk.StringVar(
            value=wizard.config_data.get("output_dir", "")
        )
        ttk.Entry(pick_frame, textvariable=self.dir_var, width=50).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(pick_frame, text="Browse...", command=self.browse).pack(
            side="left", padx=6
        )

    def browse(self) -> None:
        path = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.dir_var.get() or os.path.expanduser("~"),
            parent=self.wizard,
        )
        if path:
            self.dir_var.set(path)

    def validate(self) -> bool:
        path = self.dir_var.get().strip()
        if not path:
            messagebox.showerror(
                "No Directory",
                "Please select an output directory.",
                parent=self.wizard,
            )
            return False

        if not os.path.isdir(path):
            create = messagebox.askyesno(
                "Directory Does Not Exist",
                f"The directory does not exist:\n{path}\n\nCreate it now?",
                parent=self.wizard,
            )
            if create:
                try:
                    os.makedirs(path, exist_ok=True)
                except OSError as e:
                    messagebox.showerror(
                        "Error",
                        f"Could not create directory:\n{e}",
                        parent=self.wizard,
                    )
                    return False
            else:
                return False

        return True

    def save_to_config(self) -> None:
        self.wizard.config_data["output_dir"] = self.dir_var.get().strip()


# ---------------------------------------------------------------------------
# Step 4: Confirm & Export
# ---------------------------------------------------------------------------

class Step4Confirm(ttk.Frame):
    """Summary and export trigger with progress feedback."""

    def __init__(self, parent, wizard: WizardApp, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.wizard = wizard

        ttk.Label(
            self,
            text="Ready to Export",
            font=("", 11, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        self._build_summary()

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=10)

        self._progress_label = ttk.Label(self, text="")
        self._progress_label.pack(anchor="w")

        self._progressbar = ttk.Progressbar(self, mode="indeterminate", length=400)
        self._progressbar.pack(pady=6)

        self._export_btn = ttk.Button(
            self, text="Export", command=self.start_export, style="Accent.TButton"
        )
        self._export_btn.pack(pady=8)

    def _build_summary(self) -> None:
        config = self.wizard.config_data
        output_dir = config.get("output_dir", "(not set)")
        folders = config.get("folders", {})

        ttk.Label(self, text=f"Output directory:  {output_dir}", anchor="w").pack(anchor="w")

        mode_counts: dict = {}
        for mode in EXPORT_MODES:
            count = sum(1 for m in folders.values() if m == mode)
            if count:
                mode_counts[mode] = count

        if mode_counts:
            ttk.Label(self, text="Folders:", anchor="w").pack(anchor="w", pady=(8, 2))
            for mode, count in mode_counts.items():
                label = EXPORT_MODE_LABELS[mode]
                ttk.Label(
                    self, text=f"  {label}: {count} folder(s)", anchor="w"
                ).pack(anchor="w")
        else:
            ttk.Label(
                self, text="No folders selected (all skipped).", foreground="orange"
            ).pack(anchor="w")

    def start_export(self) -> None:
        """Disable UI, start indeterminate progressbar, launch export thread."""
        self._export_btn.config(state="disabled")
        self.wizard.disable_nav()
        self._progressbar.start(10)
        self._progress_label.config(text="Starting export...")

        config = dict(self.wizard.config_data)
        thread = threading.Thread(
            target=self._export_thread,
            args=(config,),
            daemon=True,
        )
        thread.start()

    def _export_thread(self, config: dict) -> None:
        """Run export in background thread. Marshals results back via root.after()."""
        from exporter import run_export

        def callback(current, total, subject):
            self.wizard.after(
                0,
                lambda c=current, t=total, s=subject: self.on_export_progress(c, t, s),
            )

        try:
            results = run_export(config, progress_callback=callback)
            self.wizard.after(0, lambda r=results: self.on_export_done(r, error=None))
        except Exception as e:
            self.wizard.after(0, lambda err=e: self.on_export_done([], error=err))

    def on_export_progress(self, current: int, total: int, subject: str) -> None:
        if total > 0:
            self._progress_label.config(
                text=f"Processing {current}/{total}: {subject[:60]}..."
            )
        else:
            self._progress_label.config(
                text=f"Processing item {current}: {subject[:60]}..."
            )

    def on_export_done(self, results: list, error=None) -> None:
        """Called when export thread completes."""
        self._progressbar.stop()

        if error is not None:
            messagebox.showerror(
                "Export Failed",
                f"An unexpected error occurred during export:\n{error}",
                parent=self.wizard,
            )
            self._export_btn.config(state="normal")
            self._progress_label.config(text="Export failed.")
            return

        total = len(results)
        with_amount = sum(1 for r in results if r.get("amount") is not None)
        with_errors = sum(1 for r in results if r.get("errors"))

        self._progress_label.config(text=f"Done. {total} emails exported.")

        messagebox.showinfo(
            "Export Complete",
            f"Export complete!\n\n"
            f"  Emails exported:  {total}\n"
            f"  With amounts:     {with_amount}\n"
            f"  With errors:      {with_errors}\n\n"
            f"Check export.log in the output directory for details.",
            parent=self.wizard,
        )

        self.wizard.finish()


# ---------------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------------

def run_wizard():
    """
    Instantiate WizardApp, run mainloop, return app.result.
    Returns the config dict if user completed export, None if cancelled.
    """
    app = WizardApp()
    app.mainloop()
    return app.result
