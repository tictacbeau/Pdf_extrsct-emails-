"""
main.py — Entry point for the Outlook Email Export Wizard.

Usage:
    python main.py

Requirements:
    - Windows OS with Microsoft Outlook desktop app installed and signed in
    - Python with packages from requirements.txt installed:
        pip install -r requirements.txt

Flow:
    1. Run the wizard (4-step tkinter UI)
    2. Set up logging to output_dir/export.log
    3. Export runs inside the wizard's Step 4 background thread
    4. Show completion summary (handled by wizard)
"""

import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox


def setup_logging(output_dir: str) -> None:
    """
    Configure root logger to write to output_dir/export.log (all levels)
    and stdout (INFO and above).
    """
    log_path = os.path.join(output_dir, "export.log")
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Remove any existing handlers (avoid duplicate log entries on re-run)
    root_logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler — everything
    try:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        root_logger.addHandler(fh)
    except OSError as e:
        print(f"Warning: could not open log file {log_path!r}: {e}", file=sys.stderr)

    # Console handler — INFO+
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    root_logger.addHandler(ch)


def main() -> None:
    """
    Bootstrap logging (to cwd first), run the wizard, then switch logging
    to the chosen output directory.
    """
    # Minimal early logging so startup errors are visible
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        stream=sys.stdout,
    )

    try:
        from wizard import run_wizard

        config = run_wizard()

        if config is None:
            # User cancelled the wizard
            logging.info("Wizard cancelled by user.")
            return

        output_dir = config.get("output_dir", "")
        if output_dir:
            try:
                os.makedirs(output_dir, exist_ok=True)
                setup_logging(output_dir)
            except OSError as e:
                logging.warning("Could not set up logging in output dir: %s", e)

        logging.info("Export configuration accepted. Output: %s", output_dir)
        # Note: the actual export runs inside the wizard's Step 4 background thread.
        # By the time run_wizard() returns, the export has already completed.

    except Exception as e:
        # Last-resort error display
        logging.critical("Unhandled exception: %s", e, exc_info=True)
        try:
            # Try to show a GUI error dialog
            _root = tk.Tk()
            _root.withdraw()
            messagebox.showerror(
                "Fatal Error",
                f"An unexpected error occurred:\n\n{e}\n\n"
                "Check export_error.log for details.",
            )
            _root.destroy()
        except Exception:
            pass

        # Also write to a fallback log file
        try:
            with open("export_error.log", "a", encoding="utf-8") as f:
                import traceback
                f.write(traceback.format_exc())
        except Exception:
            pass

        sys.exit(1)


if __name__ == "__main__":
    main()
