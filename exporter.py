"""
exporter.py — Email export pipeline orchestrator.

Loops over configured Outlook folders, processes each MailItem,
extracts amounts, and saves files into named output folders.

COM threading: This module's run_export() is designed to be called from
a background thread. pythoncom.CoInitialize() must be called before any
COM use in that thread, and CoUninitialize() in a finally block.
"""

import os
import logging
import tempfile
import shutil
from datetime import datetime

from outlook_connector import get_mail_items, FolderNotFoundError, get_outlook_app
from parser import extract_amount_from_file, extract_amount_from_text
from organizer import (
    create_output_folder,
    sanitize_filename,
    move_files_to_folder,
)

logger = logging.getLogger(__name__)

SAVE_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv", ".txt"}


def run_export(config: dict, progress_callback=None) -> list:
    """
    Main export loop. Intended to be called from a background thread.

    Args:
        config: Config dict with "output_dir" and "folders" keys.
        progress_callback: Optional callable(current, total, subject).

    Returns:
        List of ExportResult dicts.
    """
    _com_initialized = False  # must be set before try so finally block is safe
    try:
        import pythoncom
        pythoncom.CoInitialize()
        _com_initialized = True
    except ImportError:
        pass

    results = []
    output_dir = config.get("output_dir", "")
    folders = config.get("folders", {})

    active_folders = {
        path: mode for path, mode in folders.items() if mode != "skip"
    }

    if not active_folders:
        logger.info("No folders selected for export (all set to skip).")
        return results

    # Pre-count total items for progress reporting
    total = _count_total_items(active_folders)
    logger.info("Starting export: %d folders, ~%d items total.", len(active_folders), total)

    current = 0

    try:
        for folder_path, mode in active_folders.items():
            # Derive a relative path for mirroring folder structure in output
            # Use the path after the first segment (mailbox name)
            segments = folder_path.split("\\")
            folder_rel_path = "\\".join(segments[1:]) if len(segments) > 1 else segments[0]

            logger.info("Processing folder: %s (mode=%s)", folder_path, mode)

            try:
                for item in get_mail_items(folder_path):
                    current += 1
                    subject = "(no subject)"
                    try:
                        subject = getattr(item, "Subject", "(no subject)") or "(no subject)"
                    except Exception:
                        pass

                    if progress_callback:
                        try:
                            progress_callback(current, total, subject)
                        except Exception:
                            pass

                    try:
                        result = process_mail_item(
                            item, mode, output_dir, folder_rel_path
                        )
                        results.append(result)
                        if result.get("errors"):
                            logger.error(
                                "Errors processing '%s': %s",
                                subject, result["errors"]
                            )
                        else:
                            logger.debug("Exported: %s → %s", subject, result.get("output_folder"))
                    except Exception as e:
                        logger.error("Failed to process item '%s': %s", subject, e)
                        results.append({
                            "subject": subject,
                            "received_time": None,
                            "amount": None,
                            "output_folder": None,
                            "saved_files": [],
                            "errors": [str(e)],
                        })

            except FolderNotFoundError as e:
                logger.warning("Folder not found, skipping: %s — %s", folder_path, e)
            except Exception as e:
                logger.error("Unexpected error processing folder %s: %s", folder_path, e)
    finally:
        if _com_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    logger.info(
        "Export complete. %d items processed, %d errors.",
        len(results),
        sum(1 for r in results if r.get("errors"))
    )
    return results


def _count_total_items(active_folders: dict) -> int:
    """
    Attempt to count total items across all active folders for progress reporting.
    Returns 0 if counting fails.
    """
    total = 0
    try:
        app = get_outlook_app()
        namespace = app.GetNamespace("MAPI")
        from outlook_connector import find_folder_by_path
        for folder_path in active_folders:
            try:
                folder = find_folder_by_path(namespace, folder_path)
                total += folder.Items.Count
            except Exception:
                pass
    except Exception:
        pass
    return total


def process_mail_item(item, mode: str, output_base_dir: str, folder_rel_path: str) -> dict:
    """
    Process a single COM MailItem.

    Returns an ExportResult dict:
      {
        "subject": str,
        "received_time": datetime or None,
        "amount": float or None,
        "output_folder": str or None,
        "saved_files": list[str],
        "errors": list[str],
      }
    """
    errors = []
    staged_files = []
    body_text = ""
    subject = "(no subject)"
    received_time = datetime.now()

    try:
        subject = getattr(item, "Subject", "(no subject)") or "(no subject)"
    except Exception as e:
        errors.append(f"Could not read subject: {e}")

    try:
        rt = getattr(item, "ReceivedTime", None)
        if rt is not None:
            received_time = _com_time_to_datetime(rt)
    except Exception as e:
        errors.append(f"Could not read ReceivedTime: {e}")

    staging_dir = tempfile.mkdtemp(prefix="email_export_")
    try:
        # Save body
        if mode in ("body", "both"):
            try:
                html_body = getattr(item, "HTMLBody", "") or ""
                plain_body = getattr(item, "Body", "") or ""
                body_text = html_to_text(html_body) if html_body else plain_body
                if body_text:
                    body_path = os.path.join(staging_dir, "body.txt")
                    with open(body_path, "w", encoding="utf-8") as f:
                        f.write(body_text)
                    staged_files.append(body_path)
            except Exception as e:
                errors.append(f"Could not save body: {e}")
                logger.warning("Body save failed for '%s': %s", subject, e)

        # Save attachments
        if mode in ("attachments", "both"):
            try:
                saved = save_attachments_to_dir(item, staging_dir)
                staged_files.extend(saved)
            except Exception as e:
                errors.append(f"Could not save attachments: {e}")
                logger.warning("Attachment save failed for '%s': %s", subject, e)

        # Extract amount
        amount = extract_best_amount(staged_files, body_text)

        # Create final output folder
        try:
            final_folder = create_output_folder(
                output_base_dir, folder_rel_path, received_time, amount
            )
        except Exception as e:
            errors.append(f"Could not create output folder: {e}")
            return {
                "subject": subject,
                "received_time": received_time,
                "amount": amount,
                "output_folder": None,
                "saved_files": [],
                "errors": errors,
            }

        # Move staged files to final folder
        final_files = move_files_to_folder(staged_files, final_folder)

        return {
            "subject": subject,
            "received_time": received_time,
            "amount": amount,
            "output_folder": final_folder,
            "saved_files": final_files,
            "errors": errors,
        }
    finally:
        # Clean up the staging temp dir (files were moved, not copied)
        try:
            shutil.rmtree(staging_dir, ignore_errors=True)
        except Exception:
            pass


def html_to_text(html_body: str) -> str:
    """
    Convert HTML email body to plain text.
    Primary: html2text library.
    Fallback: BeautifulSoup get_text().
    Returns empty string if both fail.
    """
    if not html_body:
        return ""
    try:
        import html2text
        converter = html2text.HTML2Text()
        converter.ignore_links = False
        converter.ignore_images = True
        return converter.handle(html_body)
    except Exception:
        pass

    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_body, "html.parser")
        return soup.get_text(separator=" ")
    except Exception:
        pass

    return ""


def extract_best_amount(file_paths: list, body_text: str):
    """
    Extract the largest dollar amount found across all provided files and body text.
    Returns float or None.
    """
    amounts = []

    for path in file_paths:
        try:
            amt = extract_amount_from_file(path)
            if amt is not None:
                amounts.append(amt)
        except Exception as e:
            logger.debug("Amount extraction failed for %s: %s", path, e)

    try:
        amt = extract_amount_from_text(body_text)
        if amt is not None:
            amounts.append(amt)
    except Exception as e:
        logger.debug("Amount extraction from body text failed: %s", e)

    if amounts:
        return max(amounts)
    return None


def _unique_dest_path(dest_path: str) -> str:
    """
    Return dest_path if it doesn't exist; otherwise return the first free
    path with a _2, _3, … suffix inserted before the extension.
    """
    if not os.path.exists(dest_path):
        return dest_path
    from pathlib import Path
    p = Path(dest_path)
    stem, suffix = p.stem, p.suffix
    counter = 2
    while True:
        candidate = p.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return str(candidate)
        counter += 1


def save_attachments_to_dir(item, staging_dir: str) -> list:
    """
    Iterate item.Attachments (1-based COM collection).
    Save each attachment to staging_dir.
    Returns list of saved full paths.
    """
    saved = []
    try:
        attachments = item.Attachments
        count = attachments.Count
    except Exception as e:
        logger.warning("Could not access attachments: %s", e)
        return saved

    for i in range(1, count + 1):
        try:
            attachment = attachments.Item(i)
            filename = getattr(attachment, "FileName", "") or f"attachment_{i}"
            safe_name = sanitize_filename(filename) or f"attachment_{i}"

            dest_path = _unique_dest_path(os.path.join(staging_dir, safe_name))
            attachment.SaveAsFile(dest_path)
            saved.append(dest_path)
            logger.debug("Saved attachment: %s", dest_path)
        except Exception as e:
            logger.warning("Failed to save attachment %d: %s", i, e)

    return saved


def _com_time_to_datetime(com_time) -> datetime:
    """
    Convert a COM/win32 time value to a Python datetime.
    COM times are typically returned as datetime objects directly by pywin32,
    but may sometimes be pywintypes.datetime objects. Handle both.
    """
    if isinstance(com_time, datetime):
        return com_time
    try:
        # pywintypes.datetime is a subclass of datetime, this should work
        return datetime(
            com_time.year, com_time.month, com_time.day,
            com_time.hour, com_time.minute, com_time.second
        )
    except Exception:
        return datetime.now()
