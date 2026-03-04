"""
organizer.py — Output folder naming and creation.

Handles:
- Folder name formatting: YYYY-MM-DD_$X,XXX.XX or YYYY-MM-DD_unknown-amount
- Windows-safe name sanitization
- Duplicate folder resolution
- File moving
"""

import os
import re
import shutil
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

INVALID_CHARS_PATTERN = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MAX_FOLDER_NAME_LENGTH = 100  # conservative limit below Windows 260-char path limit


def format_amount_folder_name(received_time: datetime, amount) -> str:
    """
    Returns folder name string:
      amount is not None → "YYYY-MM-DD_$X,XXX.XX"  e.g. "2026-03-04_$1,234.56"
      amount is None     → "YYYY-MM-DD_unknown-amount"
    """
    date_str = received_time.strftime("%Y-%m-%d")
    if amount is not None:
        return f"{date_str}_${amount:,.2f}"
    return f"{date_str}_unknown-amount"


def sanitize_folder_name(name: str) -> str:
    """
    Replace invalid Windows path chars with '_'.
    Strip leading/trailing spaces and dots.
    Truncate to MAX_FOLDER_NAME_LENGTH characters.
    """
    sanitized = INVALID_CHARS_PATTERN.sub("_", name)
    sanitized = sanitized.strip(" .")
    sanitized = sanitized[:MAX_FOLDER_NAME_LENGTH]
    sanitized = sanitized.strip(" .")
    if not sanitized:
        return "_unnamed"
    return sanitized


def sanitize_path_segment(path: str) -> str:
    """
    Apply sanitize_folder_name() to each backslash or forward-slash separated
    segment of a relative path, then rejoin with os.sep.
    """
    segments = re.split(r'[/\\]', path)
    sanitized_segments = [sanitize_folder_name(s) for s in segments if s]
    return os.sep.join(sanitized_segments)


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a filename while preserving the file extension.
    Splits on the last '.', sanitizes stem and extension separately.
    """
    if "." in filename:
        stem, _, ext = filename.rpartition(".")
        safe_stem = sanitize_folder_name(stem)
        safe_ext = INVALID_CHARS_PATTERN.sub("_", ext).strip()
        if safe_ext:
            return f"{safe_stem}.{safe_ext}"
        return safe_stem
    return sanitize_folder_name(filename)


def resolve_duplicate_folder(folder_path: str) -> str:
    """
    If folder_path does not exist, return it unchanged.
    If it exists, try folder_path + "_2", "_3", ... until a free name is found.
    """
    if not os.path.exists(folder_path):
        return folder_path
    counter = 2
    while True:
        candidate = f"{folder_path}_{counter}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def create_output_folder(
    output_base_dir: str,
    folder_rel_path: str,
    received_time: datetime,
    amount
) -> str:
    """
    Constructs and creates:
      <output_base_dir>/<sanitized_rel_path>/<date_amount_folder>

    Handles duplicate folder names by appending _2, _3, etc.
    Returns the final created folder path.
    """
    safe_rel = sanitize_path_segment(folder_rel_path) if folder_rel_path else ""
    # format_amount_folder_name produces only safe chars (digits, hyphens, underscores, $, comma, dot)
    safe_date_folder = format_amount_folder_name(received_time, amount)

    if safe_rel:
        base_path = os.path.join(output_base_dir, safe_rel, safe_date_folder)
    else:
        base_path = os.path.join(output_base_dir, safe_date_folder)

    final_path = resolve_duplicate_folder(base_path)

    try:
        os.makedirs(final_path, exist_ok=True)
    except OSError as e:
        logger.error("Failed to create output folder %s: %s", final_path, e)
        raise

    return final_path


def move_files_to_folder(file_paths: list, destination_folder: str) -> list:
    """
    Move each file in file_paths to destination_folder using shutil.move().
    Returns list of final destination paths.
    Logs any individual file move errors but continues processing.
    """
    result = []
    for src in file_paths:
        try:
            dst = shutil.move(src, destination_folder)
            result.append(dst)
        except Exception as e:
            logger.error("Failed to move %s to %s: %s", src, destination_folder, e)
    return result
