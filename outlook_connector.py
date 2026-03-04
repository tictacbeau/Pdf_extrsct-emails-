"""
outlook_connector.py — win32com COM automation for Outlook.

IMPORTANT: This module requires:
  - Microsoft Outlook desktop app installed and signed in
  - pywin32 installed (pip install pywin32)
  - Windows OS

COM threading note: Any background thread that uses COM objects must call
pythoncom.CoInitialize() before using any COM objects and pythoncom.CoUninitialize()
when done (in a finally block).

olMailItem constant = 43 (OlItemType.olMailItem)
"""

import logging

logger = logging.getLogger(__name__)

OL_MAIL_ITEM = 43  # OlItemType.olMailItem constant


class OutlookConnectionError(Exception):
    """Raised when a COM connection to Outlook cannot be established."""


class FolderNotFoundError(Exception):
    """Raised when a folder path cannot be found in the Outlook namespace."""


def is_outlook_running() -> bool:
    """
    Check if Outlook is currently running without starting it.
    Uses GetActiveObject which only succeeds if the app is already open.
    Returns True if Outlook is running, False otherwise.
    """
    try:
        import win32com.client
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


def get_outlook_app():
    """
    Return a win32com Outlook.Application object.
    Uses Dispatch which will start Outlook if not already running.
    Raises OutlookConnectionError if COM fails.
    """
    try:
        import win32com.client
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise OutlookConnectionError(
            f"Could not connect to Outlook via COM: {e}"
        ) from e


def get_folder_tree() -> list:
    """
    Return a flat list of folder_node dicts representing ALL Outlook folders
    (all top-level accounts and their subfolders).

    Each folder_node has:
      {
        "name": str,          # folder display name
        "full_path": str,     # backslash-joined path from mailbox root
        "depth": int,         # nesting depth for UI indentation
        "children": list      # list of child folder_node dicts
      }

    The returned list is flat (all nodes at all depths), ordered depth-first.
    Children are also embedded in each node for tree rendering.
    """
    app = get_outlook_app()
    namespace = app.GetNamespace("MAPI")

    all_nodes = []
    for i in range(1, namespace.Folders.Count + 1):
        try:
            top_folder = namespace.Folders.Item(i)
            node = _recurse_folder(top_folder, depth=0, parent_path="")
            all_nodes.append(node)
            _flatten_into(node, all_nodes)
        except Exception as e:
            logger.warning("Error reading top-level folder at index %d: %s", i, e)

    return all_nodes


def _flatten_into(node: dict, result: list) -> None:
    """Recursively add all children of node into result list (depth-first)."""
    for child in node.get("children", []):
        result.append(child)
        _flatten_into(child, result)


def _recurse_folder(folder, depth: int, parent_path: str) -> dict:
    """
    Build a folder_node dict for one COM folder object.
    Recursively builds children list.
    """
    name = folder.Name
    if parent_path:
        full_path = parent_path + "\\" + name
    else:
        full_path = name

    children = []
    try:
        sub_folders = folder.Folders
        for i in range(1, sub_folders.Count + 1):
            try:
                child_folder = sub_folders.Item(i)
                child_node = _recurse_folder(child_folder, depth + 1, full_path)
                children.append(child_node)
            except Exception as e:
                logger.debug("Error reading subfolder at index %d under %s: %s",
                             i, full_path, e)
    except Exception as e:
        logger.debug("Error accessing subfolders of %s: %s", full_path, e)

    return {
        "name": name,
        "full_path": full_path,
        "depth": depth,
        "children": children,
    }


def find_folder_by_path(namespace, full_path: str):
    """
    Navigate COM namespace.Folders using backslash-split path segments.
    Returns the COM folder object, or raises FolderNotFoundError.

    Example:
      full_path = "Mailbox - Jane\\Inbox\\Orders"
      Navigates: namespace.Folders["Mailbox - Jane"] -> .Folders["Inbox"] -> .Folders["Orders"]
    """
    segments = full_path.split("\\")
    if not segments:
        raise FolderNotFoundError(f"Empty folder path: {full_path!r}")

    # Find the top-level folder (mailbox/account)
    top_name = segments[0]
    current_folder = None
    for i in range(1, namespace.Folders.Count + 1):
        try:
            f = namespace.Folders.Item(i)
            if f.Name == top_name:
                current_folder = f
                break
        except Exception:
            continue

    if current_folder is None:
        raise FolderNotFoundError(
            f"Top-level folder not found: {top_name!r} in path {full_path!r}"
        )

    # Navigate through remaining segments
    for segment in segments[1:]:
        found = None
        try:
            sub = current_folder.Folders
            for i in range(1, sub.Count + 1):
                try:
                    f = sub.Item(i)
                    if f.Name == segment:
                        found = f
                        break
                except Exception:
                    continue
        except Exception as e:
            raise FolderNotFoundError(
                f"Error navigating to {segment!r} in path {full_path!r}: {e}"
            ) from e

        if found is None:
            raise FolderNotFoundError(
                f"Subfolder not found: {segment!r} in path {full_path!r}"
            )
        current_folder = found

    return current_folder


def get_mail_items(folder_full_path: str):
    """
    Given a full_path like "Mailbox - Jane\\Inbox\\Orders",
    navigate the COM folder tree and iterate its Items collection.

    Yields raw COM MailItem objects (item.Class == OL_MAIL_ITEM = 43).
    Skips non-mail items (MeetingItem, ContactItem, etc.).
    """
    app = get_outlook_app()
    namespace = app.GetNamespace("MAPI")

    folder = find_folder_by_path(namespace, folder_full_path)

    try:
        items = folder.Items
        # Sort by received time ascending for consistent processing order
        try:
            items.Sort("[ReceivedTime]", False)
        except Exception:
            pass  # Sort is best-effort

        count = items.Count
        logger.info("Iterating %d items in folder: %s", count, folder_full_path)

        for i in range(1, count + 1):
            try:
                item = items.Item(i)
                if item.Class == OL_MAIL_ITEM:
                    yield item
            except Exception as e:
                logger.warning(
                    "Error accessing item %d in %s: %s", i, folder_full_path, e
                )
    except Exception as e:
        raise FolderNotFoundError(
            f"Error accessing items in folder {folder_full_path!r}: {e}"
        ) from e
