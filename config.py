"""Shared configuration: paths, directory creation, and asset resolution."""

import os
import sys
from pathlib import Path

APP_NAME = "DataEntryAutomation"


def is_frozen_exe() -> bool:
    """Return True when running as a PyInstaller-packaged EXE."""
    return getattr(sys, "frozen", False)


def get_app_root_dir() -> Path:
    """Return a persistent folder for app data.

    - Frozen (EXE): a folder named APP_NAME next to the EXE.
    - Source: %LOCALAPPDATA%\\DataEntryAutomation
    """
    if is_frozen_exe():
        return Path(sys.executable).parent / APP_NAME
    local_app_data = os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
    return Path(local_app_data) / APP_NAME


def get_download_dir() -> Path:
    """Return (and create) the downloads sub-folder."""
    d = get_app_root_dir() / "downloads"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_export_csv_path() -> Path:
    """Return the canonical path for the exported CSV."""
    return get_download_dir() / "pickups_last2weeks.csv"


def resource_path(rel_path: str) -> Path:
    """Resolve a path for bundled assets.

    Works both when running from source and when packaged with PyInstaller
    (which extracts resources to ``sys._MEIPASS``).
    """
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        # Running from source — assets are next to this file
        base = Path(__file__).parent
    else:
        base = Path(base)
    return base / rel_path
