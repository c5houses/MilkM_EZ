"""Auto-updater: check GitHub Releases and apply pending EXE updates."""

import os
import shutil
import sys
import threading
import tkinter.messagebox as messagebox
from pathlib import Path
from typing import Optional

import requests

GITHUB_REPO = "c5houses/MilkM_EZ"

# Read current version from the VERSION file bundled with the app
def _read_version() -> str:
    try:
        ver_path = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "VERSION"
        return ver_path.read_text(encoding="utf-8").strip()
    except Exception:  # noqa: BLE001
        return "0.0.0"


CURRENT_VERSION: str = _read_version()

# ---------------------------------------------------------------------------
# Update helpers
# ---------------------------------------------------------------------------


def _version_tuple(ver: str):
    """Convert a semver string (optionally prefixed with 'v') to a tuple of ints."""
    ver = ver.lstrip("vV")
    try:
        return tuple(int(x) for x in ver.split("."))
    except ValueError:
        return (0, 0, 0)


def check_for_update() -> Optional[tuple]:
    """Check GitHub Releases for a newer version.

    Returns ``(tag_name, download_url)`` if an update is available, else ``None``.
    Does not raise on network errors.
    """
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        latest_tag = data.get("tag_name", "")
        if _version_tuple(latest_tag) > _version_tuple(CURRENT_VERSION):
            # Find the EXE asset
            for asset in data.get("assets", []):
                if asset["name"].lower().endswith(".exe"):
                    return latest_tag, asset["browser_download_url"]
    except Exception:  # noqa: BLE001
        pass
    return None


def download_and_apply_update(download_url: str) -> None:
    """Download the new EXE from *download_url* and save it as ``<current>.new``.

    The actual replacement happens the next time ``apply_pending_update()`` is
    called (i.e. on next startup).
    """
    current_exe = Path(sys.executable)
    pending = current_exe.with_suffix(".new")
    resp = requests.get(download_url, stream=True, timeout=60)
    resp.raise_for_status()
    with open(pending, "wb") as fh:
        for chunk in resp.iter_content(chunk_size=65536):
            fh.write(chunk)


def apply_pending_update() -> None:
    """Replace the current EXE with a pending ``.new`` file if present.

    Called at startup.  Only runs when the app is frozen as an EXE.
    """
    if not getattr(sys, "frozen", False):
        return
    current_exe = Path(sys.executable)
    pending = current_exe.with_suffix(".new")
    if not pending.exists():
        return
    try:
        backup = current_exe.with_suffix(".bak")
        shutil.copy2(str(current_exe), str(backup))
        shutil.move(str(pending), str(current_exe))
    except Exception as exc:  # noqa: BLE001
        print(f"[updater] apply_pending_update failed: {exc}")


# ---------------------------------------------------------------------------
# Background check (called from GUI)
# ---------------------------------------------------------------------------


def _background_check():
    result = check_for_update()
    if result is None:
        return
    tag, url = result
    answer = messagebox.askyesno(
        "Update Available",
        f"A new version ({tag}) is available.\n\n"
        f"Current version: {CURRENT_VERSION}\n\n"
        "Download and install now? The app will restart automatically.",
    )
    if answer:
        try:
            download_and_apply_update(url)
            messagebox.showinfo(
                "Update Ready",
                "Update downloaded. Please restart the application to apply it.",
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Update Failed", str(exc))


def start_background_update_check() -> None:
    """Launch the update check in a daemon thread so the GUI is not blocked."""
    if not getattr(sys, "frozen", False):
        return
    t = threading.Thread(target=_background_check, daemon=True)
    t.start()
