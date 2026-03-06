"""PyAutoGUI automation: import a CSV file into the local EZFeed program."""

import subprocess
import time
from collections import deque
from datetime import datetime
from pathlib import Path
from typing import Optional

import pyautogui
import pyperclip

import config

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

EZFEED_PATH_DEFAULT = r"C:\EZfeed4W\EzFeed CSharp.exe"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------


class RunLogger:
    """Structured logger that writes to import.log and keeps the last 20 runs."""

    MAX_RUNS = 20

    def __init__(self):
        log_dir = config.get_app_root_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = log_dir / "import.log"
        self._runs: deque = deque(maxlen=self.MAX_RUNS)
        self._current_run: list = []
        self._load()

    # ------------------------------------------------------------------
    def _load(self):
        """Load existing runs from the log file."""
        if not self.log_path.exists():
            return
        try:
            text = self.log_path.read_text(encoding="utf-8")
            blocks = [b.strip() for b in text.split("---RUN---") if b.strip()]
            for block in blocks[-self.MAX_RUNS :]:
                self._runs.append(block)
        except Exception:  # noqa: BLE001
            pass

    def _flush(self):
        """Write all stored runs to disk."""
        try:
            self.log_path.write_text(
                "\n---RUN---\n".join(self._runs),
                encoding="utf-8",
            )
        except Exception:  # noqa: BLE001
            pass

    # ------------------------------------------------------------------
    def start(self):
        self._current_run = [f"=== Run started {datetime.now().isoformat()} ==="]

    def log(self, message: str):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        self._current_run.append(line)
        print(line)

    def finish(self, success: bool):
        status = "SUCCESS" if success else "FAILED"
        self._current_run.append(f"=== {status} {datetime.now().isoformat()} ===")
        self._runs.append("\n".join(self._current_run))
        self._flush()
        self._current_run = []


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def kill_ezfeed() -> None:
    """Force-kill the EZFeed process if it is running."""
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "EzFeed CSharp.exe"],
            capture_output=True,
        )
    except Exception:  # noqa: BLE001
        pass


def paste_text(text: str) -> None:
    """Copy *text* to clipboard then paste with Ctrl+V."""
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)


def wait_and_click(
    logger: RunLogger,
    img: str,
    desc: str,
    confidence: float = 0.8,
    delay: float = 1.0,
    timeout: float = 30.0,
) -> None:
    """Wait until *img* appears on screen and click its centre.

    *img* is a relative asset path resolved via ``config.resource_path()``.
    Raises ``RuntimeError`` if the image is not found within *timeout* seconds.
    """
    img_path = str(config.resource_path(img))
    logger.log(f"Looking for: {desc} ({img})")
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=confidence)
            if location:
                pyautogui.click(location)
                logger.log(f"Clicked: {desc}")
                time.sleep(delay)
                return
        except Exception:  # noqa: BLE001
            pass
        time.sleep(5)
    raise RuntimeError(f"Image not found on screen within {timeout}s: {desc} ({img})")


# ---------------------------------------------------------------------------
# Username ComboBox selection
# ---------------------------------------------------------------------------


def _select_username_in_combobox(
    logger: RunLogger, ezfeed_username: str, ezfeed_path: str
) -> None:
    """Select *ezfeed_username* in the EZFeed login ComboBox.

    Primary method: pywinauto Windows accessibility API — enumerates the
    ComboBox items and selects the matching one directly (no OCR, no position
    counting, works for any username).

    Fallback (if pywinauto is unavailable or fails to connect): open the
    dropdown with a click, press Home to go to the top of the list, then
    press the first character of the username so Windows jumps to the first
    item starting with that letter, and press Enter to confirm.
    """
    try:
        from pywinauto import Application  # type: ignore  # noqa: PLC0415

        logger.log("Attempting pywinauto ComboBox selection …")
        app: Optional[Application] = None
        for attempt in range(10):
            try:
                app = Application(backend="win32").connect(path=ezfeed_path)
                break
            except Exception as connect_exc:  # noqa: BLE001
                logger.log(f"  pywinauto connect attempt {attempt + 1}/10 failed: {connect_exc}")
                time.sleep(1)

        if app is None:
            raise RuntimeError(
                "Could not connect to EZFeed via pywinauto after 10 attempts."
            )

        dlg = app.top_window()
        combo = dlg.ComboBox
        combo.select(ezfeed_username)
        logger.log(f"pywinauto: selected '{ezfeed_username}'.")
        return

    except Exception as exc:  # noqa: BLE001
        logger.log(
            f"pywinauto selection failed ({exc}); "
            "falling back to arrow-key navigation …"
        )

    # ------------------------------------------------------------------
    # Fallback: open the dropdown and navigate with keyboard keys.
    # Windows ComboBox responds to a single-character keypress by jumping
    # to the first item that starts with that character.
    # ------------------------------------------------------------------
    if not ezfeed_username:
        raise ValueError("ezfeed_username must not be empty.")

    wait_and_click(logger, "assets/ezloginuser.png", "Username dropdown", timeout=30)
    time.sleep(0.5)

    # Go to the very top of the list
    pyautogui.press("home")
    time.sleep(0.2)

    # Jump to the first item starting with the same letter as the username
    first_char = ezfeed_username[0].lower()
    pyautogui.press(first_char)
    time.sleep(0.3)

    # Confirm selection
    pyautogui.press("enter")
    time.sleep(0.5)


# ---------------------------------------------------------------------------
# Elevation-aware launcher
# ---------------------------------------------------------------------------


def _launch_ezfeed(path: str, logger: RunLogger) -> None:
    """Launch EZFeed, handling the case where it requires elevation."""
    import ctypes
    import os

    # Strategy 1: Try normal subprocess launch first (works if app doesn't need admin,
    # or if we're already running as admin)
    try:
        subprocess.Popen([path])
        logger.log("Launched EZFeed via subprocess.")
        return
    except OSError as e:
        if not hasattr(e, 'winerror') or e.winerror != 740:  # 740 = ERROR_ELEVATION_REQUIRED
            raise
        logger.log("EZFeed requires elevation, using ShellExecute with 'runas' …")

    # Strategy 2: Use ShellExecuteW with 'runas' verb for elevation
    # This will show a UAC prompt if the script is not already elevated.
    # When running from Task Scheduler configured with "Run with highest privileges",
    # no UAC prompt will appear.
    ret = ctypes.windll.shell32.ShellExecuteW(
        None,           # hwnd
        "runas",        # lpOperation - request elevation
        path,           # lpFile
        None,           # lpParameters
        os.path.dirname(path) or None,  # lpDirectory - working dir
        1,              # nShowCmd - SW_SHOWNORMAL
    )
    # ShellExecuteW returns a value > 32 on success
    if ret <= 32:
        raise RuntimeError(
            f"ShellExecuteW failed to launch EZFeed (return code {ret}). "
            f"Try running the automation app as Administrator."
        )
    logger.log("Launched EZFeed via ShellExecuteW (elevated).")


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def run_ezfeed_import(
    csv_file: Path,
    ezfeed_username: str,
    ezfeed_path: str = EZFEED_PATH_DEFAULT,
    ezfeed_password: str = "",
) -> None:
    """Open EZFeed, log in, and import *csv_file*.

    Raises an exception on failure.
    """
    logger = RunLogger()
    logger.start()
    logger.log(f"CSV: {csv_file}")
    logger.log(f"EZFeed path: {ezfeed_path}")
    logger.log(f"EZFeed username: {ezfeed_username}")

    try:
        # Make sure any stale instance is gone
        kill_ezfeed()
        time.sleep(1)

        # Launch EZFeed
        logger.log("Launching EZFeed …")
        _launch_ezfeed(ezfeed_path, logger)
        time.sleep(15)

        # ------------------------------------------------------------------
        # Login — dropdown combobox (password optional)
        # ------------------------------------------------------------------
        logger.log("Logging in to EZFeed …")

        # 1. Select the username from the ComboBox
        logger.log(f"Selecting username '{ezfeed_username}' from dropdown …")
        _select_username_in_combobox(logger, ezfeed_username, ezfeed_path)
        time.sleep(0.5)

        # 2. If a password is provided, tab to the password field and enter it
        if ezfeed_password:
            logger.log("Entering EZFeed password …")
            pyautogui.press("tab")
            time.sleep(0.3)
            paste_text(ezfeed_password)
            time.sleep(0.5)

        # 3. Click the Login button
        wait_and_click(logger, "assets/ezlogin.png", "Login button", timeout=30)
        time.sleep(2)

        # ------------------------------------------------------------------
        # Navigation: Pens → Milk Weights → Import Milk From Processor
        # ------------------------------------------------------------------
        wait_and_click(logger, "assets/pens.png", "Pens icon", timeout=30)
        wait_and_click(logger, "assets/milkweights.png", "Milk Weights tab", timeout=30)
        wait_and_click(logger, "assets/import.png", "Import Milk From Processor", timeout=30)

        # ------------------------------------------------------------------
        # Browse to CSV file
        # ------------------------------------------------------------------
        wait_and_click(logger, "assets/browse.png", "Browse button", timeout=30)
        time.sleep(1)

        # Paste the full CSV path into the file-open dialog and confirm
        paste_text(str(csv_file))
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1)

        # ------------------------------------------------------------------
        # Confirm import
        # ------------------------------------------------------------------
        wait_and_click(logger, "assets/ok_import.png", "OK button", timeout=30)

        # Wait for the import to process
        logger.log("Waiting 10 seconds for import to complete …")
        time.sleep(10)

        # ------------------------------------------------------------------
        # Done — close EZFeed
        # ------------------------------------------------------------------
        logger.log("Import complete. Closing EZFeed …")
        kill_ezfeed()

        logger.finish(success=True)
        logger.log("EZFeed import finished successfully.")

    except Exception as exc:
        logger.log(f"ERROR: {exc}")
        kill_ezfeed()
        logger.finish(success=False)
        raise
