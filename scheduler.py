"""Windows Task Scheduler integration for DataEntryAutomation."""

import subprocess
import sys
from dataclasses import dataclass, field

TASK_NAME = "DataEntryAutomation"


@dataclass
class Schedule:
    time_hhmm: str
    daily: bool = field(default=True)


def _exe_path() -> str:
    """Return the path to the running executable (EXE or python interpreter)."""
    return sys.executable


def create_or_update_daily_task(schedule: Schedule) -> None:
    """Create or update a daily Windows Task Scheduler task.

    The task runs the application with the ``--run`` flag so it executes in
    headless mode (no GUI).  It is configured to run only when the user is
    logged on (required for PyAutoGUI desktop interaction).
    """
    exe = _exe_path()
    # When packaged as EXE the executable itself is the entry point.
    # When running from source we need to pass app.py as argument.
    if exe.lower().endswith(".exe") and "python" not in exe.lower():
        tr = f'"{exe}" --run'
    else:
        import pathlib
        app_script = str(pathlib.Path(__file__).parent / "app.py")
        tr = f'"{exe}" "{app_script}" --run'

    cmd = [
        "schtasks",
        "/Create",
        "/F",
        "/SC", "DAILY",
        "/TN", TASK_NAME,
        "/TR", tr,
        "/ST", schedule.time_hhmm,
        "/RL", "LIMITED",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"schtasks /Create failed (rc={result.returncode}):\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )


def delete_task() -> None:
    """Delete the scheduled task if it exists."""
    cmd = ["schtasks", "/Delete", "/F", "/TN", TASK_NAME]
    subprocess.run(cmd, capture_output=True, text=True)
