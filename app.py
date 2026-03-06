"""Tkinter GUI entry point for the Data Entry Automation app."""

import argparse
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import keyring
import keyring.errors

import config
import updater
from ezfeed_import import run_ezfeed_import
from portal_export import run_portal_export
from scheduler import Schedule, create_or_update_daily_task, delete_task

# ---------------------------------------------------------------------------
# Elevation helper
# ---------------------------------------------------------------------------


def _ensure_elevated():
    """Re-launch this script as admin if we're not already elevated.

    Only auto-elevates on Windows. On other platforms this is a no-op.
    """
    import os
    import sys

    if os.name != 'nt':
        return

    import ctypes
    try:
        is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        is_admin = False

    if is_admin:
        return  # Already running elevated

    # Re-launch ourselves with elevation
    import subprocess
    params = subprocess.list2cmdline(sys.argv)
    executable = sys.executable

    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", executable, params, None, 1
    )
    if ret > 32:
        sys.exit(0)  # Successfully re-launched elevated; exit this instance
    # If elevation was declined/failed, continue without it
    print("WARNING: Could not elevate to admin. EZFeed launch may fail.")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_NAME = config.APP_NAME
KEYRING_SERVICE = APP_NAME

REGIONS = [
    "Bongards Creameries",
    "California Dairies Inc.",
    "Cargill Feed Shop",
    "Dairy Farmers of Newfoundland and Labrador",
    "Erie Cooperative Association (TCJ)",
    "Great Plains Dairymen's Association (TCJ)",
    "Idaho Milk Products",
    "KYTN Cooperative (TCJ)",
    "Legacy Milk",
    "Liberty Milk Producers (TCJ)",
    "Michigan Milk Producers Association",
    "Minerva Dairy",
    "Nebraska Milk Producers (TCJ)",
    "Plainview Milk Products",
    "Prairie Farms",
    "United Dairymen of Arizona",
    "Upstate Niagara Cooperative Inc.",
    "White Eagle Cooperative (TCJ)",
]

DEFAULT_REGION = "California Dairies Inc."

# ---------------------------------------------------------------------------
# Keyring helpers
# ---------------------------------------------------------------------------


def _save_credentials(portal_user: str, portal_pass: str, region: str, ezfeed_user: str, ezfeed_password: str = "") -> None:
    keyring.set_password(KEYRING_SERVICE, "portal_username", portal_user)
    keyring.set_password(KEYRING_SERVICE, "portal_password", portal_pass)
    keyring.set_password(KEYRING_SERVICE, "portal_region", region)
    keyring.set_password(KEYRING_SERVICE, "ezfeed_username", ezfeed_user)
    keyring.set_password(KEYRING_SERVICE, "ezfeed_password", ezfeed_password)


def _clear_credentials() -> None:
    for key in ("portal_username", "portal_password", "portal_region", "ezfeed_username", "ezfeed_password"):
        try:
            keyring.delete_password(KEYRING_SERVICE, key)
        except keyring.errors.PasswordDeleteError:
            pass


def _load_credentials() -> dict:
    return {
        "portal_username": keyring.get_password(KEYRING_SERVICE, "portal_username") or "",
        "portal_password": keyring.get_password(KEYRING_SERVICE, "portal_password") or "",
        "portal_region": keyring.get_password(KEYRING_SERVICE, "portal_region") or DEFAULT_REGION,
        "ezfeed_username": keyring.get_password(KEYRING_SERVICE, "ezfeed_username") or "",
        "ezfeed_password": keyring.get_password(KEYRING_SERVICE, "ezfeed_password") or "",
    }


# ---------------------------------------------------------------------------
# Headless / scheduled run
# ---------------------------------------------------------------------------


def run_headless() -> None:
    """Run the full automation without a GUI (used by Task Scheduler via --run)."""
    creds = _load_credentials()
    if not creds["portal_username"] or not creds["portal_password"]:
        print("ERROR: No saved credentials found. Open the GUI and click 'Run Now' first.")
        sys.exit(1)

    print("=== Headless run started ===")
    csv_path = run_portal_export(
        creds["portal_username"],
        creds["portal_password"],
        creds["portal_region"],
    )
    run_ezfeed_import(csv_path, creds["ezfeed_username"], ezfeed_password=creds["ezfeed_password"])
    print("=== Headless run complete ===")


# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Entry Automation")
        self.geometry("560x580")
        self.resizable(False, False)

        self._build_ui()
        self._populate_saved_credentials()

        # Apply any pending update (only when frozen)
        updater.apply_pending_update()

        # Check for updates in background (only when frozen)
        updater.start_background_update_check()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        # ---- Milk Moovement Portal frame ----
        portal_frame = ttk.LabelFrame(self, text="Milk Moovement Portal")
        portal_frame.pack(fill="x", padx=12, pady=6)

        ttk.Label(portal_frame, text="Region:").grid(row=0, column=0, sticky="w", **pad)
        self.region_var = tk.StringVar(value=DEFAULT_REGION)
        self.region_combo = ttk.Combobox(
            portal_frame,
            textvariable=self.region_var,
            values=REGIONS,
            state="readonly",
            width=38,
        )
        self.region_combo.grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(portal_frame, text="Username:").grid(row=1, column=0, sticky="w", **pad)
        self.portal_user_var = tk.StringVar()
        ttk.Entry(portal_frame, textvariable=self.portal_user_var, width=40).grid(
            row=1, column=1, sticky="w", **pad
        )

        ttk.Label(portal_frame, text="Password:").grid(row=2, column=0, sticky="w", **pad)
        self.portal_pass_var = tk.StringVar()
        ttk.Entry(portal_frame, textvariable=self.portal_pass_var, show="*", width=40).grid(
            row=2, column=1, sticky="w", **pad
        )

        # ---- EZFeed frame ----
        ezfeed_frame = ttk.LabelFrame(self, text="EZFeed (local program)")
        ezfeed_frame.pack(fill="x", padx=12, pady=6)

        ttk.Label(ezfeed_frame, text="Username:").grid(row=0, column=0, sticky="w", **pad)
        self.ezfeed_user_var = tk.StringVar()
        ttk.Entry(ezfeed_frame, textvariable=self.ezfeed_user_var, width=40).grid(
            row=0, column=1, sticky="w", **pad
        )

        ttk.Label(ezfeed_frame, text="Password:").grid(row=1, column=0, sticky="w", **pad)
        self.ezfeed_pass_var = tk.StringVar()
        ttk.Entry(ezfeed_frame, textvariable=self.ezfeed_pass_var, show="*", width=40).grid(
            row=1, column=1, sticky="w", **pad
        )

        ttk.Label(
            ezfeed_frame,
            text="(Leave blank if security is not enabled)",
            foreground="gray",
        ).grid(row=2, column=0, columnspan=2, sticky="w", **pad)

        # ---- Schedule frame ----
        sched_frame = ttk.LabelFrame(self, text="Schedule (optional)")
        sched_frame.pack(fill="x", padx=12, pady=6)

        ttk.Label(sched_frame, text="Time (HH:MM):").grid(row=0, column=0, sticky="w", **pad)
        self.sched_time_var = tk.StringVar(value="07:00")
        ttk.Entry(sched_frame, textvariable=self.sched_time_var, width=10).grid(
            row=0, column=1, sticky="w", **pad
        )
        ttk.Button(sched_frame, text="Create/Update Schedule", command=self._on_create_schedule).grid(
            row=0, column=2, **pad
        )
        ttk.Button(sched_frame, text="Remove Schedule", command=self._on_delete_schedule).grid(
            row=0, column=3, **pad
        )

        # ---- Credentials buttons ----
        cred_frame = ttk.Frame(self)
        cred_frame.pack(pady=4)
        ttk.Button(cred_frame, text="Save Credentials", command=self._on_save_credentials).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(cred_frame, text="Clear Saved Credentials", command=self._on_clear_credentials).grid(
            row=0, column=1, padx=10
        )

        # ---- Status label ----
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, foreground="blue").pack(pady=4)

        # ---- Action buttons ----
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=6)
        self.run_btn = ttk.Button(btn_frame, text="Run Now", command=self._on_run_now)
        self.run_btn.grid(row=0, column=0, padx=10)
        ttk.Button(btn_frame, text="Quit", command=self.destroy).grid(row=0, column=1, padx=10)

    # ------------------------------------------------------------------
    # Credential management
    # ------------------------------------------------------------------

    def _populate_saved_credentials(self):
        creds = _load_credentials()
        self.portal_user_var.set(creds["portal_username"])
        self.portal_pass_var.set(creds["portal_password"])
        if creds["portal_region"] in REGIONS:
            self.region_var.set(creds["portal_region"])
        self.ezfeed_user_var.set(creds["ezfeed_username"])
        self.ezfeed_pass_var.set(creds["ezfeed_password"])

    def _on_save_credentials(self):
        portal_user = self.portal_user_var.get().strip()
        portal_pass = self.portal_pass_var.get().strip()
        region = self.region_var.get().strip()
        ezfeed_user = self.ezfeed_user_var.get().strip()
        ezfeed_pass = self.ezfeed_pass_var.get().strip()
        _save_credentials(portal_user, portal_pass, region, ezfeed_user, ezfeed_pass)
        messagebox.showinfo("Credentials", "Credentials saved successfully.")

    def _on_clear_credentials(self):
        _clear_credentials()
        self.portal_user_var.set("")
        self.portal_pass_var.set("")
        self.region_var.set(DEFAULT_REGION)
        self.ezfeed_user_var.set("")
        self.ezfeed_pass_var.set("")
        messagebox.showinfo("Credentials", "Saved credentials cleared.")

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _on_run_now(self):
        portal_user = self.portal_user_var.get().strip()
        portal_pass = self.portal_pass_var.get().strip()
        region = self.region_var.get().strip()
        ezfeed_user = self.ezfeed_user_var.get().strip()
        ezfeed_pass = self.ezfeed_pass_var.get().strip()

        if not portal_user or not portal_pass:
            messagebox.showwarning("Missing credentials", "Please enter your portal username and password.")
            return
        if not region:
            messagebox.showwarning("Missing region", "Please select a region.")
            return

        # Save credentials for headless/scheduled runs
        _save_credentials(portal_user, portal_pass, region, ezfeed_user, ezfeed_pass)

        self.run_btn.config(state="disabled")
        threading.Thread(
            target=self._run_automation,
            args=(portal_user, portal_pass, region, ezfeed_user, ezfeed_pass),
            daemon=True,
        ).start()

    def _run_automation(self, portal_user: str, portal_pass: str, region: str, ezfeed_user: str, ezfeed_pass: str):
        try:
            self._set_status("Exporting from portal …")
            csv_path = run_portal_export(portal_user, portal_pass, region)

            self._set_status("Importing into EZFeed …")
            run_ezfeed_import(csv_path, ezfeed_user, ezfeed_password=ezfeed_pass)

            self._set_status("Done ✓")
            messagebox.showinfo("Success", "Export and import completed successfully.")
        except Exception as exc:
            self._set_status("Failed ✗")
            messagebox.showerror("Error", str(exc))
        finally:
            self.run_btn.config(state="normal")

    def _on_create_schedule(self):
        time_val = self.sched_time_var.get().strip()
        if not time_val:
            messagebox.showwarning("Missing time", "Please enter a schedule time (HH:MM).")
            return
        try:
            create_or_update_daily_task(Schedule(time_hhmm=time_val))
            messagebox.showinfo("Schedule", f"Daily task scheduled at {time_val}.")
        except Exception as exc:
            messagebox.showerror("Schedule Error", str(exc))

    def _on_delete_schedule(self):
        try:
            delete_task()
            messagebox.showinfo("Schedule", "Scheduled task removed.")
        except Exception as exc:
            messagebox.showerror("Schedule Error", str(exc))

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    def _set_status(self, msg: str):
        """Thread-safe status update."""
        self.after(0, lambda: self.status_var.set(msg))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    _ensure_elevated()

    parser = argparse.ArgumentParser(description="Data Entry Automation")
    parser.add_argument(
        "--run",
        action="store_true",
        help="Run headless (for Task Scheduler — no GUI).",
    )
    args = parser.parse_args()

    if args.run:
        run_headless()
    else:
        app = App()
        app.mainloop()


if __name__ == "__main__":
    main()
