# Data Entry Automation

A Windows desktop application that automates two daily workflows:

1. **Portal Export** — logs in to the [Milk Moovement](https://milkmoovement.com) portal with Selenium, selects your region, and exports the last-two-weeks pickups CSV.
2. **EZFeed Import** — opens the local EZFeed program (`C:\EZfeed4W\EzFeed CSharp.exe`) with PyAutoGUI and imports the downloaded CSV.

A Tkinter GUI lets you enter credentials, pick your region, schedule a daily run, and monitor status. The app can be packaged as a standalone Windows EXE with PyInstaller.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Python 3.10 or newer | <https://python.org> |
| Microsoft Edge **or** Google Chrome | WebDriver is downloaded automatically via `webdriver-manager` |
| EZFeed installed at `C:\EZfeed4W\` | Default path; update `EZFEED_PATH_DEFAULT` in `ezfeed_import.py` if different |
| Windows 10 / 11 | Required for PyAutoGUI desktop automation and Windows Credential Manager |

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Asset Setup

PyAutoGUI uses image recognition to click buttons inside EZFeed. You must supply **7 PNG screenshots** in the `assets/` folder. Take them on the machine where EZFeed is installed:

| File | What to capture |
|---|---|
| `assets/pens.png` | The red "P" (Pens) icon in EZFeed |
| `assets/milkweights.png` | The "Milk Weights" tab |
| `assets/import.png` | The "Import Milk From Processor" link |
| `assets/browse.png` | The "Browse…" button in the import dialog |
| `assets/ok.png` | The "OK" confirmation button |
| `assets/ezfeed_username_dropdown.png` | The blue username combobox with ▼ arrow on the login screen |
| `assets/ezfeed_login.png` | The Login button (with its border) |

Keep captures tight around the target element to improve matching accuracy.

---

## Running from Source

```bash
python app.py
```

For a headless run (e.g. invoked by Task Scheduler):

```bash
python app.py --run
```

> **Note:** The first time you click **Run Now**, credentials are saved to Windows Credential Manager so that `--run` can retrieve them automatically.

---

## Building the Windows EXE

```bash
pyinstaller --onefile --windowed --name "DataEntryAutomation" --add-data "assets;assets" --add-data "VERSION;." app.py
```

The output EXE will be in `dist/DataEntryAutomation.exe`. Copy it (and the `assets/` folder) to the target machine.

---

## Scheduling (Daily Runs)

1. Enter the desired run time in the **Schedule** section of the GUI.
2. Click **Create/Update Schedule**.

This creates a Windows Task Scheduler task named `DataEntryAutomation` that runs the app with `--run` daily at the specified time. The task runs only when the user is logged on (required for PyAutoGUI desktop interaction).

To remove the schedule, click **Remove Schedule**.

---

## Auto-Update

When running as a packaged EXE the app checks the [GitHub Releases](https://github.com/c5houses/MilkM_EZ/releases) page on startup. If a newer version is found you will be prompted to download and install it. The update is applied automatically on the next launch.

---

## Notes

- PyAutoGUI requires an **interactive desktop session**. It will not work over a headless remote desktop where the screen is locked.
- EZFeed does **not** require administrator privileges.
- Download logs and debug screenshots are saved to `%LOCALAPPDATA%\DataEntryAutomation\` (or next to the EXE when packaged).