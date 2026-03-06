"""Selenium automation: log in to the Milk Moovement portal and export CSV."""

import shutil
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import config

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def norm(text: str) -> str:
    """Normalise unicode and collapse whitespace for fuzzy comparison."""
    return " ".join(unicodedata.normalize("NFKD", text).casefold().split())


def ts() -> str:
    """Return a short timestamp string for log messages."""
    return datetime.now().strftime("%H:%M:%S")


def dump_debug(driver, label: str = "error") -> None:
    """Save a screenshot and page source for post-mortem debugging."""
    try:
        out_dir = config.get_download_dir()
        shot = out_dir / f"debug_{label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        driver.save_screenshot(str(shot))
        print(f"[{ts()}] Debug screenshot saved: {shot}")
    except Exception as exc:  # noqa: BLE001
        print(f"[{ts()}] dump_debug failed: {exc}")


def find_first(driver, selectors: List[tuple]):
    """Return the first visible element that matches one of the (By, value) pairs."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    for by, value in selectors:
        try:
            el = WebDriverWait(driver, 4).until(
                EC.visibility_of_element_located((by, value))
            )
            return el
        except Exception:  # noqa: BLE001
            continue
    return None


def click_first(driver, selectors: List[tuple], description: str = "") -> bool:
    """Click the first matching visible element. Return True on success."""
    el = find_first(driver, selectors)
    if el:
        try:
            el.click()
            print(f"[{ts()}] Clicked: {description or selectors[0][1]}")
            return True
        except Exception as exc:  # noqa: BLE001
            print(f"[{ts()}] click_first click failed ({description}): {exc}")
    print(f"[{ts()}] click_first: element not found — {description or selectors}")
    return False


def wait_for_new_csv(download_dir: Path, before: set, timeout: int = 120) -> Optional[Path]:
    """Wait until a new .csv file appears in *download_dir* and is fully written."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        current = set(download_dir.glob("*.csv"))
        new_files = current - before
        if new_files:
            # Make sure no .crdownload / .tmp companions remain
            tmp = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.tmp"))
            if not tmp:
                return sorted(new_files, key=lambda p: p.stat().st_mtime)[-1]
        time.sleep(1)
    return None


def atomic_replace(src: Path, dst: Path) -> None:
    """Move *src* to *dst*, replacing *dst* if it exists."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(dst))


# ---------------------------------------------------------------------------
# Driver factory
# ---------------------------------------------------------------------------

def _build_driver(download_dir: Path):
    """Try Edge first, fall back to Chrome. Raise if neither is available."""
    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    # --- Edge ---
    try:
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.edge.service import Service as EdgeService
        from selenium import webdriver
        from webdriver_manager.microsoft import EdgeChromiumDriverManager

        opts = EdgeOptions()
        opts.add_experimental_option("prefs", prefs)
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        service = EdgeService(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service, options=opts)
        print(f"[{ts()}] Using Microsoft Edge WebDriver")
        return driver
    except Exception as exc:  # noqa: BLE001
        print(f"[{ts()}] Edge unavailable ({exc}), trying Chrome …")

    # --- Chrome ---
    try:
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.chrome.service import Service as ChromeService
        from selenium import webdriver
        from webdriver_manager.chrome import ChromeDriverManager

        opts = ChromeOptions()
        opts.add_experimental_option("prefs", prefs)
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
        print(f"[{ts()}] Using Google Chrome WebDriver")
        return driver
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError("Neither Edge nor Chrome WebDriver could be initialised.") from exc


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def run_portal_export(username: str, password: str, region_name: str) -> Path:
    """Log in to the Milk Moovement portal for *region_name* and export CSV.

    Returns the path of the saved CSV file.
    Raises an exception on any failure.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    download_dir = config.get_download_dir()
    csv_dest = config.get_export_csv_path()

    driver = _build_driver(download_dir)
    try:
        # ------------------------------------------------------------------
        # Step 1 — region selection
        # ------------------------------------------------------------------
        print(f"[{ts()}] Navigating to regions page …")
        driver.get("https://www.milkmoovement.com/regions")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Find the region link/button by normalised text
        region_norm = norm(region_name)
        found = False
        for tag in ("a", "button"):
            elements = driver.find_elements(By.TAG_NAME, tag)
            for el in elements:
                if norm(el.text) == region_norm:
                    print(f"[{ts()}] Clicking region: {el.text!r}")
                    el.click()
                    found = True
                    break
            if found:
                break

        if not found:
            dump_debug(driver, "region_not_found")
            raise RuntimeError(
                f"Region '{region_name}' not found on the regions page. "
                "Check that the region name matches exactly."
            )

        # Wait for redirect to the regional login page
        WebDriverWait(driver, 20).until(
            lambda d: "milkmoovement" in d.current_url and d.current_url != "https://www.milkmoovement.com/regions"
        )
        print(f"[{ts()}] Redirected to: {driver.current_url}")

        # ------------------------------------------------------------------
        # Step 2 — login
        # ------------------------------------------------------------------
        print(f"[{ts()}] Logging in …")
        user_field = find_first(driver, [
            (By.NAME, "username"),
            (By.NAME, "email"),
            (By.ID, "username"),
            (By.ID, "email"),
            (By.CSS_SELECTOR, "input[type='email']"),
            (By.CSS_SELECTOR, "input[type='text']"),
        ])
        if not user_field:
            dump_debug(driver, "login_no_user_field")
            raise RuntimeError("Could not find username/email field on login page.")
        user_field.clear()
        user_field.send_keys(username)

        pass_field = find_first(driver, [
            (By.NAME, "password"),
            (By.ID, "password"),
            (By.CSS_SELECTOR, "input[type='password']"),
        ])
        if not pass_field:
            dump_debug(driver, "login_no_pass_field")
            raise RuntimeError("Could not find password field on login page.")
        pass_field.clear()
        pass_field.send_keys(password)
        pass_field.send_keys(Keys.RETURN)

        # Wait for dashboard / main content
        WebDriverWait(driver, 30).until(
            lambda d: "login" not in d.current_url.lower() or
                      len(d.find_elements(By.CSS_SELECTOR, "nav, header, .dashboard")) > 0
        )
        print(f"[{ts()}] Logged in. Current URL: {driver.current_url}")
        time.sleep(2)

        # ------------------------------------------------------------------
        # Step 3 — navigate to Pickups & Labs
        # ------------------------------------------------------------------
        print(f"[{ts()}] Opening navigation menu …")
        # Try hamburger / nav menu first
        clicked_menu = click_first(driver, [
            (By.CSS_SELECTOR, "button.hamburger"),
            (By.CSS_SELECTOR, "[aria-label='menu']"),
            (By.CSS_SELECTOR, "button[class*='menu']"),
            (By.CSS_SELECTOR, "button[class*='hamburger']"),
            (By.CSS_SELECTOR, "nav button"),
            (By.XPATH, "//button[contains(@aria-label,'menu') or contains(@aria-label,'Menu')]"),
        ], "hamburger menu")

        if not clicked_menu:
            print(f"[{ts()}] No hamburger found, looking for Pickups & Labs link directly …")

        pickups_clicked = click_first(driver, [
            (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'pickups') and contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'labs')]"),
            (By.PARTIAL_LINK_TEXT, "Pickups"),
            (By.PARTIAL_LINK_TEXT, "pickups"),
        ], "Pickups & Labs")
        if not pickups_clicked:
            dump_debug(driver, "pickups_not_found")
            raise RuntimeError("Could not find 'Pickups & Labs' navigation item.")
        time.sleep(2)

        # ------------------------------------------------------------------
        # Step 4 — select "Last 2 weeks" preset
        # ------------------------------------------------------------------
        print(f"[{ts()}] Selecting 'Last 2 weeks' preset …")
        preset_clicked = click_first(driver, [
            (By.XPATH, "//*[contains(text(),'Last 2 weeks') or contains(text(),'last 2 weeks')]"),
            (By.CSS_SELECTOR, "[data-preset='last2weeks']"),
            (By.PARTIAL_LINK_TEXT, "Last 2 weeks"),
        ], "Last 2 weeks preset")
        if not preset_clicked:
            # Try opening a presets dropdown first
            click_first(driver, [
                (By.XPATH, "//*[contains(text(),'Preset') or contains(text(),'preset')]"),
                (By.CSS_SELECTOR, "[data-toggle='presets']"),
            ], "Presets dropdown")
            time.sleep(1)
            preset_clicked = click_first(driver, [
                (By.XPATH, "//*[contains(text(),'Last 2 weeks') or contains(text(),'last 2 weeks')]"),
                (By.PARTIAL_LINK_TEXT, "Last 2 weeks"),
            ], "Last 2 weeks preset (after dropdown)")
        if not preset_clicked:
            dump_debug(driver, "preset_not_found")
            raise RuntimeError("Could not select 'Last 2 weeks' preset.")
        time.sleep(2)

        # ------------------------------------------------------------------
        # Step 5 — export CSV
        # ------------------------------------------------------------------
        print(f"[{ts()}] Clicking Export button …")
        export_clicked = click_first(driver, [
            (By.XPATH, "//*[contains(text(),'Export') or contains(text(),'export')]"),
            (By.CSS_SELECTOR, "button[class*='export']"),
        ], "Export button")
        if not export_clicked:
            dump_debug(driver, "export_btn_not_found")
            raise RuntimeError("Could not find Export button.")
        time.sleep(1)

        print(f"[{ts()}] Clicking 'Export all columns as CSV' …")
        csv_export_clicked = click_first(driver, [
            (By.XPATH, "//*[contains(text(),'Export all columns as CSV') or contains(text(),'export all columns')]"),
            (By.PARTIAL_LINK_TEXT, "Export all columns"),
            (By.XPATH, "//*[contains(text(),'CSV')]"),
        ], "Export all columns as CSV")
        if not csv_export_clicked:
            dump_debug(driver, "csv_option_not_found")
            raise RuntimeError("Could not find 'Export all columns as CSV' option.")

        # ------------------------------------------------------------------
        # Step 6 — wait for download and move to destination
        # ------------------------------------------------------------------
        print(f"[{ts()}] Waiting for CSV download …")
        before = set(download_dir.glob("*.csv"))
        new_csv = wait_for_new_csv(download_dir, before, timeout=120)
        if not new_csv:
            dump_debug(driver, "download_timeout")
            raise RuntimeError("CSV download did not complete within 120 seconds.")

        atomic_replace(new_csv, csv_dest)
        print(f"[{ts()}] CSV saved to: {csv_dest}")
        return csv_dest

    except Exception:
        dump_debug(driver, "fatal")
        raise
    finally:
        try:
            driver.quit()
        except Exception:  # noqa: BLE001
            pass
