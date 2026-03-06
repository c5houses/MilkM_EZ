"""Selenium automation: log in to the Milk Moovement portal and export CSV."""

import os
import shutil
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import config

from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    WebDriverException,
)

# ---------------------------------------------------------------------------
# Region → login URL mapping
# ---------------------------------------------------------------------------

REGION_URLS = {
    "Bongards Creameries": "https://bongards.milkmoovement.io/#/login",
    "California Dairies Inc.": "https://cdi.milkmoovement.io/#/login",
    "Cargill Feed Shop": "https://cargill.milkmoovement.io/#/login",
    "Dairy Farmers of Newfoundland and Labrador": "https://dfnl.milkmoovement.io/#/login",
    "Erie Cooperative Association (TCJ)": "https://erie.milkmoovement.io/#/login",
    "Great Plains Dairymen's Association (TCJ)": "https://greatplains.milkmoovement.io/#/login",
    "Idaho Milk Products": "https://idahomilk.milkmoovement.io/#/login",
    "KYTN Cooperative (TCJ)": "https://kytn.milkmoovement.io/#/login",
    "Legacy Milk": "https://legacy.milkmoovement.io/#/login",
    "Liberty Milk Producers (TCJ)": "https://liberty.milkmoovement.io/#/login",
    "Michigan Milk Producers Association": "https://mmpa.milkmoovement.io/#/login",
    "Minerva Dairy": "https://minerva.milkmoovement.io/#/login",
    "Nebraska Milk Producers (TCJ)": "https://nebraska.milkmoovement.io/#/login",
    "Plainview Milk Products": "https://plainview.milkmoovement.io/#/login",
    "Prairie Farms": "https://prairiefarms.milkmoovement.io/#/login",
    "United Dairymen of Arizona": "https://uda.milkmoovement.io/#/login",
    "Upstate Niagara Cooperative Inc.": "https://upstateniagara.milkmoovement.io/#/login",
    "White Eagle Cooperative (TCJ)": "https://whiteeagle.milkmoovement.io/#/login",
}

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
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        shot = out_dir / f"debug_{label}_{stamp}.png"
        driver.save_screenshot(str(shot))
        print(f"[{ts()}] Debug screenshot saved: {shot}")
        html_path = out_dir / f"debug_{label}_{stamp}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"[{ts()}] Debug HTML saved: {html_path}")
    except Exception as exc:  # noqa: BLE001
        print(f"[{ts()}] dump_debug failed: {exc}")


def find_first(driver, locators, clickable=False, timeout=20, label="element"):
    """Try multiple locators and return the first element found."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    end = time.time() + timeout
    last_exc = None
    while time.time() < end:
        for by, sel in locators:
            try:
                if clickable:
                    el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((by, sel)))
                else:
                    el = WebDriverWait(driver, 2).until(EC.presence_of_element_located((by, sel)))
                return el
            except Exception as e:
                last_exc = e
        time.sleep(0.15)
    raise TimeoutException(f"Could not find {label}. Last error: {last_exc}")


def click_first(driver, locators, timeout=20, label="element", retries=3):
    """Click the first matching locator using robust click patterns."""
    last_exc = None
    for _ in range(retries):
        try:
            el = find_first(driver, locators, clickable=True, timeout=timeout, label=label)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.1)
            el.click()
            print(f"[{ts()}] Clicked: {label}")
            return True
        except (
            StaleElementReferenceException,
            ElementClickInterceptedException,
            ElementNotInteractableException,
            TimeoutException,
            WebDriverException,
        ) as e:
            last_exc = e
            # JS click fallback
            try:
                el = find_first(driver, locators, clickable=False, timeout=3, label=label)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                print(f"[{ts()}] Clicked (JS fallback): {label}")
                return True
            except Exception:
                time.sleep(0.3)
    print(f"WARNING: click_first failed for {label}. Last error: {last_exc}")
    return False


def open_and_click_menu_item_by_text(driver, desired_text, timeout=15, click_submit_changes=False):
    """Click a menu item by visible text (case-insensitive).
    If click_submit_changes=True, will also click 'SUBMIT CHANGES' if present afterward.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, timeout)
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH,
             "//*[@role='menu']//*[@role='menuitem'] | //*[@role='menuitem'] | //ul//li | //div[@role='presentation'] | //button")
        )
    )
    time.sleep(0.2)

    desired_raw = norm(desired_text)
    desired = desired_raw.lower()

    candidates = driver.find_elements(
        By.XPATH,
        "//*[@role='menuitem'] | //ul//li | //div[@role='option'] | //button | //div[@role='button']",
    )

    visible = []
    for el in candidates:
        try:
            if el.is_displayed():
                txt = " ".join((el.text or "").split()).strip()
                if txt:
                    visible.append((el, txt, txt.lower()))
        except Exception:
            pass

    # exact match (case-insensitive)
    clicked = False
    for el, txt, low in visible:
        if low == desired:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                print(f"[{ts()}] Clicked menu option: {txt}")
                clicked = True
                break
            except Exception:
                pass

    # contains match
    if not clicked:
        for el, txt, low in visible:
            if desired in low:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    driver.execute_script("arguments[0].click();", el)
                    print(f"[{ts()}] Clicked menu option (contains): {txt}")
                    clicked = True
                    break
                except Exception:
                    pass

    if not clicked:
        print(f"WARNING: Could not click menu item: '{desired_raw}'")
        return False

    # Some versions require "SUBMIT CHANGES"
    if click_submit_changes:
        time.sleep(0.4)
        for el, txt, low in visible:
            if low == "submit changes":
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    driver.execute_script("arguments[0].click();", el)
                    print(f"[{ts()}] Clicked: SUBMIT CHANGES")
                    return True
                except Exception:
                    pass
        # Re-scan quickly (menu may rerender)
        try:
            submit = driver.find_element(
                By.XPATH,
                "//*[self::button or @role='menuitem' or @role='button']"
                "[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'SUBMIT CHANGES')]",
            )
            if submit.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit)
                driver.execute_script("arguments[0].click();", submit)
                print(f"[{ts()}] Clicked: SUBMIT CHANGES")
        except Exception:
            pass

    return True


def wait_for_new_csv(download_dir: Path, before: set, timeout: int = 120) -> Optional[Path]:
    """Wait until a new .csv file appears in *download_dir* and is fully written."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        temps = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.tmp")) + list(download_dir.glob("*.download"))
        if temps:
            time.sleep(1)
            continue
        current = {p.name for p in download_dir.glob("*.csv")}
        new_names = current - before
        if new_names:
            candidates = [download_dir / n for n in new_names]
            return max(candidates, key=lambda p: p.stat().st_mtime)
        time.sleep(1)
    return None


def atomic_replace(src: Path, dst: Path) -> None:
    """Safer replace: move to temp name then os.replace (atomic)."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dst.with_suffix(dst.suffix + ".tmp")
    if tmp_dest.exists():
        tmp_dest.unlink()
    shutil.move(str(src), str(tmp_dest))
    os.replace(str(tmp_dest), str(dst))


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
        # Step 1 — navigate directly to regional login page
        # ------------------------------------------------------------------
        login_url = REGION_URLS.get(region_name)
        if not login_url:
            raise RuntimeError(
                f"Unknown region '{region_name}'. "
                f"Known regions: {', '.join(sorted(REGION_URLS.keys()))}"
            )
        print(f"[{ts()}] Navigating to {login_url} …")
        driver.get(login_url)

        # Wait for the SPA to fully render the login form
        time.sleep(5)
        print(f"[{ts()}] Page loaded. Current URL: {driver.current_url}")

        # ------------------------------------------------------------------
        # Step 2 — login
        # ------------------------------------------------------------------
        print(f"[{ts()}] Logging in …")

        # Check if already logged in
        try:
            WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.TAG_NAME, "header")))
            print(f"[{ts()}] Already logged in (header found).")
            already_logged_in = True
        except TimeoutException:
            already_logged_in = False

        if not already_logged_in:
            user_locators = [
                (By.CSS_SELECTOR, "input[type='text']"),
                (By.CSS_SELECTOR, "input[type='email']"),
                (By.CSS_SELECTOR, "input.MuiInputBase-input"),
                (By.CSS_SELECTOR, ".MuiOutlinedInput-root input"),
                (By.XPATH, "//label[contains(text(),'Username') or contains(text(),'Email')]/following::input[1]"),
                (By.XPATH, "//div[contains(@class,'MuiInputBase-root')]//input"),
                (By.NAME, "username"),
                (By.NAME, "email"),
                (By.ID, "username"),
                (By.ID, "email"),
                (By.XPATH, "//input[contains(@aria-label,'Email') or contains(@placeholder,'Email')]"),
            ]
            pass_locators = [
                (By.CSS_SELECTOR, "input[type='password']"),
                (By.CSS_SELECTOR, "input.MuiInputBase-input[type='password']"),
                (By.XPATH, "//label[contains(text(),'Password')]/following::input[1]"),
                (By.XPATH, "//div[contains(@class,'MuiInputBase-root')]//input[@type='password']"),
                (By.NAME, "password"),
                (By.ID, "password"),
                (By.XPATH, "//input[contains(@aria-label,'Password') or contains(@placeholder,'Password')]"),
            ]

            user_field = find_first(driver, user_locators, clickable=False, timeout=20, label="username/email field")
            user_field.clear()
            user_field.send_keys(username)
            time.sleep(0.4)

            pass_field = find_first(driver, pass_locators, clickable=False, timeout=20, label="password field")
            pass_field.clear()
            pass_field.send_keys(password)
            time.sleep(0.4)

            pass_field.send_keys(Keys.RETURN)
            time.sleep(1)

            # Also try clicking Log In button as backup
            try:
                click_first(driver, [
                    (By.XPATH, "//button[contains(text(),'Log In') or contains(text(),'Login') or contains(text(),'Sign In')]"),
                    (By.CSS_SELECTOR, "button[type='submit']"),
                    (By.CSS_SELECTOR, "button.MuiButton-root"),
                ], timeout=5, label="Login button")
            except Exception:
                pass

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "header")))
            print(f"[{ts()}] Login successful. Current URL: {driver.current_url}")
            time.sleep(1.2)

        # ------------------------------------------------------------------
        # Step 3 — navigate to Pickups & Labs
        # ------------------------------------------------------------------
        print(f"[{ts()}] Opening navigation menu …")
        nav_button_locators = [
            (By.CSS_SELECTOR, "button[aria-label*='menu' i]"),
            (By.CSS_SELECTOR, "button[aria-label*='navigation' i]"),
            (By.XPATH, "//header//button[contains(@aria-label,'Menu') or contains(@aria-label,'menu')]"),
            (By.XPATH, "//button[@type='button' and (contains(@aria-label,'Menu') or contains(@aria-label,'menu'))]"),
            (By.XPATH, "(//header//button)[1]"),
        ]
        if not click_first(driver, nav_button_locators, timeout=20, label="nav/hamburger button"):
            dump_debug(driver, "nav_button_not_found")
            raise RuntimeError("Could not find navigation/hamburger button.")
        time.sleep(0.6)

        print(f"[{ts()}] Clicking 'Pickups & Labs' …")
        pickups_locators = [
            (By.XPATH,
             "//*[self::a or self::button or @role='button' or @role='menuitem']"
             "[.//div[normalize-space()='Pickups & Labs'] or .//*[normalize-space()='Pickups & Labs']]"),
            (By.XPATH,
             "//*[self::a or self::button or @role='button' or @role='menuitem']"
             "[contains(normalize-space(.), 'Pickups') and contains(normalize-space(.), 'Labs')]"),
            (By.XPATH, "//*[normalize-space()='Pickups & Labs']"),
        ]
        if not click_first(driver, pickups_locators, timeout=25, label="Pickups & Labs nav item"):
            dump_debug(driver, "pickups_not_found")
            raise RuntimeError("Could not find 'Pickups & Labs' navigation item.")
        print(f"[{ts()}] Pickups & Labs page opened.")
        time.sleep(2)

        # ------------------------------------------------------------------
        # Step 4 — select "Last 2 weeks" preset
        # ------------------------------------------------------------------
        print(f"[{ts()}] Selecting preset: Last 2 weeks …")
        presets_button_locators = [
            (By.XPATH, "//button[contains(., 'Presets')]"),
            (By.XPATH, "//button[contains(., 'PRESETS')]"),
            (By.XPATH, "//button[contains(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'PRESETS')]"),
            (By.CSS_SELECTOR, "button[aria-label*='preset' i]"),
        ]
        if not click_first(driver, presets_button_locators, timeout=20, label="Presets button/dropdown"):
            dump_debug(driver, "presets_button_not_found")
            raise RuntimeError("Could not find Presets button.")
        time.sleep(0.4)

        if not open_and_click_menu_item_by_text(driver, "Last 2 weeks", timeout=15, click_submit_changes=True):
            dump_debug(driver, "preset_menu_item_not_found")
            raise RuntimeError("Could not select 'Last 2 weeks' from presets menu.")
        print(f"[{ts()}] Last 2 weeks selected.")
        time.sleep(2.0)

        # ------------------------------------------------------------------
        # Step 5 — export CSV
        # ------------------------------------------------------------------
        print(f"[{ts()}] Scrolling page to find Export button …")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)
        print(f"[{ts()}] Opening Export menu …")
        export_button_locators = [
            (By.XPATH, "//button[contains(., 'Export')]"),
            (By.XPATH, "//button[contains(., 'EXPORT')]"),
            (By.CSS_SELECTOR, "button[aria-label*='export' i]"),
            (By.XPATH, "//*[self::button or self::span][contains(., 'Export') or contains(., 'EXPORT')]/ancestor::button[1]"),
            (By.CSS_SELECTOR, "button[title*='Export' i]"),
        ]
        if not click_first(driver, export_button_locators, timeout=25, label="Export button"):
            dump_debug(driver, "export_btn_not_found")
            raise RuntimeError("Could not find Export button.")
        time.sleep(0.5)

        print(f"[{ts()}] Clicking 'Export all columns as CSV' …")
        before = {p.name for p in download_dir.glob("*.csv")}
        if not open_and_click_menu_item_by_text(driver, "Export all columns as CSV", timeout=15, click_submit_changes=False):
            dump_debug(driver, "csv_option_not_found")
            raise RuntimeError("Could not find 'Export all columns as CSV' option.")

        # ------------------------------------------------------------------
        # Step 6 — wait for download and move to destination
        # ------------------------------------------------------------------
        print(f"[{ts()}] Waiting for CSV download …")
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
