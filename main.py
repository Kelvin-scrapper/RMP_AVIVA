import os
import sys
import io
import glob
import time
import random
import logging
import traceback
import requests
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Fix Unicode encoding for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ============================================================================
# DIRECTORY SETUP
# ============================================================================
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'downloads')
LOGS_DIR     = os.path.join(BASE_DIR, 'logs')

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(LOGS_DIR,     exist_ok=True)

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, f'{timestamp}.log'), encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)

# ========== CONFIGURATION ==========
HEADLESS_MODE = True  # Set to False to show the browser window
# ====================================

# ============================================================================
# ASSERTION FUNCTIONS
# ============================================================================

def assert_with_log(condition, message):
    if not condition:
        logging.error(f"ASSERTION FAILED: {message}")
        raise AssertionError(message)

def assert_element_exists(element, element_name, context=""):
    context_msg = f" in {context}" if context else ""
    if element is None:
        msg = f"Element '{element_name}' not found{context_msg}"
        logging.error(f"ASSERTION FAILED: {msg}")
        raise AssertionError(msg)
    return element

def assert_file_exists(filepath, file_description=""):
    desc = file_description or filepath
    if not os.path.exists(filepath):
        msg = f"File not found: {desc} at {filepath}"
        logging.error(f"ASSERTION FAILED: {msg}")
        raise AssertionError(msg)
    logging.info(f"File verified: {desc}")
    return filepath

def assert_data_not_empty(data, data_name):
    if not data or len(data) == 0:
        msg = f"No data found for: {data_name}"
        logging.error(f"ASSERTION FAILED: {msg}")
        raise AssertionError(msg)
    logging.info(f"Data validated: {data_name} contains {len(data)} items")
    return data

def assert_driver_initialized(driver):
    assert_with_log(driver is not None, "WebDriver initialized")
    return driver

# ============================================================================
# ERROR HELPERS
# ============================================================================

def save_error_screenshot(driver, error_context=""):
    if driver:
        path = os.path.join(LOGS_DIR, f'error_{timestamp}_{error_context}.png')
        try:
            driver.save_screenshot(path)
            logging.error(f"Screenshot saved: {path}")
            return path
        except Exception as e:
            logging.error(f"Failed to save screenshot: {e}")
    return None

def save_page_source(driver, error_context=""):
    if driver:
        path = os.path.join(LOGS_DIR, f'error_{timestamp}_{error_context}.html')
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
            logging.error(f"Page source saved: {path}")
            return path
        except Exception as e:
            logging.error(f"Failed to save page source: {e}")
    return None

# ============================================================================
# DRIVER INITIALISATION
# ============================================================================

def initialize_driver():
    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-blink-features=AutomationControlled')

    if HEADLESS_MODE:
        options.add_argument('--headless=new')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-features=IsolateOrigins,site-per-process')
        options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
        logging.info("Running in HEADLESS mode (browser hidden)")
    else:
        logging.info("Running in VISIBLE mode (browser shown)")

    prefs = {
        "download.default_directory":                                   DOWNLOAD_DIR,
        "download.prompt_for_download":                                 False,
        "download.directory_upgrade":                                   True,
        "safebrowsing.enabled":                                         True,
        "plugins.always_open_pdf_externally":                           True,
        "profile.default_content_settings.popups":                      0,
        "profile.default_content_setting_values.automatic_downloads":   1,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    }
    options.add_experimental_option("prefs", prefs)

    logging.info(f"Downloads will be saved to: {DOWNLOAD_DIR}")
    logging.info("Initializing ChromeDriver with Selenium Manager...")

    driver = webdriver.Chrome(options=options)
    assert_driver_initialized(driver)

    if HEADLESS_MODE:
        driver.command_executor._commands["send_command"] = (
            "POST", '/session/$sessionId/chromium/send_command'
        )
        driver.execute("send_command", {
            'cmd': 'Page.setDownloadBehavior',
            'params': {'behavior': 'allow', 'downloadPath': DOWNLOAD_DIR}
        })
        logging.info("Enabled downloads in headless mode")

    logging.info("Chrome driver initialized successfully")
    return driver

# ============================================================================
# HELPERS
# ============================================================================

def human_like_scroll(driver, element=None, direction='down'):
    try:
        if element:
            driver.execute_script(
                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element
            )
            time.sleep(random.uniform(0.3, 0.7))
        else:
            scroll_amount = random.randint(200, 500)
            delta = scroll_amount if direction == 'down' else -scroll_amount
            driver.execute_script(f"window.scrollBy({{top: {delta}, behavior: 'smooth'}});")
            time.sleep(random.uniform(0.5, 1.2))
    except Exception as e:
        logging.warning(f"Scroll error: {str(e)[:100]}")


def wait_for_download(timeout=15, check_interval=0.5):
    """Wait for a new PDF to appear in DOWNLOAD_DIR and return its path."""
    initial_pdfs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf")))
    seconds = 0
    while seconds < timeout:
        time.sleep(check_interval)
        seconds += check_interval
        current_pdfs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf")))
        new_pdfs = current_pdfs - initial_pdfs
        if new_pdfs:
            latest = max(new_pdfs, key=os.path.getmtime)
            time.sleep(0.5)
            return latest
        if glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")):
            continue
    return None

# ============================================================================
# SCRAPING LOGIC
# ============================================================================

def scrape_aviva_data(driver):
    url = (
        "https://www.avivainvestors.com/en-gb/capabilities/fixed-income/"
        "emerging-markets-local-currency-bond-fund/lu0273498039-eur/"
    )
    logging.info(f"Navigating to {url}")
    driver.get(url)

    logging.info("Waiting for page to load...")
    time.sleep(2)
    human_like_scroll(driver, direction='down')

    # ── STEP 1: Accept cookies ────────────────────────────────────────────────
    logging.info("STEP 1: Accepting cookies")
    try:
        cookie_selectors = [
            (By.ID, "onetrust-accept-btn-handler"),
            (By.XPATH, "//button[contains(text(), 'Accept all cookies')]"),
            (By.XPATH, "//button[contains(text(), 'Accept All')]"),
        ]
        short_wait = WebDriverWait(driver, 5)
        accept_btn = None
        for by, sel in cookie_selectors:
            try:
                accept_btn = short_wait.until(EC.element_to_be_clickable((by, sel)))
                break
            except TimeoutException:
                continue

        if accept_btn:
            accept_btn.click()
            logging.info("Accepted all cookies")
            time.sleep(0.3)
        else:
            logging.warning("Accept Cookies button not found — may not be shown")
    except Exception as e:
        logging.warning(f"Could not accept cookies: {str(e)[:150]}")

    # ── STEP 2: Select Intermediary role ─────────────────────────────────────
    logging.info("STEP 2: Selecting Intermediary role")
    try:
        time.sleep(1)
        intermediary_selectors = [
            "//span[contains(text(), 'Intermediary')]/ancestor::label",
            "//label[contains(., 'Intermediary')]",
        ]
        intermediary_label = None
        for sel in intermediary_selectors:
            try:
                intermediary_label = driver.find_element(By.XPATH, sel)
                if intermediary_label:
                    break
            except Exception:
                continue

        if intermediary_label:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", intermediary_label
            )
            time.sleep(0.2)
            driver.execute_script("arguments[0].click();", intermediary_label)
            logging.info("Clicked Intermediary role")
            time.sleep(0.5)
        else:
            logging.warning("Intermediary role option not found")
    except Exception as e:
        logging.warning(f"Could not select Intermediary role: {str(e)[:150]}")

    # ── STEP 3: Remember me checkbox ─────────────────────────────────────────
    logging.info("STEP 3: Enabling Remember Me checkbox")
    try:
        time.sleep(0.5)
        checkbox = driver.find_element(By.ID, "disclaimer__remember-me")
        if checkbox and checkbox.is_enabled() and not checkbox.is_selected():
            driver.execute_script("arguments[0].click();", checkbox)
            logging.info("Checked 'Remember me' checkbox")
            time.sleep(0.3)
        else:
            logging.info("Checkbox already checked or disabled")
    except Exception as e:
        logging.warning(f"Could not handle checkbox: {str(e)[:150]}")

    # ── STEP 4: I Agree ───────────────────────────────────────────────────────
    logging.info("STEP 4: Clicking I Agree button")
    try:
        time.sleep(0.5)
        agree_btn = driver.find_element(
            By.CSS_SELECTOR, "a.disclaimer__btn.a-button.a-button--primary"
        )
        driver.execute_script("""
            arguments[0].classList.remove('is-disabled');
            arguments[0].setAttribute('aria-disabled', 'false');
        """, agree_btn)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", agree_btn)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", agree_btn)
        logging.info("Clicked I Agree button")
        time.sleep(3)
    except Exception as e:
        logging.warning(f"Could not click I Agree button: {str(e)[:150]}")

    logging.info(f"Current URL after navigation: {driver.current_url}")

    # Scroll to load lazy content
    for _ in range(4):
        driver.execute_script("window.scrollBy(0, 400);")
        time.sleep(0.3)

    # ── STEP 5: Scroll to Documents section & collect cards ───────────────────
    logging.info("STEP 5: Locating document cards")
    time.sleep(2)

    # Scroll all the way down so lazy-loaded document cards render
    logging.info("Scrolling page to load document cards...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(10):
        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(0.4)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
    time.sleep(2)

    # The page uses card-based layout: <li class="download-card ...">
    #   <a class="download-card__link" href="https://doc.morningstar.com/...msdoc?...">
    #     <h3 class="download-card__file-name">Fund factsheet</h3>
    #   </a>
    # </li>
    all_cards = driver.find_elements(By.CSS_SELECTOR, "li.download-card")
    logging.info(f"Found {len(all_cards)} download card(s) total")
    assert_data_not_empty(all_cards, "download cards")

    # Filter to only Fund factsheet cards
    target_cards = []
    for card in all_cards:
        try:
            name_el = card.find_element(By.CSS_SELECTOR, "h3.download-card__file-name")
            if "Fund factsheet" in name_el.text:
                link_el  = card.find_element(By.CSS_SELECTOR, "a.download-card__link")
                doc_url  = link_el.get_attribute("href")
                try:
                    date_el  = card.find_element(By.CSS_SELECTOR, "span.file-info__count")
                    doc_date = date_el.text.strip()
                except Exception:
                    doc_date = "N/A"
                target_cards.append((name_el.text.strip(), doc_date, doc_url))
        except Exception:
            continue

    logging.info(f"Found {len(target_cards)} 'Fund factsheet' card(s)")
    assert_data_not_empty(target_cards, "Fund factsheet cards")

    # Build a requests session using Selenium's cookies so Morningstar
    # recognises us as a valid Aviva Investors referrer.
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Referer": driver.current_url,
    })
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"])

    downloaded_count = 0

    for idx, (doc_name, doc_date, doc_url) in enumerate(target_cards, 1):
        logging.info(f"Document {idx}: '{doc_name}' | Date: {doc_date} | URL: {doc_url}")

        try:
            # Step A: Fetch the .msdoc viewer page
            resp = session.get(doc_url, timeout=30, allow_redirects=True)
            logging.info(f"msdoc response: {resp.status_code} | Content-Type: {resp.headers.get('Content-Type','?')}")

            content_type = resp.headers.get("Content-Type", "")

            if "application/pdf" in content_type or resp.url.endswith(".pdf"):
                # Morningstar redirected straight to the PDF
                pdf_bytes = resp.content
                pdf_url_final = resp.url
            else:
                # HTML viewer — parse for the embedded PDF URL
                import re
                pdf_url_final = None

                # Pattern 1: look for a direct .pdf href or src
                match = re.search(
                    r'(?:src|href|data)=["\']([^"\']+\.pdf[^"\']*)["\']',
                    resp.text, re.IGNORECASE
                )
                if match:
                    pdf_url_final = match.group(1)
                    if pdf_url_final.startswith("//"):
                        pdf_url_final = "https:" + pdf_url_final

                # Pattern 2: look for a "url" JSON key pointing to a PDF
                if not pdf_url_final:
                    match = re.search(
                        r'"url"\s*:\s*"([^"]+\.pdf[^"]*)"',
                        resp.text, re.IGNORECASE
                    )
                    if match:
                        pdf_url_final = match.group(1).replace("\\u0026", "&")

                # Pattern 3: Morningstar sometimes uses /document/.../download endpoint
                if not pdf_url_final:
                    hash_match = re.search(
                        r'/document/([a-f0-9]+)\.msdoc', doc_url
                    )
                    if hash_match:
                        doc_hash = hash_match.group(1)
                        pdf_url_final = (
                            f"https://doc.morningstar.com/document/{doc_hash}.pdf"
                            f"?clientId=avivainvestors&key=0011b526a18a80ef"
                        )
                        logging.info(f"Constructed direct PDF URL: {pdf_url_final}")

                if not pdf_url_final:
                    logging.error(f"Could not find PDF URL in msdoc page for: {doc_name}")
                    continue

                logging.info(f"Fetching PDF from: {pdf_url_final}")
                pdf_resp = session.get(pdf_url_final, timeout=60, allow_redirects=True)
                logging.info(
                    f"PDF response: {pdf_resp.status_code} | "
                    f"Content-Type: {pdf_resp.headers.get('Content-Type','?')} | "
                    f"Size: {len(pdf_resp.content)//1024} KB"
                )
                if pdf_resp.status_code != 200:
                    logging.error(f"Failed to download PDF: HTTP {pdf_resp.status_code}")
                    continue
                pdf_bytes = pdf_resp.content

            # Save the PDF
            safe_name = doc_name.replace(" ", "_").replace("/", "-")
            safe_date = doc_date.replace("/", "-")
            filename   = f"{safe_name}_{safe_date}.pdf"
            save_path  = os.path.join(DOWNLOAD_DIR, filename)

            with open(save_path, "wb") as f:
                f.write(pdf_bytes)

            size_kb = len(pdf_bytes) / 1024
            logging.info(f"Saved: {filename} ({size_kb:.2f} KB)")
            assert_file_exists(save_path, filename)
            downloaded_count += 1

        except Exception as e:
            logging.error(f"Error downloading '{doc_name}': {str(e)[:200]}")

    logging.info(f"Successfully downloaded {downloaded_count} / {len(target_cards)} document(s)")

    # Final summary
    pdf_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf"))
    if pdf_files:
        logging.info(f"Files in {DOWNLOAD_DIR}:")
        for pdf in pdf_files:
            size_kb = os.path.getsize(pdf) / 1024
            logging.info(f"  - {os.path.basename(pdf)} ({size_kb:.2f} KB)")
    else:
        logging.warning(f"No PDF files found in {DOWNLOAD_DIR}")

# ============================================================================
# MAIN
# ============================================================================

def main():
    driver = None
    try:
        logging.info("=" * 60)
        logging.info(f"AVIVA DATA EXTRACTION PIPELINE — {timestamp}")
        logging.info(f"Python {sys.version}")
        logging.info("=" * 60)

        driver = initialize_driver()
        scrape_aviva_data(driver)

        logging.info("=" * 60)
        logging.info("SCRIPT COMPLETED SUCCESSFULLY")
        logging.info("=" * 60)
        return 0

    except AssertionError as e:
        logging.error("=" * 60)
        logging.error("ASSERTION FAILED")
        logging.error(f"Error: {e}")
        logging.error("=" * 60)
        save_error_screenshot(driver, "assertion_failure")
        save_page_source(driver, "assertion_failure")
        return 1

    except Exception as e:
        logging.error("=" * 60)
        logging.error("UNEXPECTED ERROR")
        logging.error(f"Type: {type(e).__name__}")
        logging.error(f"Message: {e}")
        logging.error("=" * 60)
        logging.error(traceback.format_exc())
        save_error_screenshot(driver, "unexpected_error")
        save_page_source(driver, "unexpected_error")
        return 1

    finally:
        if driver:
            try:
                driver.quit()
                logging.info("Browser closed successfully")
            except Exception as e:
                logging.warning(f"Error closing browser: {e}")


if __name__ == "__main__":
    sys.exit(main())
