import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import sys
import io
import random
import os
from datetime import datetime

# Fix Unicode encoding for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ========== CONFIGURATION ==========
HEADLESS_MODE = True  # Set to True to run browser in headless mode (hidden)
# ==================================

def initialize_driver_with_retry(max_attempts=3):
    """Initialize Chrome driver with automatic version handling and retry logic"""

    def create_options():
        """Create fresh ChromeOptions (cannot be reused)"""
        import os

        # Get the directory where the script is located
        download_dir = os.path.dirname(os.path.abspath(__file__))

        options = uc.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-blink-features=AutomationControlled')

        # Enable headless mode if configured
        if HEADLESS_MODE:
            options.add_argument('--headless=new')
            # Additional arguments to make headless mode work better
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--disable-web-security')
            options.add_argument('--disable-features=IsolateOrigins,site-per-process')
            options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            print("Running in HEADLESS mode (browser hidden)")
        else:
            print("Running in VISIBLE mode (browser shown)")

        # Set download preferences to save in script directory and auto-download PDFs
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "plugins.always_open_pdf_externally": True,  # Automatically download PDFs instead of opening in browser
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1  # Allow automatic downloads
        }
        options.add_experimental_option("prefs", prefs)

        print(f"Downloads will be saved to: {download_dir}")

        return options

    # First, try to detect the current Chrome version and use it
    import re
    try:
        print("Attempting to initialize with auto-detected Chrome version...")
        driver = uc.Chrome(options=create_options(), version_main=None, use_subprocess=True)
        print("✓ Chrome driver initialized successfully!")
        return driver
    except SessionNotCreatedException as e:
        error_msg = str(e)
        print(f"✗ Version mismatch detected: {error_msg[:150]}...")

        # Extract version numbers from error
        if "only supports Chrome version" in error_msg or "Current browser version is" in error_msg:
            versions = re.findall(r'version (?:is )?(\d+)', error_msg)
            if len(versions) >= 1:
                # The last version in the error is usually the current browser version
                current_version = int(versions[-1])
                print(f"  Detected browser version: {current_version}")
                print(f"  Attempting to download matching ChromeDriver version {current_version}...")

                try:
                    driver = uc.Chrome(
                        options=create_options(),
                        version_main=current_version,
                        use_subprocess=True
                    )
                    print("✓ Chrome driver initialized with matching version!")
                    return driver
                except Exception as e2:
                    print(f"✗ Failed with version {current_version}: {str(e2)[:150]}")
    except Exception as e:
        print(f"✗ Initial attempt failed: {type(e).__name__}: {str(e)[:150]}")

    # Try alternative strategies
    strategies = [
        {'version_main': None, 'use_subprocess': False},
        {'use_subprocess': True},
    ]

    for attempt in range(max_attempts):
        for strategy_idx, strategy in enumerate(strategies):
            try:
                print(f"\nAttempt {attempt + 1}/{max_attempts}, Strategy {strategy_idx + 1}/{len(strategies)}")
                print(f"  Using strategy: {strategy}")

                driver = uc.Chrome(options=create_options(), **strategy)
                print("✓ Chrome driver initialized successfully!")
                return driver

            except Exception as e:
                print(f"✗ Failed: {type(e).__name__}: {str(e)[:150]}")
                if attempt == max_attempts - 1 and strategy_idx == len(strategies) - 1:
                    raise
                time.sleep(1)

    raise Exception("Failed to initialize Chrome driver after all attempts")

def human_like_scroll(driver, element=None, direction='down'):
    """Scroll in a human-like manner with random speeds and pauses"""
    try:
        if element:
            # Scroll to element with smooth behavior
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
            time.sleep(random.uniform(0.3, 0.7))
        else:
            # Random scroll amount
            scroll_amount = random.randint(200, 500)
            if direction == 'down':
                driver.execute_script(f"window.scrollBy({{top: {scroll_amount}, behavior: 'smooth'}});")
            else:
                driver.execute_script(f"window.scrollBy({{top: -{scroll_amount}, behavior: 'smooth'}});")
            time.sleep(random.uniform(0.5, 1.2))
    except Exception as e:
        print(f"  ⚠ Scroll error: {str(e)[:100]}")

def random_pause():
    """Add random human-like pauses"""
    time.sleep(random.uniform(0.5, 1.5))


def wait_for_download(download_dir, timeout=15, check_interval=0.5):
    """Wait for download to complete and return the downloaded file path"""
    import glob

    # Get initial PDF files to ignore
    initial_pdfs = set(glob.glob(os.path.join(download_dir, "*.pdf")))

    seconds = 0
    while seconds < timeout:
        time.sleep(check_interval)
        seconds += check_interval

        # Look for new PDF files (not .crdownload)
        current_pdfs = set(glob.glob(os.path.join(download_dir, "*.pdf")))
        new_pdfs = current_pdfs - initial_pdfs

        if new_pdfs:
            # Return the newly downloaded PDF
            latest_file = max(new_pdfs, key=os.path.getmtime)

            # Wait a bit more to ensure file is fully written
            time.sleep(0.5)
            return latest_file

        # Check if any download is in progress
        crdownload_files = glob.glob(os.path.join(download_dir, "*.crdownload"))
        if crdownload_files:
            # Download is in progress, keep waiting
            continue

    return None

def get_filename_from_url(driver):
    """Try to get the original filename from the download or page"""
    try:
        # Method 1: Check if there's a download attribute or Content-Disposition header
        script = """
        // Try to find download links with filename
        let links = document.querySelectorAll('a[download]');
        if (links.length > 0) {
            return links[0].getAttribute('download');
        }

        // Try to get filename from page title or document title
        let title = document.title;
        if (title) {
            // Clean up title to make it a valid filename
            return title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        }

        return null;
        """
        filename = driver.execute_script(script)
        return filename
    except:
        return None

def scrape_aviva_data():
    driver = None

    try:
        # Initialize driver with automatic version handling
        driver = initialize_driver_with_retry()

        # Enable downloads in headless mode
        if HEADLESS_MODE:
            download_dir = os.path.dirname(os.path.abspath(__file__))
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {
                'cmd': 'Page.setDownloadBehavior',
                'params': {
                    'behavior': 'allow',
                    'downloadPath': download_dir
                }
            }
            driver.execute("send_command", params)
            print(f"✓ Enabled downloads in headless mode")

        # Navigate to the URL
        url = "https://www.avivainvestors.com/en-gb/capabilities/fixed-income/emerging-markets-local-currency-bond-fund/lu0273498039-eur/"
        print(f"\nNavigating to {url}...")
        driver.get(url)

        # Wait for page to load
        print("Waiting for page to load...")
        time.sleep(2)

        # Initial human-like scroll
        print("Performing initial page scroll...")
        human_like_scroll(driver, direction='down')

        # Wait for elements
        wait = WebDriverWait(driver, 20)

        # STEP 1: Accept All Cookies
        print("\n=== STEP 1: Accepting All Cookies ===")
        try:
            # Try multiple selectors for the Accept All Cookies button
            cookie_selectors = [
                (By.ID, "onetrust-accept-btn-handler"),
                (By.XPATH, "//button[contains(text(), 'Accept all cookies')]"),
                (By.XPATH, "//button[contains(text(), 'Accept All')]"),
            ]

            accept_cookies_button = None
            short_wait = WebDriverWait(driver, 5)
            for by, selector in cookie_selectors:
                try:
                    accept_cookies_button = short_wait.until(
                        EC.element_to_be_clickable((by, selector))
                    )
                    print(f"✓ Found Accept Cookies button")
                    break
                except TimeoutException:
                    continue

            if accept_cookies_button:
                accept_cookies_button.click()
                print("✓ Accepted all cookies")
                time.sleep(0.3)
            else:
                print("✗ Could not find Accept Cookies button")

        except Exception as e:
            print(f"✗ Could not accept cookies: {str(e)[:150]}")

        # STEP 2: Select Intermediary Role
        print("\n=== STEP 2: Selecting Intermediary Role ===")
        try:
            # Wait for role selection to appear
            time.sleep(1)

            # Try different selectors for the Intermediary option
            intermediary_selectors = [
                "//span[contains(text(), 'Intermediary')]/ancestor::label",
                "//label[contains(., 'Intermediary')]",
            ]

            intermediary_label = None
            for selector in intermediary_selectors:
                try:
                    intermediary_label = driver.find_element(By.XPATH, selector)
                    if intermediary_label:
                        print(f"✓ Found Intermediary label")
                        break
                except:
                    continue

            if intermediary_label:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", intermediary_label)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", intermediary_label)
                print("✓ Clicked Intermediary role")
                time.sleep(0.5)
            else:
                print("✗ Could not find Intermediary role option")

        except Exception as e:
            print(f"✗ Could not select Intermediary role: {str(e)[:150]}")

        # STEP 3: Enable "Remember me" checkbox
        print("\n=== STEP 3: Enabling Remember Me Checkbox ===")
        try:
            time.sleep(0.5)
            checkbox = driver.find_element(By.ID, "disclaimer__remember-me")

            if checkbox and checkbox.is_enabled() and not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", checkbox)
                print("✓ Checked 'Remember me' checkbox")
                time.sleep(0.3)
            else:
                print("✓ Checkbox already checked or disabled")

        except Exception as e:
            print(f"✗ Could not handle checkbox: {str(e)[:150]}")

        # STEP 4: Click "I agree" button
        print("\n=== STEP 4: Clicking I Agree Button ===")
        try:
            time.sleep(0.5)
            agree_button = driver.find_element(By.CSS_SELECTOR, "a.disclaimer__btn.a-button.a-button--primary")

            if agree_button:
                # Remove disabled state if present
                driver.execute_script("""
                    arguments[0].classList.remove('is-disabled');
                    arguments[0].setAttribute('aria-disabled', 'false');
                """, agree_button)

                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", agree_button)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", agree_button)
                print("✓ Clicked I agree button")
                time.sleep(3)  # Wait for page content to load after disclaimer
            else:
                print("✗ Could not find I agree button")

        except Exception as e:
            print(f"✗ Could not click I agree button: {str(e)[:150]}")

        print("\n=== Navigation Complete ===")
        print(f"Current URL: {driver.current_url}")

        # Scroll down to load content
        print("Scrolling to load documents...")
        for _ in range(4):
            driver.execute_script("window.scrollBy(0, 400);")
            time.sleep(0.3)

        # STEP 5: Find and download documents from table
        print("\n=== STEP 5: Downloading Documents ===")
        try:
            # Wait for the document table to load
            time.sleep(4)

            # Find all table rows with document data
            # Looking for rows with the specific structure: td with header matching documentName
            document_rows = driver.find_elements(
                By.XPATH,
                "//td[@headers and contains(@headers, 'documentName')]/ancestor::tr[contains(@class, 'ec-table__row')]"
            )

            if not document_rows:
                # Try alternative selector
                document_rows = driver.find_elements(
                    By.XPATH,
                    "//td[contains(@class, 'ec-table__cell') and @data-title='Document name']/ancestor::tr"
                )

            print(f"✓ Found {len(document_rows)} document row(s)")

            downloaded_count = 0
            download_dir = os.path.dirname(os.path.abspath(__file__))

            for idx, row in enumerate(document_rows, 1):
                try:
                    # Get document name from the first cell
                    doc_name_cell = row.find_element(By.XPATH, ".//div[contains(@class, 'ec-table__cell-content')]")
                    doc_name = doc_name_cell.text.strip()

                    # Only download from "Fund factsheet" section
                    if "Fund factsheet" not in doc_name:
                        print(f"\n  Skipping {idx}: {doc_name} (not Fund factsheet)")
                        continue

                    # Get date if available
                    try:
                        date_cell = row.find_element(By.XPATH, ".//td[@data-title='Date']//div[@class='ec-table__cell-content']")
                        doc_date = date_cell.text.strip()
                    except:
                        doc_date = "N/A"

                    print(f"\n  Document {idx}: {doc_name} (Date: {doc_date})")

                    # Find the download button in this row
                    download_button = row.find_element(By.XPATH, ".//button[@type='button']")

                    # Human-like scroll to the button
                    human_like_scroll(driver, download_button)
                    time.sleep(0.3)

                    # Store current window handles before clicking
                    original_window = driver.current_window_handle
                    original_windows = driver.window_handles

                    # Get initial PDF count before clicking
                    import glob
                    initial_pdfs = set(glob.glob(os.path.join(download_dir, "*.pdf")))

                    # Click the download button
                    try:
                        download_button.click()
                        print(f"  ✓ Clicked download button for: {doc_name}")
                    except Exception as e:
                        print(f"  Regular click failed, trying JavaScript click")
                        driver.execute_script("arguments[0].click();", download_button)
                        print(f"  ✓ Clicked download button for: {doc_name} (via JavaScript)")

                    # Wait for new tab to potentially open or download to start
                    time.sleep(0.5)

                    # Check if a new tab opened
                    new_windows = driver.window_handles
                    if len(new_windows) > len(original_windows):
                        print(f"  ✓ New tab detected, switching to it...")
                        # Switch to the new tab
                        new_window = [w for w in new_windows if w not in original_windows][0]
                        driver.switch_to.window(new_window)

                        # Wait for the PDF to load in the new tab
                        print(f"  ✓ Waiting for document to load in new tab...")
                        time.sleep(1)

                        current_url = driver.current_url
                        print(f"  📄 New tab URL: {current_url}")

                        # Human-like scroll in the new tab
                        human_like_scroll(driver, direction='down')
                        time.sleep(0.3)

                        # Check if it's a direct PDF or opened in Chrome's PDF viewer
                        if '.msdoc' in current_url or 'morningstar.com/document' in current_url:
                            print(f"  ✓ Morningstar document viewer detected")
                            # Wait for the viewer to load fully
                            time.sleep(1)

                            # Try to extract the direct PDF URL from the page
                            try:
                                pdf_url_script = """
                                // Try to find the embedded PDF URL
                                let pdfUrl = null;

                                // Method 1: Check for embed element with PDF src
                                let embeds = document.querySelectorAll('embed[type="application/pdf"]');
                                if (embeds.length > 0) {
                                    pdfUrl = embeds[0].src;
                                }

                                // Method 2: Check for iframe with PDF
                                if (!pdfUrl) {
                                    let iframes = document.querySelectorAll('iframe');
                                    for (let iframe of iframes) {
                                        if (iframe.src && (iframe.src.includes('.pdf') || iframe.src.includes('pdf'))) {
                                            pdfUrl = iframe.src;
                                            break;
                                        }
                                    }
                                }

                                // Method 3: Look for object tags
                                if (!pdfUrl) {
                                    let objects = document.querySelectorAll('object[type="application/pdf"]');
                                    if (objects.length > 0) {
                                        pdfUrl = objects[0].data;
                                    }
                                }

                                // Method 4: Check if current page itself is a PDF
                                if (!pdfUrl && document.contentType === 'application/pdf') {
                                    pdfUrl = window.location.href;
                                }

                                return pdfUrl;
                                """

                                pdf_url = driver.execute_script(pdf_url_script)
                                print(f"  📋 Extracted PDF URL: {pdf_url}")

                                if pdf_url and pdf_url != current_url:
                                    print(f"  ✓ Found direct PDF URL, navigating to it for download...")
                                    driver.get(pdf_url)
                                    time.sleep(2)  # Wait for download to start
                                    print(f"  ✓ PDF download initiated")
                                else:
                                    print(f"  ⚠ Could not extract PDF URL, waiting for auto-download...")
                                    time.sleep(2)  # Wait for auto-download with plugins.always_open_pdf_externally

                            except Exception as extract_error:
                                print(f"  ⚠ PDF extraction failed: {str(extract_error)[:100]}")

                            # Switch to the PDF viewer's iframe or shadow DOM context
                            # The Chrome PDF viewer has a download button in its toolbar
                            try:
                                # Try to click the download button in Chrome's PDF viewer
                                # The download button is in the PDF viewer's shadow DOM
                                print(f"  ⚠ Attempting to click Chrome PDF viewer download button...")

                                # Use JavaScript to find and click the download button in Shadow DOM
                                download_script = """
                                // Access Chrome PDF viewer's shadow DOM
                                try {
                                    // Method 1: Try to find the PDF viewer embed element
                                    let pdfViewer = document.querySelector('embed[type="application/pdf"]');
                                    if (pdfViewer) {
                                        // Look for the toolbar in parent elements
                                        let toolbar = pdfViewer.parentElement.querySelector('pdf-viewer-toolbar');
                                        if (toolbar && toolbar.shadowRoot) {
                                            let downloadBtn = toolbar.shadowRoot.querySelector('#download');
                                            if (downloadBtn) {
                                                downloadBtn.click();
                                                return 'clicked_shadow';
                                            }
                                        }
                                    }

                                    // Method 2: Try direct selectors
                                    let downloadButton = document.querySelector('cr-icon-button#download');
                                    if (downloadButton) {
                                        downloadButton.click();
                                        return 'clicked_direct';
                                    }

                                    // Method 3: Look for download icon
                                    downloadButton = document.querySelector('button[aria-label*="Download"]');
                                    if (downloadButton) {
                                        downloadButton.click();
                                        return 'clicked_aria';
                                    }

                                    // Method 4: Look in all shadow roots
                                    let allElements = document.querySelectorAll('*');
                                    for (let el of allElements) {
                                        if (el.shadowRoot) {
                                            let btn = el.shadowRoot.querySelector('button[title*="Download"], cr-icon-button#download, [aria-label*="Download"]');
                                            if (btn) {
                                                btn.click();
                                                return 'clicked_shadow_scan';
                                            }
                                        }
                                    }

                                    return 'not_found';
                                } catch (e) {
                                    return 'error: ' + e.message;
                                }
                                """

                                result = driver.execute_script(download_script)
                                print(f"  📋 JavaScript result: {result}")

                                if 'clicked' in str(result):
                                    print(f"  ✓ Clicked download button via JavaScript ({result})")
                                    time.sleep(1)
                                else:
                                    print(f"  ⚠ Download button not found ({result})")

                            except Exception as js_error:
                                print(f"  ⚠ JavaScript method failed: {str(js_error)[:100]}")

                        elif current_url.endswith('.pdf'):
                            print(f"  ✓ Direct PDF URL detected")
                            # Direct PDF should auto-download
                            time.sleep(2)
                        else:
                            # Try to find a download link/button in the page
                            print(f"  ⚠ Looking for download elements in page...")
                            try:
                                download_selectors = [
                                    (By.XPATH, "//a[contains(@download, '')]"),
                                    (By.XPATH, "//a[contains(text(), 'Download')]"),
                                    (By.XPATH, "//button[contains(text(), 'Download')]"),
                                    (By.CSS_SELECTOR, "a[download]"),
                                ]

                                download_link = None
                                for by, selector in download_selectors:
                                    try:
                                        download_link = driver.find_element(by, selector)
                                        print(f"  ✓ Found download element: {selector}")
                                        download_link.click()
                                        print(f"  ✓ Clicked download element")
                                        time.sleep(2)
                                        break
                                    except:
                                        continue

                                if not download_link:
                                    print(f"  ⚠ No download button found")
                            except Exception as dl_error:
                                print(f"  ⚠ Could not trigger download: {str(dl_error)[:100]}")

                        # Close the new tab
                        driver.close()
                        print(f"  ✓ Closed new tab")

                        # Switch back to original window
                        driver.switch_to.window(original_window)
                        print(f"  ✓ Switched back to original tab")
                    else:
                        print(f"  ✓ Document download initiated (no new tab)")

                    # Wait for the download to complete
                    print(f"  ⏳ Waiting for download to complete...")

                    # Check for new PDFs
                    time.sleep(1.5)  # Give download a moment to start/complete
                    current_pdfs = set(glob.glob(os.path.join(download_dir, "*.pdf")))
                    new_pdfs = current_pdfs - initial_pdfs

                    if new_pdfs:
                        downloaded_file = max(new_pdfs, key=os.path.getmtime)
                        print(f"  ✓ Downloaded: {os.path.basename(downloaded_file)}")
                        file_size = os.path.getsize(downloaded_file) / 1024  # Size in KB
                        print(f"  📊 File size: {file_size:.2f} KB")
                    else:
                        # If not found yet, use wait function
                        downloaded_file = wait_for_download(download_dir, timeout=8)
                        if downloaded_file:
                            print(f"  ✓ Downloaded: {os.path.basename(downloaded_file)}")
                            file_size = os.path.getsize(downloaded_file) / 1024  # Size in KB
                            print(f"  📊 File size: {file_size:.2f} KB")
                        else:
                            print(f"  ⚠ Download may still be in progress or failed")

                    downloaded_count += 1
                    time.sleep(0.5)

                except Exception as e:
                    print(f"  ✗ Could not download document {idx}: {str(e)[:150]}")
                    # Make sure we're back on the original window
                    try:
                        driver.switch_to.window(original_window)
                    except:
                        pass
                    continue

            print(f"\n✓ Successfully processed {downloaded_count} document(s)")

            # Final check of all downloaded files
            import glob
            pdf_files = glob.glob(os.path.join(download_dir, "*.pdf"))

            if pdf_files:
                print(f"\n✓ All downloaded files in {download_dir}:")
                for pdf in pdf_files:
                    file_size = os.path.getsize(pdf) / 1024  # Size in KB
                    print(f"  - {os.path.basename(pdf)} ({file_size:.2f} KB)")
            else:
                print(f"\n⚠ No PDF files found in {download_dir}")

        except Exception as e:
            print(f"✗ Could not download documents: {str(e)[:150]}")

    except Exception as e:
        print(f"\n❌ Error occurred: {type(e).__name__}")
        print(f"Details: {e}")

        import traceback
        print("\nFull traceback:")
        traceback.print_exc()

        if driver:
            try:
                print(f"\nCurrent URL when error occurred: {driver.current_url}")
            except:
                pass

    finally:
        if driver:
            try:
                driver.quit()
                print("\n✓ Browser closed")
            except Exception as e:
                print(f"Error closing browser: {e}")

if __name__ == "__main__":
    print("=" * 70)
    print("Aviva Investors Scraper")
    print("=" * 70)
    print(f"Python version: {sys.version}")
    print(f"Platform: {sys.platform}")
    print("=" * 70 + "\n")

    scrape_aviva_data()
