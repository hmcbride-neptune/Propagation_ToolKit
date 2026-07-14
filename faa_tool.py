# take decimal coordinates and convert to DMS and insert them into the website  'https://oeaaa.faa.gov/oeaaa/oe3a/main/#/noticePrescreen' using microsoft edge the latitude goes in the
# <div id="nct-points-rows" ...><div ckass="col-2"> <div ...>Latitude</div> <input ...>

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, StaleElementReferenceException
import time
from selenium.webdriver.edge.service import Service
from pathlib import Path


"""
FAA automation script with a small Tkinter GUI for batch processing.

Input CSV columns: Sector, Latitude, Longitude, AntHgt(m)
Output CSV columns: Sector, Latitude, Longitude, AntHgt(m), Result

"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
from selenium.webdriver.edge.service import Service
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import csv
import os
import sys
import logging
from datetime import datetime


class SelectTowerError(Exception):
    """Raised when selecting the TOWER structure type fails."""
    pass


def wait_for_page_ready(driver, timeout=20):
    """Wait until document.readyState == 'complete'."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == 'complete'
        )
    except Exception:
        # If waiting for readyState fails, continue; individual element waits will still guard actions
        pass


def wait_for_no_loading(driver, timeout=15, attempts=2):
    """Poll the page for any *visible* element containing the word 'Loading' (case-insensitive).
    This uses a short JS snippet to avoid problems with hidden/offscreen elements or shadow DOM text.

    Returns True when no visible "Loading" text is found. If the string remains visible after
    `timeout` seconds the function will refresh the page and retry (up to `attempts` times).
    """

    check_js = '''
    return (function(){
        var re = /\\bLoading ...\\b/i;
        var nodes = document.querySelectorAll('*');
        for (var i=0;i<nodes.length;i++){
            var el = nodes[i];
            try{
                var text = (el.innerText || '').trim();
                if (!text) continue;
                if (!re.test(text)) continue;
                var style = window.getComputedStyle(el);
                if (!style) continue;
                if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') continue;
                var rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) return true;
            } catch (e) { /* ignore DOM access errors */ }
        }
        return false;
    })();
    '''

    end_time = time.time() + timeout
    for attempt in range(attempts):
        while time.time() < end_time:
            try:
                found = driver.execute_script(check_js)
            except Exception:
                # If script execution fails, break out so we can refresh/ retry
                found = True

            if not found:
                return True

            time.sleep(0.5)

        # still found after timeout for this attempt -> refresh and retry if attempts remain
        if attempt < attempts - 1:
            try:
                print(f"Loading indicator still present after {timeout}s — refreshing (attempt {attempt+1})")
                driver.refresh()
                wait_for_page_ready(driver, timeout=20)
                end_time = time.time() + timeout
                continue
            except Exception:
                # ignore refresh errors and fall through to final return
                pass

        # final attempt expired
        print(f"Loading indicator still present after {attempts} attempts.")
        return False


def decimal_to_dms(decimal):
    # Work with absolute value to avoid negative minutes/seconds
    sign = '-' if decimal < 0 else ''
    abs_val = abs(decimal)
    degrees = int(abs_val)
    minutes = int((abs_val - degrees) * 60)
    seconds = (abs_val - degrees - minutes / 60) * 3600

    # Handle rounding that could push seconds or minutes to the next unit
    seconds = round(seconds, 2)
    if seconds >= 60:
        seconds -= 60
        minutes += 1
    if minutes >= 60:
        minutes -= 60
        degrees += 1

    # Return in ASCII format: "DD MM SS.SS" (zero-padded minutes), include sign if negative
    return f"{sign}{degrees:02d} {minutes:02d} {seconds:05.2f}"


def _create_service(driver_path=r"C:\WebDriver\msedgedriver.exe"):
    logging.getLogger(__name__).debug(f"_create_service entered at {datetime.utcnow().isoformat()} with driver_path={driver_path}")
    start_ts = time.time()
    # If an explicit local driver path was provided and exists, use it
    driver_file = Path(driver_path)
    if driver_file.is_file():
        return Service(str(driver_file))

    # If running as a PyInstaller bundle, the bundled files are unpacked to sys._MEIPASS
    bundled_candidates = []
    try:
        if getattr(sys, 'frozen', False):
            base = getattr(sys, '_MEIPASS', None) or os.path.dirname(sys.executable)
            bundled_candidates.append(os.path.join(base, 'msedgedriver.exe'))
    except Exception:
        pass

    # Also check next to the source file (useful for --onedir or when distributing folder)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    bundled_candidates.append(os.path.join(base_dir, 'msedgedriver.exe'))
    bundled_candidates.append(os.path.join(base_dir, 'drivers', 'msedgedriver.exe'))

    for p in bundled_candidates:
        try:
            if p and os.path.isfile(p):
                logging.getLogger(__name__).info(f"Using bundled msedgedriver at {p}")
                logging.getLogger(__name__).debug(f"_create_service exiting after {time.time()-start_ts:.2f}s")
                return Service(str(p))
        except Exception:
            continue

    # Fallback to webdriver-manager (requires network)
    try:
        logging.getLogger(__name__).info("Attempting to download msedgedriver via webdriver-manager (network required)")
        from webdriver_manager.microsoft import EdgeChromiumDriverManager
        svc = Service(EdgeChromiumDriverManager().install())
        logging.getLogger(__name__).debug(f"_create_service exiting after {time.time()-start_ts:.2f}s (downloaded)")
        return svc
    except Exception as e:
        logging.getLogger(__name__).exception("webdriver-manager failed to download msedgedriver")
        raise RuntimeError(
            "Could not locate a local msedgedriver, no bundled driver was found, and webdriver-manager failed to download one.\n"
            "To create a portable executable, include a matching msedgedriver.exe next to the application executable and rebuild.\n"
            f"Searched paths: {driver_path}, {bundled_candidates}\n" + str(e)
        )


def create_driver(driver_path=r"C:\WebDriver\msedgedriver.exe"):
    logging.getLogger(__name__).info(f"create_driver: starting at {datetime.utcnow().isoformat()}")
    t0 = time.time()
    service = _create_service(driver_path)
    driver = webdriver.Edge(service=service)
    try:
        driver.maximize_window()
    except Exception:
        pass
    logging.getLogger(__name__).info(f"create_driver: driver created in {time.time()-t0:.2f}s")
    return driver


# Base URL used by the automation flow
BASE_URL = 'https://oeaaa.faa.gov/oeaaa/oe3a/main/#/noticePrescreen'

# How many times to attempt a site before giving up (1 = no retry, 2 = one retry)
SITE_MAX_ATTEMPTS = 2


def reset_page(driver):
    """Reload the prescreen page and wait until it's ready. Use between sites."""
    driver.get(BASE_URL)
    wait_for_page_ready(driver, timeout=20)
    # best-effort wait for SPA loading indicators
    try:
        wait_for_no_loading(driver, timeout=15, attempts=2)
    except Exception:
        # non-fatal; caller will continue
        pass


def process_site(driver, decimal_latitude, decimal_longitude, height_m):
    """Run the automated flow for a single site. Returns result string or raises on error."""
    url = 'https://oeaaa.faa.gov/oeaaa/oe3a/main/#/noticePrescreen'
    driver.get(url)
    wait_for_page_ready(driver, timeout=20)
    wait_for_no_loading(driver, timeout=15, attempts=2)
    wait = WebDriverWait(driver, 20)

    # Select structure type TOWER
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.ID, "oe3a-sttype-tree-btn")))
        dropdown.click()
        towers_section = wait.until(EC.element_to_be_clickable((By.ID, "cat-TOWER")))
        driver.execute_script("arguments[0].scrollIntoView(true);", towers_section)
        driver.execute_script("arguments[0].click();", towers_section)
        tower_expand = wait.until(EC.element_to_be_clickable((By.ID, "class-TOWER")))
        driver.execute_script("arguments[0].click();", tower_expand)
        tower_select = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@id='type-TOWER']//span[contains(@class,'oe3a-st-type-label')]")))
        driver.execute_script("arguments[0].click();", tower_select)
    except Exception as e:
        # Raise a specific exception so callers can offer a retry to the user
        raise SelectTowerError(f"Could not select TOWER: {e}")
    
    # Convert coordinates
    dms_latitude = decimal_to_dms(decimal_latitude)
    dms_longitude = decimal_to_dms(decimal_longitude)
    
    # Input lat/lon
    try:
        lat_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[ng-model='point.formLat2']")))
        lon_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[ng-model='point.formLon2']")))
    except Exception:
        lat_input = wait.until(EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='DMS'])[1]")))
        lon_input = wait.until(EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='DMS'])[2]")))

    lat_input.clear()
    lat_input.send_keys(dms_latitude)
    lon_input.clear()
    lon_input.send_keys(dms_longitude)

    # Datum
    try:
        datum_elem = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[ng-model='point.datum']")))
    except Exception:
        datum_elem = wait.until(EC.element_to_be_clickable((By.NAME, "datum")))
    select = Select(datum_elem)
    select.select_by_visible_text("NAD83")
    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", datum_elem)

    # Height: input expects feet; convert from meters
    height_ft = round(float(height_m) / 0.3048) + 20 if height_m is not None else ''
    try:
        height_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#case-height-entry")))
    except Exception:
        height_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[ng-model='point.structureHeight']")))
    height_input.clear()
    height_input.send_keys(str(height_ft))
    driver.execute_script("arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", height_input)

    time.sleep(1)  # brief pause to allow any reactive updates before proceeding
    # Validate -> Accept -> Submit
    validate_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='VALIDATE']")))
    driver.execute_script("arguments[0].click();", validate_btn)
    
    accept_btn = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[.='Accept Point']")))
    time.sleep(1)
    # Robust click: scroll into view, JS click, retry on intercept/stale
    clicked = False
    for attempt in range(3):
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", accept_btn)
            # try normal click first
            try:
                accept_btn.click()
            except (ElementClickInterceptedException, StaleElementReferenceException):
                # fallback to JS click
                driver.execute_script("arguments[0].click();", accept_btn)

            # give the page a moment to react and clear any loading indicators
            wait_for_no_loading(driver, timeout=15, attempts=2)

            # verify that the Submit button appears (indicates Accept succeeded)
            try:
                WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and normalize-space()='Submit']")))
                clicked = True
                break
            except Exception:
                # maybe the DOM changed; try to re-find accept_btn and retry
                try:
                    accept_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[.='Accept Point']")))
                except Exception:
                    pass
                time.sleep(0.5)
                continue
        except StaleElementReferenceException:
            # re-acquire element and retry
            try:
                accept_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[.='Accept Point']")))
            except Exception:
                pass
            continue
        except Exception as e:
            # last attempt: raise so caller can handle
            if attempt == 2:
                raise
            time.sleep(0.5)

    if not clicked:
        # final attempt using JS click (best-effort)
        try:
            driver.execute_script("arguments[0].click();", accept_btn)
            wait_for_no_loading(driver, timeout=10, attempts=1)
        except Exception:
            # allow process_site to continue; later steps will surface failure
            pass
        
    submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and normalize-space()='Submit']")))
    driver.execute_script("arguments[0].click();", submit_btn)

    time.sleep(2)

    # Read result text
    result_text = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Based on the information you provided')]"))
    ).text

    if "not required to file notice" in result_text.lower():
        return "Notice not exceeded"
    elif "required to file notice" in result_text.lower():
        return "NOTICE EXCEEDED"
    else:
        return "UNKNOWN"


def run_batch(input_csv, output_csv, driver_path=r"C:\WebDriver\msedgedriver.exe", progress_callback=None, prompt_callback=None, start_callback=None):
    """Read sites from input_csv and write results to output_csv.
    Expected input CSV columns: Sector, Latitude, Longitude, AntHgt(m)
    progress_callback(i, total, message) optional; start_callback(i, total, sector) optional.
    prompt_callback(title, message) should return True/False for retry prompts.
    """
    results = []

    # Read input CSV
    rows = []
    with open(input_csv, 'r', newline='', encoding='utf-8') as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append(r)

    total = len(rows)
    if total == 0:
        raise RuntimeError('Input CSV contains no rows')

    # Create driver once and reuse across rows
    driver = create_driver(driver_path)

    try:
        for i, row in enumerate(rows, start=1):
            # Parse expected columns; be permissive with names
            sector = row.get('Sector') or row.get('Location') or row.get('site') or row.get('Site') or ''
            try:
                lat = float(row.get('Latitude') or row.get('Lat') or row.get('Lat_dec') or row.get('lat') or 0)
            except Exception:
                lat = 0.0
            try:
                lon = float(row.get('Longitude') or row.get('Lon') or row.get('Long_dec') or row.get('lon') or 0)
            except Exception:
                lon = 0.0
            try:
                elev_m_val = float(row.get('Elev(m)') or row.get('Elev_m') or row.get('elev_m') or row.get('AntHgt(m)') or row.get('AntHgt(m)') or 0)
            except Exception:
                elev_m_val = None
            try:
                height_m_val = float(row.get('AntHgt(m)') or row.get('Height_m') or row.get('ant_hgt_m') or row.get('AntHgt') or 0)
            except Exception:
                height_m_val = None

            # Notify that this site is starting (updates current-site display only)
            if start_callback:
                try:
                    start_callback(i, total, sector)
                except Exception:
                    pass

            # Attempt the site, retrying until a correct result is obtained.
            MAX_SAFE_ATTEMPTS = 5
            SUCCESS_RESULTS = {"Notice not exceeded", "NOTICE EXCEEDED"}

            result = None
            attempt = 0
            while True:
                attempt += 1
                try:
                    result = process_site(driver, lat, lon, height_m_val)
                except SelectTowerError as e:
                    # Prompt user whether to retry selecting TOWER
                    full_err = str(e)
                    short_err = (full_err.splitlines()[0] or '')[:200]
                    retry = False
                    if prompt_callback:
                        try:
                            retry = bool(prompt_callback('ERROR', short_err))
                        except Exception:
                            retry = False
                    else:
                        try:
                            retry = messagebox.askretrycancel('ERROR', short_err)
                        except Exception:
                            retry = False

                    if retry:
                        logging.getLogger(__name__).info(f"User chose to retry selecting TOWER (attempt {attempt}) for {sector}")
                        try:
                            reset_page(driver)
                        except Exception:
                            pass
                        continue
                    else:
                        result = f"ERROR: {e}"
                        break
                except Exception as e:
                    result = f"ERROR: {e}"

                # If result is a success value, we're done
                if isinstance(result, str) and result in SUCCESS_RESULTS:
                    break

                # If result is UNKNOWN or starts with ERROR, decide whether to retry
                should_retry = False
                if isinstance(result, str) and (result.startswith('ERROR') or result == 'UNKNOWN'):
                    if attempt < MAX_SAFE_ATTEMPTS:
                        should_retry = True
                    else:
                        # ask the user whether to continue trying
                        try:
                            if prompt_callback:
                                cont = bool(prompt_callback('Repeated Failure', f"Site {sector} failed {attempt} times. Continue trying?"))
                            else:
                                cont = messagebox.askretrycancel('Repeated Failure', f"Site {sector} failed {attempt} times. Continue trying?")
                        except Exception:
                            cont = False
                        if cont:
                            attempt = 0
                            should_retry = True
                        else:
                            should_retry = False

                if should_retry:
                    logging.getLogger(__name__).info(f"Retrying {sector} (next attempt) after failure")
                    try:
                        reset_page(driver)
                    except Exception:
                        pass
                    continue

                # No more retries; break and record result (could be success or final error)
                break

            # Compose output row
            out_row = {
                'Location': sector,
                'Lat_dec': lat,
                'Long_dec': lon,
                'Latitude(N)': decimal_to_dms(lat),
                'Longitude(W)': decimal_to_dms(lon),
                'Structure': "",
                'Elev(m)': elev_m_val if elev_m_val is not None else '',
                'AGLw/oApp(m)': height_m_val if height_m_val is not None else '',
                'AGLw/App(m)': round((height_m_val or 0) + 6, 4) if height_m_val is not None else '',
                'elev+anthgt': round((height_m_val or 0) + 6 + (elev_m_val or 0), 4) if (height_m_val is not None and elev_m_val is not None) else '',
                'AntHgt(ft)': round(float(height_m_val) / 0.3048) if height_m_val is not None else '',
                'HAAT (m)': "",
                'FAA Result': result
            }

            results.append(out_row)

            # Update progress display
            if progress_callback:
                try:
                    progress_callback(i, total, result or '')
                except Exception:
                    pass

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # Write results to output CSV
    fieldnames = ['Location', 'Lat_dec', 'Long_dec', 'Latitude(N)', 'Longitude(W)', 'Structure', 'Elev(m)', 'AGLw/oApp(m)', 'AGLw/App(m)', 'elev+anthgt', 'AntHgt(ft)', 'HAAT (m)', 'FAA Result']
    os.makedirs(os.path.dirname(output_csv), exist_ok=True)
    with open(output_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in results:
            writer.writerow(r)

    return output_csv


def start_gui():
    root = tk.Tk()
    root.title('FAA Prescreener - Batch')
    root.geometry('620x260')

    # configure logging to file for GUI/exe runs (no console)
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'faa.log')
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(name)s: %(message)s',
                        handlers=[logging.FileHandler(log_file, encoding='utf-8')])

    logging.getLogger(__name__).info('GUI start invoked')

    input_var = tk.StringVar()
    project_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar()
    current_site_var = tk.StringVar()
    top_var = tk.BooleanVar(value=True)

    # Keep the window on top if requested
    try:
        root.attributes('-topmost', True)
    except Exception:
        pass

    def choose_input():
        path = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if path:
            input_var.set(path)
            # try to infer Project from path
            try:
                parts = path.replace('\\', '/').split('/')
                if 'Projects' in parts:
                    idx = parts.index('Projects')
                    if idx + 1 < len(parts):
                        proj = parts[idx + 1]
                        project_var.set(proj)
            except Exception:
                pass
            # default output based on inferred project
            proj = project_var.get() or 'Project'
            yy = time.strftime('%y%m%d')
            default_out_dir = os.path.join('C:\\ProgramData\\EDX\\Projects', proj, 'Data\\xReports')
            default_name = f"{proj} {yy} - FAA-ASR Results.csv"
            output_var.set(os.path.join(default_out_dir, default_name))

    def choose_output():
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')])
        if path:
            output_var.set(path)

    def progress(i, total, msg):
        try:
            pct = int((i / total) * 100) if total else 0
        except Exception:
            pct = 0
        status_var.set(f"{i}/{total} {pct}% - {msg}")
        current_site_var.set('')

    def start_progress(i, total, sector):
        try:
            current_site_var.set(f"Currently Evaluating: {sector}")
        except Exception:
            pass

    def run_thread():
        in_path = input_var.get()
        out_path = output_var.get()
        if not in_path:
            messagebox.showerror('Error', 'Select an input CSV first', parent=root)
            return

        def worker():
            try:
                status_var.set('Starting...')

                class PromptInvoker:
                    def __init__(self, root_widget):
                        self._root = root_widget

                    def askretrycancel(self, title, message):
                        event = threading.Event()
                        result = {'value': False}

                        def show():
                            try:
                                res = messagebox.askretrycancel(title, message, parent=self._root)
                            except Exception:
                                res = False
                            result['value'] = bool(res)
                            event.set()

                        try:
                            self._root.after(0, show)
                            event.wait()
                        except Exception:
                            return False
                        return result['value']

                invoker = PromptInvoker(root)
                try:
                    root.attributes('-topmost', bool(top_var.get()))
                except Exception:
                    pass

                run_batch(in_path, out_path, progress_callback=progress, prompt_callback=invoker.askretrycancel, start_callback=start_progress)
                status_var.set('Done.')
                try:
                    root.after(100, root.destroy)
                except Exception:
                    pass
            except Exception as e:
                logging.getLogger(__name__).exception('Unhandled exception in worker')
                short = str(e) or 'An unexpected error occurred.'
                if len(short) > 200:
                    short = short.splitlines()[0][:200] + '...'
                status_var.set(f'Error: {short}')
                try:
                    messagebox.showerror('Error', f"{short}\nSee log file for details: {log_file}", parent=root)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    # Build UI
    tk.Label(root, text='Input CSV:').pack(anchor='w', padx=10, pady=(10, 0))
    frame1 = tk.Frame(root)
    frame1.pack(fill='x', padx=10)
    tk.Entry(frame1, textvariable=input_var).pack(side='left', fill='x', expand=True)
    tk.Button(frame1, text='Browse', command=choose_input).pack(side='right')

    tk.Button(root, text='Run', command=run_thread, bg='#4caf50', fg='white').pack(pady=10)
    try:
        status_font = ('TkDefaultFont', 14, 'bold')
    except Exception:
        status_font = None

    tk.Label(root, textvariable=status_var, font=status_font).pack(fill='x', padx=10)
    tk.Label(root, textvariable=current_site_var, fg='#1565c0', font=status_font).pack(fill='x', padx=10)

    root.mainloop()


if __name__ == '__main__':
    start_gui()
