# ODOO Automation Script - Excel Integration
import os, sys, time, logging, random, pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, StaleElementReferenceException, ElementClickInterceptedException

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(funcName)s : %(lineno)d - %(message)s')
logger = logging.getLogger(__name__)
log_file = "automation_log.txt"

def logger_debug(pesan):
    waktu = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    log_message = f"{waktu} {pesan}"
    
    # Tulis ke file dengan encoding UTF-8
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"{log_message}\n")
    except Exception as e:
        # Fallback: tulis tanpa karakter khusus
        safe_message = log_message.encode('ascii', 'ignore').decode('ascii')
        with open(log_file, "a") as f:
            f.write(f"{safe_message}\n")
    
    print(log_message, flush=True)


class ExcelDataProcessor:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.data = None
        self.load_excel_data()
    
    def load_excel_data(self):
        """Load data from Excel file"""
        try:
            self.data = pd.read_excel(self.excel_file_path)
            logger.info(f"Loaded Excel file with {len(self.data)} rows")
            logger.info(f"Columns: {list(self.data.columns)}")
        except Exception as e:
            logger.error(f"Failed to load Excel file: {e}")
            raise
    
    def get_row_data(self, row_index):
        """Get data for specific row"""
        if self.data is None or row_index >= len(self.data):
            return None
        return self.data.iloc[row_index]
    
    def should_duplicate(self, current_row_index):
        """Check if next row has same kode_benda_uji and proyek"""
        if self.data is None or current_row_index >= len(self.data) - 1:
            return False
        
        current_row = self.get_row_data(current_row_index)
        next_row = self.get_row_data(current_row_index + 1)
        
        if current_row is None or next_row is None:
            return False
        
        # Get kode_benda_uji from column 3 (index 2)
        current_kode = str(current_row.iloc[2]) if len(current_row) > 2 else ""
        next_kode = str(next_row.iloc[2]) if len(next_row) > 2 else ""
        
        # Get proyek from column 4 (index 3)
        current_proyek = str(current_row.iloc[3]) if len(current_row) > 3 else ""
        next_proyek = str(next_row.iloc[3]) if len(next_row) > 3 else ""
        
        # Check if both values are the same
        same_kode = current_kode == next_kode
        same_proyek = current_proyek == next_proyek
        
        logger.info(f"Duplicate check - Current row {current_row_index + 1}:")
        logger.info(f"  Current kode_benda_uji: {current_kode}")
        logger.info(f"  Next kode_benda_uji: {next_kode}")
        logger.info(f"  Current proyek: {current_proyek}")
        logger.info(f"  Next proyek: {next_proyek}")
        logger.info(f"  Same kode: {same_kode}, Same proyek: {same_proyek}")
        
        return same_kode and same_proyek

def resource_path(relative_path: str) -> str:
    """Get resource file path for both .py and .exe execution"""
    try:
        base_path = sys._MEIPASS  # type: ignore # PyInstaller bundle
    except AttributeError:
        base_path = os.path.abspath(".")  # Direct execution
    return os.path.join(base_path, relative_path)

def generate_random_slump_test(slump_rencana):
    """Generate random slump test value (±2 from slump_rencana)"""
    try:
        base_value = float(slump_rencana)
        result = base_value + random.uniform(-1, 2)
        return str(int(round(result)))  # Convert to integer to remove decimal
    except:
        result = random.uniform(11, 14)
        return str(int(round(result)))  # Convert to integer to remove decimal

def generate_random_yield():
    """Generate random yield value between 0.97-0.99"""
    return str(round(random.uniform(0.97, 0.99), 2))

def calculate_jam_sample(base_time):
    """Calculate jam sample by adding 1:10 to 1:50 hours to base time"""
    try:
        # Check if base_time contains date (format: 'YYYY-MM-DD HH:MM:SS' or 'YYYY-MM-DD HH:MM')
        if ' ' in str(base_time):
            # Extract time part from datetime string
            time_part = str(base_time).split(' ')[1]
            # Handle seconds if present
            if time_part.count(':') == 2:
                time_part = ':'.join(time_part.split(':')[:2])  # Keep only HH:MM
            base_hour, base_minute = map(int, time_part.split(':'))
        else:
            # Original format HH:MM
            base_hour, base_minute = map(int, str(base_time).split(':'))
        
        # Add random time between 1:5 to 1:30 hours
        additional_minutes = random.randint(65, 90)
        
        total_minutes = base_hour * 60 + base_minute + additional_minutes
        final_hour = (total_minutes // 60) % 24
        final_minute = total_minutes % 60
        
        return f"{final_hour:02d}:{final_minute:02d}"
    except:
        # Fallback to random time
        hour = random.randint(10, 15)
        minute = random.randint(0, 59)
        return f"{hour:02d}:{minute:02d}"

def quick_delete_all(driver):
    """Delete all rows by clicking delete buttons"""
    deleted_count = 0
    try:
        while True:
            delete_buttons = driver.find_elements(By.CSS_SELECTOR, 'tr[data-id^="one2many_v_id_"] td.o_list_record_delete .fa-trash-o')
            if not delete_buttons:
                break
            
            first_button = delete_buttons[0]
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", first_button)
                time.sleep(0.1)
                first_button.click()
                deleted_count += 1
                time.sleep(0.2)
            except Exception as e:
                logger.error(f"Failed to click delete button: {e}")
                break
        
        logger.info(f"Quick delete: {deleted_count} rows deleted!")
        return deleted_count
    except Exception as e:
        logger.error(f"Error in quick_delete_all: {e}")
        return deleted_count

def quick_delete_excess_rows(driver, rows_to_delete):
    """Delete specific number of excess rows starting from the last row"""
    deleted_count = 0
    try:
        for i in range(rows_to_delete):
            delete_buttons = driver.find_elements(By.CSS_SELECTOR, 'tr[data-id^="one2many_v_id_"] td.o_list_record_delete .fa-trash-o')
            if not delete_buttons:
                break
            
            # Get the last button instead of first
            last_button = delete_buttons[-1]
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", last_button)
                time.sleep(0.1)
                last_button.click()
                deleted_count += 1
                time.sleep(0.2)
            except Exception as e:
                logger.error(f"Failed to click delete button: {e}")
                break
        
        logger.info(f"Quick delete excess: {deleted_count} rows deleted from bottom!")
        return deleted_count
    except Exception as e:
        logger.error(f"Error in quick_delete_excess_rows: {e}")
        return deleted_count

def select_first_row_in_modal_and_confirm(driver, wait, row_text: str | None = None, absolute_xpath: str | None = None):
    """Select first row in modal and confirm selection"""
    max_attempts = 1
    last_error = None
    
    for attempt in range(1, max_attempts + 1):
        try:
            # Find visible modal
            visible_modals = driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']")
            modal = visible_modals[-1] if visible_modals else wait.until(EC.visibility_of_element_located((
                By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"
            )))

            # Find table in modal
            table = None
            for selector in [
                ".modal-content table.o_list_view",
                ".modal-content table.table.o_list_view",
                ".modal-content table.table-condensed.table-striped.o_list_view"
            ]:
                try:
                    table = modal.find_element(By.CSS_SELECTOR, selector)
                    break
                except NoSuchElementException:
                    continue
            
            if table is None:
                try:
                    table = modal.find_element(By.XPATH, ".//table[contains(@class,'o_list_view')]")
                except NoSuchElementException:
                    raise NoSuchElementException("No list table found inside modal")

            target_row = None

            # Search by row_text first
            if row_text:
                row_xpath = f".//tbody/tr[.//td[contains(normalize-space(), \"{row_text}\")] or contains(normalize-space(), \"{row_text}\") or contains(., \"{row_text}\")]"
                matching_rows = table.find_elements(By.XPATH, row_xpath)
                if matching_rows:
                    target_row = matching_rows[0]

            # Fallback to absolute xpath
            if target_row is None and absolute_xpath:
                try:
                    target_row = driver.find_element(By.XPATH, absolute_xpath)
                except NoSuchElementException:
                    target_row = None

            # Fallback to first row
            if target_row is None:
                rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
                if not rows:
                    rows = table.find_elements(By.XPATH, ".//tbody/tr")
                if not rows:
                    raise NoSuchElementException("No rows found inside the modal list view")
                target_row = rows[0]

            # Scroll and click
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_row)
            time.sleep(0.05)

            # Find clickable element
            clickable_selectors = [
                "td.o_list_record_selector input",
                "td.o_list_record_selector",
                "td a",
                "td"
            ]
            clickable = None
            for sel in clickable_selectors:
                try:
                    clickable = target_row.find_element(By.CSS_SELECTOR, sel)
                    break
                except NoSuchElementException:
                    continue
            
            if clickable is None:
                clickable = target_row

            try:
                driver.execute_script("arguments[0].click();", clickable)
            except WebDriverException:
                ActionChains(driver).move_to_element(clickable).pause(0.05).click().perform()

            # Click Select button in modal footer
            try:
                select_btn = modal.find_element(By.CSS_SELECTOR, ".modal-footer .btn.btn-primary, .modal-footer .o_select_button")
                try:
                    driver.execute_script("arguments[0].click();", select_btn)
                except WebDriverException:
                    select_btn.click()
            except NoSuchElementException:
                pass

            # Wait for modal to disappear
            try:
                WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((
                    By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"
                )))
            except TimeoutException:
                try:
                    WebDriverWait(driver, 2).until(EC.staleness_of(modal))
                except TimeoutException:
                    if not driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"):
                        return
            return

        except StaleElementReferenceException as e:
            last_error = e
            logger.info(f"Retry selecting row in modal due to stale element (attempt {attempt}/{max_attempts})")
            time.sleep(0.2)
            continue
        except (TimeoutException, NoSuchElementException, WebDriverException) as e:
            last_error = e
            logger.info(f"Retry selecting row in modal due to transient error: {e} (attempt {attempt}/{max_attempts})")
            time.sleep(0.2)
            continue

    # Check if modal is gone (success)
    if not driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"):
        logger.info("Modal no longer visible; treating selection as successful.")
        return

    logger.error(f"Failed selecting row in modal after {max_attempts} attempts: {last_error}")
    if last_error is not None:
        raise last_error
    else:
        raise Exception("Failed selecting row in modal")

def fill_field(driver, wait, xpath, value, field_name):
    """Generic function to fill form fields"""
    logger.info(f"Filling {field_name} field with value: {value}")
    try:
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
        time.sleep(0.2)
        field.click()
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys(Keys.DELETE)
        field.send_keys(str(value))
        field.send_keys(Keys.TAB)
        
        # Verify value entered
        try:
            wait.until(lambda d: str(value) in (field.get_attribute('value') or ''))
            logger.info(f"{field_name} value entered successfully")
        except TimeoutException:
            logger.warning(f"{field_name} field did not reflect the typed value immediately")
    except TimeoutException:
        logger.error(f"Could not find {field_name} field")
        raise

def data_to_input(driver, no_urut, row_data, is_first_row=False):
    """Input data to the table row using Excel data"""
    wait = WebDriverWait(driver, 5)
    # Determine test age based on sequence number
    rencana_umur_test = "7" if no_urut in [1, 2] else "28"
    # Get kode benda uji from Excel (column 3, index 2)
    kode_benda_uji = str(row_data.iloc[2]) if len(row_data) > 2 else f" Isi Kode Benda Uji - {no_urut}"
    bentuk_benda_uji = "Silinder 15 x 30 cm"
    logger.info(f"Input Row {no_urut} on data table...")
    time.sleep(1)
    # Only click on the first row if it's the first iteration
    if is_first_row:
        first_row = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_'] td.o_list_number[data-field='nomor_urut']")
                ))
        first_row.click()

    else:
        # For rows 2-4, try to find and click the specific row
        try:
            # Try to find the row with the specific number
            rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_'] td.o_list_number[data-field='nomor_urut']")
            if len(rows) >= no_urut:
                target_row = rows[no_urut - 1]  # Index is 0-based, no_urut is 1-based
                target_row.click()
            else:
                logger.warning(f"Row {no_urut} not found, using first available row")
                target_row = rows[0] if rows else None
                if target_row:
                    target_row.click()
        except Exception as e:
            logger.warning(f"Error finding row {no_urut}: {e}")
    # Fill all fields for the row
    base_xpath = "/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/table/tbody/tr/td/div/div[2]/div[1]"
    fields_data = [
        (f"{base_xpath}/input[1]", str(no_urut), "No Urut"),
        (f"{base_xpath}/input[2]", kode_benda_uji, "Kode Benda Uji"),
        (f"{base_xpath}/div[1]/div/input", rencana_umur_test, "Rencana Umur Test"),
    ]
    for xpath, value, field_name in fields_data:
        logger.info(f"Filling {field_name} with value: {value}")
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys(Keys.DELETE)
        field.send_keys(str(value))
        time.sleep(1)

    wait_for_loading_overlay_to_disappear(driver, wait)

    if rencana_umur_test == "7":
        umur7 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[3]/li[1]/a")))
        umur7.click()
    else:
        umur28 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[3]/li/a")))
        umur28.click()

    # Fill Bentuk Benda Uji
    logger.info(f"Filling Bentuk Benda Uji field for row {no_urut}...")
    bentuk_benda_uji_field = driver.find_element(By.CSS_SELECTOR, '[data-fieldname="bentuk_benda_uji"] .o_form_input')
    bentuk_benda_uji_field.click()
    bentuk_benda_uji_field.send_keys(Keys.CONTROL, "a")
    bentuk_benda_uji_field.send_keys(Keys.DELETE) 
    bentuk_benda_uji_field.send_keys(bentuk_benda_uji)
    time.sleep(1)
    silinder_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[4]/li[1]/a")))
    silinder_select.click
    time.sleep(1)

    # Fill Tempat Pengetesan
    logger.info(f"Filling Tempat Pengetesan field for row {no_urut}...")
    tempat_field = wait.until(EC.element_to_be_clickable((By.XPATH, f"{base_xpath}/select")))
    tempat_field.click()
    time.sleep(1)
    tempat_internal_select = wait.until(EC.element_to_be_clickable((By.XPATH, f"{base_xpath}/select/option[2]")))
    tempat_internal_select.click()
    time.sleep(1)

def setup_driver():
    """Setup Chrome driver"""
    chromedriver_path = resource_path("chromedriver.exe")
    service = Service(executable_path=chromedriver_path)
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver

def login(driver, wait):
    """Handle login process"""
    username = "oenoseven@gmail.com"
    password = "rmc"
    login_url = "https://rmc.adhimix.web.id/web/login"
    logger.info("Navigating to login page...")
    driver.get(login_url)
    logger.info("Entering credentials...")
    username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
    password_field = driver.find_element(By.ID, "password")
    username_field.send_keys(username)
    password_field.send_keys(password)
    logger.info("Clicking login button...")
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn-primary")))
    login_button.click()
    time.sleep(3)
    logger.info(f"Current URL after login: {driver.current_url}")

def navigate_and_create(driver, wait):
    wait_for_loading_overlay_to_disappear(driver, wait)
    """Navigate to create page and click create button"""
    create_rbu_url = "https://rmc.adhimix.web.id/web?#min=1&limit=80&view_type=list&model=schedule.truck.mixer.benda.uji&menu_id=535"
    logger.info("Navigating to CREATE RENCANA BENDA UJI page...")
    driver.get(create_rbu_url)
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("CREATE RENCANA BENDA UJI...")
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/button[1]")))
    create_button.click()

def fill_proyek_form(driver, wait, row_data):
    """Fill main form fields using Excel data"""
    # Date field - from Excel column 1 (index 0)
    wait_for_loading_overlay_to_disappear(driver, wait)
    tgl_mulai_prod = str(row_data.iloc[0]) if len(row_data) > 0 else "None"
    logger.info(f"Filling Date form with: {tgl_mulai_prod}")
    tgl_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[1]/td[2]/div/input")))
    tgl_field.clear()
    tgl_field.send_keys(tgl_mulai_prod)
    # Proyek field - from Excel column 4 (index 3)
    proyek = str(row_data.iloc[3]) if len(row_data) > 3 else "-"
    logger.info(f"Filling Proyek field with: {proyek}")
    proyek_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[2]/td[2]/div/div/input")))
    proyek_field.clear()
    proyek_field.send_keys(proyek)
    time.sleep(3)
    
    if proyek == "JALAN TOL AKSES PATIMBAN":
        proyek_option = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[1]/li[2]")))
        proyek_option.click()
        logger.info("Selected 'JALAN TOL AKSES PATIMBAN' from dropdown")
    else:
        try:
            dropdown_option = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[1]/li[1]")))
            dropdown_option.click()
            logger.info("Proyek selected")
        except TimeoutException:
            logger.error("Proyek dropdown not found")

def fill_docket_form(driver, wait, row_data):
    """Fill docket form using Excel data"""
    # No. Docket field - from Excel column 2 (index 1)
    wait_for_loading_overlay_to_disappear(driver, wait)
    no_docket = str(row_data.iloc[1]) if len(row_data) > 1 else "None"
    logger.info(f"Filling No. Docket field with: {no_docket}")
    # Click and fill the No. Docket field
    no_docket_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[3]/td[2]/div/div/input")))
    driver.execute_script("arguments[0].scrollIntoView(true);", no_docket_field)
    time.sleep(1)
    no_docket_field.click()
    no_docket_field.clear()
    no_docket_field.send_keys(no_docket)
    time.sleep(3)
    
    try:
        time.sleep(1)
        # Check if the element exists without waiting
        xpath = f"//ul[@class='ui-autocomplete ui-front ui-menu ui-widget ui-widget-content']//a[text()='{no_docket}']"
        elements = driver.find_elements(By.XPATH, xpath)
        
        if elements and elements[0].is_displayed():
            # Element found, click it
            logger.info(f"Found {no_docket} in autocomplete dropdown")
            elements[0].click()
            logger.info(f"Successfully clicked on {no_docket} from autocomplete dropdown")
        else:
            logger.info(f"Element with text '{no_docket}' not found in autocomplete dropdown")
            logger.info("Using 'Search more...' option as fallback")
            # Search more option
            search_more = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[2]/li[8]/a")))
            search_more.click()
            # Search in modal
            modal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".modal-content")))
            try:
                search_input = modal.find_element(By.CSS_SELECTOR, ".o_searchview input.o_searchview_input")
            except NoSuchElementException:
                search_input = modal.find_element(By.CSS_SELECTOR, ".o_searchview input.o_input")
            search_input.click()
            search_input.clear()
            search_input.send_keys(no_docket)
            search_input.send_keys(Keys.ENTER)
            # Wait for search results and select
            wait.until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(@class,'modal-content')]//div[contains(@class,'o_searchview_facet')][.//span[contains(@class,'o_searchview_facet_label')][normalize-space()='No. Docket']]//div[contains(@class,'o_facet_values')]//span[contains(normalize-space(), \"{no_docket}\")]")))
            time.sleep(3)
            select_first_row_in_modal_and_confirm(driver, wait, row_text=no_docket)   
    except Exception as e:
        logger.error(f"Error in docket selection: {str(e)}")
        raise

    # Fill remaining fields using Excel data
    slump_rencana = str(row_data.iloc[6]) if len(row_data) > 6 else "12"  # Column 7 (index 6)
    slump_test = generate_random_slump_test(slump_rencana)
    yield_value = generate_random_yield()
    nama_teknisi = str(row_data.iloc[4]) if len(row_data) > 4 else "TEKNISI"  # Column 5 (index 4)
    base_jam = str(row_data.iloc[8]) if len(row_data) > 8 else "10:30"  # Column 9 (index 8)
    jam_sample = calculate_jam_sample(base_jam)
    base_xpath = "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[2]/tbody/tr"
    form_fields = [
        (f"{base_xpath}[2]/td[2]/input", slump_rencana, "Slump Rencana"),
        (f"{base_xpath}[3]/td[2]/input", slump_test, "Slump Test"),
        (f"{base_xpath}[4]/td[2]/input", yield_value, "Yield"),
        (f"{base_xpath}[5]/td[2]/input", nama_teknisi, "Nama Teknisi"),
        (f"{base_xpath}[6]/td[2]/input", jam_sample, "Jam Sample")
    ]
    for xpath, value, field_name in form_fields:
        time.sleep(1)
        fill_field(driver, wait, xpath, value, field_name)

def add_table_rows(driver, wait, row_data):
    """Add and fill table rows using Excel data"""
    logger.info("Processing table rows...")
    
    # Check how many existing rows with data-id^='one2many_v_id' exist
    existing_rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_']")
    existing_count = len(existing_rows)
    logger.info(f"Found {existing_count} existing rows")
    
    # Delete excess rows if more than 4 exist
    if existing_count > 4:
        logger.info(f"More than 4 rows exist ({existing_count}), deleting excess rows...")
        rows_to_delete = existing_count - 4
        deleted_count = quick_delete_excess_rows(driver, rows_to_delete)
        logger.info(f"Deleted {deleted_count} excess rows")
        time.sleep(1)
        # Recheck existing count after deletion
        existing_rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_']")
        existing_count = len(existing_rows)
        logger.info(f"After deletion: {existing_count} rows remain")
    
    # Get add item link
    add_item_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Add an item")))
    
    if existing_count < 4:
        # Fill existing rows first, then add new ones
        logger.info("Less than 4 rows exist, filling existing rows first...")
        
        # Fill existing rows in order
        for no_urut in range(1, existing_count + 1):
            logger.info(f"Filling existing row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1))
        
        # Add and fill remaining rows using add item link
        for no_urut in range(existing_count + 1, 5):
            logger.info(f"Adding new row {no_urut}...")
            add_item_link.click()
            time.sleep(1)  # Wait for row to be added
            data_to_input(driver, no_urut, row_data, is_first_row=False)
    else:
        # If exactly 4 rows exist, just fill them
        logger.info("Exactly 4 rows exist, filling existing rows...")
        for no_urut in range(1, 5):
            logger.info(f"Filling row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1))

def save_form(driver, wait):
    """Save the form"""
    time.sleep(2)
    tablist = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/ul")))
    tablist.click()
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Saving Rencana Benda Uji...")
    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/div[2]/button[1]")))
    logger.info("Save button found, clicking...")
    save_button.click()
    time.sleep(3)

def create_form(wait):
    """Create new form"""
    logger.info("Creating new form...")
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/button[2]")))
    logger.info("Create button found, clicking...")
    create_button.click()

def duplicate_form(driver, wait, next_row_data):
    """Duplicate form for next entry with same kode_benda_uji and proyek"""
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Duplicating Rencana Benda Uji...")
    action_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/button")))
    action_button.click()
    time.sleep(1)
    duplicate_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/ul/li/a")))
    logger.info("Duplicate button found, clicking...")
    duplicate_button.click()
    time.sleep(2)
    # Update the duplicated form with next row data
    logger.info("Updating duplicated form with next row data...")
    fill_docket_form(driver, wait, next_row_data)
    logger.info("Input data to the table row using Excel data")
    add_table_rows(driver, wait, next_row_data)

def alternative_form(driver, wait, next_row_data):
    # Update the duplicated form with next row data
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Updating duplicated form with next row data...")
    fill_docket_form(driver, wait, next_row_data)
    logger.info("Input data to the table row using Excel data")
    add_table_rows(driver, wait, next_row_data)


def wait_for_loading_overlay_to_disappear(driver, wait, max_wait=30):
    """Wait for blockUI loading overlay to disappear"""
    try:
        loading_selectors = [
            "div.blockUI.blockMsg.blockPage",
            "div.blockUI.blockOverlay",
            "div.oe_blockui_spin_container"
        ]
        for selector in loading_selectors:
            try:
                wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, selector)))
            except:
                pass
        time.sleep(0.5)
        return True
    except Exception as e:
        logger.warning(f"Error waiting for loading overlay: {e}")
        return False

def is_click_intercepted_error(error_message):
    """Check if error is due to element click intercepted"""
    error_str = str(error_message).lower()
    return "element click intercepted" in error_str and ("blockui" in error_str or "blockoverlay" in error_str)

def refresh_and_wait(driver, wait):
    """Refresh page and wait for it to load"""
    logger.info("Refreshing page...")
    driver.refresh()
    time.sleep(3)
    wait_for_loading_overlay_to_disappear(driver, wait)
    time.sleep(2)

def process_excel_row_with_retry(driver, wait, excel_processor, row_data, row_index, max_retries=3):
    """Process single Excel row with retry logic for click intercepted errors"""
    no_docket = row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing Excel row {row_index + 1} - No. Docket: {no_docket}")

    for attempt in range(max_retries):
        try:
            navigate_and_create(driver, wait)

            wait_for_loading_overlay_to_disappear(driver, wait)
            fill_proyek_form(driver, wait, row_data)

            wait_for_loading_overlay_to_disappear(driver, wait)
            fill_docket_form(driver, wait, row_data)

            wait_for_loading_overlay_to_disappear(driver, wait)
            add_table_rows(driver, wait, row_data)

            save_form(driver, wait)
            logger.info(f"Success processing row {row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""
            
        except ElementClickInterceptedException as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {row_index + 1} (No. Docket: {no_docket})")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    continue
                else:
                    logger.error(f"Max retries reached for row {row_index + 1} (No. Docket: {no_docket}). Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                # Not a blockUI error, don't retry
                return False, no_docket, error_message
                
        except Exception as e:
            error_message = str(e)
            
            # Check if it's a click intercepted error even if not ElementClickInterceptedException
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted detected on attempt {attempt + 1}/{max_retries} for row {row_index + 1}")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    continue
                else:
                    logger.error(f"Max retries reached for row {row_index + 1}. Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                logger.error(f"Error processing row {row_index + 1}: {error_message}")
                return False, no_docket, error_message
            
    return False, no_docket, "Unknown error after all retries"

def process_duplicate_row_with_retry(driver, wait, next_row_data, next_row_index, max_retries=3):
    """Process next row using duplicate form with retry logic"""
    no_docket = next_row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing row {next_row_index + 1} using duplicate form - No. Docket: {no_docket}")
    
    # Flag untuk mengontrol dua opsi fungsi
    use_duplicate = True
    
    for attempt in range(max_retries):
        try:
            # DUA OPSI: duplicate_form atau alternative_form
            if use_duplicate:
                logger.info(f"Using duplicate_form on attempt {attempt + 1}")
                duplicate_form(driver, wait, next_row_data)
            else:
                logger.info(f"Using alternative_form on attempt {attempt + 1}")
                alternative_form(driver, wait, next_row_data)
            
            # Selalu panggil save_form setelah form processing
            save_form(driver, wait)
            logger.info(f"Success processing row {next_row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""
            
        except ElementClickInterceptedException as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {next_row_index + 1}")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    navigate_and_create(driver, wait)
                    fill_proyek_form(driver, wait, next_row_data)
                    
                    # SWITCH STRATEGI: Jika duplicate gagal, gunakan alternative
                    if use_duplicate:
                        use_duplicate = False
                        logger.info("Switching from duplicate_form to alternative_form strategy")
                    
                    continue  # Kembali ke awal loop dengan strategi baru
                else:
                    logger.error(f"Max retries reached for row {next_row_index + 1}. Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                return False, no_docket, error_message
                
        except Exception as e:
            error_message = str(e)
            
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted detected on attempt {attempt + 1}/{max_retries}")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    navigate_and_create(driver, wait)
                    fill_proyek_form(driver, wait, next_row_data)
                    
                    # SWITCH STRATEGI: Dari duplicate ke alternative
                    if use_duplicate:
                        use_duplicate = False
                        logger.info("Exception occurred, switching from duplicate_form to alternative_form strategy")
                    
                    continue  # Kembali ke loop dengan fungsi berbeda
                else:
                    logger.error(f"Max retries reached for row {next_row_index + 1}. Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                logger.error(f"Error processing row {next_row_index + 1}: {error_message}")
                return False, no_docket, error_message
    
    return False, no_docket, "Unknown error after all retries"

def prepare_for_next_row(driver, wait, excel_processor, row_index):
    """Prepare for next row - either duplicate or create new form"""
    try:
        wait_for_loading_overlay_to_disappear(driver, wait)
        if excel_processor.should_duplicate(row_index):
            logger.info("Next row has same kode_benda_uji and proyek - will use duplicate")
            return "duplicate"
        else:
            logger.info("Next row has different kode_benda_uji or proyek - creating new form")
            create_form(wait)
            return "create"
        
    except Exception as e:
        logger.error(f"Error preparing for next row after {row_index + 1}: {e}")
        return "error"

def log_processing_summary(successful_rows, failed_rows, skipped_rows, last_success_info, last_failure_info):
    """Log processing summary with last success/failure details"""
    logger.info(f"{'='*60}")
    logger_debug("="*60)
    logger.info("PROCESSING SUMMARY")
    logger_debug("PROCESSING SUMMARY")
    logger.info(f"{'='*60}")
    logger_debug("="*60)
    logger.info(f"Total successful rows: {len(successful_rows)}")
    logger_debug(f"Total successful rows: {len(successful_rows)}")
    logger.info(f"Total failed rows: {len(failed_rows)}")
    logger_debug(f"Total failed rows: {len(failed_rows)}")
    logger.info(f"Total skipped rows (after retries): {len(skipped_rows)}")
    logger_debug(f"Total skipped rows (after retries): {len(skipped_rows)}")
    logger.info(f"\n{'='*120}")
    logger_debug("="*120)
    
    if successful_rows:
        logger.info(f"\nSuccessful rows:")
        for row_info in successful_rows[-5:]:  # Show last 5 successful
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']}")
        if len(successful_rows) > 5:
            logger.info(f"  ... and {len(successful_rows) - 5} more")
    
    if failed_rows:
        logger.info(f"\nFailed rows:")
        for row_info in failed_rows:
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']} - {row_info['error']}")
    
    if skipped_rows:
        logger.info(f"\nSkipped rows (after max retries):")
        for row_info in skipped_rows:
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']} - {row_info['error']}")
    
    if last_success_info:
        logger.info(f"\nLast successful row: {last_success_info['index']} - No. Docket {last_success_info['no_docket']}")
    
    if last_failure_info:
        logger.info(f"Last failed row: {last_failure_info['index']} - No. Docket {last_failure_info['no_docket']}")
    
    logger.info(f"{'='*60}")

def initialize_components(excel_file_path):
    """Initialize Excel processor and web driver"""
    try:
        if not os.path.exists(excel_file_path):
            logger.error(f"Excel file not found: {excel_file_path}")
            return None, None
        
        excel_processor = ExcelDataProcessor(excel_file_path)
        driver = setup_driver()
        
        return driver, excel_processor
    except Exception as e:
        logger.error(f"Failed to initialize components: {e}")
        return None, None


def process_all_rows(driver, wait, excel_processor):
    """Process all rows from Excel with proper tracking"""
    results = {
        'successful_rows': [],
        'failed_rows': [],
        'skipped_rows': [],
        'last_success_info': None,
        'last_failure_info': None
    }
    
    total_rows = len(excel_processor.data)
    logger.info(f"Starting to process {total_rows} rows from Excel")
    
    row_index = 0
    while row_index < total_rows:
        row_data = excel_processor.get_row_data(row_index)
        
        if row_data is None:
            logger.warning(f"Skipping empty row {row_index + 1}")
            row_index += 1
            continue
        
        # Process current row
        no_docket = row_data.get('No. Docket', 'Unknown')
        log_row_header(row_index + 1, total_rows, no_docket)
        
        success, processed_no_docket, error_message = process_excel_row_with_retry(
            driver, wait, excel_processor, row_data, row_index
        )
        
        if success:
            # Handle successful row processing
            row_index = handle_successful_row(
                driver, wait, excel_processor, results, 
                row_index, total_rows, processed_no_docket
            )
        else:
            # Handle failed row processing
            handle_failed_row(
                driver, wait, results, row_index, 
                processed_no_docket, error_message
            )
        
        row_index += 1
        time.sleep(2)  # Consider making this configurable
    
    return results


def handle_successful_row(driver, wait, excel_processor, results, 
                         row_index, total_rows, processed_no_docket):
    """Handle successful row processing and potential duplicates"""
    success_info = create_row_info(row_index + 1, processed_no_docket)
    results['successful_rows'].append(success_info)
    results['last_success_info'] = success_info
    
    logger.info(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")
    logger_debug(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")
    
    # Handle next row preparation and potential duplicates
    if row_index + 1 < total_rows:
        return handle_next_row_preparation(
            driver, wait, excel_processor, results, 
            row_index, total_rows
        )
    else:
        logger.info("This is the last row - no next action needed")
    
    return row_index


def handle_next_row_preparation(driver, wait, excel_processor, results, 
                               row_index, total_rows):
    """Handle preparation for next row and potential duplicate processing"""
    next_action = prepare_for_next_row(driver, wait, excel_processor, row_index)
    
    if next_action == "duplicate":
        return process_duplicate_sequence(
            driver, wait, excel_processor, results, 
            row_index, total_rows
        )
    elif next_action == "error":
        logger.error(f"Error preparing for next row after {row_index + 1}")
    
    return row_index


def process_duplicate_sequence(driver, wait, excel_processor, results, 
                              row_index, total_rows):
    """Process sequence of duplicate rows"""
    current_row = row_index
    
    while current_row + 1 < total_rows:
        current_row += 1
        next_row_data = excel_processor.get_row_data(current_row)
        
        if next_row_data is None:
            logger.warning(f"Next row {current_row + 1} is empty - skipping duplicate")
            break
        
        next_no_docket = next_row_data.get('No. Docket', 'Unknown')
        log_duplicate_header(current_row + 1, total_rows, next_no_docket)
        
        duplicate_success, duplicate_no_docket, duplicate_error = process_duplicate_row_with_retry(
            driver, wait, next_row_data, current_row
        )
        
        if duplicate_success:
            handle_successful_duplicate(results, current_row, duplicate_no_docket)
            
            # Check if there's another duplicate
            if current_row + 1 < total_rows:
                next_action = prepare_for_next_row(driver, wait, excel_processor, current_row)
                if next_action != "duplicate":
                    break
            else:
                break
        else:
            handle_failed_duplicate(driver, wait, results, current_row, 
                                  duplicate_no_docket, duplicate_error)
            break
    
    return current_row


def handle_successful_duplicate(results, row_index, no_docket):
    """Handle successful duplicate processing"""
    success_info = create_row_info(row_index + 1, no_docket)
    results['successful_rows'].append(success_info)
    results['last_success_info'] = success_info
    
    logger.info(f"Row {row_index + 1} successfully processed via duplicate - No. Docket: {no_docket}")
    logger_debug(f"Row {row_index + 1} successfully processed via duplicate - No. Docket: {no_docket}")


def handle_failed_duplicate(driver, wait, results, row_index, no_docket, error_message):
    """Handle failed duplicate processing"""
    if is_max_retry_error(error_message):
        skipped_info = create_error_info(row_index + 1, no_docket, error_message)
        results['skipped_rows'].append(skipped_info)
        logger.warning(f"Row {row_index + 1} skipped after max retries - No. Docket: {no_docket}")
    else:
        failure_info = create_error_info(row_index + 1, no_docket, error_message)
        results['failed_rows'].append(failure_info)
        results['last_failure_info'] = failure_info
    
    log_failed_duplicate(row_index + 1, no_docket, error_message)
    refresh_and_wait(driver, wait)


def handle_failed_row(driver, wait, results, row_index, no_docket, error_message):
    """Handle failed row processing"""
    if is_max_retry_error(error_message):
        skipped_info = create_error_info(row_index + 1, no_docket, error_message)
        results['skipped_rows'].append(skipped_info)
        logger.warning(f"Row {row_index + 1} skipped after max retries - No. Docket: {no_docket}")
    else:
        failure_info = create_error_info(row_index + 1, no_docket, error_message)
        results['failed_rows'].append(failure_info)
        results['last_failure_info'] = failure_info
    
    log_failed_row(row_index + 1, no_docket, error_message)
    refresh_and_wait(driver, wait)


def cleanup_resources(driver):
    """Clean up resources properly"""
    if driver:
        logger.info("Browser will remain open for review...")
        input("Press Enter to close the browser...")
        driver.quit()


# Helper functions for better code organization
def create_row_info(index, no_docket):
    """Create row info dictionary"""
    return {'index': index, 'no_docket': no_docket}


def create_error_info(index, no_docket, error):
    """Create error info dictionary"""
    return {'index': index, 'no_docket': no_docket, 'error': error}


def is_max_retry_error(error_message):
    """Check if error is due to max retry attempts"""
    return "after" in error_message and "attempts" in error_message


def log_row_header(row_num, total_rows, no_docket):
    """Log row processing header"""
    logger.info(f"{'='*100}")
    logger.info(f"Processing Row {row_num}/{total_rows} - No. Docket: {no_docket}")
    logger.info(f"{'='*100}")
    logger_debug(f"{'='*100}")


def log_duplicate_header(row_num, total_rows, no_docket):
    """Log duplicate row processing header"""
    logger.info(f"\n{'='*50}")
    logger.info(f"Processing Row {row_num}/{total_rows} (via duplicate) - No. Docket: {no_docket}")
    logger.info(f"{'='*50}")


def log_failed_row(row_num, no_docket, error_message):
    """Log failed row processing"""
    logger.error(f"Row {row_num} processing failed - No. Docket: {no_docket} - Error: {error_message}")
    logger.error(f"Skipped processing row {row_num} - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")
    logger_debug(f"Row {row_num} Skipped - Processing failed - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")


def log_failed_duplicate(row_num, no_docket, error_message):
    """Log failed duplicate processing"""
    logger.error(f"Row {row_num} duplicate processing failed - No. Docket: {no_docket} - Error: {error_message}")
    logger.error(f"Skipped processing row {row_num} - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")
    logger_debug(f"Row {row_num} Skipped - Duplicate Processing failed - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")


# Configuration class for better maintainability
class ProcessingConfig:
    """Configuration class for processing parameters"""
    EXCEL_FILE_PATH = "data.xlsx"
    WAIT_TIMEOUT = 10
    PROCESSING_DELAY = 2
    ERROR_MESSAGE_DUPLICATE = "Tidak ada No. Docket / Sudah Pernah diinput"

def main():
    """Main execution function with enhanced retry logic and improved structure"""
    driver = None
    excel_processor = None
    
    try:
        # Initialize components
        excel_file_path = "data.xlsx"  # Consider moving to config file
        driver, excel_processor = initialize_components(excel_file_path)
        if not driver or not excel_processor:
            return
        
        wait = WebDriverWait(driver, 10)
        login(driver, wait)
        
        # Process all rows
        results = process_all_rows(driver, wait, excel_processor)
        
        # Log final summary
        log_processing_summary(
            results['successful_rows'], 
            results['failed_rows'], 
            results['skipped_rows'],
            results['last_success_info'], 
            results['last_failure_info']
        )

    except Exception as e:
        logger.error(f"Unexpected error in main: {e}")
    finally:
        cleanup_resources(driver)

if __name__ == "__main__":
    main()