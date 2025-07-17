import os
import re
import time
import xlsxwriter
from fake_useragent import UserAgent
import undetected_chromedriver as uc
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from specifications import createSpecificationsExcel

class WebScraper:
    def __init__(self, storefront="Karshield", headless=False, status_callback=None):
        self.storefront = storefront
        self.headless = headless
        self.status_callback = status_callback
        self.driver = None
        self.product_results = []
        self.selected_product = None
        
        # File paths
        self.current_path = os.getcwd()
        self.results_folder = os.path.join(self.current_path, "results")
        self.compatibility_excel_path = os.path.join(self.results_folder, "compatibility.xlsx")
        self.extra_info_txt_path = os.path.join(self.results_folder, "extraInfo.txt")
        
        # Ensure results folder exists
        os.makedirs(self.results_folder, exist_ok=True)
        
        self.init_driver()

    # Initialize the Chrome driver
    def init_driver(self):
        self.update_status("Initializing browser...")
        
        ua = UserAgent()
        options = uc.ChromeOptions()
        options.add_argument(f'user-agent={ua.random}')
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        if self.headless:
            options.add_argument("--headless")
            
        self.driver = uc.Chrome(options=options, service=Service(ChromeDriverManager().install()))

    # Update status if callback is provided
    def update_status(self, message):
        if self.status_callback:
            self.status_callback(message)

    # Search for products by SKU and return results
    def search_products(self, sku):
        self.update_status(f"Searching for SKU: {sku}")
        
        website = f"https://www.rockauto.com/en/partsearch/?partnum={sku}"
        self.driver.get(website)
        
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'listings-container'))
            )
            all_results = self.driver.find_elements(
                By.XPATH, '//*[contains(@class, "listing-border-top-line listing-inner-content")]'
            )
            
            if not all_results:
                raise TimeoutException("No results found")
                
        except TimeoutException:
            raise Exception(f"No results found for SKU: {sku}")
        
        self.product_results = []
        for result in all_results:
            try:
                self.product_results.append(
                    dict(part_number = result.find_element(By.CLASS_NAME, "listing-final-partnumber").text,
                         manufacturer = result.find_element(By.CLASS_NAME, "listing-final-manufacturer").text,
                         category = re.split(r"\s[\(\[].*$", result.find_element(By.CLASS_NAME, "listing-text-row").text[10:])[0].strip(), element=result,
                    )
                )
            except NoSuchElementException:
                continue
        
        return self.product_results

    # Get specifications for the selected product
    def get_specifications(self, product_index):
        if product_index >= len(self.product_results):
            raise Exception("Invalid product index for specifications")
        
        self.update_status("Getting product specifications...")
        
        try:
            product = self.product_results[product_index]
            info_href = product['element'].find_element(By.CLASS_NAME, 'ra-btn-moreinfo').get_attribute('href')
            createSpecificationsExcel(info_href, self.driver)
            
        except Exception as e:
            # If specifications fail, continue with compatibility
            self.update_status(f"Specifications failed: {str(e)}")

    # Get compatibility information for the selected product
    def get_compatibility(self, product_index):
        if product_index >= len(self.product_results):
            raise Exception("Invalid product index for compatibility")
        
        self.selected_product = self.product_results[product_index]
        
        # Go back to original search, reload listing page (avoid stale elements)
        sku = self.selected_product['part_number']
        website = f"https://www.rockauto.com/en/partsearch/?partnum={sku}"
        self.driver.get(website)
        
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'listings-container'))
        )
        
        # Re-fetch elements to avoid stale references
        all_results = self.driver.find_elements(
            By.XPATH, '//*[contains(@class, "listing-border-top-line listing-inner-content")]'
        )
        
        if product_index >= len(all_results):
            raise Exception("Product no longer available")
        
        chosen_result = all_results[product_index]
        
        # Get product details
        chosen_part_number = chosen_result.find_element(By.CLASS_NAME, 'listing-final-partnumber').text
        chosen_manufacturer = chosen_result.find_element(By.CLASS_NAME, 'listing-final-manufacturer').text
        category_raw = chosen_result.find_element(By.CLASS_NAME, 'listing-text-row').text
        chosen_category = re.split(r'\s[\(\[].*$', category_raw[10:])[0].strip()
        
        # Open compatibility popup
        self.update_status("Getting vehicle compatibility...")
        part_link = chosen_result.find_element(By.XPATH, './/*[contains(@id, "vew_partnumber")]')
        part_link.click()
        
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="buyersguidepopup-outer_b"]/div/div/table'))
        )
        
        compatible_vehicles = self.driver.find_elements(
            By.XPATH, '//*[@id="buyersguidepopup-outer_b"]/div/div/table/tbody/tr'
        )
        
        # Extract vehicle information
        vehicles = []
        for vehicle in compatible_vehicles:
            try:
                make = vehicle.find_element(By.XPATH, './td[1]').text
                model = vehicle.find_element(By.XPATH, './td[2]').text
                years = vehicle.find_element(By.XPATH, './td[3]').text
                
                if "-" in years:
                    start_year, end_year = years.split("-")
                else:
                    start_year = end_year = years
                
                vehicles.append({
                    'make': make,
                    'model': model,
                    'start_year': start_year,
                    'end_year': end_year,
                    'position': "",
                    'extra': ""
                })
                
            except NoSuchElementException:
                continue
        
        # Close dialog/popup
        try:
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'dialog-close'))
            ).click()
        except TimeoutException:
            pass
        
        # Process each vehicle for detailed compatibility
        self.update_status("Processing vehicle compatibility...")
        
        results_text = f"Compatibility Results for {chosen_part_number}\n"
        results_text += f"Manufacturer: {chosen_manufacturer}\n"
        results_text += f"Category: {chosen_category}\n"
        results_text += "=" * 80 + "\n\n"
        
        # Setup Excel file
        self.setup_excel_file()
        
        for i, vehicle in enumerate(vehicles):
            self.update_status(f"Processing vehicle {i+1}/{len(vehicles)}")
            
            try:
                vehicle_info = self.process_vehicle_compatibility(
                    vehicle, chosen_part_number, chosen_manufacturer, chosen_category
                )
                
                # Write to Excel
                self.write_vehicle_to_excel(i, vehicle_info)
                
                # Add to results text
                year_str = vehicle_info['start_year'] if vehicle_info['start_year'] == vehicle_info['end_year'] else f"{vehicle_info['start_year']}-{vehicle_info['end_year']}"
                results_text += f"{vehicle_info['make']} {vehicle_info['model']} ({year_str})\n"
                results_text += f"Position: {vehicle_info['position']}\n"
                results_text += f"Engine Info: {vehicle_info['extra']}\n"
                results_text += "-" * 50 + "\n"
                
            except Exception as e:
                results_text += f"Error processing {vehicle['make']} {vehicle['model']}: {str(e)}\n"
                results_text += "-" * 50 + "\n"
        
        # Close Excel file
        self.close_excel_file()
        
        results_text += f"\nResults saved to: {self.compatibility_excel_path}\n"
        
        return results_text

    # Process compatibility for a single vehicle
    def process_vehicle_compatibility(self, vehicle, part_number, manufacturer, category):
        search_string = f"{vehicle['end_year']} {vehicle['make']} {vehicle['model']} "
        
        # Navigate to catalog
        self.driver.get("https://www.rockauto.com/en/catalog/")
        
        # Search for vehicle
        try:
            search_bar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="topsearchinput[input]"]'))
            )
            search_bar.clear()
            search_bar.send_keys(search_string)
            time.sleep(0.5)
        except TimeoutException:
            raise Exception(f"Timeout loading catalog for {search_string}")
        
        # Get engine suggestions
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'))
            )
            engines = self.driver.find_elements(
                By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'
            )
        except TimeoutException:
            raise Exception(f"No engine suggestions found for {search_string}")
        
        vehicle_info = vehicle.copy()
        
        # Process each engine (skip index 0 ‑‑ header)
        for j in range(1, len(engines)):
            engine_displacement, part_info, part_fits = self.process_engine_compatibility(
                search_string, j, part_number, manufacturer, category
            )

            # write one line to txt file per engine
            self.txt_file.write(
                f"Results for engine {j} ({engine_displacement}): "
                f"{part_info if part_info else 'No fit'}\n"
            )
            
            # update vehicle dict (position / extra)
            if part_fits and part_info:
                # split footnotes
                if ";" in part_info:
                    pos = part_info.split("; ")[0]
                    extra_info = "; ".join(part_info.split("; ")[1:])
                else:
                    pos = part_info
                    extra_info = ""

                if pos and not vehicle_info["position"]:
                    vehicle_info["position"] = pos

                if extra_info:
                    vehicle_info["extra"] = (
                        extra_info
                        if not vehicle_info["extra"]
                        else f'{vehicle_info["extra"]}; {extra_info}'
                    )

            elif not part_fits:
                nofit = f"No {engine_displacement}"
                vehicle_info["extra"] = (
                    nofit if not vehicle_info["extra"] else f'{vehicle_info["extra"]}, {nofit}'
                )

            # always append engine displacement text (so GUI shows them)
            # if engine_displacement not in vehicle_info["extra"]:
            #     vehicle["extra"] = (
            #         engine_displacement
            #         if not vehicle_info["extra"]
            #         else f'{vehicle_info["extra"]}, {engine_displacement}'
            #     )

        return vehicle_info

    # Process compatibility for a specific engine
    def process_engine_compatibility(self, search_string, engine_index, part_number, manufacturer, category):
        # Navigate back to catalog
        self.driver.get("https://www.rockauto.com/en/catalog/")
        
        # Re-enter search
        search_bar = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="topsearchinput[input]"]'))
        )
        search_bar.clear()
        search_bar.send_keys(search_string)
        time.sleep(0.5)
        
        # Get fresh engine elements
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'))
        )
        engines = self.driver.find_elements(
            By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'
        )
        
        if engine_index >= len(engines):
            return "Unknown", "", False
        
        # Click engine
        engine = engines[engine_index]
        engine_text = engine.text.strip()
        
        if not self.safe_click(engine):
            return engine_text, "", False
        
        # Get engine displacement
        try:
            crumb = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[id^='breadcrumb_location_banner_inner'] span.belem.active")
                )
            )
            engine_displacement = crumb.text.strip()
        except TimeoutException:
            engine_displacement = engine_text
        
        # Navigate to category
        if not self.navigate_to_category(category):
            return engine_displacement, "", False
        
        # Search for part
        try:
            search_bar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'filter-input'))
            )
            search_bar.clear()
            search_bar.send_keys(part_number)
            search_bar.send_keys(Keys.ENTER)
            time.sleep(1)
        except TimeoutException:
            return engine_displacement, "", False
        
        # Check if part fits
        try:
            part_listing = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, 
                    f"//td[contains(@class, 'listing-inner-content')][.//span[contains(@class, 'listing-final-manufacturer') and contains(text(), '{manufacturer}')]]"))
            )
            part_fits = True
        except TimeoutException:
            part_listing = None
            part_fits = False

        # footnote / position / extra
        part_info = ""
        if part_listing:
            try:
                notes = [e.text.strip()
                         for e in part_listing.find_elements(By.CLASS_NAME, "listing-footnote-text")
                         if e.text.strip()]
                part_info = " or ".join(dict.fromkeys(notes))  # de-dupe + preserve order
            except NoSuchElementException:
                pass

        return engine_displacement, part_info, part_fits

    # Safely click an element with multiple fallback strategies
    def safe_click(self, element):
        try:
            element.click()
            return True
        except ElementClickInterceptedException:
            try:
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except:
                try:
                    ActionChains(self.driver).move_to_element(element).click().perform()
                    return True
                except:
                    return False

    # Navigate to part category with retry logic
    def navigate_to_category(self, category, max_retries=3):
        for attempt in range(max_retries):
            try:
                # Click "Brake & Wheel Hub"
                brake_hub_link = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Brake & Wheel Hub')]"))
                )
                if not self.safe_click(brake_hub_link):
                    continue
                
                time.sleep(1)
                
                # Click the specific category
                category_link = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space(text()) = '{category}']"))
                )
                return self.safe_click(category_link)
                
            except TimeoutException:
                time.sleep(2)
        
        return False

    # Setup Excel file for writing compatibility results
    def setup_excel_file(self):
        self.compatibility_workbook = xlsxwriter.Workbook(self.compatibility_excel_path)
        self.compatibility_worksheet = self.compatibility_workbook.add_worksheet()
        self.compatibility_worksheet.set_default_row(20.25)
        self.compatibility_worksheet.set_column('A:Z', 17)
        
        # Setup formats
        header_color = "#2F75B5"
        if self.storefront.lower() == "karshield":
            header_color = "red"
        
        self.header_format = self.compatibility_workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": header_color,
            "font_color": "white",
            "border": 1
        })
        
        self.cell_format = self.compatibility_workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })
        
        # Write headers
        self.compatibility_worksheet.write("A1", "Make", self.header_format)
        self.compatibility_worksheet.write("B1", "Model", self.header_format)
        self.compatibility_worksheet.write("C1", "Year", self.header_format)
        self.compatibility_worksheet.write("D1", "Position", self.header_format)
        self.compatibility_worksheet.write("E1", "Engine", self.header_format)
        
        # Setup text file
        self.txt_file = open(self.extra_info_txt_path, 'w', encoding='utf-8')
        self.txt_file.write(f"Extra information for SKU {self.selected_product['part_number']}\n")
        self.txt_file.write("=" * 80 + "\n")

    # Write vehicle information to Excel file
    def write_vehicle_to_excel(self, row_index, vehicle_info):
        row = row_index + 1
        
        self.compatibility_worksheet.write(row, 0, vehicle_info['make'], self.cell_format)
        self.compatibility_worksheet.write(row, 1, vehicle_info['model'], self.cell_format)
        
        # Handle year range
        if vehicle_info['start_year'] == vehicle_info['end_year']:
            year_str = vehicle_info['start_year']
        else:
            year_str = f"{vehicle_info['start_year']}-{vehicle_info['end_year']}"
        
        self.compatibility_worksheet.write(row, 2, year_str, self.cell_format)
        self.compatibility_worksheet.write(row, 3, vehicle_info['position'], self.cell_format)
        self.compatibility_worksheet.write(row, 4, vehicle_info['extra'], self.cell_format)
        
        # Write to text file
        self.txt_file.write(f"{vehicle_info['make']} {vehicle_info['model']} ({year_str}): {vehicle_info['extra']}\n")
        self.txt_file.write("-" * 50 + "\n")

    # Close Excel file and text file
    def close_excel_file(self):
        if hasattr(self, 'compatibility_workbook'):
            self.compatibility_workbook.close()
        if hasattr(self, 'txt_file'):
            self.txt_file.close()

    # Close the webscraper and cleanup resources
    def close(self):
        if self.driver:
            self.driver.quit()
        
        if hasattr(self, 'compatibility_workbook'):
            try:
                self.compatibility_workbook.close()
            except:
                pass
        
        if hasattr(self, 'txt_file'):
            try:
                self.txt_file.close()
            except:
                pass