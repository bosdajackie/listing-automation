import os, re, time, xlsxwriter
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

lineBreak = "\n" + "=" * 100 + "\n"
smallLineBreak = "\n" + "-" * 100 + "\n"
print(f"{lineBreak}\nGet the vehicle compatibilities and specifications for your SKU\n{lineBreak}")

# set up chrome driver
ua = UserAgent()
options = uc.ChromeOptions()
options.add_argument(f'user-agent={ua.random}')  # disguise software agent identity
options.add_argument("--disable-blink-features=AutomationControlled")  # disables flag that marks browser as automated
# options.add_argument("--headless")  # hide browser GUI to make invisible to user
driver = uc.Chrome(options=options, service=Service(ChromeDriverManager().install()))

# file/folder paths
currentPath = os.getcwd()
resultsFolder = os.path.join(currentPath, "results")
compatibilityExcelPath = os.path.join(resultsFolder, "compatibility.xlsx")
extraInfoTxtPath = os.path.join(resultsFolder, "extraInfo.txt")

# set up excel files to write to
compatibilityWorkbook = xlsxwriter.Workbook(compatibilityExcelPath)
compatibilityWorksheet = compatibilityWorkbook.add_worksheet()
compatibilityWorksheet.set_default_row(20.25)
compatibilityWorksheet.set_column('A:Z', 17)
headerFormat = compatibilityWorkbook.add_format({
    "bold": True,
    "text_wrap": True,
    "align": "center",
    "valign": "vcenter",
    "bg_color": "#2F75B5",
    "font_color": "white",
    "border": 1
})
cellFormat = compatibilityWorkbook.add_format({
    "bold": True,
    "text_wrap": True,
    "align": "center",
    "valign": "vcenter",
    "border": 1
})

# set up text file to write extra info to
txtFile = open(extraInfoTxtPath, 'w', encoding='utf-8')

# ask user if they are listing to Autofirst, Karshield, or 365Hubs
while True:
    storefront = input(
        "Select from the following storefronts (1, 2, or 3):\n"
        "(1) AutoFirst\n(2) Karshield\n(3) 365Hubs\n"
    )
    if storefront.lower() in ['1', '2', '3', 'autofirst', 'karshield', '365hubs']:
        if storefront.lower() in ['2', 'karshield']:
            headerFormat.set_bg_color("red")
        print(f"{lineBreak}")
        break
    print("Invalid input. Please enter 1, 2, or 3.\n")

# write headers with formatting
compatibilityWorksheet.write("A1", "Make", headerFormat)
compatibilityWorksheet.write("B1", "Model", headerFormat)
compatibilityWorksheet.write("C1", "Year", headerFormat)
compatibilityWorksheet.write("D1", "Position", headerFormat)
compatibilityWorksheet.write("E1", "Engine", headerFormat)

# ask for SKU and search it on rockauto
sku = input("Enter the SKU you would like to search for:\n").strip()
print(f"{lineBreak}Loading website for SKU {sku}...\n")
txtFile.write(f"{lineBreak}\nExtra information for SKU {sku}\n{lineBreak}")
website = "https://www.rockauto.com/en/partsearch/?partnum=" + sku
driver.get(website)

# wait for page listings and fetch all listings from container
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'listings-container')))
    allResults = driver.find_elements(By.XPATH, '//*[contains(@class, "listing-border-top-line listing-inner-content")]')
    print(f"Fetching all results for SKU{smallLineBreak}")
except TimeoutException:
    print("Timeout waiting for listings container for SKU (no results found)")
    driver.quit()
    compatibilityWorkbook.close()
    txtFile.close()
    exit()

for i, result in enumerate(allResults):
    try:
        partNumber = result.find_element(By.CLASS_NAME, 'listing-final-partnumber').text
        manufacturer = result.find_element(By.CLASS_NAME, 'listing-final-manufacturer').text
        categoryRaw = result.find_element(By.CLASS_NAME, 'listing-text-row').text
        category = re.split(r'\s[\(\[].*$', categoryRaw[10:])[0].strip()

        print(f"({i+1}) {partNumber}\nManufacturer: {manufacturer}\nCategory: {category}\n{smallLineBreak}")
    except NoSuchElementException:
        print("No such element error")
print(f"{lineBreak}")

# select the index to get specifications information
while True:
    indexInput = input(f"Select the index (1-{len(allResults)}) of the listing you'd like to get part specs info from.\nIf you want to skip this process enter 0.\n")
    if indexInput.isdigit():
        index = int(indexInput)
        if index == 0:
            print("Skipped part specification grabbing process")
            if createSpecificationsExcel(None, driver): break
        if 0 < index <= len(allResults):
            chosenResult = allResults[index-1]
            infoHref = chosenResult.find_element(By.CLASS_NAME, 'ra-btn-moreinfo').get_attribute('href')
            if createSpecificationsExcel(infoHref, driver):
                break
    else: print(f"Invalid input. Please enter an index/number from 0 to {len(allResults)}")
print(f"{lineBreak}")

driver.get(website)  # go back to partnumber site
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'listings-container')))
allResults = driver.find_elements(By.XPATH, '//*[contains(@class, "listing-border-top-line listing-inner-content")]')  # re-fetch elements to avoid stale references

# select the index to get manufacturer and category information
while True:
    indexInput = input(f"Select the index (1-{len(allResults)}) of the listing you'd like to get manufacturer/category from:\n")
    if indexInput.isdigit():
        index = int(indexInput)
        if 0 < index <= len(allResults):
            chosenResult = allResults[index-1]
            chosenPartNumber = chosenResult.find_element(By.CLASS_NAME, 'listing-final-partnumber').text
            chosenManufacturer = chosenResult.find_element(By.CLASS_NAME, 'listing-final-manufacturer').text
            categoryRaw = chosenResult.find_element(By.CLASS_NAME, 'listing-text-row').text
            chosenCategory = re.split(r'\s[\(\[].*$', categoryRaw[10:])[0].strip()

            print(f"\nSelected part #: {chosenPartNumber}\nSelected manufacturer: {chosenManufacturer}\nSelected category: {chosenCategory}")
            break
    else: print(f"Invalid input. Please enter an index/number from 1 to {len(allResults)}")
print(f"{lineBreak}")

# open popup and get compatible vehicles
partLink = chosenResult.find_element(By.XPATH, './/*[contains(@id, "vew_partnumber")]')
partLink.click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="buyersguidepopup-outer_b"]/div/div/table')))
compatibleVehicles = driver.find_elements(By.XPATH, '//*[@id="buyersguidepopup-outer_b"]/div/div/table/tbody/tr')

# extract vehicle information from popup
make, model, startYear, endYear, position, extra = [], [], [], [], [], []
try:
    for vehicle in compatibleVehicles:
        try:
            carMake = vehicle.find_element(By.XPATH, './td[1]').text
            carModel = vehicle.find_element(By.XPATH, './td[2]').text
            carYears = vehicle.find_element(By.XPATH, './td[3]').text
            if "-" in carYears:
                startYearValue, endYearValue = carYears.split("-")
            else:
                startYearValue = endYearValue = carYears
            make.append(carMake)
            model.append(carModel)
            startYear.append(startYearValue)
            endYear.append(endYearValue)
            position.append("")  # placeholder
            extra.append("")  # placeholder
        except NoSuchElementException:
            print("Error extracting vehicle info")
            continue
except Exception as e:
    print(f"Error processing vehicle list: {e}")
    
try:
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'dialog-close'))).click()  # close dialog box
except TimeoutException:
    print("Timeout closing dialog")

# Safely click an element with multiple fallback strategies
def safeClick(driver, element):
    try:
        # Try normal click first
        element.click()
        return True
    except ElementClickInterceptedException:
        try:
            # Try JavaScript click
            driver.execute_script("arguments[0].click();", element)
            return True
        except:
            try:
                # Try ActionChains
                ActionChains(driver).move_to_element(element).click().perform()
                return True
            except:
                return False

# Navigate to part category with retry logic
def navigateToCategory(driver, category, max_retries=3):
    for attempt in range(max_retries):
        try:
            # Wait for and click "Brake & Wheel Hub"
            brakeHubLink = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Brake & Wheel Hub')]"))
            )
            if not safeClick(driver, brakeHubLink):
                print(f"Failed to click 'Brake & Wheel Hub' on attempt {attempt + 1}")
                continue
            
            time.sleep(1)
            
            # Wait for and click the specific category
            categoryLink = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space(text()) = '{category}']"))
            )
            if not safeClick(driver, categoryLink):
                print(f"Failed to click '{category}' on attempt {attempt + 1}")
                continue
            
            return True
        except TimeoutException:
            print(f"Timeout navigating to category on attempt {attempt + 1}")
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
    return False

# process each vehicle
for i in range(len(make)):
    try:
        searchString = endYear[i] + " " + make[i] + " " + model[i] + " "
        print(f"{smallLineBreak}\nProcessing results for: {searchString}")
        txtFile.write(f"{smallLineBreak}\n{searchString}\n")

        # navigate to catalog
        driver.get("https://www.rockauto.com/en/catalog/")
        
        # Wait for page to load and find search input
        try:
            searchBar = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="topsearchinput[input]"]'))
            )
            searchBar.clear()
            searchBar.send_keys(searchString)
            time.sleep(0.5)
        except TimeoutException:
            print(f"Timeout loading catalog or finding search input for {searchString}")
            continue

        # get autosuggested engines
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'))
            )
            engines = driver.find_elements(By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr')
        except TimeoutException:
            print(f"Timeout waiting for autosuggestions for {searchString}")
            continue

        # process each engine
        for j in range(1, len(engines)):  # skip index 0
            engineDisplacement = "Unknown"  # Default value
            
            try:
                # Navigate back to catalog for each engine
                driver.get("https://www.rockauto.com/en/catalog/")
                
                # Re-enter search string
                searchBar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@id="topsearchinput[input]"]'))
                )
                searchBar.clear()
                searchBar.send_keys(searchString)
                time.sleep(0.5)
                
                # Wait for suggestions and get fresh elements
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr'))
                )
                engines = driver.find_elements(By.XPATH, '//*[@id="autosuggestions[topsearchinput]"]/tbody/tr')
                
                # Get engine text before clicking (for better debugging)
                if j < len(engines):
                    engineText = engines[j].text.strip()
                    print(f"\nChecking {engineText}...")
                    
                    # Click the engine
                    engine = engines[j]
                    if not safeClick(driver, engine):
                        print(f"Failed to click engine {j} for {searchString}")
                        continue
                else:
                    print(f"Engine index {j} out of range for {searchString}")
                    continue

            except Exception as e:
                print(f"Error clicking engine {j}: {e}")
                continue

            # get engine displacement (L) with better error handling
            try:
                # breadcrumbsElement = WebDriverWait(driver, 10).until(
                #     EC.presence_of_element_located((By.XPATH, '//*[@id="breadcrumb_location_banner_inner[catalog]"]'))
                # )
                # breadcrumbs = breadcrumbsElement.text.split(">")
                # if len(breadcrumbs) > 0:
                #     engineDisplacement = breadcrumbs[-1].strip()
                #     print(f"Engine displacement found: {engineDisplacement}")
                # else:
                #     print("No breadcrumbs found")
                #     engineDisplacement = engineText if 'engineText' in locals() else "Unknown"

                crumb = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        # any id that starts with breadcrumb_location_banner_inner …
                        (By.CSS_SELECTOR,
                        "div[id^='breadcrumb_location_banner_inner'] span.belem.active")
                    )
                )
                engineDisplacement = crumb.text.strip()
                print(f"Engine displacement found: {engineDisplacement}")
            except TimeoutException:
                print("Timeout waiting for breadcrumb")
                engineDisplacement = engineText if 'engineText' in locals() else "Unknown"
            except Exception as e:
                print(f"Error getting engine displacement: {e}")
                engineDisplacement = engineText if 'engineText' in locals() else "Unknown"

            # navigate to category with improved error handling
            if not navigateToCategory(driver, chosenCategory):
                print(f"Failed to navigate to category {chosenCategory} for {searchString}")
                continue

            # filter by partnumber
            try:
                searchBar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'filter-input'))
                )
                searchBar.clear()
                searchBar.send_keys(chosenPartNumber)
                searchBar.send_keys(Keys.ENTER)
                time.sleep(1)
            except TimeoutException:
                print(f"Timeout finding search input for {chosenPartNumber}")
                continue

            # extract part listing
            try:
                partListing = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, 
                        f"//td[contains(@class, 'listing-inner-content')][.//span[contains(@class, 'listing-final-manufacturer') and contains(text(), '{chosenManufacturer}')]]"))
                )
                partFits = True
                print(f"Part found for {engineDisplacement}")
            except TimeoutException:
                print(f"Part listing not found for engine {engineDisplacement}")
                partListing = None
                partFits = False

            # extract position and extra info
            partInfo = ""
            if partListing:
                try:
                    footnotes = [
                        e.text.strip()
                        for e in partListing.find_elements(By.CLASS_NAME, "listing-footnote-text")
                        if e.text.strip()
                    ]
                    uniqueNotes = list(dict.fromkeys(footnotes)) # remove duplicates while preserving order
                    partInfo = " or ".join(uniqueNotes)
                except NoSuchElementException:
                    print("No footnote-text found in part listing")
                    
            if partInfo:
                if ";" in partInfo:
                    parts = partInfo.split("; ")
                    currentPosition = parts[0]
                    extraInfo = "; ".join(parts[1:]).strip() if len(parts) > 1 else ""
                else:
                    currentPosition = partInfo
                    extraInfo = ""
                    
                if not position[i]:
                    position[i] = currentPosition
                    
                if extraInfo:
                    if extra[i]:
                        extra[i] = extra[i] + "; " + extraInfo
                    else:
                        extra[i] = extraInfo
                        
                print(f"Found position: {position[i]}")
                print(f"Found extra information: {extra[i]}")
                
            # Handle case where part doesn't fit
            if not partFits:
                noFitText = f"No {engineDisplacement}"
                print(f"Adding to extra: {noFitText}")
                
                if not extra[i]:
                    extra[i] = noFitText
                else:
                    extra[i] = extra[i] + ", " + noFitText

            # write results for engine to text file
            txtFile.write(f"Results for engine {j} ({engineDisplacement}): {partInfo if partInfo else 'No fit'}\n")

        # write final row to excel (each vehicle once)
        compatibilityWorksheet.write(i + 1, 0, make[i], cellFormat)
        compatibilityWorksheet.write(i + 1, 1, model[i], cellFormat)
        compatibilityWorksheet.write(i + 1, 2, startYear[i] + "-" + endYear[i], cellFormat)
        compatibilityWorksheet.write(i + 1, 3, position[i], cellFormat)
        compatibilityWorksheet.write(i + 1, 4, extra[i], cellFormat)
        
        print(f"Final extra info for {make[i]} {model[i]}: {extra[i]}")
        
    except Exception as e:
        print(f"Error processing {searchString}: {e}")
        continue

print(f"{smallLineBreak}\n{lineBreak}\nFinished getting all results for your SKU! Please check results folder\n{lineBreak}")

driver.quit()
compatibilityWorkbook.close()
txtFile.close()