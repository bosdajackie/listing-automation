import os
import re
from collections import defaultdict
import xlsxwriter
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

def createSpecificationsExcel(href, driver):
    # file/folder paths
    currentPath = os.getcwd()
    specificationExcelPath = os.path.join(os.path.join(currentPath, "results"), "specifications.xlsx")

    # set up excel files to write to
    specificationWorkbook = xlsxwriter.Workbook(specificationExcelPath)
    specificationWorksheet = specificationWorkbook.add_worksheet()
    specificationWorksheet.set_default_row(20.25)
    specificationWorksheet.set_column('A:B', 40)
    labelFormat = specificationWorkbook.add_format({
        "bold": True,
        "text_wrap": True,
        "align": "left",
        "valign": "vcenter",
        "border": 1
    })
    valueFormat = specificationWorkbook.add_format({
        "bold": False,
        "text_wrap": True,
        "align": "left",
        "valign": "vcenter",
        "border": 1
    })

    IN_TO_MM = 25.4

    def round_to(x, decimals=3):
        return round(float(x), decimals)

    def in_to_mm(val_in):
        return round_to(float(val_in) * IN_TO_MM)

    def mm_to_in(val_mm):
        return round_to(float(val_mm) / IN_TO_MM)

    def parse_value_with_unit(value_str):
        """Parse a value string that might contain embedded units"""
        # Patterns to match numbers with units (with or without spaces)
        mm_pattern = r'(\d+(?:\.\d+)?)\s*mm'
        in_pattern = r'(\d+(?:\.\d+)?)\s*(?:in|inch|inches|")'
        
        # Check for mm units
        mm_match = re.search(mm_pattern, value_str, re.IGNORECASE)
        if mm_match:
            return float(mm_match.group(1)), 'mm'
        
        # Check for inch units
        in_match = re.search(in_pattern, value_str, re.IGNORECASE)
        if in_match:
            return float(in_match.group(1)), 'in'
        
        # If no units found, return the original string
        return value_str, 'raw'

    if not href:
        # clear old specifications file (if it exists)
        if os.path.exists(specificationExcelPath):
            os.remove(specificationExcelPath)

        # create empty workbook and close it right away
        specificationWorkbook = xlsxwriter.Workbook(specificationExcelPath)
        specificationWorkbook.close()
        print("No specification URL provided; empty file created.")
        return True
    else:
        driver.get(href)

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'moreinfotable')))
        print("\nSpecifications table found. Getting rows data...")
        table = driver.find_element(By.CLASS_NAME, 'moreinfotable')
        rows = table.find_elements(By.TAG_NAME, 'tr')

        specs = defaultdict(dict)  # { "label": {"in": val, "mm": val, "raw": val} }

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if not cells:
                continue

            raw_label = cells[0].text.strip()
            value = cells[1].text.strip()

            # First check if label contains unit info (original logic)
            if "(IN)" in raw_label.upper():
                label = raw_label.replace("(IN)", "").strip()
                specs[label]["in"] = value
            elif "(MM)" in raw_label.upper():
                label = raw_label.replace("(MM)", "").strip()
                specs[label]["mm"] = value
            else:
                # Check if value contains embedded units
                parsed_value, unit_type = parse_value_with_unit(value)
                
                if unit_type == 'mm':
                    specs[raw_label]["mm"] = str(parsed_value)
                elif unit_type == 'in':
                    specs[raw_label]["in"] = str(parsed_value)
                else:
                    specs[raw_label]["raw"] = value

        # convert and write to Excel
        row_index = 0
        for label, values in specs.items():
            # Column A (label)
            specificationWorksheet.write(row_index, 0, label, labelFormat)

            # Column B (value)
            if "raw" in values:
                raw = values["raw"]
                try:
                    number_value = float(raw)
                    if number_value.is_integer():
                        number_value = int(number_value)
                    specificationWorksheet.write_number(row_index, 1, number_value, valueFormat)
                except ValueError:
                    specificationWorksheet.write(row_index, 1, raw, valueFormat)

            elif "in" in values and "mm" in values:
                spec_value = f"{values['in']} in / {values['mm']} mm"
                specificationWorksheet.write(row_index, 1, spec_value, valueFormat)

            elif "in" in values:
                converted = in_to_mm(values["in"])
                spec_value = f"{values['in']} in / {converted} mm"
                specificationWorksheet.write(row_index, 1, spec_value, valueFormat)

            elif "mm" in values:
                converted = mm_to_in(values["mm"])
                spec_value = f"{converted} in / {values['mm']} mm"
                specificationWorksheet.write(row_index, 1, spec_value, valueFormat)

            row_index += 1

        print("Information outputted to specifications excel file")
        specificationWorkbook.close()
    except Exception as e:
        print(f"Failed to extract specifications: {e}")
        return False

    return True