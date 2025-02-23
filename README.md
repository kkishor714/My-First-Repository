# My-First-Repository
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import os
import pandas as pd
from openpyxl import Workbook
from fake_useragent import UserAgent
import random

# Specify the remote debugging port
chrome_debugger_port = 9000

# Set up Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", f"localhost:{chrome_debugger_port}")
chrome_options.add_argument("--disable-javascript")  # Disable JavaScript

# Specify the path to your Chrome driver
chrome_driver_path = r'C:\\Users\\kishore.kumar\\AppData\\Local\\SeleniumBasic\\chromedriver.exe'
chrome_service = Service(executable_path=chrome_driver_path)

# Initialize the Chrome driver
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

# Load URLs and item names from the Excel file
excel_file = r"C:\\Users\\kishore.kumar\\Downloads\\PythonSS.xlsx"
sheet_name = "Sheet1"
url_column_name = "URL"
item_column_name = "ITEM"
screenshot_folder = r"C:\\Users\\kishore.kumar\\Downloads\\abcd"
page_load_timeout = 40

# Set up the headers in the new sheet
data_headers = [
    "URL", "Title", "Seller", "Seller Link", "Image Links", "Price 1", "Price 2", "Sold", "Stock", "Ships From"
]

# Create a new workbook and worksheet for data extraction
wb = Workbook()
ws = wb.active
ws.append(data_headers)

# Load data from Excel file
df = pd.read_excel(excel_file, sheet_name=sheet_name)
urls = df[url_column_name]
item_names = df[item_column_name]

# Set up UserAgent
ua = UserAgent()

# Path to save the updated Excel file
new_excel_file = r"C:\\Users\\kishore.kumar\\Downloads\\ExtractedData.xlsx"

try:
    for url, item_name in zip(urls, item_names):
        # Validate the URL
        if not isinstance(url, str) or not url.strip():
            print(f"Skipping invalid URL: {url}")
            row_data = [url or "Invalid URL"] + ["Data not found"] * (len(data_headers) - 1)
            ws.append(row_data)
            # Save workbook after each URL
            wb.save(new_excel_file)
            continue

        row_data = [url.strip()]  # Start with the URL

        try:
            # Set a new user-agent for each iteration
            user_agent = ua.random
            print(f"Using user-agent: {user_agent}")
            chrome_options.add_argument(f'user-agent={user_agent}')

            # Navigate to the URL
            driver.get(url.strip())
            driver.set_page_load_timeout(page_load_timeout)
            time.sleep(random.uniform(3, 4))

            # Extract data from the page
            try:
                title_element = driver.find_element(By.CSS_SELECTOR, ".WBVL_7")
                row_data.append(title_element.text)
            except:
                row_data.append("Title not found")

            try:
                seller_element = driver.find_element(By.CSS_SELECTOR, ".fV3TIn")
                row_data.append(seller_element.text)
            except:
                row_data.append("Seller not found")

            try:
                store_element = driver.find_element(By.CSS_SELECTOR, ".lG5Xxv")
                row_data.append(store_element.get_attribute("href"))
            except:
                row_data.append("Store link not found")

            try:
                image_elements = driver.find_elements(By.CSS_SELECTOR, "div.YM40Nc img")
                image_src_list = [img.get_attribute("src") for img in image_elements]
                row_data.append(', '.join(image_src_list))
            except:
                row_data.append("Images not found")

            try:
                price_element1 = driver.find_element(By.CSS_SELECTOR, ".IZPeQz")
                row_data.append(price_element1.text)
            except:
                row_data.append("Price 1 not found")

            try:
                price_element2 = driver.find_element(By.CSS_SELECTOR, ".ZA5sW5")
                row_data.append(price_element2.text)
            except:
                row_data.append("Price 2 not found")

            try:
                sold_element = driver.find_element(By.CSS_SELECTOR, ".AcmPRb")
                row_data.append(sold_element.text)
            except:
                row_data.append("Sold data not found")

            try:
                stock_element = driver.find_element(By.XPATH, "//div[@class='ybxj32']/label[text()='Stok']/following-sibling::div")
                row_data.append(stock_element.text)
            except:
                row_data.append("Stock not found")

            try:
                ships_from_element = driver.find_element(By.XPATH, "//div[@class='ybxj32']/label[text()='Dikirim Dari']/following-sibling::div")
                row_data.append(ships_from_element.text)
            except:
                row_data.append("Ships From not found")

            # Append the row data to the new sheet
            ws.append(row_data)

            # Capture the final screenshot
            try:
                screenshot = driver.get_screenshot_as_png()
                filename = f"{item_name}.png"
                screenshot_path = os.path.join(screenshot_folder, filename)
                with open(screenshot_path, 'wb') as file:
                    file.write(screenshot)
            except Exception as e:
                print(f"Error capturing final screenshot for {url}: {e}")

        except Exception as e:
            print(f"Error processing URL: {url}. Exception: {e}")
            row_data.extend(["Data not found"] * (len(data_headers) - len(row_data)))
            ws.append(row_data)

        # Save workbook after processing each URL
        wb.save(new_excel_file)

finally:
    # Close the browser
    driver.quit()
    print(f"Data saved to {new_excel_file}")
