import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
import time

# Path to your manually downloaded ChromeDriver
chrome_driver_path = r'C:\Users\winnie\Desktop\code playground\chromedriver-win64\chromedriver.exe'

# Set up the Chrome options
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

# Initialize the WebDriver with the correct path and options
driver = webdriver.Chrome(service=ChromeService(chrome_driver_path), options=options)

# Open the webpage
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

# Wait for the page to load and allow manual authentication
input("Please log in and click into Investments, then press Enter to continue...")

# Read the reference numbers from the Excel file
excel_path = r'C:/Users/winnie/Desktop/code playground/RefNumbers.xlsx'
df = pd.read_excel(excel_path)

# Iterate through each reference number
for ref in df['Ref']:
    # Find the search input field and enter the reference number
    search_input = driver.find_element(By.NAME, 'ref')
    search_input.clear()
    search_input.send_keys(ref)
    print(f'Searching for: {ref}') 
    search_input.send_keys(Keys.RETURN)

    # Wait for the search results to load
    time.sleep(3)

    # Click the first row of data in the search results
    try:
        first_row = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"] a')
        first_row.click()

        # Wait for the new page to load
        time.sleep(3)

        # Go back to the search page
        driver.back()
        time.sleep(2)
    except Exception as e:
        print(f"Could not click on the first row: {e}")

# Close the browser
driver.quit()
