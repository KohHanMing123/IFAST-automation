import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time

# Use this Chrome driver if you need to manually install Chrome driver, else mostly use webdriver manager
chrome_driver_path = r'C:\Users\winnie\Desktop\code playground\chromedriver-win64\chromedriver.exe'

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

excel_path = r'C:/Users/Han Ming\Documents/Han Ming/python learning/IFAST-automation/RefNumbers.xlsx'
df = pd.read_excel(excel_path)

main_window = driver.current_window_handle

for ref in df['Ref']:
    search_input = driver.find_element(By.NAME, 'ref')
    search_input.clear()
    search_input.send_keys(ref)
    print(f'Searching for: {ref}')
    search_input.send_keys(Keys.RETURN)
    time.sleep(3)

    try:
        first_row = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"] a.newWindow:not(.left)')
        first_row.click()
        
        time.sleep(3)

        all_windows = driver.window_handles
        for window in all_windows:
            if window != main_window:
                driver.switch_to.window(window)
                break

        nbs_tab = driver.find_element(By.ID, 'addinvestment')
        nbs_tab.click()
        
        # Close the new window and switch back to the main window
        # driver.close()
        # driver.switch_to.window(main_window)
        
        time.sleep(2)
    except Exception as e:
        print(f"Could not click on the first row: {e}")

driver.quit()
