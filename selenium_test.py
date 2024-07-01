import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time

# Use this Chrome driver if you need to manually install Chrome driver, else mostly use webdriver manager
# chrome_driver_path = r'C:\Users\winnie\Desktop\code playground\chromedriver-win64\chromedriver.exe'

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")


# reading the RefNumbers excel sheet for ref no. to search
ref_excel_path = r'C:/Users/Han Ming\Documents/Han Ming/python learning/IFAST-automation/RefNumbers.xlsx'
ref_df = pd.read_excel(ref_excel_path)

# reading the NBSForm excel sheet to fill the NBS form based on the ref no. its linked to
nbs_excel_path = r'C:/Users/Han Ming\Documents/Han Ming/python learning/IFAST-automation/NBSForm.xlsx'
nbs_df = pd.read_excel(nbs_excel_path)

main_window = driver.current_window_handle

for ref in ref_df['Ref']:
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
        
        time.sleep(3)

        # find respective input fields in the form
        try:
            nbs_data = nbs_df.loc[nbs_df['Acc No.'] == ref].iloc[0]

            driver.find_element(By.ID, 'Ref').send_keys(str(nbs_data['Acc No.']))
            driver.find_element(By.ID, 'amount').send_keys(str(nbs_data['Amount Invested']))

            Select(driver.find_element(By.ID, 'providers')).select_by_visible_text(nbs_data['Provider'])
            Select(driver.find_element(By.ID, 'productlist')).select_by_visible_text(nbs_data['Product'])
            Select(driver.find_element(By.ID, 'bespoke_428')).select_by_visible_text(nbs_data['Type']) # for Type
            Select(driver.find_element(By.ID, 'User')).select_by_visible_text(nbs_data['Agent'])

            # driver.find_element(By.ID, 'providers').send_keys(nbs_data['Provider'])
            # driver.find_element(By.ID, 'productlist').send_keys(nbs_data['Product'])
            # driver.find_element(By.ID, 'bespoke_428').send_keys(nbs_data['Type'])
            # driver.find_element(By.ID, 'User').send_keys(nbs_data['Agent'])

            driver.find_element(By.ID, 'upfront_commission').send_keys(str(nbs_data['Upfront Comms']))

             # Toggle FAF's switch
            if nbs_data["FAF's"] == 1:
                driver.find_element(By.ID, 'FAF').click()
            
            driver.find_element(By.ID, 'faf_per').send_keys(str(nbs_data['FAF Percentage']))
            driver.find_element(By.ID, 'faf_frq').send_keys(nbs_data['FAF Frequency'])
        except IndexError:
            print(f"No matching NBS form data found for reference number {ref}")

        # Close the new window and switch back to the main window
        # driver.close()
        # driver.switch_to.window(main_window)
        
        time.sleep(2)
    except Exception as e:
        print(f"Could not click on the first row: {e}")

driver.quit()
