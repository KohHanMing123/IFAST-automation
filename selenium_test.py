import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time

# Chrome driver options
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

# Initialize the driver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

# Read the Excel sheets
ref_excel_path = r'C:/Users/winnie/Desktop/code playground/RefNumbers.xlsx'
ref_df = pd.read_excel(ref_excel_path)
nbs_excel_path = r'C:/Users/winnie/Desktop/code playground/NBSForm.xlsx'
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
                driver.maximize_window()
                break

        nbs_tab = driver.find_element(By.ID, 'addinvestment')
        nbs_tab.click()
        
        time.sleep(3)

        try:
            nbs_data = nbs_df.loc[nbs_df['Acc No.'] == ref].iloc[0]

            driver.find_element(By.ID, 'Ref').send_keys(str(nbs_data['Acc No.']))
    
            # Click on the next input to trigger the popup
            amount_input = driver.find_element(By.ID, 'amount')
            amount_input.click()
            
            close_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'close-modal') and contains(text(), 'Close and continue')]"))
            )
            
            close_button.click()
            
            amount_input.send_keys(str(nbs_data['Amount Invested']))

            # Click to open the dropdown
            provider_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select a Provider']")
            provider_dropdown.click()
            print("im here")

            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//span[@class='select silver-gradient glossy  expandable-list replacement tracking open']//span[@class='drop-down']"))
            )
            print("im here 2")

            # Locate the desired option in the dropdown
            desired_option_text = nbs_data['Provider']
            desired_option_xpath = f"//span[@class='select silver-gradient glossy  expandable-list replacement tracking open']//span[@class='drop-down']/span[text()='{desired_option_text}']"
            desired_option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, desired_option_xpath))
            )
            print("im here 3")

            # Scroll into view and click the desired option
            # driver.execute_script("arguments[0].scrollIntoView(true);", desired_option)
            desired_option.click()

            # Select other fields based on the NBS data
            # Select(driver.find_element(By.ID, 'productlist')).select_by_visible_text(nbs_data['Product'])
            # Select(driver.find_element(By.ID, 'bespoke_428')).select_by_visible_text(nbs_data['Type'])
            # Select(driver.find_element(By.ID, 'User')).select_by_visible_text(nbs_data['Agent'])
            # driver.find_element(By.ID, 'upfront_commission').send_keys(str(nbs_data['Upfront Comms']))

            # Toggle FAF's switch if necessary
            # if nbs_data["FAF's"] == 1:
            #     driver.find_element(By.ID, 'FAF').click()
            
            # driver.find_element(By.ID, 'faf_per').send_keys(str(nbs_data['FAF Percentage']))
            # driver.find_element(By.ID, 'faf_frq').send_keys(nbs_data['FAF Frequency'])

        except IndexError:
            print(f"No matching NBS form data found for reference number {ref}")

        time.sleep(2)

    except IndexError:
        print(f"No matching NBS form data found for reference number {ref}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

driver.quit()




# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import Select
# from webdriver_manager.chrome import ChromeDriverManager
# import time

# # Use this Chrome driver if you need to manually install Chrome driver, else mostly use webdriver manager
# # chrome_driver_path = r'C:\Users\winnie\Desktop\code playground\chromedriver-win64\chromedriver.exe'

# options = webdriver.ChromeOptions()
# options.add_argument("--start-maximized")
# options.add_experimental_option("detach", True)

# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

# driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

# input("Please log in and click into Investments, then press Enter to continue...")

# # reading the RefNumbers excel sheet for ref no. to search
# ref_excel_path = r'C:/Users/winnie/Desktop/code playground/RefNumbers.xlsx'
# ref_df = pd.read_excel(ref_excel_path)

# # reading the NBSForm excel sheet to fill the NBS form based on the ref no. its linked to
# nbs_excel_path = r'C:/Users/winnie/Desktop/code playground/NBSForm.xlsx'
# nbs_df = pd.read_excel(nbs_excel_path)

# main_window = driver.current_window_handle

# for ref in ref_df['Ref']:
#     search_input = driver.find_element(By.NAME, 'ref')
#     search_input.clear()
#     search_input.send_keys(ref)
#     print(f'Searching for: {ref}')
#     search_input.send_keys(Keys.RETURN)
#     time.sleep(3)

#     try:
#         first_row = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"] a.newWindow:not(.left)')
#         first_row.click()

#         time.sleep(3)   

#         all_windows = driver.window_handles
#         for window in all_windows:
#             if window != main_window:
#                 driver.switch_to.window(window)
#                 driver.maximize_window()
#                 break

#         nbs_tab = driver.find_element(By.ID, 'addinvestment')
#         nbs_tab.click()
        
#         time.sleep(3)

#         # find respective input fields in the form
#         try:
#             nbs_data = nbs_df.loc[nbs_df['Acc No.'] == ref].iloc[0]

#             driver.find_element(By.ID, 'Ref').send_keys(str(nbs_data['Acc No.']))
    
#             # click on next input to trigger popup
#             amount_input = driver.find_element(By.ID, 'amount')
#             amount_input.click()
            
#             close_button = WebDriverWait(driver, 10).until(
#                 EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'close-modal') and contains(text(), 'Close and continue')]"))
#             )
            
#             close_button.click()
            
#             amount_input.send_keys(str(nbs_data['Amount Invested']))

#             print("past ref and investment, trying to open provider dropdwon")
#             # p_dropdown = driver.find_element(By.XPATH, 'providers')
#             # p_dropdown.click()
#             # time.sleep(1)
            
#             provider_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select a Provider']")
#             provider_dropdown.click()
#             print("I GOTTA BE HERE RIGHT")
#             WebDriverWait(driver, 10).until(
#                 EC.visibility_of_element_located((By.XPATH, "//span[@class='select silver-gradient glossy  expandable-list replacement tracking open']//span[@class='drop-down']"))
#             )
#             print("dropdown visible")

#             # Locate the desired option in the dropdown
#             desired_option_text = nbs_data['Provider']
#             print(f"desired option from excel file is {desired_option_text}")
#             desired_option_xpath = f"//span[@class='select silver-gradient glossy  expandable-list replacement tracking open']//span[@class='drop-down']/span[text()='{desired_option_text}']"
#             desired_option = WebDriverWait(driver, 10).until(
#                 EC.element_to_be_clickable((By.XPATH, desired_option_xpath))
#             )
#             print("dropdown options here")
            


#             driver.execute_script("arguments[0].scrollIntoView(true);", desired_option)
#             print(f"this is the desired option: {desired_option}")
#             # Click on the desired option to select it
#             desired_option.click()
            
#             Select(driver.find_element(By.ID, 'productlist')).select_by_visible_text(nbs_data['Product'])
#             Select(driver.find_element(By.ID, 'bespoke_428')).select_by_visible_text(nbs_data['Type']) # for Type
#             Select(driver.find_element(By.ID, 'User')).select_by_visible_text(nbs_data['Agent'])

#             # driver.find_element(By.ID, 'providers').send_keys(nbs_data['Provider'])
#             # driver.find_element(By.ID, 'productlist').send_keys(nbs_data['Product'])
#             # driver.find_element(By.ID, 'bespoke_428').send_keys(nbs_data['Type'])
#             # driver.find_element(By.ID, 'User').send_keys(nbs_data['Agent'])

#             driver.find_element(By.ID, 'upfront_commission').send_keys(str(nbs_data['Upfront Comms']))

#              # Toggle FAF's switch
#             if nbs_data["FAF's"] == 1:
#                 driver.find_element(By.ID, 'FAF').click()
            
#             driver.find_element(By.ID, 'faf_per').send_keys(str(nbs_data['FAF Percentage']))
#             driver.find_element(By.ID, 'faf_frq').send_keys(nbs_data['FAF Frequency'])
#         except IndexError:
#             print(f"No matching NBS form data found for reference number {ref}")

#         # Close the new window and switch back to the main window
#         # driver.close()
#         # driver.switch_to.window(main_window)
        
#         time.sleep(2)
#     # except None as e:
#     #     print(f"Option '{desired_option_text}' not found in the dropdown: {str(e)}")
#     except IndexError:
#         print(f"No matching NBS form data found for reference number {ref}")
#     except Exception as e:
#         print(f"An error occurred: {str(e)}")

# driver.quit()
