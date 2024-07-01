import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

ref_excel_path = r'C:/Users/Han Ming/Documents/Han Ming/python learning/IFAST-automation/RefNumbers.xlsx'
ref_df = pd.read_excel(ref_excel_path)
nbs_excel_path = r'C:/Users/Han Ming/Documents/Han Ming/python learning/IFAST-automation/NBSForm.xlsx'
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
    
            amount_input = driver.find_element(By.ID, 'amount')
            amount_input.click()
            
            close_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'close-modal') and contains(text(), 'Close and continue')]"))
            )
            
            close_button.click()
            
            amount_input.send_keys(str(nbs_data['Amount Invested']))


            # FOR PROVIDER DROPDOWN
            provider_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select a Provider']")
            provider_dropdown.click()
            print("Dropdown opened")

            select_element = driver.find_element(By.ID, "providers")
            time.sleep(1) 

            desired_option_text = nbs_data['Provider']
            driver.execute_script("""
                var select = arguments[0];
                var desiredOption = arguments[1];
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === desiredOption) {
                        select.options[i].selected = true;
                        select.dispatchEvent(new Event('change', { 'bubbles': true }));
                        break;
                    }
                }
            """, select_element, desired_option_text)
            
            driver.execute_script("arguments[0].click();", select_element)

            print(f"Selected option: {desired_option_text}")

            time.sleep(1)

            # FOR PRODUCT DROPDOWN
            product_dropdown = driver.find_element(By.CSS_SELECTOR, 'span[id="productlist-holder"]')
            product_dropdown.click()
            print("Product dropdown opened")

            select_element_product = driver.find_element(By.ID, "productlist")
            print("aft productlist")
            time.sleep(1)

            print("starting to get data Product from excel")
            desired_product_option_text = str(nbs_data['Product'])
            print(f"desired product option text is {desired_product_option_text}")
            driver.execute_script("""
                var select = arguments[0];
                var desiredOption = arguments[1];
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === desiredOption) {
                        select.options[i].selected = true;
                        select.dispatchEvent(new Event('change', { 'bubbles': true }));
                        break;
                    }
                }
            """, select_element_product, desired_product_option_text)
            
            driver.execute_script("arguments[0].click();", select_element_product)

            print(f"Selected product option: {desired_product_option_text}")

            # FOR TYPE DROPDOWN
            type_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select from list']")
            type_dropdown.click()
            print("Type dropdown opened")

            select_element_type = driver.find_element(By.ID, "bespoke_428")
            print("element select found for type")
            time.sleep(1)
    
            desired_type_option_text = nbs_data['Type']
            print(f"desired product option text is {desired_type_option_text}")
            driver.execute_script("""
                var select = arguments[0];
                var desiredOption = arguments[1];
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === desiredOption) {
                        select.options[i].selected = true;
                        select.dispatchEvent(new Event('change', { 'bubbles': true }));
                        break;
                    }
                }
            """, select_element_type, desired_type_option_text)
            
            driver.execute_script("arguments[0].click();", select_element_type)

            print(f"Selected product option: {desired_type_option_text}")

            # FOR AGENT DROPDOWN
            agent_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Singh,  Deepak']") # All agent names are double spaced after first name,
            agent_dropdown.click()
            print("Agent dropdown opened")

            select_element_agent = driver.find_element(By.ID, "User")
            print("element select found for agent")
            time.sleep(1)
    
            desired_agent_option_text = nbs_data['Agent']
            print(f"desired product option text is {desired_agent_option_text}")
            driver.execute_script("""
                var select = arguments[0];
                var desiredOption = arguments[1];
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === desiredOption) {
                        select.options[i].selected = true;
                        select.dispatchEvent(new Event('change', { 'bubbles': true }));
                        break;
                    }
                }
            """, select_element_agent, desired_agent_option_text)
            
            driver.execute_script("arguments[0].click();", select_element_agent)

            print(f"Selected product option: {desired_agent_option_text}")

            driver.find_element(By.ID, 'upfront_commission').send_keys(str(nbs_data['Upfront Comms']))

            # Toggle FAF's switch if true, 1
            if nbs_data["FAF's"] == 1:
                driver.find_element(By.ID, 'FAF').click()
            
            driver.find_element(By.ID, 'faf_per').send_keys(str(nbs_data['FAF Percentage']))

            # FOR FAF FREQUENCY DROPDOWN
            faf_freq_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()=' Annually']")
            faf_freq_dropdown.click()
            print("faf freq dropdown opened")

            select_element_faf_freq = driver.find_element(By.ID, "faf_frq")
            print("element select found for faf freq")
            time.sleep(1)
    
            desired_faf_freq_option_text = nbs_data['FAF Frequency']
            print(f"desired product option text is {desired_faf_freq_option_text}")
            driver.execute_script("""
                var select = arguments[0];
                var desiredOption = arguments[1];
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === desiredOption) {
                        select.options[i].selected = true;
                        select.dispatchEvent(new Event('change', { 'bubbles': true }));
                        break;
                    }
                }
            """, select_element_faf_freq, desired_faf_freq_option_text)
            
            driver.execute_script("arguments[0].click();", select_element_faf_freq)

            print(f"Selected product option: {desired_faf_freq_option_text}")

            driver.find_element(By.ID, 'fl_menu').click()
            print(f"{ref} has been saved")

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
