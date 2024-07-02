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
from nbs_form import fill_nbs_form

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

# Paths for Excel files
ref_excel_path = r'C:/Users/winnie/Desktop/code playground/RefNumbers.xlsx'  # to be made dynamic
nbs_excel_path = r'C:/Users/winnie/Desktop/code playground/NBSForm.xlsx'  # to be made dynamic

ref_df = pd.read_excel(ref_excel_path)
nbs_df = pd.read_excel(nbs_excel_path)

main_window = driver.current_window_handle

for ref in ref_df['Ref']:
    search_input = driver.find_element(By.NAME, 'ref')
    search_input.clear()
    search_input.send_keys(ref)
    print(f'Searching for: {ref}')
    search_input.send_keys(Keys.RETURN)
    time.sleep(3)

    # fill_nbs_form(driver, ref, nbs_df, main_window)  # performs all nbs form filling

    f2f_valid_answer = None

    row_id = 2  
    while True:
        try:
            print(f"Processing row id {row_id}")
            client_col_next = driver.find_element(By.CSS_SELECTOR, f'tr[id="{row_id}"] a.newWindow.left')
            client_col_next.click()
            time.sleep(3)

   
            all_windows_next = driver.window_handles
            for window_next in all_windows_next:
                if window_next != main_window:
                    driver.switch_to.window(window_next)
                    driver.maximize_window()
                    break

            dropdown_xpath = '//*[@id="investment"]/fieldset/div[2]/div[6]/span/span[1]'

            face_to_face_dropdown = driver.find_element(By.XPATH, dropdown_xpath)
            current_value = face_to_face_dropdown.text.strip()
            print(f"Current value for Face to Face dropdown is {current_value}")

            if current_value in ['Yes', 'No']:
                f2f_valid_answer = current_value
                print(f"Valid answer found for Face to Face dropdown: {f2f_valid_answer}")
                driver.close()
                break 
            else:
                driver.close()
                driver.switch_to.window(main_window)
                row_id += 1  

        except Exception as e:
            print(f"Error processing row {row_id}: {str(e)}")
            driver.switch_to.window(main_window)
            row_id += 1 

    if f2f_valid_answer: # cant direct to id 1 after checking
        print(f"Proceeding with valid answer: {f2f_valid_answer}")
        
        try:
            print(f"Returning to the first row id 1")
            client_col_first = driver.find_element(By.XPATH, '//*[@id="1"]/td/a[@id="policyinfo"]')
            client_col_first.click()
            time.sleep(3)

            # all_windows_first = driver.window_handles
            # for window_first in all_windows_first:
            #     if window_first != main_window:
            #         driver.switch_to.window(window_first)
            #         driver.maximize_window()
            #         break

            f2f_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select from list']")
            f2f_dropdown.click()
            print("f2f dropdown opened")

            select_element_f2f = driver.find_element(By.ID, "bespoke_432")
            print("element select found for f2f")
            time.sleep(1)

            desired_f2f_option_text = f2f_valid_answer
            print(f"desired f2f option text is {desired_f2f_option_text}")
            
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
            """, select_element_f2f, desired_f2f_option_text)

            driver.execute_script("arguments[0].click();", select_element_f2f)
            print(f"Selected agent option: {desired_f2f_option_text}")

        except Exception as e:
            print(f"Error processing first row: {str(e)}")
            driver.switch_to.window(main_window)

driver.switch_to.window(main_window)


# this last driver.quit to close the whole program
# driver.quit()

#             driver.find_element(By.ID, 'faf_per').send_keys(str(nbs_data['FAF Percentage']))

#             # FOR FAF FREQUENCY DROPDOWN
#             faf_freq_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()=' Annually']")
#             faf_freq_dropdown.click()
#             print("faf freq dropdown opened")

#             select_element_faf_freq = driver.find_element(By.ID, "faf_frq")
#             print("element select found for faf freq")
#             time.sleep(1)
    
#             desired_faf_freq_option_text = nbs_data['FAF Frequency']
#             print(f"desired product option text is {desired_faf_freq_option_text}")
#             driver.execute_script("""
#                 var select = arguments[0];
#                 var desiredOption = arguments[1];
#                 for (var i = 0; i < select.options.length; i++) {
#                     if (select.options[i].text === desiredOption) {
#                         select.options[i].selected = true;
#                         select.dispatchEvent(new Event('change', { 'bubbles': true }));
#                         break;
#                     }
#                 }
#             """, select_element_faf_freq, desired_faf_freq_option_text)
            
#             driver.execute_script("arguments[0].click();", select_element_faf_freq)

#             print(f"Selected product option: {desired_faf_freq_option_text}")