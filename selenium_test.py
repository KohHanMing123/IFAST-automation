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
import threading
from datetime import datetime
from nbs_form import fill_nbs_form
from process_form import get_valid_f2f_answer, process_form

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

# Paths for Excel files
# ref_excel_path = r'C:/Users/winnie/Desktop/code playground/RefNumbers.xlsx'  # to be made dynamic
nbs_excel_path = r'C:/Users/winnie/Desktop/code playground/NBSForm.xlsx'  # to be made dynamic

# ref_df = pd.read_excel(ref_excel_path)
nbs_df = pd.read_excel(nbs_excel_path)

main_window = driver.current_window_handle

# for ref in nbs_df['Acc No.']:
#     if pd.isna(ref):
#         print("No more account numbers to process.")
#         break

#     search_input = driver.find_element(By.NAME, 'ref')
#     search_input.clear()
#     search_input.send_keys(ref)
#     print(f'Searching for: {ref}')
#     search_input.send_keys(Keys.RETURN)
#     time.sleep(3)

#     fill_nbs_form(driver, ref, nbs_df, main_window)  # performs all nbs form filling

#     f2f_valid_answer = get_valid_f2f_answer(driver, main_window)

#     if f2f_valid_answer:
#         process_form(driver, main_window, ref, nbs_df, f2f_valid_answer)

# to click the commission tab after no ref number left
try:
    commission_tab = driver.find_element(By.XPATH, '//*[@id="commission"]')
    commission_tab.click()
    print("Clicked on the Commission tab.")
    time.sleep(3)

    # Process each record in the "processlist" table by opening in new tabs
    # processlist = driver.find_element(By.ID, 'processlist')
    # rows = processlist.find_elements(By.TAG_NAME, 'tr')

    # for row in rows:
    #     try:
    #         process_button = row.find_element(By.XPATH, './/a[contains(@id, "Processing") and contains(text(), "Process")]')
            
    #         process_button_link = process_button.get_attribute('href')

    #         driver.execute_script(f"window.open('{process_button_link}', '_blank');")
    #         time.sleep(2)
    #     except Exception as e:
    #         print(f"Error processing row: {str(e)}")

    # print("All tabs opened. Processing each tab now...")

    # # Switch to each tab, scroll down, and click the "PROCESS NOW" button
    # all_tabs = driver.window_handles
    # for tab in all_tabs[1:]:
    #     driver.switch_to.window(tab)
    #     time.sleep(3)
    #     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    #     time.sleep(1)
    #     try:
    #         process_now_button = driver.find_element(By.XPATH, '//*[@value="PROCESS NOW"]')
    #         process_now_button.click()
    #         print(f"Clicked 'PROCESS NOW' button in tab: {tab}")
    #         time.sleep(1)
    #         # driver.close()
    #     except Exception as e:
    #         print(f"Error clicking 'PROCESS NOW' button in tab {tab}: {str(e)}")

    # driver.switch_to.window(main_window)
    # print("Processing completed.")
    
    # Process "Pay Now" buttons in "To Payout" tab
    # to_payout_tab = driver.find_element(By.XPATH, '//*[@id="To Payout"]')  
    # to_payout_tab.click()
    # print("Clicked on the To Payout tab.")
    # time.sleep(3)

    # payout_list = driver.find_element(By.ID, 'processlist')
    # payout_rows = payout_list.find_elements(By.TAG_NAME, 'tr')

    # for row in payout_rows:
    #     print(f"how many inside {payout_rows}")
    #     try:
    #         pay_now_button = row.find_element(By.XPATH, './/a[contains(@id, "Processing") and contains(text(), "Pay Now")]')
    #         pay_now_button.click()
    #         time.sleep(3)

    #         # Switch to the newly opened window
    #         new_window = driver.window_handles[-1]
    #         driver.switch_to.window(new_window)
    #         time.sleep(1)

    #         try:
    #             save_button = driver.find_element(By.XPATH, '//*[@id="fl_menu"]')
    #             save_button.click()
    #             print("Clicked 'Save' button")
    #             time.sleep(1)
    #             driver.back()  # Switch back to the main window
    #             time.sleep(2) 

    #         except Exception as e:
    #             print(f"Error clicking 'Save' button: {str(e)}")

    #     except Exception as e:
    #         print(f"Error processing row in payout list: {str(e)}")

    # print("All payouts processed.")


    # Expected In tab
    expected_in_tab = driver.find_element(By.XPATH, '//*[@id="Expected In"]')  
    expected_in_tab.click()
    print("Clicked on the Expected In tab.")

    processlist = driver.find_element(By.ID, 'processlist')
    expected_in_rows = processlist.find_elements(By.TAG_NAME, 'tr')

    for row in expected_in_rows:
        try:
            # Find the <a> tag within the row
            provider_name = row.find_element(By.XPATH, './/td/a[@id="Processing"]').text.strip()
            print(f"provider name is {provider_name}")
            
            # Check if provider_name matches the value in the NBS form's Provider column
            print(f"provider in {nbs_df['Provider'][0]}")
            if provider_name == nbs_df['Provider'][0]: 

                row.find_element(By.XPATH, './/td/a[@id="Processing"]').click()
                print(f"Clicked on '{provider_name}' link.")
                time.sleep(3)  # Adjust wait time after clicking the link

                checkboxes = driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
                for checkbox in checkboxes:
                    checkbox.click()
                    print(f"Checked checkbox: {checkbox.get_attribute('id')}")
                    
            else:
                print(f"Provider name '{provider_name}' does not match expected value.")

        except Exception as e:
            print(f"Error processing row in Expected In tab: {str(e)}")

    print("All records processed in Expected In tab.")


except Exception as e:
    print(f"Error accessing or processing the To Payout tab: {str(e)}")


# Close the main window after processing
# driver.quit()


# driver.switch_to.window(main_window)


# this last driver.quit to close the whole program
# driver.quit()

# f2f_valid_answer = None

    # row_id = 2  
    # while True:
    #     try:
    #         print(f"Processing row id {row_id}")
    #         client_col_next = driver.find_element(By.CSS_SELECTOR, f'tr[id="{row_id}"] a.newWindow.left')
    #         client_col_next.click()
    #         time.sleep(3)

   
    #         all_windows_next = driver.window_handles
    #         for window_next in all_windows_next:
    #             if window_next != main_window:
    #                 driver.switch_to.window(window_next)
    #                 driver.maximize_window()
    #                 break

    #         dropdown_xpath = '//*[@id="investment"]/fieldset/div[2]/div[6]/span/span[1]'

    #         face_to_face_dropdown = driver.find_element(By.XPATH, dropdown_xpath)
    #         current_value = face_to_face_dropdown.text.strip()
    #         print(f"Current value for Face to Face dropdown is {current_value}")

    #         if current_value in ['Yes', 'No']:
    #             f2f_valid_answer = current_value
    #             print(f"Valid answer found for Face to Face dropdown: {f2f_valid_answer}")
    #             driver.close()
    #             driver.switch_to.window(main_window)
    #             break 
    #         else:
    #             driver.close()
    #             driver.switch_to.window(main_window)
    #             row_id += 1  

    #     except Exception as e:
    #         print(f"Error processing row {row_id}: {str(e)}")
    #         driver.switch_to.window(main_window)
    #         row_id += 1 

    # print(f"Proceeding with valid answer: {f2f_valid_answer}")
        
    #     try:
    #         print(f"Returning to the first row id 1")
    #         client_col_first = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"] a.newWindow.left')
    #         client_col_first.click()
    #         time.sleep(3)

    #         all_windows_first = driver.window_handles
    #         for window_first in all_windows_first:
    #             if window_first != main_window:
    #                 driver.switch_to.window(window_first)
    #                 driver.maximize_window()
    #                 break
            
    #         time.sleep(1)
    #         # F2F DROPDOWN
    #         f2f_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Select from list']")
    #         f2f_dropdown.click()
    #         print("f2f dropdown opened")

    #         select_element_f2f = driver.find_element(By.ID, "bespoke_432")
    #         print("element select found for f2f")
    #         time.sleep(1)

    #         desired_f2f_option_text = f2f_valid_answer
    #         print(f"desired f2f option text is {desired_f2f_option_text}")
            
    #         driver.execute_script("""
    #             var select = arguments[0];
    #             var desiredOption = arguments[1];
    #             for (var i = 0; i < select.options.length; i++) {
    #                 if (select.options[i].text === desiredOption) {
    #                     select.options[i].selected = true;
    #                     select.dispatchEvent(new Event('change', { 'bubbles': true }));
    #                     break;
    #                 }
    #             }
    #         """, select_element_f2f, desired_f2f_option_text)

    #         driver.execute_script("arguments[0].click();", select_element_f2f)
    #         print(f"Selected f2f option: {desired_f2f_option_text}")

    #         # DATE ISSUED DATEPICKER
    #         date_issued = nbs_df.loc[nbs_df['Acc No.'] == ref, 'Date Issued'].values[0]
    #         date_issued_formatted = datetime.strftime(pd.to_datetime(date_issued), "%d %b, %Y")
    #         print(f"Date issued for {ref} is {date_issued_formatted}")

    #         date_picker = driver.find_element(By.ID, 'dateiss')
    #         driver.execute_script("arguments[0].removeAttribute('readonly')", date_picker)
    #         date_picker.clear()
    #         date_picker.send_keys(date_issued_formatted)

    #         print(f"Date picker set to: {date_issued_formatted}")

    #         # WRITTEN IN DROPDOWN
    #         wi_dropdown = driver.find_element(By.XPATH, '//*[@id="investment"]/fieldset/div[2]/div[12]/span/span[1]')
    #         f2f_dropdown.click()
    #         print("WI dropdown opened")

    #         select_element_wi = driver.find_element(By.ID, "country")
    #         print("element select found for WI")
    #         time.sleep(1)

    #         desired_wi_option_text = "Singapore"
    #         print(f"desired wi option text is {desired_wi_option_text}")
            
    #         driver.execute_script("""
    #             var select = arguments[0];
    #             var desiredOption = arguments[1];
    #             for (var i = 0; i < select.options.length; i++) {
    #                 if (select.options[i].text === desiredOption) {
    #                     select.options[i].selected = true;
    #                     select.dispatchEvent(new Event('change', { 'bubbles': true }));
    #                     break;
    #                 }
    #             }
    #         """, select_element_wi, desired_wi_option_text)

    #         driver.execute_script("arguments[0].click();", select_element_wi)
    #         print(f"Selected f2f option: {desired_wi_option_text}")

    #         driver.find_element(By.ID, 'city').send_keys("Asia")
            
    #         #COST CENTRE DD
    #         cc_dropdown = driver.find_element(By.XPATH, '//*[@id="investment"]/fieldset/div[2]/div[14]/span/span[1]')
    #         cc_dropdown.click()
    #         print("CC dropdown opened")

    #         select_element_cc = driver.find_element(By.ID, "costcenter")
    #         print("element select found for cc")
    #         time.sleep(1)

    #         desired_cc_option_text = "Global Singapore"
    #         print(f"desired wi option text is {desired_cc_option_text}")
            
    #         driver.execute_script("""
    #             var select = arguments[0];
    #             var desiredOption = arguments[1];
    #             for (var i = 0; i < select.options.length; i++) {
    #                 if (select.options[i].text === desiredOption) {
    #                     select.options[i].selected = true;
    #                     select.dispatchEvent(new Event('change', { 'bubbles': true }));
    #                     break;
    #                 }
    #             }
    #         """, select_element_cc, desired_cc_option_text)

    #         driver.execute_script("arguments[0].click();", select_element_cc)
    #         print(f"Selected cost centre: {desired_cc_option_text}")

    #         #STATUS DD
    #         status_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='In Transit']")
    #         status_dropdown.click()
    #         print("status dropdown opened")

    #         select_element_status = driver.find_element(By.ID, "Pstatus")
    #         print("element select found for status")
    #         time.sleep(1)

    #         desired_status_option_text = "Process Now"
    #         print(f"desired status option text is {desired_status_option_text}")
            
    #         driver.execute_script("""
    #             var select = arguments[0];
    #             var desiredOption = arguments[1];
    #             for (var i = 0; i < select.options.length; i++) {
    #                 if (select.options[i].text === desiredOption) {
    #                     select.options[i].selected = true;
    #                     select.dispatchEvent(new Event('change', { 'bubbles': true }));
    #                     break;
    #                 }
    #             }
    #         """, select_element_status, desired_status_option_text)

    #         driver.execute_script("arguments[0].click();", select_element_status)
    #         print(f"Selected cost centre: {desired_status_option_text}")

    #         # driver.find_element(By.ID, 'fl_menu').click()
    #         print(f"{ref} has been changed to Process Now")

    #         driver.close()
    #         # close process form window, switch back to the main window
    #         driver.switch_to.window(main_window)

    #         time.sleep(20)

    #     except Exception as e:
    #         print(f"Error processing first row: {str(e)}")
    #         driver.switch_to.window(main_window)