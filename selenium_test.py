import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
from nbs_form import fill_nbs_form
from process_form import get_valid_f2f_answer, process_form
from commission_actions import open_tab, processing_tab, to_payout_tab, expected_in_tab

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.get('https://global.broker-backoffice.com//modules/investments/broker/')

input("Please log in and click into Investments, then press Enter to continue...")

nbs_excel_path = r'C:/Users/Han Ming/Documents/Han Ming/python learning/IFAST-automation/NBSForm.xlsx'  # to be made dynamic

nbs_df = pd.read_excel(nbs_excel_path)

main_window = driver.current_window_handle

for ref in nbs_df['Acc No.']:
    if pd.isna(ref):
        print("No more account numbers to process.")
        break

    search_input = driver.find_element(By.NAME, 'ref')
    search_input.clear()
    search_input.send_keys(ref)
    print(f'Searching for: {ref}')
    search_input.send_keys(Keys.RETURN)
    time.sleep(3)

    fill_nbs_form(driver, ref, nbs_df, main_window)  # performs all nbs form filling

    f2f_valid_answer = get_valid_f2f_answer(driver, main_window)

    if f2f_valid_answer:
        process_form(driver, main_window, ref, nbs_df, f2f_valid_answer)

# to click the commission tab after no ref number left
try:
    # Commission Tab
    open_tab(driver, '//*[@id="commission"]')
    processing_tab(driver)

    # To Payout Tab
    open_tab(driver, '//*[@id="To Payout"]') 
    to_payout_tab(driver)

    # Expected In tab
    open_tab(driver, '//*[@id="Expected In"]')
    expected_in_tab(driver, nbs_df)

except Exception as e:
    print(f"Error accessing or processing the commission work: {str(e)}")

# this last driver.quit to close the whole program
driver.quit()



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