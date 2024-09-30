from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time

def open_tab(driver, tab_xpath):
    try:
        tab = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, tab_xpath)))
        tab.click()
        print(f"Clicked on the tab with XPath: {tab_xpath}")
        time.sleep(3)
    except TimeoutException:
        print(f"Timeout while trying to click on the tab with XPath: {tab_xpath}")

def extract_processing_refs(driver):
    processing_refs = set()

    try:
        processlist = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'processlist')))
        processing_rows = processlist.find_elements(By.CSS_SELECTOR, '#processlist > tr')

        for row in processing_rows:
            try:
                ref_number = row.find_element(By.CSS_SELECTOR, 'td:nth-child(1) > a').text.strip()
                processing_refs.add(ref_number)
                print(f"Found ref number: {ref_number}")
            except NoSuchElementException:
                print("No reference number found in this row.")

    except TimeoutException:
        print("Timeout error while extracting references from Processing tab.")
    
    return processing_refs

def processing_tab(driver):
    try:
        processlist = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'processlist')))
        processing_rows = processlist.find_elements(By.TAG_NAME, 'tr')

        if not processing_rows:
            print("No rows to process in Processing tab.")
            return

        any_rows_processed = False

        for row in processing_rows:
            try:
                process_button = row.find_element(By.XPATH, './/a[contains(@id, "Processing") and contains(text(), "Process")]')
                process_button_link = process_button.get_attribute('href')
                driver.execute_script(f"window.open('{process_button_link}', '_blank');")
                time.sleep(2)
                any_rows_processed = True
            except NoSuchElementException:
                print("No 'Process' button found in this row.")

        if not any_rows_processed:
            print("No 'Process' buttons found in any rows.")

        print("All tabs opened. Processing each tab now...")

        all_tabs = driver.window_handles
        for tab in all_tabs[1:]:
            driver.switch_to.window(tab)
            time.sleep(3)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            try:
                process_now_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@value="PROCESS NOW"]')))
                process_now_button.click()
                print(f"Clicked 'PROCESS NOW' button in tab: {tab}")
                time.sleep(1)
            except TimeoutException:
                print(f"Error clicking 'PROCESS NOW' button in tab {tab}")

        driver.switch_to.window(all_tabs[0])
        print("Processing completed.")
    except NoSuchElementException as e:
        print(f"Error in processing_tab: {str(e)}")
    except TimeoutException as e:
        print(f"Timeout error in processing_tab: {str(e)}")

def to_payout_tab(driver):
    while True:
        try:
            payout_list = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'processlist')))
            payout_rows = payout_list.find_elements(By.TAG_NAME, 'tr')
            print("Payout rows found")

            if not payout_rows:
                print("No more rows to process in Pay Now tab.")
                break

            first_row = payout_rows[0]
            try:
                pay_now_button = first_row.find_element(By.XPATH, '//*[contains(@id, "Processing") and contains(text(), "Pay Now")]')
                pay_now_button.click()
                print("Clicked on the 'Pay Now' button.")
                time.sleep(3)

                main_window = driver.current_window_handle
                all_windows = driver.window_handles
                for window in all_windows:
                    if window != main_window:
                        driver.switch_to.window(window)
                        break

                save_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fl_menu"]')))
                save_button.click()
                print("Clicked the 'Save' button.")
                time.sleep(2)

                driver.switch_to.window(main_window)
                print("Switched back to the main window.")
                time.sleep(1)
            except NoSuchElementException:
                print("No 'Pay Now' button found in the current row.")
                break

        except NoSuchElementException:
            print("No more rows to process in Pay Now tab.")
            break
        except Exception as e:
            print(f"Error processing row: {str(e)}")
            break
            
    print("Finished processing the Pay Now tab.")

def expected_in_tab(driver, nbs_df):
    try:
        expected_in_list = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'processlist')))
        expected_in_rows = expected_in_list.find_elements(By.TAG_NAME, 'tr')

        if not expected_in_rows:
            print("No rows to process in Expected In tab.")
            return

        for row in expected_in_rows:
            try:
                provider_name = row.find_element(By.XPATH, './/td/a[@id="Processing"]').text.strip()
                print(f"Provider name is {provider_name}")
                print(f"Provider in {nbs_df['Provider'][0]}")

                if provider_name == nbs_df['Provider'][0]:
                    row.find_element(By.XPATH, './/td/a[@id="Processing"]').click()
                    print(f"Clicked on '{provider_name}' link.")
                    time.sleep(3)

                    checkboxes = driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
                    for checkbox in checkboxes:
                        checkbox.click()
                        print(f"Checked checkbox: {checkbox.get_attribute('id')}")

                    update_received_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@value="Update Received"]')))
                    update_received_button.click()
                    print("Clicked the 'Update Received' button.")
                    print(f"Process done for provider: {nbs_df['Provider'][0]}")
                else:
                    print(f"Provider name '{provider_name}' does not match expected value.")
            except NoSuchElementException:
                print("No Provider link found in the current row.")
            except Exception as e:
                print(f"Error processing row in Expected In tab: {str(e)}")
    except NoSuchElementException:
        print("No rows found in the Expected In tab.")
    except Exception as e:
        print(f"Error accessing the Expected In tab: {str(e)}")

    print("All records processed in Expected In tab.")
