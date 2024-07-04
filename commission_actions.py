from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time

def open_tab(driver, tab_xpath):
    tab = driver.find_element(By.XPATH, tab_xpath)
    tab.click()
    print(f"Clicked on the tab with XPath: {tab_xpath}")
    time.sleep(3)

def processing_tab(driver):
    processlist = driver.find_element(By.ID, 'processlist')
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
            process_now_button = driver.find_element(By.XPATH, '//*[@value="PROCESS NOW"]')
            process_now_button.click()
            print(f"Clicked 'PROCESS NOW' button in tab: {tab}")
            time.sleep(1)
        except Exception as e:
            print(f"Error clicking 'PROCESS NOW' button in tab {tab}: {str(e)}")

    driver.switch_to.window(all_tabs[0])
    print("Processing completed.")

def to_payout_tab(driver):
    while True:
        try:
            payout_list = driver.find_element(By.ID, 'processlist')
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

                save_button = driver.find_element(By.XPATH, '//*[@id="fl_menu"]')
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
        processlist = driver.find_element(By.ID, 'processlist')
        expected_in_rows = processlist.find_elements(By.TAG_NAME, 'tr')

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

                    update_received_button = driver.find_element(By.XPATH, '//*[@value="Update Received"]')
                    update_received_button.click()
                    print("Clicked the 'Update Received' button.")
                    print(f"Process done for provider: {nbs_df['Provider'][0]}")
                else:
                    print(f"Provider name '{provider_name}' does not match expected value.")
            except NoSuchElementException:
                print(f"No Provider link found in the current row for provider: {provider_name}")
            except Exception as e:
                print(f"Error processing row in Expected In tab: {str(e)}")
    except NoSuchElementException:
        print("No rows found in the Expected In tab.")
    except Exception as e:
        print(f"Error accessing the Expected In tab: {str(e)}")

    print("All records processed in Expected In tab.")