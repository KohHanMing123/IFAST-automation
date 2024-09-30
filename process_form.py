import time
from datetime import datetime
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_valid_f2f_answer(driver, main_window):
    f2f_valid_answer = None
    row_id = 2

    while True:
        try:
            print(f"Processing row id {row_id}")

            client_col_next = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, f'tr[id="{row_id}"] a.newWindow.left'))
            )
            client_col_next.click()
            time.sleep(3)

            all_windows_next = driver.window_handles
            for window_next in all_windows_next:
                if window_next != main_window:
                    driver.switch_to.window(window_next)
                    break

            time.sleep(1)

            dropdown_xpath = '//*[@id="investment"]/fieldset/div[2]/div[6]/span/span[1]'
            face_to_face_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, dropdown_xpath))
            )
            current_value = face_to_face_dropdown.text.strip()
            print(f"Current value for Face to Face dropdown is {current_value}")

            # If a valid answer is found, return it
            if current_value in ['Yes', 'No']:
                f2f_valid_answer = current_value
                print(f"Valid answer found for Face to Face dropdown: {f2f_valid_answer}")
                driver.close()
                driver.switch_to.window(main_window)
                return f2f_valid_answer
            else:
                print(f"Invalid value '{current_value}' for Face to Face dropdown. Moving to next row.")
                driver.close()
                driver.switch_to.window(main_window)
                row_id += 1

        except Exception as e:
            print(f"Error processing row {row_id}: {str(e)}")
            driver.switch_to.window(main_window)
            row_id += 1

            # If there are no more rows, exit and return None
            if row_id > 5:
                print("No valid Face to Face answer found after checking all rows.")
                break

    return None

def process_form(driver, main_window, ref, nbs_df, f2f_valid_answer):
    print(f"Proceeding with valid answer: {f2f_valid_answer}")
        
    try:
        print(f"Returning to the first row id 1 USING CSS SELECTOR")
        # client_col_first = WebDriverWait(driver, 10).until(
        #     EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div[2]/span/table[3]/tbody/tr[1]/td[1]/a'))
        # )
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'tr[id="1"] a#policyinfo'))
        )
        
        client_col_first = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'tr[id="1"] a#policyinfo'))
        )
        # driver.execute_script("arguments[0].scrollIntoView(true);", client_col_first) 

        time.sleep(1) 
            
        try:
            tooltip = driver.find_element(By.CLASS_NAME, 'ui-tooltip')
            driver.execute_script("arguments[0].style.display = 'none';", tooltip)
            # time.sleep(1)
        except:
            print("Tooltip not found or already hidden.")
            pass

        client_col_first.click()
        # time.sleep(3)

        all_windows_first = driver.window_handles
        for window_first in all_windows_first:
            if window_first != main_window:
                driver.switch_to.window(window_first)
                driver.maximize_window()
                break
        
        time.sleep(1)
        
        try:
            f2f_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'select-value') and text()='Select from list']"))
            )
            f2f_dropdown.click()
            print("f2f dropdown opened")

            select_element_f2f = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "bespoke_432"))
            )
            print("Element select found for f2f")
            time.sleep(1)

            desired_f2f_option_text = f2f_valid_answer
            print(f"Desired f2f option text is {desired_f2f_option_text}")

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
            print(f"Selected f2f option: {desired_f2f_option_text}")

        except Exception as e:
            print(f"Error processing dropdown: {str(e)}")

        # DATE ISSUED DATEPICKER
        date_issued = nbs_df.loc[nbs_df['Acc No.'] == ref, 'Date Issued'].values[0]
        date_issued_formatted = datetime.strftime(pd.to_datetime(date_issued), "%d %b, %Y")
        print(f"Date issued for {ref} is {date_issued_formatted}")

        date_picker = driver.find_element(By.ID, 'dateiss')
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_picker)
        date_picker.clear()
        date_picker.send_keys(date_issued_formatted)
        date_picker.send_keys(Keys.RETURN)
        print(f"Date picker set to: {date_issued_formatted}")

        # WRITTEN IN DROPDOWN
        try:
            wi_dropdown = driver.find_element(By.XPATH, '//*[@id="investment"]/fieldset/div[2]/div[12]/span/span[1]')
            driver.execute_script("arguments[0].click();", wi_dropdown)
            print("WI dropdown opened")
        except Exception as e:
            print(f"Error clicking WI dropdown: {str(e)}")

        select_element_wi = driver.find_element(By.ID, "country")
        print("element select found for WI")
        time.sleep(1)

        desired_wi_option_text = "Singapore"
        print(f"desired wi option text is {desired_wi_option_text}")
        
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
        """, select_element_wi, desired_wi_option_text)

        driver.execute_script("arguments[0].click();", select_element_wi)
        print(f"Selected wi option: {desired_wi_option_text}")

        driver.find_element(By.ID, 'city').send_keys("Asia")

        driver.execute_script("scroll(0, 350);")
        print("scrolled to middle page")
        
        #COST CENTRE DD
        try:
            cc_dropdown = driver.find_element(By.XPATH, '//*[@id="investment"]/fieldset/div[2]/div[14]/span/span[1]')
            driver.execute_script("arguments[0].click();", wi_dropdown)
            print("CC dropdown opened")
        except Exception as e:
            print(f"Error clicking CC dropdown: {str(e)}")

        select_element_cc = driver.find_element(By.ID, "costcenter")
        print("element select found for cc")
        time.sleep(1)

        desired_cc_option_text = "Global Singapore"
        print(f"desired wi option text is {desired_cc_option_text}")
        
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
        """, select_element_cc, desired_cc_option_text)

        driver.execute_script("arguments[0].click();", select_element_cc)
        print(f"Selected cost centre: {desired_cc_option_text}")

        #STATUS DD
        try:
            status_dropdown = driver.find_element(By.XPATH, '//*[@id="investment"]/fieldset/div[2]/div[16]/span/span[1]')

            print("am i found for status")
            driver.execute_script("arguments[0].click();", status_dropdown)
            print("status dropdown opened")
        except Exception as e:
            print(f"Error clicking CC dropdown: {str(e)}")

        select_element_status = driver.find_element(By.ID, "Pstatus")
        print("element select found for status")
        time.sleep(1)

        desired_status_option_text = "Process Now"
        print(f"desired status option text is {desired_status_option_text}")
        
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
        """, select_element_status, desired_status_option_text)

        driver.execute_script("arguments[0].click();", select_element_status)
        print(f"Selected status: {desired_status_option_text}")

        driver.find_element(By.ID, 'fl_menu').click()
        print(f"{ref} has been changed to Process Now")

        driver.close()
        # close process form window, switch back to the main window
        driver.switch_to.window(main_window)
        driver.refresh()
        time.sleep(3)

    except Exception as e:
        print(f"Error processing first row: {str(e)}")
        driver.switch_to.window(main_window)