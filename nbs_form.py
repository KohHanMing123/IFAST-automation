import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def fill_nbs_form(driver, ref, nbs_df, main_window):
    try:
        first_row_client_col = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"] a.newWindow:not(.left)')
        first_row_client_col.click()

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

            # Currency DROPDOWN
            currency_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='SGD']")
            currency_dropdown.click()
            print("CURR DROP")

            select_curr_element = driver.find_element(By.ID, "Currency")
            time.sleep(1)

            desired_curr_text = nbs_data['Investment Currency']
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
            """, select_curr_element, desired_curr_text)

            driver.execute_script("arguments[0].click();", select_curr_element)

            print(f"Selected currency: {desired_curr_text}")

            time.sleep(1)

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
            # agent_dropdown = driver.find_element(By.XPATH, "//span[contains(@class, 'select-value') and text()='Singh,  Deepak']")  # All agent names are double spaced after first name, excel can just use single space eg. lastname, firstname
            agent_dropdown = driver.find_element(By.XPATH, '//*[@id="investform"]/fieldset/div/div[6]/span/span[1]') 
            agent_dropdown.click()
            print("Agent dropdown opened")

            select_element_agent = driver.find_element(By.ID, "User")
            print("element select found for agent")
            time.sleep(1)

            desired_agent_option_text = nbs_data['Agent']
            print(f"desired agent option text is {desired_agent_option_text}")
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

            print(f"Selected agent option: {desired_agent_option_text}")

            driver.find_element(By.ID, 'upfront_commission').send_keys(str(nbs_data['Upfront Comms']))

            # toggle FAF's switch if true, 1
            if nbs_data["FAF's"] == 1:
                print(nbs_data["FAF's"] == 1)
                faf_checkbox = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'onoffswitch-switch'))
                )
                faf_checkbox.click()

            time.sleep(2)

            driver.find_element(By.ID, 'fl_menu').click()
            print(f"{ref} has been saved")

            driver.close()

            # close nbs form window, switch back to the main window
            driver.switch_to.window(main_window)

            driver.refresh()

            # reload page and close the alert for resubmission
            try:
                alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert.accept()
                print("Alert closed")

            except Exception as e:
                print("No alert found or an error occurred:", str(e))

            print("ref col clicked after page reload")

        except IndexError:
            print(f"No matching NBS form data found for reference number {ref}")
    
    except IndexError:
        print(f"There is no record for {ref}")

    except Exception as e:
        print(f"An error occurred in fill_nbs_form: {str(e)}")
