import os
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

desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')

nbs_excel_path = os.path.join(desktop_dir, 'NBSForm.xlsx') # target any machines desktop folder

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
