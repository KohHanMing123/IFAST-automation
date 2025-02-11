import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def verify_new_record(driver, initial_deal_value):
    attempt = 0
    max_attempts = 2

    while attempt < max_attempts:
        try:
            time.sleep(2)
            first_row = driver.find_element(By.CSS_SELECTOR, 'tr[id="1"]')
            new_deal_value = first_row.get_attribute("deal") if first_row else None
            print(f"Attempt {attempt + 1} - New deal value: {new_deal_value}")

            if new_deal_value and new_deal_value != initial_deal_value:
                return True

        except Exception as e:
            print(f"Attempt {attempt + 1} - Error verifying record creation: {str(e)}")

        if attempt == 0:
            print("Reloading page and retrying...")
            driver.refresh()
            time.sleep(3)

        attempt += 1

    print("Record verification failed after two attempts.")
    return False

# def save_missing_references(file_path, no_records_refs, missing_refs):
#     """
#     Saves the missing reference numbers and records with errors to a text file.
#     """
#     try:
#         with open(file_path, 'w') as file:
#             if no_records_refs:
#                 file.write("New Records (No data found during NBS form filling):\n")
#                 file.writelines(f"{ref}\n" for ref in no_records_refs)
#             else:
#                 file.write("No new records\n")

#             if missing_refs:
#                 file.write("\nReferences missing due to errors:\n")
#                 file.writelines(f"{ref}\n" for ref in missing_refs)
#             else:
#                 file.write("\nNo mismatched refs")
#         print(f"Reference numbers saved to: {file_path}")
#     except Exception as e:
#         print(f"Error writing to file: {str(e)}")

# def search_reference(driver, ref):
#     """
#     Searches for a reference number in the system using the search input field.
#     """
#     try:
#         search_input = driver.find_element(By.NAME, 'ref')
#         search_input.clear()
#         search_input.send_keys(ref)
#         print(f'Searching for: {ref}')
#         search_input.send_keys(Keys.RETURN)
#         time.sleep(3)
#     except Exception as e:
#         print(f"Error searching for {ref}: {str(e)}")
