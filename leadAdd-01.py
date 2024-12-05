import logging
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import openpyxl

# Set up logging
logging.basicConfig(filename="lead_submission.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# Path to ChromeDriver
chrome_driver_path = "D:\\LeadAutomation\\chromedriver-win64\\chromedriver.exe"
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)

# Load the Excel file
workbook = openpyxl.load_workbook("D:\\LeadAutomation\\leadAutoMation.xlsx")
sheet = workbook.active

# Function to handle form submission
def submit_form(name, email, mobile, state, city):
    try:
        # Open the form page
        driver.get("https://admissions.universalai.in/?utm_source=Admission+Ace&utm_medium=offline&utm_campaign=Santosh+Kumar&utm_id=other")

        # Fill Name
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "Name"))).send_keys(name)
        time.sleep(1)

        # Fill Email
        driver.find_element(By.ID, "Email").send_keys(email)
        time.sleep(1)

        # Fill Mobile
        driver.find_element(By.ID, "Mobile").send_keys(mobile)
        time.sleep(1)

        # Fill Password
        driver.find_element(By.ID, "Password").send_keys("12345")
        time.sleep(1)

        # Select Course
        Select(driver.find_element(By.ID, "CourseId")).select_by_visible_text("MBA Programs")
        time.sleep(1)

        # Select Specialization
        Select(driver.find_element(By.ID, "SpecializationId")).select_by_visible_text("MBA Single Specialization")
        time.sleep(1)

        # Prompt manual intervention for CAPTCHA and other fields
        print(f"Complete State, City, CAPTCHA, checkbox, and Submit for {name}.")
        input("Press Enter after manually completing these steps...")

        # Check if redirected to the Thank You page
        if "thank-you" in driver.current_url:
            logging.info(f"Form successfully submitted for {name} (Thank You page detected).")
        else:
            # Check for specific error messages related to already registered email or mobile
            try:
                # Check for email already registered
                error_message_email = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Your Email ID is already registered')]"))
                )
                logging.warning(f"Form submission error for {name}: {error_message_email.text}")
                return  # Skip to the next entry

            except Exception as e1:
                logging.error(f"Error checking email error for {name}: {e1}")
            
            try:
                # Check for mobile already registered
                error_message_mobile = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Your Mobile Number is already registered')]"))
                )
                logging.warning(f"Form submission error for {name}: {error_message_mobile.text}")
                return  # Skip to the next entry

            except Exception as e2:
                logging.error(f"Error checking mobile error for {name}: {e2}")
                logging.info(f"Form submitted for {name}, but no specific errors detected.")

        # Wait before the next entry (6 seconds)
        print("Waiting 6 seconds before processing the next entry...")
        time.sleep(6)

    except Exception as e:
        logging.error(f"Error submitting form for {name}: {e}")

# Process Excel rows
for row in sheet.iter_rows(min_row=2, values_only=True):
    name, email, mobile, state, city = row[:5]
    if not all([name, email, mobile]):  # Validate required fields
        logging.warning(f"Incomplete data for row: {row}. Skipping.")
        continue

    # Call the function to submit the form
    submit_form(name, email, mobile, state, city)

# Close browser
driver.quit()
