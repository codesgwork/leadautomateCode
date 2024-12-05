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

        # Prompt manual intervention for State, City, CAPTCHA, checkbox
        print(f"Complete State, City, CAPTCHA, checkbox, and Submit for {name}.")
        input("Press Enter after manually completing these steps...")

        # Check if the email or mobile is already registered
        try:
            # Look for error message related to email or mobile
            error_message = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Your Email ID is already registered')]"))
            )
            logging.warning(f"Details already exist for {name}: {error_message.text}")
            print(f"Skipping {name} due to already existing details.")
            time.sleep(6)  # Wait for 6 seconds before moving to the next entry
            return  # Skip to next entry
        except Exception as e:
            logging.info(f"No existing details for {name}. Proceeding with submission.")

        # Wait before submitting
        time.sleep(1)

        # Click Submit Button
        driver.find_element(By.ID, "Submit").click()

        # Wait for Success or Error Message
        time.sleep(3)  # Wait for a few seconds after clicking submit to handle the response

        # Check if the form submission was successful or failed
        try:
            # If redirected to "Thank You" page
            WebDriverWait(driver, 10).until(
                EC.url_contains("thank-you")
            )
            logging.info(f"Form successfully submitted for {name}.")
            time.sleep(6)  # Wait for 6 seconds before processing the next entry
        except Exception as e:
            logging.error(f"Error during form submission for {name}: {e}")
        
        # Wait before moving to the next entry
        time.sleep(6)

    except Exception as e:
        logging.error(f"Error submitting form for {name}: {e}")
        time.sleep(6)  # Wait before processing the next entry in case of any other error

# Process Excel rows
for row in sheet.iter_rows(min_row=2, values_only=True):
    name, email, mobile, state, city = row[:5]
    if not all([name, email, mobile]):  # Validate required fields
        logging.warning(f"Incomplete data for row: {row}. Skipping.")
        continue
    submit_form(name, email, mobile, state, city)

# Close browser
driver.quit()
