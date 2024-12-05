from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import openpyxl

# Load Excel data
excel_file = "D:\\LeadAutomation\\leadAutoMation.xlsx"  # Full path of the Excel file
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Configure WebDriver
from selenium.webdriver.chrome.service import Service

service = Service('C:\\Users\\skyde\\chromedriver-win64\\chromedriver.exe')  # Path to ChromeDriver
driver = webdriver.Chrome(service=service)

driver.get("https://admissions.universalai.in/?utm_source=Admission+Ace&utm_medium=offline&utm_campaign=Santosh+Kumar&utm_id=other")

# Open a log file to record failed entries
log_file = open("D:\\LeadAutomation\\error_log.txt", "w")

# Process each row in the Excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
    name, email, mobile, state, city = row

    try:
        # Fill form fields
        driver.find_element(By.ID, "Name").send_keys(name)
        driver.find_element(By.ID, "Email").send_keys(email)
        driver.find_element(By.ID, "country_dial_codeMobile").send_keys("+91")
        driver.find_element(By.ID, "Mobile").send_keys(mobile)
        driver.find_element(By.ID, "Password").send_keys("12345")  # Fixed password

        # Select dropdowns
        Select(driver.find_element(By.ID, "StateId")).select_by_visible_text(state)
        Select(driver.find_element(By.ID, "CityId")).select_by_visible_text(city)
        driver.find_element(By.ID, "CourseId").send_keys("MBA Programs")
        driver.find_element(By.ID, "SpecializationId").send_keys("MBA Single Specialization")

        # Handle CAPTCHA (manual step required unless automated)
        input("Enter CAPTCHA and press Enter to proceed...")

        # Agree and submit
        driver.find_element(By.ID, "Agree").click()
        driver.find_element(By.ID, "registerBtn").click()

        time.sleep(3)  # Wait for submission to process
        print(f"Form submitted for {name}")

    except Exception as e:
        # Log errors for duplicate or failed entries
        log_file.write(f"Failed for {name}, {email}, {mobile}: {str(e)}\n")
        print(f"Error submitting form for {name}: {e}")

# Close log file and browser
log_file.close()
driver.quit()
