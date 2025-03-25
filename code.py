import time
import os
import logging
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, \
    UnexpectedAlertPresentException

# Configure logging
log_file = "test_log.txt"
logging.basicConfig(filename=log_file, level=logging.INFO, format="%(asctime)s - %(message)s")

# Create a directory for screenshots if it doesn't exist
screenshot_dir = "screenshots"
if not os.path.exists(screenshot_dir):
    os.makedirs(screenshot_dir)

# Load test data from Excel using openpyxl
workbook = openpyxl.load_workbook("Excel_file/Username.xlsx")
sheet = workbook['LoginTest']

# Setup WebDriver
driver = webdriver.Chrome()  # Ensure you have chromedriver installed
login_url = "https://training.pfms.gov.in/SitePages/Users/LoginDetails/Login.aspx"  # Replace with actual login URL
driver.get(login_url)
driver.maximize_window()
wait = WebDriverWait(driver, 15)


# Function to take a screenshot
def take_screenshot(username, error_type):
    screenshot_path = os.path.join(screenshot_dir, f"{error_type}_{username}.png")
    driver.save_screenshot(screenshot_path)
    logging.error(f"Screenshot saved at {screenshot_path}")


# Loop through test data (Assuming data starts from row 2)
for row in sheet.iter_rows(min_row=2, values_only=True):
    username, password, expected_result = row

    try:
        # Ensure login page is loaded
        driver.get(login_url)
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='UserName']")))
        # Find and input username
        # Enter username and password
        driver.find_element(By.XPATH, "//input[@id='UserName']").send_keys(username)
        driver.find_element(By.XPATH, "//input[@id='Password']").send_keys(password)
        driver.execute_script('window.scrollTo(0,300)')  # to scroll the page till given height
        time.sleep(15)  # Enter Captcha manually
        driver.find_element(By.XPATH, "//input[@id='ctl00_cphBody_btnLoginButton']").click()
        time.sleep(2)  # Allow time for page load
        header = driver.title


        # Handle pop-ups or alerts
        try:
            alert = driver.switch_to.alert
            logging.warning(f"Unexpected alert detected for {username}: {alert.text}")
            alert.accept()
        except UnexpectedAlertPresentException:
            logging.warning(f"Unhandled alert present for {username}")

        # Check if login was successful
        try:
            wait.until(
                EC.presence_of_element_located((By.LINK_TEXT, "Central Plan Scheme Monitoring System [CPSMS]")))  # Replace with actual landing page main element ID
            result = "Pass"
        except TimeoutException:
            result = "Fail"

        # Log results
        logging.info(f"Test case: Username={username}, Expected={expected_result}, Got={result}")

        # Take screenshot on failure
        if result != expected_result:
            take_screenshot(username, "failure")

        # Logout if login was successful
        if result == "Pass":
            try:
                wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#ctl00_ctl00_lnbLogin"))).click()  # Replace with actual logout button ID
                time.sleep(2)
                logging.info(f"Successfully logged out for {username}")
            except NoSuchElementException:
                logging.error(f"Logout failed for {username}. Reloading login page.")
                driver.get(login_url)  # Reload login page if logout fails
        else:
            driver.get(login_url)  # Reload login page on failure

    except WebDriverException as e:
        logging.error(f"Error for {username}: {str(e)}")
        take_screenshot(username, "Fail")
        driver.get(login_url)  # Reload login page on network failure

# Close the browser
driver.quit()
