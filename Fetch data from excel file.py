from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import time
import openpyxl
import logging
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import WebDriverException, NoSuchElementException, TimeoutException, \
    UnexpectedAlertPresentException

# Configure logging
log_file = "ghj.txt"
logging.basicConfig(filename=log_file, level=logging.INFO, format="%(asctime)s - %(message)s")

# Create a directory for screenshots if it doesn't exist
screenshot_dir = "screenshots"
if not os.path.exists(screenshot_dir):
    os.makedirs(screenshot_dir)

workbook = openpyxl.load_workbook("Excel_file/Username.xlsx")
sheet = workbook["LoginTest"]

driver = webdriver.Chrome()
driver.maximize_window()
login_url = ("https://training.pfms.gov.in/SitePages/Users/LoginDetails/Login.aspx")
driver.get(login_url)
time.sleep(5)

def take_screenshot(username, error_type):
    screenshot_path = os.path.join(screenshot_dir, f"{error_type}_{username}.png")
    driver.save_screenshot(screenshot_path)
    logging.error(f"Screenshot saved at {screenshot_path}")


for r in sheet.iter_rows(min_row=2, values_only=True):
    Username, Password, expected_result = r
    try:
        driver.get(login_url)
        WebDriverWait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='UserName']")))
        time.sleep(5)
        Username = driver.find_element(By.XPATH, "//input[@id='UserName']")
        Username.send_keys(Username)
        Passcode = driver.find_element(By.XPATH, "//input[@id='Password']")
        Passcode.send_keys(Password)
        actions = ActionChains(driver)
        actions.scroll_by_amount(0, 300).perform()
        time.sleep(10)
        Login = driver.find_element(By.XPATH, "//input[@value='Log In']")
        Login.click()

        # Handle pop-ups or alerts
        try:
            alert = driver.switch_to.alert
            logging.warning(f"Unexpected alert detected for {Username}: {alert.text}")
            alert.accept()
        except UnexpectedAlertPresentException:
            logging.warning(f"Unhandled alert present for {Username}")

        # Check if login was successful
        try:
            WebDriverWait.until(EC.presence_of_element_located((By.LINK_TEXT, "Central Plan Scheme Monitoring System [CPSMS]")))  # Replace with actual landing page main element ID
            time.sleep(5)
            result = "Pass"
        except TimeoutException:
            result = "Fail"

        # Log results
        logging.info(f"Test case: Username={Username}, Expected={expected_result}, Got={result}")

        # Take screenshot on failure
        if result != expected_result:
            take_screenshot(Username, "failure")

        # Logout if login was successful
        if result == "Pass":
            try:
                WebDriverWait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ctl00_lnbLogin"))).click()  # Replace with actual logout button ID
                time.sleep(5)
                logging.info(f"Successfully logged out for {Username}")
            except NoSuchElementException:
                logging.error(f"Logout failed for {Username}. Reloading login page.")
                driver.get(login_url)  # Reload login page if logout fails
        else:
            driver.get(login_url)  # Reload login page on failure

    except WebDriverException as e:
            logging.error(f"Error for {Username}: {str(e)}")
            take_screenshot(Username, "Fail")
            driver.get(login_url)  # Reload login page on network failure

# Close the browser
driver.quit()

