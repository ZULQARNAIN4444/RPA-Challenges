from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time

url = "https://rpachallenge.com/"
DOWNLOAD_PATH = r"B:\BotsDNA"
EXCEL_NAME = "challenge.xlsx"
EXCEL_PATH = os.path.join(DOWNLOAD_PATH, EXCEL_NAME)

def setup_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)

def main():
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)

    driver.get(url)

    
    if not os.path.exists(EXCEL_PATH):
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Download Excel')]")
        )).click()
        time.sleep(5)

    df = pd.read_excel(EXCEL_PATH).fillna("")

    
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/button[1]")
    )).click()

    for index, row in df.iterrows():

        # Create dictionary from Excel row
        data_map = {
            "First Name": str(row["First Name"]).strip(),
            "Last Name": str(row["Last Name "]).strip(),
            "Company Name": str(row["Company Name"]).strip(),
            "Role in Company": str(row["Role in Company"]).strip(),
            "Address": str(row["Address"]).strip(),
            "Email": str(row["Email"]).strip(),
            "Phone Number": str(row["Phone Number"]).strip()
        }

        
        labels = driver.find_elements(By.XPATH, "//label")

        for label in labels:
            field_name = label.text.strip()

            if field_name in data_map:
                input_box = label.find_element(
                    By.XPATH, "following-sibling::input"
                )

                input_box.clear()
                input_box.send_keys(data_map[field_name])

        # Click Submit
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[2]/form[1]/input[1]")
        )).click()

    print("Challenge Completed Successfully ")
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    main()


