from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

URL = "https://botsdna.com/notaries/"
DOWNLOAD_PATH = r"B:\BotsDNA"
EXCEL_NAME = "AP-ADVOCATES.xlsx"
EXCEL_PATH = os.path.join(DOWNLOAD_PATH, EXCEL_NAME)

EXCEL_DOWNLOAD_XPATH = "//html/body/center/font/a[2]"
DISTRICT_XPATH = '//*[@id="DIST"]'
NOTARY_NAME_XPATH = '//*[@id="notary"]'
AREA_PRACTICE_XPATH = '//*[@id="area"]'
SUBMIT_XPATH = "//input[@value='Submit Notary']"
TRANSACTION_XPATH = '//*[@id="TransNo"]'

def setup_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)

def clean_district(text):
    return text.upper().replace("DIST", "").strip()

def main():
    driver = setup_driver()
    wait = WebDriverWait(driver, 30)
    driver.get(URL)

    if not os.path.exists(EXCEL_PATH):
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, EXCEL_DOWNLOAD_XPATH)
        )).click()
        time.sleep(5)

    df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
    if "Transaction Number" not in df.columns:
        df["Transaction Number"] = ""

    current_district = None

    for idx, row in df.iterrows():
        sl_no = row["SL.NO."].strip()

        # District header
        if "DIST" in sl_no.upper():
            current_district = clean_district(sl_no)
            print(f" District set to: {current_district}")
            continue

        if not current_district or row["Transaction Number"].strip():
            continue

        notary = row["NOTARY ADVOCATE NAME"]
        area = row["AREA OF PRACTICE"]

        if not notary:
            continue

        print(f"Processing → {current_district} | {notary}")

        # Select district
        Select(wait.until(
            EC.element_to_be_clickable((By.XPATH, DISTRICT_XPATH))
        )).select_by_visible_text(current_district)

        # Fill fields
        wait.until(EC.visibility_of_element_located((By.XPATH, NOTARY_NAME_XPATH))).clear()
        driver.find_element(By.XPATH, NOTARY_NAME_XPATH).send_keys(notary)

        driver.find_element(By.XPATH, AREA_PRACTICE_XPATH).clear()
        driver.find_element(By.XPATH, AREA_PRACTICE_XPATH).send_keys(area)

        # Submit
        driver.find_element(By.XPATH, SUBMIT_XPATH).click()

        # Extract transaction number (Thank You page)
        txn = wait.until(
            EC.visibility_of_element_located((By.XPATH, TRANSACTION_XPATH))
        ).text.strip()

        df.at[idx, "Transaction Number"] = txn
        df.to_excel(EXCEL_PATH, index=False)

        print(f" Transaction captured: {txn}")

        #   GO BACK TO FORM PAGE
        driver.get(URL)
        wait.until(EC.element_to_be_clickable((By.XPATH, DISTRICT_XPATH)))
        time.sleep(2)

    driver.quit()
    print(" NOTARY AUTOMATION COMPLETED SUCCESSFULLY")

if __name__ == "__main__":
    main()