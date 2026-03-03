from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

BASE_URL = "https://botsdna.com/school/"
DOWNLOAD_PATH = r"B:\BotsDNA"
EXCEL_NAME = "Master Template.xlsx"
EXCEL_PATH = os.path.join(DOWNLOAD_PATH, EXCEL_NAME)

EXCEL_DOWNLOAD_XPATH = "//html/body/center/font/a[2]"

XPATHS = {
    "School Name": "//html/body/center/h1[1]",
    "School Address": "//html/body/center/table[1]/tbody/tr[1]/td[2]",
    "School Phonenumber": "//html/body/center/table[1]/tbody/tr[2]/td[2]",
    "Number of Student": "//html/body/center/table[1]/tbody/tr[3]/td[2]",
    "Prncipal Name": "//html/body/center/table[1]/tbody/tr[4]/td[2]",
    "Number of TeachingStaff": "//html/body/center/table[1]/tbody/tr[5]/td[2]",
    "Number of Non-TeachingStaff": "//html/body/center/table[1]/tbody/tr[6]/td[2]",
    "Number of School buses": "//html/body/center/table[1]/tbody/tr[7]/td[2]",
    "School Playground": "//html/body/center/table[1]/tbody/tr[8]/td[2]",
    "Facilities": "//html/body/center/table[1]/tbody/tr[9]/td[2]",
    "School Accrediation": "//html/body/center/table[1]/tbody/tr[10]/td[2]",
    "School Hostel": "//html/body/center/table[1]/tbody/tr[11]/td[2]",
    "School Canteen": "//html/body/center/table[1]/tbody/tr[12]/td[2]",
    "School Stationary": "//html/body/center/table[1]/tbody/tr[13]/td[2]",
    "School Teaching method's": "//html/body/center/table[1]/tbody/tr[14]/td[2]",
    "School Timing": "//html/body/center/table[1]/tbody/tr[15]/td[2]",
    "School Achivements": "//html/body/center/table[1]/tbody/tr[16]/td[2]",
    "School Awards": "//html/body/center/table[1]/tbody/tr[17]/td[2]",
    "School Uniform": "//html/body/center/table[1]/tbody/tr[18]/td[2]",
    "School type": "//html/body/center/table[1]/tbody/tr[19]/td[2]",
}


def setup_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    return webdriver.Chrome(options=options)


def safe_text(driver, wait, xpath):
    try:
        return wait.until(
            EC.presence_of_element_located((By.XPATH, xpath))
        ).text.strip()
    except:
        return ""

def main():
    driver = setup_driver()
    wait = WebDriverWait(driver, 15)

   
    driver.get(BASE_URL)

    if not os.path.exists(EXCEL_PATH):
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, EXCEL_DOWNLOAD_XPATH)
        )).click()
        time.sleep(5)

  
    df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")

    #  Normalize Excel headers so pandas matches them exactly
    df.columns = (
        df.columns
        .str.strip()
        .str.replace("\n", " ")
        .str.replace(r"\s+", " ", regex=True)
    )

    #  Iterate school codes
    for index, row in df.iterrows():
        school_code = row["School Code"].strip()

        if not school_code:
            continue

        print(f" Processing School Code: {school_code}")

        detail_url = f"{BASE_URL}{school_code}.html"
        driver.get(detail_url)

        #  Extract data
        for column, xpath in XPATHS.items():
            df.at[index, column] = safe_text(driver, wait, xpath)

    #  Save Excel
    df.to_excel(EXCEL_PATH, index=False)
    driver.quit()

    print(" SCHOOL AUTOMATION COMPLETED SUCCESSFULLY")

if __name__ == "__main__":
    main()