from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import time
import os
import zipfile
import glob
import re

URL = "https://botsdna.com/ActiveLoans/"
DOWNLOAD_PATH = r"B:\BotsDNA\xlsx_activeloan"
INPUT_EXCEL = os.path.join(DOWNLOAD_PATH, "input.xlsx")

INPUT_XLSX_XPATH = "//html/body/center/font/a[2]"

def get_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)

driver = get_driver()
driver.get(URL)
wait = WebDriverWait(driver, 15)

if not os.path.exists(INPUT_EXCEL):
    wait.until(EC.element_to_be_clickable((By.XPATH, INPUT_XLSX_XPATH))).click()

    for _ in range(30):
        if os.path.exists(INPUT_EXCEL):
            break
        time.sleep(1)
    else:
        raise Exception("input.xlsx download failed")


df = pd.read_excel(INPUT_EXCEL, dtype=str)

string_columns = [
    "Status", "PAN NUMBER", "Bank",
    "Branch", "Loan Taken On",
    "Amount", "EMI(month)"
]

for col in string_columns:
    if col not in df.columns:
        df[col] = ""
    df[col] = df[col].astype("string")

df["Last4"] = df["AccountNumber"].astype(str).str[-4:]


wait.until(
    EC.presence_of_element_located(
        (By.XPATH, "//html/body/center/table/tbody/tr[2]")
    )
)

for index, row in df.iterrows():
    excel_last4 = row["Last4"]
    match_found = False

    rows = driver.find_elements(
        By.XPATH,
        "//html/body/center/table/tbody/tr[position()>1]"
    )

    for i in range(2, len(rows) + 2):
        status_xpath = f"//html/body/center/table/tbody/tr[{i}]/td[1]"
        loan_xpath   = f"//html/body/center/table/tbody/tr[{i}]/td[2]/a"
        pan_xpath    = f"//html/body/center/table/tbody/tr[{i}]/td[3]"

        loan_elem = driver.find_element(By.XPATH, loan_xpath)
        loan_text = loan_elem.text.strip()

        loan_last4 = re.findall(r"\d{4}$", loan_text)

        if loan_last4 and loan_last4[0] == excel_last4:
            match_found = True

            df.at[index, "Status"] = driver.find_element(By.XPATH, status_xpath).text
            df.at[index, "PAN NUMBER"] = driver.find_element(By.XPATH, pan_xpath).text

            loan_elem.click()
            time.sleep(5)

            zip_files = glob.glob(os.path.join(DOWNLOAD_PATH, "*.zip"))
            if zip_files:
                latest_zip = max(zip_files, key=os.path.getctime)
                with zipfile.ZipFile(latest_zip, 'r') as z:
                    z.extractall(DOWNLOAD_PATH)
                os.remove(latest_zip)

            break

    if not match_found:
        print(f"No LOAN CODE match found for account ending {excel_last4}")

driver.quit()


def parse_txt(path):
    data = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if ":" in line:
                k, v = line.split(":", 1)
                data[k.strip()] = v.strip()
    return data

df.columns = df.columns.str.strip()

for txt in glob.glob(os.path.join(DOWNLOAD_PATH, "*.txt")):
    data = parse_txt(txt)

    acc = data.get("Account Number", "")
    idx = df[df["AccountNumber"].astype(str) == acc].index

    if not idx.empty:
        i = idx[0]

        if "Bank" in df.columns:
            df.at[i, "Bank"] = data.get("Bank", "")

        if "Branch" in df.columns:
            df.at[i, "Branch"] = data.get("Branch", "")

        if "Loan Taken On" in df.columns:
            df.at[i, "Loan Taken On"] = data.get("Loan Taken On", "")

        if "Amount" in df.columns:
            df.at[i, "Amount"] = data.get("Amount", "")

        if "EMI(month)" in df.columns:
            df.at[i, "EMI(month)"] = data.get("EMI(month)", "")

    os.remove(txt)


if "Loan Taken on" in df.columns:
    df.drop(columns=["Loan Taken on"], inplace=True)

df.drop(columns=["Last4"], inplace=True)
df.to_excel(INPUT_EXCEL, index=False)

print("ACTIVE LOAN AUTOMATION COMPLETED SUCCESSFULLY")
