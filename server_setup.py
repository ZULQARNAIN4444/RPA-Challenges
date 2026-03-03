from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
import glob
import pandas as pd

url = "https://botsdna.com/server/"
DOWNLOAD_PATH = r"B:\BotsDNA"

excel_download_xpath = "//a[contains(text(),'input.xlsx')]"
chose_os_xpath = "//*[@id='os']"
ram_size_xpath = "//*[@id='Ram']"
server_xpath = "//*[@id='CreateServer']"


chrome_options = Options()
prefs = {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 5)


driver.get(url)


wait.until(EC.element_to_be_clickable((By.XPATH, excel_download_xpath))).click()
time.sleep(5)

files = glob.glob(os.path.join(DOWNLOAD_PATH, "*.xlsx"))
latest_file = max(files, key=os.path.getctime)

df = pd.read_excel(latest_file)

if "status" not in df.columns:
    df["status"] = ""



for index, row in df.iterrows():

    request_id = row["RequestID"]
    os_value = str(row["OS"]).strip()
    ram_value = str(row["RAM"]).strip()
    hdd_value = str(row["HDD"]).strip()
    applications = [app.strip() for app in str(row["Applications"]).split(",")]

    print(f"\nProcessing RequestID: {request_id}")

    # Reload page fresh each time
    driver.get(url)
    wait.until(EC.presence_of_element_located((By.XPATH, chose_os_xpath)))
    time.sleep(2)


    # SELECT OS
    Select(driver.find_element(By.XPATH, chose_os_xpath)
           ).select_by_visible_text(os_value)
    time.sleep(1)

    # SELECT RAM
 
    Select(driver.find_element(By.XPATH, ram_size_xpath)
           ).select_by_visible_text(ram_value)
    time.sleep(1)

   
    # SELECT HDD (INDEX MATCH FIX)

    for i in range(1, 6):

        label_xpath = f"//tr[3]/td[2]/label[{i}]"
        input_xpath = f"//tr[3]/td[2]/input[{i}]"

        label_text = driver.find_element(By.XPATH, label_xpath).text

        clean_label = label_text.replace(" ", "").strip().lower()
        clean_excel = hdd_value.replace(" ", "").strip().lower()

        if clean_label == clean_excel:
            radio_button = driver.find_element(By.XPATH, input_xpath)
            driver.execute_script("arguments[0].click();", radio_button)
            break

    time.sleep(1)

    # SELECT APPLICATIONS (INDEX MATCH FIX)

    for app in applications:

        clean_excel_app = app.strip().lower()

        for i in range(1, 10):

            label_xpath = f"//tr[4]/td[2]/label[{i}]"
            input_xpath = f"//tr[4]/td[2]/input[{i}]"

            label_text = driver.find_element(By.XPATH, label_xpath).text
            clean_label = label_text.strip().lower()

            if clean_label == clean_excel_app:
                checkbox = driver.find_element(By.XPATH, input_xpath)

                if not checkbox.is_selected():
                    driver.execute_script("arguments[0].click();", checkbox)
                break

    time.sleep(1)

    # CLICK CREATE SERVER
   
    driver.find_element(By.XPATH, server_xpath).click()


    wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
    time.sleep(2)

    print(f"{request_id} -> server confirmed")

    df.loc[index, "status"] = "server confirmed"

    # Wait before next iteration
    time.sleep(1)

# SAVE OUTPUT

output_path = os.path.join(DOWNLOAD_PATH, "Server_Output.xlsx")
df.to_excel(output_path, index=False)

print(" All Requests Processed Successfully")

driver.quit()