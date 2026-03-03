from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import openpyxl

URL = "https://botsdna.com/locator/"

def sanitize_sheet_name(name: str) -> str:
    bad = ['\\', '/', '*', '[', ']', ':', '?']
    for ch in bad:
        name = name.replace(ch, '-')
    return name[:31].strip() or "Sheet1"


def main():
    driver = webdriver.Chrome()
    driver.get(URL)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )

    time.sleep(2)

    tables = driver.find_elements(By.TAG_NAME, "table")

    target_table = None
    headers = []

    
    for table in tables:
        ths = table.find_elements(By.TAG_NAME, "th")
        header_texts = [th.text.strip() for th in ths]

        if any("Customer" in h for h in header_texts):
            target_table = table
            headers = header_texts
            break

    if not target_table:
        driver.quit()
        raise SystemExit("Could not find Locator Table.")

    rows = target_table.find_elements(By.TAG_NAME, "tr")

    if not rows:
        driver.quit()
        raise SystemExit("No rows found.")

    
    headers = [h if h else f"Column{i}" for i, h in enumerate(headers)]

    
    headers[0] = "customer name"

    country_data = {h: [] for h in headers[1:]}

    
    for row in rows[1:]:
        cells = row.find_elements(By.TAG_NAME, "td")

        if not cells:
            continue

        customer = cells[0].text.strip()

        for i in range(1, len(cells)):
            if i < len(headers):
                value = cells[i].text.strip()

                if value and value not in ("0", "0.0", "-", "N/A"):
                    country = headers[i]
                    country_data[country].append({
                        "customer name": customer,
                        "location value": value
                    })

    driver.quit()

    
    country_data = {k: v for k, v in country_data.items() if v}

    if not country_data:
        print("No valid data found.")
        return

    with pd.ExcelWriter("customers_by_country.xlsx", engine="openpyxl") as writer:
        for country, records in country_data.items():
            df = pd.DataFrame(records)
            df.to_excel(writer, sheet_name=sanitize_sheet_name(country), index=False)

    print("File created successfully: customers_by_country.xlsx")


if __name__ == "__main__":
    main()